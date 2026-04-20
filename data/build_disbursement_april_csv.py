#!/usr/bin/env python3
"""Process April disbursement CSV → SQL for disbursement_daily.

Fixes column misalignment caused by unquoted commas in LoanPurpose.
Filters DisbStatus = 'Active' (drops Cancelled).
Uses lookups pulled from EC2 to resolve region/district + emp_id.
"""
import csv
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

CSV_PATH = Path("/Users/raghunandanmali/Downloads/ESAF-ClientDisbursementDetail(Consolidated)_NAV6752314_20042026105812.csv")
BRANCH_LOOKUP = Path("/tmp/branch_lookup.txt")
EMP_LOOKUP = Path("/tmp/emp_lookup.txt")
OUT_SQL = Path("/Users/raghunandanmali/Desktop/Coll_Db/data/disbursement_april_csv.sql")

PRODUCT_MAP = {
    "204207": "IGL", "104207": "IGL", "264203": "IGL", "234008": "IGL",
    "604001": "FIG",
    "81402": "IL",
}

# Header has 39 cols. Last 6 are fixed: Sector, DisbursmentMode, ProfileCIFID,
# OperativeAccount, ValueDate, DisbStatus. First 32 cols are also fixed.
# The middle field (LoanPurpose, idx 32) may contain unquoted commas.
HEADER_LEN = 39
TAIL_LEN = 6   # Sector, Mode, CIF, OpAcc, ValueDate, DisbStatus
HEAD_LEN = 32  # Sl No ... InsuranceFee


def fix_row(row):
    n = len(row)
    if n == HEADER_LEN:
        return row
    if n < HEADER_LEN:
        return None
    head = row[:HEAD_LEN]
    middle = ",".join(row[HEAD_LEN:n - TAIL_LEN])
    tail = row[n - TAIL_LEN:]
    return head + [middle] + tail


def parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%d-%b-%Y", "%d-%b-%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def clean_officer(name):
    if not name:
        return None
    return re.sub(r"\s*\([^)]*\)?\s*$", "", str(name)).strip() or None


def map_product(pid):
    if not pid:
        return None
    pid = str(pid).strip().split(".")[0]
    return PRODUCT_MAP.get(pid)


def load_lookups():
    branch_lookup = {}
    for line in BRANCH_LOOKUP.read_text().splitlines():
        parts = line.split("|")
        if len(parts) >= 3 and parts[0]:
            b, r, d = parts[0].strip(), parts[1].strip(), parts[2].strip()
            branch_lookup[b.upper()] = (r or None, d or None)

    emp_lookup = {}
    for line in EMP_LOOKUP.read_text().splitlines():
        parts = line.split("|")
        if len(parts) >= 3 and parts[0] and parts[2]:
            b, o, e = parts[0].strip(), parts[1].strip(), parts[2].strip()
            emp_lookup[(b.upper(), o.upper())] = e
    return branch_lookup, emp_lookup


def sql_str(v):
    if v is None or v == "":
        return "NULL"
    return "'" + str(v).replace("'", "''") + "'"


def main():
    branch_lookup, emp_lookup = load_lookups()
    print(f"branches loaded: {len(branch_lookup)}, emp mappings: {len(emp_lookup)}")

    agg = defaultdict(lambda: [0, 0.0])
    skipped = defaultdict(int)
    total = 0
    active = 0
    cancelled = 0
    unmatched_branches = set()
    unmatched_officers = set()
    dates = set()

    with CSV_PATH.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader)
        for raw in reader:
            total += 1
            row = fix_row(raw)
            if row is None or len(row) != HEADER_LEN:
                skipped["bad_row_len"] += 1
                continue

            status = (row[38] or "").strip()
            if status == "Cancelled":
                cancelled += 1
                continue
            if status != "Active":
                skipped[f"bad_status:{status}"] += 1
                continue
            active += 1

            d = parse_date(row[25])  # LoanDisbDate
            if not d:
                skipped["bad_date"] += 1
                continue

            product = map_product(row[19])  # SchemeID/ProductID
            if not product:
                skipped["bad_product"] += 1
                continue

            try:
                amount = float(str(row[26] or "0").replace(",", ""))
            except ValueError:
                skipped["bad_amount"] += 1
                continue

            branch = (row[7] or "").strip() or None  # BranchName
            officer = clean_officer(row[10])  # CreditOffcierName

            region = district = None
            if branch:
                hit = branch_lookup.get(branch.upper())
                if hit:
                    region, district = hit
                else:
                    unmatched_branches.add(branch)

            emp_id = None
            if branch and officer:
                emp_id = emp_lookup.get((branch.upper(), officer.upper()))
                if not emp_id:
                    unmatched_officers.add((branch, officer))

            key = (d, region, district, branch, emp_id, officer, product)
            agg[key][0] += 1
            agg[key][1] += amount
            dates.add(d)

    print(f"raw rows: {total}")
    print(f"  active: {active}, cancelled (dropped): {cancelled}")
    print(f"  skipped: {dict(skipped)}")
    print(f"  aggregated rows: {len(agg)}")
    print(f"  date range: {min(dates)} → {max(dates)}")
    print(f"  unmatched branches: {len(unmatched_branches)}")
    print(f"  unmatched officers: {len(unmatched_officers)}")
    if unmatched_branches:
        print(f"    examples: {list(unmatched_branches)[:5]}")
    if unmatched_officers:
        print(f"    examples: {list(unmatched_officers)[:5]}")

    min_d = min(dates).strftime("%Y-%m-%d")
    max_d = max(dates).strftime("%Y-%m-%d")

    lines = ["BEGIN;",
             f"-- Replace April CSV-derived rows ({min_d} → {max_d})",
             f"DELETE FROM disbursement_daily WHERE disb_date >= '{min_d}' AND disb_date <= '{max_d}';",
             ""]

    cols = "(disb_date, region_name, district_name, branch_name, emp_id, officer_name, product_name, disb_count, disb_amount)"
    batch = []
    for (d, region, district, branch, emp_id, officer, product), (cnt, amt) in agg.items():
        vals = (
            f"('{d.strftime('%Y-%m-%d')}', {sql_str(region)}, {sql_str(district)}, "
            f"{sql_str(branch)}, {sql_str(emp_id)}, {sql_str(officer)}, "
            f"{sql_str(product)}, {cnt}, {amt:.2f})"
        )
        batch.append(vals)
        if len(batch) >= 500:
            lines.append(f"INSERT INTO disbursement_daily {cols} VALUES")
            lines.append(",\n".join(batch) + ";")
            batch = []
    if batch:
        lines.append(f"INSERT INTO disbursement_daily {cols} VALUES")
        lines.append(",\n".join(batch) + ";")

    lines.append("COMMIT;")
    OUT_SQL.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote SQL → {OUT_SQL}")


if __name__ == "__main__":
    main()
