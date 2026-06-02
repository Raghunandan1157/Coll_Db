#!/usr/bin/env python3
"""
Apply March 2026 POS to portfolio_month_wise.

Strategy (preserves existing demand/collection on portfolio_performance):
  1. UPDATE existing perf rows (month_id=24) with new POS cols
  2. Auto-create regions/districts/branches/employees for orphan emps
  3. INSERT new (emp, product) rows with demand=0, POS=new
  4. Zero POS cols on DB-only rows (had demand, no March POS)
  5. DELETE+INSERT branch_pos for month_id=24
  6. DELETE+INSERT employee_pos for month_id=24

Run as:  python3 data/march_pos_apply.py [--host HOST]
Default host = 127.0.0.1 (local). Pass --host ec2-user@52.66.163.52 to run via SSH.
"""
import argparse
import sys
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values

MONTH_ID = 24  # MAR
EXPECTED_TOTAL_CR = 1051.89

PRODUCT_MAP = {
    "IGL": "IGL", "FIG": "FIG",
    "VVY": "IL", "MEL": "IL", "QR Loan": "IL",
    "Nirmal/Jeevadhara": "IGL", "CDL": "IGL", "Vidya Jyothi": "IGL",
}
PRODUCT_TYPE_ID = {"IGL": 1, "FIG": 2, "IL": 3}
ASSETCLASS_TO_BUCKET = {
    "Regular": "regular_pos", "SMA0": "sma0_pos", "SMA1": "sma1_pos",
    "SMA2": "pnpa_pos",
    "Sub Standard": "npa_pos", "Doubtful1": "npa_pos",
    "Doubtful2": "npa_pos", "Doubtful3": "npa_pos",
}
POS_COLS = ["regular_pos", "sma0_pos", "sma1_pos", "pnpa_pos", "npa_pos"]


def aggregate(xlsx):
    print("Reading Data sheet...", flush=True)
    df = pd.read_excel(xlsx, sheet_name="Data", engine="openpyxl",
                       usecols=["BranchName", "Region", "Division", "EMP ID",
                                "OfficerName", "Product Name", "AssetClass",
                                "PrincipalOS"])
    df = df.dropna(subset=["EMP ID", "Product Name", "AssetClass"])
    df["EMP ID"] = df["EMP ID"].astype(str).str.strip()
    df["Region"] = df["Region"].astype(str).str.strip()
    df["Division"] = df["Division"].astype(str).str.strip()
    df["BranchName"] = df["BranchName"].astype(str).str.strip()
    df["OfficerName"] = df["OfficerName"].astype(str).str.strip()

    df["product"] = df["Product Name"].map(PRODUCT_MAP)
    df["bucket"] = df["AssetClass"].map(ASSETCLASS_TO_BUCKET)
    df["product_type_id"] = df["product"].map(PRODUCT_TYPE_ID)
    if df["product_type_id"].isna().any() or df["bucket"].isna().any():
        sys.exit("FATAL: unmapped product or AssetClass")

    total_cr = df["PrincipalOS"].sum() / 1e7
    print(f"  rows={len(df):,}  total={total_cr:.2f} Cr", flush=True)

    # PERF aggregation (emp, product) plus emp metadata for orphan creation
    perf_pivot = (df.groupby(["EMP ID", "product_type_id", "bucket"])["PrincipalOS"]
                    .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in perf_pivot.columns:
            perf_pivot[c] = 0.0
    perf_pivot["total_pos"] = perf_pivot[POS_COLS].sum(axis=1)
    # Emp metadata (one row per emp_id) — pick first non-null branch info
    emp_meta = (df.groupby("EMP ID")
                  .agg(officer_name=("OfficerName", "first"),
                       region_name=("Region", "first"),
                       district_name=("Division", "first"),
                       branch_name=("BranchName", "first"))
                  .reset_index())

    # BRANCH_POS aggregation
    bp = (df.groupby(["Region", "Division", "BranchName", "product", "bucket"])["PrincipalOS"]
            .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in bp.columns: bp[c] = 0.0
    bp["total_pos"] = bp[POS_COLS].sum(axis=1)
    bp.columns = ["region_name", "district_name", "branch_name", "product_name",
                  *POS_COLS, "total_pos"]

    # EMP_POS aggregation
    ep = (df.groupby(["EMP ID", "bucket"])["PrincipalOS"]
            .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in ep.columns: ep[c] = 0.0
    ep["total_pos"] = ep[POS_COLS].sum(axis=1)
    ep.columns = ["emp_id", *POS_COLS, "total_pos"]

    # Sanity
    t1, t2, t3 = (perf_pivot["total_pos"].sum(),
                  bp["total_pos"].sum(),
                  ep["total_pos"].sum())
    assert abs(t1 - t2) < 1 and abs(t2 - t3) < 1, "agg totals mismatch"
    return perf_pivot, emp_meta, bp, ep


def apply_to_db(perf, emp_meta, bp, ep, conn):
    cur = conn.cursor()
    cur.execute("BEGIN")
    try:
        # Pre-snapshot
        cur.execute("SELECT SUM(total_pos) FROM portfolio_performance WHERE month_id=%s", (MONTH_ID,))
        pre_pos = float(cur.fetchone()[0] or 0)
        print(f"\nPRE: portfolio_performance month_id={MONTH_ID} total_pos = {pre_pos/1e7:.2f} Cr", flush=True)

        # ── 1. Auto-create regions/districts/branches/employees (mirror endpoint logic) ──
        cur.execute("SELECT region_id, region_name FROM regions")
        regions = {r[1]: r[0] for r in cur.fetchall()}
        cur.execute("SELECT district_id, district_name, region_id FROM districts")
        districts = {(r[1], r[2]): r[0] for r in cur.fetchall()}
        cur.execute("SELECT branch_id, branch_name, district_id FROM branches")
        branches = {(r[1], r[2]): r[0] for r in cur.fetchall()}
        cur.execute("SELECT emp_id FROM employees")
        emp_set = set(r[0] for r in cur.fetchall())

        new_regions = new_districts = new_branches = new_emps = 0
        for _, row in emp_meta.iterrows():
            rn = row["region_name"]; dn = row["district_name"]; bn = row["branch_name"]
            if rn not in regions:
                cur.execute("INSERT INTO regions (region_name) VALUES (%s) ON CONFLICT (region_name) DO UPDATE SET region_name=EXCLUDED.region_name RETURNING region_id", (rn,))
                regions[rn] = cur.fetchone()[0]; new_regions += 1
            rid = regions[rn]
            dkey = (dn, rid)
            if dkey not in districts:
                cur.execute("INSERT INTO districts (district_name, region_id) VALUES (%s,%s) ON CONFLICT (district_name, region_id) DO UPDATE SET district_name=EXCLUDED.district_name RETURNING district_id", (dn, rid))
                districts[dkey] = cur.fetchone()[0]; new_districts += 1
            did = districts[dkey]
            bkey = (bn, did)
            if bkey not in branches:
                cur.execute("INSERT INTO branches (branch_name, district_id) VALUES (%s,%s) ON CONFLICT (branch_name, district_id) DO UPDATE SET branch_name=EXCLUDED.branch_name RETURNING branch_id", (bn, did))
                branches[bkey] = cur.fetchone()[0]; new_branches += 1
            bid = branches[bkey]
            if row["EMP ID"] not in emp_set:
                cur.execute("INSERT INTO employees (emp_id, officer_name, branch_id) VALUES (%s,%s,%s) ON CONFLICT (emp_id) DO UPDATE SET branch_id=EXCLUDED.branch_id", (row["EMP ID"], row["officer_name"], bid))
                emp_set.add(row["EMP ID"]); new_emps += 1
        print(f"  master tables: +{new_regions} regions, +{new_districts} districts, +{new_branches} branches, +{new_emps} employees", flush=True)

        # ── 2. UPSERT portfolio_performance POS cols (preserves demand) ──
        # Build list of (month_id, emp_id, product_type_id, regular_pos, sma0, sma1, pnpa, npa, total)
        upsert_rows = [
            (MONTH_ID, r["EMP ID"], int(r["product_type_id"]),
             float(r["regular_pos"]), float(r["sma0_pos"]), float(r["sma1_pos"]),
             float(r["pnpa_pos"]), float(r["npa_pos"]), float(r["total_pos"]))
            for _, r in perf.iterrows()
        ]
        # First UPDATE existing rows
        cur.execute("""
            CREATE TEMP TABLE _pos_stage (
                month_id int, emp_id varchar(30), product_type_id int,
                regular_pos numeric, sma0_pos numeric, sma1_pos numeric,
                pnpa_pos numeric, npa_pos numeric, total_pos numeric
            ) ON COMMIT DROP
        """)
        execute_values(cur,
            "INSERT INTO _pos_stage VALUES %s",
            upsert_rows)
        # UPDATE matched
        cur.execute("""
            UPDATE portfolio_performance pp SET
              regular_pos = s.regular_pos, sma0_pos = s.sma0_pos,
              sma1_pos = s.sma1_pos, pnpa_pos = s.pnpa_pos,
              npa_pos = s.npa_pos, total_pos = s.total_pos
            FROM _pos_stage s
            WHERE pp.month_id = s.month_id AND pp.emp_id = s.emp_id
              AND pp.product_type_id = s.product_type_id
        """)
        upd = cur.rowcount
        # INSERT missing (perf rows in stage but not in pp)
        cur.execute("""
            INSERT INTO portfolio_performance (month_id, emp_id, product_type_id,
                regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
            SELECT s.month_id, s.emp_id, s.product_type_id,
                   s.regular_pos, s.sma0_pos, s.sma1_pos,
                   s.pnpa_pos, s.npa_pos, s.total_pos
            FROM _pos_stage s
            WHERE NOT EXISTS (
                SELECT 1 FROM portfolio_performance pp
                WHERE pp.month_id=s.month_id AND pp.emp_id=s.emp_id
                  AND pp.product_type_id=s.product_type_id
            )
        """)
        ins = cur.rowcount
        # Zero out POS for DB rows not in stage
        cur.execute("""
            UPDATE portfolio_performance pp SET
              regular_pos=0, sma0_pos=0, sma1_pos=0,
              pnpa_pos=0, npa_pos=0, total_pos=0
            WHERE pp.month_id=%s AND NOT EXISTS (
                SELECT 1 FROM _pos_stage s
                WHERE s.month_id=pp.month_id AND s.emp_id=pp.emp_id
                  AND s.product_type_id=pp.product_type_id
            )
        """, (MONTH_ID,))
        zeroed = cur.rowcount
        print(f"  portfolio_performance: UPDATE={upd}, INSERT={ins}, zeroed={zeroed}", flush=True)

        # ── 3. branch_pos: DELETE+INSERT for month ──
        cur.execute("DELETE FROM branch_pos WHERE month_id=%s", (MONTH_ID,))
        bp_del = cur.rowcount
        bp_rows = [(MONTH_ID, r["region_name"], r["district_name"], r["branch_name"],
                    r["product_name"],
                    float(r["regular_pos"]), float(r["sma0_pos"]), float(r["sma1_pos"]),
                    float(r["pnpa_pos"]), float(r["npa_pos"]), float(r["total_pos"]))
                   for _, r in bp.iterrows()]
        execute_values(cur,
            """INSERT INTO branch_pos (month_id, region_name, district_name, branch_name,
                 product_name, regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
               VALUES %s""", bp_rows)
        print(f"  branch_pos: DELETE={bp_del}, INSERT={len(bp_rows)}", flush=True)

        # ── 4. employee_pos: DELETE+INSERT for month ──
        cur.execute("DELETE FROM employee_pos WHERE month_id=%s", (MONTH_ID,))
        ep_del = cur.rowcount
        ep_rows = [(MONTH_ID, r["emp_id"],
                    float(r["regular_pos"]), float(r["sma0_pos"]), float(r["sma1_pos"]),
                    float(r["pnpa_pos"]), float(r["npa_pos"]), float(r["total_pos"]))
                   for _, r in ep.iterrows()]
        execute_values(cur,
            """INSERT INTO employee_pos (month_id, emp_id, regular_pos, sma0_pos,
                 sma1_pos, pnpa_pos, npa_pos, total_pos) VALUES %s""", ep_rows)
        print(f"  employee_pos: DELETE={ep_del}, INSERT={len(ep_rows)}", flush=True)

        # ── 5. Verify totals ──
        cur.execute("SELECT SUM(total_pos) FROM portfolio_performance WHERE month_id=%s", (MONTH_ID,))
        post_pp = float(cur.fetchone()[0] or 0)
        cur.execute("SELECT SUM(total_pos) FROM branch_pos WHERE month_id=%s", (MONTH_ID,))
        post_bp = float(cur.fetchone()[0] or 0)
        cur.execute("SELECT SUM(total_pos) FROM employee_pos WHERE month_id=%s", (MONTH_ID,))
        post_ep = float(cur.fetchone()[0] or 0)
        print(f"\nPOST: pp={post_pp/1e7:.2f} Cr | branch_pos={post_bp/1e7:.2f} Cr | emp_pos={post_ep/1e7:.2f} Cr", flush=True)
        # All 3 should equal expected
        for label, val in [("pp", post_pp), ("branch_pos", post_bp), ("emp_pos", post_ep)]:
            diff_cr = abs(val/1e7 - EXPECTED_TOTAL_CR)
            if diff_cr > 0.01:
                raise RuntimeError(f"{label} total {val/1e7:.4f} Cr ≠ expected {EXPECTED_TOTAL_CR} Cr (diff {diff_cr:.4f})")
        print(f"\n✓ All 3 tables match expected {EXPECTED_TOTAL_CR} Cr", flush=True)

        cur.execute("COMMIT")
        print("\nCOMMITTED.", flush=True)

    except Exception as e:
        cur.execute("ROLLBACK")
        print(f"\nROLLBACK: {e}", file=sys.stderr)
        raise
    finally:
        cur.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5432)
    ap.add_argument("--xlsx", default="/Users/raghunandanmali/Downloads/03.ESAF POS & Overdue Summary Mar'2026.xlsx")
    args = ap.parse_args()

    perf, emp_meta, bp, ep = aggregate(args.xlsx)

    conn = psycopg2.connect(host=args.host, port=args.port, user="Raghunandan1157",
                            password="raghu", database="portfolio_month_wise")
    apply_to_db(perf, emp_meta, bp, ep, conn)
    conn.close()


if __name__ == "__main__":
    main()
