#!/usr/bin/env python3
"""
March 2026 POS aggregation from ESAF Excel → portfolio_month_wise DB shape.

Reads Data sheet, maps products, aggregates by AssetClass into 6 POS columns.
Outputs 3 CSVs for 3 update targets:
  - perf_pos.csv     → portfolio_performance (emp_id, product_type_id) UPDATE
  - branch_pos.csv   → branch_pos (region, district, branch, product) DELETE+INSERT
  - emp_pos.csv      → employee_pos (emp_id) DELETE+INSERT
"""
import sys
import pandas as pd

XLSX = "/Users/raghunandanmali/Downloads/03.ESAF POS & Overdue Summary Mar'2026.xlsx"
OUT_DIR = "/Users/raghunandanmali/Desktop/Coll_Db/data"

PRODUCT_MAP = {
    "IGL": "IGL",
    "FIG": "FIG",
    "VVY": "IL",
    "MEL": "IL",
    "QR Loan": "IL",
    "Nirmal/Jeevadhara": "IGL",
    "CDL": "IGL",
    "Vidya Jyothi": "IGL",
}

PRODUCT_TYPE_ID = {"IGL": 1, "FIG": 2, "IL": 3}

ASSETCLASS_TO_BUCKET = {
    "Regular":      "regular_pos",
    "SMA0":         "sma0_pos",
    "SMA1":         "sma1_pos",
    "SMA2":         "pnpa_pos",
    "Sub Standard": "npa_pos",
    "Doubtful1":    "npa_pos",
    "Doubtful2":    "npa_pos",
    "Doubtful3":    "npa_pos",
}

POS_COLS = ["regular_pos", "sma0_pos", "sma1_pos", "pnpa_pos", "npa_pos"]


def main():
    print(f"Reading {XLSX}...", flush=True)
    df = pd.read_excel(XLSX, sheet_name="Data", engine="openpyxl",
                       usecols=["BranchName", "Region", "Division", "EMP ID",
                                "OfficerName", "Product Name", "AssetClass",
                                "PrincipalOS"])
    print(f"  rows: {len(df):,}", flush=True)

    # Validate
    unknown_products = set(df["Product Name"].dropna().unique()) - set(PRODUCT_MAP)
    unknown_classes = set(df["AssetClass"].dropna().unique()) - set(ASSETCLASS_TO_BUCKET)
    if unknown_products:
        print(f"FATAL: unknown products: {unknown_products}", file=sys.stderr)
        sys.exit(1)
    if unknown_classes:
        print(f"FATAL: unknown AssetClass: {unknown_classes}", file=sys.stderr)
        sys.exit(1)

    # Map
    df["product"] = df["Product Name"].map(PRODUCT_MAP)
    df["bucket"] = df["AssetClass"].map(ASSETCLASS_TO_BUCKET)
    df["product_type_id"] = df["product"].map(PRODUCT_TYPE_ID)

    total_cr = df["PrincipalOS"].sum() / 1e7
    print(f"  total PrincipalOS: {total_cr:.2f} Cr", flush=True)

    # Schema for portfolio_performance: region/district/branch are stored on
    # employees → branches → districts → regions; the perf row keys on emp_id.
    # In the Excel: Region = state-level (e.g., CHITRADURGA), Division = district.
    # branch_pos table: keys on (region_name, district_name, branch_name, product_name).

    # ── 1) portfolio_performance: (emp_id, product_type_id) → POS buckets ──
    perf = (df.groupby(["EMP ID", "product_type_id", "bucket"])["PrincipalOS"]
              .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in perf.columns: perf[c] = 0.0
    perf["total_pos"] = perf[POS_COLS].sum(axis=1)
    perf = perf.rename(columns={"EMP ID": "emp_id"})
    perf = perf[["emp_id", "product_type_id", *POS_COLS, "total_pos"]]
    perf.to_csv(f"{OUT_DIR}/perf_pos.csv", index=False)
    print(f"  perf_pos: {len(perf):,} rows, total={perf['total_pos'].sum()/1e7:.2f} Cr", flush=True)

    # ── 2) branch_pos: (region, division, branch, product) → POS ──
    bp = (df.groupby(["Region", "Division", "BranchName", "product", "bucket"])["PrincipalOS"]
            .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in bp.columns: bp[c] = 0.0
    bp["total_pos"] = bp[POS_COLS].sum(axis=1)
    bp = bp.rename(columns={"Region": "region_name",
                            "Division": "district_name",
                            "BranchName": "branch_name",
                            "product": "product_name"})
    bp = bp[["region_name", "district_name", "branch_name", "product_name",
             *POS_COLS, "total_pos"]]
    bp.to_csv(f"{OUT_DIR}/branch_pos.csv", index=False)
    print(f"  branch_pos: {len(bp):,} rows, total={bp['total_pos'].sum()/1e7:.2f} Cr", flush=True)

    # ── 3) employee_pos: (emp_id) → POS (no product split) ──
    ep = (df.groupby(["EMP ID", "bucket"])["PrincipalOS"]
            .sum().unstack(fill_value=0).reset_index())
    for c in POS_COLS:
        if c not in ep.columns: ep[c] = 0.0
    ep["total_pos"] = ep[POS_COLS].sum(axis=1)
    ep = ep.rename(columns={"EMP ID": "emp_id"})
    ep = ep[["emp_id", *POS_COLS, "total_pos"]]
    ep.to_csv(f"{OUT_DIR}/emp_pos.csv", index=False)
    print(f"  emp_pos: {len(ep):,} rows, total={ep['total_pos'].sum()/1e7:.2f} Cr", flush=True)

    # Sanity: all 3 totals should match
    t1 = perf['total_pos'].sum()
    t2 = bp['total_pos'].sum()
    t3 = ep['total_pos'].sum()
    print(f"\nTallies (₹): perf={t1:,.2f}  branch={t2:,.2f}  emp={t3:,.2f}", flush=True)
    if not (abs(t1 - t2) < 1 and abs(t2 - t3) < 1):
        print("FATAL: aggregation totals mismatch", file=sys.stderr)
        sys.exit(1)
    print(f"All match ✓  total = {t1/1e7:.2f} Cr", flush=True)


if __name__ == "__main__":
    main()
