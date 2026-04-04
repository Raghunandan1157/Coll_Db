#!/usr/bin/env python3
"""
Migration script: Populate v2_* tables from existing employee_performance data
using ESAF branch mapping for new Region→Division→Area→Branch hierarchy.
"""

import openpyxl
import psycopg2
from psycopg2.extras import execute_values

ESAF_FILE = '/Users/raghunandanmali/Downloads/ESAF Branch Info updated.xlsx'
SHEET_NAME = 'ESAF Original'

DB_CONFIG = {
    'host': '52.66.163.52',
    'port': 5432,
    'user': 'Raghunandan1157',
    'password': 'raghu',
    'database': 'postgres',
}

# ── 1. Load ESAF mapping ────────────────────────────────────────────────────

def load_esaf_mapping():
    wb = openpyxl.load_workbook(ESAF_FILE, read_only=True)
    ws = wb[SHEET_NAME]
    mapping = {}  # BRANCH_NAME_UPPER → {state, division, area}
    skipped = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:  # header
            continue
        branch = row[5]   # BRANCH NAME (new)
        state  = row[6]   # STATE
        area   = row[7]   # Area
        division = row[8] # DIVISION
        if not branch or not state or not division or not area:
            skipped += 1
            continue
        key = str(branch).strip().upper()
        mapping[key] = {
            'state':    str(state).strip().upper(),
            'division': str(division).strip().upper(),
            'area':     str(area).strip().upper(),
        }
    wb.close()
    print(f"ESAF mapping loaded: {len(mapping)} branches  (skipped {skipped} incomplete rows)")
    return mapping


# ── 2. Read existing data from EC2 ──────────────────────────────────────────

FETCH_QUERY = """
SELECT e.emp_id, e.officer_name, b.branch_name,
       ep.product_type_id,
       ep.regular_demand, ep.regular_collection,
       ep.demand_1_30, ep.collection_1_30,
       ep.demand_31_60, ep.collection_31_60,
       ep.pnpa_demand, ep.pnpa_collection,
       ep.npa_cases, ep.npa_act_acc, ep.npa_act_amt,
       ep.npa_clo_acc, ep.npa_clo_amt,
       ep.on_date_demand, ep.on_date_collection,
       ep.regular_demand_amt, ep.regular_collection_amt,
       ep.demand_1_30_amt, ep.collection_1_30_amt,
       ep.demand_31_60_amt, ep.collection_31_60_amt,
       ep.pnpa_demand_amt, ep.pnpa_collection_amt,
       ep.on_date_demand_amt, ep.on_date_collection_amt
FROM employee_performance ep
JOIN employees e ON ep.emp_id = e.emp_id
JOIN branches  b ON e.branch_id = b.branch_id
"""

OLD_TOTALS_QUERY = """
SELECT SUM(regular_demand), SUM(regular_collection),
       SUM(demand_1_30),    SUM(collection_1_30),
       SUM(demand_31_60),   SUM(collection_31_60),
       SUM(pnpa_demand),    SUM(pnpa_collection),
       COUNT(DISTINCT emp_id)
FROM employee_performance
"""

NEW_TOTALS_QUERY = """
SELECT SUM(regular_demand), SUM(regular_collection),
       SUM(demand_1_30),    SUM(collection_1_30),
       SUM(demand_31_60),   SUM(collection_31_60),
       SUM(pnpa_demand),    SUM(pnpa_collection),
       COUNT(DISTINCT emp_id)
FROM v2_employee_performance
"""

HIERARCHY_SUMMARY_QUERY = """
SELECT s.state_name,
       COUNT(DISTINCT dv.division_name) AS divisions,
       COUNT(DISTINCT a.area_name)      AS areas,
       COUNT(DISTINCT b.branch_name)    AS branches,
       COUNT(DISTINCT e.emp_id)         AS employees
FROM v2_states s
LEFT JOIN v2_divisions dv ON dv.state_id   = s.state_id
LEFT JOIN v2_areas     a  ON a.division_id = dv.division_id
LEFT JOIN v2_branches  b  ON b.area_id     = a.area_id
LEFT JOIN v2_employees e  ON e.branch_id   = b.branch_id
GROUP BY s.state_name
ORDER BY s.state_name
"""


def main():
    esaf = load_esaf_mapping()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = False
    cur = conn.cursor()

    # ── Fetch old data ──────────────────────────────────────────────────────
    print("\nFetching existing employee_performance rows …")
    cur.execute(FETCH_QUERY)
    rows = cur.fetchall()
    print(f"  Total rows fetched: {len(rows)}")

    # ── Map each row to new hierarchy ───────────────────────────────────────
    matched_branches   = set()
    unmatched_branches = set()
    good_rows = []

    for row in rows:
        emp_id, officer_name, branch_name, *rest = row
        key = str(branch_name).strip().upper()
        if key in esaf:
            matched_branches.add(key)
            info = esaf[key]
            good_rows.append((emp_id, officer_name, branch_name, info['state'], info['division'], info['area'], *rest))
        else:
            unmatched_branches.add(key)

    print(f"\n  Matched branches   : {len(matched_branches)}")
    print(f"  Unmatched branches : {len(unmatched_branches)}")
    if unmatched_branches:
        print("  WARNING — unmatched branches (skipped):")
        for b in sorted(unmatched_branches):
            print(f"    • {b}")
    print(f"  Rows to migrate    : {len(good_rows)}")

    # ── Capture old totals BEFORE truncate ──────────────────────────────────
    cur.execute(OLD_TOTALS_QUERY)
    old_totals = cur.fetchone()

    # ── Truncate v2 tables ──────────────────────────────────────────────────
    print("\nTruncating v2 tables …")
    cur.execute("TRUNCATE v2_employee_performance, v2_employees, v2_branches, v2_areas, v2_divisions, v2_states RESTART IDENTITY CASCADE")

    # ── Build lookup dicts for IDs ──────────────────────────────────────────
    states_map    = {}  # state_name → state_id
    divisions_map = {}  # (state_id, div_name) → div_id
    areas_map     = {}  # (div_id, area_name) → area_id
    branches_map  = {}  # (area_id, branch_name_upper) → branch_id
    employees_map = {}  # emp_id → new emp_id (same emp_id reused)

    # Collect unique combos
    for row in good_rows:
        emp_id, officer_name, branch_name, state, division, area, *_ = row

        # State
        if state not in states_map:
            cur.execute("INSERT INTO v2_states (state_name) VALUES (%s) RETURNING state_id", (state,))
            states_map[state] = cur.fetchone()[0]
        state_id = states_map[state]

        # Division
        div_key = (state_id, division)
        if div_key not in divisions_map:
            cur.execute("INSERT INTO v2_divisions (division_name, state_id) VALUES (%s, %s) RETURNING division_id",
                        (division, state_id))
            divisions_map[div_key] = cur.fetchone()[0]
        div_id = divisions_map[div_key]

        # Area
        area_key = (div_id, area)
        if area_key not in areas_map:
            cur.execute("INSERT INTO v2_areas (area_name, division_id) VALUES (%s, %s) RETURNING area_id",
                        (area, div_id))
            areas_map[area_key] = cur.fetchone()[0]
        area_id = areas_map[area_key]

        # Branch
        branch_key = (area_id, branch_name.strip().upper())
        if branch_key not in branches_map:
            cur.execute("INSERT INTO v2_branches (branch_name, area_id) VALUES (%s, %s) RETURNING branch_id",
                        (branch_name, area_id))
            branches_map[branch_key] = cur.fetchone()[0]
        branch_id = branches_map[branch_key]

        # Employee
        if emp_id not in employees_map:
            cur.execute("INSERT INTO v2_employees (emp_id, officer_name, branch_id) VALUES (%s, %s, %s) ON CONFLICT (emp_id) DO NOTHING",
                        (emp_id, officer_name, branch_id))
            employees_map[emp_id] = branch_id

    print(f"  States    inserted: {len(states_map)}")
    print(f"  Divisions inserted: {len(divisions_map)}")
    print(f"  Areas     inserted: {len(areas_map)}")
    print(f"  Branches  inserted: {len(branches_map)}")
    print(f"  Employees inserted: {len(employees_map)}")

    # ── Insert performance rows ─────────────────────────────────────────────
    print("\nInserting v2_employee_performance rows …")
    perf_rows = []
    for row in good_rows:
        (emp_id, officer_name, branch_name, state, division, area,
         product_type_id,
         regular_demand, regular_collection,
         demand_1_30, collection_1_30,
         demand_31_60, collection_31_60,
         pnpa_demand, pnpa_collection,
         npa_cases, npa_act_acc, npa_act_amt,
         npa_clo_acc, npa_clo_amt,
         on_date_demand, on_date_collection,
         regular_demand_amt, regular_collection_amt,
         demand_1_30_amt, collection_1_30_amt,
         demand_31_60_amt, collection_31_60_amt,
         pnpa_demand_amt, pnpa_collection_amt,
         on_date_demand_amt, on_date_collection_amt) = row

        perf_rows.append((
            emp_id, product_type_id,
            regular_demand, regular_collection,
            demand_1_30, collection_1_30,
            demand_31_60, collection_31_60,
            pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt,
            npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt,
            demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt,
            pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt,
        ))

    PERF_INSERT = """
        INSERT INTO v2_employee_performance (
            emp_id, product_type_id,
            regular_demand, regular_collection,
            demand_1_30, collection_1_30,
            demand_31_60, collection_31_60,
            pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt,
            npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt,
            demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt,
            pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt
        ) VALUES %s
    """
    execute_values(cur, PERF_INSERT, perf_rows, page_size=500)
    print(f"  Performance rows inserted: {len(perf_rows)}")

    conn.commit()

    # ── Dry-run comparison ──────────────────────────────────────────────────
    cur.execute(NEW_TOTALS_QUERY)
    new_totals = cur.fetchone()

    labels = [
        'regular_demand', 'regular_collection',
        'demand_1_30', 'collection_1_30',
        'demand_31_60', 'collection_31_60',
        'pnpa_demand', 'pnpa_collection',
        'DISTINCT emp_id (count)',
    ]

    print("\n" + "="*70)
    print("DRY RUN COMPARISON")
    print("="*70)
    print(f"{'Metric':<30}  {'OLD':>18}  {'NEW':>18}  {'MATCH'}")
    print("-"*70)
    all_match = True
    for label, old_val, new_val in zip(labels, old_totals, new_totals):
        match = "✓" if old_val == new_val else "✗ MISMATCH"
        if old_val != new_val:
            all_match = False
        print(f"{label:<30}  {str(old_val):>18}  {str(new_val):>18}  {match}")
    print("="*70)
    if all_match:
        print("ALL TOTALS MATCH — migration successful!\n")
    else:
        print("WARNING: Some totals do NOT match — investigate skipped rows.\n")

    # ── Hierarchy summary ───────────────────────────────────────────────────
    cur.execute(HIERARCHY_SUMMARY_QUERY)
    summary = cur.fetchall()
    print("NEW HIERARCHY SUMMARY")
    print(f"{'State':<20}  {'Divisions':>10}  {'Areas':>8}  {'Branches':>10}  {'Employees':>10}")
    print("-"*65)
    for r in summary:
        print(f"{r[0]:<20}  {r[1]:>10}  {r[2]:>8}  {r[3]:>10}  {r[4]:>10}")
    print("-"*65)
    total_div = sum(r[1] for r in summary)
    total_area = sum(r[2] for r in summary)
    total_br = sum(r[3] for r in summary)
    total_emp = sum(r[4] for r in summary)
    print(f"{'TOTAL':<20}  {total_div:>10}  {total_area:>8}  {total_br:>10}  {total_emp:>10}")

    cur.close()
    conn.close()
    print("\nDone.")


if __name__ == '__main__':
    main()
