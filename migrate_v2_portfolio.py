#!/usr/bin/env python3
"""
Migration script: Populate v2_* tables from portfolio_month_wise DB
for months APR through FEB (excluding MAR), using ESAF branch mapping.
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
    'database': 'portfolio_month_wise',
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
        branch   = row[5]   # Col 6 = BRANCH NAME
        state    = row[6]   # Col 7 = STATE
        area     = row[7]   # Col 8 = Area
        division = row[8]   # Col 9 = DIVISION
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


# ── 2. Main ─────────────────────────────────────────────────────────────────

def main():
    esaf = load_esaf_mapping()

    conn = psycopg2.connect(**DB_CONFIG)
    conn.autocommit = False
    cur = conn.cursor()

    # ── Get APR-FEB month_ids ───────────────────────────────────────────────
    cur.execute("SELECT month_id, month_label FROM months WHERE month_label != 'MAR' ORDER BY sort_order")
    apr_feb_months = cur.fetchall()
    apr_feb_month_ids = {r[0] for r in apr_feb_months}
    print(f"\nMonths to migrate: {[r[1] for r in apr_feb_months]}")

    # ── Capture OLD totals before truncate ──────────────────────────────────
    print("\nCapturing old totals …")
    cur.execute("""
        SELECT m.month_label, SUM(pp.regular_demand), SUM(pp.regular_collection), COUNT(DISTINCT pp.emp_id)
        FROM portfolio_performance pp
        JOIN months m ON pp.month_id = m.month_id
        WHERE m.month_label != 'MAR'
        GROUP BY m.month_label, m.sort_order ORDER BY m.sort_order
    """)
    old_monthly = cur.fetchall()

    # ── Truncate v2 tables ──────────────────────────────────────────────────
    print("\nTruncating v2 tables …")
    cur.execute("""
        TRUNCATE v2_employee_pos, v2_branch_pos, v2_fy_performance,
                 v2_portfolio_performance, v2_employees, v2_branches,
                 v2_areas, v2_divisions, v2_states
        RESTART IDENTITY CASCADE
    """)
    print("  Done.")

    # ── 3. Fetch portfolio_performance (APR-FEB) ────────────────────────────
    print("\nFetching portfolio_performance rows (APR-FEB) …")
    cur.execute("""
        SELECT pp.month_id, pp.emp_id, pp.product_type_id,
               pp.regular_demand, pp.regular_collection,
               pp.demand_1_30, pp.collection_1_30,
               pp.demand_31_60, pp.collection_31_60,
               pp.pnpa_demand, pp.pnpa_collection,
               pp.npa_cases, pp.npa_act_acc, pp.npa_act_amt,
               pp.npa_clo_acc, pp.npa_clo_amt,
               pp.on_date_demand, pp.on_date_collection,
               pp.regular_demand_amt, pp.regular_collection_amt,
               pp.demand_1_30_amt, pp.collection_1_30_amt,
               pp.demand_31_60_amt, pp.collection_31_60_amt,
               pp.pnpa_demand_amt, pp.pnpa_collection_amt,
               pp.on_date_demand_amt, pp.on_date_collection_amt,
               pp.regular_pos, pp.sma0_pos, pp.sma1_pos,
               pp.pnpa_pos, pp.npa_pos, pp.total_pos,
               e.officer_name, b.branch_name
        FROM portfolio_performance pp
        JOIN employees e ON pp.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
        JOIN months m ON pp.month_id = m.month_id
        WHERE m.month_label != 'MAR'
    """)
    pp_rows = cur.fetchall()
    print(f"  Fetched {len(pp_rows)} rows")

    # ── 4. Build hierarchy from all unique employees/branches seen ──────────
    # Collect all unique (emp_id, officer_name, branch_name) across all months
    unique_employees = {}  # emp_id → (officer_name, branch_name)
    for row in pp_rows:
        emp_id = row[1]
        officer_name = row[34]
        branch_name = row[35]
        if emp_id not in unique_employees:
            unique_employees[emp_id] = (officer_name, branch_name)

    print(f"\n  Unique employees: {len(unique_employees)}")

    # Map branches to ESAF hierarchy
    matched_branches = set()
    unmatched_branches = set()
    branch_to_esaf = {}  # branch_name_upper → {state, division, area}

    for emp_id, (officer_name, branch_name) in unique_employees.items():
        key = branch_name.strip().upper()
        if not key:
            unmatched_branches.add('(empty)')
            branch_to_esaf[key] = {'state': 'UNMAPPED', 'division': 'UNMAPPED', 'area': 'UNMAPPED'}
        elif key in esaf:
            matched_branches.add(key)
            branch_to_esaf[key] = esaf[key]
        else:
            unmatched_branches.add(key)

    print(f"  Matched branches   : {len(matched_branches)}")
    print(f"  Unmatched branches : {len(unmatched_branches)}")
    if unmatched_branches:
        print("  WARNING — unmatched branches (will use UNMAPPED):")
        for b in sorted(unmatched_branches):
            print(f"    • {b}")
            branch_to_esaf[b] = {'state': 'UNMAPPED', 'division': 'UNMAPPED', 'area': 'UNMAPPED'}

    # ── 5. Insert hierarchy (states → divisions → areas → branches → employees)
    states_map    = {}  # state_name → state_id
    divisions_map = {}  # (state_id, div_name) → div_id
    areas_map     = {}  # (div_id, area_name) → area_id
    branches_map  = {}  # (area_id, branch_name_upper) → branch_id
    employees_inserted = set()

    for emp_id, (officer_name, branch_name) in unique_employees.items():
        key = branch_name.strip().upper()
        info = branch_to_esaf[key]
        state = info['state']
        division = info['division']
        area = info['area']

        if state not in states_map:
            cur.execute("INSERT INTO v2_states (state_name) VALUES (%s) RETURNING state_id", (state,))
            states_map[state] = cur.fetchone()[0]
        state_id = states_map[state]

        div_key = (state_id, division)
        if div_key not in divisions_map:
            cur.execute("INSERT INTO v2_divisions (division_name, state_id) VALUES (%s, %s) RETURNING division_id",
                        (division, state_id))
            divisions_map[div_key] = cur.fetchone()[0]
        div_id = divisions_map[div_key]

        area_key = (div_id, area)
        if area_key not in areas_map:
            cur.execute("INSERT INTO v2_areas (area_name, division_id) VALUES (%s, %s) RETURNING area_id",
                        (area, div_id))
            areas_map[area_key] = cur.fetchone()[0]
        area_id = areas_map[area_key]

        display_branch = branch_name.strip() if branch_name.strip() else 'UNKNOWN'
        branch_key = (area_id, key if key else 'UNKNOWN')
        if branch_key not in branches_map:
            cur.execute("INSERT INTO v2_branches (branch_name, area_id) VALUES (%s, %s) RETURNING branch_id",
                        (display_branch, area_id))
            branches_map[branch_key] = cur.fetchone()[0]
        branch_id = branches_map[branch_key]

        if emp_id not in employees_inserted:
            cur.execute(
                "INSERT INTO v2_employees (emp_id, officer_name, branch_id) VALUES (%s, %s, %s) ON CONFLICT (emp_id) DO NOTHING",
                (emp_id, officer_name, branch_id)
            )
            employees_inserted.add(emp_id)

    print(f"\n  States    inserted: {len(states_map)}")
    print(f"  Divisions inserted: {len(divisions_map)}")
    print(f"  Areas     inserted: {len(areas_map)}")
    print(f"  Branches  inserted: {len(branches_map)}")
    print(f"  Employees inserted: {len(employees_inserted)}")

    # ── 6. Insert v2_portfolio_performance ──────────────────────────────────
    print("\nInserting v2_portfolio_performance …")
    perf_rows = []
    skipped_pp = 0
    for row in pp_rows:
        (month_id, emp_id, product_type_id,
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
         regular_pos, sma0_pos, sma1_pos,
         pnpa_pos, npa_pos, total_pos,
         officer_name, branch_name) = row

        key = branch_name.strip().upper()
        if key not in branch_to_esaf:
            skipped_pp += 1
            continue

        perf_rows.append((
            month_id, emp_id, product_type_id,
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
            regular_pos, sma0_pos, sma1_pos,
            pnpa_pos, npa_pos, total_pos,
        ))

    PERF_INSERT = """
        INSERT INTO v2_portfolio_performance (
            month_id, emp_id, product_type_id,
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
            regular_pos, sma0_pos, sma1_pos,
            pnpa_pos, npa_pos, total_pos
        ) VALUES %s
    """
    execute_values(cur, PERF_INSERT, perf_rows, page_size=500)
    print(f"  Inserted: {len(perf_rows)}  Skipped: {skipped_pp}")

    # ── 7. Migrate fy_performance ────────────────────────────────────────────
    print("\nMigrating fy_performance …")
    cur.execute("""
        SELECT fp.emp_id, fp.product_type_id,
               fp.regular_demand, fp.regular_collection,
               fp.demand_1_30, fp.collection_1_30,
               fp.demand_31_60, fp.collection_31_60,
               fp.pnpa_demand, fp.pnpa_collection,
               fp.npa_cases, fp.npa_act_acc, fp.npa_act_amt,
               fp.npa_clo_acc, fp.npa_clo_amt,
               fp.on_date_demand, fp.on_date_collection,
               fp.regular_demand_amt, fp.regular_collection_amt,
               fp.demand_1_30_amt, fp.collection_1_30_amt,
               fp.demand_31_60_amt, fp.collection_31_60_amt,
               fp.pnpa_demand_amt, fp.pnpa_collection_amt,
               fp.on_date_demand_amt, fp.on_date_collection_amt,
               fp.regular_pos, fp.sma0_pos, fp.sma1_pos,
               fp.pnpa_pos, fp.npa_pos, fp.total_pos,
               b.branch_name
        FROM fy_performance fp
        JOIN employees e ON fp.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
    """)
    fy_rows = cur.fetchall()
    print(f"  Source rows: {len(fy_rows)}")

    if fy_rows:
        fy_insert_rows = []
        for row in fy_rows:
            emp_id = row[0]
            branch_name = row[32]
            key = branch_name.strip().upper()
            if key not in branch_to_esaf:
                continue
            fy_insert_rows.append(row[:32])  # all cols except branch_name

        FY_INSERT = """
            INSERT INTO v2_fy_performance (
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
                regular_pos, sma0_pos, sma1_pos,
                pnpa_pos, npa_pos, total_pos
            ) VALUES %s
            ON CONFLICT (emp_id, product_type_id) DO NOTHING
        """
        execute_values(cur, FY_INSERT, fy_insert_rows, page_size=500)
        print(f"  Inserted: {len(fy_insert_rows)}")
    else:
        print("  No fy_performance rows to migrate (table is empty).")

    # ── 8. Migrate branch_pos (APR-FEB) ─────────────────────────────────────
    print("\nMigrating branch_pos (APR-FEB) …")
    cur.execute("""
        SELECT bp.month_id, bp.branch_name, bp.product_name,
               bp.regular_pos, bp.sma0_pos, bp.sma1_pos,
               bp.pnpa_pos, bp.npa_pos, bp.total_pos
        FROM branch_pos bp
        WHERE bp.month_id = ANY(%s)
    """, (list(apr_feb_month_ids),))
    bp_rows = cur.fetchall()
    print(f"  Source rows: {len(bp_rows)}")

    bp_insert_rows = []
    bp_unmatched = set()
    for row in bp_rows:
        month_id, branch_name, product_name, regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos = row
        key = branch_name.strip().upper()
        if key in branch_to_esaf:
            info = branch_to_esaf[key]
            bp_insert_rows.append((
                month_id,
                info['state'], info['division'], info['area'], branch_name,
                product_name,
                regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos
            ))
        else:
            bp_unmatched.add(key)

    if bp_unmatched:
        print(f"  WARNING — unmatched branches in branch_pos ({len(bp_unmatched)}):")
        for b in sorted(bp_unmatched):
            print(f"    • {b}")

    BP_INSERT = """
        INSERT INTO v2_branch_pos (
            month_id, state_name, division_name, area_name, branch_name,
            product_name, regular_pos, sma0_pos, sma1_pos,
            pnpa_pos, npa_pos, total_pos
        ) VALUES %s
        ON CONFLICT (month_id, state_name, division_name, area_name, branch_name, product_name) DO NOTHING
    """
    execute_values(cur, BP_INSERT, bp_insert_rows, page_size=500)
    print(f"  Inserted: {len(bp_insert_rows)}")

    # ── 9. Migrate employee_pos (APR-FEB) ───────────────────────────────────
    print("\nMigrating employee_pos (APR-FEB) …")
    cur.execute("""
        SELECT month_id, emp_id, regular_pos, sma0_pos, sma1_pos,
               pnpa_pos, npa_pos, total_pos
        FROM employee_pos
        WHERE month_id = ANY(%s)
    """, (list(apr_feb_month_ids),))
    ep_rows = cur.fetchall()
    print(f"  Source rows: {len(ep_rows)}")

    EP_INSERT = """
        INSERT INTO v2_employee_pos (
            month_id, emp_id, regular_pos, sma0_pos, sma1_pos,
            pnpa_pos, npa_pos, total_pos
        ) VALUES %s
        ON CONFLICT (month_id, emp_id) DO NOTHING
    """
    execute_values(cur, EP_INSERT, ep_rows, page_size=500)
    print(f"  Inserted: {len(ep_rows)}")

    conn.commit()
    print("\nCommit successful.")

    # ── 10. DRY RUN COMPARISON ──────────────────────────────────────────────
    cur.execute("""
        SELECT m.month_label, SUM(pp.regular_demand), SUM(pp.regular_collection), COUNT(DISTINCT pp.emp_id)
        FROM v2_portfolio_performance pp
        JOIN months m ON pp.month_id = m.month_id
        GROUP BY m.month_label, m.sort_order ORDER BY m.sort_order
    """)
    new_monthly = cur.fetchall()

    print("\n" + "="*80)
    print("DRY RUN COMPARISON — portfolio_performance (APR-FEB)")
    print("="*80)
    print(f"{'Month':<6}  {'OLD Demand':>14}  {'OLD Coll':>14}  {'OLD Emp':>8}  {'NEW Demand':>14}  {'NEW Coll':>14}  {'NEW Emp':>8}  MATCH")
    print("-"*100)

    old_map = {r[0]: r for r in old_monthly}
    new_map = {r[0]: r for r in new_monthly}
    all_match = True

    for month_label in [r[1] for r in apr_feb_months]:
        old = old_map.get(month_label, (month_label, 0, 0, 0))
        new = new_map.get(month_label, (month_label, 0, 0, 0))
        match = "✓" if (old[1] == new[1] and old[2] == new[2] and old[3] == new[3]) else "✗ MISMATCH"
        if "MISMATCH" in match:
            all_match = False
        print(f"{month_label:<6}  {str(old[1]):>14}  {str(old[2]):>14}  {str(old[3]):>8}  {str(new[1]):>14}  {str(new[2]):>14}  {str(new[3]):>8}  {match}")

    print("="*80)
    if all_match:
        print("ALL MONTHLY TOTALS MATCH — migration successful!\n")
    else:
        print("WARNING: Some totals do NOT match — investigate.\n")

    # ── 11. Hierarchy summary ────────────────────────────────────────────────
    cur.execute("""
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
        GROUP BY s.state_name ORDER BY s.state_name
    """)
    summary = cur.fetchall()
    print("NEW HIERARCHY SUMMARY")
    print(f"{'State':<25}  {'Divisions':>10}  {'Areas':>8}  {'Branches':>10}  {'Employees':>10}")
    print("-"*70)
    for r in summary:
        print(f"{r[0]:<25}  {r[1]:>10}  {r[2]:>8}  {r[3]:>10}  {r[4]:>10}")
    print("-"*70)
    print(f"{'TOTAL':<25}  {sum(r[1] for r in summary):>10}  {sum(r[2] for r in summary):>8}  {sum(r[3] for r in summary):>10}  {sum(r[4] for r in summary):>10}")

    cur.close()
    conn.close()
    print("\nDone.")


if __name__ == '__main__':
    main()
