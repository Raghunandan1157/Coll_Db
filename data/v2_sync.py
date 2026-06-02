#!/usr/bin/env python3
"""
Standalone v2 portfolio table sync.
Mirrors populateV2PortfolioTables in server/index.js.
"""
import argparse
import json
import sys
import psycopg2
from psycopg2.extras import execute_values

with open("/tmp/branch_v2_map.json") as f:
    BRANCH_V2_MAP = json.load(f)

UNMAPPED = {"state": "UNMAPPED", "division": "UNMAPPED", "area": "UNMAPPED"}


def sync(conn):
    cur = conn.cursor()
    cur.execute("BEGIN")
    try:
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
        """)
        rows = cur.fetchall()
        print(f"  source rows: {len(rows):,}", flush=True)

        cur.execute("""
            TRUNCATE v2_employee_pos, v2_branch_pos, v2_fy_performance,
                     v2_portfolio_performance, v2_employees, v2_branches,
                     v2_areas, v2_divisions, v2_states
            RESTART IDENTITY CASCADE
        """)

        # Unique employee → (officer_name, branch_name)
        unique_emp = {}
        for r in rows:
            emp_id, officer, branch = r[1], r[34], r[35]
            if emp_id not in unique_emp:
                unique_emp[emp_id] = (officer, branch)

        states = {}; divisions = {}; areas = {}; branches = {}
        emp_to_v2branch = {}

        for emp_id, (officer, branch) in unique_emp.items():
            key = (branch or "").strip().upper()
            info = BRANCH_V2_MAP.get(key, UNMAPPED)

            if info["state"] not in states:
                cur.execute("INSERT INTO v2_states (state_name) VALUES (%s) RETURNING state_id", (info["state"],))
                states[info["state"]] = cur.fetchone()[0]
            sid = states[info["state"]]

            dkey = (sid, info["division"])
            if dkey not in divisions:
                cur.execute("INSERT INTO v2_divisions (division_name, state_id) VALUES (%s,%s) RETURNING division_id",
                            (info["division"], sid))
                divisions[dkey] = cur.fetchone()[0]
            did = divisions[dkey]

            akey = (did, info["area"])
            if akey not in areas:
                cur.execute("INSERT INTO v2_areas (area_name, division_id) VALUES (%s,%s) RETURNING area_id",
                            (info["area"], did))
                areas[akey] = cur.fetchone()[0]
            aid = areas[akey]

            bkey = (aid, key)
            if bkey not in branches:
                cur.execute("INSERT INTO v2_branches (branch_name, area_id) VALUES (%s,%s) RETURNING branch_id",
                            ((branch or "").strip() or "UNKNOWN", aid))
                branches[bkey] = cur.fetchone()[0]
            bid = branches[bkey]
            emp_to_v2branch[emp_id] = bid

        # v2_employees
        emp_rows = [(emp, ofc, emp_to_v2branch[emp]) for emp, (ofc, _) in unique_emp.items()]
        execute_values(cur,
            "INSERT INTO v2_employees (emp_id, officer_name, branch_id) VALUES %s ON CONFLICT (emp_id) DO NOTHING",
            emp_rows, page_size=200)
        print(f"  v2_employees: {len(emp_rows)}", flush=True)

        # v2_portfolio_performance — bulk insert all rows (drop officer/branch tail cols)
        pp_rows = [tuple(r[:34]) for r in rows]
        execute_values(cur,
            """INSERT INTO v2_portfolio_performance (month_id, emp_id, product_type_id,
                regular_demand, regular_collection, demand_1_30, collection_1_30,
                demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
                npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
                on_date_demand, on_date_collection,
                regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
                demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
                on_date_demand_amt, on_date_collection_amt,
                regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
               VALUES %s""", pp_rows, page_size=200)
        print(f"  v2_portfolio_performance: {len(pp_rows)}", flush=True)

        # v2_fy_performance
        cur.execute("""
            SELECT fp.emp_id, fp.product_type_id,
                fp.regular_demand, fp.regular_collection, fp.demand_1_30, fp.collection_1_30,
                fp.demand_31_60, fp.collection_31_60, fp.pnpa_demand, fp.pnpa_collection,
                fp.npa_cases, fp.npa_act_acc, fp.npa_act_amt, fp.npa_clo_acc, fp.npa_clo_amt,
                fp.on_date_demand, fp.on_date_collection,
                fp.regular_demand_amt, fp.regular_collection_amt, fp.demand_1_30_amt, fp.collection_1_30_amt,
                fp.demand_31_60_amt, fp.collection_31_60_amt, fp.pnpa_demand_amt, fp.pnpa_collection_amt,
                fp.on_date_demand_amt, fp.on_date_collection_amt,
                fp.regular_pos, fp.sma0_pos, fp.sma1_pos, fp.pnpa_pos, fp.npa_pos, fp.total_pos
            FROM fy_performance fp WHERE fp.emp_id IN (SELECT emp_id FROM v2_employees)
        """)
        fy_rows = cur.fetchall()
        if fy_rows:
            execute_values(cur,
                """INSERT INTO v2_fy_performance (emp_id, product_type_id,
                    regular_demand, regular_collection, demand_1_30, collection_1_30,
                    demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
                    npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
                    on_date_demand, on_date_collection,
                    regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
                    demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
                    on_date_demand_amt, on_date_collection_amt,
                    regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
                   VALUES %s ON CONFLICT (emp_id, product_type_id) DO NOTHING""", fy_rows, page_size=200)
        print(f"  v2_fy_performance: {len(fy_rows)}", flush=True)

        # v2_branch_pos
        cur.execute("""
            SELECT bp.month_id, bp.branch_name, bp.product_name,
                bp.regular_pos, bp.sma0_pos, bp.sma1_pos, bp.pnpa_pos, bp.npa_pos, bp.total_pos
            FROM branch_pos bp
        """)
        bp_src = cur.fetchall()
        bp_rows = []
        for r in bp_src:
            month_id, bname, pname = r[0], r[1], r[2]
            key = (bname or "").strip().upper()
            info = BRANCH_V2_MAP.get(key, UNMAPPED)
            bp_rows.append((month_id, info["state"], info["division"], info["area"],
                            bname, pname, *r[3:]))
        if bp_rows:
            execute_values(cur,
                """INSERT INTO v2_branch_pos (month_id, state_name, division_name, area_name, branch_name,
                    product_name, regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
                   VALUES %s ON CONFLICT (month_id, state_name, division_name, area_name, branch_name, product_name) DO NOTHING""",
                bp_rows, page_size=200)
        print(f"  v2_branch_pos: {len(bp_rows)}", flush=True)

        # v2_employee_pos
        cur.execute("""
            SELECT ep.month_id, ep.emp_id, ep.regular_pos, ep.sma0_pos, ep.sma1_pos,
                ep.pnpa_pos, ep.npa_pos, ep.total_pos
            FROM employee_pos ep WHERE ep.emp_id IN (SELECT emp_id FROM v2_employees)
        """)
        ep_rows = cur.fetchall()
        if ep_rows:
            execute_values(cur,
                """INSERT INTO v2_employee_pos (month_id, emp_id, regular_pos, sma0_pos, sma1_pos,
                    pnpa_pos, npa_pos, total_pos)
                   VALUES %s ON CONFLICT (month_id, emp_id) DO NOTHING""", ep_rows, page_size=200)
        print(f"  v2_employee_pos: {len(ep_rows)}", flush=True)

        cur.execute("COMMIT")
        print("\nCOMMITTED.", flush=True)

        # Verify MAR totals
        cur.execute("SELECT ROUND(SUM(total_pos)/10000000,2) FROM v2_portfolio_performance WHERE month_id=24")
        print(f"\nv2_portfolio_performance MAR total = {cur.fetchone()[0]} Cr")
        cur.execute("SELECT ROUND(SUM(total_pos)/10000000,2) FROM v2_branch_pos WHERE month_id=24")
        print(f"v2_branch_pos MAR total = {cur.fetchone()[0]} Cr")
        cur.execute("SELECT ROUND(SUM(total_pos)/10000000,2) FROM v2_employee_pos WHERE month_id=24")
        print(f"v2_employee_pos MAR total = {cur.fetchone()[0]} Cr")

    except Exception as e:
        cur.execute("ROLLBACK")
        print(f"ROLLBACK: {e}", file=sys.stderr)
        raise
    finally:
        cur.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=5432)
    args = ap.parse_args()
    conn = psycopg2.connect(host=args.host, port=args.port, user="Raghunandan1157",
                            password="raghu", database="portfolio_month_wise")
    sync(conn)
    conn.close()
