# /hr_team Staff Sync — Design Spec

**Date:** 2026-06-08
**Status:** Approved design, pending spec review
**Author:** Claude + Raghunandan

## 1. Goal

A dedicated web page at `growwithme.navachetanalivelihoods.com/hr_team` where an HR user uploads
the staff-details Excel and, after reviewing a preview of the changes, syncs **both** databases to
match the file's "Working" roster:

- **EC2 PostgreSQL** `employee_master` (flat table) — hard-delete sync.
- **Supabase** `Grow_With_Me` (normalized HR warehouse) — soft-delete (status flip) + insert, history preserved.

This automates the manual two-DB sync run on 2026-06-08.

## 2. Non-Goals

- No change to the Flutter app or the main dashboard beyond adding one nav link.
- No editing of individual employees through this page (bulk file sync only).
- No scheduling/auto-pull — upload is user-initiated.

## 3. Placement & Infra

- Page: `~/Coll_Db/hr_team/index.html` (+ `hr_team/hr_team.css`, `hr_team/hr_team.js`), served by Apache
  at path `/hr_team` from web root `/var/www/html/coll-db/` (symlinked to `~/Coll_Db/`).
- API: two new routes on the **existing** Express server (`server/index.js`, port 3000, Apache-proxied).
- Nav: one admin-only link to `/hr_team` added to the dashboard.
- No new server/instance. Reuses EC2 box, `pg` pool, Apache proxy.

## 4. Access Control

Separate **HR passcode**, independent of the dashboard login.

- `.env`: `HR_SYNC_PASSCODE=<set by user>`.
- The page prompts for the passcode; it is sent with both API calls.
- Server validates the passcode on `preview` and `apply`; reject with 401 if wrong.
- Rate-limit apply (e.g. 5/min) to blunt brute force.

## 5. Credentials (server-side only, never sent to browser)

Added to `~/Coll_Db/.env`:

- `SUPABASE_URL=https://knbijsnghjcaocwtjvvw.supabase.co`
- `SUPABASE_SERVICE_KEY=<service_role key — user adds on EC2>`
- `HR_SYNC_PASSCODE=<user sets>`

EC2 PG creds already present (`PGHOST/PGPORT/PGUSER/PGPASSWORD/PGDATABASE`).

## 6. Data Flow (dry-run / apply split)

```
/hr_team page
  1. upload xlsx ─▶ POST /api/staff-sync/preview (passcode + file)
       server: parse → EC2 diff (pg) → Supabase diff (RPC hr_preview)
               → stash {file-derived plan} under token (15-min TTL, in-memory)
               → return diff JSON
     page renders preview:
       EC2:      +add  ✎update  ✖delete
       Supabase: ➕insert  ♻reactivate  ⏸deactivate
       ⚠ new master-data (branches/roles/designations) that will be auto-created
  2. [Confirm] ─▶ POST /api/staff-sync/apply (passcode + token)
       server: re-validate token → backup both → EC2 txn → Supabase RPC hr_apply (txn)
               → return summary
```

- `preview` performs **zero writes**.
- `apply` consumes the token once; a stale/expired token → 409, user re-previews.

## 7. Components (small, single-purpose, testable)

| File | Responsibility | Depends on |
|------|----------------|-----------|
| `server/lib/staffParser.js` | xlsx buffer → canonical Working rows. **Header auto-detect** (fixes current row-offset bug: detect the row containing `NMEmpId`/`Name`, data follows). Status map. Pure function. | `xlsx` |
| `server/lib/ec2Sync.js` | Compute diff + apply (upsert Working, delete absent) against `employee_master`. Transactional. | `pg` |
| `server/lib/supabaseSync.js` | Call `hr_preview` / `hr_apply` RPCs via `fetch` (service key, `Content-Profile: Grow_With_Me`). | `fetch` |
| `server/lib/staffBackup.js` | EC2: `CREATE TABLE employee_master_bak_<ts>`. Supabase backups handled inside `hr_apply`. | `pg`, RPC |
| `server/lib/syncTokens.js` | In-memory token store (uuid → plan + expiry). | — |
| routes in `server/index.js` | `/api/staff-sync/preview`, `/api/staff-sync/apply`. Passcode guard. | above |
| `hr_team/*` | Upload UI, passcode prompt, diff table, Confirm. | — |

## 8. Supabase RPC functions (transactional, in schema `Grow_With_Me`)

Created once via migration. Both take `p_working jsonb` (array of `{emp_id,name,gender,role,designation,branch,mobile,doj,status}`).

- `hr_preview(p_working jsonb) RETURNS jsonb` — read-only. Returns
  `{insert:[...], reactivate:[...], deactivate:[{code,to_status}], new_roles:[...], new_designations:[...], new_branches:[...]}`.
- `hr_apply(p_working jsonb) RETURNS jsonb` — `SECURITY DEFINER`, single transaction:
  1. snapshot `bak.employee_<ts>`, `bak.employee_assignment_<ts>`, `bak.employee_contact_<ts>`;
  2. auto-create missing `role` / `designation` / `branch` (case-insensitive match; typo-alias table for known typos: `Filed Officer→Field Officer`, `Executive-Operation IT→Execuitve - Operation IT`, `Office Assistant→Admin Assistant`);
  3. insert new employees (`employee` + `employee_assignment` + `employee_contact`, status active);
  4. reactivate matched-but-inactive → `status_id=1`;
  5. deactivate active-but-absent → mapped departed status;
  6. return counts.

Match key: `employee_code`. Branch→area→division→region derive via `branch_new_link` (assignment stores `branch_id`/`designation_id`/`role_id` only). New branches get next `branch_id` (no sequence).

## 9. Status Mapping

File status → Supabase `status_id`: `Working→1`, `Resigned→3`, `Abscond→6`, `Left→4`; anyone in DB but absent from file → `resigned(3)`. EC2 keeps the flat `status='Working'` for all kept rows (departed are hard-deleted).

## 10. Master-Data Handling

Auto-create unknown roles/designations/branches (user's choice), BUT the **preview lists every value that will be created** so the HR user can spot a typo and re-upload a corrected file before clicking Confirm. A small typo-alias table maps known misspellings to existing rows to reduce junk.

## 11. Safety

- Preview-before-apply (no blind writes).
- Auto-backup both DBs on every apply (timestamped restore points; EC2 table + Supabase `bak.*`).
- Both writes transactional — all-or-nothing per DB.
- Passcode gate + apply rate-limit.
- Token single-use + TTL prevents replay / stale-plan apply.
- EC2 confirmed: no FK refs to `employee_master` (safe delete). Supabase: hard delete impossible (RESTRICT/CASCADE FKs) — status-flip only.

## 12. Failure Handling

- Parse error / no Working sheet → 400 with message, no token issued.
- EC2 txn fails → rollback, abort before touching Supabase, report which stage failed.
- Supabase RPC fails → EC2 already committed; report partial state + that EC2 succeeded / Supabase rolled back. (EC2 first because it's the simpler, FK-free table; Supabase RPC is atomic on its own.)
- Wrong passcode → 401. Expired token → 409.

## 13. Testing

- **Unit:** `staffParser` against (a) this 2026 file (headers row 0) and (b) the old daily-report format (header row 1) — proves auto-detect; status mapping; EC2 diff math.
- **Integration:** `hr_preview`/`hr_apply` against a **throwaway Supabase branch**; EC2 sync against a scratch schema. Assert counts and that departed rows flip `is_working=false`, new rows appear in `v_employee_app`.
- **Manual:** dry-run preview on the real file, confirm counts match the 2026-06-08 run (insert 77 / reactivate 10 / deactivate 108 — will differ as data evolves).

## 14. Deploy

Per Coll_Db rules: commit + `git push origin master`, then rsync to EC2, then `pm2 restart all` (server changed). The Supabase RPC migration applied once via Supabase MCP/dashboard. User adds `SUPABASE_SERVICE_KEY` + `HR_SYNC_PASSCODE` to EC2 `.env` before restart.

## 15. Open Items

- Confirm EC2 Node version supports global `fetch` (Node ≥18). If not, add `node-fetch` or use `https`.
- Decide retention/cleanup of `*_bak_<ts>` tables (manual prune for now).
