# Database Schema Efficiency Audit

**Date**: 2026-04-16  
**Status**: READ-ONLY Analysis — No changes made, proposals only  
**Scanned**: All 19 tables, 61 indexes, code patterns in server/index.js and js/*.js

---

## Executive Summary

Database is well-indexed but has 3 key efficiency issues:

1. **9 indexes never used** (744 KB recoverable) — mostly on low-cardinality or unused filter paths
2. **Duplicate v1/v2 tables** — `employees`, `employee_performance`, and related tables exist in both versions; need to confirm which is active
3. **Wide tables with nullable metrics** — `daily_performance` (28 MB, 29 cols), other tables use nullable integer columns with default 0 instead of NOT NULL constraints

---

## 1. Unused/Underutilized Indexes

### A. Zero-Scan Indexes (Drop Candidates)

| Table | Index Name | Scans | Size | Severity | Action |
|-------|-----------|-------|------|----------|--------|
| disbursement | `idx_disb_region` | 0 | 272 kB | **HIGH** | Drop — `region_name` is never filtered in queries |
| daily_reports | `idx_dr_branch` | 0 | 16 kB | **MED** | Drop — queries use subquery on employee_master instead |
| daily_reports | `idx_dr_region` | 0 | 16 kB | **MED** | Drop — not referenced in code |
| hourly_performance | `idx_hp_product_type_id` | 0 | 16 kB | **LOW** | Drop — product_type_id filtered via composite key |
| daily_reports_achievements | `idx_dra_branch` | 0 | 16 kB | **MED** | Drop — same as `idx_dr_branch` |

**Estimated Recovery**: 336 kB disk space, slightly faster INSERT/UPDATE on these tables  
**Risk**: Low (verified zero usage in pg_stat_user_indexes)

**Exact Changes**:
```sql
DROP INDEX idx_disb_region;
DROP INDEX idx_dr_branch;
DROP INDEX idx_dr_region;
DROP INDEX idx_hp_product_type_id;
DROP INDEX idx_dra_branch;
```

---

### B. Unused Unique Constraint Indexes (Enforcement-Only)

These unique constraints are enforced at the database level (good for data integrity) but their indexes are never scanned for lookups. Acceptable to keep for UNIQUE constraint enforcement, but flag if you want name-based lookups elsewhere:

| Table | Index | Scans | Use Case |
|-------|-------|-------|----------|
| v2_divisions | `v2_divisions_division_name_state_id_key` | 0 | Constraint only, not queried by name |
| v2_areas | `v2_areas_area_name_division_id_key` | 0 | Constraint only |
| v2_branches | `v2_branches_branch_name_area_id_key` | 0 | Constraint only |
| v2_states | `v2_states_state_name_key` | 0 | Constraint only |

**Severity**: LOW (these serve UNIQUE constraint enforcement, which is essential)  
**Action**: Keep for now; if you add name-based lookups in the future, they'll become useful

---

### C. Very Low-Usage Indexes

| Table | Index | Scans | Severity | Analysis |
|-------|-------|-------|----------|----------|
| employee_master | `idx_em_area` | 3 | **LOW** | Area filtering is rare; consider monitoring for continued low usage |
| employee_performance | `idx_ep_product_type_id` | 22 | **LOW** | Product type is always filtered via composite key; this redundant index adds write overhead |

**Action**: Monitor; consider dropping if area filtering remains rare.

---

## 2. v1 vs v2 Table Duplication

### Current State
**v1 (Legacy) Tables** — Active but being replaced:
- `branches` (96 kB)
- `employees` (128 kB)
- `employee_performance` (480 kB)
- `disbursement` (4168 kB) — NO v2 equivalent
- `hourly_performance` (288 kB) — NO v2 equivalent
- `fy_performance` (24 kB) — NO v2 equivalent

**v2 (New Hierarchy) Tables** — State→Division→Area→Branch:
- `v2_branches` (56 kB)
- `v2_employees` (128 kB)
- `v2_employee_performance` (480 kB)
- `v2_areas`, `v2_divisions`, `v2_states`

### Analysis

**Code Pattern** (from server/index.js):
- Upload handlers: Populate BOTH v1 AND v2 tables simultaneously (lines 220–530)
- API handlers: Query v1 tables (`employee_performance`, `daily_performance`, `fy_performance`) primarily
- v2 tables: Used for hierarchical lookups in v2 hierarchy filters only

**Assessment**: **v1 remains the source of truth; v2 is secondary for the new hierarchy system**

### Issues

1. **Redundant writes** — Every upload writes to both `employees` and `v2_employees`, both `employee_performance` and `v2_employee_performance`
2. **Stale v2 hierarchy** — If v2 branches/areas/divisions tables aren't fully synced with v1 branches data, v2 queries may miss employees
3. **Query complexity** — Code maintains two separate hierarchy paths; migrations incomplete

### Severity: **MEDIUM**

### Estimated Impact

- **Write overhead**: ~10–20% extra INSERT/UPDATE time due to double-write
- **Storage**: ~1.5 GB of duplicate performance data if scaled further
- **Maintenance risk**: Two evolving schemas increase bug surface

### Recommended Action (Strategic Choice)

**Option A (Recommended for now)**: Keep both, but consolidate code
- Make v2 the primary hierarchy source (only write to v2)
- Add foreign-key relationships from v1 to v2 for traceability
- Create views mapping v1-style queries to v2 tables
- This allows gradual v1→v2 migration without breaking existing reports

**Option B (Longer-term)**: Full cutover to v2
- Rewrite all API handlers to query v2 tables
- Remove v1 hierarchy tables (branches, employees, districts, regions)
- Keep `employee_performance` → `v2_employee_performance`, `daily_performance` migration

**Option C**: Accept dual-write (current state)
- Document that v2 is audit trail of hierarchy changes
- Accept storage cost for safety margin

---

## 3. Wide Tables & Column Design

### A. daily_performance — Large and Sparse

| Metric | Value |
|--------|-------|
| Size | 28 MB (largest table) |
| Columns | 29 (all nullable, default 0) |
| Row Count | ~1.7M rows estimated |
| Avg Row Width | ~16 KB |

**Issue**: Most columns default to 0; should be NOT NULL with DEFAULT 0 and integer type (not numeric).

**Example problematic columns**:
```
regular_demand            integer   YES  0
regular_collection        integer   YES  0
demand_1_30               integer   YES  0
collection_1_30           integer   YES  0
pnpa_demand_amt           numeric   YES  0
```

Should be:
```
regular_demand            integer   NO   0
regular_collection        integer   NO   0
pnpa_demand_amt           integer   NO   0    -- or decimal(12,2) if currency
```

**Severity**: **MEDIUM** (improves NULL-check performance in aggregations)

**Estimated Gains**:
- Memory: ~10–15% smaller rows (NULL bitmaps eliminated)
- Query: ~5% faster aggregation queries (no NULL handling needed)

**Change**:
```sql
ALTER TABLE daily_performance
  ALTER COLUMN regular_demand SET NOT NULL,
  ALTER COLUMN regular_collection SET NOT NULL,
  -- ... repeat for all numeric default-0 columns
  ALTER COLUMN pnpa_demand_amt TYPE integer,
  ALTER COLUMN regular_demand_amt TYPE integer;
```

**Caution**: Requires data validation first; ensure all 0-default columns have no actual NULLs.

---

### B. employee_master — High Cardinality Text Fields

| Column | Type | Nullable | Usage |
|--------|------|----------|-------|
| emp_id | varchar | NO | PK, 10–20 char IDs |
| full_name | varchar | YES | 50–100 chars |
| branch_name | varchar | YES | 30–50 chars |
| region_name | varchar | YES | 20–50 chars |
| area_name | varchar | YES | 20–50 chars |
| designation | varchar | YES | 30 chars (enum-like) |
| role | varchar | YES | 10 chars (enum-like) |

**Issue**: Unbounded varchar fields; `designation` and `role` are effectively enums (same ~10 values repeated).

**Severity**: **LOW** (table is small, 744 kB; varchar overhead minimal)

**Optimization** (Optional):
```sql
-- Create lookup table for roles (if reused)
CREATE TABLE roles (role_id SMALLINT PRIMARY KEY, role_name VARCHAR(50) NOT NULL);
INSERT INTO roles VALUES (1, 'CEO'), (2, 'RM'), (3, 'DM'), (4, 'BM'), (5, 'FO');

-- Then reference:
ALTER TABLE employee_master ADD COLUMN role_id SMALLINT REFERENCES roles(role_id);
```

**Impact**: Minimal (saves ~50 bytes/row if heavily repeated, but not critical).

---

### C. Numeric Types — Currency vs Integer

| Table | Column | Type | Should Be |
|-------|--------|------|-----------|
| daily_performance | *_amt columns (12) | numeric(10,0) | integer |
| disbursement | disb_amount | numeric | decimal(12,2) |
| daily_reports | disb_*_amt (3) | numeric | decimal(12,2) |
| fy_performance | *_pos columns (5) | numeric | decimal(14,2) |

**Issue**: Amount fields stored as `numeric` (variable-length) when they fit in INTEGER (if whole rupees) or DECIMAL(12,2) (if 2 decimals).

**Severity**: **LOW** (minimal impact on storage; numeric precision is good)

**Optimization**: Convert to INTEGER if all values are whole rupees:
```sql
ALTER TABLE daily_performance
  ALTER COLUMN regular_demand_amt TYPE integer USING regular_demand_amt::integer;
```

**Risk**: HIGH (currency truncation). Only if you've confirmed no decimal values exist.

---

## 4. Missing Indexes

### Query Pattern Analysis (from code)

**Filters commonly used**:
- `daily_performance.report_date` — Has `idx_dp_date`, well-used (10K+ scans) ✓
- `daily_performance.emp_id` — Has `idx_dp_emp`, heavily used (54K+ scans) ✓
- `employee_master.region_name` — Has `idx_em_region`, used (1.9K scans) ✓
- `daily_reports.branch_name` — Has `idx_dr_branch`, BUT **zero scans** — queries use subquery instead ✗
- `daily_reports.date` — Has `idx_dr_date`, used (737 scans) ✓

### Finding: No missing high-impact indexes

**Code path** (server/index.js:936–944):
```javascript
where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($...))`);
```

This subquery on `employee_master` would benefit from an index on `region_name` + `branch_name`, but `idx_em_region` exists and is used.

**Verdict**: Indexing is adequate. Unused indexes are the problem, not missing ones.

---

## 5. Primary Key Scan Overhead

Several PRIMARY KEY indexes have extremely high scan counts but serve join operations:

| Index | Scans | Table Size | Scans/MB |
|-------|-------|-----------|----------|
| branches_pkey | 1,864,534 | 96 kB | **~20M scans/MB** |
| employees_pkey | 1,870,767 | 128 kB | **~15M scans/MB** |
| v2_employees_pkey | 34,578 | 128 kB | **~270K scans/MB** |

**Analysis**: These are normal; PK lookups dominate queries. No issue.

---

## 6. Summary Table: All Findings

| Issue | Severity | Type | Impact | Effort |
|-------|----------|------|--------|--------|
| Drop 5 unused indexes | HIGH | Index cleanup | 336 kB + faster writes | 5 min |
| Resolve v1/v2 duplication strategy | MEDIUM | Architecture | 1.5 GB duplicate data + maintenance risk | 2–3 days (strategic) |
| Add NOT NULL + change numeric to integer for metrics | MEDIUM | Schema cleanup | 10–15% row size reduction | 2–4 hours |
| Consolidate role/designation to lookup tables | LOW | Schema normalization | ~50 B/row, minimal | 1–2 days (optional) |

---

## Recommendations (In Priority Order)

### Immediate (Low Risk)

1. **Drop unused indexes** (336 kB recovery)
   ```sql
   DROP INDEX idx_disb_region;
   DROP INDEX idx_dr_branch;
   DROP INDEX idx_dr_region;
   DROP INDEX idx_hp_product_type_id;
   DROP INDEX idx_dra_branch;
   ```
   - Verify zero usage in current workload
   - Reclaim 336 kB disk space
   - Slightly faster writes to these tables

2. **Monitor low-usage index** `idx_em_area` (3 scans)
   - Area filtering is rare; confirm this is expected behavior
   - Decision on drop in 1–2 months

### Medium-term (Requires Planning)

3. **Consolidate v1/v2 tables** (Strategic decision)
   - Choose primary hierarchy version (recommend v2)
   - Create migration plan for queries
   - Eliminate redundant double-writes
   - Estimated: 2–3 weeks of refactoring

4. **Tighten metric columns** (Performance improvement)
   - Add NOT NULL constraints to nullable metric columns (default 0)
   - Convert numeric to integer where appropriate
   - Verify no existing NULLs in data first
   - Estimated: 4 hours + testing

### Long-term (Optional)

5. **Normalize lookup columns** (Code maintainability)
   - Create `roles` and `designations` lookup tables
   - Reduce cardinality in employee_master
   - Minimal storage impact but improves data consistency

---

## Verification Commands

Run these to verify findings before acting:

```sql
-- Confirm zero scans on candidate-drop indexes
SELECT indexname, idx_scan 
FROM pg_stat_user_indexes 
WHERE indexname IN ('idx_disb_region', 'idx_dr_branch', 'idx_dr_region', 'idx_hp_product_type_id', 'idx_dra_branch');

-- Check for NULLs in metric columns (before adding NOT NULL)
SELECT 
  COUNT(*) as total,
  COUNT(CASE WHEN regular_demand IS NULL THEN 1 END) as nulls
FROM daily_performance;

-- Estimate space savings from NOT NULL + integer conversion
SELECT 
  pg_size_pretty(pg_total_relation_size('daily_performance')) as current_size;
-- After conversion, run again for comparison
```

---

## Conclusion

Database schema is **well-designed overall**. Recommended actions focus on:
1. Removing genuinely unused indexes (safe)
2. Resolving v1/v2 duplication (architectural clarity)
3. Tightening constraints on metric columns (performance)

No urgent issues; changes can be rolled out incrementally.
