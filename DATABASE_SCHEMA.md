# Database Schema — postgres (EC2 PostgreSQL)

## Overview
19 user tables in `public` schema. Organized by domain: hierarchy, performance tracking, daily operations, and v2 (new structure).

---

## Hierarchy & Reference Tables

### regions
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| region_id | integer | NO | nextval('regions_region_id_seq'::regclass) | ✓ |
| region_name | character varying | NO | | |

### districts
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| district_id | integer | NO | nextval('districts_district_id_seq'::regclass) | ✓ |
| district_name | character varying | NO | | |
| region_id | integer | NO | | |

### branches
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| branch_id | integer | NO | nextval('branches_branch_id_seq'::regclass) | ✓ |
| branch_name | character varying | NO | | |
| district_id | integer | NO | | |

### employees
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| emp_id | character varying | NO | | ✓ |
| officer_name | character varying | NO | | |
| branch_id | integer | NO | | |

### employee_master
Employee directory with hierarchy, role, and contact details.
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| emp_id | character varying | NO | | ✓ |
| full_name | character varying | YES | | |
| role | character varying | YES | | |
| designation | character varying | YES | | |
| branch_name | character varying | YES | | |
| area_name | character varying | YES | | |
| area_manager | character varying | YES | | |
| division_name | character varying | YES | | |
| division_manager | character varying | YES | | |
| region_name | character varying | YES | | |
| mobile | character varying | YES | | |
| date_of_joining | date | YES | | |
| reporting_officer_id | character varying | YES | | |
| reporting_officer_name | character varying | YES | | |
| status | character varying | YES | 'Working'::character varying | |

### product_types
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| product_type_id | integer | NO | nextval('product_types_product_type_id_seq'::regclass) | ✓ |
| product_type_name | character varying | NO | | |

---

## Performance Tracking — Month/FY Level

### employee_performance
Aggregated employee performance metrics (demand, collection, NPA tracking).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| performance_id | integer | NO | nextval('employee_performance_performance_id_seq'::regclass) | ✓ |
| emp_id | character varying | NO | | |
| product_type_id | integer | NO | | |
| regular_demand | integer | YES | 0 | |
| regular_collection | integer | YES | 0 | |
| demand_1_30 | integer | YES | 0 | |
| collection_1_30 | integer | YES | 0 | |
| demand_31_60 | integer | YES | 0 | |
| collection_31_60 | integer | YES | 0 | |
| pnpa_demand | integer | YES | 0 | |
| pnpa_collection | integer | YES | 0 | |
| npa_cases | integer | YES | 0 | |
| npa_act_acc | integer | YES | 0 | |
| npa_act_amt | numeric | YES | 0 | |
| npa_clo_acc | integer | YES | 0 | |
| npa_clo_amt | numeric | YES | 0 | |
| on_date_demand | integer | YES | 0 | |
| on_date_collection | integer | YES | 0 | |
| regular_demand_amt | numeric | YES | 0 | |
| regular_collection_amt | numeric | YES | 0 | |
| demand_1_30_amt | numeric | YES | 0 | |
| collection_1_30_amt | numeric | YES | 0 | |
| demand_31_60_amt | numeric | YES | 0 | |
| collection_31_60_amt | numeric | YES | 0 | |
| pnpa_demand_amt | numeric | YES | 0 | |
| pnpa_collection_amt | numeric | YES | 0 | |
| on_date_demand_amt | numeric | YES | 0 | |
| on_date_collection_amt | numeric | YES | 0 | |

### fy_performance
Fiscal year performance with POS (Portfolio Outstanding) tracking.
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| performance_id | integer | NO | nextval('fy_performance_performance_id_seq'::regclass) | ✓ |
| emp_id | character varying | NO | | |
| product_type_id | integer | NO | | |
| regular_demand | integer | YES | 0 | |
| regular_collection | integer | YES | 0 | |
| demand_1_30 | integer | YES | 0 | |
| collection_1_30 | integer | YES | 0 | |
| demand_31_60 | integer | YES | 0 | |
| collection_31_60 | integer | YES | 0 | |
| pnpa_demand | integer | YES | 0 | |
| pnpa_collection | integer | YES | 0 | |
| npa_cases | integer | YES | 0 | |
| npa_act_acc | integer | YES | 0 | |
| npa_act_amt | numeric | YES | 0 | |
| npa_clo_acc | integer | YES | 0 | |
| npa_clo_amt | numeric | YES | 0 | |
| on_date_demand | integer | YES | 0 | |
| on_date_collection | integer | YES | 0 | |
| regular_demand_amt | numeric | YES | 0 | |
| regular_collection_amt | numeric | YES | 0 | |
| demand_1_30_amt | numeric | YES | 0 | |
| collection_1_30_amt | numeric | YES | 0 | |
| demand_31_60_amt | numeric | YES | 0 | |
| collection_31_60_amt | numeric | YES | 0 | |
| pnpa_demand_amt | numeric | YES | 0 | |
| pnpa_collection_amt | numeric | YES | 0 | |
| on_date_demand_amt | numeric | YES | 0 | |
| on_date_collection_amt | numeric | YES | 0 | |
| regular_pos | numeric | YES | 0 | |
| sma0_pos | numeric | YES | 0 | |
| sma1_pos | numeric | YES | 0 | |
| pnpa_pos | numeric | YES | 0 | |
| npa_pos | numeric | YES | 0 | |
| total_pos | numeric | YES | 0 | |

### hourly_performance
Intra-day performance tracking (same metric structure as employee_performance).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| performance_id | integer | NO | nextval('hourly_performance_performance_id_seq'::regclass) | ✓ |
| emp_id | character varying | NO | | |
| product_type_id | integer | NO | | |
| regular_demand | integer | YES | 0 | |
| regular_collection | integer | YES | 0 | |
| demand_1_30 | integer | YES | 0 | |
| collection_1_30 | integer | YES | 0 | |
| demand_31_60 | integer | YES | 0 | |
| collection_31_60 | integer | YES | 0 | |
| pnpa_demand | integer | YES | 0 | |
| pnpa_collection | integer | YES | 0 | |
| npa_cases | integer | YES | 0 | |
| npa_act_acc | integer | YES | 0 | |
| npa_act_amt | numeric | YES | 0 | |
| npa_clo_acc | integer | YES | 0 | |
| npa_clo_amt | numeric | YES | 0 | |
| on_date_demand | integer | YES | 0 | |
| on_date_collection | integer | YES | 0 | |
| regular_demand_amt | numeric | YES | 0 | |
| regular_collection_amt | numeric | YES | 0 | |
| demand_1_30_amt | numeric | YES | 0 | |
| collection_1_30_amt | numeric | YES | 0 | |
| demand_31_60_amt | numeric | YES | 0 | |
| collection_31_60_amt | numeric | YES | 0 | |
| pnpa_demand_amt | numeric | YES | 0 | |
| pnpa_collection_amt | numeric | YES | 0 | |
| on_date_demand_amt | numeric | YES | 0 | |
| on_date_collection_amt | numeric | YES | 0 | |

### daily_performance
Per-date performance metrics (same structure as employee_performance).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| id | integer | NO | nextval('daily_performance_id_seq'::regclass) | ✓ |
| report_date | date | NO | | |
| emp_id | character varying | NO | | |
| product_type_id | integer | NO | | |
| regular_demand | integer | YES | 0 | |
| regular_collection | integer | YES | 0 | |
| demand_1_30 | integer | YES | 0 | |
| collection_1_30 | integer | YES | 0 | |
| demand_31_60 | integer | YES | 0 | |
| collection_31_60 | integer | YES | 0 | |
| pnpa_demand | integer | YES | 0 | |
| pnpa_collection | integer | YES | 0 | |
| npa_cases | integer | YES | 0 | |
| npa_act_acc | integer | YES | 0 | |
| npa_act_amt | numeric | YES | 0 | |
| npa_clo_acc | integer | YES | 0 | |
| npa_clo_amt | numeric | YES | 0 | |
| on_date_demand | integer | YES | 0 | |
| on_date_collection | integer | YES | 0 | |
| regular_demand_amt | numeric | YES | 0 | |
| regular_collection_amt | numeric | YES | 0 | |
| demand_1_30_amt | numeric | YES | 0 | |
| collection_1_30_amt | numeric | YES | 0 | |
| demand_31_60_amt | numeric | YES | 0 | |
| collection_31_60_amt | numeric | YES | 0 | |
| pnpa_demand_amt | numeric | YES | 0 | |
| pnpa_collection_amt | numeric | YES | 0 | |
| on_date_demand_amt | numeric | YES | 0 | |
| on_date_collection_amt | numeric | YES | 0 | |

---

## Daily Operations

### daily_reports
Daily Plan targets entered by Branch Managers (FTOD, DPD, NPA, Disbursement, KYC).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| id | integer | NO | nextval('daily_reports_id_seq'::regclass) | ✓ |
| branch_name | character varying | NO | | |
| date | date | NO | | |
| region | character varying | YES | | |
| district | character varying | YES | | |
| dm_name | character varying | YES | | |
| ftod_actual | integer | YES | 0 | |
| ftod_plan | integer | YES | 0 | |
| dpd_1_30_actual | integer | YES | 0 | |
| dpd_1_30_plan | integer | YES | 0 | |
| dpd_31_60_actual | integer | YES | 0 | |
| dpd_31_60_plan | integer | YES | 0 | |
| dpd_61_90_actual | integer | YES | 0 | |
| dpd_61_90_plan | integer | YES | 0 | |
| npa_activation | integer | YES | 0 | |
| npa_closure | integer | YES | 0 | |
| fy_non_start_acc | integer | YES | 0 | |
| fy_non_start_plan | integer | YES | 0 | |
| disb_igl_acc | integer | YES | 0 | |
| disb_igl_amt | numeric | YES | 0 | |
| disb_fig_acc | integer | YES | 0 | |
| disb_fig_amt | numeric | YES | 0 | |
| disb_il_acc | integer | YES | 0 | |
| disb_il_amt | numeric | YES | 0 | |
| kyc_igl | integer | YES | 0 | |
| kyc_fig | integer | YES | 0 | |
| kyc_il | integer | YES | 0 | |
| created_at | timestamp without time zone | YES | CURRENT_TIMESTAMP | |

### daily_reports_achievements
Daily Plan achievements (same structure as daily_reports, mirrors reporting table).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| id | integer | NO | nextval('daily_reports_achievements_id_seq'::regclass) | ✓ |
| branch_name | character varying | NO | | |
| date | date | NO | | |
| region | character varying | YES | | |
| district | character varying | YES | | |
| dm_name | character varying | YES | | |
| ftod_actual | integer | YES | 0 | |
| ftod_plan | integer | YES | 0 | |
| dpd_1_30_actual | integer | YES | 0 | |
| dpd_1_30_plan | integer | YES | 0 | |
| dpd_31_60_actual | integer | YES | 0 | |
| dpd_31_60_plan | integer | YES | 0 | |
| dpd_61_90_actual | integer | YES | 0 | |
| dpd_61_90_plan | integer | YES | 0 | |
| npa_activation | integer | YES | 0 | |
| npa_closure | integer | YES | 0 | |
| fy_non_start_acc | integer | YES | 0 | |
| fy_non_start_plan | integer | YES | 0 | |
| disb_igl_acc | integer | YES | 0 | |
| disb_igl_amt | numeric | YES | 0 | |
| disb_fig_acc | integer | YES | 0 | |
| disb_fig_amt | numeric | YES | 0 | |
| disb_il_acc | integer | YES | 0 | |
| disb_il_amt | numeric | YES | 0 | |
| kyc_igl | integer | YES | 0 | |
| kyc_fig | integer | YES | 0 | |
| kyc_il | integer | YES | 0 | |
| created_at | timestamp without time zone | YES | CURRENT_TIMESTAMP | |

### disbursement
Monthly disbursement summary by product type, aggregated at branch/region level.
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| id | integer | NO | nextval('disbursement_id_seq'::regclass) | ✓ |
| db_month | character varying | NO | | |
| region_name | character varying | NO | | |
| district_name | character varying | NO | | |
| branch_name | character varying | NO | | |
| emp_id | character varying | NO | | |
| officer_name | character varying | YES | ''::character varying | |
| product_name | character varying | NO | | |
| disb_count | integer | YES | 0 | |
| disb_amount | numeric | YES | 0 | |

---

## v2 Tables (New Hierarchy Structure)

### v2_states
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| state_id | integer | NO | nextval('v2_states_state_id_seq'::regclass) | ✓ |
| state_name | character varying | NO | | |

### v2_divisions
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| division_id | integer | NO | nextval('v2_divisions_division_id_seq'::regclass) | ✓ |
| division_name | character varying | NO | | |
| state_id | integer | NO | | |

### v2_areas
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| area_id | integer | NO | nextval('v2_areas_area_id_seq'::regclass) | ✓ |
| area_name | character varying | NO | | |
| division_id | integer | NO | | |

### v2_branches
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| branch_id | integer | NO | nextval('v2_branches_branch_id_seq'::regclass) | ✓ |
| branch_name | character varying | NO | | |
| area_id | integer | NO | | |

### v2_employees
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| emp_id | character varying | NO | | ✓ |
| officer_name | character varying | NO | | |
| branch_id | integer | NO | | |

### v2_employee_performance
Performance metrics using v2 hierarchy structure (same as employee_performance).
| Column | Type | Nullable | Default | PK |
|--------|------|----------|---------|-----|
| performance_id | integer | NO | nextval('v2_employee_performance_performance_id_seq'::regclass) | ✓ |
| emp_id | character varying | NO | | |
| product_type_id | integer | NO | | |
| regular_demand | integer | YES | 0 | |
| regular_collection | integer | YES | 0 | |
| demand_1_30 | integer | YES | 0 | |
| collection_1_30 | integer | YES | 0 | |
| demand_31_60 | integer | YES | 0 | |
| collection_31_60 | integer | YES | 0 | |
| pnpa_demand | integer | YES | 0 | |
| pnpa_collection | integer | YES | 0 | |
| npa_cases | integer | YES | 0 | |
| npa_act_acc | integer | YES | 0 | |
| npa_act_amt | numeric | YES | 0 | |
| npa_clo_acc | integer | YES | 0 | |
| npa_clo_amt | numeric | YES | 0 | |
| on_date_demand | integer | YES | 0 | |
| on_date_collection | integer | YES | 0 | |
| regular_demand_amt | numeric | YES | 0 | |
| regular_collection_amt | numeric | YES | 0 | |
| demand_1_30_amt | numeric | YES | 0 | |
| collection_1_30_amt | numeric | YES | 0 | |
| demand_31_60_amt | numeric | YES | 0 | |
| collection_31_60_amt | numeric | YES | 0 | |
| pnpa_demand_amt | numeric | YES | 0 | |
| pnpa_collection_amt | numeric | YES | 0 | |
| on_date_demand_amt | numeric | YES | 0 | |
| on_date_collection_amt | numeric | YES | 0 | |

---

## Notes

- **Text IDs**: `emp_id` is `character varying`, not numeric
- **Amounts**: Stored as `numeric` for currency precision
- **Defaults**: Most metrics default to `0`; timestamps default to `CURRENT_TIMESTAMP`
- **Naming**: `dr_*` columns in original code correspond to daily_reports table; `v2_*` tables reflect new hierarchy (State → Division → Area → Branch)
- **Metrics**: Standard pattern across daily/monthly/FY tables: regular demand/collection, DPD buckets (1-30, 31-60), PNPA, NPA counts & amounts, On-date demand/collection
