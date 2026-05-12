-- Repair v2_employee_pos (mirrors employee_pos) — same scramble.
-- Usage:
--   psql -d portfolio_month_wise -v month_id=24 -f fix_v2_employee_pos_column_swap.sql

BEGIN;

SELECT 'before' AS phase, month_id,
       SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos,
       SUM(sma1_pos)    AS sma1_pos,    SUM(pnpa_pos) AS pnpa_pos,
       SUM(npa_pos)     AS npa_pos,     SUM(total_pos) AS total_pos
FROM v2_employee_pos WHERE month_id = :'month_id'::int GROUP BY month_id;

UPDATE v2_employee_pos
SET regular_pos = sma1_pos_src,
    sma0_pos    = pnpa_pos_src,
    sma1_pos    = npa_pos_src,
    pnpa_pos    = sma0_pos_src,
    npa_pos     = regular_pos_src
FROM (
  SELECT month_id, emp_id,
         regular_pos AS regular_pos_src,
         sma0_pos    AS sma0_pos_src,
         sma1_pos    AS sma1_pos_src,
         pnpa_pos    AS pnpa_pos_src,
         npa_pos     AS npa_pos_src
  FROM v2_employee_pos
  WHERE month_id = :'month_id'::int
) src
WHERE v2_employee_pos.month_id = src.month_id
  AND v2_employee_pos.emp_id   = src.emp_id;

SELECT 'after' AS phase, month_id,
       SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos,
       SUM(sma1_pos)    AS sma1_pos,    SUM(pnpa_pos) AS pnpa_pos,
       SUM(npa_pos)     AS npa_pos,     SUM(total_pos) AS total_pos
FROM v2_employee_pos WHERE month_id = :'month_id'::int GROUP BY month_id;

COMMIT;
