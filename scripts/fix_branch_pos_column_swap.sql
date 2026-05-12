-- =====================================================================
-- One-shot repair: branch_pos columns were ingested in scrambled order
-- because the Excel POS sheet uses [NPA, PNPA, Regular, SMA0, SMA1, Total]
-- while the legacy loader assumed [Regular, SMA0, SMA1, PNPA, NPA, Total].
--
-- Detected source-vs-stored permutation (from MAR 2026 cross-check vs
-- portfolio_performance which is correct):
--   stored regular_pos == real npa_pos
--   stored sma0_pos    == real pnpa_pos
--   stored sma1_pos    == real regular_pos
--   stored pnpa_pos    == real sma0_pos
--   stored npa_pos     == real sma1_pos
--   stored total_pos   == real total_pos (unchanged, kept right slot)
--
-- This script unscrambles existing rows in branch_pos.
-- IMPORTANT: run ONCE per affected month. Re-running will re-scramble.
--
-- Usage on EC2:
--   psql -d portfolio_month_wise -v month_id=24 -f fix_branch_pos_column_swap.sql
-- Or hard-code WHERE month_id below.
-- =====================================================================

BEGIN;

-- Sanity snapshot before
SELECT 'before' AS phase, month_id,
       SUM(regular_pos) AS regular_pos,
       SUM(sma0_pos)    AS sma0_pos,
       SUM(sma1_pos)    AS sma1_pos,
       SUM(pnpa_pos)    AS pnpa_pos,
       SUM(npa_pos)     AS npa_pos,
       SUM(total_pos)   AS total_pos
FROM branch_pos
WHERE month_id = :'month_id'::int
GROUP BY month_id;

UPDATE branch_pos
SET regular_pos = sma1_pos_src,
    sma0_pos    = pnpa_pos_src,
    sma1_pos    = npa_pos_src,
    pnpa_pos    = sma0_pos_src,
    npa_pos     = regular_pos_src
FROM (
  SELECT pos_id,
         regular_pos AS regular_pos_src,
         sma0_pos    AS sma0_pos_src,
         sma1_pos    AS sma1_pos_src,
         pnpa_pos    AS pnpa_pos_src,
         npa_pos     AS npa_pos_src
  FROM branch_pos
  WHERE month_id = :'month_id'::int
) src
WHERE branch_pos.pos_id = src.pos_id;

-- Sanity snapshot after
SELECT 'after' AS phase, month_id,
       SUM(regular_pos) AS regular_pos,
       SUM(sma0_pos)    AS sma0_pos,
       SUM(sma1_pos)    AS sma1_pos,
       SUM(pnpa_pos)    AS pnpa_pos,
       SUM(npa_pos)     AS npa_pos,
       SUM(total_pos)   AS total_pos
FROM branch_pos
WHERE month_id = :'month_id'::int
GROUP BY month_id;

-- Cross-check against portfolio_performance (authoritative)
SELECT 'expected' AS phase, pp.month_id,
       SUM(pp.regular_pos) AS regular_pos,
       SUM(pp.sma0_pos)    AS sma0_pos,
       SUM(pp.sma1_pos)    AS sma1_pos,
       SUM(pp.pnpa_pos)    AS pnpa_pos,
       SUM(pp.npa_pos)     AS npa_pos,
       SUM(pp.total_pos)   AS total_pos
FROM portfolio_performance pp
WHERE pp.month_id = :'month_id'::int
GROUP BY pp.month_id;

-- Inspect the 3 result rows above. If 'after' matches 'expected', COMMIT.
-- If they differ, ROLLBACK and re-investigate column permutation.
COMMIT;
