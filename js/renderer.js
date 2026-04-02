/**
 * renderer.js — Renders parsed workbook data as formatted HTML tables.
 */

/**
 * Main entry: render tab navigation and first sheet.
 * @param {object} workbook — SheetJS workbook
 * @param {HTMLElement} tabsContainer
 * @param {HTMLElement} reportContainer
 */
function renderReport(workbook, tabsContainer, reportContainer) {
  const sheetNames = workbook.SheetNames;
  if (!sheetNames.length) {
    reportContainer.innerHTML = '<p>No sheets found in this workbook.</p>';
    return;
  }

  tabsContainer.innerHTML = '';
  tabsContainer.hidden = false;

  // Create tabs
  sheetNames.forEach((name, idx) => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab' + (idx === 0 ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', () => {
      tabsContainer.querySelectorAll('.sheet-tab').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      renderSheetByName(workbook, name, reportContainer);
    });
    tabsContainer.appendChild(tab);
  });

  // Render first sheet
  renderSheetByName(workbook, sheetNames[0], reportContainer);
}

/**
 * Parse and render a sheet by name.
 */
function renderSheetByName(workbook, sheetName, container) {
  const parsed = parseSheet(workbook, sheetName);
  if (!parsed) {
    container.innerHTML = '<p>Could not parse sheet: ' + sheetName + '</p>';
    return;
  }
  renderSheet(parsed, container);
}

/**
 * Build merged-cell lookup for a sheet.
 * Returns a Map: "r,c" -> { skip: true } or { rowSpan, colSpan }
 */
function buildMergeMap(merges) {
  const map = new Map();
  for (const m of merges) {
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        if (r === m.startRow && c === m.startCol) {
          map.set(r + ',' + c, { rowSpan: m.rowSpan, colSpan: m.colSpan });
        } else {
          map.set(r + ',' + c, { skip: true });
        }
      }
    }
  }
  return map;
}

/**
 * Wide sheet (25-col) CORRECT column map based on actual Excel structure:
 *
 * BUCKET GROUP          | Cols         | Sub-headers
 * ----------------------|--------------|----------------------------------------------
 * REGULAR DEMAND        | C(2)-F(5)   | DEMAND(2), COLLECTION(3), FTOD(4), COLL%(5)
 * 1-30 DPD             | G(6)-J(9)   | DEMAND(6), COLLECTION(7), BALANCE(8), COLL%(9)
 * 31-60 DPD            | K(10)-N(13) | DEMAND(10), COLLECTION(11), BALANCE(12), COLL%(13)
 * PNPA                 | O(14)-R(17) | DEMAND(14), COLLECTION(15), BALANCE(16), COLL%(17)
 * NPA                  | S(18)-W(22) | DEMAND(18), ACT.ACCT(19), ACT.AMT(20), CLS.ACCT(21), CLS.AMT(22)
 * METRICS              | X(23)-Y(24) | RANK(23), PERFORMANCE(24)
 */

/** Demand columns (peach): C(2), G(6), K(10), O(14), S(18) */
const DEMAND_COLS_WIDE = new Set([2, 6, 10, 14, 18]);

/** Collection columns (green): D(3), H(7), L(11), P(15) — NOT NPA cols */
const COLLECTION_COLS_WIDE = new Set([3, 7, 11, 15]);

/** FTOD / Balance columns (light peach): E(4), I(8), M(12), Q(16) */
const BALANCE_COLS_WIDE = new Set([4, 8, 12, 16]);

/** Collection % columns (yellow): F(5), J(9), N(13), R(17) */
const PERCENT_COLS_WIDE = new Set([5, 9, 13, 17]);

/** NPA sub-columns (blue-ish): T(19)=Act.Acct, U(20)=Act.Amt, V(21)=Cls.Acct, W(22)=Cls.Amt */
const NPA_SUB_COLS_WIDE = new Set([19, 20, 21, 22]);

/** Conditional formatting targets — same as percent cols */
const COND_FORMAT_COLS_WIDE = new Set([5, 9, 13, 17]);

/** Rank column: X => 0-indexed: 23 */
const RANK_COL = 23;
/** Performance column: Y => 0-indexed: 24 */
const PERFORMANCE_COL = 24;

/**
 * Narrow sheet (5 cols: A-E) column mapping:
 *   A(0) = spacer/title, B(1) = Name, C(2) = Demand, D(3) = Collection, E(4) = Collection %
 */
const DEMAND_COL_NARROW = 2;
const COLLECTION_COL_NARROW = 3;
const PERCENT_COL_NARROW = 4;

/**
 * Render a full parsed sheet into the container.
 */
function renderSheet(parsed, container) {
  container.innerHTML = '';

  const isWide = parsed.colCount >= 15;
  const mergeMap = buildMergeMap(parsed.merges);

  // Collect data rows per section for conditional formatting
  const sectionDataValues = collectSectionValues(parsed);

  // Build a single table for the sheet
  const table = document.createElement('table');
  table.className = 'report-table';

  for (let r = 0; r < parsed.rows.length; r++) {
    const rowInfo = parsed.rows[r];
    if (rowInfo.isEmpty && rowInfo.type === 'empty') {
      // Render empty row as spacing
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = parsed.colCount;
      td.style.border = 'none';
      td.style.height = '8px';
      td.innerHTML = '&nbsp;';
      tr.appendChild(td);
      table.appendChild(tr);
      continue;
    }

    const tr = document.createElement('tr');
    tr.className = getRowClass(rowInfo.type);

    for (let c = 0; c < parsed.colCount; c++) {
      const mergeKey = r + ',' + c;
      const mergeInfo = mergeMap.get(mergeKey);
      if (mergeInfo && mergeInfo.skip) continue;

      const cell = rowInfo.cells[c];
      const td = document.createElement('td');

      if (mergeInfo) {
        if (mergeInfo.rowSpan > 1) td.rowSpan = mergeInfo.rowSpan;
        if (mergeInfo.colSpan > 1) td.colSpan = mergeInfo.colSpan;
      }

      // Cell value
      const rawVal = cell ? cell.v : '';
      const displayVal = cell ? (cell.w || String(rawVal)) : '';

      // Determine formatting — works for both wide (25-col) and narrow (5-col) sheets
      const isPercent = isWide ? PERCENT_COLS_WIDE.has(c) : c === PERCENT_COL_NARROW;
      const isDemand = isWide ? DEMAND_COLS_WIDE.has(c) : c === DEMAND_COL_NARROW;
      const isCollection = isWide ? COLLECTION_COLS_WIDE.has(c) : c === COLLECTION_COL_NARROW;
      const isRank = isWide && c === RANK_COL;
      const isDataRow = rowInfo.type === 'data' || rowInfo.type === 'officer';
      const isNumVal = cell && isNumeric(cell);

      // Format display value
      let formatted = displayVal;
      if (isDataRow && isNumVal && cell) {
        formatted = formatNumber(cell.v, isPercent);
      } else if (displayVal === '-' || (typeof rawVal === 'string' && rawVal.trim() === '-')) {
        formatted = '-';
      }

      td.textContent = formatted;

      // Apply CSS classes
      const classes = [];

      if (isNumVal || isPercent) {
        classes.push('num');
      }

      // Column-based coloring for data rows (both wide and narrow)
      var isBalance = isWide && BALANCE_COLS_WIDE.has(c);
      var isNpaSub = isWide && NPA_SUB_COLS_WIDE.has(c);
      if (isDataRow) {
        if (isDemand) classes.push('bg-peach');
        else if (isCollection) classes.push('bg-green');
        else if (isBalance) classes.push('bg-peach-light');
        else if (isPercent) classes.push('bg-yellow');
        else if (isNpaSub) classes.push('bg-blue');
        else if (isWide && c === RANK_COL) classes.push('col-rank');
        else if (c <= 1) classes.push('bg-white');
      }

      // Conditional formatting for percentage columns
      const condFormatCols = isWide ? COND_FORMAT_COLS_WIDE : new Set([PERCENT_COL_NARROW]);
      if (isDataRow && condFormatCols.has(c) && isNumVal && cell) {
        const condClass = getConditionalClass(cell.v, c, r, sectionDataValues);
        if (condClass) classes.push(condClass);
      }

      // Performance column conditional formatting (wide sheets only, col Y=24)
      if (isDataRow && isWide && c === PERFORMANCE_COL && cell) {
        const perfText = String(cell.v || cell.w || '').toUpperCase().trim();
        if (perfText.includes('ABOVE') || perfText.includes('TOP')) {
          classes.push('cond-above');
        } else if (perfText.includes('BELOW') || perfText.includes('BOTTOM')) {
          classes.push('cond-below');
        } else if (perfText === 'N/A' || perfText === 'NA' || perfText === '-') {
          classes.push('cond-na');
        }
      }

      if (classes.length) td.className = classes.join(' ');
      tr.appendChild(td);
    }

    table.appendChild(tr);
  }

  container.appendChild(table);
}

/**
 * Map row type to CSS class.
 */
function getRowClass(type) {
  switch (type) {
    case 'grandTotal': return 'row-grand-total';
    case 'branchSubtotal': return 'row-branch-subtotal';
    case 'divider': return 'row-divider';
    case 'title': return 'row-title';
    case 'header': return 'row-header';
    case 'metricHeader': return 'row-metric-header';
    default: return '';
  }
}

/**
 * Collect numeric values per section per percentage column, for conditional formatting.
 * Returns Map: sectionIndex -> Map: colIndex -> { values: [{v, row}], sorted }
 */
function collectSectionValues(parsed) {
  const result = new Map();
  const isWide = parsed.colCount >= 15;
  const percentCols = isWide ? COND_FORMAT_COLS_WIDE : new Set([PERCENT_COL_NARROW]);

  // Find which section each row belongs to
  const rowSectionMap = new Map();
  parsed.sections.forEach((section, secIdx) => {
    section.rows.forEach((rowIdx) => {
      rowSectionMap.set(rowIdx, secIdx);
    });
  });

  for (let r = 0; r < parsed.rows.length; r++) {
    const rowInfo = parsed.rows[r];
    if (rowInfo.type !== 'data') continue;

    const secIdx = rowSectionMap.get(r);
    if (secIdx === undefined) continue;

    if (!result.has(secIdx)) result.set(secIdx, new Map());
    const secMap = result.get(secIdx);

    for (const col of percentCols) {
      const cell = rowInfo.cells[col];
      if (!cell || !isNumeric(cell)) continue;

      if (!secMap.has(col)) secMap.set(col, []);
      secMap.get(col).push({ v: cell.v, row: r });
    }
  }

  // Sort each column's values and determine top/bottom 3
  for (const [secIdx, secMap] of result) {
    for (const [col, entries] of secMap) {
      if (entries.length < 2) continue;
      const sorted = [...entries].sort((a, b) => b.v - a.v);
      const topRows = new Set(sorted.slice(0, 3).map((e) => e.row));
      const bottomRows = new Set(sorted.slice(-3).map((e) => e.row));
      secMap.set(col, { entries, topRows, bottomRows });
    }
  }

  return { result, rowSectionMap };
}

/**
 * Get conditional formatting class for a percentage cell.
 */
function getConditionalClass(value, col, rowIndex, sectionData) {
  const { result, rowSectionMap } = sectionData;
  const secIdx = rowSectionMap.get(rowIndex);
  if (secIdx === undefined) return null;

  const secMap = result.get(secIdx);
  if (!secMap) return null;

  const colData = secMap.get(col);
  if (!colData || !colData.topRows) return null;

  if (colData.topRows.has(rowIndex)) return 'cond-top';
  if (colData.bottomRows.has(rowIndex)) return 'cond-bottom';
  return null;
}
