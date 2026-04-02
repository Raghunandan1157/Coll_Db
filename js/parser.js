/**
 * parser.js — Excel parsing using SheetJS
 * Reads uploaded files, extracts cell values, merged ranges, and detects sections.
 */

/**
 * Read an uploaded File and return a SheetJS workbook object.
 * @param {File} file
 * @returns {Promise<object>} workbook
 */
function parseWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {
          type: 'array',
          cellStyles: true,
          cellFormula: false,
          cellNF: true,
        });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Parse a single worksheet into a structured object.
 * @param {object} workbook — SheetJS workbook
 * @param {string} sheetName
 * @returns {object} { name, rows[][], merges[], colCount, rowCount, sections[] }
 */
function parseSheet(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return null;

  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  const rowCount = range.e.r + 1;
  const colCount = range.e.c + 1;

  // Build 2D array of cell values
  const rows = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [];
    for (let c = 0; c < colCount; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell) {
        row.push({
          v: cell.v !== undefined ? cell.v : (cell.w || ''),
          w: cell.w || '',
          t: cell.t || 's',
          raw: cell,
        });
      } else {
        row.push(null);
      }
    }
    rows.push(row);
  }

  // Parse merged cell ranges
  const merges = (ws['!merges'] || []).map((m) => ({
    startRow: m.s.r,
    startCol: m.s.c,
    endRow: m.e.r,
    endCol: m.e.c,
    rowSpan: m.e.r - m.s.r + 1,
    colSpan: m.e.c - m.s.c + 1,
  }));

  // Classify each row
  const classifiedRows = rows.map((row, idx) => ({
    index: idx,
    cells: row,
    type: classifyRow(row, idx, colCount),
    isEmpty: isEmptyRow(row),
  }));

  // Detect sections
  const sections = detectSections(classifiedRows, colCount);

  return {
    name: sheetName,
    rows: classifiedRows,
    merges,
    colCount,
    rowCount,
    sections,
  };
}

/**
 * Check if a row is effectively empty (all null or whitespace).
 */
function isEmptyRow(row) {
  return row.every((cell) => {
    if (!cell) return true;
    const v = String(cell.v).trim();
    return v === '' || v === 'undefined';
  });
}

/**
 * Get the text value of a cell safely.
 */
function cellText(cell) {
  if (!cell) return '';
  if (cell.w) return String(cell.w).trim();
  if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
  return '';
}

/**
 * Get the first non-empty cell text in a row.
 */
function firstCellText(row) {
  for (const cell of row) {
    const t = cellText(cell);
    if (t) return t;
  }
  return '';
}

/**
 * Get combined text of the first few cells for pattern matching.
 */
function rowTextPreview(row) {
  return row
    .slice(0, 5)
    .map((c) => cellText(c))
    .join(' ')
    .toUpperCase();
}

// Known section title patterns
const SECTION_PATTERNS = [
  { pattern: /REGION\s*WISE/i, name: 'REGION WISE' },
  { pattern: /DISTRICT\s*WISE/i, name: 'DISTRICT WISE' },
  { pattern: /BRANCH\s*WISE/i, name: 'BRANCH WISE' },
  { pattern: /\bIGL\b/i, name: 'IGL' },
  { pattern: /\bFIG\b/i, name: 'FIG' },
  { pattern: /\bI\.?L\b/i, name: 'IL' },
  { pattern: /OFFICER\s*NAME/i, name: 'OFFICER' },
];

// Patterns for Grand Total rows
const GRAND_TOTAL_PATTERN = /GRAND\s*TOTAL/i;
const TOTAL_PATTERN = /^TOTAL$|^TOTAL\s/i;

/**
 * Classify a row into one of: title, header, metricHeader, grandTotal, branchSubtotal, divider, officer, data, empty.
 */
function classifyRow(row, rowIndex, colCount) {
  if (isEmptyRow(row)) return 'empty';

  const preview = rowTextPreview(row);
  const first = firstCellText(row);
  const upper = first.toUpperCase();

  // Section divider: IGL, FIG, IL — typically short text spanning merged cells
  if (/^(IGL|FIG|I\.?L)$/i.test(upper)) return 'divider';

  // Grand Total
  if (GRAND_TOTAL_PATTERN.test(preview)) return 'grandTotal';

  // Title: contains REGION WISE, DISTRICT WISE, BRANCH WISE, OFFICER NAME, etc.
  for (const sp of SECTION_PATTERNS) {
    if (sp.pattern.test(preview)) return 'title';
  }

  // Header row detection: look for "Sl" or "Demand" or "Collection" or "%" in cells
  if (/\b(Sl\.?|SL\.?)\b/i.test(preview) && /DEMAND|COLLECTION/i.test(preview)) return 'header';

  // Metric header: contains pattern like "CURRENT MONTH" / "PROGRESSIVE" / "LAST YEAR" merged cells
  if (/CURRENT\s*MONTH|PROGRESSIVE|LAST\s*YEAR|PERFORMANCE/i.test(preview)) return 'metricHeader';

  // Branch subtotal / Total row
  if (TOTAL_PATTERN.test(upper) || /^TOTAL$/i.test(upper)) return 'branchSubtotal';

  // Check for subtotal patterns like "Branch Total", "District Total"
  if (/TOTAL/i.test(upper) && !/GRAND/i.test(upper)) return 'branchSubtotal';

  // Data rows — if there are enough filled cells, it's a data row
  const filledCount = row.filter((c) => c && cellText(c) !== '').length;
  if (filledCount >= 2) return 'data';

  return 'empty';
}

/**
 * Detect section boundaries in the classified rows.
 * A section starts at a title or divider row and runs until the next section or end.
 * @returns {Array<{name, startRow, endRow, titleRow, headerRows[], dataRows[]}>}
 */
function detectSections(classifiedRows) {
  const sections = [];
  let currentSection = null;

  for (let i = 0; i < classifiedRows.length; i++) {
    const row = classifiedRows[i];

    // Start a new section at a title or divider
    if (row.type === 'title' || row.type === 'divider') {
      if (currentSection) {
        currentSection.endRow = i - 1;
        sections.push(currentSection);
      }
      const preview = rowTextPreview(row.cells);
      let sectionName = firstCellText(row.cells) || 'Section';
      for (const sp of SECTION_PATTERNS) {
        if (sp.pattern.test(preview)) {
          sectionName = sp.name;
          break;
        }
      }
      currentSection = {
        name: sectionName,
        startRow: i,
        endRow: null,
        titleRow: i,
        rows: [],
      };
    }

    if (currentSection) {
      currentSection.rows.push(i);
    }
  }

  // Close last section
  if (currentSection) {
    currentSection.endRow = classifiedRows.length - 1;
    sections.push(currentSection);
  }

  return sections;
}

/**
 * Check if a cell's value is numeric (for formatting purposes).
 */
function isNumeric(cell) {
  if (!cell) return false;
  if (cell.t === 'n') return true;
  const v = cell.v;
  return typeof v === 'number' && isFinite(v);
}

/**
 * Format a number with commas, or as percentage.
 * @param {*} value
 * @param {boolean} isPercentage
 * @returns {string}
 */
