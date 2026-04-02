/**
 * employeeParser.js — Extracts employee/officer data from the last sheet
 * of a Daily Collection Report workbook.
 *
 * All functions are global (no ES modules). Depends on the XLSX global
 * from SheetJS being available.
 */

// Keywords used to locate the header row
var _EP_HEADER_KEYWORDS = ['OFFICER', 'NAME', 'DEMAND', 'COLLECTION', 'SL'];

// Patterns for the ID column header
var _EP_ID_PATTERNS = [/^SL\.?$/i, /^S\.?\s*NO\.?$/i, /^SR\.?$/i, /^ID$/i, /^EMP/i, /^SL\s*NO/i];

// Patterns for the Name column header
var _EP_NAME_PATTERNS = [/OFFICER/i, /NAME/i, /EMPLOYEE/i];

// Rows matching these patterns should be skipped during extraction
var _EP_SKIP_PATTERNS = [
  /^GRAND\s*TOTAL$/i,
  /^TOTAL$/i,
  /^TOTAL\s/i,
  /^REGION\s*WISE/i,
  /^DISTRICT\s*WISE/i,
  /^BRANCH\s*WISE/i,
  /^IGL$/i,
  /^FIG$/i,
  /^I\.?L$/i,
];

/**
 * Safely convert a value to a trimmed string.
 * Handles null, undefined, numbers, and objects.
 * @param {*} val
 * @returns {string}
 */
function _epStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

/**
 * Test whether a cell value matches any pattern in a list.
 * @param {string} text
 * @param {RegExp[]} patterns
 * @returns {boolean}
 */
function _epMatchesAny(text, patterns) {
  for (var i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) return true;
  }
  return false;
}

/**
 * Determine whether a row array looks like a header row for officer/employee data.
 * A header row must contain at least one ID-like keyword and one data keyword
 * (DEMAND or COLLECTION).
 * @param {Array} row — raw array of cell values
 * @returns {boolean}
 */
function _epIsHeaderRow(row) {
  if (!row || !row.length) return false;

  var hasIdOrName = false;
  var hasData = false;

  for (var c = 0; c < row.length; c++) {
    var t = _epStr(row[c]).toUpperCase();
    if (!t) continue;

    if (/SL|NAME|OFFICER|EMPLOYEE|EMP|SR|S\.?\s*NO/i.test(t)) {
      hasIdOrName = true;
    }
    if (/DEMAND|COLLECTION/i.test(t)) {
      hasData = true;
    }
  }

  return hasIdOrName && hasData;
}

/**
 * Auto-detect the column index for the ID field from a header row.
 * @param {Array} headerRow
 * @returns {number} column index, or -1 if not found
 */
function _epDetectIdCol(headerRow) {
  for (var c = 0; c < headerRow.length; c++) {
    var t = _epStr(headerRow[c]);
    if (t && _epMatchesAny(t, _EP_ID_PATTERNS)) return c;
  }
  // Fallback: column 0
  return 0;
}

/**
 * Auto-detect the column index for the Name field from a header row.
 * @param {Array} headerRow
 * @returns {number} column index, or -1 if not found
 */
function _epDetectNameCol(headerRow) {
  for (var c = 0; c < headerRow.length; c++) {
    var t = _epStr(headerRow[c]);
    if (t && _epMatchesAny(t, _EP_NAME_PATTERNS)) return c;
  }
  // Fallback: column 1 (or 0 if only one column)
  return headerRow.length > 1 ? 1 : 0;
}

/**
 * Check if a row is effectively empty (all values are null, undefined, or blank).
 * @param {Array} row
 * @returns {boolean}
 */
function _epIsEmptyRow(row) {
  if (!row || !row.length) return true;
  for (var c = 0; c < row.length; c++) {
    if (_epStr(row[c]) !== '') return false;
  }
  return true;
}

/**
 * Check if a row should be skipped (total rows, section titles, dividers).
 * @param {Array} row
 * @param {number} idCol
 * @param {number} nameCol
 * @returns {boolean}
 */
function _epShouldSkipRow(row, idCol, nameCol) {
  // Check the first few cells against skip patterns
  var checkCols = [0, 1, idCol, nameCol];
  var seen = {};
  for (var i = 0; i < checkCols.length; i++) {
    var ci = checkCols[i];
    if (ci < 0 || ci >= row.length || seen[ci]) continue;
    seen[ci] = true;
    var t = _epStr(row[ci]);
    if (t && _epMatchesAny(t, _EP_SKIP_PATTERNS)) return true;
  }
  return false;
}

/**
 * Extract employees from the last sheet of a workbook.
 *
 * @param {object} workbook — SheetJS workbook object
 * @returns {{employees: Array<{id: string, name: string, rowData: Array}>, headers: Array, sheetName: string}}
 */
function extractEmployees(workbook) {
  var result = { employees: [], headers: [], sheetName: '' };

  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
    return result;
  }

  var sheetName = workbook.SheetNames[workbook.SheetNames.length - 1];
  result.sheetName = sheetName;

  var ws = workbook.Sheets[sheetName];
  if (!ws) return result;

  // Convert to 2D array of raw values
  var rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  if (!rows || !rows.length) return result;

  // Scan for the LAST header row (officer-level data is usually near the bottom)
  var headerRowIdx = -1;
  for (var r = 0; r < rows.length; r++) {
    if (_epIsHeaderRow(rows[r])) {
      headerRowIdx = r;
    }
  }

  // If no header found, return empty
  if (headerRowIdx < 0) return result;

  var headerRow = rows[headerRowIdx];
  result.headers = headerRow.map(function (v) { return _epStr(v); });

  // Detect ID and Name columns from the header
  var idCol = _epDetectIdCol(headerRow);
  var nameCol = _epDetectNameCol(headerRow);

  // Ensure idCol and nameCol are different; if they collide, adjust
  if (idCol === nameCol) {
    if (idCol === 0 && headerRow.length > 1) {
      nameCol = 1;
    } else if (idCol > 0) {
      idCol = 0;
    }
  }

  // Extract data rows after the header
  var seenIds = {};
  for (var r = headerRowIdx + 1; r < rows.length; r++) {
    var row = rows[r];

    // Skip empty rows
    if (_epIsEmptyRow(row)) continue;

    // Skip total/title/divider rows
    if (_epShouldSkipRow(row, idCol, nameCol)) continue;

    var id = _epStr(row[idCol]);
    var name = _epStr(row[nameCol]);

    // Skip if ID or Name is empty
    if (!id || !name) continue;

    // Skip if ID looks like a header keyword repeated
    if (_epMatchesAny(id, _EP_ID_PATTERNS)) continue;

    // Deduplicate by ID (keep first occurrence)
    var idLower = id.toLowerCase();
    if (seenIds[idLower]) continue;
    seenIds[idLower] = true;

    result.employees.push({
      id: id,
      name: name,
      rowData: row,
    });
  }

  return result;
}

/**
 * Search an employees array for an employee by ID.
 * Comparison is case-insensitive and trimmed.
 *
 * @param {Array<{id: string, name: string, rowData: Array}>} employees
 * @param {string} id — the ID to search for
 * @returns {{id: string, name: string, rowData: Array}|null}
 */
function findEmployeeById(employees, id) {
  if (!employees || !employees.length || !id) return null;

  var needle = String(id).trim().toLowerCase();
  if (!needle) return null;

  for (var i = 0; i < employees.length; i++) {
    if (employees[i].id.toLowerCase() === needle) {
      return employees[i];
    }
  }
  return null;
}

/**
 * Get the name of the last sheet in a workbook.
 *
 * @param {object} workbook — SheetJS workbook object
 * @returns {string} sheet name, or empty string if workbook is invalid
 */
function getLastSheetName(workbook) {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
    return '';
  }
  return workbook.SheetNames[workbook.SheetNames.length - 1];
}
