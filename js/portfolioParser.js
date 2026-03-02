/**
 * portfolioParser.js — Parses the portfolio Excel workbook (6 sheets)
 * into structured data for the portfolio UI tab.
 *
 * Sheets:
 *   1. OverAll            — 25-col full report
 *   2. OverAll_On-Date    — 5-col on-date narrow report
 *   3. tom_OverAll_On-Date — 5-col tomorrow on-date
 *   4. FY_25-26           — 25-col FY report
 *   5. FY_25-26_On-Date   — 5-col FY on-date
 *   6. tom_FY_25-26_On-Date — 5-col FY tomorrow on-date
 *
 * Usage:
 *   var result = parsePortfolioWorkbook(workbook);
 *   // result.overall, result.overallOnDate, result.tomOverallOnDate,
 *   // result.fy, result.fyOnDate, result.tomFyOnDate
 *
 * Each parsed sheet object has:
 *   .rows           — raw 2D array from sheet_to_json
 *   .isWide         — true for 25-col sheets, false for 5-col
 *   .sections       — { region: {start,end,grandTotalRow}, district: {...}, branch: {...} }
 *   .productBounds  — (wide only) { igl: {regStart,regEnd,boStart,boEnd}, fig: {...}, il: {...} }
 *
 * All global functions — no ES modules.
 */

/* ===== Column layout constants ===== */

var PORT_REG = { demand: 2, collection: 3, ftod: 4, pct: 5 };

var PORT_BUCKETS = [
  { name: 'Bucket 1', range: '1 - 30 DPD', key: '1-30',  color: '#34D399', demand: 6,  collection: 7,  balance: 8,  pct: 9  },
  { name: 'Bucket 2', range: '31 - 60 DPD', key: '31-60', color: '#FBBF24', demand: 10, collection: 11, balance: 12, pct: 13 },
  { name: 'PNPA',     range: 'Pre-NPA',     key: 'pnpa',  color: '#FB923C', demand: 14, collection: 15, balance: 16, pct: 17 },
  { name: 'NPA',      range: 'NPA',         key: 'npa',   color: '#F87171', demand: 18, actAcct: 19, actAmt: 20, closAcct: 21, closAmt: 22 }
];

var PORT_RANK_COL = 23;
var PORT_PERF_COL = 24;

/* Narrow (on-date) column indices */
var PORT_NARROW = { demand: 2, collection: 3, pct: 4 };

/* ===== Internal helpers ===== */

function _pNumVal(v) { return typeof v === 'number' ? v : 0; }

function _pNorm(s) {
  return s.toUpperCase().replace(/[\s\-\u00A0]+/g, '').trim();
}

function _pFuzzy(a, b) {
  var au = a.toUpperCase().trim();
  var bu = b.toUpperCase().trim();
  if (au === bu) return true;
  var an = _pNorm(a);
  var bn = _pNorm(b);
  if (an === bn) return true;
  return false;
}

/* ===== Section detection ===== */

/**
 * Scan rows for REGION/DISTRICT/BRANCH section headers and Grand Total rows.
 * Returns { region: {start, end, grandTotalRow}, district: {...}, branch: {...} }
 * Only captures the FIRST occurrence of each (the "All products" tables).
 */
function _pDetectSections(rows) {
  var sections = {};
  var current = null;

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (row == null || row.length === 0) continue;
    var c0 = String(row[0] || '').trim().toUpperCase();
    var c1 = String(row[1] || '').trim().toUpperCase();
    var hdr = c0 + ' ' + c1;

    // Detect section starts
    if (hdr.indexOf('REGION') >= 0 && (hdr.indexOf('WISE') >= 0 || hdr.indexOf('NAME') >= 0)) {
      if (sections.region == null) {
        current = 'region';
        sections.region = { start: r, end: null, grandTotalRow: null };
      }
      continue;
    }
    if (hdr.indexOf('DISTRICT') >= 0 && (hdr.indexOf('WISE') >= 0 || hdr.indexOf('NAME') >= 0)) {
      if (sections.district == null) {
        current = 'district';
        sections.district = { start: r, end: null, grandTotalRow: null };
      }
      continue;
    }
    if (hdr.indexOf('BRANCH') >= 0 && hdr.indexOf('OFFICER') < 0 && (hdr.indexOf('WISE') >= 0 || hdr.indexOf('NAME') >= 0)) {
      if (sections.branch == null) {
        current = 'branch';
        sections.branch = { start: r, end: null, grandTotalRow: null };
      }
      continue;
    }

    // Grand Total marks end of current section
    if (c1 === 'GRAND TOTAL' || c0 === 'GRAND TOTAL') {
      if (current && sections[current] && sections[current].grandTotalRow == null) {
        sections[current].grandTotalRow = r;
        sections[current].end = r;
        current = null;
      }
    }

    // Stop scanning after branch Grand Total (product tables come after)
    if (sections.region && sections.district && sections.branch &&
        sections.region.end != null && sections.district.end != null && sections.branch.end != null) {
      break;
    }
  }

  return sections;
}

/**
 * Detect product table boundaries (IGL, FIG, IL) in a wide sheet.
 * Returns { igl: {regStart,regEnd,boStart,boEnd}, fig: {...}, il: {...} }
 */
function _pDetectProducts(rows) {
  var result = {};
  var boStarts = [];

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (row == null || row.length === 0) continue;
    var c0 = String(row[0] || '').toUpperCase().trim();
    var c1 = String(row[1] || '').toUpperCase().trim();

    // Region-level product header: "REGION WISE - IGL REPORT" in col 1
    if (c1.indexOf('REGION WISE') >= 0 && c1.indexOf('REPORT') >= 0) {
      var prod = null;
      if (c1.indexOf('IGL') >= 0) prod = 'igl';
      else if (c1.indexOf('FIG') >= 0) prod = 'fig';
      else if (c1.indexOf('IL') >= 0) prod = 'il';
      if (prod) {
        if (result[prod] == null) result[prod] = {};
        result[prod].regStart = r;
      }
    }

    // Branch+Officer product header in col 0
    if (c0.indexOf('BRANCH') >= 0 && c0.indexOf('OFFICER') >= 0 && c0.indexOf('COLLECTION REPORT') >= 0) {
      var prod2 = null;
      if (c0.indexOf('- IGL') >= 0) prod2 = 'igl';
      else if (c0.indexOf('- FIG') >= 0) prod2 = 'fig';
      else if (c0.indexOf('- IL') >= 0) prod2 = 'il';
      if (prod2) {
        if (result[prod2] == null) result[prod2] = {};
        if (result[prod2].boStart == null) {
          result[prod2].boStart = r;
          boStarts.push({ key: prod2, row: r });
        }
      }
    }
  }

  // Compute BO end rows
  boStarts.sort(function (a, b) { return a.row - b.row; });
  for (var i = 0; i < boStarts.length; i++) {
    var endRow = (i + 1 < boStarts.length) ? boStarts[i + 1].row : rows.length;
    result[boStarts[i].key].boEnd = endRow;
  }

  // Compute region section end rows (scan for Grand Total)
  var products = ['igl', 'fig', 'il'];
  for (var p = 0; p < products.length; p++) {
    var k = products[p];
    if (result[k] && result[k].regStart != null) {
      result[k].regEnd = result[k].regStart + 50;
      for (var rr = result[k].regStart; rr < Math.min(result[k].regStart + 50, rows.length); rr++) {
        var row2 = rows[rr];
        if (row2 == null) continue;
        if (String(row2[1] || '').trim().toUpperCase() === 'GRAND TOTAL') {
          result[k].regEnd = rr + 1;
          break;
        }
      }
    }
  }

  return result;
}

/* ===== Row extraction helpers ===== */

/**
 * Collect all data rows from a section type ('region', 'district', 'branch').
 * Uses sections object for bounds.
 */
function _pGetSectionRows(rows, sections, sectionType) {
  var sec = sections[sectionType];
  if (sec == null) return [];
  var results = [];
  for (var r = sec.start; r <= sec.end; r++) {
    var row = rows[r];
    if (row == null || row.length === 0) continue;
    var c0 = String(row[0] || '').trim().toUpperCase();
    var c1 = String(row[1] || '').trim().toUpperCase();
    // Skip headers, blank, serial number, total rows
    if (c1 === '' && c0 === '') continue;
    if (c1 === 'GRAND TOTAL' || c0 === 'GRAND TOTAL') continue;
    if (/^(SL|SL\.|SL NO|S\.NO|SR)$/i.test(c0)) continue;
    if (/WISE|NAME|DEMAND|COLLECTION|FTOD|BALANCE|REPORT/i.test(c1) && typeof row[2] !== 'number') continue;
    if (c1 === '' || /^\d+$/.test(c1)) continue;
    // Must have some numeric data
    if (typeof row[2] !== 'number' && typeof row[3] !== 'number') continue;
    results.push({ name: c1 || c0, row: row });
  }
  return results;
}

/**
 * Find the Grand Total row in a section.
 */
function _pGetGrandTotal(rows, sections, sectionType) {
  var sec = sections[sectionType];
  if (sec == null || sec.grandTotalRow == null) return null;
  return rows[sec.grandTotalRow];
}

/**
 * Find a role-specific row:
 *   CEO  -> Grand Total (from region section)
 *   RM   -> matching region row
 *   DM   -> matching district row
 *   BM   -> matching branch row
 */
function portFindRoleRow(parsed, role, location) {
  var rows = parsed.rows;
  var sections = parsed.sections;
  if (role === 'CEO') {
    return _pGetGrandTotal(rows, sections, 'region');
  }
  var sectionMap = { RM: 'region', DM: 'district', BM: 'branch' };
  var sectionType = sectionMap[role];
  if (sectionType == null) return null;
  var locUpper = (location || '').toUpperCase().trim();
  var sectionRows = _pGetSectionRows(rows, sections, sectionType);
  for (var i = 0; i < sectionRows.length; i++) {
    if (_pFuzzy(sectionRows[i].name, locUpper)) return sectionRows[i].row;
  }
  return null;
}

/**
 * Find children for drill-down.
 *   CEO -> regions, RM -> districts (filtered by hierarchy), DM -> branches, BM -> officers
 */
function portFindChildren(parsed, hierarchy, role, location) {
  var rows = parsed.rows;
  var sections = parsed.sections;
  if (role === 'CEO') {
    return _pGetSectionRows(rows, sections, 'region');
  }
  if (role === 'FO') return [];
  var locUpper = (location || '').toUpperCase().trim();

  if (role === 'RM') {
    var hierDistricts = null;
    for (var rk in hierarchy) {
      if (_pFuzzy(rk, locUpper)) { hierDistricts = Object.keys(hierarchy[rk]); break; }
    }
    if (hierDistricts == null) return [];
    var allDistricts = _pGetSectionRows(rows, sections, 'district');
    var filtered = [];
    for (var i = 0; i < allDistricts.length; i++) {
      for (var j = 0; j < hierDistricts.length; j++) {
        if (_pFuzzy(allDistricts[i].name, hierDistricts[j])) {
          filtered.push(allDistricts[i]);
          break;
        }
      }
    }
    return filtered;
  }

  if (role === 'DM') {
    var hierBranches = null;
    for (var rk in hierarchy) {
      for (var dk in hierarchy[rk]) {
        if (_pFuzzy(dk, locUpper)) { hierBranches = hierarchy[rk][dk]; break; }
      }
      if (hierBranches) break;
    }
    if (hierBranches == null) return [];
    var allBranches = _pGetSectionRows(rows, sections, 'branch');
    var filtered2 = [];
    for (var i = 0; i < allBranches.length; i++) {
      for (var j = 0; j < hierBranches.length; j++) {
        if (_pFuzzy(allBranches[i].name, hierBranches[j])) {
          filtered2.push(allBranches[i]);
          break;
        }
      }
    }
    return filtered2;
  }

  // BM -> officers from BO table (not in sections, scan from row 232+ area)
  // For portfolio we don't have a separate BO section detector in sections,
  // so scan from after the branch Grand Total
  if (role === 'BM') {
    return _pFindOfficers(rows, locUpper);
  }

  return [];
}

/**
 * Find officers for a branch from the Branch+Officer section (wide sheets).
 */
function _pFindOfficers(rows, branchUpper) {
  var officers = [];
  var inBO = false;
  var foundBranch = false;

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    if (row == null || row.length === 0) continue;
    var c0 = String(row[0] || '').trim();
    var c1 = String(row[1] || '').trim();
    var c0u = c0.toUpperCase();
    var c1u = c1.toUpperCase();

    // Detect BO section header
    if (c0u.indexOf('BRANCH') >= 0 && c0u.indexOf('OFFICER') >= 0 && c0u.indexOf('COLLECTION REPORT') >= 0) {
      // Only use first (All products) BO section
      if (inBO) break;
      inBO = true;
      continue;
    }
    if (c0u === 'EMP ID' || c1u === 'BRANCH / OFFICER NAME') continue;
    if (c1u === 'DEMAND' || c1u === 'COLLECTION') continue;

    if (inBO === false) continue;

    var hasEmpId = c0.length > 0 && /^[A-Z]{2}\d+/i.test(c0);

    // Branch header row (no emp ID, name in col 1, has data)
    if (hasEmpId === false && c1.length > 0 && typeof row[2] === 'number') {
      if (_pFuzzy(c1u, branchUpper)) {
        foundBranch = true;
      } else if (foundBranch) {
        break; // past our branch
      }
      continue;
    }

    if (hasEmpId && foundBranch) {
      officers.push({ name: c1.trim(), empId: c0, row: row });
    }
  }

  return officers;
}

/* ===== Product-specific row finders (wide sheets) ===== */

/**
 * Find the role row in a product table.
 */
function portFindProductRow(parsed, product, role, location, hierarchy) {
  var bounds = parsed.productBounds;
  if (bounds == null) return null;
  var b = bounds[product];
  if (b == null) return null;
  var rows = parsed.rows;

  // CEO: Grand Total from product region section
  if (role === 'CEO' && b.regStart != null) {
    for (var r = b.regStart; r < b.regEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      if (String(row[1] || '').trim().toUpperCase() === 'GRAND TOTAL') return row;
    }
    return null;
  }

  // RM: region row from product region section
  if (role === 'RM' && b.regStart != null) {
    var loc = (location || '').toUpperCase().trim();
    for (var r = b.regStart; r < b.regEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      var c1 = String(row[1] || '').trim();
      if (c1 && _pFuzzy(c1, loc)) return row;
    }
    return null;
  }

  // DM: aggregate branches from product BO table
  if (role === 'DM') {
    return _pAggregateBranches(rows, b, location, hierarchy);
  }

  // BM: find branch header in product BO table
  if (role === 'BM') {
    var bu = (location || '').toUpperCase().trim();
    for (var r = b.boStart; r < b.boEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      var c0 = String(row[0] || '').trim();
      var c1 = String(row[1] || '').trim();
      if (c0 === '' && c1 && _pFuzzy(c1, bu)) return row;
    }
    return null;
  }

  // FO: find by emp ID
  var needleId = String(location || '').trim().toUpperCase();
  for (var r = b.boStart; r < b.boEnd; r++) {
    var row = rows[r];
    if (row == null) continue;
    var c0 = String(row[0] || '').trim().toUpperCase();
    if (needleId && c0 === needleId) return row;
  }
  return null;
}

/**
 * Aggregate branch data for a district from a product BO table.
 */
function _pAggregateBranches(rows, b, districtName, hierarchy) {
  var distUpper = (districtName || '').toUpperCase().trim();
  var branches = null;
  for (var rk in hierarchy) {
    for (var dk in hierarchy[rk]) {
      if (_pFuzzy(dk, distUpper)) { branches = hierarchy[rk][dk]; break; }
    }
    if (branches) break;
  }
  if (branches == null || branches.length === 0) return null;

  var branchSet = {};
  for (var i = 0; i < branches.length; i++) {
    branchSet[_pNorm(branches[i])] = true;
  }

  var sumRow = new Array(25);
  sumRow[0] = null;
  sumRow[1] = districtName;
  for (var c = 2; c < 25; c++) sumRow[c] = 0;

  var found = false;
  for (var r = b.boStart; r < b.boEnd; r++) {
    var row = rows[r];
    if (row == null) continue;
    var c0 = String(row[0] || '').trim();
    var c1 = String(row[1] || '').trim();
    if (c0 === '' && c1 && branchSet[_pNorm(c1)]) {
      found = true;
      for (var c = 2; c < Math.min(row.length, 24); c++) {
        if (typeof row[c] === 'number') sumRow[c] += row[c];
      }
    }
  }
  if (found === false) return null;

  // Recompute percentage columns
  if (sumRow[2] > 0) sumRow[5] = sumRow[3] / sumRow[2];
  if (sumRow[6] > 0) sumRow[9] = sumRow[7] / sumRow[6];
  if (sumRow[10] > 0) sumRow[13] = sumRow[11] / sumRow[10];
  if (sumRow[14] > 0) sumRow[17] = sumRow[15] / sumRow[14];

  return sumRow;
}

/**
 * Find product children for drill-down.
 */
function portFindProductChildren(parsed, product, role, location, hierarchy) {
  var bounds = parsed.productBounds;
  if (bounds == null) return [];
  var b = bounds[product];
  if (b == null) return [];
  var rows = parsed.rows;

  if (role === 'CEO' && b.regStart != null) {
    var children = [];
    for (var r = b.regStart; r < b.regEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      var c1 = String(row[1] || '').trim();
      if (c1 === '') continue;
      var c1u = c1.toUpperCase();
      if (c1u === 'GRAND TOTAL' || c1u.indexOf('REGION') >= 0 || c1u.indexOf('REPORT') >= 0) continue;
      if (c1u === 'DEMAND' || c1u === 'COLLECTION' || c1u === 'FTOD') continue;
      if (typeof row[2] !== 'number') continue;
      children.push({ name: c1u, row: row });
    }
    return children;
  }

  if (role === 'RM') {
    var locUpper = (location || '').toUpperCase().trim();
    var regionData = null;
    for (var rk in hierarchy) {
      if (_pFuzzy(rk, locUpper)) { regionData = hierarchy[rk]; break; }
    }
    if (regionData == null) return [];
    var children2 = [];
    for (var dk in regionData) {
      var aggRow = _pAggregateBranches(rows, b, dk, hierarchy);
      if (aggRow) children2.push({ name: dk.toUpperCase(), row: aggRow });
    }
    return children2;
  }

  if (role === 'DM') {
    var locUpper2 = (location || '').toUpperCase().trim();
    var branchNames = null;
    for (var rk in hierarchy) {
      for (var dk in hierarchy[rk]) {
        if (_pFuzzy(dk, locUpper2)) { branchNames = hierarchy[rk][dk]; break; }
      }
      if (branchNames) break;
    }
    if (branchNames == null) return [];
    var branchSet = {};
    for (var i = 0; i < branchNames.length; i++) {
      branchSet[_pNorm(branchNames[i])] = true;
    }
    var children3 = [];
    for (var r = b.boStart; r < b.boEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      var c0 = String(row[0] || '').trim();
      var c1 = String(row[1] || '').trim();
      if (c0 === '' && c1 && branchSet[_pNorm(c1)]) {
        children3.push({ name: c1.toUpperCase(), row: row });
      }
    }
    return children3;
  }

  if (role === 'BM') {
    var bu = (location || '').toUpperCase().trim();
    var officers = [];
    var foundBranch = false;
    for (var r = b.boStart; r < b.boEnd; r++) {
      var row = rows[r];
      if (row == null) continue;
      var c0 = String(row[0] || '').trim();
      var c1 = String(row[1] || '').trim();
      var c0u = c0.toUpperCase();
      var c1u = c1.toUpperCase();
      if (c0u === 'EMP ID' || c1u === 'BRANCH / OFFICER NAME') continue;
      if (c1u === 'DEMAND' || c1u === 'COLLECTION') continue;
      var hasEmpId = c0.length > 0 && /^[A-Z]{2}\d+/i.test(c0);
      if (hasEmpId === false && c1.length > 0 && row.length > 3) {
        if (_pFuzzy(c1u, bu)) foundBranch = true;
        else if (foundBranch) break;
        continue;
      }
      if (hasEmpId && foundBranch) {
        officers.push({ name: c1.trim(), empId: c0, row: row });
      }
    }
    return officers;
  }

  return [];
}

/* ===== Narrow (on-date) sheet helpers ===== */

/**
 * Find role row in a narrow (5-col) sheet.
 * Same section logic but only 5 columns.
 */
function portFindNarrowRoleRow(parsed, role, location) {
  return portFindRoleRow(parsed, role, location);
}

/**
 * Find children in a narrow sheet.
 */
function portFindNarrowChildren(parsed, hierarchy, role, location) {
  return portFindChildren(parsed, hierarchy, role, location);
}

/* ===== Main parser entry point ===== */

/**
 * Parse a SheetJS workbook into structured portfolio data.
 * @param {object} workbook — SheetJS workbook object
 * @returns {object} with keys: overall, overallOnDate, tomOverallOnDate, fy, fyOnDate, tomFyOnDate
 *   Each has: { rows, isWide, sections, productBounds (if wide) }
 *   Returns null if workbook has no recognizable sheets.
 */
function parsePortfolioWorkbook(workbook) {
  if (workbook == null || workbook.SheetNames == null) return null;

  var sheetMap = {
    overall:           { match: 'OverAll',              wide: true  },
    overallOnDate:     { match: 'OverAll_On-Date',      wide: false },
    tomOverallOnDate:  { match: 'tom_OverAll_On-Date',  wide: false },
    fy:                { match: 'FY_25-26',             wide: true  },
    fyOnDate:          { match: 'FY_25-26_On-Date',     wide: false },
    tomFyOnDate:       { match: 'tom_FY_25-26_On-Date', wide: false }
  };

  var result = {};
  var foundAny = false;

  for (var key in sheetMap) {
    var cfg = sheetMap[key];
    var sheetName = _pFindSheet(workbook, cfg.match);
    if (sheetName == null) {
      result[key] = null;
      continue;
    }

    var ws = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows == null || rows.length < 2) {
      result[key] = null;
      continue;
    }

    var parsed = {
      rows: rows,
      isWide: cfg.wide,
      sections: _pDetectSections(rows),
      productBounds: cfg.wide ? _pDetectProducts(rows) : null
    };

    result[key] = parsed;
    foundAny = true;
  }

  return foundAny ? result : null;
}

/**
 * Find a sheet by name (case-insensitive exact match, then includes match).
 */
function _pFindSheet(workbook, target) {
  var targetUpper = target.toUpperCase();
  // Exact match first
  for (var i = 0; i < workbook.SheetNames.length; i++) {
    if (workbook.SheetNames[i].toUpperCase() === targetUpper) return workbook.SheetNames[i];
  }
  // Includes match
  for (var i = 0; i < workbook.SheetNames.length; i++) {
    if (workbook.SheetNames[i].toUpperCase().indexOf(targetUpper) >= 0) return workbook.SheetNames[i];
  }
  return null;
}

/* ===== Convenience: get report date from sheet title ===== */

/**
 * Extract the report date string from a parsed sheet's first row.
 * e.g. "31-01-2026" from "REGION - WISE COLLECTION REPORT - as on 31-01-2026 (OverAll)"
 */
function portGetReportDate(parsed) {
  if (parsed == null || parsed.rows == null || parsed.rows.length === 0) return null;
  for (var r = 0; r < Math.min(3, parsed.rows.length); r++) {
    var row = parsed.rows[r];
    if (row == null) continue;
    for (var c = 0; c < Math.min(3, row.length); c++) {
      var val = String(row[c] || '');
      var m = val.match(/(\d{2}-\d{2}-\d{4})/);
      if (m) return m[1];
    }
  }
  return null;
}
