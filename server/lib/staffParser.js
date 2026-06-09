/**
 * staffParser.js — Parse the "Working staff details" Excel into canonical rows.
 *
 * Header auto-detect: the 2026 staff file has headers in row 0; older
 * daily-report exports put a number row first (headers in row 1). We scan the
 * first few rows for the header row (the one containing NMEmpId / Emp Id / Name)
 * instead of hardcoding an offset.
 */

const XLSX = require("xlsx");

function _str(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

// Column name patterns -> canonical field. First match wins.
const COLS = {
  emp_id: ["nmempid", "emp id", "empid", "employee id"],
  name: ["name(asperaadhar)", "name (as per aadhar", "name(asperaadhaar)", "name"],
  status: ["status"],
  role: ["role", "position"],
  designation: ["designation"],
  branch: ["branch"],
  area: ["area name"],
  area_manager: ["area manager"],
  division: ["division name"],
  division_manager: ["division manager"],
  region: ["region name", "region"],
  department: ["department"],
  mobile: ["personalmobile", "personal mobile", "mobile"],
  doj: ["date of joining"],
  gender: ["gender"],
};

function _looksLikeHeader(row) {
  if (!row || !row.length) return false;
  const cells = row.map((c) => _str(c).toLowerCase());
  const hasId = cells.some((c) => c.includes("nmempid") || c === "empid" || c.includes("emp id"));
  const hasName = cells.some((c) => c.includes("name"));
  return hasId && hasName;
}

function _findHeaderRow(rows) {
  const scan = Math.min(rows.length, 8);
  for (let r = 0; r < scan; r++) {
    if (_looksLikeHeader(rows[r])) return r;
  }
  return -1;
}

function _buildColMap(headerRow) {
  const headers = headerRow.map((h) => _str(h).toLowerCase());
  const map = {};
  for (const [field, patterns] of Object.entries(COLS)) {
    map[field] = -1;
    for (const p of patterns) {
      // "Date of Joining" must not match "Group Date of Joining".
      const idx = headers.findIndex(
        (h) => h.includes(p) && !(field === "doj" && h.includes("group"))
      );
      if (idx >= 0) { map[field] = idx; break; }
    }
  }
  return map;
}

function _fmt(y, m, d) {
  const p = (n) => String(n).padStart(2, "0");
  return `${y}-${p(m)}-${p(d)}`;
}

function _toDate(v) {
  if (v === null || v === undefined || v === "") return null;
  // cellDates gives a Date at LOCAL midnight — read local parts (no TZ shift).
  if (v instanceof Date && !isNaN(v)) return _fmt(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === "number" && v > 1000) {
    // Excel serial date — anchor at the Excel epoch in UTC, read UTC parts.
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return isNaN(d) ? null : _fmt(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }
  const s = String(v).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/); // already date-only
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const d = new Date(s);
  return isNaN(d) ? null : _fmt(d.getFullYear(), d.getMonth() + 1, d.getDate());
}

/**
 * Parse an xlsx buffer into canonical employee rows.
 * @param {Buffer} buffer
 * @returns {{ rows: Array<object>, working: Array<object>, sheet: string, total: number }}
 * @throws {Error} on missing/empty sheet or missing emp_id column
 */
function parseStaffWorkbook(buffer) {
  // No cellDates: keep date cells as raw Excel serials and convert them
  // deterministically (UTC epoch) to avoid local-timezone off-by-one shifts.
  const wb = XLSX.read(buffer, { type: "buffer" });
  // Prefer a sheet whose name hints at staff/working; else first sheet.
  const sheetName =
    wb.SheetNames.find((s) => /work|staff|sheet1/i.test(s)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("Workbook has no readable sheet");

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) throw new Error("Sheet is empty");

  const headerIdx = _findHeaderRow(rows);
  if (headerIdx < 0) throw new Error("Could not find a header row (need NMEmpId + Name columns)");

  const col = _buildColMap(rows[headerIdx]);
  if (col.emp_id < 0) throw new Error("Could not find the Employee ID (NMEmpId) column");

  const get = (row, field) => (col[field] >= 0 ? _str(row[col[field]]) : "");

  const out = [];
  const seen = new Set();
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const empId = get(row, "emp_id");
    if (!empId || seen.has(empId)) continue;
    seen.add(empId);
    out.push({
      emp_id: empId,
      name: get(row, "name") || null,
      role: get(row, "role") || null,
      designation: get(row, "designation") || null,
      branch: get(row, "branch") || null,
      area: get(row, "area") || null,
      area_manager: get(row, "area_manager") || null,
      division: get(row, "division") || null,
      division_manager: get(row, "division_manager") || null,
      region: get(row, "region") || null,
      department: get(row, "department") || null,
      mobile: get(row, "mobile") || null,
      doj: col.doj >= 0 ? _toDate(row[col.doj]) : null,
      gender: get(row, "gender") || null,
      status: get(row, "status") || null,
    });
  }

  const working = out.filter((e) => String(e.status || "").trim().toLowerCase() === "working");
  return { rows: out, working, sheet: sheetName, total: out.length };
}

module.exports = { parseStaffWorkbook };
