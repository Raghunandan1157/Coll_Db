/**
 * collection.js — Collection tab renderer for the employee dashboard.
 * Mirrors portfolio.js exactly: sheet toggle, on-date preview,
 * rank/performance badges, DPD buckets, NPA details, and drill-down.
 *
 * Data source: default workbook (getWorkbookWithFallback)
 * Container: collectionContent
 */
(function () {
  /* ========== Column Indices (0-indexed from Excel) ========== */
  var COL = {
    name: 1,
    regDemand: 2, regCollection: 3, regFtod: 4, regPct: 5,
    dpd1Demand: 6, dpd1Collection: 7, dpd1Balance: 8, dpd1Pct: 9,
    dpd2Demand: 10, dpd2Collection: 11, dpd2Balance: 12, dpd2Pct: 13,
    pnpaDemand: 14, pnpaCollection: 15, pnpaBalance: 16, pnpaPct: 17,
    npaDemand: 18, npaActAcct: 19, npaActAmt: 20, npaClsAcct: 21, npaClsAmt: 22,
    rank: 23, performance: 24
  };

  /* ========== Hierarchy (same as employee.js / portfolio.js) ========== */
  var HIERARCHY = {"ANDRA PRADESH":{"KADAPA":["BUDWAL","DHARMAVARAM","KADAPA","KADIRI"]},"DHARWAD":{"BADAMI":["BADAMI","GAJENDRAGAD","NARAGUNDA","RAMDURGA"],"BALLARI":["BALLARI","KUDATHINI","SANDURU","SIRUGUPPA"],"BELAGAVI":["BAILHONGAL","BELAGAVI","GOKAK","KITTUR","YARAGATTI"],"CHIKKODI":["ATHANI","CHIKKODI","MUDALAGI","NIPPANI"],"DAVANAGERE":["DAVANAGERE","HARIHARA","HONNALI","SANTHEBENNURU"],"DHARWAD":["DHARWAD","HUBLI","HUBLI-2","KALGHATGI"],"GADAG":["GADAG","LAXMESHWAR","MUNDARAGI"],"KUDLIGI":["HARAPANAHALLI","KHANAHOSAHALLI","KOTTURU","KUDLIGI"],"VIJAYANAGARA":["HAGARIBOMMANAHALLI","HOSPET","HUVENAHADAGALLI"]},"KALBURGI":{"BIDAR":["AURAD","BHALKI","BIDAR","BIDAR-2"],"HUMNABAD":["BASAVAKALYAN","HULSOOR","HUMNABAD","KAMALAPURA"],"INDI":["ALMEL","CHADCHAN","INDI"],"KALBURGI":["AFZALPUR","ALAND","JEVARGI","KALABURAGI","KALBURGI-2"],"KUSHTAGI":["BAGALKOT","GANGAVATHI","HUNGUND","KOPPAL","KUSHTAGI"],"LINGSUGUR":["DEVADURGA","LINGSUGUR","MANVI","RAICHUR","SINDHNUR","SIRWAR"],"SEDAM":["CHINCHOLI","KALAGI","SEDAM","SHAHAPUR","YADGIR"],"VIJAYAPURA":["BILAGI","JAMAKHANDI","LOKAPUR","MUDDEBIHAL","SINDAGI","TALIKOTI","TIKOTA","VIJAYAPUR"]},"TELANGANA":{"MAHABOOBNAGAR":["GADWAL","MAHABUB NAGAR","MARIKAL","TANDUR"],"SANGAREDDY":["KODANGAL","NARAYANKHED","SANGAREDDY","ZAHEERABAD"]},"TUMKUR":{"BENGALORE -RURAL":["DABUSPET","DODDABALLAPURA","GOWRIBIDANUR"],"BENGALORE -URBAN":["CHANDAPURA","HEBBAL","J P NAGAR","KENGERI"],"CHIKKABALLAPUR":["BAGEPALLI","CHIKBALLAPURA","CHINTAMANI","DEVANAHALLI","SRINIVASPURA"],"CHIKKAMAGALURU":["CHIKKAMAGALURU","MUDIGERE","NR PURA"],"CHITRADURGA":["CHALLAKERE","CHITRADURGA","HIRIYUR","JAGALORE"],"HOLALKERE":["CHANNAGIRI","HOLAKERE","HOSADURGA"],"KADUR":["AJJAMPURA","KADUR","PANCHANHALLI","TARIKERE"],"KOLAR":["BANGARPET","BETHAMANGALA","KOLAR","MALUR"],"TIPTUR":["CHIKKANAYAKANAHALLI","GUBBI","HULIYAR","TIPTUR","TUREVEKERE"],"TUMKUR":["KORATAGERE","KUNIGAL","MADHUGIRI","SIRA","TUMKUR"]}};

  /* ========== Session ========== */
  var session = getEmployeeSession();

  /* ========== Formatters ========== */
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    if (typeof v === 'string') return v.trim() || '-';
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return v.toLocaleString('en-IN');
      return v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(v);
  }
  function fmtPct(v) {
    if (v == null || v === '' || v === '-') return '-';
    if (typeof v === 'number') return (v * 100).toFixed(2) + '%';
    return String(v);
  }
  function numVal(v) { return (typeof v === 'number' && isFinite(v)) ? v : 0; }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }

  /* ========== Fuzzy name matching ========== */
  function normalizeForMatch(s) {
    return s.toUpperCase().replace(/[\s\-\u00A0]+/g, '').trim();
  }
  function fuzzyMatch(a, b) {
    var an = normalizeForMatch(a);
    var bn = normalizeForMatch(b);
    if (an === bn) return true;
    if (an.length >= 8 && bn.length >= 8 && Math.abs(an.length - bn.length) <= 1) {
      return levenshtein(an, bn) <= 1;
    }
    return false;
  }
  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var prev = [], curr = [];
    for (var j = 0; j <= a.length; j++) prev[j] = j;
    for (var i = 1; i <= b.length; i++) {
      curr[0] = i;
      for (var j2 = 1; j2 <= a.length; j2++) {
        if (b[i - 1] === a[j2 - 1]) curr[j2] = prev[j2 - 1];
        else curr[j2] = 1 + Math.min(prev[j2 - 1], prev[j2], curr[j2 - 1]);
      }
      var tmp = prev; prev = curr; curr = tmp;
    }
    return prev[a.length];
  }

  /* ========== Sheet Detection ========== */
  function categorizeSheets(sheetNames) {
    var result = { overall: null, overallOnDate: null, overallTom: null,
                   fy: null, fyOnDate: null, fyTom: null };
    for (var i = 0; i < sheetNames.length; i++) {
      var n = sheetNames[i].toUpperCase();
      if (n.indexOf('TOM') >= 0 && n.indexOf('FY') >= 0) result.fyTom = sheetNames[i];
      else if (n.indexOf('TOM') >= 0) result.overallTom = sheetNames[i];
      else if (n.indexOf('FY') >= 0 && n.indexOf('ON') >= 0) result.fyOnDate = sheetNames[i];
      else if (n.indexOf('ON') >= 0 && n.indexOf('DATE') >= 0) result.overallOnDate = sheetNames[i];
      else if (n.indexOf('FY') >= 0) result.fy = sheetNames[i];
      else if (n.indexOf('OVERALL') >= 0) result.overall = sheetNames[i];
    }
    return result;
  }

  /* ========== Data Extraction ========== */
  function parseRows(workbook, sheetName) {
    if (!sheetName) return null;
    var ws = workbook.Sheets[sheetName];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1 });
  }

  function extractReportDate(rows) {
    for (var r = 0; r < Math.min(5, rows.length); r++) {
      var row = rows[r];
      if (!row) continue;
      for (var c = 0; c < row.length; c++) {
        var v = String(row[c] || '');
        var m = v.match(/as on\s+(\d{2}-\d{2}-\d{4})/i);
        if (m) return m[1];
      }
    }
    return null;
  }

  /* ---------- Section parsing ---------- */
  function getAllSectionRows(rows, sectionType) {
    var results = [];
    var currentSection = null;
    var foundTarget = false;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase();
      var c1 = String(row[1] || '').trim().toUpperCase();
      var hdr = c0 + ' ' + c1;

      if (hdr.includes('REGION') && (hdr.includes('WISE') || hdr.includes('NAME'))) {
        if (currentSection === sectionType && foundTarget) break;
        currentSection = 'region'; continue;
      }
      if (hdr.includes('DISTRICT') && (hdr.includes('WISE') || hdr.includes('NAME'))) {
        if (currentSection === sectionType && foundTarget) break;
        currentSection = 'district'; continue;
      }
      if (hdr.includes('BRANCH') && !hdr.includes('OFFICER') && (hdr.includes('WISE') || hdr.includes('NAME'))) {
        if (currentSection === sectionType && foundTarget) break;
        currentSection = 'branch'; continue;
      }
      if (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL')) {
        if (currentSection === sectionType && foundTarget) break;
        currentSection = null; continue;
      }
      if (currentSection !== sectionType) continue;
      if (c0 === 'SL' || c0 === 'SL.' || c1 === 'SL' || c1 === 'SL.') continue;
      if (/^TOTAL/i.test(c0)) continue;

      var name = String(row[1] || '').trim().toUpperCase();
      if (!name || /^\d+$/.test(name)) name = String(row[0] || '').trim().toUpperCase();
      if (!name || /^\d+$/.test(name)) continue;
      if (/^(SL|NAME|OFFICER|DEMAND|COLLECTION|FTOD|BALANCE|ACCOUNT|AMOUNT)/i.test(name)) continue;

      results.push({ name: name, row: row });
      foundTarget = true;
    }
    return results;
  }

  /* ---------- Find row by role ---------- */
  function findRoleRow(rows, role, location) {
    var currentSection = null;
    var locationUpper = (location || '').toUpperCase().trim();
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase();
      var c1 = String(row[1] || '').trim().toUpperCase();
      var hdr = c0 + ' ' + c1;

      if (hdr.includes('REGION') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'region'; continue; }
      if (hdr.includes('DISTRICT') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'district'; continue; }
      if (hdr.includes('BRANCH') && !hdr.includes('OFFICER') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'branch'; continue; }

      if (role === 'CEO') {
        if (currentSection === 'region' && (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL'))) return row;
        continue;
      }

      if (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL')) { currentSection = null; continue; }

      var target = role === 'RM' ? 'region' : role === 'DM' ? 'district' : role === 'BM' ? 'branch' : null;
      if (currentSection !== target) continue;

      var nameInRow = String(row[1] || '').trim().toUpperCase();
      if (!nameInRow || /^\d+$/.test(nameInRow)) nameInRow = String(row[0] || '').trim().toUpperCase();
      if (fuzzyMatch(nameInRow, locationUpper)) return row;
    }
    return null;
  }

  /* ---------- Find employee row (FO) ---------- */
  function findEmpRow(rows) {
    if (session.role && session.role !== 'FO') {
      return findRoleRow(rows, session.role, session.location);
    }
    var needle = (session.name || '').toUpperCase().trim();
    var needleId = String(session.id || '').trim();
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      for (var c = 0; c < Math.min(row.length, 5); c++) {
        var cv = String(row[c] || '').trim();
        if (cv.toUpperCase() === needle || cv === needleId) return row;
      }
    }
    return null;
  }

  /* ---------- Find on-date row (narrow sheets) ---------- */
  function findOnDateRow(rows) {
    if (!rows) return null;
    var currentSection = null;
    var locationUpper = (session.location || '').toUpperCase().trim();

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase();
      var c1 = String(row[1] || '').trim().toUpperCase();
      var hdr = c0 + ' ' + c1;

      if (hdr.includes('REGION') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'region'; continue; }
      if (hdr.includes('DISTRICT') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'district'; continue; }
      if (hdr.includes('BRANCH') && !hdr.includes('OFFICER') && (hdr.includes('WISE') || hdr.includes('NAME'))) { currentSection = 'branch'; continue; }

      if (session.role === 'CEO') {
        if (currentSection === 'region' && (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL'))) return row;
        continue;
      }

      if (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL')) { currentSection = null; continue; }

      var target = session.role === 'RM' ? 'region' : session.role === 'DM' ? 'district' : session.role === 'BM' ? 'branch' : null;
      if (currentSection !== target) continue;

      var nameInRow = String(row[1] || '').trim().toUpperCase();
      if (!nameInRow || /^\d+$/.test(nameInRow)) nameInRow = String(row[0] || '').trim().toUpperCase();
      if (target && fuzzyMatch(nameInRow, locationUpper)) return row;
    }
    return null;
  }

  /* ---------- Find children for drill-down ---------- */
  function findChildren(rows, role, location) {
    if (role === 'FO') return [];
    var children = [];
    var locationUpper = (location || '').toUpperCase().trim();

    if (role === 'CEO') {
      children = getAllSectionRows(rows, 'region');
    } else if (role === 'RM') {
      var hierarchyDistricts = null;
      for (var rk in HIERARCHY) {
        if (fuzzyMatch(rk, locationUpper)) { hierarchyDistricts = Object.keys(HIERARCHY[rk]); break; }
      }
      if (hierarchyDistricts) {
        var allDistricts = getAllSectionRows(rows, 'district');
        for (var i = 0; i < allDistricts.length; i++) {
          for (var j = 0; j < hierarchyDistricts.length; j++) {
            if (fuzzyMatch(allDistricts[i].name, hierarchyDistricts[j])) { children.push(allDistricts[i]); break; }
          }
        }
      }
    } else if (role === 'DM') {
      var hierarchyBranches = null;
      for (var rk2 in HIERARCHY) {
        for (var dk in HIERARCHY[rk2]) {
          if (fuzzyMatch(dk, locationUpper)) { hierarchyBranches = HIERARCHY[rk2][dk]; break; }
        }
        if (hierarchyBranches) break;
      }
      if (hierarchyBranches) {
        var allBranches = getAllSectionRows(rows, 'branch');
        for (var i2 = 0; i2 < allBranches.length; i2++) {
          for (var j2 = 0; j2 < hierarchyBranches.length; j2++) {
            if (fuzzyMatch(allBranches[i2].name, hierarchyBranches[j2])) { children.push(allBranches[i2]); break; }
          }
        }
      }
    } else if (role === 'BM') {
      children = findOfficersForBranch(rows, locationUpper);
    }

    children.sort(function (a, b) {
      var pctA = numVal(a.row[COL.regDemand]) > 0 ? numVal(a.row[COL.regCollection]) / numVal(a.row[COL.regDemand]) : 0;
      var pctB = numVal(b.row[COL.regDemand]) > 0 ? numVal(b.row[COL.regCollection]) / numVal(b.row[COL.regDemand]) : 0;
      return pctB - pctA;
    });
    return children;
  }

  /* ---------- Find officers for branch ---------- */
  function findOfficersForBranch(rows, branchUpper) {
    var officers = [];
    var inOfficerSection = false;
    var foundBranch = false;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase();
      var c1 = String(row[1] || '').trim().toUpperCase();
      var hdr = c0 + ' ' + c1;
      if (hdr.includes('BRANCH') && hdr.includes('OFFICER') && hdr.includes('NAME')) { inOfficerSection = true; continue; }
      if (!inOfficerSection) continue;
      if (c0 === 'EMP ID' || c1 === 'BRANCH / OFFICER NAME') continue;
      if (c0 === '' && c1 === '') continue;
      if (c1 === 'DEMAND' || c1 === 'COLLECTION') continue;

      var hasEmpId = c0.length > 0 && /^[A-Z]{2}\d+/.test(c0);
      if (!hasEmpId && c1.length > 0 && row.length > 3) {
        if (fuzzyMatch(c1, branchUpper)) foundBranch = true;
        else if (foundBranch) break;
        continue;
      }
      if (hasEmpId && foundBranch) {
        officers.push({ name: String(row[1] || '').trim(), empId: String(row[0] || '').trim(), row: row });
      }
    }
    return officers;
  }

  /* ---------- Product boundaries ---------- */
  function detectProductBounds(rows) {
    var result = {};
    var boStarts = [];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var c0 = String(row[0] || '').toUpperCase().trim();
      var c1 = String(row[1] || '').toUpperCase().trim();
      if (c1.includes('REGION WISE') && c1.includes('REPORT')) {
        var prod = c1.includes('IGL') ? 'igl' : c1.includes('FIG') ? 'fig' : c1.includes('IL') ? 'il' : null;
        if (prod) { if (!result[prod]) result[prod] = {}; result[prod].regStart = r; }
      }
      if (c0.includes('BRANCH') && c0.includes('OFFICER') && c0.includes('COLLECTION REPORT')) {
        var prod2 = c0.includes('- IGL') ? 'igl' : c0.includes('- FIG') ? 'fig' : c0.includes('- IL') ? 'il' : null;
        if (prod2) { if (!result[prod2]) result[prod2] = {}; if (result[prod2].boStart == null) { result[prod2].boStart = r; boStarts.push({ key: prod2, row: r }); } }
      }
    }
    boStarts.sort(function (a, b) { return a.row - b.row; });
    for (var i = 0; i < boStarts.length; i++) {
      result[boStarts[i].key].boEnd = (i + 1 < boStarts.length) ? boStarts[i + 1].row : rows.length;
    }
    var products = ['igl', 'fig', 'il'];
    for (var p = 0; p < products.length; p++) {
      var k = products[p];
      if (result[k] && result[k].regStart != null) {
        result[k].regEnd = result[k].regStart + 50;
        for (var r2 = result[k].regStart; r2 < Math.min(result[k].regStart + 50, rows.length); r2++) {
          if (rows[r2] && String(rows[r2][1] || '').trim().toUpperCase() === 'GRAND TOTAL') { result[k].regEnd = r2 + 1; break; }
        }
      }
    }
    return result;
  }

  /* ---------- Product row finder ---------- */
  function findProductRow(rows, bounds, product) {
    var b = bounds[product];
    if (!b) return null;
    if (session.role === 'CEO' && b.regStart != null) {
      for (var r = b.regStart; r < b.regEnd; r++) {
        if (rows[r] && String(rows[r][1] || '').trim().toUpperCase() === 'GRAND TOTAL') return rows[r];
      }
      return null;
    }
    if (session.role === 'RM' && b.regStart != null) {
      var loc = (session.location || '').toUpperCase().trim();
      for (var r2 = b.regStart; r2 < b.regEnd; r2++) {
        var c1 = String(rows[r2] && rows[r2][1] || '').trim();
        if (c1 && fuzzyMatch(c1, loc)) return rows[r2];
      }
      return null;
    }
    if (session.role === 'DM') return aggregateDistrictProduct(rows, bounds, product, session.location);
    if (session.role === 'BM') {
      var brU = (session.location || '').toUpperCase().trim();
      for (var r3 = b.boStart; r3 < b.boEnd; r3++) {
        var row = rows[r3]; if (!row) continue;
        var ec0 = String(row[0] || '').trim(), ec1 = String(row[1] || '').trim();
        if (!ec0 && ec1 && fuzzyMatch(ec1, brU)) return row;
      }
      return null;
    }
    var nId = String(session.id || '').trim().toUpperCase();
    var nName = (session.name || '').toUpperCase().trim();
    for (var r4 = b.boStart; r4 < b.boEnd; r4++) {
      var row2 = rows[r4]; if (!row2) continue;
      var rc0 = String(row2[0] || '').trim().toUpperCase();
      var rc1 = String(row2[1] || '').trim().toUpperCase();
      if ((nId && rc0 === nId) || (nName && rc1 === nName)) return row2;
    }
    return null;
  }

  function aggregateDistrictProduct(rows, bounds, product, districtName) {
    var b = bounds[product]; if (!b) return null;
    var distU = (districtName || '').toUpperCase().trim();
    var branches = null;
    for (var rk in HIERARCHY) {
      for (var dk in HIERARCHY[rk]) { if (fuzzyMatch(dk, distU)) { branches = HIERARCHY[rk][dk]; break; } }
      if (branches) break;
    }
    if (!branches || !branches.length) return null;
    var bSet = {};
    for (var i = 0; i < branches.length; i++) bSet[normalizeForMatch(branches[i])] = true;
    var sum = new Array(25);
    sum[0] = null; sum[1] = districtName;
    for (var c = 2; c < 25; c++) sum[c] = 0;
    var found = false;
    for (var r = b.boStart; r < b.boEnd; r++) {
      var row = rows[r]; if (!row) continue;
      var ec0 = String(row[0] || '').trim(), ec1 = String(row[1] || '').trim();
      if (!ec0 && ec1 && bSet[normalizeForMatch(ec1)]) {
        found = true;
        for (var c2 = 2; c2 < Math.min(row.length, 24); c2++) {
          if (typeof row[c2] === 'number') sum[c2] += row[c2];
        }
      }
    }
    if (!found) return null;
    if (sum[2] > 0) sum[5] = sum[3] / sum[2];
    if (sum[6] > 0) sum[9] = sum[7] / sum[6];
    if (sum[10] > 0) sum[13] = sum[11] / sum[10];
    if (sum[14] > 0) sum[17] = sum[15] / sum[14];
    return sum;
  }

  /* ---------- Product children ---------- */
  function findProductChildren(rows, bounds, product) {
    var b = bounds[product]; if (!b) return [];
    if (session.role === 'CEO' && b.regStart != null) {
      var ch = [];
      for (var r = b.regStart; r < b.regEnd; r++) {
        var row = rows[r]; if (!row) continue;
        var c1 = String(row[1] || '').trim();
        if (!c1) continue;
        var c1u = c1.toUpperCase();
        if (c1u === 'GRAND TOTAL' || c1u.includes('REGION') || c1u.includes('REPORT') || c1u === 'DEMAND' || c1u === 'COLLECTION' || c1u === 'FTOD') continue;
        if (typeof row[2] !== 'number') continue;
        ch.push({ name: c1u, row: row });
      }
      return ch;
    }
    if (session.role === 'RM') {
      var locU = (session.location || '').toUpperCase().trim();
      var rd = null;
      for (var rk in HIERARCHY) { if (fuzzyMatch(rk, locU)) { rd = HIERARCHY[rk]; break; } }
      if (!rd) return [];
      var ch2 = [];
      for (var dk in rd) {
        var agg = aggregateDistrictProduct(rows, bounds, product, dk);
        if (agg) ch2.push({ name: dk.toUpperCase(), row: agg });
      }
      return ch2;
    }
    if (session.role === 'DM') {
      var locU2 = (session.location || '').toUpperCase().trim();
      var br = null;
      for (var rk2 in HIERARCHY) { for (var dk2 in HIERARCHY[rk2]) { if (fuzzyMatch(dk2, locU2)) { br = HIERARCHY[rk2][dk2]; break; } } if (br) break; }
      if (!br) return [];
      var bSet = {};
      for (var i = 0; i < br.length; i++) bSet[normalizeForMatch(br[i])] = true;
      var ch3 = [];
      for (var r2 = b.boStart; r2 < b.boEnd; r2++) {
        var row2 = rows[r2]; if (!row2) continue;
        var ec0 = String(row2[0] || '').trim(), ec1 = String(row2[1] || '').trim();
        if (!ec0 && ec1 && bSet[normalizeForMatch(ec1)]) ch3.push({ name: ec1.toUpperCase(), row: row2 });
      }
      return ch3;
    }
    if (session.role === 'BM') {
      var brU = (session.location || '').toUpperCase().trim();
      var off = [];
      var fb = false;
      for (var r3 = b.boStart; r3 < b.boEnd; r3++) {
        var row3 = rows[r3]; if (!row3) continue;
        var ec02 = String(row3[0] || '').trim(), ec12 = String(row3[1] || '').trim();
        var u0 = ec02.toUpperCase(), u1 = ec12.toUpperCase();
        if (u0 === 'EMP ID' || u1 === 'BRANCH / OFFICER NAME' || u1 === 'DEMAND' || u1 === 'COLLECTION') continue;
        var hasId = ec02.length > 0 && /^[A-Z]{2}\d+/.test(u0);
        if (!hasId && ec12.length > 0 && row3.length > 3) {
          if (fuzzyMatch(u1, brU)) fb = true; else if (fb) break;
          continue;
        }
        if (hasId && fb) off.push({ name: ec12.trim(), empId: ec02, row: row3 });
      }
      return off;
    }
    return [];
  }

  /* ===================== RENDERING ===================== */
  var _collState = { view: 'overall', product: 'all', rows: null, fyRows: null,
                     onDateRows: null, tomRows: null, fyOnDateRows: null, fyTomRows: null,
                     bounds: null, fyBounds: null, sheetMap: null };

  function getActiveRows() { return _collState.view === 'fy' ? _collState.fyRows : _collState.rows; }
  function getActiveBounds() { return _collState.view === 'fy' ? _collState.fyBounds : _collState.bounds; }
  function getActiveOnDate() { return _collState.view === 'fy' ? _collState.fyOnDateRows : _collState.onDateRows; }
  function getActiveTom() { return _collState.view === 'fy' ? _collState.fyTomRows : _collState.tomRows; }

  /* ---------- Main render ---------- */
  function renderCollection() {
    var rows = getActiveRows();
    var container = document.getElementById('collectionContent');
    if (!rows) { container.innerHTML = noDataHtml(); return; }

    var empRow;
    if (_collState.product === 'all') {
      empRow = findEmpRow(rows);
    } else {
      empRow = findProductRow(rows, getActiveBounds(), _collState.product);
    }

    if (!empRow) { container.innerHTML = viewToggleHtml() + productPillsHtml() + noDataHtml('No data found for this view'); return; }

    var html = '';

    // View toggle
    html += viewToggleHtml();

    // Report date & Rank/Performance header
    var reportDate = extractReportDate(rows);
    var rank = empRow[COL.rank];
    var perf = empRow[COL.performance];
    html += headerCardHtml(reportDate, rank, perf);

    // Snapshot
    html += snapshotHtml(empRow);

    // Product pills
    html += productPillsHtml();

    // Regular demand card
    html += regDemandHtml(empRow);

    // DPD Bucket cards
    html += dpdBucketsHtml(empRow);

    // NPA card
    html += npaCardHtml(empRow);

    // On-Date & Tomorrow
    var onDateData = getActiveOnDate();
    var tomData = getActiveTom();
    if (onDateData || tomData) {
      html += onDateSectionHtml(onDateData, tomData);
    }

    // Sub-units
    if (session.role && session.role !== 'FO') {
      var childRoleMap = { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
      var childRole = childRoleMap[session.role];
      if (childRole) {
        var children;
        if (_collState.product === 'all') {
          children = findChildren(rows, session.role, session.location);
        } else {
          children = findProductChildren(rows, getActiveBounds(), _collState.product);
          children.sort(function (a, b) {
            var pA = numVal(a.row[COL.regDemand]) > 0 ? numVal(a.row[COL.regCollection]) / numVal(a.row[COL.regDemand]) : 0;
            var pB = numVal(b.row[COL.regDemand]) > 0 ? numVal(b.row[COL.regCollection]) / numVal(b.row[COL.regDemand]) : 0;
            return pB - pA;
          });
        }
        if (children.length) html += subUnitsHtml(children, childRole);
      }
    }

    container.innerHTML = html;
    attachHandlers();
  }

  /* ---------- HTML builders ---------- */
  function noDataHtml(msg) {
    return '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="font-size:36px;margin-bottom:12px;">&#128202;</div>' +
      '<div style="color:#6B7A99;font-size:14px;">' + (msg || 'No collection data uploaded yet.') + '</div></div>';
  }

  function viewToggleHtml() {
    var ov = _collState.view === 'overall';
    return '<div class="pf-view-toggle emp-fade">' +
      '<button class="pf-view-btn' + (ov ? ' active' : '') + '" data-coll-view="overall">OverAll</button>' +
      '<button class="pf-view-btn' + (!ov ? ' active' : '') + '" data-coll-view="fy">FY 25-26</button>' +
    '</div>';
  }

  function headerCardHtml(reportDate, rank, perf) {
    var perfText = String(perf || '').trim();
    var perfUpper = perfText.toUpperCase();
    var isAbove = perfUpper.includes('ABOVE');
    var isBelow = perfUpper.includes('BELOW');
    var perfColor = isAbove ? '#34D399' : isBelow ? '#F87171' : '#6B7A99';
    var perfIcon = isAbove ? '&#9650;' : isBelow ? '&#9660;' : '';
    var perfLabel = isAbove ? 'Above Average' : isBelow ? 'Below Average' : (perfText || '-');

    var rankVal = (rank != null && rank !== '-' && rank !== '') ? rank : '-';

    return '<div class="pf-header-card emp-fade">' +
      '<div class="pf-header-top">' +
        '<div class="pf-report-date">' +
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' +
          '<span>' + (reportDate || 'Month End') + '</span>' +
        '</div>' +
        '<div class="pf-badges">' +
          (rankVal !== '-' ? '<div class="pf-rank-badge">#' + esc(String(rankVal)) + '</div>' : '') +
          '<div class="pf-perf-badge" style="color:' + perfColor + ';border-color:' + perfColor + '30">' +
            perfIcon + ' ' + esc(perfLabel) +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function snapshotHtml(row) {
    var totalDem = numVal(row[COL.regDemand]) + numVal(row[COL.dpd1Demand]) + numVal(row[COL.dpd2Demand]) + numVal(row[COL.pnpaDemand]) + numVal(row[COL.npaDemand]);
    var totalCol = numVal(row[COL.regCollection]) + numVal(row[COL.dpd1Collection]) + numVal(row[COL.dpd2Collection]) + numVal(row[COL.pnpaCollection]) + numVal(row[COL.npaActAcct]);
    var overallPct = totalDem > 0 ? ((totalCol / totalDem) * 100).toFixed(1) : '0';

    return '<div class="emp-snapshot emp-fade">' +
      '<div class="emp-snapshot-label">Snapshot</div>' +
      '<div class="emp-snapshot-row">' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Demand</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalDem) + '</div>' +
        '</div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Collection</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalCol) + '</div>' +
        '</div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Overall %</div>' +
          '<div class="emp-snap-val" style="color:#4F8CFF;font-size:24px;">' + overallPct + '%</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function productPillsHtml() {
    var p = _collState.product;
    return '<div class="emp-product-filter">' +
      '<button class="emp-product-pill' + (p === 'all' ? ' active' : '') + '" data-coll-product="all">All</button>' +
      '<button class="emp-product-pill' + (p === 'igl' ? ' active' : '') + '" data-coll-product="igl">IGL</button>' +
      '<button class="emp-product-pill' + (p === 'fig' ? ' active' : '') + '" data-coll-product="fig">FIG</button>' +
      '<button class="emp-product-pill' + (p === 'il' ? ' active' : '') + '" data-coll-product="il">IL</button>' +
    '</div>';
  }

  function regDemandHtml(row) {
    var dem = numVal(row[COL.regDemand]);
    var col = numVal(row[COL.regCollection]);
    var ftod = numVal(row[COL.regFtod]);
    var pct = row[COL.regPct];
    var pctVal = typeof pct === 'number' ? (pct * 100) : 0;
    var noData = (dem === 0 && col === 0);
    var pctColor = noData ? '#6B7A99' : pctVal >= 99 ? '#34D399' : pctVal >= 95 ? '#FBBF24' : '#F87171';
    var barW = Math.min(pctVal, 100);

    return '<div class="emp-data-section">' +
      '<div class="emp-section-title">Regular Demand vs Collection</div>' +
      '<div class="pf-reg-card emp-fade">' +
        '<div class="pf-reg-main">' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val">' + fmtNum(dem) + '</div>' +
            '<div class="pf-reg-lbl">Demand</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:#4F8CFF">' + fmtNum(col) + '</div>' +
            '<div class="pf-reg-lbl">Collection</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:#FB923C">' + fmtNum(ftod) + '</div>' +
            '<div class="pf-reg-lbl">FTOD</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:' + pctColor + '">' + (noData ? '-' : pctVal.toFixed(2) + '%') + '</div>' +
            '<div class="pf-reg-lbl">Collection %</div>' +
          '</div>' +
        '</div>' +
        '<div class="pf-progress-track">' +
          '<div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function dpdBucketsHtml(row) {
    var buckets = [
      { name: 'Bucket 1', range: '1 \u2014 30 DPD', color: '#34D399', d: COL.dpd1Demand, c: COL.dpd1Collection, b: COL.dpd1Balance, p: COL.dpd1Pct },
      { name: 'Bucket 2', range: '31 \u2014 60 DPD', color: '#FBBF24', d: COL.dpd2Demand, c: COL.dpd2Collection, b: COL.dpd2Balance, p: COL.dpd2Pct },
      { name: 'PNPA', range: 'Pre-NPA', color: '#FB923C', d: COL.pnpaDemand, c: COL.pnpaCollection, b: COL.pnpaBalance, p: COL.pnpaPct }
    ];

    var html = '<div class="emp-section-title" style="margin-top:20px;">DPD Bucket Allocation</div>';
    html += '<div class="emp-buckets">';

    for (var i = 0; i < buckets.length; i++) {
      var bk = buckets[i];
      var dem = numVal(row[bk.d]);
      var col = numVal(row[bk.c]);
      var bal = numVal(row[bk.b]);
      var pct = row[bk.p];
      var pctVal = typeof pct === 'number' ? (pct * 100) : 0;
      var noData = (dem === 0 && col === 0);
      var pctColor = noData ? '#6B7A99' : pctVal >= 50 ? '#34D399' : pctVal >= 25 ? '#FBBF24' : '#F87171';
      var barW = noData ? 0 : Math.min(pctVal, 100);

      html += '<div class="emp-bucket-card emp-fade" style="animation-delay:' + (0.1 + i * 0.05) + 's">' +
        '<div class="emp-bucket-indicator" style="background:' + bk.color + '"></div>' +
        '<div class="emp-bucket-info">' +
          '<div class="emp-bucket-name">' + bk.name + '</div>' +
          '<div class="emp-bucket-sub">' + bk.range + '</div>' +
        '</div>' +
        '<div class="emp-bucket-stats">' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:' + bk.color + '">' + fmtNum(dem) + '</div>' +
            '<div class="emp-bucket-lbl">Demand</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val emp-pos-val">' + fmtNum(col) + '</div>' +
            '<div class="emp-bucket-lbl">Collection</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:#6B7A99">' + fmtNum(bal) + '</div>' +
            '<div class="emp-bucket-lbl">Balance</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:' + pctColor + '">' + (noData ? '-' : pctVal.toFixed(2) + '%') + '</div>' +
            '<div class="emp-bucket-lbl">Coll %</div>' +
          '</div>' +
        '</div>' +
        (barW === 0
          ? '<div class="emp-bucket-bar" style="width:100%;height:2px;background:rgba(255,255,255,0.08)"></div>'
          : '<div class="emp-bucket-bar" style="width:' + Math.max(barW, 8) + '%;height:4px;opacity:0.85;background:' + bk.color + '"></div>') +
      '</div>';
    }

    html += '</div>';
    return html;
  }

  function npaCardHtml(row) {
    var demand = numVal(row[COL.npaDemand]);
    var actAcct = row[COL.npaActAcct];
    var actAmt = row[COL.npaActAmt];
    var clsAcct = row[COL.npaClsAcct];
    var clsAmt = row[COL.npaClsAmt];

    return '<div class="pf-npa-card emp-fade" style="animation-delay:0.25s">' +
      '<div class="pf-npa-header">' +
        '<div class="emp-bucket-indicator" style="background:#F87171"></div>' +
        '<div class="emp-bucket-info">' +
          '<div class="emp-bucket-name">NPA</div>' +
          '<div class="emp-bucket-sub">Non-Performing Assets</div>' +
        '</div>' +
        '<div class="pf-npa-demand">' +
          '<div class="emp-bucket-val" style="color:#F87171">' + fmtNum(demand) + '</div>' +
          '<div class="emp-bucket-lbl">Demand</div>' +
        '</div>' +
      '</div>' +
      '<div class="pf-npa-grid">' +
        '<div class="pf-npa-section">' +
          '<div class="pf-npa-section-title" style="color:#34D399">Activation</div>' +
          '<div class="pf-npa-row">' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(actAcct) + '</div><div class="pf-npa-lbl">Account</div></div>' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(actAmt) + '</div><div class="pf-npa-lbl">Amount</div></div>' +
          '</div>' +
        '</div>' +
        '<div class="pf-npa-divider"></div>' +
        '<div class="pf-npa-section">' +
          '<div class="pf-npa-section-title" style="color:#4F8CFF">Closure</div>' +
          '<div class="pf-npa-row">' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(clsAcct) + '</div><div class="pf-npa-lbl">Account</div></div>' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(clsAmt) + '</div><div class="pf-npa-lbl">Amount</div></div>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function onDateSectionHtml(onDateRows, tomRows) {
    var odRow = onDateRows ? findOnDateRow(onDateRows) : null;
    var tmRow = tomRows ? findOnDateRow(tomRows) : null;

    var odDem = odRow ? odRow[2] : null;
    var odCol = odRow ? odRow[3] : null;
    var odPct = odRow ? odRow[4] : null;
    var tmDem = tmRow ? tmRow[2] : null;
    var tmCol = tmRow ? tmRow[3] : null;
    var tmPct = tmRow ? tmRow[4] : null;

    var html = '<div class="emp-section-title" style="margin-top:20px;">On-Date Status</div>' +
      '<div class="pf-ondate-grid emp-fade" style="animation-delay:0.3s">';

    // Today card
    html += '<div class="pf-ondate-card">' +
      '<div class="pf-ondate-label">Today</div>' +
      '<div class="pf-ondate-stats">' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val">' + fmtNum(odDem) + '</div><div class="pf-ondate-lbl">Demand</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#4F8CFF">' + fmtNum(odCol) + '</div><div class="pf-ondate-lbl">Collection</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#34D399">' + fmtPct(odPct) + '</div><div class="pf-ondate-lbl">Coll %</div></div>' +
      '</div>' +
    '</div>';

    // Tomorrow card
    html += '<div class="pf-ondate-card pf-ondate-tomorrow">' +
      '<div class="pf-ondate-label">Tomorrow</div>' +
      '<div class="pf-ondate-stats">' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val">' + fmtNum(tmDem) + '</div><div class="pf-ondate-lbl">Demand</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#4F8CFF">' + fmtNum(tmCol) + '</div><div class="pf-ondate-lbl">Collection</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#34D399">' + fmtPct(tmPct) + '</div><div class="pf-ondate-lbl">Coll %</div></div>' +
      '</div>' +
    '</div>';

    html += '</div>';
    return html;
  }

  function subUnitsHtml(children, childRole) {
    if (!children.length) return '';
    var roleLabel = { RM: 'Regions', DM: 'Districts', BM: 'Branches', FO: 'Officers' };
    var avType = { RM: 'region', DM: 'district', BM: 'branch', FO: 'officer' };
    var label = roleLabel[childRole] || 'Team';
    var av = avType[childRole] || 'branch';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label + '<span class="emp-team-count">' + children.length + '</span></div>';

    for (var i = 0; i < children.length; i++) {
      var ch = children[i];
      var dem = numVal(ch.row[COL.regDemand]);
      var col = numVal(ch.row[COL.regCollection]);
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      var pct = pctRaw.toFixed(2);
      var pctColor = pctRaw >= 95 ? '#34D399' : pctRaw >= 80 ? '#FBBF24' : '#F87171';
      var rank = ch.row[COL.rank];
      var rankBadge = (rank != null && rank !== '-' && rank !== '') ? '<span class="pf-sub-rank">#' + esc(String(rank)) + '</span>' : '';
      var initial = ch.name.charAt(0).toUpperCase();
      var dataAttr = childRole === 'FO'
        ? 'data-emp-id="' + esc(ch.empId || '') + '" data-emp-name="' + esc(ch.name) + '"'
        : 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(ch.name) + '"';

      html += '<div class="emp-sub-card" ' + dataAttr + '>' +
        '<div class="emp-sub-avatar ' + av + '">' + initial + '</div>' +
        '<div class="emp-sub-info">' +
          '<div class="emp-sub-name">' + esc(ch.name) + '</div>' +
          '<div class="emp-sub-meta">' +
            '<span>D: ' + fmtNum(dem) + '</span>' +
            '<span>C: ' + fmtNum(col) + '</span>' +
            rankBadge +
          '</div>' +
        '</div>' +
        '<div class="emp-sub-pct" style="color:' + pctColor + '">' + pct + '%</div>' +
        '<div class="emp-sub-arrow">&#8250;</div>' +
      '</div>';
    }

    html += '</div>';
    return html;
  }

  /* ========== Event Handlers ========== */
  var _collHandlersAttached = false;

  function attachHandlers() {
    var container = document.getElementById('collectionContent');
    if (!container) return;

    // View toggle
    container.querySelectorAll('[data-coll-view]').forEach(function (btn) {
      btn.onclick = function () {
        _collState.view = btn.dataset.collView;
        renderCollection();
      };
    });

    // Product pills
    container.querySelectorAll('[data-coll-product]').forEach(function (pill) {
      pill.onclick = function () {
        _collState.product = pill.dataset.collProduct;
        renderCollection();
      };
    });

    // Sub-unit drill-down (only attach once to avoid duplicates)
    if (!_collHandlersAttached) {
      _collHandlersAttached = true;
      container.addEventListener('click', function (ev) {
        var card = ev.target.closest('.emp-sub-card');
        if (!card) return;
        // Only handle if we're in the collection tab
        var collectionTab = document.getElementById('collectionTab');
        if (!collectionTab || !collectionTab.classList.contains('active')) return;

        card.style.background = 'rgba(79,140,255,0.12)';
        card.style.pointerEvents = 'none';
        var arrow = card.querySelector('.emp-sub-arrow');
        if (arrow) arrow.innerHTML = '<div style="width:16px;height:16px;border:2px solid rgba(79,140,255,0.2);border-top-color:#4F8CFF;border-radius:50%;animation:spin .7s linear infinite;"></div>';

        if (card.dataset.childRole) {
          pushRoleNav(card.dataset.childRole, card.dataset.childLocation);
          window.location.reload();
        } else if (card.dataset.empId) {
          var stack = getRoleNavStack();
          var current = getEmployeeSession();
          stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
          localStorage.setItem('roleNavStack', JSON.stringify(stack));
          localStorage.removeItem('roleAuth');
          localStorage.removeItem('roleName');
          localStorage.removeItem('roleLocation');
          localStorage.setItem('employeeId', card.dataset.empId);
          localStorage.setItem('employeeName', card.dataset.empName);
          window.location.reload();
        }
      });
    }
  }

  /* ========== Load Collection Data ========== */
  window._loadCollectionTab = async function () {
    try {
      await initDB();
      var wb = await getWorkbookWithFallback();
      if (!wb || !wb.data) {
        document.getElementById('collectionContent').innerHTML = noDataHtml();
        return;
      }

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });
      var sheetMap = categorizeSheets(workbook.SheetNames);
      _collState.sheetMap = sheetMap;

      // Parse all sheets
      _collState.rows = parseRows(workbook, sheetMap.overall);
      _collState.fyRows = parseRows(workbook, sheetMap.fy);
      _collState.onDateRows = parseRows(workbook, sheetMap.overallOnDate);
      _collState.tomRows = parseRows(workbook, sheetMap.overallTom);
      _collState.fyOnDateRows = parseRows(workbook, sheetMap.fyOnDate);
      _collState.fyTomRows = parseRows(workbook, sheetMap.fyTom);

      // Detect product boundaries
      if (_collState.rows) _collState.bounds = detectProductBounds(_collState.rows);
      if (_collState.fyRows) _collState.fyBounds = detectProductBounds(_collState.fyRows);

      renderCollection();

      // Show "Last updated" timestamp
      if (wb.uploadedAt) {
        var d = new Date(wb.uploadedAt);
        var day = String(d.getDate()).padStart(2, '0');
        var mon = String(d.getMonth() + 1).padStart(2, '0');
        var hr = String(d.getHours()).padStart(2, '0');
        var min = String(d.getMinutes()).padStart(2, '0');
        var el = document.getElementById('lastUpdated');
        if (el) el.textContent = 'Updated ' + day + '-' + mon + '-' + d.getFullYear() + ' ' + hr + ':' + min;
      }
    } catch (err) {
      console.error('Collection load failed:', err);
      document.getElementById('collectionContent').innerHTML = noDataHtml('Failed to load collection data.');
    }
  };
})();
