(function () {
  'use strict';

  // Inject CSS if not already present
  if (!document.getElementById('analytical-styles')) {
    var style = document.createElement('style');
    style.id = 'analytical-styles';
    style.textContent = [
      '.anal-container { padding: 12px 10px 80px; }',
      '.anal-page-title { font-size: 20px; font-weight: 700; color: #E8ECF4; margin-bottom: 4px; }',
      '.anal-page-subtitle { font-size: 12px; color: #6B7A99; margin-bottom: 16px; }',
      '.anal-btn-row { display: flex; gap: 10px; margin-bottom: 20px; }',
      '.anal-btn {',
      '  padding: 10px 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08);',
      '  background: #131825; color: #6B7A99; font-size: 14px; font-weight: 600;',
      '  cursor: pointer; transition: all 0.2s; flex: 1; text-align: center;',
      '}',
      '.anal-btn.active {',
      '  background: rgba(79,140,255,0.12); color: #4F8CFF; border-color: rgba(79,140,255,0.3);',
      '}',
      '.anal-btn:active { transform: scale(0.97); }',
      '.anal-content { }',
      '.anal-section { margin-bottom: 24px; }',
      '.anal-section-header {',
      '  display: flex; align-items: center; gap: 10px;',
      '  margin-bottom: 14px; padding: 0 2px;',
      '}',
      '.anal-section-icon {',
      '  width: 32px; height: 32px; border-radius: 10px;',
      '  display: flex; align-items: center; justify-content: center;',
      '  font-size: 15px; flex-shrink: 0;',
      '}',
      '.anal-section-icon.top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-section-icon.bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-section-title { font-size: 15px; font-weight: 700; color: #E8ECF4; }',
      '.anal-section-count {',
      '  font-size: 11px; font-weight: 500; color: #6B7A99;',
      '  background: rgba(107,122,153,0.12); padding: 2px 8px; border-radius: 10px; margin-left: auto;',
      '}',
      '.anal-empty { text-align: center; color: #6B7A99; padding: 40px 0; font-size: 13px; }',
      '.anal-list { display: flex; flex-direction: column; gap: 10px; }',
      '.anal-card {',
      '  background: #131825;',
      '  border: 1px solid rgba(255,255,255,0.06);',
      '  border-radius: 14px;',
      '  padding: 14px;',
      '  transition: border-color 0.2s, transform 0.2s;',
      '}',
      '.anal-card:active {',
      '  border-color: rgba(255,255,255,0.12);',
      '  transform: scale(0.99);',
      '}',
      '.anal-card-header {',
      '  display: flex; align-items: center; gap: 8px; margin-bottom: 10px;',
      '}',
      '.anal-rank {',
      '  font-size: 11px; font-weight: 700; padding: 3px 7px;',
      '  border-radius: 8px; min-width: 28px; text-align: center; flex-shrink: 0;',
      '}',
      '.anal-branch-avatar {',
      '  width: 30px; height: 30px; border-radius: 50%;',
      '  display: flex; align-items: center; justify-content: center;',
      '  font-weight: 700; font-size: 13px; flex-shrink: 0;',
      '}',
      '.anal-avatar-top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-avatar-bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-branch-name {',
      '  flex: 1; font-size: 13px; font-weight: 600; color: #E8ECF4;',
      '  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;',
      '}',
      '.anal-pct {',
      '  font-size: 16px; font-weight: 800; flex-shrink: 0;',
      '  font-family: "DM Sans", sans-serif;',
      '}',
      '.anal-card-stats {',
      '  display: flex; justify-content: space-between; margin-bottom: 8px;',
      '}',
      '.anal-stat {',
      '  display: flex; flex-direction: column; align-items: center; gap: 1px; flex: 1;',
      '}',
      '.anal-stat-lbl {',
      '  font-size: 9px; text-transform: uppercase; color: #6B7A99; letter-spacing: 0.5px;',
      '}',
      '.anal-stat-val {',
      '  font-size: 13px; font-weight: 700; color: #E8ECF4;',
      '}',
      '.anal-divider {',
      '  height: 1px; background: rgba(255,255,255,0.04); margin: 16px 0;',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  var HIERARCHY = {"ANDRA PRADESH":{"KADAPA":["BUDWAL","DHARMAVARAM","KADAPA","KADIRI"]},"DHARWAD":{"BADAMI":["BADAMI","GAJENDRAGAD","NARAGUNDA","RAMDURGA"],"BALLARI":["BALLARI","KUDATHINI","SANDURU","SIRUGUPPA"],"BELAGAVI":["BAILHONGAL","BELAGAVI","GOKAK","KITTUR","YARAGATTI"],"CHIKKODI":["ATHANI","CHIKKODI","MUDALAGI","NIPPANI"],"DAVANAGERE":["DAVANAGERE","HARIHARA","HONNALI","SANTHEBENNURU"],"DHARWAD":["DHARWAD","HUBLI","HUBLI-2","KALGHATGI"],"GADAG":["GADAG","LAXMESHWAR","MUNDARAGI"],"KUDLIGI":["HARAPANAHALLI","KHANAHOSAHALLI","KOTTURU","KUDLIGI"],"VIJAYANAGARA":["HAGARIBOMMANAHALLI","HOSPET","HUVENAHADAGALLI"]},"KALBURGI":{"BIDAR":["AURAD","BHALKI","BIDAR","BIDAR-2"],"HUMNABAD":["BASAVAKALYAN","HULSOOR","HUMNABAD","KAMALAPURA"],"INDI":["ALMEL","CHADCHAN","INDI"],"KALBURGI":["AFZALPUR","ALAND","JEVARGI","KALABURAGI","KALBURGI-2"],"KUSHTAGI":["BAGALKOT","GANGAVATHI","HUNGUND","KOPPAL","KUSHTAGI"],"LINGSUGUR":["DEVADURGA","LINGSUGUR","MANVI","RAICHUR","SINDHNUR","SIRWAR"],"SEDAM":["CHINCHOLI","KALAGI","SEDAM","SHAHAPUR","YADGIR"],"VIJAYAPURA":["BILAGI","JAMAKHANDI","LOKAPUR","MUDDEBIHAL","SINDAGI","TALIKOTI","TIKOTA","VIJAYAPUR"]},"TELANGANA":{"MAHABOOBNAGAR":["GADWAL","MAHABUB NAGAR","MARIKAL","TANDUR"],"SANGAREDDY":["KODANGAL","NARAYANKHED","SANGAREDDY","ZAHEERABAD"]},"TUMKUR":{"BENGALORE -RURAL":["DABUSPET","DODDABALLAPURA","GOWRIBIDANUR"],"BENGALORE -URBAN":["CHANDAPURA","HEBBAL","J P NAGAR","KENGERI"],"CHIKKABALLAPUR":["BAGEPALLI","CHIKBALLAPURA","CHINTAMANI","DEVANAHALLI","SRINIVASPURA"],"CHIKKAMAGALURU":["CHIKKAMAGALURU","MUDIGERE","NR PURA"],"CHITRADURGA":["CHALLAKERE","CHITRADURGA","HIRIYUR","JAGALORE"],"HOLALKERE":["CHANNAGIRI","HOLAKERE","HOSADURGA"],"KADUR":["AJJAMPURA","KADUR","PANCHANHALLI","TARIKERE"],"KOLAR":["BANGARPET","BETHAMANGALA","KOLAR","MALUR"],"TIPTUR":["CHIKKANAYAKANAHALLI","GUBBI","HULIYAR","TIPTUR","TUREVEKERE"],"TUMKUR":["KORATAGERE","KUNIGAL","MADHUGIRI","SIRA","TUMKUR"]}};

  var COL = {
    name: 1,
    regDemand: 2, regCollection: 3, regFtod: 4, regPct: 5,
    dpd1Demand: 6, dpd1Collection: 7, dpd1Balance: 8, dpd1Pct: 9,
    dpd2Demand: 10, dpd2Collection: 11, dpd2Balance: 12, dpd2Pct: 13,
    pnpaDemand: 14, pnpaCollection: 15, pnpaBalance: 16, pnpaPct: 17,
    npaDemand: 18, npaActAcct: 19, npaActAmt: 20, npaClsAcct: 21, npaClsAmt: 22,
    rank: 23, performance: 24
  };

  function numVal(v) { return typeof v === 'number' ? v : (parseFloat(String(v).replace(/,/g, '')) || 0); }
  function fmtNum(n) { return n.toLocaleString('en-IN'); }
  function esc(s) { var d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
  function fuzzyMatch(a, b) { return a.replace(/[\s\-]/g, '').indexOf(b.replace(/[\s\-]/g, '')) >= 0 || b.replace(/[\s\-]/g, '').indexOf(a.replace(/[\s\-]/g, '')) >= 0; }

  // Get all branches under a region from HIERARCHY
  function getRegionBranches(regionName) {
    var rUpper = regionName.toUpperCase().trim();
    for (var rk in HIERARCHY) {
      if (fuzzyMatch(rk, rUpper)) {
        var branches = [];
        var districts = HIERARCHY[rk];
        for (var dk in districts) {
          for (var i = 0; i < districts[dk].length; i++) {
            branches.push(districts[dk][i].toUpperCase());
          }
        }
        return branches;
      }
    }
    return [];
  }

  // Get all branches under a district from HIERARCHY
  function getDistrictBranches(districtName) {
    var dUpper = districtName.toUpperCase().trim();
    for (var rk in HIERARCHY) {
      var districts = HIERARCHY[rk];
      for (var dk in districts) {
        if (fuzzyMatch(dk, dUpper)) {
          var branches = [];
          for (var i = 0; i < districts[dk].length; i++) {
            branches.push(districts[dk][i].toUpperCase());
          }
          return branches;
        }
      }
    }
    return [];
  }

  // Generic section row extractor — mirrors collection.js getAllSectionRows
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

  function filterByNames(sectionRows, allowedNames) {
    var filtered = [];
    for (var i = 0; i < sectionRows.length; i++) {
      var bName = sectionRows[i].name.toUpperCase().trim();
      for (var j = 0; j < allowedNames.length; j++) {
        if (fuzzyMatch(bName, allowedNames[j])) {
          filtered.push(sectionRows[i]);
          break;
        }
      }
    }
    return filtered;
  }

  function computeData(sectionRows) {
    var data = [];
    for (var i = 0; i < sectionRows.length; i++) {
      var b = sectionRows[i];
      var dem = numVal(b.row[COL.regDemand]);
      var col = numVal(b.row[COL.regCollection]);
      var bal = dem - col;
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      data.push({ name: b.name, demand: dem, collection: col, balance: bal, pct: pctRaw });
    }
    data.sort(function (a, b) { return a.pct - b.pct; });
    return data;
  }

  function buildCardHtml(item, rankNum, type, delay) {
    var pctColor = item.pct >= 95 ? '#34D399' : item.pct >= 80 ? '#FBBF24' : '#F87171';
    var barW = Math.min(item.pct, 100);
    var initial = item.name.charAt(0).toUpperCase();
    var avatarClass = type === 'top' ? 'anal-avatar-top' : 'anal-avatar-bottom';

    var html = '<div class="anal-card emp-fade" style="animation-delay:' + delay + 's">';
    html += '<div class="anal-card-header">';
    html += '<div class="anal-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>';
    html += '<div class="anal-branch-avatar ' + avatarClass + '">' + initial + '</div>';
    html += '<div class="anal-branch-name">' + esc(item.name) + '</div>';
    html += '<div class="anal-pct" style="color:' + pctColor + '">' + item.pct.toFixed(2) + '%</div>';
    html += '</div>';
    html += '<div class="anal-card-stats">';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Demand</span><span class="anal-stat-val">' + fmtNum(item.demand) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Collection</span><span class="anal-stat-val" style="color:#4F8CFF">' + fmtNum(item.collection) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Balance</span><span class="anal-stat-val" style="color:#FB923C">' + fmtNum(item.balance) + '</span></div>';
    html += '</div>';
    html += '<div class="pf-progress-track"><div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div></div>';
    html += '</div>';
    return html;
  }

  function buildSectionHtml(title, iconSvg, type, items) {
    var html = '<div class="anal-section">';
    html += '<div class="anal-section-header">';
    html += '<div class="anal-section-icon ' + type + '">' + iconSvg + '</div>';
    html += '<div class="anal-section-title">' + title + '</div>';
    html += '<div class="anal-section-count">' + items.length + '</div>';
    html += '</div>';

    if (items.length === 0) {
      html += '<div class="anal-empty">No data available</div>';
    } else {
      html += '<div class="anal-list">';
      for (var j = 0; j < items.length; j++) {
        html += buildCardHtml(items[j], j + 1, type, j * 0.04);
      }
      html += '</div>';
    }
    html += '</div>';
    return html;
  }

  // Stored state for re-rendering
  var _analState = { rows: null, role: null, location: null };

  function renderAnalytical() {
    var rows = _analState.rows;
    var role = _analState.role;
    var location = _analState.location;
    var container = document.getElementById('analyticalContent');
    if (!container || !rows) return;

    var upArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>';
    var downArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></svg>';

    var html = '<div class="anal-container">';
    html += '<div class="anal-page-title">Analytical Tool</div>';
    html += '<div class="anal-page-subtitle">Performance analysis based on regular collection %</div>';

    // Buttons row
    html += '<div class="anal-btn-row">';
    html += '<button class="anal-btn active" data-anal-view="branches">';
    html += '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:-2px;margin-right:4px;"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>';
    html += 'Branches</button>';
    html += '</div>';

    // Content
    html += '<div class="anal-content">';

    if (role === 'CEO') {
      var branchData = computeData(getAllSectionRows(rows, 'branch'));
      var top10 = branchData.slice(-10).reverse();
      var bottom10 = branchData.slice(0, 10);

      html += buildSectionHtml('Top 10 Branches', upArrow, 'top', top10);
      html += '<div class="anal-divider"></div>';
      html += buildSectionHtml('Bottom 10 Branches', downArrow, 'bottom', bottom10);
    }

    if (role === 'RM') {
      var allBranches = getAllSectionRows(rows, 'branch');
      var regionBranchNames = getRegionBranches(location);
      var filtered = filterByNames(allBranches, regionBranchNames);
      var regionData = computeData(filtered);
      var top10rm = regionData.slice(-10).reverse();

      html += buildSectionHtml('Top 10 Branches \u2014 ' + esc(location), upArrow, 'top', top10rm);
    }

    if (role === 'DM') {
      var allBranches2 = getAllSectionRows(rows, 'branch');
      var districtBranchNames = getDistrictBranches(location);
      var filtered2 = filterByNames(allBranches2, districtBranchNames);
      var districtData = computeData(filtered2);
      var top5 = districtData.slice(-5).reverse();
      var bottom5 = districtData.slice(0, 5);

      html += buildSectionHtml('Top 5 Branches \u2014 ' + esc(location), upArrow, 'top', top5);
      html += '<div class="anal-divider"></div>';
      html += buildSectionHtml('Bottom 5 Branches \u2014 ' + esc(location), downArrow, 'bottom', bottom5);
    }

    html += '</div>';
    html += '</div>';
    container.innerHTML = html;
  }

  // Load function exposed globally
  window._loadAnalyticalTab = async function () {
    try {
      var session = typeof getEmployeeSession === 'function' ? getEmployeeSession() : null;
      if (!session) return;

      var role = session.role;

      // Only show for CEO, RM, DM
      if (role !== 'CEO' && role !== 'RM' && role !== 'DM') {
        var navItem = document.getElementById('analyticalNavItem');
        if (navItem) navItem.style.display = 'none';
        return;
      }

      // Show sidebar nav item
      var navItem2 = document.getElementById('analyticalNavItem');
      if (navItem2) navItem2.style.display = '';

      if (typeof getWorkbookWithFallback !== 'function') return;
      var wb = await getWorkbookWithFallback();
      if (!wb || !wb.data) return;

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });
      if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return;

      var overallSheet = null;
      for (var i = 0; i < workbook.SheetNames.length; i++) {
        var sn = workbook.SheetNames[i];
        if (/overall/i.test(sn) && !/on.?date/i.test(sn) && !/^tom_/i.test(sn) && !/^fy/i.test(sn)) {
          overallSheet = sn;
          break;
        }
      }
      if (!overallSheet) overallSheet = workbook.SheetNames[0];

      var sheet = workbook.Sheets[overallSheet];
      if (!sheet) return;

      var rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      _analState.rows = rows;
      _analState.role = role;
      _analState.location = session.location;
      renderAnalytical();
    } catch (err) {
      console.error('Analytical load failed:', err);
    }
  };
})();
