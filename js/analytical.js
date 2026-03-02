(function () {
  'use strict';

  if (!document.getElementById('analytical-styles')) {
    var style = document.createElement('style');
    style.id = 'analytical-styles';
    style.textContent = [
      '.anal-container { padding: 12px 10px 80px; }',
      '.anal-page-title { font-size: 20px; font-weight: 700; color: #E8ECF4; margin-bottom: 4px; }',
      '.anal-page-subtitle { font-size: 12px; color: #6B7A99; margin-bottom: 16px; }',
      '.anal-btn-row { display: flex; gap: 10px; margin-bottom: 20px; }',
      '.anal-btn {',
      '  padding: 10px 16px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.08);',
      '  background: #131825; color: #6B7A99; font-size: 13px; font-weight: 600;',
      '  cursor: pointer; transition: all 0.2s; flex: 1; text-align: center;',
      '  display: flex; align-items: center; justify-content: center; gap: 6px;',
      '}',
      '.anal-btn.active { background: rgba(79,140,255,0.12); color: #4F8CFF; border-color: rgba(79,140,255,0.3); }',
      '.anal-btn:active { transform: scale(0.97); }',
      '.anal-section { margin-bottom: 24px; }',
      '.anal-section-header { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; padding: 0 2px; }',
      '.anal-section-icon { width: 32px; height: 32px; border-radius: 10px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }',
      '.anal-section-icon.top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-section-icon.bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-section-title { font-size: 15px; font-weight: 700; color: #E8ECF4; }',
      '.anal-section-count { font-size: 11px; font-weight: 500; color: #6B7A99; background: rgba(107,122,153,0.12); padding: 2px 8px; border-radius: 10px; margin-left: auto; }',
      '.anal-empty { text-align: center; color: #6B7A99; padding: 40px 0; font-size: 13px; }',
      '.anal-list { display: flex; flex-direction: column; gap: 10px; }',
      '.anal-card { background: #131825; border: 1px solid rgba(255,255,255,0.06); border-radius: 14px; padding: 14px; transition: border-color 0.2s, transform 0.2s; }',
      '.anal-card:active { border-color: rgba(255,255,255,0.12); transform: scale(0.99); }',
      '.anal-card-header { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }',
      '.anal-rank { font-size: 11px; font-weight: 700; padding: 3px 7px; border-radius: 8px; min-width: 28px; text-align: center; flex-shrink: 0; }',
      '.anal-branch-avatar { width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px; flex-shrink: 0; }',
      '.anal-avatar-top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-avatar-bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-branch-name { flex: 1; font-size: 13px; font-weight: 600; color: #E8ECF4; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }',
      '.anal-pct { font-size: 16px; font-weight: 800; flex-shrink: 0; font-family: "DM Sans", sans-serif; }',
      '.anal-card-stats { display: flex; justify-content: space-between; margin-bottom: 8px; }',
      '.anal-stat { display: flex; flex-direction: column; align-items: center; gap: 1px; flex: 1; }',
      '.anal-stat-lbl { font-size: 9px; text-transform: uppercase; color: #6B7A99; letter-spacing: 0.5px; }',
      '.anal-stat-val { font-size: 13px; font-weight: 700; color: #E8ECF4; }',
      '.anal-divider { height: 1px; background: rgba(255,255,255,0.04); margin: 16px 0; }',
      // FO card styles
      '.fo-card { background: #131825; border: 1px solid rgba(255,255,255,0.06); border-radius: 14px; padding: 14px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s; }',
      '.fo-card:active { transform: scale(0.99); }',
      '.fo-header { display: flex; align-items: center; gap: 10px; }',
      '.fo-rank { font-size: 11px; font-weight: 700; padding: 3px 7px; border-radius: 8px; min-width: 28px; text-align: center; flex-shrink: 0; }',
      '.fo-avatar { width: 30px; height: 30px; border-radius: 50%; background: rgba(248,113,113,0.12); color: #F87171; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px; flex-shrink: 0; }',
      '.fo-info { flex: 1; min-width: 0; }',
      '.fo-name { font-size: 13px; font-weight: 600; color: #E8ECF4; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }',
      '.fo-empid { font-size: 10px; color: #6B7A99; }',
      '.fo-pct { font-size: 16px; font-weight: 800; flex-shrink: 0; font-family: "DM Sans", sans-serif; }',
      '.fo-stats { display: flex; justify-content: space-between; margin-top: 10px; }',
      '.fo-actions { display: flex; gap: 8px; margin-top: 12px; }',
      '.fo-call-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: rgba(52,211,153,0.1); color: #34D399; font-size: 13px; font-weight: 600; text-decoration: none; border: 1px solid rgba(52,211,153,0.2); transition: all 0.2s; }',
      '.fo-call-btn:active { transform: scale(0.97); background: rgba(52,211,153,0.2); }',
      '.fo-no-phone { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: rgba(107,122,153,0.08); color: #6B7A99; font-size: 12px; border: 1px solid rgba(107,122,153,0.12); }',
      '.fo-view-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: rgba(79,140,255,0.1); color: #4F8CFF; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid rgba(79,140,255,0.2); transition: all 0.2s; }',
      '.fo-view-btn:active { transform: scale(0.97); background: rgba(79,140,255,0.2); }',
      '.fo-branch-tag { font-size: 10px; color: #4F8CFF; background: rgba(79,140,255,0.1); padding: 2px 6px; border-radius: 6px; margin-left: 4px; }'
    ].join('\n');
    document.head.appendChild(style);
  }

  var HIERARCHY = {"ANDRA PRADESH":{"KADAPA":["BUDWAL","DHARMAVARAM","KADAPA","KADIRI"]},"DHARWAD":{"BADAMI":["BADAMI","GAJENDRAGAD","NARAGUNDA","RAMDURGA"],"BALLARI":["BALLARI","KUDATHINI","SANDURU","SIRUGUPPA"],"BELAGAVI":["BAILHONGAL","BELAGAVI","GOKAK","KITTUR","YARAGATTI"],"CHIKKODI":["ATHANI","CHIKKODI","MUDALAGI","NIPPANI"],"DAVANAGERE":["DAVANAGERE","HARIHARA","HONNALI","SANTHEBENNURU"],"DHARWAD":["DHARWAD","HUBLI","HUBLI-2","KALGHATGI"],"GADAG":["GADAG","LAXMESHWAR","MUNDARAGI"],"KUDLIGI":["HARAPANAHALLI","KHANAHOSAHALLI","KOTTURU","KUDLIGI"],"VIJAYANAGARA":["HAGARIBOMMANAHALLI","HOSPET","HUVENAHADAGALLI"]},"KALBURGI":{"BIDAR":["AURAD","BHALKI","BIDAR","BIDAR-2"],"HUMNABAD":["BASAVAKALYAN","HULSOOR","HUMNABAD","KAMALAPURA"],"INDI":["ALMEL","CHADCHAN","INDI"],"KALBURGI":["AFZALPUR","ALAND","JEVARGI","KALABURAGI","KALBURGI-2"],"KUSHTAGI":["BAGALKOT","GANGAVATHI","HUNGUND","KOPPAL","KUSHTAGI"],"LINGSUGUR":["DEVADURGA","LINGSUGUR","MANVI","RAICHUR","SINDHNUR","SIRWAR"],"SEDAM":["CHINCHOLI","KALAGI","SEDAM","SHAHAPUR","YADGIR"],"VIJAYAPURA":["BILAGI","JAMAKHANDI","LOKAPUR","MUDDEBIHAL","SINDAGI","TALIKOTI","TIKOTA","VIJAYAPUR"]},"TELANGANA":{"MAHABOOBNAGAR":["GADWAL","MAHABUB NAGAR","MARIKAL","TANDUR"],"SANGAREDDY":["KODANGAL","NARAYANKHED","SANGAREDDY","ZAHEERABAD"]},"TUMKUR":{"BENGALORE -RURAL":["DABUSPET","DODDABALLAPURA","GOWRIBIDANUR"],"BENGALORE -URBAN":["CHANDAPURA","HEBBAL","J P NAGAR","KENGERI"],"CHIKKABALLAPUR":["BAGEPALLI","CHIKBALLAPURA","CHINTAMANI","DEVANAHALLI","SRINIVASPURA"],"CHIKKAMAGALURU":["CHIKKAMAGALURU","MUDIGERE","NR PURA"],"CHITRADURGA":["CHALLAKERE","CHITRADURGA","HIRIYUR","JAGALORE"],"HOLALKERE":["CHANNAGIRI","HOLAKERE","HOSADURGA"],"KADUR":["AJJAMPURA","KADUR","PANCHANHALLI","TARIKERE"],"KOLAR":["BANGARPET","BETHAMANGALA","KOLAR","MALUR"],"TIPTUR":["CHIKKANAYAKANAHALLI","GUBBI","HULIYAR","TIPTUR","TUREVEKERE"],"TUMKUR":["KORATAGERE","KUNIGAL","MADHUGIRI","SIRA","TUMKUR"]}};

  var SUPABASE_URL = 'https://zovnmmdfthpbubrorsgh.supabase.co';
  var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inpvdm5tbWRmdGhwYnVicm9yc2doIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE1NzE3ODgsImV4cCI6MjA3NzE0Nzc4OH0.92BH2sjUOgkw6iSRj1_4gt0p3eThg3QT4VK-Q4EdmBE';

  var COL = {
    name: 1, regDemand: 2, regCollection: 3, regFtod: 4, regPct: 5,
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

  function getRegionBranches(regionName) {
    var rUpper = regionName.toUpperCase().trim();
    for (var rk in HIERARCHY) {
      if (fuzzyMatch(rk, rUpper)) {
        var branches = [];
        for (var dk in HIERARCHY[rk]) { for (var i = 0; i < HIERARCHY[rk][dk].length; i++) branches.push(HIERARCHY[rk][dk][i].toUpperCase()); }
        return branches;
      }
    }
    return [];
  }

  function getDistrictBranches(districtName) {
    var dUpper = districtName.toUpperCase().trim();
    for (var rk in HIERARCHY) { for (var dk in HIERARCHY[rk]) { if (fuzzyMatch(dk, dUpper)) return HIERARCHY[rk][dk].map(function(b) { return b.toUpperCase(); }); } }
    return [];
  }

  function getAllSectionRows(rows, sectionType) {
    var results = [], currentSection = null, foundTarget = false;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r]; if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase(), c1 = String(row[1] || '').trim().toUpperCase(), hdr = c0 + ' ' + c1;
      if (hdr.includes('REGION') && (hdr.includes('WISE') || hdr.includes('NAME'))) { if (currentSection === sectionType && foundTarget) break; currentSection = 'region'; continue; }
      if (hdr.includes('DISTRICT') && (hdr.includes('WISE') || hdr.includes('NAME'))) { if (currentSection === sectionType && foundTarget) break; currentSection = 'district'; continue; }
      if (hdr.includes('BRANCH') && !hdr.includes('OFFICER') && (hdr.includes('WISE') || hdr.includes('NAME'))) { if (currentSection === sectionType && foundTarget) break; currentSection = 'branch'; continue; }
      if (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL')) { if (currentSection === sectionType && foundTarget) break; currentSection = null; continue; }
      if (currentSection !== sectionType) continue;
      if (c0 === 'SL' || c0 === 'SL.' || c1 === 'SL' || c1 === 'SL.') continue;
      if (/^TOTAL/i.test(c0)) continue;
      var name = String(row[1] || '').trim().toUpperCase();
      if (!name || /^\d+$/.test(name)) name = String(row[0] || '').trim().toUpperCase();
      if (!name || /^\d+$/.test(name)) continue;
      if (/^(SL|NAME|OFFICER|DEMAND|COLLECTION|FTOD|BALANCE|ACCOUNT|AMOUNT)/i.test(name)) continue;
      results.push({ name: name, row: row }); foundTarget = true;
    }
    return results;
  }

  // Parse ALL field officers from all branches
  function getAllOfficers(rows) {
    var officers = [];
    var inOfficerSection = false;
    var currentBranch = '';
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r]; if (!row || !row.length) continue;
      var c0 = String(row[0] || '').trim().toUpperCase();
      var c1 = String(row[1] || '').trim().toUpperCase();
      var hdr = c0 + ' ' + c1;
      if (hdr.includes('BRANCH') && hdr.includes('OFFICER') && hdr.includes('NAME')) { inOfficerSection = true; continue; }
      if (!inOfficerSection) continue;
      if (c0 === 'EMP ID' || c1 === 'BRANCH / OFFICER NAME' || c1 === 'DEMAND' || c1 === 'COLLECTION') continue;
      var hasEmpId = c0.length > 0 && /^[A-Z]{2}\d+/.test(c0);
      if (!hasEmpId && c1.length > 0 && row.length > 3) {
        currentBranch = c1; continue;
      }
      if (hasEmpId && currentBranch) {
        officers.push({ name: String(row[1] || '').trim(), empId: String(row[0] || '').trim(), branch: currentBranch, row: row });
      }
    }
    return officers;
  }

  function filterByNames(sectionRows, allowedNames) {
    var filtered = [];
    for (var i = 0; i < sectionRows.length; i++) {
      var bName = sectionRows[i].name.toUpperCase().trim();
      for (var j = 0; j < allowedNames.length; j++) { if (fuzzyMatch(bName, allowedNames[j])) { filtered.push(sectionRows[i]); break; } }
    }
    return filtered;
  }

  function computeData(sectionRows) {
    var data = [];
    for (var i = 0; i < sectionRows.length; i++) {
      var b = sectionRows[i], dem = numVal(b.row[COL.regDemand]), col = numVal(b.row[COL.regCollection]);
      data.push({ name: b.name, demand: dem, collection: col, balance: dem - col, pct: dem > 0 ? (col / dem) * 100 : 0 });
    }
    data.sort(function (a, b) { return a.pct - b.pct; });
    return data;
  }

  function computeFOData(officers, bucket) {
    var data = [];
    for (var i = 0; i < officers.length; i++) {
      var o = officers[i];
      var dem, col, pct;
      if (bucket === 'dpd1') { dem = numVal(o.row[COL.dpd1Demand]); col = numVal(o.row[COL.dpd1Collection]); }
      else if (bucket === 'dpd2') { dem = numVal(o.row[COL.dpd2Demand]); col = numVal(o.row[COL.dpd2Collection]); }
      else if (bucket === 'pnpa') { dem = numVal(o.row[COL.pnpaDemand]); col = numVal(o.row[COL.pnpaCollection]); }
      else if (bucket === 'npa') { dem = numVal(o.row[COL.npaDemand]); col = numVal(o.row[COL.npaActAmt]); }
      else { dem = numVal(o.row[COL.regDemand]); col = numVal(o.row[COL.regCollection]); }
      // Skip officers with 0 demand and 0 collection
      if (dem === 0 && col === 0) continue;
      pct = dem > 0 ? (col / dem) * 100 : 0;
      data.push({
        name: o.name, empId: o.empId, branch: o.branch,
        demand: dem, collection: col, balance: dem - col, pct: pct,
        dpd1Dem: numVal(o.row[COL.dpd1Demand]), dpd1Col: numVal(o.row[COL.dpd1Collection]), dpd1Pct: numVal(o.row[COL.dpd1Demand]) > 0 ? (numVal(o.row[COL.dpd1Collection]) / numVal(o.row[COL.dpd1Demand])) * 100 : 0,
        dpd2Dem: numVal(o.row[COL.dpd2Demand]), dpd2Col: numVal(o.row[COL.dpd2Collection]), dpd2Pct: numVal(o.row[COL.dpd2Demand]) > 0 ? (numVal(o.row[COL.dpd2Collection]) / numVal(o.row[COL.dpd2Demand])) * 100 : 0,
        pnpaDem: numVal(o.row[COL.pnpaDemand]), pnpaCol: numVal(o.row[COL.pnpaCollection]), pnpaPct: numVal(o.row[COL.pnpaDemand]) > 0 ? (numVal(o.row[COL.pnpaCollection]) / numVal(o.row[COL.pnpaDemand])) * 100 : 0,
        npaDem: numVal(o.row[COL.npaDemand]), npaCol: numVal(o.row[COL.npaActAmt]), npaPct: numVal(o.row[COL.npaDemand]) > 0 ? (numVal(o.row[COL.npaActAmt]) / numVal(o.row[COL.npaDemand])) * 100 : 0,
        npaActAcct: numVal(o.row[COL.npaActAcct]), npaActAmt: numVal(o.row[COL.npaActAmt])
      });
    }
    data.sort(function (a, b) { return a.pct - b.pct; });
    return data;
  }

  // Fetch phone numbers from Supabase
  var _phoneCache = {};
  async function fetchPhones(empIds) {
    var toFetch = empIds.filter(function(id) { return !_phoneCache[id.toUpperCase()]; });
    if (!toFetch.length) return;
    try {
      // Fetch all employees and match locally for case-insensitive lookup
      var batchSize = 50;
      for (var b = 0; b < toFetch.length; b += batchSize) {
        var batch = toFetch.slice(b, b + batchSize);
        var orFilter = batch.map(function(id) { return 'emp_id.ilike.' + id; }).join(',');
        var url = SUPABASE_URL + '/rest/v1/employees?select=emp_id,mobile&or=(' + orFilter + ')';
        var resp = await fetch(url, { headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY } });
        var data = await resp.json();
        if (Array.isArray(data)) { for (var i = 0; i < data.length; i++) { if (data[i].emp_id && data[i].mobile) _phoneCache[data[i].emp_id.toUpperCase()] = data[i].mobile; } }
      }
    } catch (e) { console.error('Phone fetch failed:', e); }
  }

  // State
  var _analState = { rows: null, role: null, location: null, activeView: null, expandedFO: null, foBucket: 'regular', foLoading: false };

  function buildCardHtml(item, rankNum, type, delay) {
    var pctColor = item.pct >= 95 ? '#34D399' : item.pct >= 80 ? '#FBBF24' : '#F87171';
    var barW = Math.min(item.pct, 100);
    var avatarClass = type === 'top' ? 'anal-avatar-top' : 'anal-avatar-bottom';
    return '<div class="anal-card emp-fade" style="animation-delay:' + delay + 's">' +
      '<div class="anal-card-header">' +
      '<div class="anal-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>' +
      '<div class="anal-branch-avatar ' + avatarClass + '">' + item.name.charAt(0) + '</div>' +
      '<div class="anal-branch-name">' + esc(item.name) + '</div>' +
      '<div class="anal-pct" style="color:' + pctColor + '">' + item.pct.toFixed(2) + '%</div>' +
      '</div>' +
      '<div class="anal-card-stats">' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Demand</span><span class="anal-stat-val">' + fmtNum(item.demand) + '</span></div>' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Collection</span><span class="anal-stat-val" style="color:#4F8CFF">' + fmtNum(item.collection) + '</span></div>' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Balance</span><span class="anal-stat-val" style="color:#FB923C">' + fmtNum(item.balance) + '</span></div>' +
      '</div>' +
      '<div class="pf-progress-track"><div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div></div>' +
      '</div>';
  }

  function buildSectionHtml(title, iconSvg, type, items) {
    var html = '<div class="anal-section"><div class="anal-section-header">' +
      '<div class="anal-section-icon ' + type + '">' + iconSvg + '</div>' +
      '<div class="anal-section-title">' + title + '</div>' +
      '<div class="anal-section-count">' + items.length + '</div></div>';
    if (!items.length) { html += '<div class="anal-empty">No data available</div>'; }
    else { html += '<div class="anal-list">'; for (var j = 0; j < items.length; j++) html += buildCardHtml(items[j], j + 1, type, j * 0.04); html += '</div>'; }
    return html + '</div>';
  }

  function buildFOCardHtml(fo, rankNum, delay, expanded) {
    var pctColor = fo.pct >= 95 ? '#34D399' : fo.pct >= 80 ? '#FBBF24' : '#F87171';
    var barW = Math.min(fo.pct, 100);
    var phone = _phoneCache[fo.empId.toUpperCase()] || '';

    var html = '<div class="fo-card emp-fade" data-fo-empid="' + esc(fo.empId) + '" data-fo-name="' + esc(fo.name) + '" style="animation-delay:' + delay + 's">';
    html += '<div class="fo-header">';
    html += '<div class="fo-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>';
    html += '<div class="fo-avatar">' + fo.name.charAt(0).toUpperCase() + '</div>';
    html += '<div class="fo-info"><div class="fo-name">' + esc(fo.name) + '<span class="fo-branch-tag">' + esc(fo.branch) + '</span></div>';
    html += '<div class="fo-empid">' + esc(fo.empId) + '</div></div>';
    html += '<div class="fo-pct" style="color:' + pctColor + '">' + fo.pct.toFixed(2) + '%</div>';
    html += '</div>';

    html += '<div class="fo-stats">';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Demand</span><span class="anal-stat-val">' + fmtNum(fo.demand) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Collection</span><span class="anal-stat-val" style="color:#4F8CFF">' + fmtNum(fo.collection) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Balance</span><span class="anal-stat-val" style="color:#FB923C">' + fmtNum(fo.balance) + '</span></div>';
    html += '</div>';
    html += '<div class="pf-progress-track" style="margin-top:8px"><div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div></div>';

    // Action buttons — call + view dashboard
    html += '<div class="fo-actions">';
    if (phone) {
      html += '<a href="tel:' + esc(phone) + '" class="fo-call-btn" onclick="event.stopPropagation();">';
      html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>';
      html += 'Call</a>';
    } else if (_analState.foLoading) {
      html += '<div class="fo-no-phone" style="color:#4F8CFF;border-color:rgba(79,140,255,0.15)"><svg width="12" height="12" viewBox="0 0 24 24" style="animation:spin .7s linear infinite"><circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="3" stroke-dasharray="30 70"/></svg> Fetching from database...</div>';
    } else {
      html += '<div class="fo-no-phone">No phone</div>';
    }
    html += '<div class="fo-view-btn" onclick="event.stopPropagation(); this.closest(\'.fo-card\').click();">';
    html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
    html += 'View Dashboard</div>';
    html += '</div>';

    html += '</div>';
    return html;
  }

  function renderPage() {
    var container = document.getElementById('analyticalContent');
    if (!container) return;

    var upArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>';
    var downArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></svg>';

    var view = _analState.activeView;
    var role = _analState.role;
    var location = _analState.location;
    var rows = _analState.rows;

    var html = '<div class="anal-container">';

    // Buttons
    html += '<div class="anal-btn-row">';
    html += '<button class="anal-btn' + (view === 'branches' ? ' active' : '') + '" data-anal-view="branches">';
    html += '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>';
    html += 'Branches</button>';
    html += '<button class="anal-btn' + (view === 'fo' ? ' active' : '') + '" data-anal-view="fo">';
    html += '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
    html += 'Field Officer</button>';
    html += '</div>';

    // Content based on active view
    if (view === 'branches') {
      if (role === 'CEO') {
        var bd = computeData(getAllSectionRows(rows, 'branch'));
        html += buildSectionHtml('Top 10 Branches', upArrow, 'top', bd.slice(-10).reverse());
        html += '<div class="anal-divider"></div>';
        html += buildSectionHtml('Bottom 10 Branches', downArrow, 'bottom', bd.slice(0, 10));
      } else if (role === 'RM') {
        var allB = getAllSectionRows(rows, 'branch');
        var rData = computeData(filterByNames(allB, getRegionBranches(location)));
        html += buildSectionHtml('Top 10 Branches \u2014 ' + esc(location), upArrow, 'top', rData.slice(-10).reverse());
      } else if (role === 'DM') {
        var allB2 = getAllSectionRows(rows, 'branch');
        var dData = computeData(filterByNames(allB2, getDistrictBranches(location)));
        html += buildSectionHtml('Top 5 Branches \u2014 ' + esc(location), upArrow, 'top', dData.slice(-5).reverse());
        html += '<div class="anal-divider"></div>';
        html += buildSectionHtml('Bottom 5 Branches \u2014 ' + esc(location), downArrow, 'bottom', dData.slice(0, 5));
      }
    } else if (view === 'fo') {
      {
        var allOfficers = getAllOfficers(rows);
        var allowedBranches = null;
        var label = 'All';

        if (role === 'RM') {
          allowedBranches = getRegionBranches(location); label = esc(location);
        } else if (role === 'DM') {
          allowedBranches = getDistrictBranches(location); label = esc(location);
        }

        if (allowedBranches) {
          allOfficers = allOfficers.filter(function(o) {
            for (var k = 0; k < allowedBranches.length; k++) { if (fuzzyMatch(o.branch, allowedBranches[k])) return true; }
            return false;
          });
        }

        var bk = _analState.foBucket;
        var bucketKey = bk === 'dpd1' ? 'dpd1' : bk === 'dpd2' ? 'dpd2' : bk === 'pnpa' ? 'pnpa' : bk === 'npa' ? 'npa' : null;
        var foData = computeFOData(allOfficers, bucketKey);
        var lowPerformers = foData.slice(0, 20);
        var bucketLabel = bk === 'dpd1' ? '1-30' : bk === 'dpd2' ? '31-60' : bk === 'pnpa' ? 'PNPA' : bk === 'npa' ? 'NPA' : 'Regular';

        // Bucket filter pills
        html += '<div class="anal-btn-row" style="margin-bottom:16px;flex-wrap:wrap">';
        html += '<button class="anal-btn' + (bk === 'regular' ? ' active' : '') + '" data-fo-bucket="regular" style="padding:8px 10px;font-size:12px">Regular</button>';
        html += '<button class="anal-btn' + (bk === 'dpd1' ? ' active' : '') + '" data-fo-bucket="dpd1" style="padding:8px 10px;font-size:12px">1-30</button>';
        html += '<button class="anal-btn' + (bk === 'dpd2' ? ' active' : '') + '" data-fo-bucket="dpd2" style="padding:8px 10px;font-size:12px">31-60</button>';
        html += '<button class="anal-btn' + (bk === 'pnpa' ? ' active' : '') + '" data-fo-bucket="pnpa" style="padding:8px 10px;font-size:12px">PNPA</button>';
        html += '<button class="anal-btn' + (bk === 'npa' ? ' active' : '') + '" data-fo-bucket="npa" style="padding:8px 10px;font-size:12px">NPA</button>';
        html += '</div>';

        html += '<div class="anal-section"><div class="anal-section-header">';
        html += '<div class="anal-section-icon bottom">' + downArrow + '</div>';
        html += '<div class="anal-section-title">Low Performing (' + bucketLabel + ')' + (role !== 'CEO' ? ' \u2014 ' + label : '') + '</div>';
        html += '<div class="anal-section-count">' + lowPerformers.length + '</div></div>';

        if (!lowPerformers.length) {
          html += '<div class="anal-empty">No field officer data available</div>';
        } else {
          html += '<div class="anal-list">';
          for (var f = 0; f < lowPerformers.length; f++) {
            var isExpanded = _analState.expandedFO === lowPerformers[f].empId;
            html += buildFOCardHtml(lowPerformers[f], f + 1, f * 0.04, isExpanded);
          }
          html += '</div>';
        }
        html += '</div>';
      }
    }
    }

    html += '</div>';
    container.innerHTML = html;
    attachHandlers();
  }

  function attachHandlers() {
    var container = document.getElementById('analyticalContent');
    if (!container) return;

    // Button clicks
    container.querySelectorAll('[data-anal-view]').forEach(function(btn) {
      btn.onclick = async function() {
        var newView = btn.dataset.analView;
        _analState.activeView = newView;
        _analState.expandedFO = null;
        _analState.foBucket = 'regular';

        // Show FO content immediately, phones load async
        if (newView === 'fo' && _analState.rows) {
          _analState.foLoading = true;
          renderPage(); // render with "Fetching..." placeholders
          var allOfficers = getAllOfficers(_analState.rows);
          var empIds = allOfficers.map(function(o) { return o.empId; });
          await fetchPhones(empIds);
          _analState.foLoading = false;
          renderPage(); // re-render with actual phone numbers
        } else {
          renderPage();
        }
      };
    });

    // Bucket filter pills
    container.querySelectorAll('[data-fo-bucket]').forEach(function(btn) {
      btn.onclick = function() {
        _analState.foBucket = btn.dataset.foBucket;
        _analState.expandedFO = null;
        renderPage();
      };
    });

    // FO card click — navigate to FO's full dashboard
    container.querySelectorAll('[data-fo-empid]').forEach(function(card) {
      card.onclick = function() {
        var empId = card.dataset.foEmpid;
        var empName = card.dataset.foName || '';
        // Save current session to nav stack
        var stack = typeof getRoleNavStack === 'function' ? getRoleNavStack() : [];
        var current = typeof getEmployeeSession === 'function' ? getEmployeeSession() : null;
        if (current) {
          stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
          localStorage.setItem('roleNavStack', JSON.stringify(stack));
        }
        // Switch to FO session
        localStorage.removeItem('roleAuth');
        localStorage.removeItem('roleName');
        localStorage.removeItem('roleLocation');
        localStorage.setItem('employeeId', empId);
        localStorage.setItem('employeeName', empName);
        // Flag to return to analytical tab
        localStorage.setItem('returnToAnalytical', 'true');
        window.location.reload();
      };
    });
  }

  window._loadAnalyticalTab = async function () {
    try {
      var session = typeof getEmployeeSession === 'function' ? getEmployeeSession() : null;
      if (!session) return;
      var role = session.role;

      if (role !== 'CEO' && role !== 'RM' && role !== 'DM') {
        var navItem = document.getElementById('analyticalNavItem');
        if (navItem) navItem.style.display = 'none';
        return;
      }

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
        if (/overall/i.test(sn) && !/on.?date/i.test(sn) && !/^tom_/i.test(sn) && !/^fy/i.test(sn)) { overallSheet = sn; break; }
      }
      if (!overallSheet) overallSheet = workbook.SheetNames[0];

      var sheet = workbook.Sheets[overallSheet];
      if (!sheet) return;

      _analState.rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      _analState.role = role;
      _analState.location = session.location;
      _analState.activeView = null;
      _analState.expandedFO = null;
      renderPage();
    } catch (err) {
      console.error('Analytical load failed:', err);
    }
  };
})();
