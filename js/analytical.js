(function () {
  'use strict';

  // Inject CSS if not already present
  if (!document.getElementById('analytical-styles')) {
    var style = document.createElement('style');
    style.id = 'analytical-styles';
    style.textContent = [
      '.anal-container { padding: 16px; }',
      '.anal-header { margin-bottom: 20px; }',
      '.anal-title { font-size: 20px; font-weight: 700; color: #E8ECF4; }',
      '.anal-subtitle { font-size: 13px; color: #6B7A99; margin-top: 4px; }',
      '.anal-empty { text-align: center; color: #6B7A99; padding: 40px 0; }',
      '.anal-list { display: flex; flex-direction: column; gap: 12px; }',
      '.anal-card {',
      '  background: #131825;',
      '  border: 1px solid rgba(255,255,255,0.06);',
      '  border-radius: 14px;',
      '  padding: 16px;',
      '  transition: border-color 0.2s, transform 0.2s;',
      '}',
      '.anal-card:hover {',
      '  border-color: rgba(255,255,255,0.12);',
      '  transform: translateY(-1px);',
      '}',
      '.anal-card-header {',
      '  display: flex;',
      '  align-items: center;',
      '  gap: 10px;',
      '  margin-bottom: 12px;',
      '}',
      '.anal-rank {',
      '  font-size: 12px;',
      '  font-weight: 700;',
      '  padding: 4px 8px;',
      '  border-radius: 8px;',
      '  min-width: 32px;',
      '  text-align: center;',
      '}',
      '.anal-branch-avatar {',
      '  width: 32px;',
      '  height: 32px;',
      '  border-radius: 50%;',
      '  background: rgba(79,140,255,0.15);',
      '  color: #4F8CFF;',
      '  display: flex;',
      '  align-items: center;',
      '  justify-content: center;',
      '  font-weight: 700;',
      '  font-size: 14px;',
      '  flex-shrink: 0;',
      '}',
      '.anal-branch-name {',
      '  flex: 1;',
      '  font-size: 14px;',
      '  font-weight: 600;',
      '  color: #E8ECF4;',
      '  white-space: nowrap;',
      '  overflow: hidden;',
      '  text-overflow: ellipsis;',
      '}',
      '.anal-pct {',
      '  font-size: 18px;',
      '  font-weight: 800;',
      '  font-family: "DM Sans", sans-serif;',
      '}',
      '.anal-card-stats {',
      '  display: flex;',
      '  justify-content: space-between;',
      '  margin-bottom: 10px;',
      '}',
      '.anal-stat {',
      '  display: flex;',
      '  flex-direction: column;',
      '  align-items: center;',
      '  gap: 2px;',
      '}',
      '.anal-stat-lbl {',
      '  font-size: 10px;',
      '  text-transform: uppercase;',
      '  color: #6B7A99;',
      '  letter-spacing: 0.5px;',
      '}',
      '.anal-stat-val {',
      '  font-size: 14px;',
      '  font-weight: 700;',
      '  color: #E8ECF4;',
      '}'
    ].join('\n');
    document.head.appendChild(style);
  }

  // Reuse same COL indices as collection.js
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

  // Get all branch rows — mirrors collection.js getAllSectionRows logic exactly
  function getAllBranchRows(rows) {
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
        if (currentSection === 'branch' && foundTarget) break;
        currentSection = 'region'; continue;
      }
      if (hdr.includes('DISTRICT') && (hdr.includes('WISE') || hdr.includes('NAME'))) {
        if (currentSection === 'branch' && foundTarget) break;
        currentSection = 'district'; continue;
      }
      if (hdr.includes('BRANCH') && !hdr.includes('OFFICER') && (hdr.includes('WISE') || hdr.includes('NAME'))) {
        if (currentSection === 'branch' && foundTarget) break;
        currentSection = 'branch'; continue;
      }
      if (c0.includes('GRAND TOTAL') || c1.includes('GRAND TOTAL')) {
        if (currentSection === 'branch' && foundTarget) break;
        currentSection = null; continue;
      }
      if (currentSection !== 'branch') continue;
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

  function renderAnalytical(rows) {
    var container = document.getElementById('analyticalContent');
    if (!container) return;

    var branches = getAllBranchRows(rows);

    // Calculate collection % for each branch
    var branchData = [];
    for (var i = 0; i < branches.length; i++) {
      var b = branches[i];
      var dem = numVal(b.row[COL.regDemand]);
      var col = numVal(b.row[COL.regCollection]);
      var bal = dem - col;
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      branchData.push({
        name: b.name,
        demand: dem,
        collection: col,
        balance: bal,
        pct: pctRaw
      });
    }

    // Sort ascending by percentage (lowest first) and take bottom 10
    branchData.sort(function (a, b) { return a.pct - b.pct; });
    var lowest10 = branchData.slice(0, 10);

    // Build HTML
    var html = '<div class="anal-container">';
    html += '<div class="anal-header">';
    html += '<div class="anal-title">Analytical Tool</div>';
    html += '<div class="anal-subtitle">Lowest 10 Performing Branches</div>';
    html += '</div>';

    if (lowest10.length === 0) {
      html += '<div class="anal-empty">No branch data available</div>';
    } else {
      html += '<div class="anal-list">';
      for (var j = 0; j < lowest10.length; j++) {
        var item = lowest10[j];
        var pctColor = item.pct >= 95 ? '#34D399' : item.pct >= 80 ? '#FBBF24' : '#F87171';
        var rankNum = j + 1;
        var barW = Math.min(item.pct, 100);
        var initial = item.name.charAt(0).toUpperCase();

        html += '<div class="anal-card emp-fade" style="animation-delay:' + (j * 0.05) + 's">';
        html += '<div class="anal-card-header">';
        html += '<div class="anal-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>';
        html += '<div class="anal-branch-avatar">' + initial + '</div>';
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
      }
      html += '</div>';
    }

    html += '</div>';
    container.innerHTML = html;
  }

  // Load function exposed globally — mirrors collection.js _loadCollectionTab pattern
  window._loadAnalyticalTab = async function () {
    try {
      if (typeof getWorkbookWithFallback !== 'function') return;
      var wb = await getWorkbookWithFallback();
      if (!wb || !wb.data) return;

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });
      if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return;

      // Find the OverAll sheet (not On-Date, not FY, not tom_)
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
      renderAnalytical(rows);
    } catch (err) {
      console.error('Analytical load failed:', err);
    }
  };
})();
