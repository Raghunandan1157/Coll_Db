(function () {
  'use strict';

  // Inject CSS if not already present
  if (!document.getElementById('analytical-styles')) {
    var style = document.createElement('style');
    style.id = 'analytical-styles';
    style.textContent = [
      '.anal-container { padding: 12px 10px 80px; }',
      '.anal-page-title { font-size: 20px; font-weight: 700; color: #E8ECF4; margin-bottom: 4px; }',
      '.anal-page-subtitle { font-size: 12px; color: #6B7A99; margin-bottom: 20px; }',
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
      '  height: 1px; background: rgba(255,255,255,0.04); margin: 20px 0;',
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
      html += '<div class="anal-empty">No branch data available</div>';
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
      branchData.push({ name: b.name, demand: dem, collection: col, balance: bal, pct: pctRaw });
    }

    // Sort by percentage
    branchData.sort(function (a, b) { return a.pct - b.pct; });

    var bottom10 = branchData.slice(0, 10);
    var top10 = branchData.slice(-10).reverse();

    // SVG icons
    var upArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>';
    var downArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></svg>';

    // Build page
    var html = '<div class="anal-container">';
    html += '<div class="anal-page-title">Analytical Tool</div>';
    html += '<div class="anal-page-subtitle">Branch performance analysis based on regular collection %</div>';

    // Top 10 section
    html += buildSectionHtml('Top 10 Branches', upArrow, 'top', top10);

    // Divider
    html += '<div class="anal-divider"></div>';

    // Bottom 10 section
    html += buildSectionHtml('Bottom 10 Branches', downArrow, 'bottom', bottom10);

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
