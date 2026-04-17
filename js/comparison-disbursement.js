/**
 * comparison-disbursement.js — Disbursement month-over-month comparison.
 * Day-of-month matching (1-1, 2-2 ...). No weekday column.
 * Layout: Summary strip → Dual-line chart → Day-by-day card list.
 */
(function () {
  var MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var AMBER = '#F59E0B', AMBER_DARK = '#D97706';
  var PREV_COLOR = '#94A3B8';

  function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
  function numVal(v) { var n = Number(v); return isFinite(n) ? n : 0; }
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    var n = Number(v); return isFinite(n) ? n.toLocaleString('en-IN') : '-';
  }
  function fmtCr(v) { var n = Number(v) || 0; return (n / 10000000).toFixed(2) + ' Cr'; }
  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];}); }

  function scopeParams() {
    var s = typeof getEmployeeSession === 'function' ? getEmployeeSession() : {};
    var p = [];
    var r = s.role;
    if ((r === 'RM' || r === 'SM') && s.location) p.push('region=' + encodeURIComponent(s.location));
    else if ((r === 'DM' || r === 'DvM') && s.location) p.push('division=' + encodeURIComponent(s.location));
    else if (r === 'AM' && s.location) p.push('area=' + encodeURIComponent(s.location));
    else if (r === 'BM' && s.location) p.push('branch=' + encodeURIComponent(s.location));
    else if ((!r || r === 'FO') && s.id) p.push('emp_id=' + encodeURIComponent(s.id));
    return p;
  }

  function buildRangeUrl(fromISO, toISO) {
    var p = ['from=' + encodeURIComponent(fromISO), 'to=' + encodeURIComponent(toISO)].concat(scopeParams());
    return '/api/disbursement/daily/by-date-range?' + p.join('&');
  }

  function toDayMap(rows, year, month) {
    var map = {};
    if (!rows) return map;
    for (var i = 0; i < rows.length; i++) {
      var iso = String(rows[i].disb_date).substring(0, 10);
      var parts = iso.split('-');
      if (parseInt(parts[0], 10) !== year || parseInt(parts[1], 10) !== month) continue;
      var day = parseInt(parts[2], 10);
      map[day] = { accounts: numVal(rows[i].total_count), amount: numVal(rows[i].total_amount) };
    }
    return map;
  }

  function injectStyles() {
    if (document.getElementById('compDisbStyles')) return;
    var s = document.createElement('style');
    s.id = 'compDisbStyles';
    s.textContent = [
      '.cdc-wrap{max-width:1200px;margin:0 auto;padding:20px 16px;}',
      '.cdc-head{margin-bottom:18px;}',
      '.cdc-title{font-size:20px;font-weight:700;color:#1E293B;margin:0;}',
      '.cdc-sub{font-size:13px;color:#64748B;margin:4px 0 0;}',
      '.cdc-strip{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:20px;}',
      '.cdc-card{background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 1px 4px rgba(15,23,42,0.06);border:1px solid #F1F5F9;}',
      '.cdc-card-label{font-size:11px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px;}',
      '.cdc-card-val{font-size:26px;font-weight:700;color:#1E293B;line-height:1.15;}',
      '.cdc-card-sub{font-size:12px;color:#64748B;margin-top:6px;}',
      '.cdc-card.prev .cdc-card-val{color:#64748B;}',
      '.cdc-card.cur .cdc-card-val{color:' + AMBER_DARK + ';}',
      '.cdc-card.chg{display:flex;flex-direction:column;justify-content:center;}',
      '.cdc-chip{display:inline-flex;align-items:center;gap:6px;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;}',
      '.cdc-chip.up{background:#F0FDF4;color:#059669;}',
      '.cdc-chip.down{background:#FEF2F2;color:#DC2626;}',
      '.cdc-chip.flat{background:#F1F5F9;color:#64748B;}',
      '.cdc-chart-wrap{background:#fff;border-radius:14px;padding:18px;box-shadow:0 1px 4px rgba(15,23,42,0.06);border:1px solid #F1F5F9;margin-bottom:20px;}',
      '.cdc-chart-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px;}',
      '.cdc-chart-title{font-size:14px;font-weight:700;color:#1E293B;}',
      '.cdc-legend{display:flex;gap:14px;font-size:12px;color:#64748B;}',
      '.cdc-dot{display:inline-block;width:10px;height:10px;border-radius:50%;vertical-align:middle;margin-right:6px;}',
      '.cdc-list{display:flex;flex-direction:column;gap:8px;}',
      '.cdc-list-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;}',
      '.cdc-list-title{font-size:14px;font-weight:700;color:#1E293B;}',
      '.cdc-row{display:flex;align-items:center;gap:16px;background:#fff;border-radius:12px;padding:12px 16px;box-shadow:0 1px 3px rgba(15,23,42,0.05);border:1px solid #F1F5F9;}',
      '.cdc-row.alt{background:#FAFAFA;}',
      '.cdc-day{flex-shrink:0;width:40px;height:40px;border-radius:50%;background:#F1F5F9;color:#1E293B;font-size:18px;font-weight:700;display:flex;align-items:center;justify-content:center;}',
      '.cdc-row-body{flex:1;display:flex;gap:16px;align-items:center;flex-wrap:wrap;}',
      '.cdc-side{display:flex;flex-direction:column;min-width:120px;}',
      '.cdc-side-label{font-size:10px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.06em;}',
      '.cdc-side-val{font-size:14px;font-weight:700;color:#1E293B;margin-top:2px;}',
      '.cdc-side-sub{font-size:11px;color:#64748B;margin-top:1px;}',
      '.cdc-side.cur .cdc-side-val{color:' + AMBER_DARK + ';}',
      '.cdc-side.prev .cdc-side-val{color:#475569;}',
      '.cdc-badge{font-size:10px;padding:2px 8px;border-radius:999px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;}',
      '.cdc-badge.new{background:#EEF2FF;color:#4F46E5;}',
      '.cdc-badge.miss{background:#FEF3C7;color:#92400E;}',
      '.cdc-row .cdc-chip{margin-left:auto;}',
      '.cdc-btn{padding:6px 12px;font-size:12px;font-weight:600;background:#fff;border:1px solid #E2E8F0;border-radius:8px;color:#1E293B;cursor:pointer;font-family:inherit;}',
      '.cdc-btn:hover{background:#F8FAFC;}',
      '.cdc-empty{text-align:center;padding:40px;color:#64748B;}',
      '@media (max-width:640px){',
      '.cdc-strip{grid-template-columns:1fr;}',
      '.cdc-card-val{font-size:22px;}',
      '.cdc-row{flex-wrap:wrap;}',
      '.cdc-row-body{width:100%;}',
      '.cdc-side{min-width:100px;}',
      '.cdc-row .cdc-chip{margin-left:0;}',
      '}'
    ].join('');
    document.head.appendChild(s);
  }

  function pct(prev, cur) {
    if (!prev) return null;
    return ((cur - prev) / prev) * 100;
  }

  function chipHtml(prev, cur, unit) {
    if (prev == null && cur == null) return '<span class="cdc-chip flat">—</span>';
    if (!prev && cur) return '<span class="cdc-chip up">&#9650; new</span>';
    if (prev && !cur) return '<span class="cdc-chip down">&#9660; missing</span>';
    var d = pct(prev, cur);
    if (d == null) return '<span class="cdc-chip flat">—</span>';
    var cls = d > 0.1 ? 'up' : d < -0.1 ? 'down' : 'flat';
    var arrow = d > 0.1 ? '&#9650;' : d < -0.1 ? '&#9660;' : '';
    return '<span class="cdc-chip ' + cls + '">' + arrow + ' ' + Math.abs(d).toFixed(1) + '%' + (unit ? ' ' + unit : '') + '</span>';
  }

  function summaryHtml(tot, months) {
    var amtChip = chipHtml(tot.prevAmt, tot.curAmt, '');
    var accChip = chipHtml(tot.prevAcc, tot.curAcc, '');
    var html = '<div class="cdc-strip">';
    html += '<div class="cdc-card prev"><div class="cdc-card-label">' + esc(months.prev.name) + '</div>';
    html += '<div class="cdc-card-val">' + fmtCr(tot.prevAmt) + '</div>';
    html += '<div class="cdc-card-sub">' + fmtNum(tot.prevAcc) + ' accounts</div></div>';

    html += '<div class="cdc-card cur"><div class="cdc-card-label">' + esc(months.cur.name) + '</div>';
    html += '<div class="cdc-card-val">' + fmtCr(tot.curAmt) + '</div>';
    html += '<div class="cdc-card-sub">' + fmtNum(tot.curAcc) + ' accounts</div></div>';

    html += '<div class="cdc-card chg"><div class="cdc-card-label">Change</div>';
    html += '<div style="margin-top:2px;">' + amtChip + '</div>';
    html += '<div class="cdc-card-sub" style="margin-top:8px;">Accounts ' + accChip + '</div>';
    html += '</div>';
    html += '</div>';
    return html;
  }

  function dualLineSvg(prevMap, curMap, months) {
    var daysPrev = lastDay(months.prev.year, months.prev.month);
    var daysCur = lastDay(months.cur.year, months.cur.month);
    var maxDays = Math.max(daysPrev, daysCur);

    var W = 700, H = 260;
    var padL = 40, padR = 20, padT = 20, padB = 34;
    var plotW = W - padL - padR;
    var plotH = H - padT - padB;

    var maxV = 0;
    for (var d = 1; d <= maxDays; d++) {
      if (prevMap[d] && prevMap[d].amount > maxV) maxV = prevMap[d].amount;
      if (curMap[d] && curMap[d].amount > maxV) maxV = curMap[d].amount;
    }
    if (maxV <= 0) maxV = 1;

    function xFor(day) { return padL + ((day - 1) * plotW) / (maxDays - 1); }
    function yFor(amt) { return padT + (1 - amt / maxV) * plotH; }

    function buildPath(map) {
      var parts = [];
      var started = false;
      for (var d = 1; d <= maxDays; d++) {
        if (!map[d]) { started = false; continue; }
        var x = xFor(d), y = yFor(map[d].amount);
        parts.push((started ? 'L' : 'M') + x.toFixed(1) + ' ' + y.toFixed(1));
        started = true;
      }
      return parts.join(' ');
    }

    function buildArea(map) {
      var base = padT + plotH;
      var seg = '';
      var openX = null;
      for (var d = 1; d <= maxDays; d++) {
        if (map[d]) {
          var x = xFor(d), y = yFor(map[d].amount);
          if (openX === null) { seg += 'M' + x.toFixed(1) + ' ' + base.toFixed(1) + ' L' + x.toFixed(1) + ' ' + y.toFixed(1); openX = x; }
          else seg += ' L' + x.toFixed(1) + ' ' + y.toFixed(1);
        } else if (openX !== null) {
          var prevX = xFor(d - 1);
          seg += ' L' + prevX.toFixed(1) + ' ' + base.toFixed(1) + ' Z ';
          openX = null;
        }
      }
      if (openX !== null) {
        var lastX = xFor(maxDays);
        seg += ' L' + lastX.toFixed(1) + ' ' + base.toFixed(1) + ' Z';
      }
      return seg;
    }

    function circles(map, color, isCur) {
      var out = '';
      var label = isCur ? months.cur.name : months.prev.name;
      for (var d = 1; d <= maxDays; d++) {
        if (!map[d]) continue;
        var x = xFor(d), y = yFor(map[d].amount);
        out += '<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="3" fill="' + color + '" stroke="#fff" stroke-width="1.5">' +
               '<title>' + esc(label) + ' · Day ' + d + ' · ' + fmtCr(map[d].amount) + ' · ' + fmtNum(map[d].accounts) + ' acc</title></circle>';
      }
      return out;
    }

    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" style="width:100%;height:clamp(200px,26vw,280px);display:block;overflow:visible;">';
    svg += '<defs><linearGradient id="cdcCurGrad" x1="0" y1="0" x2="0" y2="1">' +
      '<stop offset="0%" stop-color="' + AMBER + '" stop-opacity="0.30"/>' +
      '<stop offset="100%" stop-color="' + AMBER + '" stop-opacity="0"/>' +
    '</linearGradient></defs>';

    // X-axis labels (every 5 days + 1 + last)
    var ticks = [1];
    for (var t = 5; t <= maxDays; t += 5) ticks.push(t);
    if (ticks[ticks.length - 1] !== maxDays) ticks.push(maxDays);
    for (var ti = 0; ti < ticks.length; ti++) {
      var tx = xFor(ticks[ti]);
      svg += '<text x="' + tx.toFixed(1) + '" y="' + (H - 14) + '" text-anchor="middle" style="font-size:11px;fill:#94A3B8;">' + ticks[ti] + '</text>';
    }
    // Y-axis mid + top gridline
    var gridY1 = padT + plotH / 2;
    svg += '<line x1="' + padL + '" y1="' + gridY1 + '" x2="' + (W - padR) + '" y2="' + gridY1 + '" stroke="#F1F5F9" stroke-width="1"/>';
    svg += '<line x1="' + padL + '" y1="' + padT + '" x2="' + (W - padR) + '" y2="' + padT + '" stroke="#F1F5F9" stroke-width="1"/>';
    svg += '<text x="' + (padL - 6) + '" y="' + (padT + 4) + '" text-anchor="end" style="font-size:10px;fill:#94A3B8;">' + fmtCr(maxV) + '</text>';
    svg += '<text x="' + (padL - 6) + '" y="' + (gridY1 + 4) + '" text-anchor="end" style="font-size:10px;fill:#94A3B8;">' + fmtCr(maxV / 2) + '</text>';

    // Cur area + cur line (amber) behind, prev line on top
    svg += '<path d="' + buildArea(curMap) + '" fill="url(#cdcCurGrad)"/>';
    svg += '<path d="' + buildPath(prevMap) + '" fill="none" stroke="' + PREV_COLOR + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" stroke-dasharray="4 3"/>';
    svg += '<path d="' + buildPath(curMap) + '" fill="none" stroke="' + AMBER + '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>';
    svg += circles(prevMap, PREV_COLOR, false);
    svg += circles(curMap, AMBER, true);
    svg += '</svg>';
    return svg;
  }

  function listHtml(prevMap, curMap, months) {
    var days = [];
    for (var d = 1; d <= 31; d++) if (prevMap[d] || curMap[d]) days.push(d);
    if (!days.length) return '<div class="cdc-empty">No disbursement data for these months.</div>';

    var html = '';
    for (var i = 0; i < days.length; i++) {
      var day = days[i];
      var p = prevMap[day] || null;
      var c = curMap[day] || null;
      var alt = i % 2 === 1 ? ' alt' : '';
      html += '<div class="cdc-row' + alt + '">';
      html += '<div class="cdc-day">' + day + '</div>';
      html += '<div class="cdc-row-body">';

      // Prev side
      html += '<div class="cdc-side prev">';
      html += '<span class="cdc-side-label">' + esc(months.prev.name) + '</span>';
      if (p) {
        html += '<span class="cdc-side-val">' + fmtCr(p.amount) + '</span>';
        html += '<span class="cdc-side-sub">' + fmtNum(p.accounts) + ' acc</span>';
      } else {
        html += '<span class="cdc-side-val" style="color:#CBD5E1;">—</span>';
        html += '<span class="cdc-side-sub"><span class="cdc-badge new">new on ' + esc(months.cur.name.split(' ')[0]) + '</span></span>';
      }
      html += '</div>';

      // Cur side
      html += '<div class="cdc-side cur">';
      html += '<span class="cdc-side-label">' + esc(months.cur.name) + '</span>';
      if (c) {
        html += '<span class="cdc-side-val">' + fmtCr(c.amount) + '</span>';
        html += '<span class="cdc-side-sub">' + fmtNum(c.accounts) + ' acc</span>';
      } else {
        html += '<span class="cdc-side-val" style="color:#CBD5E1;">—</span>';
        html += '<span class="cdc-side-sub"><span class="cdc-badge miss">no data</span></span>';
      }
      html += '</div>';

      html += chipHtml(p ? p.amount : null, c ? c.amount : null, '');
      html += '</div></div>';
    }
    return html;
  }

  function computeCsv(prevMap, curMap, months) {
    var lines = ['Day,' + months.prev.name + ' Accounts,' + months.prev.name + ' Amount,' + months.cur.name + ' Accounts,' + months.cur.name + ' Amount,Diff %'];
    for (var d = 1; d <= 31; d++) {
      if (!prevMap[d] && !curMap[d]) continue;
      var p = prevMap[d] || { accounts: '', amount: '' };
      var c = curMap[d] || { accounts: '', amount: '' };
      var diff = (p.amount && c.amount) ? (((c.amount - p.amount) / p.amount) * 100).toFixed(1) : '';
      lines.push([d, p.accounts, p.amount, c.accounts, c.amount, diff].join(','));
    }
    return lines.join('\n');
  }

  window._cdcDownloadCsv = function (csv, filename) {
    var blob = new Blob([csv], { type: 'text/csv' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 100);
  };

  window.renderDisbComparison = function () {
    injectStyles();
    var container = document.getElementById('comparisonContent');
    if (!container) return;
    var body = document.getElementById('compDisbBody');
    if (!body) {
      var bar = document.getElementById('compSubTabBar');
      body = document.createElement('div');
      body.id = 'compDisbBody';
      if (bar && bar.parentNode === container) container.appendChild(body);
      else { container.innerHTML = ''; container.appendChild(body); }
    }
    body.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:' + AMBER + ';border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading...</div></div>';

    var datesUrl = '/api/disbursement/daily/dates';
    var sp = scopeParams();
    if (sp.length) datesUrl += '?' + sp.join('&');

    fetch(datesUrl).then(function (r) { return r.json(); }).then(function (dates) {
      if (!dates || !dates.length) {
        body.innerHTML = '<div class="cdc-wrap"><div class="cdc-empty">No disbursement data available.</div></div>';
        return;
      }
      var sorted = dates.slice().sort();
      var latest = String(sorted[sorted.length - 1]).substring(0, 10);
      var lp = latest.split('-');
      var cy = parseInt(lp[0], 10), cm = parseInt(lp[1], 10);
      var py = cm === 1 ? cy - 1 : cy, pm = cm === 1 ? 12 : cm - 1;
      var months = {
        cur: { year: cy, month: cm, name: MONTH_NAMES[cm] + ' ' + cy },
        prev: { year: py, month: pm, name: MONTH_NAMES[pm] + ' ' + py }
      };
      var curFrom = cy + '-' + pad2(cm) + '-01';
      var curTo = cy + '-' + pad2(cm) + '-' + pad2(lastDay(cy, cm));
      var prevFrom = py + '-' + pad2(pm) + '-01';
      var prevTo = py + '-' + pad2(pm) + '-' + pad2(lastDay(py, pm));

      Promise.all([
        fetch(buildRangeUrl(prevFrom, prevTo)).then(function (r) { return r.json(); }),
        fetch(buildRangeUrl(curFrom, curTo)).then(function (r) { return r.json(); })
      ]).then(function (arr) {
        var prevMap = toDayMap(arr[0], py, pm);
        var curMap = toDayMap(arr[1], cy, cm);

        var tot = { prevAcc: 0, prevAmt: 0, curAcc: 0, curAmt: 0 };
        for (var d = 1; d <= 31; d++) {
          if (prevMap[d]) { tot.prevAcc += prevMap[d].accounts; tot.prevAmt += prevMap[d].amount; }
          if (curMap[d]) { tot.curAcc += curMap[d].accounts; tot.curAmt += curMap[d].amount; }
        }

        var csv = computeCsv(prevMap, curMap, months);
        window.__cdcCsv = csv;
        var filename = 'disbursement-comparison-' + months.prev.name.replace(' ', '') + '-vs-' + months.cur.name.replace(' ', '') + '.csv';

        var html = '<div class="cdc-wrap">';
        html += '<div class="cdc-head">';
        html += '<h2 class="cdc-title">Disbursement Comparison</h2>';
        html += '<p class="cdc-sub">' + esc(months.prev.name) + ' vs ' + esc(months.cur.name) + ' · Day-of-month matched</p>';
        html += '</div>';

        html += summaryHtml(tot, months);

        html += '<div class="cdc-chart-wrap">';
        html += '<div class="cdc-chart-head">';
        html += '<div class="cdc-chart-title">Daily Disbursement Trend</div>';
        html += '<div class="cdc-legend">';
        html += '<span><span class="cdc-dot" style="background:' + PREV_COLOR + ';"></span>' + esc(months.prev.name) + '</span>';
        html += '<span><span class="cdc-dot" style="background:' + AMBER + ';"></span>' + esc(months.cur.name) + '</span>';
        html += '</div></div>';
        html += dualLineSvg(prevMap, curMap, months);
        html += '</div>';

        html += '<div class="cdc-list">';
        html += '<div class="cdc-list-head">';
        html += '<div class="cdc-list-title">Day-by-Day</div>';
        html += '<button class="cdc-btn" onclick="window._cdcDownloadCsv(window.__cdcCsv, \'' + esc(filename) + '\')">Download CSV</button>';
        html += '</div>';
        html += listHtml(prevMap, curMap, months);
        html += '</div>';

        html += '</div>';
        body.innerHTML = html;
      }).catch(function () {
        body.innerHTML = '<div class="cdc-wrap"><div class="cdc-empty">Failed to load disbursement data.</div></div>';
      });
    }).catch(function () {
      body.innerHTML = '<div class="cdc-wrap"><div class="cdc-empty">Failed to load disbursement dates.</div></div>';
    });
  };
})();
