/**
 * comparison-disbursement.js — Disbursement month-over-month comparison.
 * Mirrors Collection comparison visual language. Day-of-month matching (no weekday).
 */
(function () {
  var MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
  function numVal(v) { var n = Number(v); return isFinite(n) ? n : 0; }
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    var n = Number(v); return isFinite(n) ? n.toLocaleString('en-IN') : '-';
  }
  function fmtCr(v) { var n = Number(v) || 0; return (n / 10000000).toFixed(2) + ' Cr'; }
  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }
  function ordinal(n) {
    var s = ['th','st','nd','rd'], v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];}); }

  var _view = 'cards', _dayIdx = -1;
  var _months = null, _prevMap = {}, _curMap = {}, _days = [];

  function scopeParams() {
    var s = typeof getEmployeeSession === 'function' ? getEmployeeSession() : {};
    var p = [];
    var r = s.role;
    var _isNew = typeof isNewStructure === 'function' ? isNewStructure() : (window.getDataStructure && window.getDataStructure() === 'new');
    if (!_isNew) p.push('structure=old');
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

  function diffPct(prev, cur) {
    if (!prev) return null;
    return ((cur - prev) / prev) * 100;
  }

  function diffCell(prev, cur, tdStyle) {
    var base = 'padding:8px;text-align:right;border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;font-weight:700;' + (tdStyle || '');
    if (prev == null && cur == null) return '<td style="' + base + 'color:#CBD5E1;">-</td>';
    if (!prev && cur) return '<td style="' + base + 'color:#6366F1;">new</td>';
    if (prev && !cur) return '<td style="' + base + 'color:#DC2626;">missing</td>';
    var d = diffPct(prev, cur);
    var color = d >= 0 ? '#059669' : '#EF4444';
    var arrow = d > 0 ? '&#9650; ' : d < 0 ? '&#9660; ' : '';
    return '<td style="' + base + 'color:' + color + ';">' + arrow + Math.abs(d).toFixed(1) + '%</td>';
  }

  /* ========== CARDS VIEW ========== */
  function renderCards() {
    var day = _days[_dayIdx];
    if (!day) return '<div style="text-align:center;padding:60px;color:#64748B;">No data for this day.</div>';
    var p = _prevMap[day] || null;
    var c = _curMap[day] || null;
    if (!p && !c) return '<div style="text-align:center;padding:60px;color:#64748B;">No data for this day.</div>';

    var buckets = [
      { name: 'Accounts', color: '#6366F1', pVal: p ? p.accounts : 0, cVal: c ? c.accounts : 0, fmt: fmtNum, unit: 'accounts' },
      { name: 'Amount',   color: '#F59E0B', pVal: p ? p.amount : 0,   cVal: c ? c.amount : 0,   fmt: fmtCr,  unit: '' }
    ];

    var html = '<style>';
    html += '.cdc-card{transition:box-shadow .2s;}.cdc-card:hover{box-shadow:0 4px 16px rgba(0,0,0,0.08);}';
    html += '</style>';

    html += '<div style="position:relative;border:1px solid #E2E8F0;border-radius:16px;overflow:hidden;background:#fff;">';

    // Vertical band backgrounds
    html += '<div style="position:absolute;top:0;bottom:0;left:180px;right:0;display:flex;pointer-events:none;z-index:0;">';
    html += '<div style="flex:1;background:rgba(139,92,246,0.04);border-left:1px solid #F1F5F9;border-right:1px solid #F1F5F9;"></div>';
    html += '<div style="flex:1;background:rgba(16,185,129,0.04);border-right:1px solid #F1F5F9;"></div>';
    html += '<div style="width:140px;"></div>';
    html += '</div>';

    // Header row
    html += '<div style="display:flex;align-items:stretch;border-bottom:2px solid #E2E8F0;position:relative;z-index:1;">';
    html += '<div style="min-width:180px;padding:12px 20px;"></div>';
    html += '<div style="flex:1;text-align:center;padding:10px 0;font-size:12px;font-weight:700;color:#7C3AED;text-transform:uppercase;letter-spacing:1.5px;background:rgba(139,92,246,0.06);">';
    html += _months.prev.name;
    html += '<div style="font-size:10px;font-weight:500;color:#A78BFA;margin-top:1px;">' + (p ? ordinal(day) + ' ' + MONTH_NAMES[_months.prev.month] : 'No data') + '</div>';
    html += '</div>';
    html += '<div style="flex:1;text-align:center;padding:10px 0;font-size:12px;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:1.5px;background:rgba(16,185,129,0.06);">';
    html += _months.cur.name + '<div style="font-size:10px;font-weight:500;color:#34D399;margin-top:1px;">' + (c ? ordinal(day) + ' ' + MONTH_NAMES[_months.cur.month] : 'No data') + '</div>';
    html += '</div>';
    html += '<div style="width:140px;text-align:center;padding:12px 0;font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Diff</div>';
    html += '</div>';

    for (var i = 0; i < buckets.length; i++) {
      var b = buckets[i];
      var diff = b.cVal - b.pVal;
      var dPct = diffPct(b.pVal, b.cVal);
      var dColor = diff >= 0 ? '#059669' : '#EF4444';
      var dBg = diff >= 0 ? '#F0FDF4' : '#FEF2F2';
      var dArrow = diff > 0 ? '&#9650; ' : diff < 0 ? '&#9660; ' : '';
      var borderTop = i > 0 ? 'border-top:1px solid #F1F5F9;' : '';

      html += '<div class="cdc-card" style="display:flex;align-items:stretch;position:relative;z-index:1;' + borderTop + '">';
      html += '<div style="min-width:180px;padding:20px 20px;display:flex;align-items:center;gap:10px;">';
      html += '<div style="width:4px;height:32px;border-radius:4px;background:' + b.color + ';flex-shrink:0;"></div>';
      html += '<div style="font-size:15px;font-weight:700;color:#1E293B;">' + b.name + '</div>';
      html += '</div>';

      html += '<div style="flex:1;padding:16px 12px;text-align:center;display:flex;flex-direction:column;justify-content:center;">';
      html += '<div style="font-family:\'Playfair Display\',serif;font-size:24px;font-weight:700;color:#FB923C;">' + (p ? b.fmt(b.pVal) : '-') + '</div>';
      if (p && b.unit) html += '<div style="font-size:11px;color:#A78BFA;margin-top:2px;">' + b.unit + '</div>';
      html += '</div>';

      html += '<div style="flex:1;padding:16px 12px;text-align:center;display:flex;flex-direction:column;justify-content:center;">';
      html += '<div style="font-family:\'Playfair Display\',serif;font-size:24px;font-weight:700;color:#FB923C;">' + (c ? b.fmt(b.cVal) : '-') + '</div>';
      if (c && b.unit) html += '<div style="font-size:11px;color:#34D399;margin-top:2px;">' + b.unit + '</div>';
      html += '</div>';

      html += '<div style="width:140px;padding:16px 12px;text-align:center;display:flex;align-items:center;justify-content:center;">';
      if (p && c) {
        html += '<div style="background:' + dBg + ';border-radius:10px;padding:8px 14px;">';
        html += '<div style="font-size:17px;font-weight:800;color:' + dColor + ';">' + dArrow + (dPct != null ? Math.abs(dPct).toFixed(1) + '%' : '-') + '</div>';
        html += '<div style="font-size:9px;text-transform:uppercase;letter-spacing:1px;color:' + dColor + ';margin-top:1px;">' + (diff >= 0 ? 'Higher' : 'Lower') + '</div>';
        html += '</div>';
      } else if (!p && c) {
        html += '<div style="background:#EEF2FF;border-radius:10px;padding:8px 14px;color:#4F46E5;font-size:13px;font-weight:700;">new</div>';
      } else if (p && !c) {
        html += '<div style="background:#FEF2F2;border-radius:10px;padding:8px 14px;color:#DC2626;font-size:13px;font-weight:700;">missing</div>';
      } else {
        html += '<div style="color:#CBD5E1;font-size:13px;">-</div>';
      }
      html += '</div></div>';
    }

    html += '</div>';
    return html;
  }

  /* ========== TABLE VIEW ========== */
  function renderTable() {
    var tot = { pAcc: 0, pAmt: 0, cAcc: 0, cAmt: 0 };
    var prevLastDay = 0, curLastDay = 0;
    for (var pd in _prevMap) if (+pd > prevLastDay) prevLastDay = +pd;
    for (var cd in _curMap) if (+cd > curLastDay) curLastDay = +cd;
    var prevCum = 0, curCum = 0;
    var beyondCell = '<td style="padding:8px;text-align:right;border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;font-weight:700;color:#CBD5E1;">&mdash;</td>';

    var html = '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E2E8F0;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:#F8FAFC;">';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#7C3AED;font-weight:700;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.04);" colspan="3">' + esc(_months.prev.name) + '</th>';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#059669;font-weight:700;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.04);" colspan="3">' + esc(_months.cur.name) + '</th>';
    html += '<th rowspan="2" style="padding:10px 12px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:11px;font-weight:700;border-left:2px solid #CBD5E1;background:#F8FAFC;">Diff %</th>';
    html += '</tr><tr style="background:#F8FAFC;">';
    var heads = ['Date', 'Accounts', 'Amount'];
    for (var h = 0; h < 2; h++) {
      for (var s = 0; s < heads.length; s++) {
        var bl = s === 0 ? (h === 0 ? 'border-left:1px solid #E2E8F0;' : 'border-left:2px solid #CBD5E1;') : '';
        var bg = h === 0 ? 'background:rgba(139,92,246,0.04);' : 'background:rgba(16,185,129,0.04);';
        var align = s === 0 ? 'text-align:center;' : 'text-align:right;';
        html += '<th style="padding:6px 8px;' + align + 'border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;' + bl + bg + '">' + heads[s] + '</th>';
      }
    }
    html += '</tr></thead><tbody>';

    if (!_days.length) {
      html += '<tr><td colspan="7" style="padding:40px;text-align:center;color:#64748B;">No disbursement data for these months.</td></tr>';
    }

    for (var r = 0; r < _days.length; r++) {
      var day = _days[r];
      var p = _prevMap[day] || null;
      var c = _curMap[day] || null;
      var rowBg = r % 2 === 0 ? '' : 'background:#FAFAFA;';
      if (p) { tot.pAcc += p.accounts; tot.pAmt += p.amount; prevCum += p.amount; }
      if (c) { tot.cAcc += c.accounts; tot.cAmt += c.amount; curCum += c.amount; }

      var showPrevAmt = day <= prevLastDay;
      var showCurAmt = day <= curLastDay;
      var beyondCur = day > curLastDay;

      html += '<tr style="' + rowBg + '">';
      // Prev block
      html += '<td style="padding:8px 6px;text-align:center;font-weight:500;color:#7C3AED;border-bottom:1px solid #F1F5F9;border-left:1px solid #E2E8F0;font-size:12px;white-space:nowrap;background:rgba(139,92,246,0.02);' + (p ? '' : 'opacity:0.4;') + '">' + ordinal(day) + ' ' + MONTH_NAMES[_months.prev.month] + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:500;border-bottom:1px solid #F1F5F9;background:rgba(139,92,246,0.02);font-variant-numeric:tabular-nums;' + (p ? 'color:#1E293B;' : 'color:#CBD5E1;') + '">' + (p ? fmtNum(p.accounts) : '-') + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:600;border-bottom:1px solid #F1F5F9;background:rgba(139,92,246,0.02);font-variant-numeric:tabular-nums;' + (showPrevAmt ? 'color:#7C3AED;' : 'color:#CBD5E1;') + '">' + (showPrevAmt ? fmtCr(prevCum) : '-') + '</td>';
      // Cur block
      html += '<td style="padding:8px 6px;text-align:center;font-weight:500;color:#059669;border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;font-size:12px;white-space:nowrap;background:rgba(16,185,129,0.02);' + (c ? '' : 'opacity:0.4;') + '">' + ordinal(day) + ' ' + MONTH_NAMES[_months.cur.month] + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:500;border-bottom:1px solid #F1F5F9;background:rgba(16,185,129,0.02);font-variant-numeric:tabular-nums;' + (c ? 'color:#1E293B;' : 'color:#CBD5E1;') + '">' + (c ? fmtNum(c.accounts) : '-') + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:600;border-bottom:1px solid #F1F5F9;background:rgba(16,185,129,0.02);font-variant-numeric:tabular-nums;' + (showCurAmt ? 'color:#059669;' : 'color:#CBD5E1;') + '">' + (showCurAmt ? fmtCr(curCum) : '-') + '</td>';
      if (beyondCur) html += beyondCell;
      else html += diffCell(showPrevAmt ? prevCum : null, showCurAmt ? curCum : null);
      html += '</tr>';
    }

    if (_days.length) {
      html += '<tr style="background:#F1F5F9;border-top:2px solid #CBD5E1;">';
      html += '<td style="padding:10px 8px;text-align:center;font-weight:800;color:#1E293B;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.08);">Total</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:700;color:#1E293B;background:rgba(139,92,246,0.08);font-variant-numeric:tabular-nums;">' + fmtNum(tot.pAcc) + '</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:800;color:#7C3AED;background:rgba(139,92,246,0.08);font-variant-numeric:tabular-nums;">' + fmtCr(tot.pAmt) + '</td>';
      html += '<td style="padding:10px 8px;text-align:center;font-weight:800;color:#1E293B;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.08);">Total</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:700;color:#1E293B;background:rgba(16,185,129,0.08);font-variant-numeric:tabular-nums;">' + fmtNum(tot.cAcc) + '</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:800;color:#059669;background:rgba(16,185,129,0.08);font-variant-numeric:tabular-nums;">' + fmtCr(tot.cAmt) + '</td>';
      html += diffCell(tot.pAmt || null, tot.cAmt || null, 'font-weight:800;');
      html += '</tr>';
    }

    html += '</tbody></table></div>';
    return html;
  }

  /* ========== MAIN RENDER ========== */
  function render() {
    var body = document.getElementById('compDisbBody');
    if (!body) return;
    if (!_months) {
      body.innerHTML = '<div style="text-align:center;padding:80px 20px;color:#64748B;">No disbursement data available.</div>';
      return;
    }
    var isCards = _view === 'cards';
    var day = _days[_dayIdx];
    var p = day ? (_prevMap[day] || null) : null;
    var c = day ? (_curMap[day] || null) : null;

    var html = '<div style="max-width:1400px;margin:0 auto;padding:16px;">';

    // Header
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px;">';
    html += '<div><h2 style="font-size:20px;font-weight:700;color:#1E293B;margin:0;">Disbursement Comparison</h2>';
    html += '<p style="font-size:13px;color:#64748B;margin:4px 0 0;">' + esc(_months.prev.name) + ' vs ' + esc(_months.cur.name) + ' &mdash; Day wise</p></div>';

    // Day nav
    html += '<div style="display:flex;align-items:center;gap:6px;">';
    html += '<button onclick="window._cdcNav(-1)" style="width:34px;height:34px;border:1px solid #E2E8F0;border-radius:8px;background:#fff;cursor:pointer;font-size:16px;color:#F59E0B;font-weight:700;">&larr;</button>';
    html += '<div style="background:#F1F5F9;border:1px solid #E2E8F0;padding:8px 18px;border-radius:10px;text-align:center;min-width:200px;">';
    if (day) {
      html += '<div style="font-size:15px;font-weight:700;color:#1E293B;">Day ' + day + '</div>';
      html += '<div style="font-size:11px;color:#94A3B8;margin-top:2px;">';
      html += (p ? ordinal(day) + ' ' + MONTH_NAMES[_months.prev.month] : 'No data') + '  vs  ' + (c ? ordinal(day) + ' ' + MONTH_NAMES[_months.cur.month] : 'No data');
      html += '</div>';
    } else {
      html += '<div style="font-size:13px;color:#94A3B8;">No day selected</div>';
    }
    html += '</div>';
    html += '<button onclick="window._cdcNav(1)" style="width:34px;height:34px;border:1px solid #E2E8F0;border-radius:8px;background:#fff;cursor:pointer;font-size:16px;color:#F59E0B;font-weight:700;">&rarr;</button>';
    html += '</div>';

    // Toggle
    html += '<div style="display:inline-flex;border:1px solid #E2E8F0;border-radius:8px;overflow:hidden;">';
    html += '<button onclick="window._cdcSetView(\'cards\')" style="padding:6px 14px;font-size:12px;font-weight:' + (isCards ? '600' : '500') + ';color:' + (isCards ? '#D97706' : '#64748B') + ';background:' + (isCards ? '#FFFBEB' : '#fff') + ';border:none;cursor:pointer;font-family:inherit;">Cards</button>';
    html += '<button onclick="window._cdcSetView(\'table\')" style="padding:6px 14px;font-size:12px;font-weight:' + (!isCards ? '600' : '500') + ';color:' + (!isCards ? '#D97706' : '#64748B') + ';background:' + (!isCards ? '#FFFBEB' : '#fff') + ';border:none;border-left:1px solid #E2E8F0;cursor:pointer;font-family:inherit;">Table</button>';
    html += '</div></div>';

    html += isCards ? renderCards() : renderTable();
    html += '</div>';
    body.innerHTML = html;
  }

  window._cdcNav = function (dir) {
    if (!_days.length) return;
    var n = _dayIdx + dir;
    if (n >= 0 && n < _days.length) { _dayIdx = n; render(); }
  };
  window._cdcSetView = function (v) { _view = v; render(); };

  window.renderDisbComparison = function () {
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
    body.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#F59E0B;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading...</div></div>';

    var datesUrl = '/api/disbursement/daily/dates';
    var sp = scopeParams();
    if (sp.length) datesUrl += '?' + sp.join('&');

    fetch(datesUrl).then(function (r) { return r.json(); }).then(function (dates) {
      if (!dates || !dates.length) {
        body.innerHTML = '<div style="text-align:center;padding:80px 20px;color:#64748B;">No disbursement data available.</div>';
        return;
      }
      var sorted = dates.slice().sort();
      var latest = String(sorted[sorted.length - 1]).substring(0, 10);
      var lp = latest.split('-');
      var cy = parseInt(lp[0], 10), cm = parseInt(lp[1], 10);
      var py = cm === 1 ? cy - 1 : cy, pm = cm === 1 ? 12 : cm - 1;
      _months = {
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
        _prevMap = toDayMap(arr[0], py, pm);
        _curMap = toDayMap(arr[1], cy, cm);
        _days = [];
        for (var d = 1; d <= 31; d++) if (_prevMap[d] || _curMap[d]) _days.push(d);
        // Default to latest cur-month day with data; else last day in _days
        var defaultIdx = _days.length - 1;
        for (var i = _days.length - 1; i >= 0; i--) {
          if (_curMap[_days[i]]) { defaultIdx = i; break; }
        }
        _dayIdx = defaultIdx;
        render();
      }).catch(function () {
        body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load disbursement data.</div>';
      });
    }).catch(function () {
      body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load disbursement dates.</div>';
    });
  };
})();
