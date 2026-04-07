/**
 * comparison.js — Month-over-month comparison tab.
 * Compares by weekday occurrence: 1-Mon vs 1-Mon, 2-Tue vs 2-Tue, etc.
 */
(function () {

  var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    var n = Number(v); return isFinite(n) ? n.toLocaleString('en-IN') : '-';
  }
  function pctStr(d, c) { return d > 0 ? ((c / d) * 100).toFixed(1) + '%' : '-'; }
  function pctColor(d, c) {
    if (!d) return '#64748B'; var p = (c / d) * 100;
    return p >= 95 ? '#34D399' : p >= 80 ? '#FBBF24' : '#F87171';
  }
  function fmtDate(dateStr) {
    if (!dateStr) return '-';
    var p = dateStr.split('-'); return parseInt(p[2]) + ' ' + MONTH_NAMES[parseInt(p[1])];
  }

  /**
   * Build working days for a month with weekday occurrence labels.
   * Each day gets a label like "1 - Mon" (1st Monday), "2 - Tue" (2nd Tuesday).
   * Only includes days that have data in dateMap.
   */
  function buildLabeledDays(dateMap, year, month) {
    var days = [], daysInMonth = new Date(year, month, 0).getDate();
    var weekdayCount = {};
    for (var d = 1; d <= daysInMonth; d++) {
      var mm = String(month).padStart(2, '0'), dd = String(d).padStart(2, '0');
      var ds = year + '-' + mm + '-' + dd;
      if (!dateMap[ds]) continue;
      var dow = new Date(year, month - 1, d).getDay();
      var dn = DAY_NAMES[dow];
      weekdayCount[dn] = (weekdayCount[dn] || 0) + 1;
      days.push({
        date: ds, dayNum: d, dayName: dn,
        occurrence: weekdayCount[dn],
        label: weekdayCount[dn] + ' - ' + dn
      });
    }
    return days;
  }

  /**
   * Build a lookup: label -> day object, for matching across months.
   */
  function buildLabelMap(days) {
    var map = {};
    for (var i = 0; i < days.length; i++) map[days[i].label] = days[i];
    return map;
  }

  function getMonths(allDates) {
    if (!allDates.length) return null;
    var sorted = allDates.slice().sort(), latest = sorted[sorted.length - 1];
    var p = latest.substring(0, 10).split('-'), cy = parseInt(p[0]), cm = parseInt(p[1]);
    var py = cm === 1 ? cy - 1 : cy, pm = cm === 1 ? 12 : cm - 1;
    return {
      cur: { year: cy, month: cm, name: MONTH_NAMES[cm] + ' ' + cy },
      prev: { year: py, month: pm, name: MONTH_NAMES[pm] + ' ' + py }
    };
  }

  /* ========== STATE ========== */
  var _compData = null, _compView = 'cards', _compDayIdx = -1;
  var _dateMap = {}, _months = null, _curDays = [], _prevDays = [];
  var _prevLabelMap = {};

  function initState() {
    _dateMap = {};
    if (!_compData) return;
    for (var i = 0; i < _compData.length; i++)
      _dateMap[String(_compData[i].report_date).substring(0, 10)] = _compData[i];
    var allDates = Object.keys(_dateMap).sort();
    _months = getMonths(allDates);
    if (!_months) return;
    _curDays = buildLabeledDays(_dateMap, _months.cur.year, _months.cur.month);
    _prevDays = buildLabeledDays(_dateMap, _months.prev.year, _months.prev.month);
    _prevLabelMap = buildLabelMap(_prevDays);
    if (_compDayIdx < 0) _compDayIdx = _curDays.length - 1;
    if (_compDayIdx >= _curDays.length) _compDayIdx = _curDays.length - 1;
    if (_compDayIdx < 0) _compDayIdx = 0;
  }

  /* ========== CARD VIEW ========== */
  function renderCards() {
    var curDay = _curDays[_compDayIdx] || null;
    if (!curDay) return '<div style="text-align:center;padding:60px;color:#64748B;">No data for this day.</div>';

    // Match by weekday label: find the same "1 - Mon" in previous month
    var prevDay = _prevLabelMap[curDay.label] || null;

    var cur = _dateMap[curDay.date] || null;
    var prev = prevDay ? _dateMap[prevDay.date] : null;

    if (!prev && !cur) return '<div style="text-align:center;padding:60px;color:#64748B;">No data for this day.</div>';

    var buckets = [
      { name: 'Regular (FTOD)', color: '#059669', pD: prev ? prev.regular_demand : 0, pC: prev ? prev.regular_collection : 0, cD: cur ? cur.regular_demand : 0, cC: cur ? cur.regular_collection : 0 },
      { name: 'SMA-0 (1-30)', color: '#34D399', pD: prev ? prev.demand_1_30 : 0, pC: prev ? prev.collection_1_30 : 0, cD: cur ? cur.demand_1_30 : 0, cC: cur ? cur.collection_1_30 : 0 },
      { name: 'SMA-1 (31-60)', color: '#FBBF24', pD: prev ? prev.demand_31_60 : 0, pC: prev ? prev.collection_31_60 : 0, cD: cur ? cur.demand_31_60 : 0, cC: cur ? cur.collection_31_60 : 0 },
      { name: 'Pre-NPA', color: '#FB923C', pD: prev ? prev.pnpa_demand : 0, pC: prev ? prev.pnpa_collection : 0, cD: cur ? cur.pnpa_demand : 0, cC: cur ? cur.pnpa_collection : 0 },
      { name: 'NPA', color: '#F87171', pD: prev ? prev.npa_cases : 0, pC: prev ? prev.npa_act_acc : 0, cD: cur ? cur.npa_cases : 0, cC: cur ? cur.npa_act_acc : 0 }
    ];

    var html = '<style>';
    html += '.comp-card{transition:box-shadow .2s;}.comp-card:hover{box-shadow:0 4px 16px rgba(0,0,0,0.08);}';
    html += '.comp-card:hover .comp-dc{max-height:20px;opacity:1;margin-top:3px;}';
    html += '.comp-dc{max-height:0;opacity:0;overflow:hidden;transition:all .25s;font-size:10px;color:#64748B;}';
    html += '</style>';

    // Wrapper with vertical colored bands
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
    if (prevDay) html += '<div style="font-size:10px;font-weight:500;color:#A78BFA;margin-top:1px;">' + fmtDate(prevDay.date) + ' (' + prevDay.label + ')</div>';
    else html += '<div style="font-size:10px;color:#A78BFA;margin-top:1px;">No matching day</div>';
    html += '</div>';
    html += '<div style="flex:1;text-align:center;padding:10px 0;font-size:12px;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:1.5px;background:rgba(16,185,129,0.06);">';
    html += _months.cur.name + '<div style="font-size:10px;font-weight:500;color:#34D399;margin-top:1px;">' + fmtDate(curDay.date) + ' (' + curDay.label + ')</div>';
    html += '</div>';
    html += '<div style="width:140px;text-align:center;padding:12px 0;font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:1px;">Balance</div>';
    html += '</div>';

    // Bucket rows
    for (var i = 0; i < buckets.length; i++) {
      var b = buckets[i];
      var pBal = b.pD - b.pC, cBal = b.cD - b.cC;
      var diff = cBal - pBal;
      var dColor = diff <= 0 ? '#059669' : '#EF4444';
      var dBg = diff <= 0 ? '#F0FDF4' : '#FEF2F2';
      var dArrow = diff < 0 ? '&#9660; ' : diff > 0 ? '&#9650; ' : '';
      var borderTop = i > 0 ? 'border-top:1px solid #F1F5F9;' : '';

      html += '<div class="comp-card" style="display:flex;align-items:stretch;position:relative;z-index:1;' + borderTop + '">';

      // Bucket name
      html += '<div style="min-width:180px;padding:20px 20px;display:flex;align-items:center;gap:10px;">';
      html += '<div style="width:4px;height:32px;border-radius:4px;background:' + b.color + ';flex-shrink:0;"></div>';
      html += '<div style="font-size:15px;font-weight:700;color:#1E293B;">' + b.name + '</div>';
      html += '</div>';

      // Prev month value
      html += '<div style="flex:1;padding:16px 12px;text-align:center;display:flex;flex-direction:column;justify-content:center;">';
      html += '<div style="font-family:\'Playfair Display\',serif;font-size:24px;font-weight:700;color:#FB923C;">' + fmtNum(pBal) + '</div>';
      html += '<div style="font-size:12px;font-weight:600;color:' + pctColor(b.pD, b.pC) + ';margin-top:2px;">' + pctStr(b.pD, b.pC) + '</div>';
      html += '<div class="comp-dc">D: ' + fmtNum(b.pD) + ' &middot; C: ' + fmtNum(b.pC) + '</div>';
      html += '</div>';

      // Cur month value
      html += '<div style="flex:1;padding:16px 12px;text-align:center;display:flex;flex-direction:column;justify-content:center;">';
      html += '<div style="font-family:\'Playfair Display\',serif;font-size:24px;font-weight:700;color:#FB923C;">' + fmtNum(cBal) + '</div>';
      html += '<div style="font-size:12px;font-weight:600;color:' + pctColor(b.cD, b.cC) + ';margin-top:2px;">' + pctStr(b.cD, b.cC) + '</div>';
      html += '<div class="comp-dc">D: ' + fmtNum(b.cD) + ' &middot; C: ' + fmtNum(b.cC) + '</div>';
      html += '</div>';

      // Balance
      html += '<div style="width:140px;padding:16px 12px;text-align:center;display:flex;align-items:center;justify-content:center;">';
      html += '<div style="background:' + dBg + ';border-radius:10px;padding:8px 14px;">';
      html += '<div style="font-size:17px;font-weight:800;color:' + dColor + ';">' + dArrow + fmtNum(Math.abs(diff)) + '</div>';
      html += '<div style="font-size:9px;text-transform:uppercase;letter-spacing:1px;color:' + dColor + ';margin-top:1px;">' + (diff <= 0 ? 'Improved' : 'Higher') + '</div>';
      html += '</div></div>';

      html += '</div>';
    }

    html += '</div>';
    return html;
  }

  /* ========== TABLE VIEW ========== */
  function renderTable() {
    // Build ordered labels: Mon→Tue→Wed→Thu→Fri→Sat→Sun for each occurrence (1,2,3...)
    var DOW_ORDER = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    var prevMap = buildLabelMap(_prevDays);
    var curMap = buildLabelMap(_curDays);

    // Find max occurrence across both months
    var maxOcc = 0;
    for (var i = 0; i < _prevDays.length; i++) if (_prevDays[i].occurrence > maxOcc) maxOcc = _prevDays[i].occurrence;
    for (var i = 0; i < _curDays.length; i++) if (_curDays[i].occurrence > maxOcc) maxOcc = _curDays[i].occurrence;

    // Always include all 7 days (Mon-Sun) for each occurrence
    var allLabels = [];
    for (var occ = 1; occ <= maxOcc; occ++) {
      for (var d = 0; d < DOW_ORDER.length; d++) {
        allLabels.push(occ + ' - ' + DOW_ORDER[d]);
      }
    }

    var html = '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E2E8F0;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:#F8FAFC;">';
    html += '<th style="padding:10px 12px;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:11px;" rowspan="2">Day</th>';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#7C3AED;font-weight:700;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.04);" colspan="5">' + _months.prev.name + '</th>';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#059669;font-weight:700;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.04);" colspan="5">' + _months.cur.name + '</th>';
    html += '</tr><tr style="background:#F8FAFC;">';
    var sh = ['Date', 'RD', 'RC', 'FTOD', 'Coll%'];
    for (var h = 0; h < 2; h++) for (var s = 0; s < sh.length; s++) {
      var bl = s === 0 ? (h === 0 ? 'border-left:1px solid #E2E8F0;' : 'border-left:2px solid #CBD5E1;') : '';
      var bg = h === 0 ? 'background:rgba(139,92,246,0.04);' : 'background:rgba(16,185,129,0.04);';
      var align = s === 0 ? 'text-align:center;' : 'text-align:right;';
      html += '<th style="padding:6px 8px;' + align + 'border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;' + bl + bg + '">' + sh[s] + '</th>';
    }
    html += '</tr></thead><tbody>';

    var lastOcc = 0;
    for (var r = 0; r < allLabels.length; r++) {
      var label = allLabels[r];
      var pv = prevMap[label] || null;
      var cu = curMap[label] || null;
      var pvD = pv ? _dateMap[pv.date] : null;
      var cuD = cu ? _dateMap[cu.date] : null;
      var bg = r % 2 === 0 ? '' : 'background:#FAFAFA;';

      // Week separator when occurrence number increases
      var occ = parseInt(label);
      if (occ > lastOcc && lastOcc > 0) {
        html += '<tr><td colspan="11" style="height:3px;background:#E2E8F0;padding:0;"></td></tr>';
      }
      lastOcc = occ;

      html += '<tr style="' + bg + '">';
      html += '<td style="padding:8px 12px;font-weight:600;color:#1E293B;border-bottom:1px solid #F1F5F9;white-space:nowrap;">' + label + '</td>';
      // Prev month: date + data
      html += '<td style="padding:8px 6px;text-align:center;font-weight:500;color:#7C3AED;border-bottom:1px solid #F1F5F9;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.02);font-size:12px;">' + (pv ? pv.dayNum : '-') + '</td>';
      html += tCells(pvD, '', 'background:rgba(139,92,246,0.02);');
      // Cur month: date + data
      html += '<td style="padding:8px 6px;text-align:center;font-weight:500;color:#059669;border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.02);font-size:12px;">' + (cu ? cu.dayNum : '-') + '</td>';
      html += tCells(cuD, '', 'background:rgba(16,185,129,0.02);');
      html += '</tr>';
    }
    html += '</tbody></table></div>';
    return html;
  }

  function tCells(d, fb, bg) {
    if (!d) { var e = ''; for (var i = 0; i < 4; i++) e += '<td style="padding:8px;text-align:right;color:#CBD5E1;border-bottom:1px solid #F1F5F9;' + (i === 0 ? fb : '') + bg + '">-</td>'; return e; }
    var rd = Number(d.regular_demand)||0, rc = Number(d.regular_collection)||0;
    var ftod = rd - rc;
    var pc = pctColor(rd, rc), s = 'font-variant-numeric:tabular-nums;';
    var h = '';
    h += '<td style="padding:8px;text-align:right;font-weight:500;color:#1E293B;border-bottom:1px solid #F1F5F9;' + s + fb + bg + '">' + fmtNum(rd) + '</td>';
    h += '<td style="padding:8px;text-align:right;font-weight:500;color:#059669;border-bottom:1px solid #F1F5F9;' + s + bg + '">' + fmtNum(rc) + '</td>';
    h += '<td style="padding:8px;text-align:right;font-weight:600;color:#FB923C;border-bottom:1px solid #F1F5F9;' + s + bg + '">' + fmtNum(ftod) + '</td>';
    h += '<td style="padding:8px;text-align:right;font-weight:700;color:' + pc + ';border-bottom:1px solid #F1F5F9;' + bg + '">' + pctStr(rd, rc) + '</td>';
    return h;
  }

  /* ========== MAIN RENDER ========== */
  function render() {
    var container = document.getElementById('comparisonContent');
    if (!_compData || !_compData.length || !_months) {
      container.innerHTML = '<div style="text-align:center;padding:80px 20px;color:#64748B;">No daily data available.</div>';
      return;
    }
    var isCards = _compView === 'cards';
    var curDay = _curDays[_compDayIdx] || null;
    var prevDay = curDay ? (_prevLabelMap[curDay.label] || null) : null;

    var html = '<div style="max-width:1400px;margin:0 auto;padding:16px;">';

    // Header
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px;">';
    html += '<div><h2 style="font-size:20px;font-weight:700;color:#1E293B;margin:0;">Monthly Comparison</h2>';
    html += '<p style="font-size:13px;color:#64748B;margin:4px 0 0;">' + _months.prev.name + ' vs ' + _months.cur.name + ' &mdash; Day wise</p></div>';

    // Date nav — shows weekday label
    html += '<div style="display:flex;align-items:center;gap:6px;">';
    html += '<button onclick="window._compNav(-1)" style="width:34px;height:34px;border:1px solid #E2E8F0;border-radius:8px;background:#fff;cursor:pointer;font-size:16px;color:#059669;font-weight:700;">&larr;</button>';
    html += '<div style="background:#F1F5F9;border:1px solid #E2E8F0;padding:8px 18px;border-radius:10px;text-align:center;min-width:200px;">';
    if (curDay) {
      html += '<div style="font-size:15px;font-weight:700;color:#1E293B;">' + curDay.label + '</div>';
      html += '<div style="font-size:11px;color:#94A3B8;margin-top:2px;">';
      html += (prevDay ? fmtDate(prevDay.date) : 'No data') + '  vs  ' + fmtDate(curDay.date);
      html += '</div>';
    }
    html += '</div>';
    html += '<button onclick="window._compNav(1)" style="width:34px;height:34px;border:1px solid #E2E8F0;border-radius:8px;background:#fff;cursor:pointer;font-size:16px;color:#059669;font-weight:700;">&rarr;</button>';
    html += '</div>';

    // Toggle
    html += '<div style="display:inline-flex;border:1px solid #E2E8F0;border-radius:8px;overflow:hidden;">';
    html += '<button onclick="window._setCompView(\'cards\')" style="padding:6px 14px;font-size:12px;font-weight:' + (isCards ? '600' : '500') + ';color:' + (isCards ? '#059669' : '#64748B') + ';background:' + (isCards ? '#F0FDF4' : '#fff') + ';border:none;cursor:pointer;font-family:inherit;">Cards</button>';
    html += '<button onclick="window._setCompView(\'table\')" style="padding:6px 14px;font-size:12px;font-weight:' + (!isCards ? '600' : '500') + ';color:' + (!isCards ? '#059669' : '#64748B') + ';background:' + (!isCards ? '#F0FDF4' : '#fff') + ';border:none;border-left:1px solid #E2E8F0;cursor:pointer;font-family:inherit;">Table</button>';
    html += '</div></div>';

    html += isCards ? renderCards() : renderTable();
    html += '</div>';
    container.innerHTML = html;
  }

  window._compNav = function (dir) {
    var n = _compDayIdx + dir;
    if (n >= 0 && n < _curDays.length) { _compDayIdx = n; render(); }
  };
  window._setCompView = function (v) { _compView = v; render(); };

  window._loadComparisonTab = function () {
    var c = document.getElementById('comparisonContent');
    c.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading...</div></div>';
    // Build role-based filter params (same as daily API)
    var session = typeof getEmployeeSession === 'function' ? getEmployeeSession() : {};
    var params = [];
    var r = session.role;
    if ((r === 'RM' || r === 'SM') && session.location) params.push('region=' + encodeURIComponent(session.location));
    else if ((r === 'DM' || r === 'DvM' || r === 'AM') && session.location) params.push('district=' + encodeURIComponent(session.location));
    else if (r === 'BM' && session.location) params.push('branch=' + encodeURIComponent(session.location));
    else if (r === 'FO' && session.id) params.push('emp_id=' + encodeURIComponent(session.id));
    var url = '/api/comparison' + (params.length ? '?' + params.join('&') : '');

    fetch(url).then(function(r){return r.json();}).then(function(data){
      _compData = data; _compDayIdx = -1; initState(); render();
    }).catch(function(){ c.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load data.</div>'; });
  };
})();
