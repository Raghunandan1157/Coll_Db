/**
 * comparison.js — Month-over-month comparison tab.
 * Compares by weekday occurrence: 1-Mon vs 1-Mon, 2-Tue vs 2-Tue, etc.
 */
(function () {

  var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function ordinal(n) {
    var s = ['th','st','nd','rd'], v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  }
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
   * Pure math: which occurrence of its weekday is this date?
   * e.g. March 21 2026 (Sat) → 3  (3rd Saturday of March)
   * No loops, no counters — just calendar arithmetic.
   */
  function getOccurrence(year, month, dayOfMonth) {
    var dow = new Date(year, month - 1, dayOfMonth).getDay();
    var firstDow = new Date(year, month - 1, 1).getDay();
    var firstOfThisDow = 1 + ((dow - firstDow + 7) % 7);
    return Math.floor((dayOfMonth - firstOfThisDow) / 7) + 1;
  }

  /**
   * Reverse: given a weekday name and occurrence, return the day-of-month.
   * e.g. getDateForLabel(2026, 3, 'Sat', 1) → 7  (1st Sat of March 2026)
   * Returns null if that occurrence doesn't exist in the month.
   */
  function getDateForLabel(year, month, dayName, occurrence) {
    var dowIndex = DAY_NAMES.indexOf(dayName);
    var firstDow = new Date(year, month - 1, 1).getDay();
    var firstOfThisDow = 1 + ((dowIndex - firstDow + 7) % 7);
    var d = firstOfThisDow + (occurrence - 1) * 7;
    return d <= new Date(year, month, 0).getDate() ? d : null;
  }

  /**
   * Build labeled days for a month. Only includes days with data in dateMap,
   * but occurrence is always from the real calendar (never from data order).
   */
  function buildLabeledDays(dateMap, year, month) {
    var days = [], daysInMonth = new Date(year, month, 0).getDate();
    for (var d = 1; d <= daysInMonth; d++) {
      var mm = String(month).padStart(2, '0'), dd = String(d).padStart(2, '0');
      var ds = year + '-' + mm + '-' + dd;
      if (!dateMap[ds]) continue;
      var dow = new Date(year, month - 1, d).getDay();
      var dn = DAY_NAMES[dow];
      var occ = getOccurrence(year, month, d);
      days.push({
        date: ds, dayNum: d, dayName: dn,
        occurrence: occ,
        label: occ + ' - ' + dn
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
  var _dateMap = {}, _dailyMap = {}, _months = null, _curDays = [], _prevDays = [];
  var _prevLabelMap = {};
  var _prevCumMap = {}, _curCumMap = {};
  var _sortCol = 'curDate', _sortAsc = true; // sorting state: 'day', 'prevDate', 'curDate'

  var DELTA_FIELDS = ['regular_demand','regular_collection','demand_1_30','collection_1_30',
    'demand_31_60','collection_31_60','pnpa_demand','pnpa_collection','npa_cases','npa_act_acc','npa_act_amt'];

  /**
   * Convert cumulative data into daily contributions.
   * For each month, sort dates chronologically, then each day's value =
   * cumulative[day] - cumulative[previous day]. First day = itself.
   */
  function buildDailyMap() {
    _dailyMap = {};
    if (!_months) return;
    var monthList = [_months.prev, _months.cur];
    for (var m = 0; m < monthList.length; m++) {
      var mo = monthList[m];
      var dates = [];
      for (var ds in _dateMap) {
        var p = ds.split('-');
        if (parseInt(p[0]) === mo.year && parseInt(p[1]) === mo.month) dates.push(ds);
      }
      dates.sort();
      for (var i = 0; i < dates.length; i++) {
        var cur = _dateMap[dates[i]];
        var prev = i > 0 ? _dateMap[dates[i - 1]] : null;
        var daily = { report_date: cur.report_date };
        for (var f = 0; f < DELTA_FIELDS.length; f++) {
          var key = DELTA_FIELDS[f];
          daily[key] = (Number(cur[key]) || 0) - (prev ? (Number(prev[key]) || 0) : 0);
        }
        _dailyMap[dates[i]] = daily;
      }
    }
  }

  /**
   * Pre-compute cumulative running totals for each label (Mon→Sun order)
   * so card view can look up the same values as the table.
   */
  function buildCumulativeMaps() {
    _prevCumMap = {}; _curCumMap = {};
    if (!_months) return;
    var DOW_ORDER = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    var prevMap = buildLabelMap(_prevDays);
    var curMap = buildLabelMap(_curDays);
    var latestCurDayNum = _curDays.length > 0 ? _curDays[_curDays.length - 1].dayNum : 0;
    var maxOcc = Math.max(
      Math.ceil(new Date(_months.prev.year, _months.prev.month, 0).getDate() / 7),
      Math.ceil(new Date(_months.cur.year, _months.cur.month, 0).getDate() / 7)
    );

    // Build all labels
    var allLabels = [];
    for (var occ = 1; occ <= maxOcc; occ++) {
      for (var d = 0; d < DOW_ORDER.length; d++) {
        allLabels.push(occ + ' - ' + DOW_ORDER[d]);
      }
    }

    // Each month accumulates in its OWN chronological date order
    // Sort by PREVIOUS month date for prev accumulation
    var prevSorted = allLabels.slice().sort(function(a, b) {
      var pa = a.split(' - '), pb = b.split(' - ');
      var da = getDateForLabel(_months.prev.year, _months.prev.month, pa[1], parseInt(pa[0]));
      var db = getDateForLabel(_months.prev.year, _months.prev.month, pb[1], parseInt(pb[0]));
      return (da || 99) - (db || 99);
    });
    var pRun = {};
    for (var f = 0; f < DELTA_FIELDS.length; f++) pRun[DELTA_FIELDS[f]] = 0;
    for (var i = 0; i < prevSorted.length; i++) {
      var label = prevSorted[i];
      var parts = label.split(' - ');
      var pv = prevMap[label] || null;
      var pvD = pv ? _dailyMap[pv.date] : null;
      var curDateNum = getDateForLabel(_months.cur.year, _months.cur.month, parts[1], parseInt(parts[0]));
      var isFuture = curDateNum === null || curDateNum > latestCurDayNum;
      if (pvD && !isFuture) {
        for (var f = 0; f < DELTA_FIELDS.length; f++) pRun[DELTA_FIELDS[f]] += Number(pvD[DELTA_FIELDS[f]]) || 0;
        var pc = {}; for (var f = 0; f < DELTA_FIELDS.length; f++) pc[DELTA_FIELDS[f]] = pRun[DELTA_FIELDS[f]];
        _prevCumMap[label] = pc;
      }
    }

    // Sort by CURRENT month date for cur accumulation
    var curSorted = allLabels.slice().sort(function(a, b) {
      var pa = a.split(' - '), pb = b.split(' - ');
      var da = getDateForLabel(_months.cur.year, _months.cur.month, pa[1], parseInt(pa[0]));
      var db = getDateForLabel(_months.cur.year, _months.cur.month, pb[1], parseInt(pb[0]));
      return (da || 99) - (db || 99);
    });
    var cRun = {};
    for (var f = 0; f < DELTA_FIELDS.length; f++) cRun[DELTA_FIELDS[f]] = 0;
    for (var i = 0; i < curSorted.length; i++) {
      var label = curSorted[i];
      var cu = curMap[label] || null;
      var cuD = cu ? _dailyMap[cu.date] : null;
      if (cuD) {
        for (var f = 0; f < DELTA_FIELDS.length; f++) cRun[DELTA_FIELDS[f]] += Number(cuD[DELTA_FIELDS[f]]) || 0;
        var cc = {}; for (var f = 0; f < DELTA_FIELDS.length; f++) cc[DELTA_FIELDS[f]] = cRun[DELTA_FIELDS[f]];
        _curCumMap[label] = cc;
      }
    }
  }

  function initState() {
    _dateMap = {};
    if (!_compData) return;
    for (var i = 0; i < _compData.length; i++)
      _dateMap[String(_compData[i].report_date).substring(0, 10)] = _compData[i];
    var allDates = Object.keys(_dateMap).sort();
    _months = getMonths(allDates);
    if (!_months) return;
    buildDailyMap();
    _curDays = buildLabeledDays(_dateMap, _months.cur.year, _months.cur.month);
    _prevDays = buildLabeledDays(_dateMap, _months.prev.year, _months.prev.month);
    _prevLabelMap = buildLabelMap(_prevDays);
    buildCumulativeMaps();
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

    var cur = _curCumMap[curDay.label] || null;
    var prev = _prevCumMap[curDay.label] || null;

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

    // Max weeks from the calendar itself (28 days = 4, 29-31 days = 5)
    var maxOcc = Math.max(
      Math.ceil(new Date(_months.prev.year, _months.prev.month, 0).getDate() / 7),
      Math.ceil(new Date(_months.cur.year, _months.cur.month, 0).getDate() / 7)
    );

    // Latest day with data in current month — anything after this is "future"
    var latestCurDayNum = _curDays.length > 0 ? _curDays[_curDays.length - 1].dayNum : 0;

    // Always include all 7 days (Mon-Sun) for each occurrence
    var allLabels = [];
    for (var occ = 1; occ <= maxOcc; occ++) {
      for (var d = 0; d < DOW_ORDER.length; d++) {
        allLabels.push(occ + ' - ' + DOW_ORDER[d]);
      }
    }

    // Sort allLabels if a sort column is active
    if (_sortCol) {
      var sortData = allLabels.map(function(label) {
        var parts = label.split(' - ');
        var occNum = parseInt(parts[0]), dayName = parts[1];
        var prevDateNum = getDateForLabel(_months.prev.year, _months.prev.month, dayName, occNum);
        var curDateNum = getDateForLabel(_months.cur.year, _months.cur.month, dayName, occNum);
        var dayIdx = DOW_ORDER.indexOf(dayName);
        return { label: label, occ: occNum, dayIdx: dayIdx, prevDateNum: prevDateNum || 99, curDateNum: curDateNum || 99 };
      });
      sortData.sort(function(a, b) {
        var va, vb;
        if (_sortCol === 'day') { va = a.occ * 10 + a.dayIdx; vb = b.occ * 10 + b.dayIdx; }
        else if (_sortCol === 'prevDate') { va = a.prevDateNum; vb = b.prevDateNum; }
        else if (_sortCol === 'curDate') { va = a.curDateNum; vb = b.curDateNum; }
        return _sortAsc ? va - vb : vb - va;
      });
      allLabels = sortData.map(function(d) { return d.label; });
    }

    // Sort arrow helper
    function sortArrow(col) {
      if (_sortCol !== col) return ' <span style="opacity:0.3;font-size:9px;">&#9650;&#9660;</span>';
      return _sortAsc ? ' <span style="font-size:9px;">&#9650;</span>' : ' <span style="font-size:9px;">&#9660;</span>';
    }

    // Sortable header styles
    var sortHdrBase = 'cursor:pointer;user-select:none;transition:background .2s;';
    var sortHdrHover = 'comp-sort-hdr';

    var html = '<style>';
    html += '@keyframes compWaveShimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}';
    html += '.comp-sort-col{background:linear-gradient(90deg,transparent 0%,rgba(99,102,241,0.06) 25%,rgba(99,102,241,0.10) 50%,rgba(99,102,241,0.06) 75%,transparent 100%);background-size:200% 100%;animation:compWaveShimmer 4s ease-in-out infinite;}';
    html += '.comp-sort-hdr{position:relative;}';
    html += '.comp-sort-hdr:hover{background:rgba(99,102,241,0.10)!important;}';
    html += '</style>';

    html += '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E2E8F0;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:#F8FAFC;">';
    html += '<th class="comp-sort-hdr" onclick="window._compSort(\'day\')" style="padding:10px 12px;border-bottom:2px solid #E2E8F0;color:#6366F1;font-size:11px;font-weight:700;' + sortHdrBase + 'background:rgba(99,102,241,0.06);" rowspan="2">Day' + sortArrow('day') + '</th>';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#7C3AED;font-weight:700;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.04);" colspan="5">' + _months.prev.name + '</th>';
    html += '<th style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#059669;font-weight:700;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.04);" colspan="5">' + _months.cur.name + '</th>';
    html += '</tr><tr style="background:#F8FAFC;">';
    var sh = ['Date', 'RD', 'RC', 'FTOD', 'Coll%'];
    for (var h = 0; h < 2; h++) for (var s = 0; s < sh.length; s++) {
      var bl = s === 0 ? (h === 0 ? 'border-left:1px solid #E2E8F0;' : 'border-left:2px solid #CBD5E1;') : '';
      var bg = h === 0 ? 'background:rgba(139,92,246,0.04);' : 'background:rgba(16,185,129,0.04);';
      var align = s === 0 ? 'text-align:center;' : 'text-align:right;';
      if (s === 0) {
        // Date column — sortable + highlighted
        var sortKey = h === 0 ? 'prevDate' : 'curDate';
        html += '<th class="comp-sort-hdr" onclick="window._compSort(\'' + sortKey + '\')" style="padding:6px 8px;' + align + 'border-bottom:2px solid #E2E8F0;color:#6366F1;font-size:10px;font-weight:700;' + sortHdrBase + bl + 'background:rgba(99,102,241,0.06);">' + sh[s] + sortArrow(sortKey) + '</th>';
      } else {
        html += '<th style="padding:6px 8px;' + align + 'border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;' + bl + bg + '">' + sh[s] + '</th>';
      }
    }
    html += '</tr></thead><tbody>';

    // Running cumulative totals — each side accumulates independently
    var prevRD = 0, prevRC = 0, curRD = 0, curRC = 0;
    var isSorted = !!_sortCol;

    var lastOcc = 0;
    for (var r = 0; r < allLabels.length; r++) {
      var label = allLabels[r];
      var pv = prevMap[label] || null;
      var cu = curMap[label] || null;
      var pvD = pv ? _dailyMap[pv.date] : null;
      var cuD = cu ? _dailyMap[cu.date] : null;
      var bg = r % 2 === 0 ? '' : 'background:#FAFAFA;';

      // Week separator when occurrence number increases (only in default order)
      var occ = parseInt(label);
      if (!isSorted && occ > lastOcc && lastOcc > 0) {
        html += '<tr><td colspan="11" style="height:3px;background:#E2E8F0;padding:0;"></td></tr>';
      }
      lastOcc = occ;

      // Compute calendar dates from pure math (always shown, even without data)
      var labelParts = label.split(' - ');
      var occNum = parseInt(labelParts[0]);
      var dayName = labelParts[1];
      var prevDateNum = getDateForLabel(_months.prev.year, _months.prev.month, dayName, occNum);
      var curDateNum = getDateForLabel(_months.cur.year, _months.cur.month, dayName, occNum);

      // Future = current month date hasn't happened yet (after latest data day)
      var isFuture = curDateNum === null || curDateNum > latestCurDayNum;
      var dimStyle = isFuture ? 'opacity:0.3;' : '';

      // Always accumulate inline in display order
      var showPrev = pvD && !isFuture;
      var showCur = !!cuD;
      if (showPrev) { prevRD += Number(pvD.regular_demand) || 0; prevRC += Number(pvD.regular_collection) || 0; }
      if (showCur) { curRD += Number(cuD.regular_demand) || 0; curRC += Number(cuD.regular_collection) || 0; }
      var prevCum = showPrev ? { regular_demand: prevRD, regular_collection: prevRC } : null;
      var curCum = showCur ? { regular_demand: curRD, regular_collection: curRC } : null;
      var prevDisplay = prevCum || (isFuture && pvD ? pvD : null);

      html += '<tr style="' + bg + '">';
      html += '<td class="comp-sort-col" style="padding:8px 12px;font-weight:600;color:#1E293B;border-bottom:1px solid #F1F5F9;white-space:nowrap;">' + label + '</td>';
      // Prev month: date (always from calendar) + data — dimmed if future
      html += '<td class="comp-sort-col" style="padding:8px 6px;text-align:center;font-weight:500;color:#7C3AED;border-bottom:1px solid #F1F5F9;border-left:1px solid #E2E8F0;font-size:12px;white-space:nowrap;' + dimStyle + '">' + (prevDateNum ? ordinal(prevDateNum) + ' ' + MONTH_NAMES[_months.prev.month] : '-') + '</td>';
      html += tCells(prevDisplay, '', 'background:rgba(139,92,246,0.02);' + dimStyle);
      // Cur month: date (always from calendar) + cumulative data
      html += '<td class="comp-sort-col" style="padding:8px 6px;text-align:center;font-weight:500;color:#059669;border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;font-size:12px;white-space:nowrap;">' + (curDateNum ? ordinal(curDateNum) + ' ' + MONTH_NAMES[_months.cur.month] : '-') + '</td>';
      html += tCells(curCum, '', 'background:rgba(16,185,129,0.02);');
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
    var container = document.getElementById('compCollectionBody') || document.getElementById('comparisonContent');
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
    if (!document.getElementById('compCollectionBody')) return;
    var n = _compDayIdx + dir;
    if (n >= 0 && n < _curDays.length) { _compDayIdx = n; render(); }
  };
  window._setCompView = function (v) { _compView = v; render(); };
  window._compSort = function (col) {
    if (_sortCol === col) { _sortAsc = !_sortAsc; }
    else { _sortCol = col; _sortAsc = true; }
    render();
  };

  function renderSubTabBar(active) {
    function btn(key, label, disabled) {
      var isActive = active === key;
      var bg = isActive ? '#0F172A' : (disabled ? '#F1F5F9' : '#fff');
      var color = isActive ? '#fff' : (disabled ? '#CBD5E1' : '#1E293B');
      var cursor = disabled ? 'not-allowed' : 'pointer';
      var onclick = disabled ? '' : ' onclick="window._setCompSub(\'' + key + '\')"';
      return '<button' + onclick + ' style="padding:8px 18px;font-size:13px;font-weight:600;background:' + bg + ';color:' + color + ';border:none;cursor:' + cursor + ';font-family:inherit;">' + label + '</button>';
    }
    var html = '<div id="compSubTabBar" style="display:flex;justify-content:center;padding:14px 16px 0;">';
    html += '<div style="display:inline-flex;border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;background:#fff;box-shadow:0 1px 2px rgba(15,23,42,0.04);">';
    html += btn('collection', 'Collection', false);
    html += '<div style="width:1px;background:#E2E8F0;"></div>';
    html += btn('disbursement', 'Disbursement', false);
    html += '<div style="width:1px;background:#E2E8F0;"></div>';
    html += btn('placeholder', '\u2026', true);
    html += '</div></div>';
    return html;
  }

  function dispatchSub(sub) {
    var c = document.getElementById('comparisonContent');
    if (!c) return;
    c.innerHTML = renderSubTabBar(sub);
    var body = document.createElement('div');
    body.id = sub === 'disbursement' ? 'compDisbBody' : 'compCollectionBody';
    c.appendChild(body);
    if (sub === 'disbursement') {
      if (typeof window.renderDisbComparison === 'function') window.renderDisbComparison();
      else body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Disbursement comparison module not loaded.</div>';
    } else {
      loadCollectionSub(body);
    }
  }

  window._setCompSub = function (sub) {
    try { localStorage.setItem('compSubTab', sub); } catch (e) {}
    dispatchSub(sub);
  };

  function loadCollectionSub(body) {
    body.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading...</div></div>';
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
    }).catch(function(){ body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load data.</div>'; });
  }

  window._loadComparisonTab = function () {
    var sub = 'collection';
    try { sub = localStorage.getItem('compSubTab') || 'collection'; } catch (e) {}
    if (sub !== 'collection' && sub !== 'disbursement') sub = 'collection';
    dispatchSub(sub);
  };
})();
