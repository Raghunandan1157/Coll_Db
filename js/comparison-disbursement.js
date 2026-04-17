/**
 * comparison-disbursement.js — Month-over-month Disbursement comparison.
 * Day-of-month matching (1-1, 2-2 ...). No weekday column.
 */
(function () {
  var MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function pad2(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
  function numVal(v) { var n = Number(v); return isFinite(n) ? n : 0; }
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    var n = Number(v); return isFinite(n) ? n.toLocaleString('en-IN') : '-';
  }
  function fmtCr(v) {
    var n = Number(v) || 0;
    return (n / 10000000).toFixed(2) + ' Cr';
  }
  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }

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
      map[day] = {
        accounts: numVal(rows[i].total_count),
        amount: numVal(rows[i].total_amount)
      };
    }
    return map;
  }

  function diffPct(prev, cur) {
    if (!prev) return null;
    return ((cur - prev) / prev) * 100;
  }

  function diffCellHtml(prev, cur) {
    if (prev == null && cur == null) return '<td style="padding:8px;text-align:right;color:#CBD5E1;border-bottom:1px solid #F1F5F9;">-</td>';
    if (!prev) {
      return '<td style="padding:8px;text-align:right;font-weight:700;color:#6366F1;border-bottom:1px solid #F1F5F9;">new</td>';
    }
    var d = diffPct(prev, cur || 0);
    var color = d >= 0 ? '#059669' : '#EF4444';
    var arrow = d > 0 ? '&#9650; ' : d < 0 ? '&#9660; ' : '';
    return '<td style="padding:8px;text-align:right;font-weight:700;color:' + color + ';border-bottom:1px solid #F1F5F9;">' + arrow + Math.abs(d).toFixed(1) + '%</td>';
  }

  function renderTable(prevMap, curMap, months) {
    var days = [];
    for (var d = 1; d <= 31; d++) {
      if (prevMap[d] || curMap[d]) days.push(d);
    }

    var tot = { prevAcc: 0, prevAmt: 0, curAcc: 0, curAmt: 0 };
    var html = '<div style="overflow-x:auto;border-radius:12px;border:1px solid #E2E8F0;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:#F8FAFC;">';
    html += '<th rowspan="2" style="padding:10px 12px;border-bottom:2px solid #E2E8F0;color:#6366F1;font-size:11px;font-weight:700;background:rgba(99,102,241,0.06);text-align:center;">Date</th>';
    html += '<th colspan="2" style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#7C3AED;font-weight:700;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.04);">' + months.prev.name + '</th>';
    html += '<th colspan="2" style="padding:10px;text-align:center;border-bottom:1px solid #E2E8F0;color:#059669;font-weight:700;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.04);">' + months.cur.name + '</th>';
    html += '<th rowspan="2" style="padding:10px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:11px;font-weight:700;border-left:2px solid #CBD5E1;background:#F8FAFC;">Diff %</th>';
    html += '</tr><tr style="background:#F8FAFC;">';
    html += '<th style="padding:6px 8px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.04);">Accounts</th>';
    html += '<th style="padding:6px 8px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;background:rgba(139,92,246,0.04);">Amount</th>';
    html += '<th style="padding:6px 8px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.04);">Accounts</th>';
    html += '<th style="padding:6px 8px;text-align:right;border-bottom:2px solid #E2E8F0;color:#64748B;font-size:10px;background:rgba(16,185,129,0.04);">Amount</th>';
    html += '</tr></thead><tbody>';

    if (!days.length) {
      html += '<tr><td colspan="6" style="padding:40px;text-align:center;color:#64748B;">No disbursement data for these months.</td></tr>';
    }

    for (var r = 0; r < days.length; r++) {
      var day = days[r];
      var p = prevMap[day] || null;
      var c = curMap[day] || null;
      var bg = r % 2 === 0 ? '' : 'background:#FAFAFA;';
      if (p) { tot.prevAcc += p.accounts; tot.prevAmt += p.amount; }
      if (c) { tot.curAcc += c.accounts; tot.curAmt += c.amount; }

      html += '<tr style="' + bg + '">';
      html += '<td style="padding:8px 12px;text-align:center;font-weight:700;color:#1E293B;border-bottom:1px solid #F1F5F9;">' + day + '</td>';
      html += '<td style="padding:8px;text-align:right;color:' + (p ? '#1E293B' : '#CBD5E1') + ';border-bottom:1px solid #F1F5F9;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.02);font-variant-numeric:tabular-nums;">' + (p ? fmtNum(p.accounts) : '-') + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:600;color:' + (p ? '#7C3AED' : '#CBD5E1') + ';border-bottom:1px solid #F1F5F9;background:rgba(139,92,246,0.02);font-variant-numeric:tabular-nums;">' + (p ? fmtCr(p.amount) : '-') + '</td>';
      html += '<td style="padding:8px;text-align:right;color:' + (c ? '#1E293B' : '#CBD5E1') + ';border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.02);font-variant-numeric:tabular-nums;">' + (c ? fmtNum(c.accounts) : '-') + '</td>';
      html += '<td style="padding:8px;text-align:right;font-weight:600;color:' + (c ? '#059669' : '#CBD5E1') + ';border-bottom:1px solid #F1F5F9;background:rgba(16,185,129,0.02);font-variant-numeric:tabular-nums;">' + (c ? fmtCr(c.amount) : '-') + '</td>';
      html += diffCellHtml(p ? p.amount : null, c ? c.amount : null).replace('border-bottom:1px solid #F1F5F9;', 'border-bottom:1px solid #F1F5F9;border-left:2px solid #CBD5E1;');
      html += '</tr>';
    }

    if (days.length) {
      html += '<tr style="background:#F1F5F9;border-top:2px solid #CBD5E1;">';
      html += '<td style="padding:10px 12px;text-align:center;font-weight:800;color:#1E293B;">Total</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:700;color:#1E293B;border-left:1px solid #E2E8F0;background:rgba(139,92,246,0.08);font-variant-numeric:tabular-nums;">' + fmtNum(tot.prevAcc) + '</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:800;color:#7C3AED;background:rgba(139,92,246,0.08);font-variant-numeric:tabular-nums;">' + fmtCr(tot.prevAmt) + '</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:700;color:#1E293B;border-left:2px solid #CBD5E1;background:rgba(16,185,129,0.08);font-variant-numeric:tabular-nums;">' + fmtNum(tot.curAcc) + '</td>';
      html += '<td style="padding:10px 8px;text-align:right;font-weight:800;color:#059669;background:rgba(16,185,129,0.08);font-variant-numeric:tabular-nums;">' + fmtCr(tot.curAmt) + '</td>';
      html += diffCellHtml(tot.prevAmt || null, tot.curAmt || null).replace('border-bottom:1px solid #F1F5F9;', 'border-left:2px solid #CBD5E1;font-weight:800;');
      html += '</tr>';
    }

    html += '</tbody></table></div>';
    return html;
  }

  window.renderDisbComparison = function () {
    var container = document.getElementById('comparisonContent');
    if (!container) return;
    // Preserve sub-tab bar if already rendered by dispatcher; mount a body div below it
    var body = document.getElementById('compDisbBody');
    if (!body) {
      var bar = document.getElementById('compSubTabBar');
      body = document.createElement('div');
      body.id = 'compDisbBody';
      if (bar && bar.parentNode === container) container.appendChild(body);
      else { container.innerHTML = ''; container.appendChild(body); }
    }
    body.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading...</div></div>';

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

        var html = '<div style="max-width:1400px;margin:0 auto;padding:16px;">';
        html += '<div style="margin-bottom:20px;">';
        html += '<h2 style="font-size:20px;font-weight:700;color:#1E293B;margin:0;">Disbursement Comparison</h2>';
        html += '<p style="font-size:13px;color:#64748B;margin:4px 0 0;">' + months.prev.name + ' vs ' + months.cur.name + ' &mdash; Day wise (date-matched)</p>';
        html += '</div>';
        html += renderTable(prevMap, curMap, months);
        html += '</div>';
        body.innerHTML = html;
      }).catch(function () {
        body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load disbursement data.</div>';
      });
    }).catch(function () {
      body.innerHTML = '<div style="text-align:center;padding:80px;color:#64748B;">Failed to load disbursement dates.</div>';
    });
  };
})();
