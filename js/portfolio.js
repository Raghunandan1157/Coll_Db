/**
 * portfolio.js — Portfolio tab renderer (month-wise).
 * Shows: Month selector, product pills, bucket summary table,
 * and hierarchical drill-down (same as collection).
 *
 * Data source: /api/portfolio/* (portfolio_month_wise DB)
 * Container: portfolioContent
 */
(function () {

  /* ========== Session ========== */
  var session = getEmployeeSession();

  /* ========== Structure helper ========== */
  function isNewStructure() { return window.getDataStructure && window.getDataStructure() === 'new'; }

  /* ========== Formatters ========== */
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    var n = typeof v === 'string' ? parseFloat(v) : v;
    if (typeof n !== 'number' || !isFinite(n)) return String(v);
    if (Number.isInteger(n)) return n.toLocaleString('en-IN');
    return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function fmtAmt(v) {
    var n = numVal(v);
    if (n === 0) return '-';
    if (n >= 10000000) return '\u20B9' + (n / 10000000).toFixed(2) + ' Cr';
    if (n >= 100000) return '\u20B9' + (n / 100000).toFixed(2) + ' L';
    if (n >= 1000) return '\u20B9' + (n / 1000).toFixed(2) + ' K';
    return '\u20B9' + n.toLocaleString('en-IN');
  }
  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }

  /* ========== API ========== */
  function apiFetch(url) {
    return fetch(url).then(function (r) {
      if (!r.ok) throw new Error('API ' + r.status);
      return r.json();
    });
  }

  /* ========== State ========== */
  var _pf = window._pf = {
    view: localStorage.getItem('pfView') || 'overall',
    months: [],
    month: localStorage.getItem('pfMonth') || '',
    product: localStorage.getItem('pfProduct') || 'all',
    summary: null,
    children: [],
    pos: null,
    posChildren: []
  };

  /* ========== Query builders ========== */
  function summaryParams() {
    var p = [];
    if (_pf.view !== 'fy' && _pf.month) p.push('month=' + encodeURIComponent(_pf.month));
    if (_pf.product && _pf.product !== 'all') p.push('product_type=' + encodeURIComponent(_pf.product.toUpperCase()));
    if (isNewStructure()) {
      var r = session.role;
      if ((r === 'SM' || r === 'RM') && session.location) p.push('state=' + encodeURIComponent(session.location));
      else if ((r === 'DvM' || r === 'DM') && session.location) p.push('division=' + encodeURIComponent(session.location));
      else if (r === 'AM' && session.location) p.push('area=' + encodeURIComponent(session.location));
      else if (r === 'BM' && session.location) p.push('branch=' + encodeURIComponent(session.location));
      else if ((!r || r === 'FO') && session.id) p.push('emp_id=' + encodeURIComponent(session.id));
    } else {
      if (session.role === 'RM' && session.location) p.push('region=' + encodeURIComponent(session.location));
      else if (session.role === 'DM' && session.location) p.push('district=' + encodeURIComponent(session.location));
      else if (session.role === 'BM' && session.location) p.push('branch=' + encodeURIComponent(session.location));
      else if ((!session.role || session.role === 'FO') && session.id) p.push('emp_id=' + encodeURIComponent(session.id));
    }
    return p.length ? '?' + p.join('&') : '';
  }

  function childrenUrl() {
    var pt = (_pf.product && _pf.product !== 'all') ? _pf.product.toUpperCase() : '';
    var base = [];
    if (_pf.view !== 'fy' && _pf.month) base.push('month=' + encodeURIComponent(_pf.month));
    if (pt) base.push('product_type=' + encodeURIComponent(pt));

    if (isNewStructure()) {
      var v2Base = _pf.view === 'fy' ? '/api/v2/fy' : '/api/v2/portfolio';
      var r = session.role;
      if (r === 'CEO') return v2Base + '/by-state' + (base.length ? '?' + base.join('&') : '');
      if ((r === 'SM' || r === 'RM') && session.location) {
        base.push('state=' + encodeURIComponent(session.location));
        return v2Base + '/by-division?' + base.join('&');
      }
      if ((r === 'DvM' || r === 'DM') && session.location) {
        base.push('division=' + encodeURIComponent(session.location));
        return v2Base + '/by-area?' + base.join('&');
      }
      if (r === 'AM' && session.location) {
        base.push('area=' + encodeURIComponent(session.location));
        return v2Base + '/by-branch?' + base.join('&');
      }
      if (r === 'BM' && session.location) {
        base.push('branch=' + encodeURIComponent(session.location));
        return v2Base + '/by-employee?' + base.join('&');
      }
      return null;
    }

    var apiBase = _pf.view === 'fy' ? '/api/fy' : '/api/portfolio';
    if (session.role === 'CEO') return apiBase + '/by-region' + (base.length ? '?' + base.join('&') : '');
    if (session.role === 'RM' && session.location) {
      base.push('region=' + encodeURIComponent(session.location));
      return apiBase + '/by-district?' + base.join('&');
    }
    if (session.role === 'DM' && session.location) {
      base.push('district=' + encodeURIComponent(session.location));
      return apiBase + '/by-branch?' + base.join('&');
    }
    if (session.role === 'BM' && session.location) {
      base.push('branch=' + encodeURIComponent(session.location));
      return apiBase + '/by-employee?' + base.join('&');
    }
    return null;
  }

  /* ========== Bucket computation ========== */
  function computeBuckets(d) {
    var regAcc = numVal(d.regular_demand);
    var pos = _pf.pos || {};
    var regAmt = numVal(pos.regular_pos) || numVal(d.regular_pos) || numVal(d.regular_demand_amt);
    var sma0Acc = numVal(d.demand_1_30);
    var sma0Amt = numVal(pos.sma0_pos) || numVal(d.sma0_pos) || numVal(d.demand_1_30_amt);
    var sma1Acc = numVal(d.demand_31_60);
    var sma1Amt = numVal(pos.sma1_pos) || numVal(d.sma1_pos) || numVal(d.demand_31_60_amt);
    var sma2Acc = numVal(d.pnpa_demand);
    var sma2Amt = numVal(pos.pnpa_pos) || numVal(d.pnpa_pos) || numVal(d.pnpa_demand_amt);
    var npaAcc = numVal(d.npa_cases);
    var npaAmt = numVal(pos.npa_pos) || numVal(d.npa_pos) || numVal(d.npa_act_amt);

    var totalAcc = regAcc + sma0Acc + sma1Acc + sma2Acc + npaAcc;
    var totalAmt = regAmt + sma0Amt + sma1Amt + sma2Amt + npaAmt;

    function pct(v, t) { return t > 0 ? ((v / t) * 100).toFixed(2) + '%' : '-'; }

    return [
      { bucket: 'Regular',     acc: regAcc,   amt: regAmt,   pctAcc: pct(regAcc, totalAcc),  pctAmt: pct(regAmt, totalAmt),  color: '#10B981' },
      { bucket: 'SMA-0',       acc: sma0Acc,  amt: sma0Amt,  pctAcc: pct(sma0Acc, totalAcc), pctAmt: pct(sma0Amt, totalAmt), color: '#34D399' },
      { bucket: 'SMA-1',       acc: sma1Acc,  amt: sma1Amt,  pctAcc: pct(sma1Acc, totalAcc), pctAmt: pct(sma1Amt, totalAmt), color: '#FBBF24' },
      { bucket: 'SMA-2',       acc: sma2Acc,  amt: sma2Amt,  pctAcc: pct(sma2Acc, totalAcc), pctAmt: pct(sma2Amt, totalAmt), color: '#FB923C' },
      { bucket: 'NPA',         acc: npaAcc,   amt: npaAmt,   pctAcc: pct(npaAcc, totalAcc),  pctAmt: pct(npaAmt, totalAmt),  color: '#F87171' },
      { bucket: 'Grand Total', acc: totalAcc, amt: totalAmt, pctAcc: '100%',                  pctAmt: '100%',                  color: '#6366F1', bold: true }
    ];
  }

  /* ========== Load & Render ========== */
  window.loadPortfolio = function() { loadAndRender(); };
  function loadAndRender() {
    var container = document.getElementById('portfolioContent');
    container.innerHTML = '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#6366F1;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div>' +
      '<div style="color:#64748B;font-size:14px;">Loading portfolio data...</div></div>';

    // Load months first if not yet loaded
    var monthPromise = _pf.months.length ? Promise.resolve() : apiFetch(isNewStructure() ? '/api/v2/portfolio/months' : '/api/portfolio/months').then(function (m) {
      _pf.months = m;
      if (!_pf.month && m.length) {
        // Pick latest month that has data: try from end, check summary
        var candidateIdx = m.length - 1;
        function tryMonth(idx) {
          if (idx < 0) { _pf.month = m[m.length - 1].month_label; localStorage.setItem('pfMonth', _pf.month); return Promise.resolve(); }
          var label = m[idx].month_label;
          return apiFetch((isNewStructure() ? '/api/v2/portfolio' : '/api/portfolio') + '/summary?month=' + encodeURIComponent(label)).then(function(d) {
            var hasData = d && Object.keys(d).some(function(k) { return d[k] !== null && d[k] !== 0 && d[k] !== '0' && d[k] !== '0.00'; });
            if (hasData) { _pf.month = label; localStorage.setItem('pfMonth', label); return; }
            return tryMonth(idx - 1);
          });
        }
        return tryMonth(candidateIdx);
      }
    });

    monthPromise.then(function () {
      var apiBase = isNewStructure()
        ? (_pf.view === 'fy' ? '/api/v2/fy' : '/api/v2/portfolio')
        : (_pf.view === 'fy' ? '/api/fy' : '/api/portfolio');
      var summUrl = apiBase + '/summary' + summaryParams();
      var chUrl = childrenUrl();
      var promises = [apiFetch(summUrl)];
      if (chUrl && session.role && session.role !== 'FO') promises.push(apiFetch(chUrl));
      else promises.push(Promise.resolve([]));

      // Fetch POS data from branch_pos table
      var posParams = [];
      if (isNewStructure()) {
        if ((session.role === 'SM' || session.role === 'RM') && session.location) posParams.push('state=' + encodeURIComponent(session.location));
        else if (session.role === 'DvM' && session.location) posParams.push('division=' + encodeURIComponent(session.location));
        else if (session.role === 'AM' && session.location) posParams.push('area=' + encodeURIComponent(session.location));
        else if (session.role === 'BM' && session.location) posParams.push('branch=' + encodeURIComponent(session.location));
      } else {
        if (session.role === 'RM' && session.location) posParams.push('region=' + encodeURIComponent(session.location));
        else if (session.role === 'DM' && session.location) posParams.push('district=' + encodeURIComponent(session.location));
        else if (session.role === 'BM' && session.location) posParams.push('branch=' + encodeURIComponent(session.location));
      }
      // POS always at company/branch level — no emp_id filter
      var posMonth = _pf.month || (_pf.months.length ? _pf.months[_pf.months.length-1].month_label : '');
      if (_pf.view === 'fy') { if (posMonth) posParams.push('month=' + encodeURIComponent(posMonth)); } else if (posMonth) { posParams.push('month=' + encodeURIComponent(posMonth)); } if (_pf.product && _pf.product !== 'all') posParams.push('product_type=' + encodeURIComponent(_pf.product.toUpperCase())); var posBase = isNewStructure() ? '/api/v2/portfolio/pos-summary' : '/api/portfolio/pos-summary'; var posUrl = posBase + (posParams.length ? '?' + posParams.join('&') : '');
      promises.push(apiFetch(posUrl));

      // POS by child unit
      var posChildUrl = null;
      if (isNewStructure()) {
        var posV2Base = '/api/v2/portfolio';
        var posM2 = _pf.view === 'fy' ? (_pf.months.length ? _pf.months[_pf.months.length-1].month_label : '') : _pf.month;
        var posMParam2 = posM2 ? 'month=' + encodeURIComponent(posM2) : '';
        if (session.role === 'CEO') posChildUrl = posV2Base + '/pos-by-state' + (posMParam2 ? '?' + posMParam2 : '');
        else if ((session.role === 'SM' || session.role === 'RM') && session.location) posChildUrl = posV2Base + '/pos-by-division?' + (posMParam2 ? posMParam2 + '&' : '') + 'state=' + encodeURIComponent(session.location);
        else if (session.role === 'DvM' && session.location) posChildUrl = posV2Base + '/pos-by-area?' + (posMParam2 ? posMParam2 + '&' : '') + 'division=' + encodeURIComponent(session.location);
        else if (session.role === 'AM' && session.location) posChildUrl = posV2Base + '/pos-by-branch?' + (posMParam2 ? posMParam2 + '&' : '') + 'area=' + encodeURIComponent(session.location);
      } else {
        if (session.role === 'CEO') posChildUrl = '/api/portfolio/pos-by-region' + (function(){ var m = _pf.view === 'fy' ? (_pf.months.length ? _pf.months[_pf.months.length-1].month_label : '') : _pf.month; return m ? '?month=' + encodeURIComponent(m) : ''; })();
        else if (session.role === 'RM' && session.location) posChildUrl = '/api/portfolio/pos-by-district?'+ ((function(){ var m = _pf.view === 'fy' ? (_pf.months.length ? _pf.months[_pf.months.length-1].month_label : '') : _pf.month; return m ? 'month=' + encodeURIComponent(m) + '&' : ''; })()) + 'region=' + encodeURIComponent(session.location);
        else if (session.role === 'DM' && session.location) posChildUrl = '/api/portfolio/pos-by-branch?' + ((function(){ var m = _pf.view === 'fy' ? (_pf.months.length ? _pf.months[_pf.months.length-1].month_label : '') : _pf.month; return m ? 'month=' + encodeURIComponent(m) + '&' : ''; })()) + 'district=' + encodeURIComponent(session.location);
      }
      if (posChildUrl) promises.push(apiFetch(posChildUrl));
      else promises.push(Promise.resolve([]));

      return Promise.all(promises);
    }).then(function (res) {
      _pf.summary = res[0];
      _pf.children = res[1] || [];
      _pf.pos = res[2] || {};
      _pf.posChildren = res[3] || [];

      // Fallback chain for FO with no data
      var needsSummaryFallback = false;
      var needsPosFallback = false;
      var s = _pf.summary;
      if ((!session.role || session.role === 'FO') && session.id) {
        var hasData = s && (parseInt(s.regular_demand||0) > 0 || parseInt(s.demand_1_30||0) > 0 || parseInt(s.npa_cases||0) > 0);
        // No summary fallback — keep employee's actual account data (0 is fine)
        // POS always company-level, no fallback needed
      }

      var fallbacks = [];
      var monthParam = _pf.view !== 'fy' && _pf.month ? 'month=' + encodeURIComponent(_pf.month) : '';
      if (needsSummaryFallback) {
        fallbacks.push(apiFetch('/api/portfolio/summary' + (monthParam ? '?' + monthParam : '')).then(function(d) { _pf.summary = d; }));
      }
      if (needsPosFallback) {
        fallbacks.push(apiFetch('/api/portfolio/pos-summary' + (monthParam ? '?' + monthParam : '')).then(function(d) { _pf.pos = d || {}; }));
      }

      if (fallbacks.length > 0) {
        Promise.all(fallbacks).then(function() { render(); }).catch(function() { render(); });
      } else {
        render();
      }
    }).catch(function (err) {
      console.error('Portfolio load failed:', err);
      container.innerHTML = noDataHtml('Failed to load portfolio data.');
    });
  }

  function render() {
    var container = document.getElementById('portfolioContent');
    var d = _pf.summary;

    if (!d || (typeof d === 'object' && Object.keys(d).length === 0)) {
      container.innerHTML = '<div class="desktop-content-area">' + viewToggleHtml() + monthSelectorHtml() + productPillsHtml() + noDataHtml('No portfolio data for this selection.') + '</div>';
      attachHandlers();
      return;
    }

    var html = '';
    html += viewToggleHtml();
    html += monthSelectorHtml();
    html += productPillsHtml();
    html += summaryCardsHtml(d);
    html += bucketTableHtml(d);

    // Sub-units (hierarchical drill-down)
    var children = _pf.children || [];
    if (session.role && session.role !== 'FO' && children.length > 0) {
      var childRoleMap = isNewStructure()
        ? { CEO: 'SM', SM: 'DvM', RM: 'DvM', DvM: 'AM', DM: 'AM', AM: 'BM', BM: 'FO' }
        : { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
      var childRole = childRoleMap[session.role];
      if (childRole) {
        children.sort(function (a, b) {
          var tA = numVal(a.regular_demand) + numVal(a.demand_1_30) + numVal(a.demand_31_60) + numVal(a.pnpa_demand) + numVal(a.npa_cases);
          var tB = numVal(b.regular_demand) + numVal(b.demand_1_30) + numVal(b.demand_31_60) + numVal(b.pnpa_demand) + numVal(b.npa_cases);
          return tB - tA;
        });
        html += subUnitsHtml(children, childRole);
      }
    }

    container.innerHTML = '<div class="desktop-content-area">' + html + '</div>';
    attachHandlers();
  }

  /* ========== HTML Builders ========== */

  function noDataHtml(msg) {
    return '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="font-size:36px;margin-bottom:12px;">&#128202;</div>' +
      '<div style="color:#64748B;font-size:14px;">' + (msg || 'No portfolio data available.') + '</div></div>';
  }

  function viewToggleHtml() {
    var ov = _pf.view === 'overall';
    return '<div class="pf-view-toggle emp-fade">' +
      '<button class="pf-view-btn' + (ov ? ' active' : '') + '" data-pf-view="overall">OverAll</button>' +
      '<button class="pf-view-btn' + (!ov ? ' active' : '') + '" data-pf-view="fy">FY 25-26</button>' +
    '</div>';
  }

  function monthSelectorHtml() { return ""; }

  function productPillsHtml() {
    var pills = [
      { key: 'all', label: 'All' },
      { key: 'igl', label: 'IGL' },
      { key: 'fig', label: 'FIG' },
      { key: 'il',  label: 'IL' }
    ];
    var html = '<div class="emp-fade" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px;">';
    for (var i = 0; i < pills.length; i++) {
      var p = pills[i];
      var active = _pf.product === p.key;
      html += '<button data-pf-product="' + p.key + '" style="' +
        'padding:6px 16px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;border:none;transition:all .2s;' +
        (active ? 'background:#6366F1;color:#fff;' : 'background:#F1F5F9;color:#64748B;') +
      '">' + p.label + '</button>';
    }
    html += '</div>';
    return html;
  }

  function summaryCardsHtml(d) {
    var buckets = computeBuckets(d);
    var grand = buckets[buckets.length - 1];
    return '<div class="emp-fade" style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px;">' +
      '<div style="background:linear-gradient(135deg,#6366F1,#818CF8);border-radius:14px;padding:18px 20px;color:#fff;">' +
        '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">Total Account</div>' +
        '<div style="font-size:24px;font-weight:700;">' + fmtNum(grand.acc) + '</div>' +
      '</div>' +
      '<div style="background:linear-gradient(135deg,#6366F1,#818CF8);border-radius:14px;padding:18px 20px;color:#fff;">' +
        '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">POS (Amount)</div>' +
        '<div style="font-size:24px;font-weight:700;">' + fmtAmt(grand.amt) + '</div>' +
      '</div>' +
    '</div>';
  }

  function bucketTableHtml(d) {
    var buckets = computeBuckets(d);
    var html = '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);overflow:hidden;margin-bottom:16px;">' +
      '<div style="padding:14px 16px;border-bottom:1px solid #F1F5F9;font-weight:700;font-size:14px;color:#1E293B;">Bucket-wise Portfolio</div>' +
      '<div style="overflow-x:auto;">' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
      '<thead><tr style="background:#F8FAFC;">' +
        '<th style="padding:10px 14px;text-align:left;color:#64748B;font-weight:600;white-space:nowrap;">Bucket</th>' +
        '<th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;white-space:nowrap;">Account</th>' +
        '<th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;white-space:nowrap;">% Contrib</th>' +
        '<th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;white-space:nowrap;">POS (Amount)</th>' +
        '<th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;white-space:nowrap;">% Contrib</th>' +
      '</tr></thead><tbody>';

    for (var i = 0; i < buckets.length; i++) {
      var b = buckets[i];
      var isGrand = b.bold;
      var bg = isGrand ? 'background:#F8FAFC;' : '';
      var fw = isGrand ? 'font-weight:700;' : '';
      html += '<tr style="border-bottom:1px solid #F1F5F9;' + bg + '">' +
        '<td style="padding:10px 14px;' + fw + '"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + b.color + ';margin-right:8px;vertical-align:middle;"></span>' + esc(b.bucket) + '</td>' +
        '<td style="padding:10px 14px;text-align:right;' + fw + '">' + fmtNum(b.acc) + '</td>' +
        '<td style="padding:10px 14px;text-align:right;' + fw + 'color:' + b.color + ';">' + b.pctAcc + '</td>' +
        '<td style="padding:10px 14px;text-align:right;' + fw + '">' + fmtAmt(b.amt) + '</td>' +
        '<td style="padding:10px 14px;text-align:right;' + fw + 'color:' + b.color + ';">' + b.pctAmt + '</td>' +
      '</tr>';
    }

    html += '</tbody></table></div></div>';
    return html;
  }

  function getSubunitView() {
    return localStorage.getItem('subunitView') || 'card';
  }

  function subunitToggleHtml() {
    var v = getSubunitView();
    return '<div class="subunit-view-toggle">' +
      '<button class="subunit-view-btn' + (v === 'card' ? ' active' : '') + '" data-subview="card">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>' +
        'Cards</button>' +
      '<button class="subunit-view-btn' + (v === 'table' ? ' active' : '') + '" data-subview="table">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>' +
        'Table</button>' +
    '</div>';
  }

  function getPfChildName(ch, childRole) {
    if (isNewStructure()) {
      if (childRole === 'SM') return ch.state_name || '';
      if (childRole === 'DvM') return ch.division_name || '';
      if (childRole === 'AM') return ch.area_name || '';
      if (childRole === 'BM') return ch.branch_name || '';
      if (childRole === 'FO') return ch.officer_name || ch.name || '';
    } else {
      if (childRole === 'RM') return ch.region_name || '';
      if (childRole === 'DM') return ch.district_name || '';
      if (childRole === 'BM') return ch.branch_name || '';
      if (childRole === 'FO') return ch.officer_name || ch.name || '';
    }
    return '';
  }

  function getPfDataAttr(ch, childRole, childName) {
    if (childRole === 'FO') {
      return 'data-emp-id="' + esc(ch.emp_id || '') + '" data-emp-name="' + esc(childName) + '"';
    }
    return 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(childName) + '"';
  }

  function pfChildData(ch, childRole) {
    var totalAcc = numVal(ch.regular_demand) + numVal(ch.demand_1_30) + numVal(ch.demand_31_60) + numVal(ch.pnpa_demand) + numVal(ch.npa_cases);
    var childKey = isNewStructure()
      ? (ch.state_name || ch.division_name || ch.area_name || ch.branch_name || '')
      : (ch.region_name || ch.district_name || ch.branch_name || '');
    var posChild = (_pf.posChildren || []).find(function(p) {
      return (isNewStructure()
        ? (p.state_name || p.division_name || p.area_name || p.branch_name || '')
        : (p.region_name || p.district_name || p.branch_name || '')) === childKey;
    });
    var totalAmt = posChild ? numVal(posChild.total_pos) : (numVal(ch.total_pos) || (numVal(ch.regular_demand_amt) + numVal(ch.demand_1_30_amt) + numVal(ch.demand_31_60_amt) + numVal(ch.pnpa_demand_amt) + numVal(ch.npa_act_amt)));
    var npaRatio = totalAcc > 0 ? numVal(ch.npa_cases) / totalAcc : 0;
    var pctColor = npaRatio <= 0.02 ? '#34D399' : npaRatio <= 0.05 ? '#FBBF24' : '#F87171';
    return { totalAcc: totalAcc, totalAmt: totalAmt, npa: numVal(ch.npa_cases), pctColor: pctColor };
  }

  function subUnitsHtml(children, childRole) {
    if (!children.length) return '';
    var roleLabel = isNewStructure()
      ? { SM: 'Regions', DvM: 'Divisions', AM: 'Areas', BM: 'Branches', FO: 'Officers' }
      : { RM: 'Regions', DM: 'Districts', BM: 'Branches', FO: 'Officers' };
    var avType = isNewStructure()
      ? { SM: 'state', DvM: 'division', AM: 'area', BM: 'branch', FO: 'officer' }
      : { RM: 'region', DM: 'district', BM: 'branch', FO: 'officer' };
    var label = roleLabel[childRole] || 'Team';
    var av = avType[childRole] || 'branch';
    var isTable = getSubunitView() === 'table';

    var avColors = { state: 'background:#F3E8FF;color:#7C3AED;', region: 'background:#F3E8FF;color:#7C3AED;', division: 'background:#DBEAFE;color:#2563EB;', area: 'background:#ECFDF5;color:#059669;', district: 'background:#DBEAFE;color:#2563EB;', branch: 'background:#ECFDF5;color:#059669;', officer: 'background:#FEF3C7;color:#D97706;' };
    var avStyle = avColors[av] || 'background:#F1F5F9;color:#64748B;';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label + '<span class="emp-team-count">' + children.length + '</span>' + subunitToggleHtml() + '</div>';

    if (isTable) {
      html += '<div class="subunit-table-wrap"><table class="subunit-table">';
      html += '<thead><tr>' +
        '<th>Unit <span class="sort-icon">&#8597;</span></th>' +
        '<th class="num-col">Accounts <span class="sort-icon">&#8597;</span></th>' +
        '<th class="num-col">POS <span class="sort-icon">&#8597;</span></th>' +
        '<th class="num-col">NPA <span class="sort-icon">&#8597;</span></th>' +
        '<th style="width:24px;"></th>' +
      '</tr></thead><tbody>';

      for (var i = 0; i < children.length; i++) {
        var ch = children[i];
        var d = pfChildData(ch, childRole);
        var childName = getPfChildName(ch, childRole);
        var dataAttr = getPfDataAttr(ch, childRole, childName);
        var initial = childName.charAt(0).toUpperCase();

        html += '<tr class="emp-sub-card" ' + dataAttr + '>' +
          '<td class="name-col"><span class="tbl-avatar" style="' + avStyle + '">' + initial + '</span>' + esc(childName) + '</td>' +
          '<td class="num-col">' + fmtNum(d.totalAcc) + '</td>' +
          '<td class="num-col">' + fmtAmt(d.totalAmt) + '</td>' +
          '<td class="num-col"><span class="tbl-pct" style="color:' + d.pctColor + '">' + fmtNum(d.npa) + '</span></td>' +
          '<td class="tbl-arrow">&#8250;</td>' +
        '</tr>';
      }
      html += '</tbody></table></div>';
    } else {
      html += '<div class="desktop-grid-3">';
      for (var i = 0; i < children.length; i++) {
        var ch = children[i];
        var d = pfChildData(ch, childRole);
        var childName = getPfChildName(ch, childRole);
        var dataAttr = getPfDataAttr(ch, childRole, childName);
        var initial = childName.charAt(0).toUpperCase();

        html += '<div class="emp-sub-card desktop-sub-card" ' + dataAttr + '>' +
          '<div class="emp-sub-avatar ' + av + '">' + initial + '</div>' +
          '<div class="emp-sub-info">' +
            '<div class="emp-sub-name">' + esc(childName) + '</div>' +
            '<div class="emp-sub-meta">' +
              '<span>A/c: ' + fmtNum(d.totalAcc) + '</span>' +
              '<span>POS: ' + fmtAmt(d.totalAmt) + '</span>' +
            '</div>' +
          '</div>' +
          '<div class="emp-sub-pct" style="color:' + d.pctColor + '">' + fmtNum(d.npa) + ' NPA</div>' +
          '<div class="emp-sub-arrow">&#8250;</div>' +
        '</div>';
      }
      html += '</div>';
    }

    html += '</div>';
    return html;
  }

  /* ========== Event Handlers ========== */
  var _pfHandlersAttached = false;

  function attachHandlers() {
    var container = document.getElementById('portfolioContent');
    if (!container) return;

    // View toggle
    container.querySelectorAll('[data-pf-view]').forEach(function (btn) {
      btn.onclick = function () {
        _pf.view = btn.dataset.pfView;
        localStorage.setItem('pfView', _pf.view);
        loadAndRender();
      };
    });

    // Month selector
    container.querySelectorAll('[data-pf-month]').forEach(function (btn) {
      btn.onclick = function () {
        _pf.month = btn.dataset.pfMonth;
        localStorage.setItem('pfMonth', _pf.month);
        loadAndRender();
      };
    });

    // Product pills
    container.querySelectorAll('[data-pf-product]').forEach(function (pill) {
      pill.onclick = function () {
        _pf.product = pill.dataset.pfProduct;
        localStorage.setItem('pfProduct', _pf.product);
        loadAndRender();
      };
    });

    // Subunit view toggle (card/table)
    container.querySelectorAll('[data-subview]').forEach(function (btn) {
      btn.onclick = function () {
        localStorage.setItem('subunitView', btn.dataset.subview);
        render();
      };
    });

    // Table column sorting
    container.querySelectorAll('.subunit-table thead th').forEach(function (th, colIdx) {
      th.onclick = function () {
        var table = th.closest('.subunit-table');
        var tbody = table.querySelector('tbody');
        var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
        var asc = th.classList.contains('sorted-asc');
        table.querySelectorAll('th').forEach(function (h) { h.classList.remove('sorted', 'sorted-asc', 'sorted-desc'); });
        th.classList.add('sorted', asc ? 'sorted-desc' : 'sorted-asc');
        rows.sort(function (a, b) {
          var cellA = a.cells[colIdx].textContent.replace(/[,%₹L]/g, '').trim();
          var cellB = b.cells[colIdx].textContent.replace(/[,%₹L]/g, '').trim();
          var numA = parseFloat(cellA), numB = parseFloat(cellB);
          if (!isNaN(numA) && !isNaN(numB)) return asc ? numB - numA : numA - numB;
          return asc ? cellB.localeCompare(cellA) : cellA.localeCompare(cellB);
        });
        rows.forEach(function (r) { tbody.appendChild(r); });
      };
    });

    // Sub-unit drill-down
    if (!_pfHandlersAttached) {
      _pfHandlersAttached = true;
      container.addEventListener('click', function (ev) {
        var card = ev.target.closest('.emp-sub-card');
        if (!card) return;
        var portfolioTab = document.getElementById('portfolioTab');
        if (!portfolioTab || !portfolioTab.classList.contains('active')) return;

        card.style.background = '#EEF2FF';
        card.style.pointerEvents = 'none';
        var arrow = card.querySelector('.emp-sub-arrow');
        if (arrow) arrow.innerHTML = '<div style="width:16px;height:16px;border:2px solid #C7D2FE;border-top-color:#6366F1;border-radius:50%;animation:spin .7s linear infinite;"></div>';

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

  /* ========== Entry point ========== */
  window._loadPortfolioTab = function () {
    try {
      loadAndRender();
    } catch (err) {
      console.error('Portfolio load failed:', err);
      document.getElementById('portfolioContent').innerHTML = noDataHtml('Failed to load portfolio data.');
    }
  };
})();
