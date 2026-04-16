/**
 * collection.js — Collection tab renderer for the employee dashboard.
 * Mirrors portfolio.js exactly: product pills, on-date preview,
 * DPD buckets, NPA details, and drill-down.
 *
 * Data source: PostgreSQL REST API
 * Container: collectionContent
 */
(function () {

  /* ========== Session ========== */
  var session = getEmployeeSession();

  /* ========== Structure helper ========== */
  function isNewStructure() { return window.getDataStructure && window.getDataStructure() === 'new'; }

  /* ========== Employee name map from master ========== */
  var _empNames = window._empNameMap || {};
  if (window._empNameMapPromise) {
    window._empNameMapPromise.then(function(m) { _empNames = m || {}; });
  } else if (!window._empNameMap) {
    fetch('/api/employees/names').then(function(r) { return r.json(); }).then(function(m) { _empNames = m || {}; window._empNameMap = m; }).catch(function(e) { console.warn('Failed to load employee names:', e); });
  }
  function getFullName(empId, fallback) {
    var mapped = _empNames[empId];
    if (mapped && mapped.trim()) return mapped;
    return fallback || empId || '';
  }

  /* ========== Formatters ========== */
  function fmtNum(v) {
    if (v == null || v === '' || v === '-') return '-';
    if (typeof v === 'string') {
      var parsed = parseFloat(v);
      if (!isNaN(parsed)) {
        if (Number.isInteger(parsed)) return parsed.toLocaleString('en-IN');
        return parsed.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      return v.trim() || '-';
    }
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return v.toLocaleString('en-IN');
      return v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(v);
  }
  function fmtPct(v) {
    if (v == null || v === '' || v === '-') return '-';
    if (typeof v === 'number') return (v * 100).toFixed(2) + '%';
    return String(v);
  }
  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }

  /* ========== API helper ========== */
  function apiFetch(url) {
    return fetch(url).then(function (resp) {
      if (!resp.ok) throw new Error('API error: ' + resp.status);
      return resp.json();
    });
  }

  /* ========== Build query params for summary endpoint ========== */
  function summaryParams(productType) {
    var params = [];
    if (!isNewStructure()) {
      params.push('structure=old');
    }
    if (productType && productType !== 'all') {
      params.push('product_type=' + encodeURIComponent(productType.toUpperCase()));
    }
    // When a date is selected, we use /api/daily which needs old-structure params
    if (isNewStructure() && !_collState.date) {
      var r = session.role;
      if ((r === 'SM' || r === 'RM') && session.location) {
        params.push('state=' + encodeURIComponent(session.location));
      } else if ((r === 'DvM' || r === 'DM') && session.location) {
        params.push('division=' + encodeURIComponent(session.location));
      } else if (r === 'AM' && session.location) {
        params.push('area=' + encodeURIComponent(session.location));
      } else if (r === 'BM' && session.location) {
        params.push('branch=' + encodeURIComponent(session.location));
      } else if ((!r || r === 'FO') && session.id) {
        params.push('emp_id=' + encodeURIComponent(session.id));
      }
    } else {
      // Old structure params (also used when date is selected in new structure)
      var r = session.role;
      if ((r === 'RM' || r === 'SM') && session.location) {
        params.push('region=' + encodeURIComponent(session.location));
      } else if ((r === 'DM' || r === 'DvM') && session.location) {
        params.push('district=' + encodeURIComponent(session.location));
      } else if (r === 'AM' && session.location) {
        // AM maps to district level in daily API
        params.push('district=' + encodeURIComponent(session.location));
      } else if (r === 'BM' && session.location) {
        params.push('branch=' + encodeURIComponent(session.location));
      } else if ((!r || r === 'FO') && session.id) {
        params.push('emp_id=' + encodeURIComponent(session.id));
      }
    }
    return params.length ? '?' + params.join('&') : '';
  }

  /* ========== Build query params for sub-units endpoint ========== */
  function subUnitsUrl(productType) {
    var pt = (productType && productType !== 'all') ? productType.toUpperCase() : '';
    var ptParam = pt ? 'product_type=' + encodeURIComponent(pt) : '';
    var structParam = isNewStructure() ? '' : 'structure=old&';
    var dateParam = _collState.date ? 'date=' + encodeURIComponent(_collState.date) + '&scope=' + (_collState.view === 'fy' ? 'fy' : 'oa') + '&' + structParam : '';

    // When a date is selected, always use /api/daily (works for both structures)
    // Daily API only supports old-structure filters: region, district, branch, emp_id
    if (_collState.date) {
      var dailyBase = '/api/daily';
      var r = session.role;
      if (r === 'CEO') {
        return dailyBase + '/by-region?' + dateParam + (ptParam || '');
      }
      if ((r === 'SM' || r === 'RM') && session.location) {
        var rp = [ptParam, 'region=' + encodeURIComponent(session.location)].filter(Boolean);
        return dailyBase + '/by-district?' + dateParam + rp.join('&');
      }
      if ((r === 'DvM' || r === 'DM' || r === 'AM') && session.location) {
        // DvM/DM/AM all map to district level in daily API
        var dp = [ptParam, 'district=' + encodeURIComponent(session.location)].filter(Boolean);
        return dailyBase + '/by-branch?' + dateParam + dp.join('&');
      }
      if (r === 'BM' && session.location) {
        var bp2 = [ptParam, 'branch=' + encodeURIComponent(session.location)].filter(Boolean);
        return dailyBase + '/by-employee?' + dateParam + bp2.join('&');
      }
      return dailyBase + '/by-region?' + dateParam + (ptParam || '');
    }

    if (isNewStructure()) {
      var v2Base = (_collState.view === 'fy') ? '/api/v2/fy' : '/api/v2/collection';
      var r = session.role;
      if (r === 'CEO') {
        var parts = [];
        if (ptParam) parts.push(ptParam);
        return v2Base + '/by-state' + (parts.length ? '?' + parts.join('&') : '');
      }
      if ((r === 'SM' || r === 'RM') && session.location) {
        var p = [ptParam, 'state=' + encodeURIComponent(session.location)].filter(Boolean);
        return v2Base + '/by-division?' + p.join('&');
      }
      if ((r === 'DvM' || r === 'DM') && session.location) {
        var p2 = [ptParam, 'division=' + encodeURIComponent(session.location)].filter(Boolean);
        return v2Base + '/by-area?' + p2.join('&');
      }
      if (r === 'AM' && session.location) {
        var p3 = [ptParam, 'area=' + encodeURIComponent(session.location)].filter(Boolean);
        return v2Base + '/by-branch?' + p3.join('&');
      }
      if (r === 'BM' && session.location) {
        var p4 = [ptParam, 'branch=' + encodeURIComponent(session.location)].filter(Boolean);
        return v2Base + '/by-employee?' + p4.join('&');
      }
      return null;
    }

    var apiBase = _collState.view === 'fy' ? '/api/fy' : '/api/collection';
    if (session.role === 'CEO') {
      return apiBase + '/by-region?' + dateParam + (ptParam || '');
    }
    if (session.role === 'RM' && session.location) {
      var rp = [ptParam, 'region=' + encodeURIComponent(session.location)].filter(Boolean);
      return apiBase + '/by-district?' + dateParam + rp.join('&');
    }
    if (session.role === 'DM' && session.location) {
      var dp = [ptParam, 'district=' + encodeURIComponent(session.location)].filter(Boolean);
      return apiBase + '/by-branch?' + dateParam + dp.join('&');
    }
    if (session.role === 'BM' && session.location) {
      var bp = [ptParam, 'branch=' + encodeURIComponent(session.location)].filter(Boolean);
      return apiBase + '/by-employee?' + dateParam + bp.join('&');
    }
    return null;
  }

  /* ===================== STATE ===================== */
  var _collState = { view: 'overall', product: 'all', date: null, availableDates: [], summaryData: null, childrenData: null };
  window._collState = _collState;

  /* ===================== RENDERING ===================== */

  /* ---------- Main data loader ---------- */
  function loadAndRender() {
    // Persist selected date so it survives page reloads (Tab toggle, drill-down, etc.)
    if (_collState.date) localStorage.setItem('collSelectedDate', _collState.date);

    var container = document.getElementById('collectionContent');
    container.innerHTML = '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div>' +
      '<div style="color:#64748B;font-size:14px;">Loading collection data...</div></div>';

    var product = _collState.product;
    var apiBase;
    if (_collState.date) {
      apiBase = '/api/daily';
    } else if (isNewStructure()) {
      apiBase = _collState.view === 'fy' ? '/api/v2/fy' : '/api/v2/collection';
    } else {
      apiBase = _collState.view === 'fy' ? '/api/fy' : '/api/collection';
    }
    var params = summaryParams(product);
    if (_collState.date) {
      params += (params ? '&' : '?') + 'date=' + encodeURIComponent(_collState.date);
      // Pass scope so backend filters OA vs FY product types
      params += '&scope=' + (_collState.view === 'fy' ? 'fy' : 'oa');
    }
    var summaryUrl = apiBase + '/summary' + params;
    var childUrl = subUnitsUrl(product);

    var promises = [apiFetch(summaryUrl)];
    if (childUrl && session.role && session.role !== 'FO') {
      promises.push(apiFetch(childUrl));
    } else {
      promises.push(Promise.resolve([]));
    }

    Promise.all(promises).then(function (results) {
      _collState.summaryData = results[0];
      _collState.childrenData = results[1] || [];
      renderCollection();
    }).catch(function (err) {
      console.error('Collection load failed:', err);
      container.innerHTML = noDataHtml('Failed to load collection data.');
    });
  }

  /* ---------- Main render ---------- */
  function renderCollection() {
    var container = document.getElementById('collectionContent');
    var data = _collState.summaryData;

    if (!data || (typeof data === 'object' && Object.keys(data).length === 0)) {
      container.innerHTML = '<div class="desktop-content-area">' + viewToggleHtml() + dateSelectorHtml() + productPillsHtml() + noDataHtml('No data found for this view') + '</div>';
      attachHandlers();
      return;
    }

    var html = '';

    // View toggle
    html += viewToggleHtml();

    // Date selector (daily data)
    html += dateSelectorHtml();

    // Header card (no rank/perf from summary API — show date only)
    html += headerCardHtml(null, null, null);

    // Snapshot
    html += snapshotHtml(data);

    // Product pills
    html += productPillsHtml();

    // Regular demand card
    html += regDemandHtml(data);

    // DPD Bucket cards
    html += dpdBucketsHtml(data);

    // NPA card
    html += npaCardHtml(data);

    // On-Date section
    var odDem = numVal(data.on_date_demand);
    var odCol = numVal(data.on_date_collection);
    var odDemAmt = numVal(data.on_date_demand_amt);
    var odColAmt = numVal(data.on_date_collection_amt);
    if (odDem > 0 || odCol > 0 || odDemAmt > 0 || odColAmt > 0) {
      html += onDateSectionHtml(data);
    }

    // Sub-units
    var children = _collState.childrenData || [];
    if (session.role && session.role !== 'FO' && children.length > 0) {
      var childRoleMap;
      if (_collState.date) {
        // Daily API only has 4-level hierarchy (region/district/branch/employee)
        childRoleMap = { CEO: 'RM', SM: 'DM', RM: 'DM', DvM: 'BM', DM: 'BM', AM: 'FO', BM: 'FO' };
      } else if (isNewStructure()) {
        childRoleMap = { CEO: 'SM', SM: 'DvM', RM: 'DvM', DvM: 'AM', DM: 'AM', AM: 'BM', BM: 'FO' };
      } else {
        childRoleMap = { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
      }
      var childRole = childRoleMap[session.role];
      if (childRole) {
        // Sort children by collection percentage descending
        children.sort(function (a, b) {
          var demA = numVal(a.regular_demand);
          var demB = numVal(b.regular_demand);
          var pctA = demA > 0 ? numVal(a.regular_collection) / demA : 0;
          var pctB = demB > 0 ? numVal(b.regular_collection) / demB : 0;
          return pctB - pctA;
        });
        html += subUnitsHtml(children, childRole);
      }
    }

    container.innerHTML = '<div class="desktop-content-area">' + html + '</div>';
    attachHandlers();
  }

  /* ---------- HTML builders ---------- */
  function noDataHtml(msg) {
    return '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="font-size:36px;margin-bottom:12px;">&#128202;</div>' +
      '<div style="color:#64748B;font-size:14px;">' + (msg || 'No collection data uploaded yet.') + '</div></div>';
  }

  function dateSelectorHtml() { return ""; }

  function viewToggleHtml() {
    var ov = _collState.view === 'overall';
    // Dynamic FY label based on selected date (Indian FY: Apr-Mar)
    var fyLabel = 'FY 25-26';
    var fyDate = _collState.date ? new Date(_collState.date) : new Date();
    var fyMonth = fyDate.getMonth() + 1; // 1-12
    var fyYear = fyDate.getFullYear();
    var fyStartYear = fyMonth >= 4 ? fyYear : fyYear - 1;
    fyLabel = 'FY ' + String(fyStartYear).slice(-2) + '-' + String(fyStartYear + 1).slice(-2);
    return '<div class="pf-view-toggle emp-fade">' +
      '<button class="pf-view-btn' + (ov ? ' active' : '') + '" data-coll-view="overall">OverAll</button>' +
      '<button class="pf-view-btn' + (!ov ? ' active' : '') + '" data-coll-view="fy">' + fyLabel + '</button>' +
    '</div>';
  }

  function headerCardHtml(reportDate, rank, perf) {
    var perfText = String(perf || '').trim();
    var perfUpper = perfText.toUpperCase();
    var isAbove = perfUpper.includes('ABOVE');
    var isBelow = perfUpper.includes('BELOW');
    var perfColor = isAbove ? '#34D399' : isBelow ? '#F87171' : '#64748B';
    var perfIcon = isAbove ? '&#9650;' : isBelow ? '&#9660;' : '';
    var perfLabel = isAbove ? 'Above Average' : isBelow ? 'Below Average' : (perfText || '-');

    var rankVal = (rank != null && rank !== '-' && rank !== '') ? rank : '-';

    return '<div class="pf-header-card emp-fade">' +
      '<div class="pf-header-top">' +
        '<div class="pf-report-date">' +
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>' +
          '<span>' + (reportDate || 'Month End') + '</span>' +
        '</div>' +
        '<div class="pf-badges">' +
          (rankVal !== '-' ? '<div class="pf-rank-badge">#' + esc(String(rankVal)) + '</div>' : '') +
          '<div class="pf-perf-badge" style="color:' + perfColor + ';border-color:' + perfColor + '30">' +
            perfIcon + ' ' + esc(perfLabel) +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function snapshotHtml(data) {
    var totalDem = numVal(data.regular_demand);
    var totalCol = numVal(data.regular_collection);
    var overallPct = totalDem > 0 ? ((totalCol / totalDem) * 100).toFixed(1) : '0';

    return '<div class="emp-snapshot emp-fade">' +
      '<div class="emp-snapshot-label">Snapshot</div>' +
      '<div class="desktop-grid-2">' +
      '<div class="emp-snapshot-row">' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Demand</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalDem) + '</div>' +
        '</div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Collection</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalCol) + '</div>' +
        '</div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Overall %</div>' +
          '<div class="emp-snap-val" style="color:#059669;font-size:24px;">' + overallPct + '%</div>' +
        '</div>' +
      '</div>' +
      '</div>' +
    '</div>';
  }

  function productPillsHtml() {
    var p = _collState.product;
    return '<div class="emp-product-filter desktop-pills-row">' +
      '<button class="emp-product-pill' + (p === 'all' ? ' active' : '') + '" data-coll-product="all">All</button>' +
      '<button class="emp-product-pill' + (p === 'igl' ? ' active' : '') + '" data-coll-product="igl">IGL</button>' +
      '<button class="emp-product-pill' + (p === 'fig' ? ' active' : '') + '" data-coll-product="fig">FIG</button>' +
      '<button class="emp-product-pill' + (p === 'il' ? ' active' : '') + '" data-coll-product="il">IL</button>' +
    '</div>';
  }

  function regDemandHtml(data) {
    var dem = numVal(data.regular_demand);
    var col = numVal(data.regular_collection);
    var ftod = dem - col;
    var pctVal = dem > 0 ? (col / dem) * 100 : 0;
    var noData = (dem === 0 && col === 0);
    var pctColor = noData ? '#64748B' : pctVal >= 99 ? '#34D399' : pctVal >= 95 ? '#FBBF24' : '#F87171';
    var barW = Math.min(pctVal, 100);

    return '<div class="emp-data-section desktop-card">' +
      '<div class="emp-section-title">Regular Demand vs Collection</div>' +
      '<div class="pf-reg-card emp-fade">' +
        '<div class="pf-reg-main">' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val">' + fmtNum(dem) + '</div>' +
            '<div class="pf-reg-lbl">Demand</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:#059669">' + fmtNum(col) + '</div>' +
            '<div class="pf-reg-lbl">Collection</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:#FB923C">' + fmtNum(ftod) + '</div>' +
            '<div class="pf-reg-lbl">FTOD</div>' +
          '</div>' +
          '<div class="pf-reg-stat">' +
            '<div class="pf-reg-val" style="color:' + pctColor + '">' + (noData ? '-' : pctVal.toFixed(2) + '%') + '</div>' +
            '<div class="pf-reg-lbl">Collection %</div>' +
          '</div>' +
        '</div>' +
        '<div class="pf-progress-track">' +
          '<div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function dpdBucketsHtml(data) {
    var buckets = [
      { name: 'SMA-0', range: '1 \u2014 30 DPD', color: '#34D399',
        d: 'demand_1_30', c: 'collection_1_30' },
      { name: 'SMA-1', range: '31 \u2014 60 DPD', color: '#FBBF24',
        d: 'demand_31_60', c: 'collection_31_60' },
      { name: 'SMA-2', range: 'Pre-NPA', color: '#FB923C',
        d: 'pnpa_demand', c: 'pnpa_collection' }
    ];

    var html = '<div class="emp-section-title" style="margin-top:20px;">SMA Bucket Allocation</div>';
    html += '<div class="emp-buckets desktop-grid-2">';

    for (var i = 0; i < buckets.length; i++) {
      var bk = buckets[i];
      var dem = numVal(data[bk.d]);
      var col = numVal(data[bk.c]);
      var bal = dem - col;
      var pctVal = dem > 0 ? (col / dem) * 100 : 0;
      var noData = (dem === 0 && col === 0);
      var pctColor = noData ? '#64748B' : pctVal >= 50 ? '#34D399' : pctVal >= 25 ? '#FBBF24' : '#F87171';
      var barW = noData ? 0 : Math.min(pctVal, 100);

      html += '<div class="emp-bucket-card emp-fade" style="animation-delay:' + (0.1 + i * 0.05) + 's">' +
        '<div class="emp-bucket-indicator" style="background:' + bk.color + '"></div>' +
        '<div class="emp-bucket-info">' +
          '<div class="emp-bucket-name">' + bk.name + '</div>' +
          '<div class="emp-bucket-sub">' + bk.range + '</div>' +
        '</div>' +
        '<div class="emp-bucket-stats">' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:' + bk.color + '">' + fmtNum(dem) + '</div>' +
            '<div class="emp-bucket-lbl">Demand</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val emp-pos-val">' + fmtNum(col) + '</div>' +
            '<div class="emp-bucket-lbl">Collection</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:#64748B">' + fmtNum(bal) + '</div>' +
            '<div class="emp-bucket-lbl">Balance</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:' + pctColor + '">' + (noData ? '-' : pctVal.toFixed(2) + '%') + '</div>' +
            '<div class="emp-bucket-lbl">Coll %</div>' +
          '</div>' +
        '</div>' +
        (barW === 0
          ? '<div class="emp-bucket-bar" style="width:100%;height:2px;background:#E2E8F0"></div>'
          : '<div class="emp-bucket-bar" style="width:' + Math.max(barW, 8) + '%;height:4px;opacity:0.85;background:' + bk.color + '"></div>') +
      '</div>';
    }

    html += '</div>';
    return html;
  }

  function npaCardHtml(data) {
    var demand = numVal(data.npa_cases);
    var actAcct = numVal(data.npa_act_acc);
    var actAmt = parseFloat(data.npa_act_amt) || 0;
    var clsAcct = numVal(data.npa_clo_acc);
    var clsAmt = parseFloat(data.npa_clo_amt) || 0;

    return '<div class="pf-npa-card desktop-card emp-fade" style="animation-delay:0.25s">' +
      '<div class="pf-npa-header">' +
        '<div class="emp-bucket-indicator" style="background:#F87171"></div>' +
        '<div class="emp-bucket-info">' +
          '<div class="emp-bucket-name">NPA</div>' +
          '<div class="emp-bucket-sub">Non-Performing Assets</div>' +
        '</div>' +
        '<div class="pf-npa-demand">' +
          '<div class="emp-bucket-val" style="color:#F87171">' + fmtNum(demand) + '</div>' +
          '<div class="emp-bucket-lbl">Demand</div>' +
        '</div>' +
      '</div>' +
      '<div class="pf-npa-grid">' +
        '<div class="pf-npa-section">' +
          '<div class="pf-npa-section-title" style="color:#34D399">Activation</div>' +
          '<div class="pf-npa-row">' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(actAcct) + '</div><div class="pf-npa-lbl">Account</div></div>' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(actAmt) + '</div><div class="pf-npa-lbl">Amount</div></div>' +
          '</div>' +
        '</div>' +
        '<div class="pf-npa-divider"></div>' +
        '<div class="pf-npa-section">' +
          '<div class="pf-npa-section-title" style="color:#059669">Closure</div>' +
          '<div class="pf-npa-row">' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(clsAcct) + '</div><div class="pf-npa-lbl">Account</div></div>' +
            '<div class="pf-npa-item"><div class="pf-npa-val">' + fmtNum(clsAmt) + '</div><div class="pf-npa-lbl">Amount</div></div>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  function onDateSectionHtml(data) {
    var odDem = numVal(data.on_date_demand);
    var odCol = numVal(data.on_date_collection);
    var odPct = odDem > 0 ? (odCol / odDem) : 0;

    var html = '<div class="emp-section-title" style="margin-top:20px;">On-Date Status</div>' +
      '<div class="pf-ondate-grid desktop-grid-2 emp-fade" style="animation-delay:0.3s">';

    // Today card
    html += '<div class="pf-ondate-card">' +
      '<div class="pf-ondate-label">Today</div>' +
      '<div class="pf-ondate-stats">' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val">' + fmtNum(odDem) + '</div><div class="pf-ondate-lbl">Demand</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#059669">' + fmtNum(odCol) + '</div><div class="pf-ondate-lbl">Collection</div></div>' +
        '<div class="pf-ondate-stat"><div class="pf-ondate-val" style="color:#34D399">' + fmtPct(odPct) + '</div><div class="pf-ondate-lbl">Coll %</div></div>' +
      '</div>' +
    '</div>';

    html += '</div>';
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

  function getChildName(ch, childRole) {
    // Check both old and new structure field names (daily API returns old, v2 returns new)
    if (childRole === 'SM') return ch.state_name || ch.region_name || '';
    if (childRole === 'RM') return ch.region_name || ch.state_name || '';
    if (childRole === 'DvM') return ch.division_name || ch.district_name || '';
    if (childRole === 'DM') return ch.district_name || ch.division_name || '';
    if (childRole === 'AM') return ch.area_name || '';
    if (childRole === 'BM') return ch.branch_name || '';
    if (childRole === 'FO') return getFullName(ch.emp_id, ch.name || ch.officer_name || '');
    return '';
  }

  function getDataAttr(ch, childRole, childName) {
    if (childRole === 'FO') {
      return 'data-emp-id="' + esc(ch.emp_id || '') + '" data-emp-name="' + esc(childName) + '"';
    }
    return 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(childName) + '"';
  }

  function subUnitsCardsHtml(children, childRole, av) {
    var html = '<div class="desktop-grid-3">';
    for (var i = 0; i < children.length; i++) {
      var ch = children[i];
      var dem = numVal(ch.regular_demand);
      var col = numVal(ch.regular_collection);
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      var pct = pctRaw.toFixed(2);
      var pctColor = pctRaw >= 95 ? '#34D399' : pctRaw >= 80 ? '#FBBF24' : '#F87171';
      var childName = getChildName(ch, childRole);
      var dataAttr = getDataAttr(ch, childRole, childName);
      var initial = childName.charAt(0).toUpperCase();

      html += '<div class="emp-sub-card desktop-sub-card" ' + dataAttr + '>' +
        '<div class="emp-sub-avatar ' + av + '">' + initial + '</div>' +
        '<div class="emp-sub-info">' +
          '<div class="emp-sub-name">' + esc(childName) + '</div>' +
          '<div class="emp-sub-meta">' +
            '<span>D: ' + fmtNum(dem) + '</span>' +
            '<span>C: ' + fmtNum(col) + '</span>' +
          '</div>' +
        '</div>' +
        '<div class="emp-sub-pct" style="color:' + pctColor + '">' + pct + '%</div>' +
        '<div class="emp-sub-arrow">&#8250;</div>' +
      '</div>';
    }
    html += '</div>';
    return html;
  }

  function subUnitsTableHtml(children, childRole, av) {
    var avColors = { state: 'background:#F3E8FF;color:#7C3AED;', region: 'background:#F3E8FF;color:#7C3AED;', division: 'background:#DBEAFE;color:#2563EB;', area: 'background:#ECFDF5;color:#059669;', district: 'background:#DBEAFE;color:#2563EB;', branch: 'background:#ECFDF5;color:#059669;', officer: 'background:#FEF3C7;color:#D97706;' };
    var avStyle = avColors[av] || 'background:#F1F5F9;color:#64748B;';

    var html = '<div class="subunit-table-wrap"><table class="subunit-table">';
    html += '<thead><tr>' +
      '<th>Unit <span class="sort-icon">&#8597;</span></th>' +
      '<th class="num-col">Demand <span class="sort-icon">&#8597;</span></th>' +
      '<th class="num-col">Collection <span class="sort-icon">&#8597;</span></th>' +
      '<th class="num-col">Balance <span class="sort-icon">&#8597;</span></th>' +
      '<th class="num-col">Coll% <span class="sort-icon">&#8597;</span></th>' +
      '<th style="width:24px;"></th>' +
    '</tr></thead><tbody>';

    for (var i = 0; i < children.length; i++) {
      var ch = children[i];
      var dem = numVal(ch.regular_demand);
      var col = numVal(ch.regular_collection);
      var bal = dem - col;
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      var pct = pctRaw.toFixed(2);
      var pctColor = pctRaw >= 95 ? '#34D399' : pctRaw >= 80 ? '#FBBF24' : '#F87171';
      var childName = getChildName(ch, childRole);
      var dataAttr = getDataAttr(ch, childRole, childName);
      var initial = childName.charAt(0).toUpperCase();

      html += '<tr class="emp-sub-card" ' + dataAttr + '>' +
        '<td class="name-col"><span class="tbl-avatar" style="' + avStyle + '">' + initial + '</span>' + esc(childName) + '</td>' +
        '<td class="num-col">' + fmtNum(dem) + '</td>' +
        '<td class="num-col" style="color:#059669;font-weight:600;">' + fmtNum(col) + '</td>' +
        '<td class="num-col" style="color:#FB923C;">' + fmtNum(bal) + '</td>' +
        '<td class="num-col"><span class="tbl-pct" style="color:' + pctColor + '">' + pct + '%</span></td>' +
        '<td class="tbl-arrow">&#8250;</td>' +
      '</tr>';
    }

    html += '</tbody></table></div>';
    return html;
  }

  function subUnitsHtml(children, childRole) {
    if (!children.length) return '';
    var roleLabel = isNewStructure()
      ? { SM: 'Regions', DvM: 'Divisions', AM: 'Areas', BM: 'Branches', FO: 'Officers' }
      : { RM: 'Regions', DM: 'Divisions', BM: 'Branches', FO: 'Officers' };
    var avType = isNewStructure()
      ? { SM: 'state', DvM: 'division', AM: 'area', BM: 'branch', FO: 'officer' }
      : { RM: 'region', DM: 'district', BM: 'branch', FO: 'officer' };
    var label = roleLabel[childRole] || 'Team';
    var av = avType[childRole] || 'branch';
    var isTable = getSubunitView() === 'table';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label + '<span class="emp-team-count">' + children.length + '</span>' + subunitToggleHtml() + '</div>';

    if (isTable) {
      html += subUnitsTableHtml(children, childRole, av);
    } else {
      html += subUnitsCardsHtml(children, childRole, av);
    }

    html += '</div>';
    return html;
  }

  /* ========== Event Handlers ========== */
  var _collHandlersAttached = false;

  function attachHandlers() {
    var container = document.getElementById('collectionContent');
    if (!container) return;

    // View toggle
    container.querySelectorAll('[data-coll-view]').forEach(function (btn) {
      btn.onclick = function () {
        _collState.view = btn.dataset.collView;
        localStorage.setItem('collView', btn.dataset.collView);
        loadAndRender();
      };
    });

    // Date selector
    container.querySelectorAll('[data-coll-date]').forEach(function (btn) {
      btn.onclick = function () {
        _collState.date = btn.dataset.collDate || null;
        if (_collState.date) localStorage.setItem('collSelectedDate', _collState.date);
        else localStorage.removeItem('collSelectedDate');
        loadAndRender();
      };
    });

    // Product pills
    container.querySelectorAll('[data-coll-product]').forEach(function (pill) {
      pill.onclick = function () {
        _collState.product = pill.dataset.collProduct;
        localStorage.setItem('collProduct', pill.dataset.collProduct);
        loadAndRender();
      };
    });

    // Subunit view toggle (card/table)
    container.querySelectorAll('[data-subview]').forEach(function (btn) {
      btn.onclick = function () {
        localStorage.setItem('subunitView', btn.dataset.subview);
        renderCollection();
      };
    });

    // Table column sorting
    container.querySelectorAll('.subunit-table thead th').forEach(function (th, colIdx) {
      th.onclick = function () {
        var table = th.closest('.subunit-table');
        var tbody = table.querySelector('tbody');
        var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
        var asc = th.classList.contains('sorted-asc');
        // Clear all sorted states
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

    // Sub-unit drill-down (only attach once to avoid duplicates)
    if (!_collHandlersAttached) {
      _collHandlersAttached = true;
      container.addEventListener('click', function (ev) {
        var card = ev.target.closest('.emp-sub-card');
        if (!card) return;
        // Only handle if we're in the collection tab
        var collectionTab = document.getElementById('collectionTab');
        if (!collectionTab || !collectionTab.classList.contains('active')) return;

        card.style.background = '#ECFDF5';
        card.style.pointerEvents = 'none';
        var arrow = card.querySelector('.emp-sub-arrow');
        if (arrow) arrow.innerHTML = '<div style="width:16px;height:16px;border:2px solid #A7F3D0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;"></div>';

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

  /* ========== Load Collection Data ========== */
  window._reloadCollection = loadAndRender;
  window._loadCollectionTab = function () {
    try {
      // Restore saved collection state
      var savedView = localStorage.getItem('collView');
      if (savedView && (savedView === 'overall' || savedView === 'fy')) {
        _collState.view = savedView;
      }
      var savedProduct = localStorage.getItem('collProduct');
      if (savedProduct && ['all', 'igl', 'fig', 'il'].indexOf(savedProduct) !== -1) {
        _collState.product = savedProduct;
      }

      // Restore saved date from localStorage (preserves date across drill-down navigation & Tab toggle)
      var savedDate = localStorage.getItem('collSelectedDate');
      if (savedDate) _collState.date = savedDate;

      // Always fetch available dates (needed for arrow navigation)
      apiFetch('/api/daily/dates').then(function(dates) {
        if (dates && dates.length) {
          _collState.availableDates = dates.map(function(d) { return d.substring(0, 10); });
          // If no saved date, use latest
          if (!_collState.date) {
            _collState.date = _collState.availableDates[0];
          }
        }
        // Sync the header date badge
        var badge = document.getElementById('dateBadgeText');
        if (badge && _collState.date) {
          var p = _collState.date.split('-');
          badge.textContent = p[2] + '-' + p[1] + '-' + p[0];
        }
        loadAndRender();
      }).catch(function() { loadAndRender(); });
    } catch (err) {
      console.error('Collection load failed:', err);
      document.getElementById('collectionContent').innerHTML = noDataHtml('Failed to load collection data.');
    }
  };
})();
