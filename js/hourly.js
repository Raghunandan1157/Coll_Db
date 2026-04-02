/**
 * hourly.js -- Hourly tab renderer for the employee dashboard.
 * Mirrors portfolio.js exactly: product pills, on-date preview,
 * DPD buckets, NPA details, and drill-down.
 *
 * Data source: PostgreSQL REST API
 * Container: hourlyContent
 */
(function () {

  /* ========== Session ========== */
  var session = getEmployeeSession();

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
    if (productType && productType !== 'all') {
      params.push('product_type=' + encodeURIComponent(productType.toUpperCase()));
    }
    if (session.role === 'RM' && session.location) {
      params.push('region=' + encodeURIComponent(session.location));
    } else if (session.role === 'DM' && session.location) {
      params.push('district=' + encodeURIComponent(session.location));
    } else if (session.role === 'BM' && session.location) {
      params.push('branch=' + encodeURIComponent(session.location));
    } else if ((!session.role || session.role === 'FO') && session.id) {
      params.push('emp_id=' + encodeURIComponent(session.id));
    }
    return params.length ? '?' + params.join('&') : '';
  }

  /* ========== Build query params for sub-units endpoint ========== */
  function subUnitsUrl(productType) {
    var pt = (productType && productType !== 'all') ? productType.toUpperCase() : '';
    var ptParam = pt ? 'product_type=' + encodeURIComponent(pt) : '';

    if (session.role === 'CEO') {
      return '/api/hourly/by-region' + (ptParam ? '?' + ptParam : '');
    }
    if (session.role === 'RM' && session.location) {
      var p = [ptParam, 'region=' + encodeURIComponent(session.location)].filter(Boolean);
      return '/api/hourly/by-district?' + p.join('&');
    }
    if (session.role === 'DM' && session.location) {
      var p2 = [ptParam, 'district=' + encodeURIComponent(session.location)].filter(Boolean);
      return '/api/hourly/by-branch?' + p2.join('&');
    }
    if (session.role === 'BM' && session.location) {
      var p3 = [ptParam, 'branch=' + encodeURIComponent(session.location)].filter(Boolean);
      return '/api/hourly/by-employee?' + p3.join('&');
    }
    return null;
  }

  /* ===================== STATE ===================== */
  var _hourlyState = { view: 'overall', product: 'all', summaryData: null, childrenData: null };

  /* ===================== RENDERING ===================== */

  /* ---------- Main data loader ---------- */
  function loadAndRender() {
    var container = document.getElementById('hourlyContent');
    container.innerHTML = '<div style="text-align:center;padding:80px 20px;">' +
      '<div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#059669;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div>' +
      '<div style="color:#64748B;font-size:14px;">Loading hourly data...</div></div>';

    var product = _hourlyState.product;
    var summaryUrl = '/api/hourly/summary' + summaryParams(product);
    var childUrl = subUnitsUrl(product);

    var promises = [apiFetch(summaryUrl)];
    if (childUrl && session.role && session.role !== 'FO') {
      promises.push(apiFetch(childUrl));
    } else {
      promises.push(Promise.resolve([]));
    }

    Promise.all(promises).then(function (results) {
      _hourlyState.summaryData = results[0];
      _hourlyState.childrenData = results[1] || [];
      renderCollection();
    }).catch(function (err) {
      console.error('Hourly load failed:', err);
      container.innerHTML = noDataHtml('Failed to load hourly data.');
    });
  }

  /* ---------- Main render ---------- */
  function renderCollection() {
    var container = document.getElementById('hourlyContent');
    var data = _hourlyState.summaryData;

    if (!data || (typeof data === 'object' && Object.keys(data).length === 0)) {
      container.innerHTML = '<div class="desktop-content-area">' + viewToggleHtml() + productPillsHtml() + noDataHtml('No data found for this view') + '</div>';
      attachHandlers();
      return;
    }

    var html = '';

    // View toggle
    html += viewToggleHtml();

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
    var children = _hourlyState.childrenData || [];
    if (session.role && session.role !== 'FO' && children.length > 0) {
      var childRoleMap = { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
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
      '<div style="color:#64748B;font-size:14px;">' + (msg || 'No hourly data uploaded yet.') + '</div></div>';
  }

  function viewToggleHtml() {
    var ov = _hourlyState.view === 'overall';
    return '<div class="pf-view-toggle emp-fade">' +
      '<button class="pf-view-btn' + (ov ? ' active' : '') + '" data-hourly-view="overall">OverAll</button>' +
      '<button class="pf-view-btn' + (!ov ? ' active' : '') + '" data-hourly-view="fy">FY 25-26</button>' +
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
    var totalDem = numVal(data.regular_demand) + numVal(data.demand_1_30) + numVal(data.demand_31_60) + numVal(data.pnpa_demand) + numVal(data.npa_cases);
    var totalCol = numVal(data.regular_collection) + numVal(data.collection_1_30) + numVal(data.collection_31_60) + numVal(data.pnpa_collection) + numVal(data.npa_act_acc);
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
    var p = _hourlyState.product;
    return '<div class="emp-product-filter desktop-pills-row">' +
      '<button class="emp-product-pill' + (p === 'all' ? ' active' : '') + '" data-hourly-product="all">All</button>' +
      '<button class="emp-product-pill' + (p === 'igl' ? ' active' : '') + '" data-hourly-product="igl">IGL</button>' +
      '<button class="emp-product-pill' + (p === 'fig' ? ' active' : '') + '" data-hourly-product="fig">FIG</button>' +
      '<button class="emp-product-pill' + (p === 'il' ? ' active' : '') + '" data-hourly-product="il">IL</button>' +
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

  function subUnitsHtml(children, childRole) {
    if (!children.length) return '';
    var roleLabel = { RM: 'Regions', DM: 'Districts', BM: 'Branches', FO: 'Officers' };
    var avType = { RM: 'region', DM: 'district', BM: 'branch', FO: 'officer' };
    var label = roleLabel[childRole] || 'Team';
    var av = avType[childRole] || 'branch';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label + '<span class="emp-team-count">' + children.length + '</span></div>' +
      '<div class="desktop-grid-3">';

    for (var i = 0; i < children.length; i++) {
      var ch = children[i];
      var dem = numVal(ch.regular_demand);
      var col = numVal(ch.regular_collection);
      var pctRaw = dem > 0 ? (col / dem) * 100 : 0;
      var pct = pctRaw.toFixed(2);
      var pctColor = pctRaw >= 95 ? '#34D399' : pctRaw >= 80 ? '#FBBF24' : '#F87171';

      // Determine the display name and data attributes based on child role
      var childName = '';
      var dataAttr = '';
      if (childRole === 'RM') {
        childName = ch.region_name || '';
      } else if (childRole === 'DM') {
        childName = ch.district_name || '';
      } else if (childRole === 'BM') {
        childName = ch.branch_name || '';
      } else if (childRole === 'FO') {
        childName = ch.name || '';
      }

      if (childRole === 'FO') {
        dataAttr = 'data-emp-id="' + esc(ch.emp_id || '') + '" data-emp-name="' + esc(childName) + '"';
      } else {
        dataAttr = 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(childName) + '"';
      }

      var initial = childName.charAt(0).toUpperCase();
      var rankBadge = '';

      html += '<div class="emp-sub-card desktop-sub-card" ' + dataAttr + '>' +
        '<div class="emp-sub-avatar ' + av + '">' + initial + '</div>' +
        '<div class="emp-sub-info">' +
          '<div class="emp-sub-name">' + esc(childName) + '</div>' +
          '<div class="emp-sub-meta">' +
            '<span>D: ' + fmtNum(dem) + '</span>' +
            '<span>C: ' + fmtNum(col) + '</span>' +
            rankBadge +
          '</div>' +
        '</div>' +
        '<div class="emp-sub-pct" style="color:' + pctColor + '">' + pct + '%</div>' +
        '<div class="emp-sub-arrow">&#8250;</div>' +
      '</div>';
    }

    html += '</div></div>';
    return html;
  }

  /* ========== Event Handlers ========== */
  var _hourlyHandlersAttached = false;

  function attachHandlers() {
    var container = document.getElementById('hourlyContent');
    if (!container) return;

    // View toggle
    container.querySelectorAll('[data-hourly-view]').forEach(function (btn) {
      btn.onclick = function () {
        _hourlyState.view = btn.dataset.hourlyView;
        localStorage.setItem('hourlyView', btn.dataset.hourlyView);
        loadAndRender();
      };
    });

    // Product pills
    container.querySelectorAll('[data-hourly-product]').forEach(function (pill) {
      pill.onclick = function () {
        _hourlyState.product = pill.dataset.hourlyProduct;
        localStorage.setItem('hourlyProduct', pill.dataset.hourlyProduct);
        loadAndRender();
      };
    });

    // Sub-unit drill-down (only attach once to avoid duplicates)
    if (!_hourlyHandlersAttached) {
      _hourlyHandlersAttached = true;
      container.addEventListener('click', function (ev) {
        var card = ev.target.closest('.emp-sub-card');
        if (!card) return;
        // Only handle if we are in the hourly tab
        var hourlyTab = document.getElementById('hourlyTab');
        if (!hourlyTab || !hourlyTab.classList.contains('active')) return;

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
  window._loadHourlyTab = function () {
    try {
      // Restore saved hourly state
      var savedView = localStorage.getItem('hourlyView');
      if (savedView && (savedView === 'overall' || savedView === 'fy')) {
        _hourlyState.view = savedView;
      }
      var savedProduct = localStorage.getItem('hourlyProduct');
      if (savedProduct && ['all', 'igl', 'fig', 'il'].indexOf(savedProduct) !== -1) {
        _hourlyState.product = savedProduct;
      }

      loadAndRender();
    } catch (err) {
      console.error('Hourly load failed:', err);
      document.getElementById('hourlyContent').innerHTML = noDataHtml('Failed to load hourly data.');
    }
  };
})();
