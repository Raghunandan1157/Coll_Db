(function () {
  'use strict';

  if (!document.getElementById('analytical-styles')) {
    var style = document.createElement('style');
    style.id = 'analytical-styles';
    style.textContent = [
      '.anal-container { padding: 12px 10px 80px; }',
      '.anal-page-title { font-size: 20px; font-weight: 700; color: #1E293B; margin-bottom: 4px; }',
      '.anal-page-subtitle { font-size: 12px; color: #64748B; margin-bottom: 16px; }',
      '.anal-btn-row { display: flex; gap: 10px; margin-bottom: 20px; }',
      '.anal-btn {',
      '  padding: 10px 16px; border-radius: 12px; border: 1px solid #E2E8F0;',
      '  background: #FFFFFF; color: #64748B; font-size: 13px; font-weight: 600;',
      '  cursor: pointer; transition: all 0.2s; flex: 1; text-align: center;',
      '  display: flex; align-items: center; justify-content: center; gap: 6px;',
      '}',
      '.anal-btn.active { background: #ECFDF5; color: #059669; border-color: #A7F3D0; }',
      '.anal-btn:active { transform: scale(0.97); }',
      '.anal-section { margin-bottom: 24px; }',
      '.anal-section-header { display: flex; align-items: center; gap: 10px; margin-bottom: 14px; padding: 0 2px; }',
      '.anal-section-icon { width: 32px; height: 32px; border-radius: 10px; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }',
      '.anal-section-icon.top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-section-icon.bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-section-title { font-size: 15px; font-weight: 700; color: #1E293B; }',
      '.anal-section-count { font-size: 11px; font-weight: 500; color: #64748B; background: #F1F5F9; padding: 2px 8px; border-radius: 10px; margin-left: auto; }',
      '.anal-empty { text-align: center; color: #64748B; padding: 40px 0; font-size: 13px; }',
      '.anal-list { display: flex; flex-direction: column; gap: 10px; }',
      '.anal-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 14px; transition: border-color 0.2s, transform 0.2s; }',
      '.anal-card:active { border-color: #CBD5E1; transform: scale(0.99); }',
      '.anal-card-header { display: flex; align-items: center; gap: 8px; margin-bottom: 10px; }',
      '.anal-rank { font-size: 11px; font-weight: 700; padding: 3px 7px; border-radius: 8px; min-width: 28px; text-align: center; flex-shrink: 0; }',
      '.anal-branch-avatar { width: 30px; height: 30px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px; flex-shrink: 0; }',
      '.anal-avatar-top { background: rgba(52,211,153,0.12); color: #34D399; }',
      '.anal-avatar-bottom { background: rgba(248,113,113,0.12); color: #F87171; }',
      '.anal-branch-name { flex: 1; font-size: 13px; font-weight: 600; color: #1E293B; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }',
      '.anal-pct { font-size: 16px; font-weight: 800; flex-shrink: 0; font-family: "DM Sans", sans-serif; }',
      '.anal-card-stats { display: flex; justify-content: space-between; margin-bottom: 8px; }',
      '.anal-stat { display: flex; flex-direction: column; align-items: center; gap: 1px; flex: 1; }',
      '.anal-stat-lbl { font-size: 9px; text-transform: uppercase; color: #64748B; letter-spacing: 0.5px; }',
      '.anal-stat-val { font-size: 13px; font-weight: 700; color: #1E293B; }',
      '.anal-divider { height: 1px; background: #F1F5F9; margin: 16px 0; }',
      // FO card styles
      '.fo-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 14px; margin-bottom: 10px; cursor: pointer; transition: all 0.2s; }',
      '.fo-card:active { transform: scale(0.99); }',
      '.fo-header { display: flex; align-items: center; gap: 10px; }',
      '.fo-rank { font-size: 11px; font-weight: 700; padding: 3px 7px; border-radius: 8px; min-width: 28px; text-align: center; flex-shrink: 0; }',
      '.fo-avatar { width: 30px; height: 30px; border-radius: 50%; background: rgba(248,113,113,0.12); color: #F87171; display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px; flex-shrink: 0; }',
      '.fo-info { flex: 1; min-width: 0; }',
      '.fo-name { font-size: 13px; font-weight: 600; color: #1E293B; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }',
      '.fo-empid { font-size: 10px; color: #64748B; }',
      '.fo-pct { font-size: 16px; font-weight: 800; flex-shrink: 0; font-family: "DM Sans", sans-serif; }',
      '.fo-stats { display: flex; justify-content: space-between; margin-top: 10px; }',
      '.fo-actions { display: flex; gap: 8px; margin-top: 12px; }',
      '.fo-call-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: rgba(52,211,153,0.1); color: #34D399; font-size: 13px; font-weight: 600; text-decoration: none; border: 1px solid rgba(52,211,153,0.2); transition: all 0.2s; }',
      '.fo-call-btn:active { transform: scale(0.97); background: rgba(52,211,153,0.2); }',
      '.fo-no-phone { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: #F1F5F9; color: #64748B; font-size: 12px; border: 1px solid #F1F5F9; }',
      '.fo-view-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 6px; padding: 10px; border-radius: 10px; background: #F0FDF4; color: #059669; font-size: 13px; font-weight: 600; cursor: pointer; border: 1px solid #A7F3D0; transition: all 0.2s; }',
      '.fo-view-btn:active { transform: scale(0.97); background: #A7F3D0; }',
      '.fo-branch-tag { font-size: 10px; color: #059669; background: #F0FDF4; padding: 2px 6px; border-radius: 6px; margin-left: 4px; }'
    ].join('\n');
    document.head.appendChild(style);
  }

  var SUPABASE_URL = 'https://zovnmmdfthpbubrorsgh.supabase.co';
  var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inpvdm5tbWRmdGhwYnVicm9yc2doIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE1NzE3ODgsImV4cCI6MjA3NzE0Nzc4OH0.92BH2sjUOgkw6iSRj1_4gt0p3eThg3QT4VK-Q4EdmBE';

  function numVal(v) { return typeof v === 'number' ? v : (parseFloat(String(v).replace(/,/g, '')) || 0); }
  function fmtNum(n) { return n.toLocaleString('en-IN'); }
  function esc(s) { var d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

  /* ── API helpers ── */

  function buildQuery(base, params) {
    var parts = [];
    for (var k in params) {
      if (params[k] !== undefined && params[k] !== null && params[k] !== '') {
        parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
      }
    }
    return base + (parts.length ? '?' + parts.join('&') : '');
  }

  function apiFetch(url) {
    return fetch(url).then(function (resp) {
      if (!resp.ok) throw new Error('API error: ' + resp.status);
      return resp.json();
    });
  }

  function fetchBranches(productType, district) {
    var params = { product_type: productType };
    if (district) params.district = district;
    return apiFetch(buildQuery('/api/collection/by-branch', params));
  }

  function fetchDistricts(productType, region) {
    var params = { product_type: productType };
    if (region) params.region = region;
    return apiFetch(buildQuery('/api/collection/by-district', params));
  }

  function fetchRegions(productType) {
    return apiFetch(buildQuery('/api/collection/by-region', { product_type: productType }));
  }

  function fetchEmployees(productType, branch) {
    var params = { product_type: productType };
    if (branch) params.branch = branch;
    return apiFetch(buildQuery('/api/collection/by-employee', params));
  }

  /* ── Metric extraction from an API row ── */

  function extractBucketMetrics(row, bucket) {
    var dem, col;
    if (bucket === 'dpd1') {
      dem = numVal(row.demand_1_30);
      col = numVal(row.collection_1_30);
    } else if (bucket === 'dpd2') {
      dem = numVal(row.demand_31_60);
      col = numVal(row.collection_31_60);
    } else if (bucket === 'pnpa') {
      dem = numVal(row.pnpa_demand);
      col = numVal(row.pnpa_collection);
    } else if (bucket === 'npa') {
      dem = numVal(row.npa_cases);
      col = parseFloat(row.npa_act_amt) || 0;
    } else {
      // regular
      dem = numVal(row.regular_demand);
      col = numVal(row.regular_collection);
    }
    return { demand: dem, collection: col };
  }

  function computeBranchData(apiRows) {
    var data = [];
    for (var i = 0; i < apiRows.length; i++) {
      var r = apiRows[i];
      var dem = numVal(r.regular_demand);
      var col = numVal(r.regular_collection);
      data.push({
        name: (r.branch_name || '').toUpperCase(),
        demand: dem,
        collection: col,
        balance: dem - col,
        pct: dem > 0 ? (col / dem) * 100 : 0
      });
    }
    data.sort(function (a, b) { return a.pct - b.pct; });
    return data;
  }

  function computeFODataFromAPI(apiRows, bucket) {
    var data = [];
    for (var i = 0; i < apiRows.length; i++) {
      var o = apiRows[i];
      var m = extractBucketMetrics(o, bucket);
      var dem = m.demand;
      var col = m.collection;
      if (dem === 0 && col === 0) continue;
      var pct = dem > 0 ? (col / dem) * 100 : 0;

      // Regular
      var regDem = numVal(o.regular_demand);
      var regCol = numVal(o.regular_collection);
      // DPD1
      var dpd1Dem = numVal(o.demand_1_30);
      var dpd1Col = numVal(o.collection_1_30);
      // DPD2
      var dpd2Dem = numVal(o.demand_31_60);
      var dpd2Col = numVal(o.collection_31_60);
      // PNPA
      var pnpaDem = numVal(o.pnpa_demand);
      var pnpaCol = numVal(o.pnpa_collection);
      // NPA
      var npaDem = numVal(o.npa_cases);
      var npaActAmt = parseFloat(o.npa_act_amt) || 0;
      var npaActAcct = numVal(o.npa_act_acc);

      data.push({
        name: o.name || '',
        empId: o.emp_id || '',
        branch: (o.branch_name || '').toUpperCase(),
        demand: dem,
        collection: col,
        balance: dem - col,
        pct: pct,
        dpd1Dem: dpd1Dem, dpd1Col: dpd1Col, dpd1Pct: dpd1Dem > 0 ? (dpd1Col / dpd1Dem) * 100 : 0,
        dpd2Dem: dpd2Dem, dpd2Col: dpd2Col, dpd2Pct: dpd2Dem > 0 ? (dpd2Col / dpd2Dem) * 100 : 0,
        pnpaDem: pnpaDem, pnpaCol: pnpaCol, pnpaPct: pnpaDem > 0 ? (pnpaCol / pnpaDem) * 100 : 0,
        npaDem: npaDem, npaCol: npaActAmt, npaPct: npaDem > 0 ? (npaActAmt / npaDem) * 100 : 0,
        npaActAcct: npaActAcct, npaActAmt: npaActAmt
      });
    }
    data.sort(function (a, b) { return a.pct - b.pct; });
    return data;
  }

  /* ── Fetch phone numbers from Supabase (kept as-is) ── */

  var _phoneCache = {};
  async function fetchPhones(empIds) {
    var toFetch = empIds.filter(function (id) { return !_phoneCache[id.toUpperCase()]; });
    if (!toFetch.length) return;
    try {
      var batchSize = 50;
      for (var b = 0; b < toFetch.length; b += batchSize) {
        var batch = toFetch.slice(b, b + batchSize);
        var orFilter = batch.map(function (id) { return 'emp_id.ilike.' + id; }).join(',');
        var url = SUPABASE_URL + '/rest/v1/employees?select=emp_id,mobile&or=(' + orFilter + ')';
        var controller = new AbortController();
        var timer = setTimeout(function () { controller.abort(); }, 8000);
        var resp = await fetch(url, { headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY }, signal: controller.signal });
        clearTimeout(timer);
        var data = await resp.json();
        if (Array.isArray(data)) { for (var i = 0; i < data.length; i++) { if (data[i].emp_id && data[i].mobile) _phoneCache[data[i].emp_id.toUpperCase()] = data[i].mobile; } }
      }
    } catch (e) { console.error('Phone fetch failed:', e); }
  }

  /* ── State ── */

  var _analState = {
    role: null,
    location: null,
    empId: null,
    activeView: null,
    expandedFO: null,
    foBucket: 'regular',
    foLoading: false,
    productType: 'IGL',
    // Cached API data per view
    branchData: null,
    foData: null
  };

  /* ── HTML builders ── */

  function buildCardHtml(item, rankNum, type, delay) {
    var pctColor = item.pct >= 95 ? '#34D399' : item.pct >= 80 ? '#FBBF24' : '#F87171';
    var barW = Math.min(item.pct, 100);
    var avatarClass = type === 'top' ? 'anal-avatar-top' : 'anal-avatar-bottom';
    return '<div class="anal-card desktop-analytical-card emp-fade" style="animation-delay:' + delay + 's">' +
      '<div class="anal-card-header">' +
      '<div class="anal-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>' +
      '<div class="anal-branch-avatar ' + avatarClass + '">' + item.name.charAt(0) + '</div>' +
      '<div class="anal-branch-name">' + esc(item.name) + '</div>' +
      '<div class="anal-pct" style="color:' + pctColor + '">' + item.pct.toFixed(2) + '%</div>' +
      '</div>' +
      '<div class="anal-card-stats">' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Demand</span><span class="anal-stat-val">' + fmtNum(item.demand) + '</span></div>' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Collection</span><span class="anal-stat-val" style="color:#059669">' + fmtNum(item.collection) + '</span></div>' +
      '<div class="anal-stat"><span class="anal-stat-lbl">Balance</span><span class="anal-stat-val" style="color:#FB923C">' + fmtNum(item.balance) + '</span></div>' +
      '</div>' +
      '<div class="pf-progress-track"><div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div></div>' +
      '</div>';
  }

  function buildSectionHtml(title, iconSvg, type, items) {
    var html = '<div class="anal-section"><div class="anal-section-header desktop-section-header">' +
      '<div class="anal-section-icon ' + type + '">' + iconSvg + '</div>' +
      '<div class="anal-section-title">' + title + '</div>' +
      '<div class="anal-section-count">' + items.length + '</div></div>';
    if (!items.length) { html += '<div class="anal-empty">No data available</div>'; }
    else { html += '<div class="anal-list">'; for (var j = 0; j < items.length; j++) html += buildCardHtml(items[j], j + 1, type, j * 0.04); html += '</div>'; }
    return html + '</div>';
  }

  function buildFOCardHtml(fo, rankNum, delay, expanded) {
    var pctColor = fo.pct >= 95 ? '#34D399' : fo.pct >= 80 ? '#FBBF24' : '#F87171';
    var barW = Math.min(fo.pct, 100);
    var phone = _phoneCache[fo.empId.toUpperCase()] || '';

    var html = '<div class="fo-card desktop-analytical-card emp-fade" data-fo-empid="' + esc(fo.empId) + '" data-fo-name="' + esc(fo.name) + '" style="animation-delay:' + delay + 's">';
    html += '<div class="fo-header">';
    html += '<div class="fo-rank" style="background:' + pctColor + '20;color:' + pctColor + '">#' + rankNum + '</div>';
    html += '<div class="fo-avatar">' + (fo.name.charAt(0) || '?').toUpperCase() + '</div>';
    html += '<div class="fo-info"><div class="fo-name">' + esc(fo.name) + '<span class="fo-branch-tag">' + esc(fo.branch) + '</span></div>';
    html += '<div class="fo-empid">' + esc(fo.empId) + '</div></div>';
    html += '<div class="fo-pct" style="color:' + pctColor + '">' + fo.pct.toFixed(2) + '%</div>';
    html += '</div>';

    html += '<div class="fo-stats">';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Demand</span><span class="anal-stat-val">' + fmtNum(fo.demand) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Collection</span><span class="anal-stat-val" style="color:#059669">' + fmtNum(fo.collection) + '</span></div>';
    html += '<div class="anal-stat"><span class="anal-stat-lbl">Balance</span><span class="anal-stat-val" style="color:#FB923C">' + fmtNum(fo.balance) + '</span></div>';
    html += '</div>';
    html += '<div class="pf-progress-track" style="margin-top:8px"><div class="pf-progress-fill" style="width:' + barW + '%;background:' + pctColor + '"></div></div>';

    // Action buttons — call + view dashboard
    html += '<div class="fo-actions">';
    if (phone) {
      html += '<a href="tel:' + esc(phone) + '" class="fo-call-btn" onclick="event.stopPropagation();">';
      html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>';
      html += 'Call</a>';
    } else if (_analState.foLoading) {
      html += '<div class="fo-no-phone" style="color:#059669;border-color:#ECFDF5"><svg width="12" height="12" viewBox="0 0 24 24" style="animation:spin .7s linear infinite"><circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="3" stroke-dasharray="30 70"/></svg> Fetching from database...</div>';
    } else {
      html += '<div class="fo-no-phone">No phone</div>';
    }
    html += '<div class="fo-view-btn" onclick="event.stopPropagation(); this.closest(\'.fo-card\').click();">';
    html += '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
    html += 'View Dashboard</div>';
    html += '</div>';

    html += '</div>';
    return html;
  }

  /* ── Render ── */

  function renderPage() {
    var container = document.getElementById('analyticalContent');
    if (!container) return;

    var upArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>';
    var downArrow = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/></svg>';

    var view = _analState.activeView;
    var role = _analState.role;
    var location = _analState.location;

    var html = '<div class="anal-container desktop-content-area">';

    // Buttons
    html += '<div class="anal-btn-row">';
    html += '<button class="anal-btn' + (view === 'branches' ? ' active' : '') + '" data-anal-view="branches">';
    html += '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>';
    html += 'Branches</button>';
    html += '<button class="anal-btn' + (view === 'fo' ? ' active' : '') + '" data-anal-view="fo">';
    html += '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
    html += 'Field Officer</button>';
    html += '</div>';

    // Content based on active view
    if (view === 'branches') {
      var bd = _analState.branchData || [];
      if (role === 'CEO') {
        html += '<div class="desktop-grid-2">';
        html += buildSectionHtml('Top 10 Branches', upArrow, 'top', bd.slice(-10).reverse());
        html += '<div class="anal-divider"></div>';
        html += buildSectionHtml('Bottom 10 Branches', downArrow, 'bottom', bd.slice(0, 10));
        html += '</div>';
      } else if (role === 'RM') {
        html += '<div class="desktop-grid-2">';
        html += buildSectionHtml('Top 10 Branches \u2014 ' + esc(location), upArrow, 'top', bd.slice(-10).reverse());
        html += '</div>';
      } else if (role === 'DM') {
        html += '<div class="desktop-grid-2">';
        html += buildSectionHtml('Top 5 Branches \u2014 ' + esc(location), upArrow, 'top', bd.slice(-5).reverse());
        html += '<div class="anal-divider"></div>';
        html += buildSectionHtml('Bottom 5 Branches \u2014 ' + esc(location), downArrow, 'bottom', bd.slice(0, 5));
        html += '</div>';
      }
    } else if (view === 'fo') {
      var bk = _analState.foBucket;
      var foData = _analState.foData || [];
      var lowPerformers = foData.slice(0, 20);
      var bucketLabel = bk === 'dpd1' ? 'SMA-0' : bk === 'dpd2' ? 'SMA-1' : bk === 'pnpa' ? 'SMA-2' : bk === 'npa' ? 'NPA' : 'Regular';
      var label = (role !== 'CEO' && location) ? esc(location) : '';

      // Bucket filter pills
      html += '<div class="anal-btn-row desktop-pills-row" style="margin-bottom:16px;flex-wrap:wrap">';
      html += '<button class="anal-btn' + (bk === 'regular' ? ' active' : '') + '" data-fo-bucket="regular" style="padding:8px 10px;font-size:12px">Regular</button>';
      html += '<button class="anal-btn' + (bk === 'dpd1' ? ' active' : '') + '" data-fo-bucket="dpd1" style="padding:8px 10px;font-size:12px">SMA-0</button>';
      html += '<button class="anal-btn' + (bk === 'dpd2' ? ' active' : '') + '" data-fo-bucket="dpd2" style="padding:8px 10px;font-size:12px">SMA-1</button>';
      html += '<button class="anal-btn' + (bk === 'pnpa' ? ' active' : '') + '" data-fo-bucket="pnpa" style="padding:8px 10px;font-size:12px">SMA-2</button>';
      html += '<button class="anal-btn' + (bk === 'npa' ? ' active' : '') + '" data-fo-bucket="npa" style="padding:8px 10px;font-size:12px">NPA</button>';
      html += '</div>';

      html += '<div class="anal-section"><div class="anal-section-header desktop-section-header">';
      html += '<div class="anal-section-icon bottom">' + downArrow + '</div>';
      html += '<div class="anal-section-title">Low Performing (' + bucketLabel + ')' + (label ? ' \u2014 ' + label : '') + '</div>';
      html += '<div class="anal-section-count">' + lowPerformers.length + '</div></div>';

      if (!lowPerformers.length) {
        html += '<div class="anal-empty">No field officer data available</div>';
      } else {
        html += '<div class="anal-list desktop-grid-3">';
        for (var f = 0; f < lowPerformers.length; f++) {
          var isExpanded = _analState.expandedFO === lowPerformers[f].empId;
          html += buildFOCardHtml(lowPerformers[f], f + 1, f * 0.04, isExpanded);
        }
        html += '</div>';
      }
      html += '</div>';
    }

    html += '</div>';
    container.innerHTML = html;
    attachHandlers();
  }

  /* ── Data loading via API ── */

  async function loadBranchData() {
    var role = _analState.role;
    var location = _analState.location;
    var pt = _analState.productType;
    var params = { product_type: pt };

    if (role === 'RM') {
      params.region = location;
    } else if (role === 'DM') {
      params.district = location;
    }
    // CEO: no filter — all branches

    var apiRows = await apiFetch(buildQuery('/api/collection/by-branch', params));
    _analState.branchData = computeBranchData(apiRows);
  }

  async function loadFOData(bucket) {
    var role = _analState.role;
    var location = _analState.location;
    var pt = _analState.productType;
    var params = { product_type: pt };

    if (role === 'BM') {
      params.branch = location;
    } else if (role === 'DM') {
      // Fetch branches for this district first, then fetch employees for each branch
      var districtBranches = await apiFetch(buildQuery('/api/collection/by-branch', { product_type: pt, district: location }));
      var allEmployees = [];
      for (var i = 0; i < districtBranches.length; i++) {
        var branchName = districtBranches[i].branch_name;
        if (branchName) {
          var emps = await apiFetch(buildQuery('/api/collection/by-employee', { product_type: pt, branch: branchName }));
          allEmployees = allEmployees.concat(emps);
        }
      }
      _analState.foData = computeFODataFromAPI(allEmployees, bucket);
      return;
    } else if (role === 'RM') {
      // Fetch branches for this region's districts, then employees
      var regionDistricts = await apiFetch(buildQuery('/api/collection/by-district', { product_type: pt, region: location }));
      var allEmps = [];
      for (var d = 0; d < regionDistricts.length; d++) {
        var distName = regionDistricts[d].district_name;
        if (distName) {
          var dBranches = await apiFetch(buildQuery('/api/collection/by-branch', { product_type: pt, district: distName }));
          for (var b = 0; b < dBranches.length; b++) {
            var bName = dBranches[b].branch_name;
            if (bName) {
              var bEmps = await apiFetch(buildQuery('/api/collection/by-employee', { product_type: pt, branch: bName }));
              allEmps = allEmps.concat(bEmps);
            }
          }
        }
      }
      _analState.foData = computeFODataFromAPI(allEmps, bucket);
      return;
    }
    // CEO or BM: direct fetch
    var apiRows = await apiFetch(buildQuery('/api/collection/by-employee', params));
    _analState.foData = computeFODataFromAPI(apiRows, bucket);
  }

  /* ── Event handlers ── */

  function attachHandlers() {
    var container = document.getElementById('analyticalContent');
    if (!container) return;

    // View toggle buttons
    container.querySelectorAll('[data-anal-view]').forEach(function (btn) {
      btn.onclick = async function () {
        var newView = btn.dataset.analView;
        _analState.activeView = newView;
        _analState.expandedFO = null;
        _analState.foBucket = 'regular';

        try {
          if (newView === 'branches') {
            await loadBranchData();
            renderPage();
          } else if (newView === 'fo') {
            _analState.foLoading = true;
            // Load FO data first, render with "Fetching..." for phones
            await loadFOData('regular');
            renderPage();
            // Now fetch phones asynchronously
            try {
              var empIds = (_analState.foData || []).map(function (o) { return o.empId; });
              await fetchPhones(empIds);
            } catch (e) { console.error('FO phone load error:', e); }
            _analState.foLoading = false;
            renderPage();
          }
        } catch (e) {
          console.error('Data load error:', e);
          renderPage();
        }
      };
    });

    // Bucket filter pills
    container.querySelectorAll('[data-fo-bucket]').forEach(function (btn) {
      btn.onclick = async function () {
        var newBucket = btn.dataset.foBucket;
        _analState.foBucket = newBucket;
        _analState.expandedFO = null;
        _analState.foLoading = false;
        try {
          await loadFOData(newBucket);
        } catch (e) { console.error('FO bucket load error:', e); }
        renderPage();
      };
    });

    // FO card click — navigate to FO's full dashboard
    container.querySelectorAll('[data-fo-empid]').forEach(function (card) {
      card.onclick = function () {
        var empId = card.dataset.foEmpid;
        var empName = card.dataset.foName || '';
        // Save current session to nav stack
        var stack = typeof getRoleNavStack === 'function' ? getRoleNavStack() : [];
        var current = typeof getEmployeeSession === 'function' ? getEmployeeSession() : null;
        if (current) {
          stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
          localStorage.setItem('roleNavStack', JSON.stringify(stack));
        }
        // Switch to FO session
        localStorage.removeItem('roleAuth');
        localStorage.removeItem('roleName');
        localStorage.removeItem('roleLocation');
        localStorage.setItem('employeeId', empId);
        localStorage.setItem('employeeName', empName);
        // Flag to return to analytical tab
        localStorage.setItem('returnToAnalytical', 'true');
        window.location.reload();
      };
    });
  }

  /* ── Entry point ── */

  window._loadAnalyticalTab = async function () {
    try {
      var session = typeof getEmployeeSession === 'function' ? getEmployeeSession() : null;
      if (!session) return;
      var role = session.role;

      // Always show analytical tool in sidebar for all roles
      var navItem2 = document.getElementById('analyticalNavItem');
      if (navItem2) navItem2.style.display = '';

      _analState.role = role;
      _analState.location = session.location;
      _analState.empId = session.id;

      // Always start analytical tool fresh — no auto-restoring previous view
      _analState.activeView = null;
      _analState.foBucket = 'regular';
      _analState.expandedFO = null;
      _analState.branchData = null;
      _analState.foData = null;

      renderPage();
    } catch (err) {
      console.error('Analytical load failed:', err);
    }
  };
})();
