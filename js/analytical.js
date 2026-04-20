/**
 * analytical.js — Analytical Tool tab.
 * Lowest 10% (MTD only) per role; Disbursement | Collection filter; bucket sub-tabs for Collection.
 * Container: analyticalContent
 */
(function () {
  var session = getEmployeeSession();

  var ROLE_VIS = {
    CEO: ['FO', 'BM', 'AM', 'DM', 'RM'],
    RM:  ['FO', 'BM', 'AM', 'DM'],
    SM:  ['FO', 'BM', 'AM', 'DM'],
    DM:  ['FO', 'BM', 'AM'],
    DvM: ['FO', 'BM', 'AM'],
    AM:  ['FO', 'BM'],
    BM:  ['FO']
  };

  var ROLE_CFG = {
    FO: { collUrl: '/api/daily/by-employee',           disbUrl: '/api/disbursement/daily/by-employee', nameField: 'name',          subField: 'branch_name'   },
    BM: { collUrl: '/api/daily/by-branch',             disbUrl: '/api/disbursement/daily/by-branch',   nameField: 'branch_name',   subField: 'district_name' },
    AM: { collUrl: '/api/daily/by-area',               disbUrl: '/api/disbursement/daily/by-area',     nameField: 'area_name',     subField: 'division_name' },
    DM: { collUrl: '/api/daily/by-division',           disbUrl: '/api/disbursement/daily/by-division', nameField: 'division_name', subField: 'region_name'   },
    RM: { collUrl: '/api/daily/by-state',              disbUrl: '/api/disbursement/daily/by-region',   nameField: 'state_name',    subField: null            }
  };

  var BUCKETS = [
    { key: 'regular', label: 'Regular', d: 'regular_demand', c: 'regular_collection' },
    { key: '1-30',    label: '1-30',    d: 'demand_1_30',    c: 'collection_1_30'    },
    { key: '31-60',   label: '31-60',   d: 'demand_31_60',   c: 'collection_31_60'   },
    { key: 'pnpa',    label: 'PNPA',    d: 'pnpa_demand',    c: 'pnpa_collection'    },
    { key: 'npa',     label: 'NPA',     d: 'npa_cases',      c: 'npa_clo_acc'        }
  ];

  /* ===== Helpers ===== */
  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }
  function apiFetch(url) { return fetch(url).then(function (r) { if (!r.ok) throw new Error('API ' + r.status); return r.json(); }); }
  function fmtAmt(n) {
    n = numVal(n);
    if (n >= 10000000) return '₹' + (n / 10000000).toFixed(2) + ' Cr';
    if (n >= 100000)   return '₹' + (n / 100000).toFixed(2) + ' L';
    if (n >= 1000)     return '₹' + (n / 1000).toFixed(1) + ' K';
    return '₹' + Math.round(n).toLocaleString('en-IN');
  }

  function lowestN(count) {
    if (!count || count <= 0) return 0;
    var v = count * 0.1;
    return v < 1 ? 1 : Math.round(v);
  }

  function loadState() {
    var raw = {};
    try { raw = JSON.parse(localStorage.getItem('analyticalState') || '{}'); } catch (e) {}
    var visible = ROLE_VIS[session.role] || [];
    var mode = (raw.mode === 'disbursement') ? 'disbursement' : 'collection';
    var bucket = BUCKETS.map(function (b) { return b.key; }).indexOf(raw.bucket) >= 0 ? raw.bucket : 'regular';
    var role = visible.indexOf(raw.role) >= 0 ? raw.role : (visible[0] || null);
    return { mode: mode, bucket: bucket, role: role };
  }
  function saveState(s) { localStorage.setItem('analyticalState', JSON.stringify(s)); }

  function mtdRange() {
    var d = (window._collState && window._collState.date) || localStorage.getItem('collSelectedDate');
    if (!d) {
      var now = new Date();
      var y = now.getFullYear();
      var m = String(now.getMonth() + 1).padStart(2, '0');
      var dd = String(now.getDate()).padStart(2, '0');
      d = y + '-' + m + '-' + dd;
    }
    return { from: d.slice(0, 8) + '01', to: d };
  }

  function scopeParams() {
    var p = [];
    var r = session.role;
    if ((r === 'RM' || r === 'SM') && session.location) p.push('state=' + encodeURIComponent(session.location));
    else if ((r === 'DM' || r === 'DvM') && session.location) p.push('division=' + encodeURIComponent(session.location));
    else if (r === 'AM' && session.location) p.push('area=' + encodeURIComponent(session.location));
    else if (r === 'BM' && session.location) p.push('branch=' + encodeURIComponent(session.location));
    return p;
  }

  /* ===== Render ===== */
  function shellHtml(state, counts, range) {
    var visible = ROLE_VIS[session.role] || [];
    var html = '<div class="desktop-content-area">';

    html += '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:16px 20px;margin-bottom:14px;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">';
    html += '<div>';
    html += '<div style="font-size:16px;font-weight:700;color:#1E293B;">Analytical Tool — Lowest 10% (MTD)</div>';
    html += '<div style="font-size:11px;color:#94A3B8;margin-top:4px;">' + esc(range.from) + ' → ' + esc(range.to) + '</div>';
    html += '</div>';
    html += '<div style="display:inline-flex;background:#F1F5F9;border-radius:18px;padding:2px;">';
    html += '<button data-an-mode="collection" style="padding:6px 16px;border-radius:16px;font-size:12px;font-weight:700;cursor:pointer;border:none;letter-spacing:.04em;' + (state.mode === 'collection' ? 'background:#059669;color:#fff;' : 'background:transparent;color:#64748B;') + '">Collection</button>';
    html += '<button data-an-mode="disbursement" style="padding:6px 16px;border-radius:16px;font-size:12px;font-weight:700;cursor:pointer;border:none;letter-spacing:.04em;' + (state.mode === 'disbursement' ? 'background:#F59E0B;color:#fff;' : 'background:transparent;color:#64748B;') + '">Disbursement</button>';
    html += '</div>';
    html += '</div>';

    if (state.mode === 'collection') {
      html += '<div style="margin-top:12px;display:flex;gap:6px;flex-wrap:wrap;">';
      for (var bi = 0; bi < BUCKETS.length; bi++) {
        var b = BUCKETS[bi];
        var ba = state.bucket === b.key;
        html += '<button data-an-bucket="' + b.key + '" style="padding:5px 14px;border-radius:14px;font-size:11px;font-weight:700;cursor:pointer;border:none;letter-spacing:.04em;' + (ba ? 'background:#1E293B;color:#fff;' : 'background:#F1F5F9;color:#64748B;') + '">' + esc(b.label) + '</button>';
      }
      html += '</div>';
    }

    html += '<div style="margin-top:12px;display:flex;gap:6px;flex-wrap:wrap;">';
    if (visible.length === 0) {
      html += '<div style="font-size:13px;color:#94A3B8;">Analytical Tool is not available for your role.</div>';
    }
    for (var ri = 0; ri < visible.length; ri++) {
      var rk = visible[ri];
      var n = lowestN(counts[rk] || 0);
      var ra = state.role === rk;
      var disabled = n === 0;
      html += '<button data-an-role="' + rk + '"' + (disabled ? ' disabled' : '') + ' style="padding:7px 14px;border-radius:18px;font-size:12px;font-weight:700;cursor:' + (disabled ? 'not-allowed' : 'pointer') + ';border:none;letter-spacing:.04em;opacity:' + (disabled ? '.45' : '1') + ';' + (ra ? 'background:#6366F1;color:#fff;' : 'background:#F1F5F9;color:#1E293B;') + '">' + esc(rk) + ' <span style="opacity:.7;font-weight:600;">(' + n + ')</span></button>';
    }
    html += '</div>';
    html += '</div>';

    html += '<div id="anList" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:14px 16px;min-height:200px;">';
    html += '<div style="text-align:center;padding:60px 20px;color:#94A3B8;font-size:13px;">Loading…</div>';
    html += '</div>';

    html += '</div>';
    return html;
  }

  function rowHtml(item, cfg, state, idx) {
    var name = item[cfg.nameField] || item.name || item.region_name || item.state_name || '—';
    var sub = cfg.subField ? (item[cfg.subField] || '') : '';

    var leftMetric = '';
    var rightMetric = '';
    if (state.mode === 'collection') {
      var bucket = BUCKETS.filter(function (b) { return b.key === state.bucket; })[0];
      var dem = numVal(item[bucket.d]);
      var col = numVal(item[bucket.c]);
      var pct = dem > 0 ? (col / dem * 100) : 0;
      var pctColor = pct >= 90 ? '#059669' : pct >= 75 ? '#F59E0B' : '#EF4444';
      var lblD = bucket.key === 'npa' ? 'Cases' : 'D';
      var lblC = bucket.key === 'npa' ? 'Closed' : 'C';
      leftMetric = '<span style="font-size:11px;color:#64748B;">' + lblD + ' ' + Math.round(dem).toLocaleString('en-IN') + ' / ' + lblC + ' ' + Math.round(col).toLocaleString('en-IN') + '</span>';
      rightMetric = '<span style="font-size:14px;font-weight:700;color:' + pctColor + ';">' + pct.toFixed(1) + '%</span>';
    } else {
      var amt = numVal(item.total_amount);
      var cnt = numVal(item.total_count);
      leftMetric = '<span style="font-size:11px;color:#64748B;">' + cnt + ' acc</span>';
      rightMetric = '<span style="font-size:14px;font-weight:700;color:#F59E0B;">' + fmtAmt(amt) + '</span>';
    }

    var dataAttrs = '';
    if (cfg === ROLE_CFG.FO && item.emp_id) {
      dataAttrs = ' data-emp-id="' + esc(item.emp_id) + '" data-emp-name="' + esc(name) + '"';
    } else if (cfg !== ROLE_CFG.FO) {
      dataAttrs = ' data-child-role="' + (state.role) + '" data-child-location="' + esc(name) + '"';
    }

    return '<div class="emp-sub-card analytical-row"' + dataAttrs + ' style="background:#fff;border-radius:8px;padding:10px 12px;display:flex;justify-content:space-between;align-items:center;gap:10px;cursor:pointer;border:1px solid #F1F5F9;margin-bottom:6px;transition:background .15s;">' +
      '<div style="display:flex;align-items:center;gap:10px;min-width:0;flex:1;">' +
        '<span style="font-size:11px;font-weight:700;color:#94A3B8;min-width:18px;text-align:right;">' + (idx + 1) + '.</span>' +
        '<div style="min-width:0;">' +
          '<div style="font-size:13px;font-weight:700;color:#1E293B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="' + esc(name) + '">' + esc(name) + '</div>' +
          '<div style="font-size:11px;color:#94A3B8;display:flex;gap:8px;align-items:center;">' +
            (sub ? '<span>' + esc(sub) + '</span>' : '') +
            leftMetric +
          '</div>' +
        '</div>' +
      '</div>' +
      '<div style="flex-shrink:0;">' + rightMetric + '</div>' +
    '</div>';
  }

  function loadList(state, range, container) {
    var listEl = container.querySelector('#anList');
    if (!listEl) return;
    if (!state.role) {
      listEl.innerHTML = '<div style="text-align:center;padding:60px 20px;color:#94A3B8;font-size:13px;">No role selected.</div>';
      return;
    }

    var cfg = ROLE_CFG[state.role];
    var url = state.mode === 'collection' ? cfg.collUrl : cfg.disbUrl;
    var qp = ['date_from=' + encodeURIComponent(range.from), 'date_to=' + encodeURIComponent(range.to)].concat(scopeParams());
    if (state.mode === 'collection') qp.push('scope=oa');

    apiFetch(url + '?' + qp.join('&')).then(function (rows) {
      rows = rows || [];
      if (state.mode === 'collection') {
        var bucket = BUCKETS.filter(function (b) { return b.key === state.bucket; })[0];
        rows = rows.filter(function (r) { return numVal(r[bucket.d]) > 0; });
        rows.sort(function (a, b) {
          var pa = numVal(a[bucket.c]) / numVal(a[bucket.d]);
          var pb = numVal(b[bucket.c]) / numVal(b[bucket.d]);
          return pa - pb;
        });
      } else {
        rows = rows.filter(function (r) { return numVal(r.total_amount) > 0 || numVal(r.total_count) > 0; });
        rows.sort(function (a, b) { return numVal(a.total_amount) - numVal(b.total_amount); });
      }
      var n = lowestN(state._workingCount || 0);
      if (n === 0) {
        listEl.innerHTML = '<div style="text-align:center;padding:60px 20px;color:#94A3B8;font-size:13px;">No staff in this role.</div>';
        return;
      }
      var slice = rows.slice(0, n);
      if (!slice.length) {
        listEl.innerHTML = '<div style="text-align:center;padding:60px 20px;color:#94A3B8;font-size:13px;">No data for this selection.</div>';
        return;
      }
      var html = '<div style="font-size:11px;color:#94A3B8;font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:10px;">Showing lowest ' + slice.length + ' of ' + (state._workingCount || 0) + ' ' + state.role + ' staff (10%)</div>';
      var cfgRef = cfg;
      html += slice.map(function (it, i) { return rowHtml(it, cfgRef, state, i); }).join('');
      listEl.innerHTML = html;
      attachRowHandlers(container);
    }).catch(function (err) {
      console.warn('Analytical list fetch failed:', url, err);
      listEl.innerHTML = '<div style="text-align:center;padding:60px 20px;color:#EF4444;font-size:13px;">Failed to load.</div>';
    });
  }

  function attachShellHandlers(container, state, counts, range) {
    container.querySelectorAll('[data-an-mode]').forEach(function (btn) {
      btn.onclick = function () {
        state.mode = btn.dataset.anMode;
        saveState({ mode: state.mode, bucket: state.bucket, role: state.role });
        renderTab();
      };
    });
    container.querySelectorAll('[data-an-bucket]').forEach(function (btn) {
      btn.onclick = function () {
        state.bucket = btn.dataset.anBucket;
        saveState({ mode: state.mode, bucket: state.bucket, role: state.role });
        renderTab();
      };
    });
    container.querySelectorAll('[data-an-role]').forEach(function (btn) {
      btn.onclick = function () {
        if (btn.disabled) return;
        state.role = btn.dataset.anRole;
        saveState({ mode: state.mode, bucket: state.bucket, role: state.role });
        renderTab();
      };
    });
  }

  function attachRowHandlers(container) {
    container.querySelectorAll('.analytical-row').forEach(function (card) {
      card.onclick = function () {
        if (card.dataset.childRole) {
          if (typeof pushRoleNav === 'function') {
            pushRoleNav(card.dataset.childRole, card.dataset.childLocation);
            window.location.reload();
          }
        } else if (card.dataset.empId) {
          if (typeof getRoleNavStack === 'function') {
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
        }
      };
    });
  }

  function renderTab() {
    var container = document.getElementById('analyticalContent');
    if (!container) return;
    var visible = ROLE_VIS[session.role] || [];
    if (visible.length === 0) {
      container.innerHTML = '<div class="desktop-content-area"><div style="text-align:center;padding:60px 20px;color:#64748B;font-size:14px;">Analytical Tool is not available for your role.</div></div>';
      return;
    }
    var state = loadState();
    var range = mtdRange();

    var qs = scopeParams();
    var countsUrl = '/api/employees/working-counts' + (qs.length ? '?' + qs.join('&') : '');

    apiFetch(countsUrl).then(function (counts) {
      counts = counts || {};
      // If currently selected role has 0 working count, fall back to first non-zero.
      if (!counts[state.role] || counts[state.role] === 0) {
        for (var i = 0; i < visible.length; i++) {
          if (counts[visible[i]] && counts[visible[i]] > 0) { state.role = visible[i]; break; }
        }
        saveState({ mode: state.mode, bucket: state.bucket, role: state.role });
      }
      state._workingCount = counts[state.role] || 0;
      container.innerHTML = shellHtml(state, counts, range);
      attachShellHandlers(container, state, counts, range);
      loadList(state, range, container);
    }).catch(function (err) {
      console.warn('Working counts failed:', err);
      container.innerHTML = '<div class="desktop-content-area"><div style="text-align:center;padding:60px 20px;color:#EF4444;font-size:14px;">Failed to load role counts.</div></div>';
    });
  }

  window._loadAnalyticalTab = renderTab;
})();
