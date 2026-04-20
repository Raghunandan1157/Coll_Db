/**
 * analytical.js — Analytical Tool tab.
 * Lowest Performers panel (FTD/MTD) for CEO + RM/SM.
 * Reuses _collState.date and availableDates from collection.js.
 * Container: analyticalContent
 */
(function () {
  var session = getEmployeeSession();

  var PRIORITY_CONFIG = {
    CEO: [
      { label: 'Branches',  role: 'BM', n: 5, url: '/api/daily/by-branch',   nameField: 'branch_name'   },
      { label: 'Areas',     role: 'AM', n: 3, url: '/api/daily/by-area',     nameField: 'area_name'     },
      { label: 'Divisions', role: 'DM', n: 2, url: '/api/daily/by-division', nameField: 'division_name' },
      { label: 'Regions',   role: 'RM', n: 1, url: '/api/daily/by-state',    nameField: 'state_name'    }
    ],
    RM: [
      { label: 'Officers',  role: 'FO', n: 5, url: '/api/daily/by-employee', nameField: 'name'          },
      { label: 'Branches',  role: 'BM', n: 3, url: '/api/daily/by-branch',   nameField: 'branch_name'   },
      { label: 'Areas',     role: 'AM', n: 2, url: '/api/daily/by-area',     nameField: 'area_name'     },
      { label: 'Divisions', role: 'DM', n: 1, url: '/api/daily/by-division', nameField: 'division_name' }
    ]
  };
  PRIORITY_CONFIG.SM = PRIORITY_CONFIG.RM;

  function cfgForRole(role) { return PRIORITY_CONFIG[role] || null; }
  function isPriorityRole(role) { return !!cfgForRole(role); }

  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }
  function apiFetch(url) { return fetch(url).then(function (r) { if (!r.ok) throw new Error('API ' + r.status); return r.json(); }); }

  function getState() {
    var saved = localStorage.getItem('analyticalScope');
    return {
      date: localStorage.getItem('collSelectedDate') || (window._collState && window._collState.date) || null,
      scope: (saved === 'mtd') ? 'mtd' : 'ftd',
      availableDates: (window._collState && window._collState.availableDates) || []
    };
  }

  function rangeParams(s) {
    if (!s.date) return null;
    if (s.scope === 'mtd') {
      var mStart = s.date.slice(0, 8) + '01';
      return 'date_from=' + encodeURIComponent(mStart) + '&date_to=' + encodeURIComponent(s.date);
    }
    return 'date=' + encodeURIComponent(s.date);
  }

  function scopeFilter() {
    var p = ['scope=oa'];
    if ((session.role === 'RM' || session.role === 'SM') && session.location) {
      p.push('state=' + encodeURIComponent(session.location));
    }
    return p.join('&');
  }

  function rowHtml(item, c) {
    var dem = numVal(item.regular_demand);
    var col = numVal(item.regular_collection);
    var pct = dem > 0 ? (col / dem * 100) : 0;
    var pctColor = pct >= 90 ? '#059669' : pct >= 75 ? '#F59E0B' : '#EF4444';
    var name = item[c.nameField] || item.name || '—';
    var sub = '';
    if (c.role === 'BM') sub = item.district_name || '';
    else if (c.role === 'AM') sub = item.division_name || '';
    else if (c.role === 'DM') sub = item.region_name || '';
    else if (c.role === 'FO') sub = item.branch_name || '';

    var dataAttrs = '';
    if (c.role === 'FO' && item.emp_id) {
      dataAttrs = ' data-emp-id="' + esc(item.emp_id) + '" data-emp-name="' + esc(name) + '"';
    } else {
      dataAttrs = ' data-child-role="' + c.role + '" data-child-location="' + esc(name) + '"';
    }

    return '<div class="emp-sub-card analytical-row"' + dataAttrs + ' style="background:#fff;border-radius:8px;padding:10px 12px;display:flex;justify-content:space-between;align-items:center;gap:8px;cursor:pointer;border:1px solid #F1F5F9;transition:all .15s;">' +
      '<div style="min-width:0;flex:1;">' +
        '<div style="font-size:13px;font-weight:700;color:#1E293B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="' + esc(name) + '">' + esc(name) + '</div>' +
        (sub ? '<div style="font-size:11px;color:#94A3B8;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + esc(sub) + '</div>' : '') +
      '</div>' +
      '<div style="font-size:14px;font-weight:700;color:' + pctColor + ';flex-shrink:0;">' + pct.toFixed(1) + '%</div>' +
    '</div>';
  }

  function shellHtml(s, cfg) {
    var html = '<div class="desktop-content-area">';

    // Header card
    html += '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:16px 20px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">';
    html += '<div>';
    html += '<div style="font-size:16px;font-weight:700;color:#1E293B;display:flex;align-items:center;gap:8px;">';
    html += '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#F87171;"></span>Lowest Performers';
    html += '</div>';
    html += '<div style="font-size:11px;color:#94A3B8;margin-top:4px;">' + (s.date ? esc(s.date) : 'No date selected') + ' · ' + (s.scope === 'mtd' ? 'Month-To-Date' : 'For The Day') + '</div>';
    html += '</div>';
    html += '<div style="display:inline-flex;background:#F1F5F9;border-radius:18px;padding:2px;">';
    html += '<button data-an-scope="ftd" style="padding:6px 16px;border-radius:16px;font-size:12px;font-weight:700;cursor:pointer;border:none;letter-spacing:.04em;' + (s.scope === 'ftd' ? 'background:#059669;color:#fff;' : 'background:transparent;color:#64748B;') + '">FTD</button>';
    html += '<button data-an-scope="mtd" style="padding:6px 16px;border-radius:16px;font-size:12px;font-weight:700;cursor:pointer;border:none;letter-spacing:.04em;' + (s.scope === 'mtd' ? 'background:#059669;color:#fff;' : 'background:transparent;color:#64748B;') + '">MTD</button>';
    html += '</div></div>';

    if (!s.date) {
      html += '<div style="text-align:center;padding:60px 20px;color:#64748B;font-size:14px;">Select a date in the Collection tab first.</div>';
      html += '</div>';
      return html;
    }

    html += '<div id="analyticalSlots" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px;">';
    for (var i = 0; i < cfg.length; i++) {
      var c = cfg[i];
      html += '<div data-an-slot="' + c.role + '" style="background:#F8FAFC;border-radius:12px;padding:12px 14px;min-height:120px;">';
      html += '<div style="font-size:12px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.05em;margin-bottom:10px;">';
      html += esc(c.label) + ' <span style="color:#94A3B8;">(lowest ' + c.n + ')</span>';
      html += '</div>';
      html += '<div class="an-rows" style="display:flex;flex-direction:column;gap:6px;">';
      html += '<div style="height:50px;background:#fff;border-radius:6px;animation:pulse 1.4s ease-in-out infinite;"></div>';
      html += '</div></div>';
    }
    html += '</div>';
    html += '</div>';
    return html;
  }

  function loadData(s, cfg, container) {
    var rangeQ = rangeParams(s);
    var scopeQ = scopeFilter();
    cfg.forEach(function (c) {
      var url = c.url + '?' + rangeQ + '&' + scopeQ;
      apiFetch(url).then(function (rows) {
        rows = (rows || []).filter(function (r) { return numVal(r.regular_demand) > 0; });
        rows.sort(function (a, b) {
          var pa = numVal(a.regular_collection) / numVal(a.regular_demand);
          var pb = numVal(b.regular_collection) / numVal(b.regular_demand);
          return pa - pb;
        });
        var worst = rows.slice(0, c.n);
        var slot = container.querySelector('[data-an-slot="' + c.role + '"] .an-rows');
        if (!slot) return;
        if (!worst.length) {
          slot.innerHTML = '<div style="font-size:12px;color:#94A3B8;text-align:center;padding:14px 0;">No data</div>';
          return;
        }
        slot.innerHTML = worst.map(function (it) { return rowHtml(it, c); }).join('');
      }).catch(function (err) {
        var slot = container.querySelector('[data-an-slot="' + c.role + '"] .an-rows');
        if (slot) slot.innerHTML = '<div style="font-size:12px;color:#EF4444;text-align:center;padding:14px 0;">Load failed</div>';
        console.warn('Analytical fetch failed:', c.url, err);
      });
    });
  }

  function attachHandlers(container) {
    container.querySelectorAll('[data-an-scope]').forEach(function (btn) {
      btn.onclick = function () {
        var next = btn.dataset.anScope;
        if (next !== 'ftd' && next !== 'mtd') return;
        localStorage.setItem('analyticalScope', next);
        window._loadAnalyticalTab();
      };
    });

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

  window._loadAnalyticalTab = function () {
    var container = document.getElementById('analyticalContent');
    if (!container) return;
    if (!isPriorityRole(session.role)) {
      container.innerHTML = '<div class="desktop-content-area"><div style="text-align:center;padding:60px 20px;color:#64748B;font-size:14px;">Analytical Tool is available for CEO and RM roles only.</div></div>';
      return;
    }
    var s = getState();
    var cfg = cfgForRole(session.role);
    container.innerHTML = shellHtml(s, cfg);
    attachHandlers(container);
    if (s.date) loadData(s, cfg, container);
  };
})();
