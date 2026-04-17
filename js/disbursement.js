/**
 * disbursement.js — Disbursement tab renderer (API-based).
 * Data source: /api/disbursement/* (PostgreSQL)
 * Container: disbursementContent
 */
(function () {
  var session = getEmployeeSession();

  /* Employee name map */
  var _empNames = window._empNameMap || {};
  function getFullName(empId, fallback) { return _empNames[empId] || fallback || empId || ''; }

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
    if (n >= 1000) return '\u20B9' + (n / 1000).toFixed(1) + ' K';
    return '\u20B9' + n.toLocaleString('en-IN');
  }
  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function fmtATS(v) {
    var n = numVal(v);
    if (n === 0) return '-';
    return '\u20B9' + Math.round(n).toLocaleString('en-IN');
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, function(c) { return '&#' + c.charCodeAt(0) + ';'; }); }
  function apiFetch(url) { return fetch(url).then(function(r) { if (!r.ok) throw new Error('API ' + r.status); return r.json(); }); }

  var _db = window._db = {
    months: [],
    month: localStorage.getItem('dbMonth') || '',
    date: localStorage.getItem('dbSelectedDate') || null,
    availableDates: [],
    product: localStorage.getItem('dbProduct') || 'all',
    summary: null,
    children: [],
    byMonth: []
  };

  function dbBase() { return _db.date ? '/api/disbursement/daily' : '/api/disbursement'; }

  function queryParams() {
    var p = [];
    if (_db.date) p.push('date=' + encodeURIComponent(_db.date));
    else if (_db.month) p.push('month=' + encodeURIComponent(_db.month));
    if (_db.product && _db.product !== 'all') p.push('product_name=' + encodeURIComponent(_db.product.toUpperCase()));
    // Role-based filtering — server resolves via employee_master
    if ((session.role === 'RM' || session.role === 'SM') && session.location) p.push('region=' + encodeURIComponent(session.location));
    else if ((session.role === 'DM' || session.role === 'DvM') && session.location) p.push('division=' + encodeURIComponent(session.location));
    else if (session.role === 'AM' && session.location) p.push('area=' + encodeURIComponent(session.location));
    else if (session.role === 'BM' && session.location) p.push('branch=' + encodeURIComponent(session.location));
    else if ((!session.role || session.role === 'FO') && session.id) p.push('emp_id=' + encodeURIComponent(session.id));
    return p.length ? '?' + p.join('&') : '';
  }

  function childrenUrl() {
    var base = [];
    if (_db.date) base.push('date=' + encodeURIComponent(_db.date));
    else if (_db.month) base.push('month=' + encodeURIComponent(_db.month));
    if (_db.product && _db.product !== 'all') base.push('product_name=' + encodeURIComponent(_db.product.toUpperCase()));

    var root = dbBase();
    if (session.role === 'CEO') return root + '/by-region' + (base.length ? '?' + base.join('&') : '');
    if ((session.role === 'RM' || session.role === 'SM') && session.location) {
      base.push('region=' + encodeURIComponent(session.location));
      return root + '/by-district?' + base.join('&');
    }
    if ((session.role === 'DM' || session.role === 'DvM') && session.location) {
      base.push('division=' + encodeURIComponent(session.location));
      return root + '/by-branch?' + base.join('&');
    }
    if (session.role === 'AM' && session.location) {
      base.push('area=' + encodeURIComponent(session.location));
      return root + '/by-branch?' + base.join('&');
    }
    if (session.role === 'BM' && session.location) {
      base.push('branch=' + encodeURIComponent(session.location));
      return root + '/by-employee?' + base.join('&');
    }
    return null;
  }

  window.loadDisbursement = function() { loadAndRender(); };
  function loadAndRender() {
    var container = document.getElementById('disbursementContent');
    if (!container) return;
    container.innerHTML = '<div style="text-align:center;padding:80px 20px;"><div style="width:32px;height:32px;border:3px solid #E2E8F0;border-top-color:#F59E0B;border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 12px;"></div><div style="color:#64748B;font-size:14px;">Loading disbursement data...</div></div>';

    // Fetch availableDates once (needed for keyboard nav + first-load default).
    var initPromise;
    if (!_db.availableDates || !_db.availableDates.length) {
      initPromise = apiFetch('/api/disbursement/daily/dates').then(function(data) {
        var list = Array.isArray(data) ? data.map(function(x) { return String(x.disb_date || x.date || x).slice(0, 10); }) : [];
        _db.availableDates = list;
        // First-load default: pick newest if neither date nor month selected yet.
        if (!_db.date && !_db.month && list.length) {
          _db.date = list[0];
          localStorage.setItem('dbSelectedDate', list[0]);
        }
      }).catch(function() { _db.availableDates = _db.availableDates || []; });
    } else {
      initPromise = Promise.resolve();
    }

    initPromise.then(function() { runLoad(); });
  }

  function runLoad() {
    var container = document.getElementById('disbursementContent');
    if (!container) return;
    var monthPromise;
    if (_db.date) {
      monthPromise = Promise.resolve();
    } else if (_db.months.length) {
      monthPromise = Promise.resolve();
    } else {
      monthPromise = apiFetch('/api/disbursement/months').then(function(m) {
        _db.months = m || [];
        if (!_db.month && m.length) {
          // Pick latest month that has data
          var cidx = m.length - 1;
          function tryDbMonth(idx) {
            if (idx < 0) { _db.month = m[m.length - 1]; localStorage.setItem('dbMonth', _db.month); return Promise.resolve(); }
            var label = m[idx];
            return apiFetch('/api/disbursement/summary?month=' + encodeURIComponent(label)).then(function(d) {
              var hasData = d && ((d.total_amount && parseFloat(d.total_amount) > 0) || (d.total_accounts && parseInt(d.total_accounts) > 0));
              if (hasData) { _db.month = label; localStorage.setItem('dbMonth', label); return; }
              return tryDbMonth(idx - 1);
            }).catch(function() { _db.month = label; localStorage.setItem('dbMonth', label); });
          }
          return tryDbMonth(cidx);
        }
      });
    }

    monthPromise.then(function() {
      var base = dbBase();
      var promises = [
        apiFetch(base + '/summary' + queryParams()),
        _db.date ? Promise.resolve([]) : apiFetch(base + '/by-month' + queryParams()),
      ];
      var chUrl = childrenUrl();
      if (chUrl && session.role && session.role !== 'FO') promises.push(apiFetch(chUrl));
      else promises.push(Promise.resolve([]));

      // Product breakdown
      promises.push(apiFetch(base + '/by-product' + queryParams()));

      return Promise.all(promises);
    }).then(function(res) {
      _db.summary = res[0];
      _db.byMonth = res[1] || [];
      _db.children = res[2] || [];
      var byProduct = res[3] || [];
      _db.byProduct = byProduct;
      render(byProduct);
    }).catch(function(err) {
      console.error('Disbursement load failed:', err);
      container.innerHTML = '<div style="text-align:center;padding:80px 20px;color:#64748B;">Failed to load disbursement data.</div>';
    });
  }

  function getDbSubunitView() {
    return localStorage.getItem('subunitView') || 'card';
  }

  function dbSubunitToggleHtml() {
    var v = getDbSubunitView();
    return '<div class="subunit-view-toggle">' +
      '<button class="subunit-view-btn' + (v === 'card' ? ' active' : '') + '" data-subview="card">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>' +
        'Cards</button>' +
      '<button class="subunit-view-btn' + (v === 'table' ? ' active' : '') + '" data-subview="table">' +
        '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>' +
        'Table</button>' +
    '</div>';
  }

  function dbSubUnitsHtml(children, childRole) {
    if (!children.length) return '';
    var roleLabel = { RM: 'Regions', DM: 'Divisions', BM: 'Branches', FO: 'Officers' };
    var label = roleLabel[childRole] || 'Team';
    var isTable = getDbSubunitView() === 'table';
    var avStyle = 'background:#FFFBEB;color:#F59E0B;';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label + '<span class="emp-team-count">' + children.length + '</span>' + dbSubunitToggleHtml() + '</div>';

    if (isTable) {
      html += '<div class="subunit-table-wrap"><table class="subunit-table">';
      html += '<thead><tr>' +
        '<th>Unit <span class="sort-icon">&#8597;</span></th>' +
        '<th class="num-col">Accounts <span class="sort-icon">&#8597;</span></th>' +
        '<th class="num-col">Amount <span class="sort-icon">&#8597;</span></th>' +
        '<th style="width:24px;"></th>' +
      '</tr></thead><tbody>';

      for (var i = 0; i < children.length; i++) {
        var ch = children[i];
        var childName = '';
        if (childRole === 'RM') childName = ch.region_name || '';
        else if (childRole === 'DM') childName = ch.district_name || '';
        else if (childRole === 'BM') childName = ch.branch_name || '';
        else if (childRole === 'FO') childName = getFullName(ch.emp_id, ch.name || ch.officer_name || '');
        var initial = childName.charAt(0).toUpperCase();
        var dataAttr = childRole === 'FO'
          ? 'data-emp-id="' + esc(ch.emp_id || '') + '" data-emp-name="' + esc(childName) + '"'
          : 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(childName) + '"';

        html += '<tr class="emp-sub-card" ' + dataAttr + '>' +
          '<td class="name-col"><span class="tbl-avatar" style="' + avStyle + '">' + initial + '</span>' + esc(childName) + '</td>' +
          '<td class="num-col">' + fmtNum(numVal(ch.total_count)) + '</td>' +
          '<td class="num-col" style="color:#F59E0B;font-weight:600;">' + fmtAmt(numVal(ch.total_amount)) + '</td>' +
          '<td class="tbl-arrow">&#8250;</td>' +
        '</tr>';
      }
      html += '</tbody></table></div>';
    } else {
      html += '<div class="desktop-grid-3">';
      for (var i = 0; i < children.length; i++) {
        var ch = children[i];
        var childName = '';
        if (childRole === 'RM') childName = ch.region_name || '';
        else if (childRole === 'DM') childName = ch.district_name || '';
        else if (childRole === 'BM') childName = ch.branch_name || '';
        else if (childRole === 'FO') childName = getFullName(ch.emp_id, ch.name || ch.officer_name || '');
        var initial = childName.charAt(0).toUpperCase();
        var dataAttr = childRole === 'FO'
          ? 'data-emp-id="' + esc(ch.emp_id || '') + '" data-emp-name="' + esc(childName) + '"'
          : 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(childName) + '"';

        html += '<div class="emp-sub-card desktop-sub-card" ' + dataAttr + '>';
        html += '<div class="emp-sub-avatar" style="background:#FFFBEB;color:#F59E0B;">' + initial + '</div>';
        html += '<div class="emp-sub-info"><div class="emp-sub-name">' + esc(childName) + '</div>';
        html += '<div class="emp-sub-meta"><span>A/c: ' + fmtNum(numVal(ch.total_count)) + '</span><span>Amt: ' + fmtAmt(numVal(ch.total_amount)) + '</span></div></div>';
        html += '<div class="emp-sub-pct" style="color:#F59E0B">' + fmtAmt(numVal(ch.total_amount)) + '</div>';
        html += '<div class="emp-sub-arrow">&#8250;</div></div>';
      }
      html += '</div>';
    }

    html += '</div>';
    return html;
  }

  function render(byProduct) {
    var container = document.getElementById('disbursementContent');
    var d = _db.summary;
    if (!d || (numVal(d.total_count) === 0 && numVal(d.total_amount) === 0)) {
      container.innerHTML = '<div class="desktop-content-area">' + monthSelectorHtml() + productPillsHtml() + '<div style="text-align:center;padding:80px 20px;color:#64748B;">No disbursement data for this selection.</div></div>';
      attachHandlers();
      return;
    }

    var totalCount = numVal(d.total_count);
    var totalAmount = numVal(d.total_amount);
    var ats = totalCount > 0 ? totalAmount / totalCount : 0;

    var html = '';
    html += monthSelectorHtml();
    html += productPillsHtml();

    // Summary cards
    html += '<div class="emp-fade" style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:16px;">';
    html += '<div style="background:linear-gradient(135deg,#F59E0B,#FBBF24);border-radius:14px;padding:18px 20px;color:#fff;">';
    html += '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">Accounts</div>';
    html += '<div style="font-size:24px;font-weight:700;">' + fmtNum(totalCount) + '</div></div>';
    html += '<div style="background:linear-gradient(135deg,#F59E0B,#FBBF24);border-radius:14px;padding:18px 20px;color:#fff;">';
    html += '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">Amount</div>';
    html += '<div style="font-size:24px;font-weight:700;">' + fmtAmt(totalAmount) + '</div></div>';
    html += '<div style="background:linear-gradient(135deg,#F59E0B,#FBBF24);border-radius:14px;padding:18px 20px;color:#fff;">';
    html += '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">ATS</div>';
    html += '<div style="font-size:24px;font-weight:700;">' + fmtATS(ats) + '</div></div>';
    html += '</div>';

    // Product breakdown table
    if (byProduct && byProduct.length > 0) {
      html += '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);overflow:hidden;margin-bottom:16px;">';
      html += '<div style="padding:14px 16px;border-bottom:1px solid #F1F5F9;font-weight:700;font-size:14px;color:#1E293B;">Product-wise Disbursement</div>';
      html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
      html += '<thead><tr style="background:#F8FAFC;"><th style="padding:10px 14px;text-align:left;color:#64748B;font-weight:600;">Product</th><th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;">Accounts</th><th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;">Amount</th><th style="padding:10px 14px;text-align:right;color:#64748B;font-weight:600;">ATS</th></tr></thead><tbody>';
      var colors = { IGL: '#10B981', FIG: '#6366F1', IL: '#F59E0B' };
      for (var i = 0; i < byProduct.length; i++) {
        var p = byProduct[i];
        var cnt = numVal(p.total_count);
        var amt = numVal(p.total_amount);
        var color = colors[p.product_name] || '#64748B';
        html += '<tr style="border-bottom:1px solid #F1F5F9;"><td style="padding:10px 14px;"><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + color + ';margin-right:8px;vertical-align:middle;"></span>' + esc(p.product_name) + '</td>';
        html += '<td style="padding:10px 14px;text-align:right;">' + fmtNum(cnt) + '</td>';
        html += '<td style="padding:10px 14px;text-align:right;">' + fmtAmt(amt) + '</td>';
        html += '<td style="padding:10px 14px;text-align:right;">' + fmtATS(cnt > 0 ? amt/cnt : 0) + '</td></tr>';
      }
      html += '</tbody></table></div>';
    }

    // Month-wise trend with bar chart
    if (_db.byMonth && _db.byMonth.length > 1) {
      var maxAmt = 0;
      for (var i = 0; i < _db.byMonth.length; i++) {
        var a = numVal(_db.byMonth[i].total_amount);
        if (a > maxAmt) maxAmt = a;
      }
      html += '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);overflow:hidden;margin-bottom:16px;">';
      html += '<div style="padding:14px 16px;border-bottom:1px solid #F1F5F9;font-weight:700;font-size:14px;color:#1E293B;">Month-wise Trend</div>';
      html += '<div style="padding:16px;overflow-x:auto;">';
      // Bar chart
      html += '<div style="display:flex;align-items:flex-end;gap:6px;height:180px;padding:0 4px;">';
      for (var i = 0; i < _db.byMonth.length; i++) {
        var m = _db.byMonth[i];
        var amt = numVal(m.total_amount);
        var pct = maxAmt > 0 ? (amt / maxAmt * 100) : 0;
        var barH = Math.max(pct * 1.5, 8);
        html += '<div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;">';
        html += '<div style="font-size:10px;font-weight:600;color:#1E293B;">' + fmtAmt(amt) + '</div>';
        html += '<div style="width:100%;max-width:48px;height:' + barH + 'px;background:linear-gradient(180deg,#F59E0B,#FBBF24);border-radius:6px 6px 0 0;transition:height .3s;"></div>';
        html += '<div style="font-size:10px;font-weight:600;color:#64748B;white-space:nowrap;">' + esc(m.db_month) + '</div>';
        html += '<div style="font-size:9px;color:#94A3B8;">' + fmtNum(numVal(m.total_count)) + '</div>';
        html += '</div>';
      }
      html += '</div></div></div>';
    }

    // Sub-units (drill-down)
    var children = _db.children || [];
    if (session.role && session.role !== 'FO' && children.length > 0) {
      var childRoleMap = { CEO: 'RM', RM: 'DM', SM: 'DM', DM: 'BM', DvM: 'BM', AM: 'FO', BM: 'FO' };
      var childRole = childRoleMap[session.role];
      if (childRole) {
        children.sort(function(a, b) { return numVal(b.total_amount) - numVal(a.total_amount); });
        html += dbSubUnitsHtml(children, childRole);
      }
    }

    container.innerHTML = '<div class="desktop-content-area">' + html + '</div>';
    attachHandlers();
  }

  function monthSelectorHtml() { return ""; }

  function productPillsHtml() {
    var pills = [{key:'all',label:'All'},{key:'igl',label:'IGL'},{key:'fig',label:'FIG'},{key:'il',label:'IL'}];
    var html = '<div class="emp-fade" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:16px;">';
    for (var i = 0; i < pills.length; i++) {
      var p = pills[i];
      var active = _db.product === p.key;
      html += '<button data-db-product="' + p.key + '" style="padding:6px 16px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;border:none;transition:all .2s;' + (active ? 'background:#F59E0B;color:#fff;' : 'background:#F1F5F9;color:#64748B;') + '">' + p.label + '</button>';
    }
    html += '</div>';
    return html;
  }

  var _dbHandlersAttached = false;
  function attachHandlers() {
    var container = document.getElementById('disbursementContent');
    if (!container) return;

    container.querySelectorAll('[data-db-month]').forEach(function(btn) {
      btn.onclick = function() {
        _db.month = btn.dataset.dbMonth;
        localStorage.setItem('dbMonth', _db.month);
        loadAndRender();
      };
    });

    container.querySelectorAll('[data-db-product]').forEach(function(pill) {
      pill.onclick = function() {
        _db.product = pill.dataset.dbProduct;
        localStorage.setItem('dbProduct', _db.product);
        loadAndRender();
      };
    });

    // Subunit view toggle (card/table)
    container.querySelectorAll('[data-subview]').forEach(function (btn) {
      btn.onclick = function () {
        localStorage.setItem('subunitView', btn.dataset.subview);
        render(_db.byProduct);
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

    if (!_dbHandlersAttached) {
      _dbHandlersAttached = true;
      container.addEventListener('click', function(ev) {
        var card = ev.target.closest('.emp-sub-card');
        if (!card) return;
        var disbTab = document.getElementById('disbursementTab');
        if (!disbTab || !disbTab.classList.contains('active')) return;
        card.style.background = '#FFFBEB';
        card.style.pointerEvents = 'none';
        var arrow = card.querySelector('.emp-sub-arrow');
        if (arrow) arrow.innerHTML = '<div style="width:16px;height:16px;border:2px solid #FDE68A;border-top-color:#F59E0B;border-radius:50%;animation:spin .7s linear infinite;"></div>';
        if (card.dataset.childRole) {
          pushRoleNav(card.dataset.childRole, card.dataset.childLocation);
          window.location.reload();
        } else if (card.dataset.empId) {
          var stack = getRoleNavStack();
          var current = getEmployeeSession();
          stack.push({role:current.role,location:current.location,name:current.name,id:current.id});
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

  window._loadDisbursementTab = function() {
    try { loadAndRender(); } catch(e) { console.error('Disbursement load failed:', e); }
  };

  // ── Keyboard navigation: Left/Right = prev/next date; Up/Down = next/prev month ──
  function dbArrowHandler(e) {
    var key = e.key;
    if (key !== 'ArrowLeft' && key !== 'ArrowRight' && key !== 'ArrowUp' && key !== 'ArrowDown') return;
    // Only on disbursement tab
    var disbTab = document.getElementById('disbursementTab');
    if (!disbTab || !disbTab.classList.contains('active')) return;
    // Skip if a text input / editable element is focused
    var ae = document.activeElement;
    if (ae) {
      var tag = (ae.tagName || '').toUpperCase();
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
      if (ae.isContentEditable) return;
    }
    // Skip if any modal is visible
    var mp = document.getElementById('monthPickerOverlay');
    if (mp && mp.style.display === 'flex') return;
    var cal = document.getElementById('calendarOverlay');
    if (cal && cal.classList.contains('show')) return;
    // Need a populated list + current selection
    var list = _db.availableDates || [];
    if (!list.length) return;
    var cur = _db.date;
    if (!cur) return;
    var idx = list.indexOf(cur);
    if (idx < 0) return;

    var target = null;
    if (key === 'ArrowLeft') {
      // older date = higher idx in desc-sorted list
      if (idx < list.length - 1) target = list[idx + 1];
    } else if (key === 'ArrowRight') {
      // newer date = lower idx
      if (idx > 0) target = list[idx - 1];
    } else if (key === 'ArrowUp') {
      // next (newer) month — walk to lower idx, first hit with ym > curYm
      var curYmU = cur.slice(0, 7);
      for (var i = idx - 1; i >= 0; i--) {
        if (list[i].slice(0, 7) > curYmU) { target = list[i]; break; }
      }
    } else if (key === 'ArrowDown') {
      // prev (older) month — walk to higher idx, first hit with ym < curYm
      var curYmD = cur.slice(0, 7);
      for (var j = idx + 1; j < list.length; j++) {
        if (list[j].slice(0, 7) < curYmD) { target = list[j]; break; }
      }
    }
    if (!target) return;
    e.preventDefault();
    _db.date = target;
    _db.month = '';
    localStorage.setItem('dbSelectedDate', target);
    localStorage.setItem('dbMonth', '');
    if (typeof window.updateArrowVisibility === 'function') window.updateArrowVisibility('disbursement');
    loadAndRender();
  }

  if (!window._dbKeyboardBound) {
    window._dbKeyboardBound = true;
    window.addEventListener('keydown', dbArrowHandler);
  }
})();
