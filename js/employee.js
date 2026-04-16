(function () {
  if (!isEmployeeAuthenticated()) { window.location.href = 'index.html'; return; }

  var session = getEmployeeSession();

  // Header subtitle — role-aware
  var subtitleEl = document.getElementById('emp-subtitle');
  if (session.role === 'CEO') {
    subtitleEl.textContent = 'CEO \u2014 Overall';
  } else if (session.role === 'RM') {
    subtitleEl.textContent = 'Regional Manager \u2014 ' + session.location;
  } else if (session.role === 'DM') {
    subtitleEl.textContent = 'Division Manager \u2014 ' + session.location;
  } else if (session.role === 'BM') {
    subtitleEl.textContent = 'Branch Manager \u2014 ' + session.location;
  } else if (session.role === 'SM') {
    subtitleEl.textContent = 'Regional Manager \u2014 ' + session.location;
  } else if (session.role === 'DvM') {
    subtitleEl.textContent = 'Division Manager \u2014 ' + session.location;
  } else if (session.role === 'AM') {
    subtitleEl.textContent = 'Area Manager \u2014 ' + session.location;
  } else {
    subtitleEl.textContent = session.name;
  }

  // Date badge — show current date
  var now = new Date();
  var dd = String(now.getDate()).padStart(2, '0');
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var yyyy = now.getFullYear();
  document.getElementById('dateBadgeText').textContent = dd + '-' + mm + '-' + yyyy;

  // Logout button
  document.getElementById('logout-btn').onclick = logout;

  // Fetch and show phone number for FO viewed from analytical
  var SUPABASE_URL = 'https://zovnmmdfthpbubrorsgh.supabase.co';
  var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inpvdm5tbWRmdGhwYnVicm9yc2doIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjE1NzE3ODgsImV4cCI6MjA3NzE0Nzc4OH0.92BH2sjUOgkw6iSRj1_4gt0p3eThg3QT4VK-Q4EdmBE';
  async function _showFOPhone(empId) {
    try {
      var url = SUPABASE_URL + '/rest/v1/employees?select=emp_id,mobile&or=(emp_id.ilike.' + empId + ')';
      var resp = await fetch(url, { headers: { 'apikey': SUPABASE_KEY, 'Authorization': 'Bearer ' + SUPABASE_KEY } });
      var data = await resp.json();
      if (Array.isArray(data) && data.length && data[0].mobile) {
        var phone = data[0].mobile;
        var phoneBar = document.createElement('div');
        phoneBar.style.cssText = 'display:flex;align-items:center;justify-content:center;gap:8px;padding:10px 20px;background:rgba(52,211,153,0.06);border-bottom:1px solid rgba(52,211,153,0.15);';
        phoneBar.innerHTML = '<a href="tel:' + phone + '" style="display:flex;align-items:center;gap:8px;color:#34D399;font-size:14px;font-weight:600;text-decoration:none;">' +
          '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>' +
          'Call ' + phone + '</a>';
        var header = document.querySelector('.emp-header');
        if (header) header.after(phoneBar);
      }
    } catch (e) { console.error('Phone fetch failed:', e); }
  }

  /* ---------- Formatters (for disbursement tab) ---------- */
  function fmtNum(v) {
    if (v == null || v === '') return '-';
    if (typeof v === 'string') return v.trim() || '-';
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return v.toLocaleString('en-IN');
      return v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(v);
  }
  function numVal(v) { return typeof v === 'number' ? v : 0; }

  /* ---------- Disbursement tab helpers ---------- */
  var REG = { demand: 2, collection: 3, ftod: 4, pct: 5 };
  var BUCKETS = [
    { name: 'SMA-0', range: '1 \u2014 30 DPD', key: '1-30',  color: '#34D399', demand: 6,  collection: 7  },
    { name: 'SMA-1', range: '31 \u2014 60 DPD', key: '31-60', color: '#FBBF24', demand: 10, collection: 11 },
    { name: 'SMA-2',     range: 'Pre-NPA',           key: 'pnpa',  color: '#FB923C', demand: 14, collection: 15 },
    { name: 'NPA',      range: 'NPA',               key: 'npa',   color: '#F87171', demand: 18, collection: 19 }
  ];

  function renderDashboard(empRow, containerId) {
    var totalDemand = numVal(empRow[REG.demand]);
    var totalPos    = numVal(empRow[REG.collection]);
    for (var i = 0; i < BUCKETS.length; i++) {
      totalDemand += numVal(empRow[BUCKETS[i].demand]);
      totalPos    += numVal(empRow[BUCKETS[i].collection]);
    }
    var html = '<div class="emp-snapshot emp-fade">' +
      '<div class="emp-snapshot-label">Snapshot</div>' +
      '<div class="emp-snapshot-row">' +
        '<div class="emp-snapshot-metric"><div class="emp-snap-lbl">Total Demand</div><div class="emp-snap-val">' + fmtNum(totalDemand) + '</div></div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric"><div class="emp-snap-lbl">Total Collection</div><div class="emp-snap-val">' + fmtNum(totalPos) + '</div></div>' +
      '</div></div>';
    document.getElementById(containerId).innerHTML = html;
  }

  var noDataHtml = '<div style="text-align:center;padding:80px 20px;">' +
    '<p style="color:#64748B;font-size:32px;margin-bottom:8px;">&#128202;</p>' +
    '<p style="color:#64748B;font-size:14px;">No data uploaded yet.</p></div>';

  /* ---------- Breadcrumbs ---------- */
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }

  // Floating back button
  window.goBackOneLevel = function() {
    var stack = getRoleNavStack();
    if (!stack.length) return;
    if (localStorage.getItem('returnToAnalytical')) {
      localStorage.removeItem('returnToAnalytical');
      localStorage.setItem('openAnalytical', 'true');
    }
    popRoleNav();
    window.location.reload();
  };

  function renderBreadcrumbs() {
    var navStack = getRoleNavStack();
    var bar = document.getElementById('breadcrumbBar');
    var floatBtn = document.getElementById('floatingBackBtn');
    if (!bar) return;
    if (!navStack.length) {
      bar.style.display = 'none';
      if (floatBtn) floatBtn.style.display = 'none';
      return;
    }
    bar.style.display = '';
    if (floatBtn) floatBtn.style.display = 'block';
    var html = '<span class="emp-bc-back" onclick="goBackOneLevel()" style="cursor:pointer;padding:4px 10px;border-radius:6px;background:#059669;color:#fff;font-weight:700;font-size:13px;margin-right:4px;">&#8592; Back</span>';
    for (var i = 0; i < navStack.length; i++) {
      var item = navStack[i];
      var label = item.role === 'CEO' ? 'CEO' : item.location || item.name || item.role;
      html += '<span class="emp-bc-item" data-nav-index="' + i + '">' + esc(label) + '</span>';
      html += '<span class="emp-bc-sep">&#8250;</span>';
    }
    var currentLabel = session.role === 'CEO' ? 'CEO' : session.role === 'FO' ? session.name : session.location;
    html += '<span class="emp-bc-item current">' + esc(currentLabel) + '</span>';
    bar.innerHTML = html;
    bar.onclick = function (ev) {
      var item = ev.target.closest('.emp-bc-item[data-nav-index]');
      if (!item) return;
      var idx = parseInt(item.dataset.navIndex, 10);
      // If navigating back from analytical drill-down, return to analytical tab
      if (localStorage.getItem('returnToAnalytical')) {
        localStorage.removeItem('returnToAnalytical');
        localStorage.setItem('openAnalytical', 'true');
      }
      restoreNavTo(idx);
      window.location.reload();
    };
  }

  /* ---------- Load Data ---------- */
  async function loadData() {
    try {
      await initDB();

      // Render breadcrumbs (shared across tabs)
      renderBreadcrumbs();

      // Load collection tab (new collection.js handles everything)
      if (typeof _loadCollectionTab === 'function') _loadCollectionTab();

      // Load hourly tab
      if (typeof _loadHourlyTab === 'function') _loadHourlyTab();

      // Load portfolio tab
      if (typeof _loadPortfolioTab === 'function') _loadPortfolioTab();

      // Load disbursement tab
      if (typeof _loadDisbursementTab === 'function') _loadDisbursementTab();

      // Load comparison tab
      if (typeof _loadComparisonTab === 'function') _loadComparisonTab();

      // Load analytical tab
      if (typeof _loadAnalyticalTab === 'function') _loadAnalyticalTab();

      // Check if returning from analytical FO drill-down
      var returnFlag = localStorage.getItem('returnToAnalytical');
      if (returnFlag) {
        // Show back button
        var headerLeft = document.querySelector('.emp-header-left');
        if (headerLeft) {
          var backBtn = document.createElement('button');
          backBtn.style.cssText = 'font-size:20px;line-height:1;margin-right:10px;padding:8px 12px;border-radius:10px;background:#F0FDF4;border:1px solid #A7F3D0;color:#059669;cursor:pointer;flex-shrink:0;';
          backBtn.innerHTML = '&#8592;';
          backBtn.title = 'Back to Analytical';
          backBtn.onclick = function() {
            localStorage.removeItem('returnToAnalytical');
            popRoleNav();
            localStorage.setItem('openAnalytical', 'true');
            window.location.reload();
          };
          headerLeft.insertBefore(backBtn, headerLeft.firstChild);
        }
        // Show phone number for this FO
        if (session.role === 'FO' && session.id) {
          _showFOPhone(session.id);
        }
      }

      // Restore active tab on reload
      if (localStorage.getItem('openAnalytical')) {
        localStorage.removeItem('openAnalytical');
        switchEmpTab('analytical');
      } else {
        // Always start on Collection tab on reload
        switchEmpTab('collection');
      }
    } catch (err) {
      console.error('Load failed:', err);
      showNoData();
    }
  }

  function showNoData() {
    document.getElementById('no-data').classList.remove('hidden');
  }

  // Tab switching (global for onclick)
  var tabTitles = { portfolio: 'Portfolio', disbursement: 'Disbursement', collection: 'Collection', hourly: 'Hourly', comparison: 'Comparison', analytical: 'Analytical Tool', ai: 'AI', dailyreports: 'Daily Reports', editbranch: 'EDIT_BRANCH_DAILY_REPORT', settings: 'Settings' };
  window.switchEmpTab = function (tab) {
    // Switch tab content
    document.querySelectorAll('.emp-tab-content').forEach(function (tc) {
      tc.classList.remove('active');
    });
    document.getElementById(tab + 'Tab').classList.add('active');

    // Bottom tab bar — show for main tabs, hide for analytical
    var isMainTab = (tab === 'portfolio' || tab === 'disbursement' || tab === 'collection' || tab === 'hourly');
    var tabBar = document.querySelector('.emp-tab-bar');
    if (tabBar) tabBar.style.display = isMainTab ? '' : 'none';
    document.querySelectorAll('.emp-tab-item').forEach(function (t) {
      t.classList.toggle('active', isMainTab && t.dataset.tab === tab);
    });

    // Sidebar — highlight active item (home = collection/portfolio/disbursement)
    var sidebarTab = isMainTab ? 'home' : tab;
    document.querySelectorAll('[data-sidebar-tab]').forEach(function (s) {
      s.classList.toggle('active', s.dataset.sidebarTab === sidebarTab);
    });
    // Desktop sidebar — sync active highlight
    document.querySelectorAll('.desktop-sidebar-item[data-dtab]').forEach(function (el) {
      el.classList.toggle('active', el.dataset.dtab === tab);
    });

    // Header title
    var titleEl = document.getElementById('emp-header-title');
    if (titleEl) titleEl.textContent = tabTitles[tab] || tab;

    // Persist active tab for reload (don't persist 'analytical' — always default to collection on next login)
    if (tab !== 'analytical') localStorage.setItem('activeTab', tab);
  };

  loadData();
})();
