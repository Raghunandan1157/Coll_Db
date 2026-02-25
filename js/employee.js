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
    subtitleEl.textContent = 'District Manager \u2014 ' + session.location;
  } else if (session.role === 'BM') {
    subtitleEl.textContent = 'Branch Manager \u2014 ' + session.location;
  } else {
    subtitleEl.textContent = 'Collection \u2014 ' + session.name;
  }

  // Date badge — show current date
  var now = new Date();
  var dd = String(now.getDate()).padStart(2, '0');
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var yyyy = now.getFullYear();
  document.getElementById('dateBadgeText').textContent = dd + '-' + mm + '-' + yyyy;

  // Logout button
  document.getElementById('logout-btn').onclick = logout;

  // Column indices (0-indexed from Excel)
  var REG = { demand: 2, collection: 3, ftod: 4, pct: 5 };
  var BUCKETS = [
    { name: 'Bucket 1', range: '1 \u2014 30 DPD', key: '1-30',  color: '#34D399', demand: 6,  collection: 7  },
    { name: 'Bucket 2', range: '31 \u2014 60 DPD', key: '31-60', color: '#FBBF24', demand: 10, collection: 11 },
    { name: 'Bucket 3', range: '61 \u2014 90 DPD', key: '61-90', color: '#FB923C', demand: 14, collection: 15 },
    { name: 'NPA',      range: '90+ DPD',          key: 'npa',   color: '#F87171', demand: 18, collection: 19 }
  ];

  /* ---------- Formatters ---------- */
  function fmtNum(v) {
    if (v == null || v === '') return '-';
    if (typeof v === 'string') return v.trim() || '-';
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return v.toLocaleString('en-IN');
      return v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(v);
  }
  function fmtCr(v) {
    if (v == null || v === '') return '-';
    if (typeof v === 'string') return v.trim() || '-';
    if (typeof v === 'number') {
      var n;
      if (Math.abs(v) >= 10000000) n = v / 10000000;
      else if (Math.abs(v) >= 100000) n = v / 100000;
      else n = v;
      return n % 1 === 0 ? String(Math.round(n)) : n.toFixed(2);
    }
    return String(v);
  }
  function crUnit(v) {
    if (typeof v !== 'number') return '';
    if (Math.abs(v) >= 10000000) return 'Cr';
    if (Math.abs(v) >= 100000)   return 'L';
    return '';
  }
  function numVal(v) { return typeof v === 'number' ? v : 0; }
  function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }

  /* ---------- Render Dashboard ---------- */
  function renderDashboard(empRow) {
    var totalDemand = numVal(empRow[REG.demand]);
    var totalPos    = numVal(empRow[REG.collection]);
    for (var i = 0; i < BUCKETS.length; i++) {
      totalDemand += numVal(empRow[BUCKETS[i].demand]);
      totalPos    += numVal(empRow[BUCKETS[i].collection]);
    }

    var html = '';

    // ── SNAPSHOT ──
    html += '<div class="emp-snapshot emp-fade">' +
      '<div class="emp-snapshot-label">Snapshot</div>' +
      '<div class="emp-snapshot-row">' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Demand</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalDemand) + '</div>' +
        '</div>' +
        '<div class="emp-snap-divider"></div>' +
        '<div class="emp-snapshot-metric">' +
          '<div class="emp-snap-lbl">Total Collection</div>' +
          '<div class="emp-snap-val">' + fmtNum(totalPos) + '</div>' +
        '</div>' +
      '</div>' +
    '</div>';

    // ── DPD BUCKET ALLOCATION ──
    html += '<div class="emp-section-title">DPD Bucket Allocation</div>';

    html += '<div class="emp-reg-summary emp-fade">' +
      '<div class="emp-reg-label">Reg DPD</div>' +
      '<div class="emp-reg-values">' +
        '<div class="emp-reg-stat">' +
          '<div class="emp-reg-val">' + fmtNum(empRow[REG.demand]) + '</div>' +
          '<div class="emp-reg-lbl">Demand</div>' +
        '</div>' +
        '<div class="emp-reg-stat">' +
          '<div class="emp-reg-val emp-accent-text">' + fmtNum(empRow[REG.collection]) + '</div>' +
          '<div class="emp-reg-lbl">Collection</div>' +
        '</div>' +
      '</div>' +
    '</div>';

    html += '<div class="emp-buckets">';
    for (var i = 0; i < BUCKETS.length; i++) {
      var bk = BUCKETS[i];
      var demand     = empRow[bk.demand];
      var collection = empRow[bk.collection];
      var barW = numVal(demand) > 0 ? Math.min(Math.round((numVal(collection) / numVal(demand)) * 100), 100) : 0;

      html += '<div class="emp-bucket-card emp-fade" data-bucket="' + bk.key + '" style="animation-delay:' + (0.1 + i * 0.05) + 's">' +
        '<div class="emp-bucket-indicator" style="background:' + bk.color + '"></div>' +
        '<div class="emp-bucket-info">' +
          '<div class="emp-bucket-name">' + bk.name + '</div>' +
          '<div class="emp-bucket-sub">' + bk.range + '</div>' +
        '</div>' +
        '<div class="emp-bucket-stats">' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val" style="color:' + bk.color + '">' + fmtNum(demand) + '</div>' +
            '<div class="emp-bucket-lbl">Demand</div>' +
          '</div>' +
          '<div class="emp-bucket-stat">' +
            '<div class="emp-bucket-val emp-pos-val">' + fmtNum(collection) + '</div>' +
            '<div class="emp-bucket-lbl">Collection</div>' +
          '</div>' +
        '</div>' +
        '<div class="emp-bucket-bar" style="width:' + barW + '%;background:' + bk.color + '"></div>' +
      '</div>';
    }
    html += '</div>';

    // ── COLLECTION % BAR CHART ──
    var barLabels = ['1-30', '31-60', '61-90', 'NPA'];
    var barColors = ['#34D399', '#FBBF24', '#FB923C', '#F87171'];
    html += '<div class="emp-bar-visual emp-fade" style="animation-delay:0.35s">' +
      '<div class="emp-bv-title">Collection %</div>';
    for (var i = 0; i < barLabels.length; i++) {
      var d = numVal(empRow[BUCKETS[i].demand]);
      var c = numVal(empRow[BUCKETS[i].collection]);
      var pct = d > 0 ? Math.min(Math.round((c / d) * 100), 100) : 0;
      var inside = pct >= 15;
      html += '<div class="emp-bar-row">' +
        '<div class="emp-bar-label">' + barLabels[i] + '</div>' +
        '<div class="emp-bar-track">' +
          '<div class="emp-bar-fill" style="width:' + pct + '%;background:' + barColors[i] + ';">' + (inside ? pct + '%' : '') + '</div>' +
          (inside ? '' : '<span class="emp-bar-outside" style="color:' + barColors[i] + '">' + pct + '%</span>') +
        '</div>' +
      '</div>';
    }
    html += '</div>';

    document.getElementById('collectionContent').innerHTML = html;
  }

  /* ---------- Find row for role-based lookup ---------- */
  function findRoleRow(rows, role, location) {
    var currentSection = null;
    var locationUpper = location.toUpperCase().trim();

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var firstCell = String(row[0] || '').trim().toUpperCase();
      var secondCell = String(row[1] || '').trim().toUpperCase();

      var headerText = firstCell + ' ' + secondCell;
      if (headerText.includes('REGION') && (headerText.includes('WISE') || headerText.includes('NAME'))) {
        currentSection = 'region'; continue;
      }
      if (headerText.includes('DISTRICT') && (headerText.includes('WISE') || headerText.includes('NAME'))) {
        currentSection = 'district'; continue;
      }
      if (headerText.includes('BRANCH') && !headerText.includes('OFFICER') && (headerText.includes('WISE') || headerText.includes('NAME'))) {
        currentSection = 'branch'; continue;
      }

      if (role === 'CEO') {
        if (currentSection === 'region' && (firstCell.includes('GRAND TOTAL') || secondCell.includes('GRAND TOTAL'))) {
          return row;
        }
        continue;
      }

      if (firstCell.includes('GRAND TOTAL') || secondCell.includes('GRAND TOTAL')) {
        currentSection = null; continue;
      }

      var targetSection = null;
      if (role === 'RM') targetSection = 'region';
      else if (role === 'DM') targetSection = 'district';
      else if (role === 'BM') targetSection = 'branch';

      if (currentSection !== targetSection) continue;

      var nameInRow = String(row[1] || '').trim().toUpperCase();
      if (!nameInRow || /^\d+$/.test(nameInRow)) {
        nameInRow = String(row[0] || '').trim().toUpperCase();
      }

      if (nameInRow === locationUpper) {
        return row;
      }
    }

    return null;
  }

  /* ---------- Hierarchy ---------- */
  var _hierarchy = null;

  function loadHierarchy() {
    if (_hierarchy) return Promise.resolve(_hierarchy);
    return fetch('data/hierarchy.json?v=' + Date.now())
      .then(function (r) { return r.json(); })
      .then(function (data) { _hierarchy = data; return data; })
      .catch(function () { _hierarchy = {}; return {}; });
  }

  /* ---------- Find children for drill-down ---------- */
  function findChildrenForRole(rows, hierarchy, role, location) {
    if (role === 'FO') return [];

    var children = [];
    var locationUpper = (location || '').toUpperCase().trim();

    if (role === 'CEO') {
      // Children = all regions
      var regionNames = Object.keys(hierarchy);
      for (var i = 0; i < regionNames.length; i++) {
        var rn = regionNames[i];
        var row = findRoleRow(rows, 'RM', rn);
        if (row) children.push({ name: rn, row: row });
      }
    } else if (role === 'RM') {
      // Children = districts in this region
      var regionData = null;
      for (var rk in hierarchy) {
        if (rk.toUpperCase() === locationUpper) { regionData = hierarchy[rk]; break; }
      }
      if (regionData) {
        var districtNames = Object.keys(regionData);
        for (var i = 0; i < districtNames.length; i++) {
          var dn = districtNames[i];
          var row = findRoleRow(rows, 'DM', dn);
          if (row) children.push({ name: dn, row: row });
        }
      }
    } else if (role === 'DM') {
      // Children = branches in this district
      var branchList = null;
      for (var rk in hierarchy) {
        for (var dk in hierarchy[rk]) {
          if (dk.toUpperCase() === locationUpper) { branchList = hierarchy[rk][dk]; break; }
        }
        if (branchList) break;
      }
      if (branchList) {
        for (var i = 0; i < branchList.length; i++) {
          var bn = branchList[i];
          var row = findRoleRow(rows, 'BM', bn);
          if (row) children.push({ name: bn, row: row });
        }
      }
    } else if (role === 'BM') {
      // Children = officers under this branch from the officer section
      children = findOfficersForBranch(rows, locationUpper);
    }

    // Sort by collection % descending
    children.sort(function (a, b) {
      var pctA = numVal(a.row[REG.demand]) > 0 ? numVal(a.row[REG.collection]) / numVal(a.row[REG.demand]) : 0;
      var pctB = numVal(b.row[REG.demand]) > 0 ? numVal(b.row[REG.collection]) / numVal(b.row[REG.demand]) : 0;
      return pctB - pctA;
    });

    return children;
  }

  /* ---------- Find officers for a branch ---------- */
  function findOfficersForBranch(rows, branchUpper) {
    var officers = [];
    var inOfficerSection = false;
    var currentBranch = null;
    var foundBranch = false;

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var firstCell = String(row[0] || '').trim().toUpperCase();
      var secondCell = String(row[1] || '').trim().toUpperCase();

      // Detect officer section start
      if (firstCell.includes('BRANCH') && firstCell.includes('OFFICER') && firstCell.includes('NAME')) {
        inOfficerSection = true; continue;
      }
      if (!inOfficerSection) continue;

      // Branch header row: empty col0 (or not an empId), branch name in col1 with data
      var col0 = String(row[0] || '').trim();
      var col1 = String(row[1] || '').trim();

      // Skip header/sub-header rows
      if (firstCell === 'EMP ID' || secondCell === 'BRANCH / OFFICER NAME') continue;
      if (firstCell === '' && secondCell === '') continue;
      if (secondCell === 'DEMAND' || secondCell === 'COLLECTION') continue;

      // Is this a branch header? (no empId in col0, name in col1, has numeric data)
      var hasEmpId = col0.length > 0 && /^[A-Z]{2}\d+/.test(col0.toUpperCase());

      if (!hasEmpId && col1.length > 0 && row.length > 3) {
        // This is a branch header row
        currentBranch = col1.toUpperCase();
        if (currentBranch === branchUpper) {
          foundBranch = true;
        } else if (foundBranch) {
          // We've moved past our branch
          break;
        }
        continue;
      }

      // Officer row — has empId
      if (hasEmpId && foundBranch) {
        var officerName = col1.trim();
        officers.push({
          name: officerName,
          empId: col0,
          row: row
        });
      }
    }

    return officers;
  }

  /* ---------- Render sub-unit cards ---------- */
  function renderSubUnits(children, childRole) {
    if (!children.length) return '';

    var roleLabel = { RM: 'Regions', DM: 'Districts', BM: 'Branches', FO: 'Officers' };
    var avatarType = { RM: 'region', DM: 'district', BM: 'branch', FO: 'officer' };
    var label = roleLabel[childRole] || 'Team';
    var avType = avatarType[childRole] || 'branch';

    var html = '<div class="emp-team-section">' +
      '<div class="emp-team-title">' + label +
        '<span class="emp-team-count">' + children.length + '</span>' +
      '</div>';

    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      var demand = numVal(child.row[REG.demand]);
      var collection = numVal(child.row[REG.collection]);
      var pct = demand > 0 ? Math.round((collection / demand) * 100) : 0;
      var pctColor = pct >= 95 ? '#34D399' : pct >= 80 ? '#FBBF24' : '#F87171';
      var displayName = child.name;
      var initial = displayName.charAt(0).toUpperCase();
      var dataAttr = childRole === 'FO'
        ? 'data-emp-id="' + esc(child.empId || '') + '" data-emp-name="' + esc(child.name) + '"'
        : 'data-child-role="' + esc(childRole) + '" data-child-location="' + esc(child.name) + '"';

      html += '<div class="emp-sub-card" ' + dataAttr + '>' +
        '<div class="emp-sub-avatar ' + avType + '">' + initial + '</div>' +
        '<div class="emp-sub-info">' +
          '<div class="emp-sub-name">' + esc(displayName) + '</div>' +
          '<div class="emp-sub-meta">' +
            '<span>D: ' + fmtNum(demand) + '</span>' +
            '<span>C: ' + fmtNum(collection) + '</span>' +
          '</div>' +
        '</div>' +
        '<div class="emp-sub-pct" style="color:' + pctColor + '">' + pct + '%</div>' +
        '<div class="emp-sub-arrow">&#8250;</div>' +
      '</div>';
    }

    html += '</div>';
    return html;
  }

  /* ---------- Render breadcrumbs ---------- */
  function renderBreadcrumbs() {
    var navStack = getRoleNavStack();
    var bar = document.getElementById('breadcrumbBar');
    if (!bar) return;

    // Only show if there's navigation history
    if (!navStack.length) {
      bar.style.display = 'none';
      return;
    }

    bar.style.display = '';
    var html = '';

    for (var i = 0; i < navStack.length; i++) {
      var item = navStack[i];
      var label = item.role === 'CEO' ? 'CEO' : item.location || item.name || item.role;
      html += '<span class="emp-bc-item" data-nav-index="' + i + '">' + esc(label) + '</span>';
      html += '<span class="emp-bc-sep">&#8250;</span>';
    }

    // Current (non-clickable)
    var currentLabel = session.role === 'CEO' ? 'CEO' :
                       session.role === 'FO' ? session.name :
                       session.location;
    html += '<span class="emp-bc-item current">' + esc(currentLabel) + '</span>';

    bar.innerHTML = html;

    // Click handler for breadcrumb items
    bar.onclick = function (ev) {
      var item = ev.target.closest('.emp-bc-item[data-nav-index]');
      if (!item) return;
      var idx = parseInt(item.dataset.navIndex, 10);
      restoreNavTo(idx);
      window.location.reload();
    };
  }

  /* ---------- Click handler for sub-unit cards ---------- */
  function attachSubUnitClickHandlers() {
    var container = document.getElementById('collectionContent');
    if (!container) return;

    container.addEventListener('click', function (ev) {
      var card = ev.target.closest('.emp-sub-card');
      if (!card) return;

      card.style.background = 'rgba(79,140,255,0.12)';
      card.style.pointerEvents = 'none';

      if (card.dataset.childRole) {
        // Role drill-down (CEO→RM, RM→DM, DM→BM)
        pushRoleNav(card.dataset.childRole, card.dataset.childLocation);
        window.location.reload();
      } else if (card.dataset.empId) {
        // BM→FO drill-down
        var stack = getRoleNavStack();
        var current = getEmployeeSession();
        stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
        localStorage.setItem('roleNavStack', JSON.stringify(stack));
        // Set FO session
        localStorage.removeItem('roleAuth');
        localStorage.removeItem('roleName');
        localStorage.removeItem('roleLocation');
        localStorage.setItem('employeeId', card.dataset.empId);
        localStorage.setItem('employeeName', card.dataset.empName);
        window.location.reload();
      }
    });
  }

  /* ---------- Load Data ---------- */
  async function loadData() {
    try {
      await initDB();
      var wb = await getWorkbookWithFallback();
      if (!wb || !wb.data) { showNoData(); return; }

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });

      // Find "Overall" sheet
      var targetSheet = null;
      for (var s = 0; s < workbook.SheetNames.length; s++) {
        if (workbook.SheetNames[s].toUpperCase().includes('OVERALL')) {
          targetSheet = workbook.Sheets[workbook.SheetNames[s]];
          break;
        }
      }
      if (!targetSheet) {
        for (var s = 0; s < workbook.SheetNames.length; s++) {
          var ws = workbook.Sheets[workbook.SheetNames[s]];
          if (ws && ws['!ref']) {
            var dec = XLSX.utils.decode_range(ws['!ref']);
            if (dec.e.c >= 15) { targetSheet = ws; break; }
          }
        }
      }
      if (!targetSheet) { showNoData(); return; }

      var rows = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });
      if (!rows || rows.length < 2) { showNoData(); return; }

      var empRow = null;

      if (session.role && session.role !== 'FO') {
        empRow = findRoleRow(rows, session.role, session.location);
      } else {
        var needle   = session.name.toUpperCase().trim();
        var needleId = String(session.id).trim();

        for (var r = 0; r < rows.length; r++) {
          var row = rows[r];
          if (!row || !row.length) continue;
          for (var c = 0; c < Math.min(row.length, 5); c++) {
            var cv = String(row[c] || '').trim();
            if (cv.toUpperCase() === needle || cv === needleId) { empRow = row; break; }
          }
          if (empRow) break;
        }
      }

      if (!empRow) { showNoData(); return; }

      renderDashboard(empRow);
      renderBreadcrumbs();

      // Load hierarchy and render sub-units
      if (session.role !== 'FO') {
        var hierarchy = await loadHierarchy();
        var childRoleMap = { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
        var childRole = childRoleMap[session.role];
        if (childRole) {
          var children = findChildrenForRole(rows, hierarchy, session.role, session.location);
          if (children.length) {
            var subHtml = renderSubUnits(children, childRole);
            document.getElementById('collectionContent').innerHTML += subHtml;
          }
        }
      }

      attachSubUnitClickHandlers();

      // Show "Last updated" timestamp
      if (wb.uploadedAt) {
        var d = new Date(wb.uploadedAt);
        var day = String(d.getDate()).padStart(2, '0');
        var mon = String(d.getMonth() + 1).padStart(2, '0');
        var hr = String(d.getHours()).padStart(2, '0');
        var min = String(d.getMinutes()).padStart(2, '0');
        var el = document.getElementById('lastUpdated');
        if (el) el.textContent = 'Updated ' + day + '-' + mon + '-' + d.getFullYear() + ' ' + hr + ':' + min;
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
  window.switchEmpTab = function (tab) {
    document.querySelectorAll('.emp-tab-item').forEach(function (t) {
      t.classList.toggle('active', t.dataset.tab === tab);
    });
    document.querySelectorAll('.emp-tab-content').forEach(function (tc) {
      tc.classList.remove('active');
    });
    document.getElementById(tab + 'Tab').classList.add('active');
  };

  loadData();
})();
