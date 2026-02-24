(function () {
  if (!isEmployeeAuthenticated()) { window.location.href = 'index.html'; return; }

  var session = getEmployeeSession();

  // Header subtitle
  document.getElementById('emp-subtitle').textContent = 'Collection \u2014 ' + session.name;

  // Date badge — show current month-end
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

  /* ---------- Render ---------- */
  function renderDashboard(empRow) {
    // Compute totals
    var totalDemand = numVal(empRow[REG.demand]);
    var totalPos      = numVal(empRow[REG.collection]);
    for (var i = 0; i < BUCKETS.length; i++) {
      totalDemand += numVal(empRow[BUCKETS[i].demand]);
      totalPos      += numVal(empRow[BUCKETS[i].collection]);
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
          '<div class="emp-snap-val">' + fmtCr(totalPos) + '<span class="emp-snap-unit">' + crUnit(totalPos) + '</span></div>' +
        '</div>' +
      '</div>' +
    '</div>';

    // ── DPD BUCKET ALLOCATION ──
    html += '<div class="emp-section-title">DPD Bucket Allocation</div>';

    // Reg DPD summary row
    html += '<div class="emp-reg-summary emp-fade">' +
      '<div class="emp-reg-label">Reg DPD</div>' +
      '<div class="emp-reg-values">' +
        '<div class="emp-reg-stat">' +
          '<div class="emp-reg-val">' + fmtNum(empRow[REG.demand]) + '</div>' +
          '<div class="emp-reg-lbl">Demand</div>' +
        '</div>' +
        '<div class="emp-reg-stat">' +
          '<div class="emp-reg-val emp-accent-text">' + fmtCr(empRow[REG.collection]) + '<span class="emp-reg-unit"> ' + crUnit(empRow[REG.collection]) + '</span></div>' +
          '<div class="emp-reg-lbl">Collection</div>' +
        '</div>' +
      '</div>' +
    '</div>';

    // Bucket cards
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
            '<div class="emp-bucket-val emp-pos-val">' + fmtCr(collection) + '<span class="emp-cr-unit">' + crUnit(collection) + '</span></div>' +
            '<div class="emp-bucket-lbl">Collection</div>' +
          '</div>' +
        '</div>' +
        '<div class="emp-bucket-bar" style="width:' + barW + '%;background:' + bk.color + '"></div>' +
      '</div>';
    }
    html += '</div>';

    // ── ACCOUNTS DISTRIBUTION BAR CHART ──
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

  /* ---------- Load Data ---------- */
  async function loadData() {
    try {
      await initDB();
      var wb = await getWorkbook();
      if (!wb || !wb.data) { showNoData(); return; }

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });
      var needle   = session.name.toUpperCase().trim();
      var needleId = String(session.id).trim();

      // Find "Overall" sheet
      var targetSheet = null;
      for (var s = 0; s < workbook.SheetNames.length; s++) {
        if (workbook.SheetNames[s].toUpperCase().includes('OVERALL')) {
          targetSheet = workbook.Sheets[workbook.SheetNames[s]];
          break;
        }
      }
      // Fallback: first wide sheet (>= 15 cols)
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

      // Find employee row
      var empRow = null;
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        if (!row || !row.length) continue;
        for (var c = 0; c < Math.min(row.length, 5); c++) {
          var cv = String(row[c] || '').trim();
          if (cv.toUpperCase() === needle || cv === needleId) { empRow = row; break; }
        }
        if (empRow) break;
      }
      if (!empRow) { showNoData(); return; }

      renderDashboard(empRow);
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
