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
  function renderDashboard(empRow, containerId) {
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

    document.getElementById(containerId || 'collectionContent').innerHTML = html;
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

  /* ---------- Hierarchy (embedded) ---------- */
  var HIERARCHY = {"ANDRA PRADESH":{"KADAPA":["BUDWAL","DHARMAVARAM","KADAPA","KADIRI"]},"DHARWAD":{"BADAMI":["BADAMI","GAJENDRAGAD","NARAGUNDA","RAMDURGA"],"BALLARI":["BALLARI","KUDATHINI","SANDURU","SIRUGUPPA"],"BELAGAVI":["BAILHONGAL","BELAGAVI","GOKAK","KITTUR","YARAGATTI"],"CHIKKODI":["ATHANI","CHIKKODI","MUDALAGI","NIPPANI"],"DAVANAGERE":["DAVANAGERE","HARIHARA","HONNALI","SANTHEBENNURU"],"DHARWAD":["DHARWAD","HUBLI","HUBLI-2","KALGHATGI"],"GADAG":["GADAG","LAXMESHWAR","MUNDARAGI"],"KUDLIGI":["HARAPANAHALLI","KHANAHOSAHALLI","KOTTURU","KUDLIGI"],"VIJAYANAGARA":["HAGARIBOMMANAHALLI","HOSPET","HUVENAHADAGALLI"]},"KALBURGI":{"BIDAR":["AURAD","BHALKI","BIDAR","BIDAR-2"],"HUMNABAD":["BASAVAKALYAN","HULSOOR","HUMNABAD","KAMALAPURA"],"INDI":["ALMEL","CHADCHAN","INDI"],"KALBURGI":["AFZALPUR","ALAND","JEVARGI","KALABURAGI","KALBURGI-2"],"KUSHTAGI":["BAGALKOT","GANGAVATHI","HUNGUND","KOPPAL","KUSHTAGI"],"LINGSUGUR":["DEVADURGA","LINGSUGUR","MANVI","RAICHUR","SINDHNUR","SIRWAR"],"SEDAM":["CHINCHOLI","KALAGI","SEDAM","SHAHAPUR","YADGIR"],"VIJAYAPURA":["BILAGI","JAMAKHANDI","LOKAPUR","MUDDEBIHAL","SINDAGI","TALIKOTI","TIKOTA","VIJAYAPUR"]},"TELANGANA":{"MAHABOOBNAGAR":["GADWAL","MAHABUB NAGAR","MARIKAL","TANDUR"],"SANGAREDDY":["KODANGAL","NARAYANKHED","SANGAREDDY","ZAHEERABAD"]},"TUMKUR":{"BENGALORE -RURAL":["DABUSPET","DODDABALLAPURA","GOWRIBIDANUR"],"BENGALORE -URBAN":["CHANDAPURA","HEBBAL","J P NAGAR","KENGERI"],"CHIKKABALLAPUR":["BAGEPALLI","CHIKBALLAPURA","CHINTAMANI","DEVANAHALLI","SRINIVASPURA"],"CHIKKAMAGALURU":["CHIKKAMAGALURU","MUDIGERE","NR PURA"],"CHITRADURGA":["CHALLAKERE","CHITRADURGA","HIRIYUR","JAGALORE"],"HOLALKERE":["CHANNAGIRI","HOLAKERE","HOSADURGA"],"KADUR":["AJJAMPURA","KADUR","PANCHANHALLI","TARIKERE"],"KOLAR":["BANGARPET","BETHAMANGALA","KOLAR","MALUR"],"TIPTUR":["CHIKKANAYAKANAHALLI","GUBBI","HULIYAR","TIPTUR","TUREVEKERE"],"TUMKUR":["KORATAGERE","KUNIGAL","MADHUGIRI","SIRA","TUMKUR"]}};

  /* ---------- Fuzzy name matching ---------- */
  function normalizeForMatch(s) {
    return s.toUpperCase().replace(/[\s\-\u00A0]+/g, '').trim();
  }

  function editDistance(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var prev = [], curr = [];
    for (var j = 0; j <= a.length; j++) prev[j] = j;
    for (var i = 1; i <= b.length; i++) {
      curr[0] = i;
      for (var j = 1; j <= a.length; j++) {
        if (b[i - 1] === a[j - 1]) curr[j] = prev[j - 1];
        else curr[j] = 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
      }
      var tmp = prev; prev = curr; curr = tmp;
    }
    return prev[a.length];
  }

  function fuzzyMatch(a, b) {
    var au = a.toUpperCase().trim();
    var bu = b.toUpperCase().trim();
    if (au === bu) return true;
    var an = normalizeForMatch(a);
    var bn = normalizeForMatch(b);
    if (an === bn) return true;
    // Allow edit distance 1 for strings >= 8 chars (avoids short-name false positives)
    if (an.length >= 8 && bn.length >= 8 && Math.abs(an.length - bn.length) <= 1) {
      return editDistance(an, bn) <= 1;
    }
    return false;
  }

  /* ---------- Collect all data rows from a section (first occurrence only) ---------- */
  function getAllSectionRows(rows, sectionType) {
    var results = [];
    var currentSection = null;
    var foundTarget = false;
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
      if (firstCell.includes('GRAND TOTAL') || secondCell.includes('GRAND TOTAL')) {
        // If we already collected rows for our target section, stop here
        if (currentSection === sectionType && foundTarget) break;
        currentSection = null; continue;
      }
      if (currentSection !== sectionType) continue;

      // Skip header/serial rows
      if (firstCell === 'SL' || firstCell === 'SL.' || firstCell === 'SL NO' || firstCell === 'S.NO' || firstCell === 'SR') continue;
      if (/^TOTAL/i.test(firstCell)) continue;

      var nameInRow = String(row[1] || '').trim().toUpperCase();
      if (!nameInRow || /^\d+$/.test(nameInRow)) {
        nameInRow = String(row[0] || '').trim().toUpperCase();
      }
      if (!nameInRow || /^\d+$/.test(nameInRow)) continue;
      if (/^(SL|NAME|OFFICER|DEMAND|COLLECTION)/i.test(nameInRow)) continue;

      results.push({ name: nameInRow, row: row });
      foundTarget = true;
    }
    return results;
  }

  /* ---------- Find children for drill-down ---------- */
  function findChildrenForRole(rows, hierarchy, role, location) {
    if (role === 'FO') return [];

    var children = [];
    var locationUpper = (location || '').toUpperCase().trim();

    if (role === 'CEO') {
      // Get all region rows directly from the sheet — no hierarchy name-matching needed
      children = getAllSectionRows(rows, 'region');
    } else if (role === 'RM') {
      // Find hierarchy entry for this region (fuzzy match handles name differences)
      var hierarchyDistricts = null;
      for (var rk in hierarchy) {
        if (fuzzyMatch(rk, locationUpper)) { hierarchyDistricts = Object.keys(hierarchy[rk]); break; }
      }
      if (hierarchyDistricts) {
        var allDistricts = getAllSectionRows(rows, 'district');
        for (var i = 0; i < allDistricts.length; i++) {
          for (var j = 0; j < hierarchyDistricts.length; j++) {
            if (fuzzyMatch(allDistricts[i].name, hierarchyDistricts[j])) {
              children.push(allDistricts[i]);
              break;
            }
          }
        }
      }
    } else if (role === 'DM') {
      // Find hierarchy entry for this district (fuzzy match)
      var hierarchyBranches = null;
      for (var rk in hierarchy) {
        for (var dk in hierarchy[rk]) {
          if (fuzzyMatch(dk, locationUpper)) { hierarchyBranches = hierarchy[rk][dk]; break; }
        }
        if (hierarchyBranches) break;
      }
      if (hierarchyBranches) {
        var allBranches = getAllSectionRows(rows, 'branch');
        for (var i = 0; i < allBranches.length; i++) {
          for (var j = 0; j < hierarchyBranches.length; j++) {
            if (fuzzyMatch(allBranches[i].name, hierarchyBranches[j])) {
              children.push(allBranches[i]);
              break;
            }
          }
        }
      }
    } else if (role === 'BM') {
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

      // Detect officer section start (header may be in col 0 or col 1)
      var hdrText = firstCell + ' ' + secondCell;
      if (hdrText.includes('BRANCH') && hdrText.includes('OFFICER') && hdrText.includes('NAME')) {
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
      // Show loading spinner inside card
      var arrow = card.querySelector('.emp-sub-arrow');
      if (arrow) arrow.innerHTML = '<div style="width:16px;height:16px;border:2px solid rgba(79,140,255,0.2);border-top-color:#4F8CFF;border-radius:50%;animation:spin .7s linear infinite;"></div>';

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

  /* ---------- Find target sheet helper ---------- */
  function findOverallSheet(workbook) {
    for (var s = 0; s < workbook.SheetNames.length; s++) {
      if (workbook.SheetNames[s].toUpperCase().includes('OVERALL')) {
        return workbook.Sheets[workbook.SheetNames[s]];
      }
    }
    for (var s = 0; s < workbook.SheetNames.length; s++) {
      var ws = workbook.Sheets[workbook.SheetNames[s]];
      if (ws && ws['!ref']) {
        var dec = XLSX.utils.decode_range(ws['!ref']);
        if (dec.e.c >= 15) return ws;
      }
    }
    return null;
  }

  /* ---------- Find employee row helper ---------- */
  function findEmpRow(rows) {
    if (session.role && session.role !== 'FO') {
      return findRoleRow(rows, session.role, session.location);
    }
    var needle   = session.name.toUpperCase().trim();
    var needleId = String(session.id).trim();
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      for (var c = 0; c < Math.min(row.length, 5); c++) {
        var cv = String(row[c] || '').trim();
        if (cv.toUpperCase() === needle || cv === needleId) return row;
      }
    }
    return null;
  }

  /* ---------- Load a category tab (portfolio / disbursement) ---------- */
  var noDataHtml = '<div style="text-align:center;padding:80px 20px;">' +
    '<p style="color:#6B7A99;font-size:32px;margin-bottom:8px;">&#128202;</p>' +
    '<p style="color:#6B7A99;font-size:14px;">No data uploaded yet.</p></div>';

  async function loadTabData(category, containerId) {
    try {
      var wb = await getWorkbookByCategory(category);
      if (!wb || !wb.data) {
        document.getElementById(containerId).innerHTML = noDataHtml;
        return;
      }

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });
      var sheet = findOverallSheet(workbook);
      if (!sheet) { document.getElementById(containerId).innerHTML = noDataHtml; return; }

      var rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!rows || rows.length < 2) { document.getElementById(containerId).innerHTML = noDataHtml; return; }

      var empRow = findEmpRow(rows);
      if (!empRow) { document.getElementById(containerId).innerHTML = noDataHtml; return; }

      renderDashboard(empRow, containerId);
    } catch (err) {
      console.error(category + ' tab load failed:', err);
      document.getElementById(containerId).innerHTML = noDataHtml;
    }
  }

  /* ---------- Load Data ---------- */
  async function loadData() {
    try {
      await initDB();
      var wb = await getWorkbookWithFallback();
      if (!wb || !wb.data) { showNoData(); return; }

      var workbook = XLSX.read(new Uint8Array(wb.data), { type: 'array', cellFormula: false, cellNF: true });

      var targetSheet = findOverallSheet(workbook);
      if (!targetSheet) { showNoData(); return; }

      var rows = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });
      if (!rows || rows.length < 2) { showNoData(); return; }

      var empRow = findEmpRow(rows);
      if (!empRow) { showNoData(); return; }

      renderDashboard(empRow);
      renderBreadcrumbs();

      // Auto-switch to Collection tab since that's where data renders
      switchEmpTab('collection');

      // Render sub-units for drill-down
      if (session.role !== 'FO') {
        var childRoleMap = { CEO: 'RM', RM: 'DM', DM: 'BM', BM: 'FO' };
        var childRole = childRoleMap[session.role];
        if (childRole) {
          var children = findChildrenForRole(rows, HIERARCHY, session.role, session.location);
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
      // Load portfolio & disbursement tabs
      loadTabData('portfolio', 'portfolioContent');
      loadTabData('disbursement', 'disbursementContent');

    } catch (err) {
      console.error('Load failed:', err);
      showNoData();
    }
  }

  function showNoData() {
    document.getElementById('no-data').classList.remove('hidden');
  }

  // Tab switching (global for onclick)
  var tabTitles = { portfolio: 'Portfolio', disbursement: 'Disbursement', collection: 'Collection' };
  window.switchEmpTab = function (tab) {
    document.querySelectorAll('.emp-tab-item').forEach(function (t) {
      t.classList.toggle('active', t.dataset.tab === tab);
    });
    document.querySelectorAll('.emp-tab-content').forEach(function (tc) {
      tc.classList.remove('active');
    });
    document.getElementById(tab + 'Tab').classList.add('active');
    var titleEl = document.getElementById('emp-header-title');
    if (titleEl) titleEl.textContent = tabTitles[tab] || tab;
  };

  loadData();
})();
