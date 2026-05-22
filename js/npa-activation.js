/**
 * npa-activation.js — NPA Activation tab.
 * Upload an NPA xlsx, pick a run, drill Region → Division → Area → Branch → rows.
 * Container: npaactivationContent
 */
(function () {
  var session = (typeof getEmployeeSession === 'function') ? getEmployeeSession() : {};

  /* ===== State ===== */
  var state = {
    runs: [],
    runId: null,
    level: 'region',   // 'region' | 'division' | 'area' | 'branch' | 'rows'
    crumb: { region: null, division: null, area: null, branch: null },
    data: [],
    rows: [],
    loading: false,
    error: null,
    uploading: false,
    uploadMsg: null
  };

  /* ===== Helpers ===== */
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return '&#' + c.charCodeAt(0) + ';'; }); }
  function numVal(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string') { var p = parseFloat(v); return isFinite(p) ? p : 0; }
    return 0;
  }
  function fmtInt(v) {
    var n = numVal(v);
    return n.toLocaleString('en-IN');
  }
  function fmtAmt(n) {
    n = numVal(n);
    if (n >= 10000000) return '₹' + (n / 10000000).toFixed(2) + ' Cr';
    if (n >= 100000)   return '₹' + (n / 100000).toFixed(2) + ' L';
    if (n >= 1000)     return '₹' + (n / 1000).toFixed(1) + ' K';
    return '₹' + Math.round(n).toLocaleString('en-IN');
  }
  function fmtDT(s) {
    if (!s) return '';
    try {
      var d = new Date(s);
      if (isNaN(d.getTime())) return String(s);
      return d.toLocaleString('en-IN', { year: 'numeric', month: 'short', day: '2-digit', hour: '2-digit', minute: '2-digit' });
    } catch (e) { return String(s); }
  }
  function apiFetch(url) {
    return fetch(url, { credentials: 'same-origin' }).then(function (r) {
      if (!r.ok) throw new Error('API ' + r.status);
      return r.json();
    });
  }

  /* ===== Data loaders ===== */
  function loadRuns() {
    state.loading = true; state.error = null; render();
    return apiFetch('/api/npa/runs').then(function (res) {
      var runs = Array.isArray(res) ? res : (res && res.runs) || [];
      state.runs = runs;
      if (runs.length && !state.runId) state.runId = runs[0].run_id || runs[0].id || runs[0];
      state.loading = false;
      render();
      if (state.runId) loadLevel();
    }).catch(function (e) {
      state.loading = false;
      state.error = 'Failed to load runs: ' + (e && e.message ? e.message : e);
      render();
    });
  }

  function loadLevel() {
    if (!state.runId) { state.data = []; render(); return; }
    state.loading = true; state.error = null; state.rows = []; render();
    var qp = 'run_id=' + encodeURIComponent(state.runId);
    var url;
    if (state.level === 'region') {
      url = '/api/npa/by-region?' + qp;
    } else if (state.level === 'division') {
      url = '/api/npa/by-division?' + qp + '&region=' + encodeURIComponent(state.crumb.region || '');
    } else if (state.level === 'area') {
      url = '/api/npa/by-area?' + qp + '&division=' + encodeURIComponent(state.crumb.division || '');
    } else if (state.level === 'branch') {
      url = '/api/npa/by-branch?' + qp + '&area=' + encodeURIComponent(state.crumb.area || '');
    } else if (state.level === 'rows') {
      url = '/api/npa/rows?' + qp + '&branch=' + encodeURIComponent(state.crumb.branch || '');
    }
    return apiFetch(url).then(function (res) {
      var arr = Array.isArray(res) ? res : (res && (res.rows || res.data)) || [];
      if (state.level === 'rows') state.rows = arr;
      else state.data = arr;
      state.loading = false;
      render();
    }).catch(function (e) {
      state.loading = false;
      state.error = 'Failed to load ' + state.level + ': ' + (e && e.message ? e.message : e);
      render();
    });
  }

  /* ===== Upload ===== */
  function doUpload(file) {
    if (!file) return;
    state.uploading = true; state.uploadMsg = null; render();
    var fd = new FormData();
    fd.append('file', file);
    fetch('/api/npa/upload', { method: 'POST', body: fd, credentials: 'same-origin' })
      .then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      })
      .then(function (res) {
        state.uploading = false;
        state.uploadMsg = { type: 'success', text: 'Upload OK' + (res && res.run_id ? ' (run ' + res.run_id + ')' : '') };
        render();
        loadRuns();
      })
      .catch(function (e) {
        state.uploading = false;
        state.uploadMsg = { type: 'error', text: 'Upload failed: ' + (e && e.message ? e.message : e) };
        render();
      });
  }

  /* ===== Drilldown navigation ===== */
  function drillInto(row) {
    if (state.level === 'region') {
      state.crumb.region = row.region || row.name || row.region_name;
      state.level = 'division';
    } else if (state.level === 'division') {
      state.crumb.division = row.division || row.name || row.division_name;
      state.level = 'area';
    } else if (state.level === 'area') {
      state.crumb.area = row.area || row.name || row.area_name;
      state.level = 'branch';
    } else if (state.level === 'branch') {
      state.crumb.branch = row.branch || row.name || row.branch_name;
      state.level = 'rows';
    } else {
      return;
    }
    loadLevel();
  }

  function jumpTo(level) {
    state.level = level;
    if (level === 'region') { state.crumb.region = null; state.crumb.division = null; state.crumb.area = null; state.crumb.branch = null; }
    else if (level === 'division') { state.crumb.division = null; state.crumb.area = null; state.crumb.branch = null; }
    else if (level === 'area') { state.crumb.area = null; state.crumb.branch = null; }
    else if (level === 'branch') { state.crumb.branch = null; }
    loadLevel();
  }

  /* ===== Render ===== */
  function levelHeader() {
    if (state.level === 'region') return 'Region';
    if (state.level === 'division') return 'Division';
    if (state.level === 'area') return 'Area';
    if (state.level === 'branch') return 'Branch';
    return 'Accounts';
  }

  function crumbHtml() {
    var parts = [];
    parts.push('<span class="npa-crumb-item" data-npa-jump="region" style="cursor:pointer;color:#059669;font-weight:600;">All Regions</span>');
    if (state.crumb.region) parts.push('<span class="npa-crumb-item" data-npa-jump="division" style="cursor:pointer;color:#059669;font-weight:600;">' + esc(state.crumb.region) + '</span>');
    if (state.crumb.division) parts.push('<span class="npa-crumb-item" data-npa-jump="area" style="cursor:pointer;color:#059669;font-weight:600;">' + esc(state.crumb.division) + '</span>');
    if (state.crumb.area) parts.push('<span class="npa-crumb-item" data-npa-jump="branch" style="cursor:pointer;color:#059669;font-weight:600;">' + esc(state.crumb.area) + '</span>');
    if (state.crumb.branch) parts.push('<span style="color:#1E293B;font-weight:700;">' + esc(state.crumb.branch) + '</span>');
    return parts.join('<span style="color:#CBD5E1;margin:0 6px;">/</span>');
  }

  function tableHtml() {
    if (state.level === 'rows') return rowsTableHtml();
    var header = levelHeader();
    var rows = state.data || [];
    if (!rows.length) {
      return '<div style="padding:32px;text-align:center;color:#94A3B8;font-size:14px;">No data for this run.</div>';
    }
    var html = '<div style="overflow:auto;border:1px solid #E2E8F0;border-radius:12px;background:#fff;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<thead><tr style="background:#F8FAFC;">';
    html += '<th style="padding:10px 14px;text-align:left;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:11px;">' + esc(header) + '</th>';
    html += '<th style="padding:10px 14px;text-align:right;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:11px;">NPA Count</th>';
    html += '<th style="padding:10px 14px;text-align:right;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:11px;">Branch Count</th>';
    html += '<th style="padding:10px 14px;text-align:right;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:11px;">Account Count</th>';
    html += '<th style="padding:10px 14px;text-align:right;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:11px;">Collection Sum</th>';
    html += '</tr></thead><tbody>';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var name = r[state.level] || r.name || r[state.level + '_name'] || '';
      var npaCount = r.npa_count != null ? r.npa_count : r.count;
      var branchCount = r.branch_count != null ? r.branch_count : '';
      var acctCount = r.account_count != null ? r.account_count : (r.accounts != null ? r.accounts : '');
      var collSum = r.collection_sum != null ? r.collection_sum : (r.collection != null ? r.collection : 0);
      html += '<tr data-npa-row="' + i + '" style="cursor:pointer;border-bottom:1px solid #F1F5F9;">';
      html += '<td style="padding:10px 14px;color:#059669;font-weight:600;">' + esc(name) + '</td>';
      html += '<td style="padding:10px 14px;text-align:right;color:#334155;">' + fmtInt(npaCount) + '</td>';
      html += '<td style="padding:10px 14px;text-align:right;color:#334155;">' + (branchCount === '' ? '—' : fmtInt(branchCount)) + '</td>';
      html += '<td style="padding:10px 14px;text-align:right;color:#334155;">' + (acctCount === '' ? '—' : fmtInt(acctCount)) + '</td>';
      html += '<td style="padding:10px 14px;text-align:right;color:#1E293B;font-weight:600;">' + fmtAmt(collSum) + '</td>';
      html += '</tr>';
    }
    html += '</tbody></table></div>';
    return html;
  }

  function rowsTableHtml() {
    var rows = state.rows || [];
    if (!rows.length) {
      return '<div style="padding:32px;text-align:center;color:#94A3B8;font-size:14px;">No accounts for this branch.</div>';
    }
    // Derive columns from first row keys
    var keys = Object.keys(rows[0]);
    var html = '<div style="overflow:auto;border:1px solid #E2E8F0;border-radius:12px;background:#fff;max-height:70vh;">';
    html += '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
    html += '<thead><tr style="background:#F8FAFC;position:sticky;top:0;z-index:1;">';
    for (var k = 0; k < keys.length; k++) {
      html += '<th style="padding:8px 12px;text-align:left;font-weight:700;color:#64748B;border-bottom:1px solid #E2E8F0;text-transform:uppercase;letter-spacing:.3px;font-size:10px;white-space:nowrap;">' + esc(keys[k]) + '</th>';
    }
    html += '</tr></thead><tbody>';
    for (var i = 0; i < rows.length; i++) {
      html += '<tr style="border-bottom:1px solid #F1F5F9;">';
      for (var kk = 0; kk < keys.length; kk++) {
        var v = rows[i][keys[kk]];
        html += '<td style="padding:7px 12px;color:#334155;white-space:nowrap;">' + esc(v == null ? '' : v) + '</td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table></div>';
    return html;
  }

  function runSelectHtml() {
    if (!state.runs.length) {
      return '<span style="font-size:13px;color:#94A3B8;">No runs yet — upload a file to start.</span>';
    }
    var html = '<label style="font-size:12px;color:#64748B;font-weight:600;margin-right:8px;">Run:</label>';
    html += '<select id="npaRunSelect" style="padding:7px 12px;border:1px solid #E2E8F0;border-radius:10px;font-size:13px;color:#334155;background:#fff;outline:none;min-width:260px;">';
    for (var i = 0; i < state.runs.length; i++) {
      var r = state.runs[i];
      var id = r.run_id || r.id || r;
      var label = (r.filename ? r.filename + ' — ' : '') + fmtDT(r.created_at || r.uploaded_at || r.date || '');
      if (!label.trim() || label.trim() === '—') label = 'Run ' + id;
      html += '<option value="' + esc(id) + '"' + (String(id) === String(state.runId) ? ' selected' : '') + '>' + esc(label) + '</option>';
    }
    html += '</select>';
    return html;
  }

  function uploadHtml() {
    var msg = '';
    if (state.uploadMsg) {
      var color = state.uploadMsg.type === 'success' ? '#059669' : '#EF4444';
      msg = '<span style="margin-left:12px;font-size:12px;color:' + color + ';font-weight:600;">' + esc(state.uploadMsg.text) + '</span>';
    }
    var btnLabel = state.uploading ? 'Uploading…' : 'Upload NPA xlsx';
    return '' +
      '<div style="display:flex;align-items:center;flex-wrap:wrap;gap:10px;">' +
        '<input id="npaFileInput" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style="font-size:13px;" ' + (state.uploading ? 'disabled' : '') + '>' +
        '<button id="npaUploadBtn" style="padding:8px 16px;background:#059669;color:#fff;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;' + (state.uploading ? 'opacity:.6;cursor:not-allowed;' : '') + '" ' + (state.uploading ? 'disabled' : '') + '>' + esc(btnLabel) + '</button>' +
        msg +
      '</div>';
  }

  function render() {
    var c = document.getElementById('npaactivationContent');
    if (!c) return;
    var html = '';
    html += '<div class="desktop-content-area" style="padding:8px 4px;">';

    // Header + upload
    html += '<div class="emp-fade" style="background:#fff;border-radius:14px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:16px 20px;margin-bottom:14px;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">';
    html += '<div>';
    html += '<div style="font-size:16px;font-weight:700;color:#1E293B;">NPA Activation</div>';
    html += '<div style="font-size:11px;color:#94A3B8;margin-top:4px;">Upload & drill Region → Division → Area → Branch → Accounts.</div>';
    html += '</div>';
    html += '<div>' + runSelectHtml() + '</div>';
    html += '</div>';
    html += '<div style="margin-top:14px;">' + uploadHtml() + '</div>';
    html += '</div>';

    // Breadcrumb
    html += '<div style="background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:10px 16px;margin-bottom:12px;font-size:13px;">';
    html += crumbHtml();
    html += '</div>';

    // Data area
    if (state.error) {
      html += '<div style="padding:20px;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:10px;color:#B91C1C;font-size:13px;">' + esc(state.error) + '</div>';
    } else if (state.loading) {
      html += '<div style="padding:40px;text-align:center;color:#94A3B8;font-size:13px;">Loading…</div>';
    } else if (!state.runId) {
      html += '<div style="padding:40px;text-align:center;color:#94A3B8;font-size:13px;">Upload an NPA xlsx to begin.</div>';
    } else {
      html += tableHtml();
    }

    html += '</div>';
    c.innerHTML = html;
    bind();
  }

  function bind() {
    var c = document.getElementById('npaactivationContent');
    if (!c) return;

    var sel = c.querySelector('#npaRunSelect');
    if (sel) sel.onchange = function () {
      state.runId = sel.value;
      state.level = 'region';
      state.crumb = { region: null, division: null, area: null, branch: null };
      loadLevel();
    };

    var btn = c.querySelector('#npaUploadBtn');
    var fileIn = c.querySelector('#npaFileInput');
    if (btn && fileIn) btn.onclick = function () {
      var f = fileIn.files && fileIn.files[0];
      if (!f) { state.uploadMsg = { type: 'error', text: 'Select a file first.' }; render(); return; }
      doUpload(f);
    };

    var crumbItems = c.querySelectorAll('[data-npa-jump]');
    for (var i = 0; i < crumbItems.length; i++) {
      (function (el) {
        el.onclick = function () { jumpTo(el.getAttribute('data-npa-jump')); };
      })(crumbItems[i]);
    }

    var rowEls = c.querySelectorAll('[data-npa-row]');
    for (var j = 0; j < rowEls.length; j++) {
      (function (el) {
        var idx = parseInt(el.getAttribute('data-npa-row'), 10);
        el.onclick = function () {
          var row = state.data[idx];
          if (row) drillInto(row);
        };
        el.onmouseenter = function () { el.style.background = '#F0FDF4'; };
        el.onmouseleave = function () { el.style.background = ''; };
      })(rowEls[j]);
    }
  }

  /* ===== Public entry ===== */
  window._loadNPAActivationTab = function () {
    // Re-render immediately (cheap) and refresh runs list on each open.
    render();
    loadRuns();
  };
})();
