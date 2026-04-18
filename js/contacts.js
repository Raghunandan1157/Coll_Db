/* ───────────────────────────────────────────────────────────────────
 * Contacts cache — single source of phone numbers for FOs and managers
 * across all dashboard tabs (Collection, Portfolio, Disbursement, Hourly).
 *
 * Loads /api/employees once on page init and indexes by:
 *   - emp_id          → FO/staff phone lookup
 *   - role+location   → manager phone lookup (RM of region X, BM of branch Y, etc.)
 * ─────────────────────────────────────────────────────────────────── */
(function () {
  'use strict';

  var _byEmpId = {};      // empId(uppercase) → row
  var _byRoleLoc = {};    // role|location(uppercase) → row (first match wins)
  var _loadPromise = null;
  var _loaded = false;

  function norm(s) { return String(s == null ? '' : s).trim().toUpperCase(); }

  function indexRow(row) {
    if (row.emp_id) _byEmpId[norm(row.emp_id)] = row;
    var role = norm(row.role);
    if (!role) return;
    // Map role → location field on the row. Same employee can manage multiple
    // levels in legacy/new structures — we index every non-empty hierarchy field
    // under the role key for lookup flexibility.
    var locKeys = [];
    if (role === 'CEO')                            locKeys.push('');
    if (role === 'RM' || role === 'SM')            locKeys.push(row.region);
    if (role === 'DM' || role === 'DvM')           locKeys.push(row.division);
    if (role === 'AM')                             locKeys.push(row.area);
    if (role === 'BM')                             locKeys.push(row.branch);
    for (var i = 0; i < locKeys.length; i++) {
      var loc = norm(locKeys[i]);
      var key = role + '|' + loc;
      if (!_byRoleLoc[key]) _byRoleLoc[key] = row;
    }
  }

  function load() {
    if (_loadPromise) return _loadPromise;
    _loadPromise = fetch('/api/employees')
      .then(function (r) { return r.ok ? r.json() : []; })
      .then(function (rows) {
        if (!Array.isArray(rows)) rows = [];
        for (var i = 0; i < rows.length; i++) indexRow(rows[i]);
        _loaded = true;
      })
      .catch(function (e) {
        console.error('Contacts cache load failed:', e);
        _loaded = true; // mark loaded so UI doesn't block forever
      });
    return _loadPromise;
  }

  function getPhone(empId) {
    if (!empId) return null;
    var row = _byEmpId[norm(empId)];
    return (row && row.mobile) ? String(row.mobile).trim() : null;
  }

  function getEmployee(empId) {
    if (!empId) return null;
    return _byEmpId[norm(empId)] || null;
  }

  // role: 'RM' | 'SM' | 'DM' | 'DvM' | 'AM' | 'BM' | 'CEO'
  // location: region/division/area/branch name (string, case-insensitive)
  function getManager(role, location) {
    if (!role) return null;
    var key = norm(role) + '|' + norm(location || '');
    return _byRoleLoc[key] || null;
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // Phone icon SVG — reused across helpers
  var PHONE_SVG = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>';

  // Compact pill-style call button — for FO rows in card layouts.
  // event.stopPropagation() lets the user tap the call link without triggering
  // the row's drill-down click handler.
  function renderCallBtn(empId) {
    var phone = getPhone(empId);
    if (!phone) return '';
    return '<a href="tel:' + esc(phone) + '" class="contact-call-btn" onclick="event.stopPropagation();" title="Call ' + esc(phone) + '">' +
      PHONE_SVG +
      '<span class="contact-call-num">' + esc(phone) + '</span>' +
      '</a>';
  }

  // For non-FO subunit rows — show the manager's name + phone.
  // Falls back to nothing if no manager record found.
  function renderManagerBadge(role, location) {
    var mgr = getManager(role, location);
    if (!mgr || !mgr.mobile) return '';
    var labelMap = { CEO: 'CEO', RM: 'RM', SM: 'RM', DM: 'DM', DvM: 'DvM', AM: 'AM', BM: 'BM' };
    var lbl = labelMap[role] || role;
    var name = mgr.name || mgr.full_name || '';
    return '<a href="tel:' + esc(mgr.mobile) + '" class="contact-mgr-badge" onclick="event.stopPropagation();" title="Call ' + esc(name) + '">' +
      PHONE_SVG +
      '<span class="contact-mgr-text">' + esc(lbl) + ': ' + esc(name) + '</span>' +
      '</a>';
  }

  // Inject styles once
  if (!document.getElementById('contacts-styles')) {
    var style = document.createElement('style');
    style.id = 'contacts-styles';
    style.textContent =
      '.contact-call-btn { display:inline-flex; align-items:center; gap:4px; padding:3px 8px; border-radius:8px; background:rgba(52,211,153,0.12); color:#059669; font-size:11px; font-weight:600; text-decoration:none; border:1px solid rgba(52,211,153,0.25); transition:background .15s; line-height:1; }' +
      '.contact-call-btn:active { background:rgba(52,211,153,0.25); }' +
      '.contact-call-btn .contact-call-num { font-family:"DM Sans", sans-serif; letter-spacing:0.2px; }' +
      '.contact-mgr-badge { display:inline-flex; align-items:center; gap:4px; padding:2px 7px; border-radius:7px; background:#F1F5F9; color:#475569; font-size:10.5px; font-weight:500; text-decoration:none; line-height:1.4; margin-top:3px; }' +
      '.contact-mgr-badge:hover { background:#E2E8F0; }' +
      '.contact-cell { white-space:nowrap; }';
    document.head.appendChild(style);
  }

  window._contactsCache = {
    load: load,
    getPhone: getPhone,
    getEmployee: getEmployee,
    getManager: getManager,
    renderCallBtn: renderCallBtn,
    renderManagerBadge: renderManagerBadge,
    isLoaded: function () { return _loaded; }
  };
})();
