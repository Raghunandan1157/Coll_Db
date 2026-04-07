/**
 * settings.js v2 — Settings tab: data structure toggle with localStorage persistence.
 * Vanilla JS, no frameworks. Matches coding patterns from employee.js / collection.js.
 */
(function () {

  var STORAGE_KEY = 'dataStructure';
  var DEFAULT_STRUCTURE = 'new';

  /* ========== Public API ========== */

  window.getDataStructure = function () {
    return localStorage.getItem(STORAGE_KEY) || DEFAULT_STRUCTURE;
  };

  /* ========== Toast ========== */

  function showToast(msg) {
    var toast = document.getElementById('structureToast');
    if (!toast) return;
    toast.textContent = msg;
    toast.style.cssText = 'position:fixed;bottom:80px;left:50%;transform:translateX(-50%) translateY(0);' +
      'background:#1E293B;color:#F8FAFC;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:500;' +
      'box-shadow:0 4px 16px rgba(0,0,0,0.25);z-index:9999;opacity:1;transition:opacity 0.4s ease;pointer-events:none;';
    clearTimeout(toast._hideTimer);
    toast._hideTimer = setTimeout(function () {
      toast.style.opacity = '0';
    }, 2600);
  }

  /* ========== UI updater ========== */

  function updateUI(value) {
    var oldBtn = document.getElementById('structureOld');
    var newBtn = document.getElementById('structureNew');
    var oldBadge = document.getElementById('structureOldBadge');
    var newBadge = document.getElementById('structureNewBadge');
    if (!oldBtn || !newBtn) return;

    if (value === 'fy2025-26') {
      oldBtn.classList.add('active');
      newBtn.classList.remove('active');
      if (oldBadge) oldBadge.style.display = '';
      if (newBadge) newBadge.style.display = 'none';
    } else {
      newBtn.classList.add('active');
      oldBtn.classList.remove('active');
      if (newBadge) newBadge.style.display = '';
      if (oldBadge) oldBadge.style.display = 'none';
    }
  }

  /* ========== setDataStructure ========== */

  function setDataStructure(value) {
    var current = localStorage.getItem(STORAGE_KEY) || DEFAULT_STRUCTURE;
    if (value === current) return;
    localStorage.setItem(STORAGE_KEY, value);
    updateUI(value);
    var label = value === 'fy2025-26' ? 'FY 2025-26 Structure' : 'FY 2026-27 Structure';
    showToast('Switching to ' + label + '...');
    // Reset to original login role (pop entire drill-down stack)
    var stack = [];
    try { stack = JSON.parse(localStorage.getItem('roleNavStack') || '[]'); } catch (e) {}
    if (stack.length) {
      var root = stack[0]; // original login role
      localStorage.setItem('roleName', root.role || 'CEO');
      localStorage.setItem('roleLocation', root.location || '');
      if (root.id) { localStorage.setItem('employeeId', root.id); localStorage.setItem('employeeName', root.name || ''); }
      else { localStorage.removeItem('employeeId'); localStorage.removeItem('employeeName'); }
    }
    localStorage.removeItem('roleNavStack');
    // Clear stale state before reload
    localStorage.removeItem('pfMonth');
    localStorage.removeItem('activeTab');
    localStorage.removeItem('dbMonth');
    localStorage.removeItem('dbProduct');
    localStorage.removeItem('roleState');
    localStorage.removeItem('roleDivision');
    localStorage.removeItem('roleArea');
    // Reload page so all tab modules re-initialize with new structure
    setTimeout(function () { window.location.reload(); }, 500);
  }

  /* ========== initSettings ========== */

  function initSettings() {
    var current = window.getDataStructure();
    updateUI(current);

    var oldBtn = document.getElementById('structureOld');
    var newBtn = document.getElementById('structureNew');

    if (oldBtn) {
      oldBtn.onclick = function () {
        setDataStructure('fy2025-26');
      };
    }

    if (newBtn) {
      newBtn.onclick = function () {
        setDataStructure('new');
      };
    }
  }

  /* ========== Bootstrap ========== */

  window._initSettings = initSettings;

  // Auto-init if DOM is already ready, else wait
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSettings);
  } else {
    initSettings();
  }

})();
