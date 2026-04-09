/**
 * settings.js v3 — Settings tab: data structure toggle with localStorage persistence.
 * Vanilla JS, no frameworks. Matches coding patterns from employee.js / collection.js.
 */
(function () {

  var STORAGE_KEY = 'dataStructure';
  var DEFAULT_STRUCTURE = 'new';

  /* ========== Public API ========== */

  window.getDataStructure = function () {
    return localStorage.getItem(STORAGE_KEY) || DEFAULT_STRUCTURE;
  };

  /* ========== Wave + Popup overlay ========== */

  function showStructureTransition(newValue) {
    var isNew = newValue === 'new';
    var label = isNew ? 'FY 2026-27' : 'FY 2025-26';
    var subtitle = isNew
      ? 'Region \u2192 Division \u2192 Area \u2192 Branch'
      : 'Region \u2192 District \u2192 Branch';
    var color = isNew ? '#059669' : '#D97706';
    var bgFrom = isNew ? '#ECFDF5' : '#FFFBEB';

    // Remove existing overlay
    var old = document.getElementById('structureOverlay');
    if (old) old.remove();

    // Inject keyframes once
    if (!document.getElementById('structureWaveCSS')) {
      var style = document.createElement('style');
      style.id = 'structureWaveCSS';
      style.textContent =
        '@keyframes structureWaveIn{0%{transform:scale(0);opacity:0}40%{transform:scale(1.1);opacity:1}60%{transform:scale(1)}100%{transform:scale(1)}}' +
        '@keyframes structureWaveRipple{0%{transform:scale(0.3);opacity:0.6}100%{transform:scale(2.5);opacity:0}}' +
        '@keyframes structureFadeUp{0%{opacity:0;transform:translateY(20px)}100%{opacity:1;transform:translateY(0)}}' +
        '@keyframes structureProgressBar{0%{width:0}100%{width:100%}}' +
        '@keyframes structureFadeOut{0%{opacity:1}100%{opacity:0}}';
      document.head.appendChild(style);
    }

    var overlay = document.createElement('div');
    overlay.id = 'structureOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;' +
      'background:rgba(15,23,42,0.6);backdrop-filter:blur(6px);animation:structureWaveIn 0.5s ease-out;';

    overlay.innerHTML =
      // Ripple wave
      '<div style="position:absolute;width:300px;height:300px;border-radius:50%;border:3px solid ' + color + ';' +
        'animation:structureWaveRipple 1.5s ease-out forwards;opacity:0.4;"></div>' +
      '<div style="position:absolute;width:300px;height:300px;border-radius:50%;border:2px solid ' + color + ';' +
        'animation:structureWaveRipple 1.5s 0.3s ease-out forwards;opacity:0.3;"></div>' +
      '<div style="position:absolute;width:300px;height:300px;border-radius:50%;border:1px solid ' + color + ';' +
        'animation:structureWaveRipple 1.5s 0.6s ease-out forwards;opacity:0.2;"></div>' +

      // Card
      '<div style="position:relative;background:white;border-radius:20px;padding:40px 48px;text-align:center;' +
        'box-shadow:0 25px 60px rgba(0,0,0,0.25);max-width:380px;width:90%;animation:structureFadeUp 0.5s 0.2s ease-out both;">' +

        // Icon circle
        '<div style="width:64px;height:64px;border-radius:50%;background:' + bgFrom + ';display:flex;align-items:center;' +
          'justify-content:center;margin:0 auto 20px;border:2px solid ' + color + ';">' +
          '<svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="' + color + '" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
            '<path d="M4 4h6v6H4zM14 4h6v6h-6zM4 14h6v6H4zM14 14h6v6h-6z"/>' +
          '</svg>' +
        '</div>' +

        // Title
        '<div style="font-size:13px;font-weight:500;color:#94A3B8;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:8px;">Switching to</div>' +
        '<div style="font-size:28px;font-weight:700;color:#0F172A;margin-bottom:6px;">' + label + '</div>' +
        '<div style="font-size:14px;color:#64748B;margin-bottom:24px;">' + subtitle + '</div>' +

        // Progress bar
        '<div style="height:4px;background:#E2E8F0;border-radius:4px;overflow:hidden;margin-bottom:16px;">' +
          '<div style="height:100%;background:' + color + ';border-radius:4px;animation:structureProgressBar 2.5s ease-in-out forwards;"></div>' +
        '</div>' +

        '<div style="font-size:12px;color:#94A3B8;">Loading new structure...</div>' +
      '</div>';

    document.body.appendChild(overlay);

    // Fade out and reload after 3s
    setTimeout(function () {
      overlay.style.animation = 'structureFadeOut 0.4s ease-out forwards';
      setTimeout(function () {
        overlay.remove();
        window.location.reload();
      }, 400);
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

    // Show wave transition + popup
    showStructureTransition(value);

    // Reset to original login role (pop entire drill-down stack)
    var stack = [];
    try { stack = JSON.parse(localStorage.getItem('roleNavStack') || '[]'); } catch (e) {}
    if (stack.length) {
      var root = stack[0];
      localStorage.setItem('roleName', root.role || 'CEO');
      localStorage.setItem('roleLocation', root.location || '');
      if (root.id) { localStorage.setItem('employeeId', root.id); localStorage.setItem('employeeName', root.name || ''); }
      else { localStorage.removeItem('employeeId'); localStorage.removeItem('employeeName'); }
    }
    localStorage.removeItem('roleNavStack');
    localStorage.removeItem('pfMonth');
    localStorage.removeItem('activeTab');
    localStorage.removeItem('dbMonth');
    localStorage.removeItem('dbProduct');
    localStorage.removeItem('roleState');
    localStorage.removeItem('roleDivision');
    localStorage.removeItem('roleArea');
    // Reload happens inside showStructureTransition after 3s
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

    // Tab key toggles between old and new structure
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Tab') {
        e.preventDefault();
        var current = window.getDataStructure();
        if (current === 'new') {
          setDataStructure('fy2025-26');
        } else {
          setDataStructure('new');
        }
      }
    });
  }

  /* ========== Bootstrap ========== */

  window._initSettings = initSettings;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSettings);
  } else {
    initSettings();
  }

})();
