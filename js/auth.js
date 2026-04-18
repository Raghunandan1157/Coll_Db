// SHA-256 hash of the admin passcode (hashed "4321")
var ADMIN_HASH = 'fe2592b42a727e977f055947385b709cc82b16b9a87f88c6abf3900d65d0cdc3';

// Pure-JS SHA-256 fallback for non-secure contexts (HTTP)
var _sha256_fallback = (function () {
  function rr(n, x) { return (x >>> n) | (x << (32 - n)); }
  var K = [
    0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
    0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
    0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
    0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
    0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2
  ];
  return function sha256(str) {
    var msg = unescape(encodeURIComponent(str));
    var n = msg.length, i, j;
    var words = [];
    for (i = 0; i < n; i++) words[i >> 2] |= (msg.charCodeAt(i) & 0xff) << (24 - (i % 4) * 8);
    words[n >> 2] |= 0x80 << (24 - (n % 4) * 8);
    var bitLen = n * 8;
    var totalLen = (((n + 8) >> 6) + 1) * 16;
    words[totalLen - 1] = bitLen;
    var H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19];
    for (i = 0; i < totalLen; i += 16) {
      var W = [], a=H[0],b=H[1],c=H[2],d=H[3],e=H[4],f=H[5],g=H[6],h=H[7];
      for (j = 0; j < 64; j++) {
        if (j < 16) W[j] = words[i + j] | 0;
        else {
          var s0 = rr(7,W[j-15]) ^ rr(18,W[j-15]) ^ (W[j-15]>>>3);
          var s1 = rr(17,W[j-2]) ^ rr(19,W[j-2]) ^ (W[j-2]>>>10);
          W[j] = (W[j-16] + s0 + W[j-7] + s1) | 0;
        }
        var S1 = rr(6,e) ^ rr(11,e) ^ rr(25,e);
        var ch = (e & f) ^ (~e & g);
        var t1 = (h + S1 + ch + K[j] + W[j]) | 0;
        var S0 = rr(2,a) ^ rr(13,a) ^ rr(22,a);
        var maj = (a & b) ^ (a & c) ^ (b & c);
        var t2 = (S0 + maj) | 0;
        h=g; g=f; f=e; e=(d+t1)|0; d=c; c=b; b=a; a=(t1+t2)|0;
      }
      H[0]=(H[0]+a)|0; H[1]=(H[1]+b)|0; H[2]=(H[2]+c)|0; H[3]=(H[3]+d)|0;
      H[4]=(H[4]+e)|0; H[5]=(H[5]+f)|0; H[6]=(H[6]+g)|0; H[7]=(H[7]+h)|0;
    }
    var hex = '';
    for (i = 0; i < 8; i++) {
      for (j = 3; j >= 0; j--) hex += ((H[i] >> (j*8)) & 0xff).toString(16).padStart(2,'0');
    }
    return hex;
  };
})();

function authenticateAdmin(passcode) {
  // Use crypto.subtle if available (HTTPS), otherwise pure-JS fallback (HTTP)
  if (typeof crypto !== 'undefined' && crypto.subtle) {
    var encoder = new TextEncoder();
    var data = encoder.encode(passcode);
    return crypto.subtle.digest('SHA-256', data).then(function (buf) {
      var hash = Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
      if (hash === ADMIN_HASH) {
        localStorage.setItem('adminAuth', 'true');
        return true;
      }
      return false;
    });
  } else {
    // Fallback for non-secure contexts (HTTP)
    var hash = _sha256_fallback(passcode);
    if (hash === ADMIN_HASH) {
      localStorage.setItem('adminAuth', 'true');
      return Promise.resolve(true);
    }
    return Promise.resolve(false);
  }
}

async function authenticateEmployee(employeeId) {
  // Try IndexedDB first
  try {
    var employee = await getEmployee(employeeId);
    if (employee) {
      localStorage.setItem('employeeId', employee.id);
      localStorage.setItem('employeeName', employee.name);
      return employee;
    }
  } catch(e) {}
  // Fallback: fetch from API directly
  try {
    var res = await fetch('/api/employees');
    var data = await res.json();
    var emp = data.find(function(e){ return e.emp_id === employeeId; });
    if (emp) {
      localStorage.setItem('employeeId', emp.emp_id);
      localStorage.setItem('employeeName', emp.name);
      return {id: emp.emp_id, name: emp.name};
    }
  } catch(e) {}
  // Last resort: just set the ID directly
  localStorage.setItem('employeeId', employeeId);
  localStorage.setItem('employeeName', employeeId);
  return {id: employeeId, name: employeeId};
}

function isNewStructure() {
  return localStorage.getItem('dataStructure') === 'new';
}

function setRoleSession(role, location, structure) {
  localStorage.setItem('roleAuth', 'true');
  localStorage.setItem('roleName', role);
  localStorage.setItem('roleLocation', location);
  // Store which structure this session uses so downstream pages can branch correctly
  var struct = structure || localStorage.getItem('dataStructure') || 'fy2025-26';
  localStorage.setItem('roleStructure', struct);
  // For new-structure roles, map location to the appropriate level key
  if (struct === 'new') {
    localStorage.removeItem('roleState');
    localStorage.removeItem('roleDivision');
    localStorage.removeItem('roleArea');
    if (role === 'SM') localStorage.setItem('roleState', location);
    else if (role === 'DvM') localStorage.setItem('roleDivision', location);
    else if (role === 'AM') localStorage.setItem('roleArea', location);
  }
}

function pushRoleNav(childRole, childLocation) {
  var stack = getRoleNavStack();
  var current = getEmployeeSession();
  stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  if (childRole !== 'FO') {
    localStorage.removeItem('employeeId');
    localStorage.removeItem('employeeName');
    setRoleSession(childRole, childLocation);
  } else {
    localStorage.removeItem('roleAuth');
    localStorage.removeItem('roleName');
    localStorage.removeItem('roleLocation');
    localStorage.setItem('employeeId', childLocation);
    localStorage.setItem('employeeName', childRole === 'FO' ? childLocation : '');
  }
}

function popRoleNav() {
  var stack = getRoleNavStack();
  if (!stack.length) return null;
  var prev = stack.pop();
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  localStorage.removeItem('employeeId');
  localStorage.removeItem('employeeName');
  localStorage.removeItem('roleAuth');
  localStorage.removeItem('roleName');
  localStorage.removeItem('roleLocation');
  if (prev.role === 'FO') {
    localStorage.setItem('employeeId', prev.id);
    localStorage.setItem('employeeName', prev.name);
  } else {
    setRoleSession(prev.role, prev.location);
  }
  return prev;
}

function restoreNavTo(index) {
  var stack = getRoleNavStack();
  if (index < 0 || index >= stack.length) return;
  var target = stack[index];
  stack = stack.slice(0, index);
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  localStorage.removeItem('employeeId');
  localStorage.removeItem('employeeName');
  localStorage.removeItem('roleAuth');
  localStorage.removeItem('roleName');
  localStorage.removeItem('roleLocation');
  if (target.role === 'FO') {
    localStorage.setItem('employeeId', target.id);
    localStorage.setItem('employeeName', target.name);
  } else {
    setRoleSession(target.role, target.location);
  }
}

function getRoleNavStack() {
  try { return JSON.parse(localStorage.getItem('roleNavStack') || '[]'); }
  catch (e) { return []; }
}

function clearRoleNavStack() {
  localStorage.removeItem('roleNavStack');
}

function isAdminAuthenticated() {
  return localStorage.getItem('adminAuth') === 'true';
}

function isEmployeeAuthenticated() {
  return !!localStorage.getItem('employeeId') || localStorage.getItem('roleAuth') === 'true';
}

function getEmployeeSession() {
  var role = localStorage.getItem('roleName');
  if (role) {
    return {
      id: null,
      name: localStorage.getItem('roleLocation') || '',
      role: role,
      location: localStorage.getItem('roleLocation') || '',
      structure: localStorage.getItem('roleStructure') || 'fy2025-26',
      state: localStorage.getItem('roleState') || null,
      division: localStorage.getItem('roleDivision') || null,
      area: localStorage.getItem('roleArea') || null
    };
  }
  return {
    id: localStorage.getItem('employeeId'),
    name: localStorage.getItem('employeeName'),
    role: 'FO',
    location: '',
    structure: localStorage.getItem('dataStructure') || 'fy2025-26',
    state: null,
    division: null,
    area: null
  };
}

function logout() {
  localStorage.removeItem('adminAuth');
  localStorage.removeItem('employeeId');
  localStorage.removeItem('employeeName');
  localStorage.removeItem('roleAuth');
  localStorage.removeItem('roleName');
  localStorage.removeItem('roleLocation');
  localStorage.removeItem('roleStructure');
  localStorage.removeItem('roleState');
  localStorage.removeItem('roleDivision');
  localStorage.removeItem('roleArea');
  localStorage.removeItem('roleNavStack');
  localStorage.removeItem('activeTab');
  localStorage.removeItem('collProduct');
  localStorage.removeItem('collView');
  window.location.href = 'index.html';
}
