// SHA-256 hash of the admin passcode (hashed "4321")
var ADMIN_HASH = 'fe2592b42a727e977f055947385b709cc82b16b9a87f88c6abf3900d65d0cdc3';

async function _sha256(str) {
  var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
}

function authenticateAdmin(passcode) {
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
}

async function authenticateEmployee(employeeId) {
  var employee = await getEmployee(employeeId);
  if (employee) {
    localStorage.setItem('employeeId', employee.id);
    localStorage.setItem('employeeName', employee.name);
    return employee;
  }
  return null;
}

function setRoleSession(role, location) {
  localStorage.setItem('roleAuth', 'true');
  localStorage.setItem('roleName', role);
  localStorage.setItem('roleLocation', location);
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
      location: localStorage.getItem('roleLocation') || ''
    };
  }
  return {
    id: localStorage.getItem('employeeId'),
    name: localStorage.getItem('employeeName'),
    role: 'FO',
    location: ''
  };
}

function logout() {
  localStorage.removeItem('adminAuth');
  localStorage.removeItem('employeeId');
  localStorage.removeItem('employeeName');
  localStorage.removeItem('roleAuth');
  localStorage.removeItem('roleName');
  localStorage.removeItem('roleLocation');
  window.location.href = 'index.html';
}
