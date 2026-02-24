// SHA-256 hash of the admin passcode (hashed "4321")
var ADMIN_HASH = 'fe2592b42a727e977f055947385b709cc82b16b9a87f88c6abf3900d65d0cdc3';

async function _sha256(str) {
  var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
}

function authenticateAdmin(passcode) {
  // Sync check using pre-computed hash
  var encoder = new TextEncoder();
  var data = encoder.encode(passcode);
  // Use sync comparison for prompt flow
  return crypto.subtle.digest('SHA-256', data).then(function (buf) {
    var hash = Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
    if (hash === ADMIN_HASH) {
      sessionStorage.setItem('adminAuth', 'true');
      return true;
    }
    return false;
  });
}

async function authenticateEmployee(employeeId) {
  var employee = await getEmployee(employeeId);
  if (employee) {
    sessionStorage.setItem('employeeId', employee.id);
    sessionStorage.setItem('employeeName', employee.name);
    return employee;
  }
  return null;
}

function isAdminAuthenticated() {
  return sessionStorage.getItem('adminAuth') === 'true';
}

function isEmployeeAuthenticated() {
  return !!sessionStorage.getItem('employeeId');
}

function getEmployeeSession() {
  return {
    id: sessionStorage.getItem('employeeId'),
    name: sessionStorage.getItem('employeeName')
  };
}

function logout() {
  sessionStorage.clear();
  window.location.href = 'index.html';
}
