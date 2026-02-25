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

function pushRoleNav(childRole, childLocation) {
  var stack = getRoleNavStack();
  var current = getEmployeeSession();
  stack.push({ role: current.role, location: current.location, name: current.name, id: current.id });
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  // Clear FO keys if navigating to a role
  if (childRole !== 'FO') {
    localStorage.removeItem('employeeId');
    localStorage.removeItem('employeeName');
    setRoleSession(childRole, childLocation);
  } else {
    // FO: store employee info
    localStorage.removeItem('roleAuth');
    localStorage.removeItem('roleName');
    localStorage.removeItem('roleLocation');
    localStorage.setItem('employeeId', childLocation); // childLocation = empId for FO
    localStorage.setItem('employeeName', childRole === 'FO' ? childLocation : '');
  }
}

function popRoleNav() {
  var stack = getRoleNavStack();
  if (!stack.length) return null;
  var prev = stack.pop();
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  // Restore session
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
  // Trim stack to that point
  stack = stack.slice(0, index);
  localStorage.setItem('roleNavStack', JSON.stringify(stack));
  // Restore session
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
  localStorage.removeItem('roleNavStack');
  window.location.href = 'index.html';
}
