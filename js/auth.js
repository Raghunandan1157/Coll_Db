var ADMIN_PASSCODE = "4321";

function authenticateAdmin(passcode) {
  if (passcode === ADMIN_PASSCODE) {
    sessionStorage.setItem('adminAuth', 'true');
    return true;
  }
  return false;
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
