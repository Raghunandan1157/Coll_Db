/**
 * admin.js — Admin panel logic for uploading reports,
 * extracting employees, and rendering the full workbook.
 */
(function () {
  // Auth guard
  if (!isAdminAuthenticated()) {
    window.location.href = 'index.html';
    return;
  }

  /* ---------- GitHub sync config ---------- */
  var GH_OWNER = 'Raghunandan1157';
  var GH_REPO = 'Coll_Db';
  var GH_FILE_PATH = 'data/report.xlsx';
  var GH_API = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/contents/' + GH_FILE_PATH;

  function getGHToken() { return localStorage.getItem('gh_token'); }
  function saveGHToken(t) { localStorage.setItem('gh_token', t); }

  function promptForToken() {
    var t = prompt('Enter your GitHub token to enable cloud sync.\nThis is saved only in this browser.');
    if (t && t.trim()) { saveGHToken(t.trim()); return t.trim(); }
    return null;
  }

  /**
   * Push an ArrayBuffer to GitHub as data/report.xlsx.
   * Gets the current file SHA first (required for updates), then PUTs the new content.
   */
  async function pushToGitHub(arrayBuffer) {
    var token = getGHToken();
    if (!token) token = promptForToken();
    if (!token) throw new Error('No GitHub token — file saved locally only.');

    // Get current file SHA
    var shaRes = await fetch(GH_API, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    var sha = null;
    if (shaRes.ok) {
      var shaData = await shaRes.json();
      sha = shaData.sha;
    }

    // Convert ArrayBuffer to base64
    var bytes = new Uint8Array(arrayBuffer);
    var binary = '';
    for (var i = 0; i < bytes.length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    var base64 = btoa(binary);

    // Push to GitHub
    var body = {
      message: 'Update report — ' + new Date().toLocaleDateString('en-IN'),
      content: base64
    };
    if (sha) body.sha = sha;

    var putRes = await fetch(GH_API, {
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    if (!putRes.ok) {
      var err = await putRes.json().catch(function () { return {}; });
      throw new Error('GitHub push failed: ' + (err.message || putRes.status));
    }
    return true;
  }

  // DOM references
  var logoutBtn = document.getElementById('logout-btn');
  var uploadZone = document.getElementById('upload-zone');
  var fileInput = document.getElementById('file-input');
  var fileInfo = document.getElementById('file-info');
  var fileNameSpan = document.getElementById('file-name');
  var clearBtn = document.getElementById('clear-btn');
  var uploadStatus = document.getElementById('upload-status');
  var employeesSection = document.getElementById('employees-section');
  var empCount = document.getElementById('emp-count');
  var employeeTbody = document.getElementById('employee-tbody');
  var reportSection = document.getElementById('report-section');
  var sheetTabs = document.getElementById('sheet-tabs');
  var reportContainer = document.getElementById('report-container');

  // Logout
  logoutBtn.onclick = logout;

  // Upload zone — click to browse
  uploadZone.addEventListener('click', function () {
    fileInput.click();
  });

  // Drag & drop events
  uploadZone.addEventListener('dragover', function (e) {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
  });

  uploadZone.addEventListener('dragleave', function (e) {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
  });

  uploadZone.addEventListener('drop', function (e) {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    var files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  });

  // File input change
  fileInput.addEventListener('change', function () {
    if (fileInput.files.length > 0) {
      handleFile(fileInput.files[0]);
    }
  });

  /**
   * Process an uploaded Excel file.
   * @param {File} file
   */
  async function handleFile(file) {
    // Validate extension
    var name = file.name.toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      showStatus('Please upload an Excel file (.xlsx or .xls).', false);
      return;
    }

    // Show file info bar
    fileNameSpan.textContent = file.name;
    fileInfo.hidden = false;
    uploadZone.style.display = 'none';

    try {
      // Read file as ArrayBuffer
      var arrayBuffer = await readFileAsArrayBuffer(file);

      // Clear old data before saving new
      await clearAllData();
      _fetchingRemote = null;

      // Save to IndexedDB
      await saveWorkbook(arrayBuffer, file.name);

      // Parse with SheetJS
      var workbook = XLSX.read(new Uint8Array(arrayBuffer), {
        type: 'array',
        cellStyles: true,
        cellFormula: false,
        cellNF: true
      });

      // Extract employees and display
      await displayWorkbookData(workbook, file.name);

      // Push to GitHub for cross-device sync
      showStatus('Syncing to cloud...', true);
      try {
        await pushToGitHub(arrayBuffer);
        showStatus('Uploaded & synced to all devices!', true);
      } catch (ghErr) {
        console.warn('GitHub sync failed:', ghErr);
        showStatus('Saved locally. Cloud sync failed: ' + ghErr.message, false);
      }

    } catch (err) {
      console.error('Upload failed:', err);
      showStatus('Failed to process file: ' + err.message, false);
    }
  }

  /**
   * Read a File object as an ArrayBuffer.
   * @param {File} file
   * @returns {Promise<ArrayBuffer>}
   */
  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        resolve(e.target.result);
      };
      reader.onerror = function () {
        reject(reader.error);
      };
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Extract employees from a workbook, save them, populate the table,
   * and render the full report.
   * @param {object} workbook — SheetJS workbook
   * @param {string} fileName — display name
   */
  async function displayWorkbookData(workbook, fileName) {
    try {
      // Extract employees
      var empResult = extractEmployees(workbook);
      var employees = empResult.employees;

      // Build employee list for IndexedDB
      var empList = employees.map(function (emp) {
        return {
          id: emp.id,
          name: emp.name,
          rowData: emp.rowData,
          headers: empResult.headers
        };
      });

      // Save employees to IndexedDB
      await saveEmployees(empList);

      // Show success status
      showStatus(
        'Successfully extracted ' + employees.length +
        ' employees from sheet: ' + empResult.sheetName,
        true
      );

      // Populate employee table
      populateEmployeeTable(employees);

      // Show employees section
      employeesSection.classList.remove('hidden');
      empCount.textContent = '(' + employees.length + ')';

      // Show report section and render
      reportSection.classList.remove('hidden');
      renderReport(workbook, sheetTabs, reportContainer);

    } catch (err) {
      console.error('Error processing workbook:', err);
      showStatus('Error extracting data: ' + err.message, false);
    }
  }

  /**
   * Populate the employee table body with extracted employees.
   * @param {Array} employees
   */
  function populateEmployeeTable(employees) {
    employeeTbody.innerHTML = '';
    for (var i = 0; i < employees.length; i++) {
      var tr = document.createElement('tr');

      var tdNum = document.createElement('td');
      tdNum.textContent = i + 1;
      tr.appendChild(tdNum);

      var tdId = document.createElement('td');
      tdId.textContent = employees[i].id;
      tr.appendChild(tdId);

      var tdName = document.createElement('td');
      tdName.textContent = employees[i].name;
      tr.appendChild(tdName);

      employeeTbody.appendChild(tr);
    }
  }

  /**
   * Show a status message below the upload zone.
   * @param {string} message
   * @param {boolean} isSuccess
   */
  function showStatus(message, isSuccess) {
    uploadStatus.textContent = message;
    uploadStatus.className = 'upload-status ' + (isSuccess ? 'success' : 'info');
    uploadStatus.hidden = false;
  }

  // Clear / re-upload
  clearBtn.addEventListener('click', async function () {
    // Delete old data from IndexedDB
    try {
      await clearAllData();
      _fetchingRemote = null;
    } catch (e) {
      console.error('Clear failed:', e);
    }

    // Reset file input
    fileInput.value = '';

    // Reset UI
    fileInfo.hidden = true;
    uploadZone.style.display = '';
    uploadStatus.hidden = true;

    // Hide data sections
    employeesSection.classList.add('hidden');
    reportSection.classList.add('hidden');
    employeeTbody.innerHTML = '';
    reportContainer.innerHTML = '';
    sheetTabs.innerHTML = '';
    sheetTabs.hidden = true;
    empCount.textContent = '';
  });

  // On page load — init DB and restore existing data
  async function init() {
    try {
      await initDB();

      var dataExists = await hasData();
      if (!dataExists) return;

      var wb = await getWorkbook();
      if (!wb || !wb.data) return;

      // Parse the stored workbook
      var workbook = XLSX.read(new Uint8Array(wb.data), {
        type: 'array',
        cellStyles: true,
        cellFormula: false,
        cellNF: true
      });

      // Show stored file name
      fileNameSpan.textContent = wb.fileName || 'Previously uploaded file';
      fileInfo.hidden = false;
      uploadZone.style.display = 'none';

      // Extract employees and display
      displayWorkbookData(workbook, wb.fileName);

    } catch (err) {
      console.error('Initialization error:', err);
    }
  }

  init();
})();
