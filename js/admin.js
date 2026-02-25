/**
 * admin.js — Admin panel with 3 side-by-side upload cards
 * (Portfolio, Disbursement, Collection).
 * Collection uses full flow (employees + GitHub sync).
 * Portfolio & Disbursement store the file for future use.
 */
(function () {
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

  async function pushToGitHub(arrayBuffer) {
    var token = getGHToken();
    if (!token) token = promptForToken();
    if (!token) throw new Error('No GitHub token — file saved locally only.');

    var shaRes = await fetch(GH_API, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    var sha = null;
    if (shaRes.ok) {
      var shaData = await shaRes.json();
      sha = shaData.sha;
    }

    var bytes = new Uint8Array(arrayBuffer);
    var binary = '';
    for (var i = 0; i < bytes.length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    var base64 = btoa(binary);

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

  /* ---------- DOM refs ---------- */
  var employeesSection = document.getElementById('employees-section');
  var empCount = document.getElementById('emp-count');
  var employeeTbody = document.getElementById('employee-tbody');
  var reportSection = document.getElementById('report-section');
  var sheetTabs = document.getElementById('sheet-tabs');
  var reportContainer = document.getElementById('report-container');

  document.getElementById('logout-btn').onclick = logout;

  /* ---------- Wire up each upload card ---------- */
  var cards = document.querySelectorAll('.upload-card');
  for (var i = 0; i < cards.length; i++) {
    (function (card) {
      var category = card.dataset.category;
      var uploadZone = card.querySelector('.upload-zone');
      var fileInput = uploadZone.querySelector('input[type="file"]');
      var fileInfo = card.querySelector('.card-file-info');
      var fileNameSpan = card.querySelector('.card-file-name');
      var clearBtn = card.querySelector('.card-clear-btn');

      uploadZone.addEventListener('click', function () { fileInput.click(); });

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
        if (e.dataTransfer.files.length > 0) {
          handleCategoryFile(category, e.dataTransfer.files[0], card);
        }
      });

      fileInput.addEventListener('change', function () {
        if (fileInput.files.length > 0) {
          handleCategoryFile(category, fileInput.files[0], card);
        }
      });

      clearBtn.addEventListener('click', function () {
        if (!confirm('Remove this uploaded file?')) return;
        clearCategoryData(category, card);
      });
    })(cards[i]);
  }

  /* ---------- Handle file upload ---------- */
  async function handleCategoryFile(category, file, card) {
    var name = file.name.toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      showCardStatus(card, 'Upload .xlsx or .xls only.', false);
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      showCardStatus(card, 'File too large. Maximum 10 MB.', false);
      return;
    }

    var fileInfo = card.querySelector('.card-file-info');
    var fileNameSpan = card.querySelector('.card-file-name');
    var uploadZone = card.querySelector('.upload-zone');

    fileNameSpan.textContent = file.name;
    fileInfo.hidden = false;
    uploadZone.style.display = 'none';

    try {
      var arrayBuffer = await readFileAsArrayBuffer(file);

      // Validate that the file is a real Excel file
      try {
        XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
      } catch (xlsxErr) {
        showCardStatus(card, 'Invalid file. Upload a valid Excel file.', false);
        fileInfo.hidden = true;
        uploadZone.style.display = '';
        return;
      }

      if (category === 'collection') {
        await clearAllData();
        resetFetchCache();
        await saveWorkbook(arrayBuffer, file.name);

        var workbook = XLSX.read(new Uint8Array(arrayBuffer), {
          type: 'array', cellStyles: true, cellFormula: false, cellNF: true
        });
        await displayWorkbookData(workbook, file.name);

        showCardStatus(card, 'Syncing to cloud...', true);
        try {
          await pushToGitHub(arrayBuffer);
          showCardStatus(card, 'Uploaded & synced!', true);
        } catch (ghErr) {
          console.warn('GitHub sync failed:', ghErr);
          showCardStatus(card, 'Saved locally. Sync failed.', false);
        }
      } else {
        await saveWorkbookByCategory(arrayBuffer, file.name, category);
        showCardStatus(card, 'Uploaded!', true);
      }
    } catch (err) {
      console.error('Upload failed:', err);
      showCardStatus(card, 'Failed: ' + err.message, false);
    }
  }

  function showCardStatus(card, message, isSuccess) {
    var el = card.querySelector('.card-upload-status');
    el.textContent = message;
    el.className = 'card-upload-status ' + (isSuccess ? 'success' : 'info');
    el.hidden = false;
  }

  async function clearCategoryData(category, card) {
    try {
      if (category === 'collection') {
        await clearAllData();
        resetFetchCache();
        employeesSection.classList.add('hidden');
        reportSection.classList.add('hidden');
        employeeTbody.innerHTML = '';
        reportContainer.innerHTML = '';
        sheetTabs.innerHTML = '';
        sheetTabs.hidden = true;
        empCount.textContent = '';
      } else {
        await clearWorkbookByCategory(category);
      }
    } catch (e) {
      console.error('Clear failed:', e);
    }

    var fileInput = card.querySelector('input[type="file"]');
    fileInput.value = '';
    card.querySelector('.card-file-info').hidden = true;
    card.querySelector('.upload-zone').style.display = '';
    card.querySelector('.card-upload-status').hidden = true;
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) { resolve(e.target.result); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  /* ---------- Display workbook data (collection) ---------- */
  async function displayWorkbookData(workbook, fileName) {
    try {
      var employees = _extractEmployeesAllSheets(workbook);
      var empResult = extractEmployees(workbook);

      var empList = employees.map(function (emp) {
        return { id: emp.id, name: emp.name, rowData: emp.rowData, headers: empResult.headers };
      });

      await saveEmployees(empList);
      populateEmployeeTable(employees);
      employeesSection.classList.remove('hidden');
      empCount.textContent = '(' + employees.length + ')';
      reportSection.classList.remove('hidden');
      renderReport(workbook, sheetTabs, reportContainer);
    } catch (err) {
      console.error('Error processing workbook:', err);
    }
  }

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

  /* ---------- Init — restore saved data ---------- */
  async function init() {
    try {
      await initDB();

      // Restore collection
      var dataExists = await hasData();
      if (dataExists) {
        var wb = await getWorkbook();
        if (wb && wb.data) {
          var workbook = XLSX.read(new Uint8Array(wb.data), {
            type: 'array', cellStyles: true, cellFormula: false, cellNF: true
          });
          var collCard = document.querySelector('.upload-card[data-category="collection"]');
          if (collCard) {
            collCard.querySelector('.card-file-name').textContent = wb.fileName || 'Previously uploaded';
            collCard.querySelector('.card-file-info').hidden = false;
            collCard.querySelector('.upload-zone').style.display = 'none';
          }
          displayWorkbookData(workbook, wb.fileName);
        }
      }

      // Restore portfolio
      var pWb = await getWorkbookByCategory('portfolio');
      if (pWb) {
        var pCard = document.querySelector('.upload-card[data-category="portfolio"]');
        if (pCard) {
          pCard.querySelector('.card-file-name').textContent = pWb.fileName || 'Previously uploaded';
          pCard.querySelector('.card-file-info').hidden = false;
          pCard.querySelector('.upload-zone').style.display = 'none';
        }
      }

      // Restore disbursement
      var dWb = await getWorkbookByCategory('disbursement');
      if (dWb) {
        var dCard = document.querySelector('.upload-card[data-category="disbursement"]');
        if (dCard) {
          dCard.querySelector('.card-file-name').textContent = dWb.fileName || 'Previously uploaded';
          dCard.querySelector('.card-file-info').hidden = false;
          dCard.querySelector('.upload-zone').style.display = 'none';
        }
      }
    } catch (err) {
      console.error('Initialization error:', err);
    }
  }

  init();
})();
