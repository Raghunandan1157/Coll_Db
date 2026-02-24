/**
 * storage.js — IndexedDB wrapper for persisting workbook and employee data
 * across browser sessions. All functions are global (no ES modules).
 *
 * DB: "CollectionReportDB" v1
 * Stores:
 *   - "workbook"   keyPath "id"  — holds the uploaded Excel file as ArrayBuffer
 *   - "employees"  keyPath "id"  — holds parsed employee records
 */

// Cached database instance so we only open once
var _dbInstance = null;

/**
 * Opens (or creates) the IndexedDB database.
 * On first run, creates the "workbook" and "employees" object stores.
 * Subsequent calls return the cached db instance.
 * @returns {Promise<IDBDatabase>}
 */
function initDB() {
  if (_dbInstance) {
    return Promise.resolve(_dbInstance);
  }

  return new Promise(function (resolve, reject) {
    var request = indexedDB.open("CollectionReportDB", 1);

    request.onupgradeneeded = function (event) {
      var db = event.target.result;

      // Create workbook store if it doesn't already exist
      if (!db.objectStoreNames.contains("workbook")) {
        db.createObjectStore("workbook", { keyPath: "id" });
      }

      // Create employees store if it doesn't already exist
      if (!db.objectStoreNames.contains("employees")) {
        db.createObjectStore("employees", { keyPath: "id" });
      }
    };

    request.onsuccess = function (event) {
      _dbInstance = event.target.result;

      // If the connection is unexpectedly closed, clear the cache
      _dbInstance.onclose = function () {
        _dbInstance = null;
      };

      resolve(_dbInstance);
    };

    request.onerror = function (event) {
      console.error("IndexedDB open failed:", event.target.error);
      reject(event.target.error);
    };
  });
}

/**
 * Saves (upserts) the workbook ArrayBuffer and metadata into the "workbook" store.
 * @param {ArrayBuffer} arrayBuffer — the raw file bytes
 * @param {string} fileName — original file name
 * @returns {Promise<void>}
 */
function saveWorkbook(arrayBuffer, fileName) {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("workbook", "readwrite");
      var store = tx.objectStore("workbook");

      var record = {
        id: "current",
        data: arrayBuffer,
        fileName: fileName,
        uploadedAt: new Date()
      };

      var request = store.put(record);

      request.onsuccess = function () {
        resolve();
      };

      request.onerror = function (event) {
        console.error("saveWorkbook failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Retrieves the current workbook record from the "workbook" store.
 * @returns {Promise<{id: string, data: ArrayBuffer, fileName: string, uploadedAt: Date}|null>}
 */
function getWorkbook() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("workbook", "readonly");
      var store = tx.objectStore("workbook");
      var request = store.get("current");

      request.onsuccess = function () {
        resolve(request.result || null);
      };

      request.onerror = function (event) {
        console.error("getWorkbook failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Saves an array of employee records into the "employees" store.
 * Clears any existing records first, then bulk-inserts all new ones
 * within a single transaction.
 * @param {Array<{id: string, name: string, rowData: object, headers: string[]}>} employeeList
 * @returns {Promise<void>}
 */
function saveEmployees(employeeList) {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("employees", "readwrite");
      var store = tx.objectStore("employees");

      // Clear all existing employee records first
      var clearRequest = store.clear();

      clearRequest.onsuccess = function () {
        // Bulk-insert all employees within the same transaction
        for (var i = 0; i < employeeList.length; i++) {
          store.put(employeeList[i]);
        }
      };

      clearRequest.onerror = function (event) {
        console.error("saveEmployees clear failed:", event.target.error);
        reject(event.target.error);
      };

      // Resolve or reject when the entire transaction completes
      tx.oncomplete = function () {
        resolve();
      };

      tx.onerror = function (event) {
        console.error("saveEmployees transaction failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Retrieves a single employee by id from the "employees" store.
 * Tries an exact key lookup first. If not found, falls back to a
 * case-insensitive search across all records.
 * @param {string} employeeId
 * @returns {Promise<{id: string, name: string, rowData: object, headers: string[]}|null>}
 */
function getEmployee(employeeId) {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("employees", "readonly");
      var store = tx.objectStore("employees");

      // Try exact match first
      var request = store.get(employeeId);

      request.onsuccess = function () {
        if (request.result) {
          resolve(request.result);
          return;
        }

        // Exact match not found — try case-insensitive search
        var allRequest = store.getAll();

        allRequest.onsuccess = function () {
          var records = allRequest.result || [];
          var needle = employeeId.toLowerCase();
          var match = null;

          for (var i = 0; i < records.length; i++) {
            if (String(records[i].id).toLowerCase() === needle) {
              match = records[i];
              break;
            }
          }

          resolve(match);
        };

        allRequest.onerror = function (event) {
          console.error("getEmployee fallback search failed:", event.target.error);
          reject(event.target.error);
        };
      };

      request.onerror = function (event) {
        console.error("getEmployee failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Retrieves all employee records from the "employees" store.
 * @returns {Promise<Array<{id: string, name: string, rowData: object, headers: string[]}>>}
 */
function getAllEmployees() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("employees", "readonly");
      var store = tx.objectStore("employees");
      var request = store.getAll();

      request.onsuccess = function () {
        resolve(request.result || []);
      };

      request.onerror = function (event) {
        console.error("getAllEmployees failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Checks whether a workbook record exists in the store.
 * @returns {Promise<boolean>}
 */
function hasData() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("workbook", "readonly");
      var store = tx.objectStore("workbook");
      var request = store.count("current");

      request.onsuccess = function () {
        resolve(request.result > 0);
      };

      request.onerror = function (event) {
        console.error("hasData failed:", event.target.error);
        reject(event.target.error);
      };
    });
  });
}

/**
 * Clears all data from both workbook and employees stores.
 * @returns {Promise<void>}
 */
function clearAllData() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(["workbook", "employees"], "readwrite");
      tx.objectStore("workbook").clear();
      tx.objectStore("employees").clear();
      tx.oncomplete = function () { resolve(); };
      tx.onerror = function (event) { reject(event.target.error); };
    });
  });
}

/* ---------- Remote fallback ---------- */
var _fetchingRemote = null;

/**
 * Extract employees by trying ALL sheets in the workbook,
 * not just the last one. Prefers sheets with "Overall" in the name.
 */
function _extractEmployeesAllSheets(workbook) {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) return [];

  // Try "Overall" sheet first
  for (var s = 0; s < workbook.SheetNames.length; s++) {
    if (workbook.SheetNames[s].toUpperCase().includes('OVERALL')) {
      var parsed = extractEmployees(
        { SheetNames: [workbook.SheetNames[s]], Sheets: workbook.Sheets }
      );
      if (parsed.employees && parsed.employees.length) return parsed.employees;
    }
  }

  // Try each sheet until we find employees
  for (var s = 0; s < workbook.SheetNames.length; s++) {
    var parsed = extractEmployees(
      { SheetNames: [workbook.SheetNames[s]], Sheets: workbook.Sheets }
    );
    if (parsed.employees && parsed.employees.length) return parsed.employees;
  }

  return [];
}

/**
 * Fetches data/report.xlsx from the server (cache-busted), parses it,
 * saves workbook + employees to IndexedDB, and returns the workbook record.
 * @returns {Promise<{id: string, data: ArrayBuffer, fileName: string, uploadedAt: Date}|null>}
 */
function _fetchRemoteReport() {
  if (_fetchingRemote) return _fetchingRemote;

  // Fetch directly from GitHub API — no-store forces fresh every time
  var apiUrl = 'https://api.github.com/repos/Raghunandan1157/Coll_Db/contents/data/report.xlsx?t=' + Date.now();

  _fetchingRemote = fetch(apiUrl, {
      headers: { 'Accept': 'application/vnd.github.v3.raw' },
      cache: 'no-store'
    })
    .then(function (res) {
      if (!res.ok) throw new Error('No remote report (' + res.status + ')');
      return res.arrayBuffer();
    })
    .then(function (buf) {
      return clearAllData().then(function () {
        return saveWorkbook(buf, 'report.xlsx');
      }).then(function () {
        if (typeof XLSX !== 'undefined' && typeof extractEmployees === 'function') {
          var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellFormula: false, cellNF: true });
          var emps = _extractEmployeesAllSheets(wb);
          if (emps.length) {
            return saveEmployees(emps).then(function () {
              return { id: 'current', data: buf, fileName: 'report.xlsx', uploadedAt: new Date() };
            });
          }
        }
        return { id: 'current', data: buf, fileName: 'report.xlsx', uploadedAt: new Date() };
      });
    })
    .catch(function (err) {
      console.error('Remote report fetch failed:', err);
      _fetchingRemote = null;
      return null;
    });

  return _fetchingRemote;
}

/**
 * Gets the workbook — always fetches latest from remote first,
 * falls back to IndexedDB if remote fails.
 */
function getWorkbookWithFallback() {
  _fetchingRemote = null; // reset so we always re-fetch
  return _fetchRemoteReport().then(function (wb) {
    if (wb) return wb;
    return getWorkbook(); // fallback to local cache
  });
}

/**
 * Gets all employees — always fetches latest from remote first,
 * falls back to IndexedDB if remote fails.
 */
function getAllEmployeesWithFallback() {
  _fetchingRemote = null; // reset so we always re-fetch
  return _fetchRemoteReport().then(function () {
    return getAllEmployees();
  }).then(function (emps) {
    if (emps && emps.length) return emps;
    // Remote had no employees, try local cache
    return getAllEmployees();
  });
}
