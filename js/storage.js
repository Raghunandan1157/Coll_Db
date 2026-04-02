/**
 * storage.js — IndexedDB wrapper for caching employee session data.
 * All functions are global (no ES modules).
 *
 * DB: "CollectionReportDB" v1
 * Stores:
 *   - "employees"  keyPath "id"  — holds employee records fetched from the server
 */

// Cached database instance so we only open once
var _dbInstance = null;

/**
 * Opens (or creates) the IndexedDB database.
 * On first run, creates the "employees" object store.
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

      // Legacy workbook store — keep for upgrade compatibility
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
 * Saves an array of employee records into the "employees" store.
 * Clears any existing records first, then bulk-inserts all new ones
 * within a single transaction.
 * @param {Array<{id: string, name: string}>} employeeList
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
 * @returns {Promise<Object|null>}
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
 * @returns {Promise<Array>}
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
 * Checks whether any employee records exist in the store.
 * @returns {Promise<boolean>}
 */
function hasData() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("employees", "readonly");
      var store = tx.objectStore("employees");
      var request = store.count();

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
 * Clears all data from the employees store.
 * @returns {Promise<void>}
 */
function clearAllData() {
  return initDB().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction("employees", "readwrite");
      tx.objectStore("employees").clear();
      tx.oncomplete = function () { resolve(); };
      tx.onerror = function (event) { reject(event.target.error); };
    });
  });
}

/**
 * Fetches all employees from the server REST API and caches them in IndexedDB.
 * @returns {Promise<Array>}
 */
function getAllEmployeesFromServer() {
  return fetch('/api/employees', { cache: 'no-store' })
    .then(function (res) {
      if (!res.ok) throw new Error('Failed to fetch employees (' + res.status + ')');
      return res.json();
    })
    .then(function (employees) {
      if (Array.isArray(employees) && employees.length) {
        var mapped = employees.map(function(e){ return {id: e.emp_id || e.id, name: e.name, branch: e.branch, district: e.district, region: e.region}; });
      return saveEmployees(mapped).then(function () {
          return employees;
        });
      }
      return employees || [];
    })
    .catch(function (err) {
      console.error('Server employee fetch failed, falling back to cache:', err);
      return getAllEmployees();
    });
}

/**
 * Stub — portfolio.js and disbursement.js call this.
 * Returns null since workbooks are no longer used (data comes from REST APIs).
 * @param {string} category
 * @returns {Promise<null>}
 */
function getWorkbookByCategoryWithFallback(category) {
  return Promise.resolve(null);
}

/**
 * No-op stub — admin.js calls this to invalidate the fetch cache.
 * No longer needed since data comes from REST APIs.
 */
function resetFetchCache() {
  // no-op
}
