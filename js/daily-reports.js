// Service worker disabled for EC2 deployment

// --- AUTO-LOGIN FROM EMPLOYEE DASHBOARD ---
// Read role from the existing employee dashboard login (localStorage)
(function autoLoginFromDashboard() {
    var roleAuth = localStorage.getItem('roleAuth');
    var roleName = localStorage.getItem('roleName');
    var roleLocation = localStorage.getItem('roleLocation');
    var employeeId = localStorage.getItem('employeeId');
    var employeeName = localStorage.getItem('employeeName');
    
    if (!roleAuth && !employeeId) return; // Not logged in — let normal login show
    
    // Map dashboard roles to Daily Plan roles: RM → DM → AM → BM
    var mappedRole = 'CEO';
    var mappedUser = 'CEO';

    if (roleName === 'CEO' || roleName === 'RM' || roleName === 'SM') {
        mappedRole = 'CEO';
        mappedUser = 'CEO';
    } else if (roleName === 'DM' || roleName === 'DvM') {
        mappedRole = 'DM';
        mappedUser = roleLocation || employeeName || 'DM';
    } else if (roleName === 'AM') {
        mappedRole = 'AM';
        mappedUser = roleLocation || employeeName || 'AM';
    } else if (roleName === 'BM') {
        mappedRole = 'BM';
        mappedUser = roleLocation || employeeName || 'BM';
    } else if (roleName === 'FO' || employeeId) {
        mappedRole = 'BM'; // FO sees same as BM (own branch)
        mappedUser = roleLocation || employeeName || employeeId;
    }
    
    // Store for the Daily Plan to pick up
    window._dashboardAutoLogin = { role: mappedRole, user: mappedUser, roleName: roleName, location: roleLocation };
})();


// --- EC2 API WRAPPER (replaces Neon direct connection) ---
async function neonQuery(sqlText, params = []) {
    // Route SQL queries to appropriate EC2 API endpoints
    const sql = sqlText.trim().toLowerCase();
    
    try {
        // SELECT queries
        if (sql.startsWith('select')) {
            // Extract table and conditions
            const isAchievements = sql.includes('daily_reports_achievements');
            const isReports = sql.includes('daily_reports') && !isAchievements;
            const isCount = sql.includes('count(');
            
            // EXISTS check: SELECT COUNT(*)... WHERE date = $1 AND branch_name = $2
            if (isCount && params.length >= 2) {
                const table = isAchievements ? 'achievements' : 'reports';
                const resp = await fetch('/api/daily-plan/exists?table=' + table + '&date=' + encodeURIComponent(params[0]) + '&branch=' + encodeURIComponent(params[1]));
                const data = await resp.json();
                return { data: [{ cnt: data.exists ? 1 : 0 }], error: null, count: 1 };
            }
            
            // Data query: extract date range from params
            var url = '/api/daily-plan/' + (isAchievements ? 'achievements' : 'reports') + '?';
            // Try to extract from/to dates from the SQL WHERE clause
            if (params.length >= 2) {
                url += 'from=' + encodeURIComponent(params[0]) + '&to=' + encodeURIComponent(params[1]);
            } else if (params.length === 1) {
                url += 'from=' + encodeURIComponent(params[0]) + '&to=' + encodeURIComponent(params[0]);
            }
            
            const resp = await fetch(url);
            if (!resp.ok) {
                const errBody = await resp.json().catch(() => ({}));
                throw new Error(errBody.error || 'Server error ' + resp.status);
            }
            const result = await resp.json();
            return { data: result.data || [], error: null, count: result.count || 0 };
        }
        
        // INSERT/UPSERT queries
        if (sql.startsWith('insert')) {
            const isAchievements = sql.includes('daily_reports_achievements');
            // Extract column names and values from the SQL
            const colMatch = sql.match(/\(([^)]+)\)\s*values/i);
            if (!colMatch) throw new Error('Could not parse INSERT columns');
            const cols = colMatch[1].split(',').map(c => c.trim().replace(/"/g, ''));
            
            const data = {};
            cols.forEach((col, i) => { data[col] = params[i] != null ? params[i] : 0; });
            
            const resp = await fetch('/api/daily-plan/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ table: isAchievements ? 'achievements' : 'reports', data })
            });
            if (!resp.ok) {
                const errBody = await resp.json().catch(() => ({}));
                throw new Error(errBody.error || 'Server error ' + resp.status);
            }
            const result = await resp.json();
            if (!result.success) throw new Error(result.error || 'Save failed');
            return { data: [], error: null, count: 1 };
        }
        
        // SELECT 1 (connection test)
        if (sql === 'select 1') {
            return { data: [{ '?column?': 1 }], error: null, count: 1 };
        }
        
        console.warn('Unhandled SQL query:', sqlText);
        return { data: [], error: null, count: 0 };
    } catch (err) {
        console.error('API query failed:', err);
        throw err;
    }
}

// --- DATA CACHE (show last data instantly while fresh data loads) ---
function getCachedData(key) {
    try {
        const raw = localStorage.getItem('nlpl_cache_' + key);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        // Cache is valid for 12 hours
        if (Date.now() - parsed.ts > 12 * 60 * 60 * 1000) return null;
        return parsed.data;
    } catch (e) { return null; }
}

function setCachedData(key, data) {
    try {
        localStorage.setItem('nlpl_cache_' + key, JSON.stringify({ ts: Date.now(), data: data }));
    } catch (e) { /* localStorage full — ignore */ }
}

// --- DATA & STATE ---
// V2 hierarchy: Branch → Area → Division → Region (state)
const defaultTSVData = `Branch\tArea\tDivision\tRegion
BUDWAL\tKADAPA\tANDRA PRADESH\tANDRA PRADESH
DHARMAVARAM\tKADAPA\tANDRA PRADESH\tANDRA PRADESH
KADAPA\tKADAPA\tANDRA PRADESH\tANDRA PRADESH
KADIRI\tKADAPA\tANDRA PRADESH\tANDRA PRADESH
CHALLAKERE\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
CHANNAGIRI\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
CHITRADURGA\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
HIRIYUR\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
HOLAKERE\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
HOSADURGA\tCHITRADURGA\tCHITRADURGA\tCHITRADURGA
DAVANAGERE\tDAVANAGERE\tCHITRADURGA\tCHITRADURGA
HARIHARA\tDAVANAGERE\tCHITRADURGA\tCHITRADURGA
HONNALI\tDAVANAGERE\tCHITRADURGA\tCHITRADURGA
SANTHEBENNURU\tDAVANAGERE\tCHITRADURGA\tCHITRADURGA
AJJAMPURA\tKADUR\tCHITRADURGA\tCHITRADURGA
KADUR\tKADUR\tCHITRADURGA\tCHITRADURGA
PANCHANHALLI\tKADUR\tCHITRADURGA\tCHITRADURGA
TARIKERE\tKADUR\tCHITRADURGA\tCHITRADURGA
BALLARI\tBALLARI\tHOSPET\tCHITRADURGA
KUDATHINI\tBALLARI\tHOSPET\tCHITRADURGA
SANDURU\tBALLARI\tHOSPET\tCHITRADURGA
SIRUGUPPA\tBALLARI\tHOSPET\tCHITRADURGA
HAGARIBOMMANAHALLI\tHOSPET\tHOSPET\tCHITRADURGA
HOSPET\tHOSPET\tHOSPET\tCHITRADURGA
HUVENAHADAGALLI\tHOSPET\tHOSPET\tCHITRADURGA
KUDLIGI\tHOSPET\tHOSPET\tCHITRADURGA
HARAPANAHALLI\tKOTTURU\tHOSPET\tCHITRADURGA
JAGALORE\tKOTTURU\tHOSPET\tCHITRADURGA
KHANAHOSAHALLI\tKOTTURU\tHOSPET\tCHITRADURGA
KOTTURU\tKOTTURU\tHOSPET\tCHITRADURGA
BAGALKOT\tBAGALKOT\tBELAGAVI\tDHARWAD
BILAGI\tBAGALKOT\tBELAGAVI\tDHARWAD
JAMAKHANDI\tBAGALKOT\tBELAGAVI\tDHARWAD
LOKAPUR\tBAGALKOT\tBELAGAVI\tDHARWAD
BAILHONGAL\tBELAGAVI\tBELAGAVI\tDHARWAD
BELAGAVI\tBELAGAVI\tBELAGAVI\tDHARWAD
GOKAK\tBELAGAVI\tBELAGAVI\tDHARWAD
KITTUR\tBELAGAVI\tBELAGAVI\tDHARWAD
YARAGATTI\tBELAGAVI\tBELAGAVI\tDHARWAD
ATHANI\tCHIKKODI\tBELAGAVI\tDHARWAD
CHIKKODI\tCHIKKODI\tBELAGAVI\tDHARWAD
MUDALAGI\tCHIKKODI\tBELAGAVI\tDHARWAD
NIPPANI\tCHIKKODI\tBELAGAVI\tDHARWAD
BADAMI\tBADAMI\tHUBLI\tDHARWAD
GAJENDRAGAD\tBADAMI\tHUBLI\tDHARWAD
NARAGUNDA\tBADAMI\tHUBLI\tDHARWAD
RAMDURGA\tBADAMI\tHUBLI\tDHARWAD
DHARWAD\tDHARWAD\tHUBLI\tDHARWAD
HUBLI\tDHARWAD\tHUBLI\tDHARWAD
HUBLI-2\tDHARWAD\tHUBLI\tDHARWAD
KALGHATGI\tDHARWAD\tHUBLI\tDHARWAD
GADAG\tGADAG\tHUBLI\tDHARWAD
LAXMESHWAR\tGADAG\tHUBLI\tDHARWAD
MUNDARAGI\tGADAG\tHUBLI\tDHARWAD
GANGAVATHI\tKUSHTAGI\tHUBLI\tDHARWAD
HUNGUND\tKUSHTAGI\tHUBLI\tDHARWAD
KOPPAL\tKUSHTAGI\tHUBLI\tDHARWAD
KUSHTAGI\tKUSHTAGI\tHUBLI\tDHARWAD
AURAD\tBIDAR\tBIDAR\tKALBURGI
BHALKI\tBIDAR\tBIDAR\tKALBURGI
BIDAR\tBIDAR\tBIDAR\tKALBURGI
BIDAR-2\tBIDAR\tBIDAR\tKALBURGI
BASAVAKALYAN\tHUMNABAD\tBIDAR\tKALBURGI
HULSOOR\tHUMNABAD\tBIDAR\tKALBURGI
HUMNABAD\tHUMNABAD\tBIDAR\tKALBURGI
KAMALAPURA\tHUMNABAD\tBIDAR\tKALBURGI
LINGSUGUR\tLINGSUGUR\tBIDAR\tKALBURGI
MANVI\tLINGSUGUR\tBIDAR\tKALBURGI
RAICHUR\tLINGSUGUR\tBIDAR\tKALBURGI
SINDHNUR\tLINGSUGUR\tBIDAR\tKALBURGI
SIRWAR\tLINGSUGUR\tBIDAR\tKALBURGI
CHINCHOLI\tSEDAM\tBIDAR\tKALBURGI
KALAGI\tSEDAM\tBIDAR\tKALBURGI
SEDAM\tSEDAM\tBIDAR\tKALBURGI
ALMEL\tINDI\tKALBURGI\tKALBURGI
CHADCHAN\tINDI\tKALBURGI\tKALBURGI
INDI\tINDI\tKALBURGI\tKALBURGI
AFZALPUR\tKALBURGI\tKALBURGI\tKALBURGI
ALAND\tKALBURGI\tKALBURGI\tKALBURGI
JEVARGI\tKALBURGI\tKALBURGI\tKALBURGI
KALABURAGI\tKALBURGI\tKALBURGI\tKALBURGI
KALBURGI-2\tKALBURGI\tKALBURGI\tKALBURGI
DEVADURGA\tSHAHAPUR\tKALBURGI\tKALBURGI
SHAHAPUR\tSHAHAPUR\tKALBURGI\tKALBURGI
YADGIR\tSHAHAPUR\tKALBURGI\tKALBURGI
MUDDEBIHAL\tVIJAYAPUR\tKALBURGI\tKALBURGI
SINDAGI\tVIJAYAPUR\tKALBURGI\tKALBURGI
TALIKOTI\tVIJAYAPUR\tKALBURGI\tKALBURGI
TIKOTA\tVIJAYAPUR\tKALBURGI\tKALBURGI
VIJAYAPUR\tVIJAYAPUR\tKALBURGI\tKALBURGI
GADWAL\tMAHABOOBNAGAR\tTELANGANA\tTELANGANA
MAHABUB NAGAR\tMAHABOOBNAGAR\tTELANGANA\tTELANGANA
MARIKAL\tMAHABOOBNAGAR\tTELANGANA\tTELANGANA
TANDUR\tMAHABOOBNAGAR\tTELANGANA\tTELANGANA
KODANGAL\tSANGAREDDY\tTELANGANA\tTELANGANA
NARAYANKHED\tSANGAREDDY\tTELANGANA\tTELANGANA
SANGAREDDY\tSANGAREDDY\tTELANGANA\tTELANGANA
ZAHEERABAD\tSANGAREDDY\tTELANGANA\tTELANGANA
CHANDAPURA\tBANGALORE URBAN\tDODDABALLAPURA\tTUMKUR
HEBBAL\tBANGALORE URBAN\tDODDABALLAPURA\tTUMKUR
J P NAGAR\tBANGALORE URBAN\tDODDABALLAPURA\tTUMKUR
KENGERI\tBANGALORE URBAN\tDODDABALLAPURA\tTUMKUR
BAGEPALLI\tCHIKBALLAPURA\tDODDABALLAPURA\tTUMKUR
CHIKBALLAPURA\tCHIKBALLAPURA\tDODDABALLAPURA\tTUMKUR
CHINTAMANI\tCHIKBALLAPURA\tDODDABALLAPURA\tTUMKUR
SRINIVASPURA\tCHIKBALLAPURA\tDODDABALLAPURA\tTUMKUR
DEVANAHALLI\tDODDABALLAPURA\tDODDABALLAPURA\tTUMKUR
DODDABALLAPURA\tDODDABALLAPURA\tDODDABALLAPURA\tTUMKUR
GOWRIBIDANUR\tDODDABALLAPURA\tDODDABALLAPURA\tTUMKUR
BANGARPET\tKOLAR\tDODDABALLAPURA\tTUMKUR
BETHAMANGALA\tKOLAR\tDODDABALLAPURA\tTUMKUR
KOLAR\tKOLAR\tDODDABALLAPURA\tTUMKUR
MALUR\tKOLAR\tDODDABALLAPURA\tTUMKUR
CHIKKAMAGALURU\tCHIKKAMAGALURU\tTUMKUR\tTUMKUR
MUDIGERE\tCHIKKAMAGALURU\tTUMKUR\tTUMKUR
NR PURA\tCHIKKAMAGALURU\tTUMKUR\tTUMKUR
CHIKKANAYAKANAHALLI\tTIPTUR\tTUMKUR\tTUMKUR
GUBBI\tTIPTUR\tTUMKUR\tTUMKUR
HULIYAR\tTIPTUR\tTUMKUR\tTUMKUR
TIPTUR\tTIPTUR\tTUMKUR\tTUMKUR
TUREVEKERE\tTIPTUR\tTUMKUR\tTUMKUR
DABUSPET\tTUMKUR\tTUMKUR\tTUMKUR
KORATAGERE\tTUMKUR\tTUMKUR\tTUMKUR
KUNIGAL\tTUMKUR\tTUMKUR\tTUMKUR
MADHUGIRI\tTUMKUR\tTUMKUR\tTUMKUR
SIRA\tTUMKUR\tTUMKUR\tTUMKUR
TUMKUR\tTUMKUR\tTUMKUR\tTUMKUR`;

let state = {
    currentUser: null,
    role: null,
    rawData: null,
    targetData: {},
    branchDetails: {}, // Main Store
    activeTab: 'dashboard',
    viewStack: [],
    fullTree: {},
    systemDate: new Date().toISOString().split('T')[0], // Default to Today
    isLoading: false, // Track loading state
    realtimeChannels: [], // Track subscriptions for cleanup
    viewMode: 'PLAN', // 'PLAN' (Plan only) or 'REVIEW' (Plan vs Achievement)
    reportLevel: 'REGION', // 'BRANCH', 'AREA', 'DIVISION', 'REGION'
    dateFrom: null, // Range Start
    dateTo: null,   // Range End
    isDragging: false,
    dragStartDate: null,
    _fetchController: null, // AbortController for in-flight fetchData
    _pollingInterval: null // Polling interval for data refresh (replaces Neon realtime)
};

// --- HELPER FUNCTIONS ---

// Lazy-load external scripts on demand (only downloads once, caches for reuse)
const _loadedScripts = {};
function loadScript(url) {
    if (_loadedScripts[url]) return _loadedScripts[url];
    _loadedScripts[url] = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = url;
        script.onload = resolve;
        script.onerror = () => reject(new Error('Failed to load: ' + url));
        document.head.appendChild(script);
    });
    return _loadedScripts[url];
}

// Indian number format: 12,34,567
function formatIndianNumber(num) {
    if (num === '' || num === null || num === undefined) return '';
    const str = String(num).replace(/,/g, '');
    if (isNaN(str) || str === '') return '';
    const [intPart, decPart] = str.split('.');
    const lastThree = intPart.slice(-3);
    const otherNumbers = intPart.slice(0, -3);
    const formatted = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + (otherNumbers ? ',' : '') + lastThree;
    return decPart !== undefined ? formatted + '.' + decPart : formatted;
}

// Strip commas for database save
function unformatNumber(str) {
    if (str === '' || str === null || str === undefined) return '';
    return String(str).replace(/,/g, '');
}

function getPreviousMonthName(dateStr) {
    const targetDate = dateStr ? new Date(dateStr) : new Date();
    // Fix: Set to 1st of month to avoid month-end rollover issues (e.g. Mar 31 -> Mar 3)
    targetDate.setDate(1);
    // Go to previous month
    targetDate.setMonth(targetDate.getMonth() - 1);

    const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
    return monthNames[targetDate.getMonth()];
}

// getSlippedLabel removed — now using '1-30 DPD' literal

// Get dynamic Fiscal Year label (e.g., "FY 2025-26")
// Fiscal year in India runs from April to March
function getFiscalYearLabel(dateStr) {
    const targetDate = dateStr ? new Date(dateStr) : new Date();
    const year = targetDate.getFullYear();
    const month = targetDate.getMonth(); // 0-indexed (0 = Jan, 3 = Apr)

    // If month is Jan-Mar (0-2), we're in the FY that started in previous year
    // If month is Apr-Dec (3-11), we're in the FY that started this year
    const fyStartYear = month < 3 ? year - 1 : year;
    const fyEndYear = fyStartYear + 1;

    // Format: "FY 2025-26"
    return `FY ${fyStartYear}-${String(fyEndYear).slice(-2)}`;
}

function calculateTotalCollectionPercentage() {
    let totalPlan = 0;
    let totalAchieve = 0;

    // Use getReportRows logic to respect current filtering
    // (We reuse the existing logic which respects state.role and state.reportLevel)
    // However, getReportRows returns aggregated rows. We can sum them up.

    // To be efficient and precise, let's iterate branchDetails directly
    // but filter by the current VIEW constraints (Role/Level)
    // Actually, getReportRows('BRANCH') already does the filtering for DM/Region
    // Let's use getAggregatedReportData('BRANCH', true/false) logic or just iterate rawData

    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');
    const idxDistrict = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'district');
    const idxDM = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'dm name');

    // DM Restriction
    const isDM = state.role === 'DM';
    const dmBranches = isDM ? getDMBranches() : null;

    state.rawData.rows.forEach(r => {
        const name = r[idxBranch];
        if (!name) return;

        // DM Filter
        if (isDM && !dmBranches.includes(name)) return;

        // Also respect Regional drill down?
        // The requirements say "when filters/date/level changes".
        // The Report Builder has a "Report Level" filter (Branch/District/Region).
        // But the data source is always the same (filtered by Role).
        // The "Achievement %" is typically "Total Achievement / Total Plan" for the *scope of the report*.
        // Since the report usually covers "All Regions" (for CEO) or "My District" (for DM),
        // we should calculate based on ALL available data for the current user.
        // Unless "level changes" implies showing the achievement for that specific level?
        // But the button generates the report for *all* entities at that level.
        // So global aggregation for the user is correct.

        const entry = state.branchDetails[name];
        if (!entry) return;

        // Sum Plans
        if (entry.target) {
            totalPlan += (Number(entry.target.ftod_plan) || 0) +
                (Number(entry.target.dpd_1_30_plan) || 0) +
                (Number(entry.target.dpd_31_60_plan) || 0) +
                (Number(entry.target.dpd_61_90_plan) || 0);
        }

        // Sum Achievements
        if (entry.achievement) {
            totalAchieve += (Number(entry.achievement.ftod_actual) || 0) +
                (Number(entry.achievement.dpd_1_30_actual) || 0) +
                (Number(entry.achievement.dpd_31_60_actual) || 0) +
                (Number(entry.achievement.dpd_61_90_actual) || 0);
        }
    });

    if (totalPlan === 0) return 0;
    return Math.round((totalAchieve / totalPlan) * 100);
}

// --- PERSISTENCE ---
function loadPersistedDateRange() {
    try {
        const saved = localStorage.getItem('nlpl_date_range');
        if (saved) {
            const parsed = JSON.parse(saved);
            state.dateFrom = parsed.from;
            state.dateTo = parsed.to;
            state.systemDate = parsed.to; // Keep systemDate as end of range for legacy compatibility
            return true;
        }
    } catch (e) {
        console.error("Error loading date range:", e);
    }
    return false;
}

function saveDateRange(from, to) {
    state.dateFrom = from;
    state.dateTo = to;
    if (from && to) {
        localStorage.setItem('nlpl_date_range', JSON.stringify({ from, to }));
    }
}

// --- PERFORMANCE UTILITIES ---
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function throttle(func, limit) {
    let inThrottle;
    return function (...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

// Debounced fetch for rapid date changes (date picker, arrow nav, quick buttons)
const debouncedFetchData = debounce(() => fetchData(), 300);

// --- LOADING STATE MANAGEMENT ---
function setLoading(isLoading, message = 'Loading...') {
    state.isLoading = isLoading;
    const loader = document.getElementById('global-loader');
    if (loader) {
        if (isLoading) {
            loader.querySelector('.loader-text').textContent = message;
            loader.classList.add('visible');
        } else {
            loader.classList.remove('visible');
        }
    }
}

// --- ERROR BOUNDARY ---
function safeExecute(fn, fallback = null, context = 'Operation') {
    try {
        const result = fn();
        if (result instanceof Promise) {
            return result.catch(err => {
                console.error(`${context} Error:`, err);
                showToast(`${context} failed. Please try again.`, 'alert');
                return fallback;
            });
        }
        return result;
    } catch (err) {
        console.error(`${context} Error:`, err);
        showToast(`${context} failed. Please try again.`, 'alert');
        return fallback;
    }
}

// --- INIT ---
window.addEventListener("DOMContentLoaded", async () => {
    // Load Persisted Range or Default to Today
    if (!loadPersistedDateRange()) {
        const today = new Date();
        const isoToday = formatDateISO(today);
        state.dateFrom = isoToday;
        state.dateTo = isoToday;
        state.systemDate = isoToday;
    }
    const { headers, rows } = parseTSV(defaultTSVData);
    state.rawData = { headers, rows };
    state.fullTree = buildHierarchy(headers, rows);

    // Populate DM Dropdown
    populateDMDropdown(rows, headers);

    // Load Local Plans (Keep local for now as per requirement focus on Branch Details)
    const saved = localStorage.getItem('dosa_targets');
    if (saved) state.targetData = JSON.parse(saved);

    // --- AUTO-LOGIN from employee dashboard (skip login overlay) ---
    if (window._dashboardAutoLogin) {
        const info = window._dashboardAutoLogin;
        console.log('Auto-login from dashboard:', info);
        state.role = info.role;
        state.currentUser = info.user;
        
        // Hide overlays
        const loginOverlay = document.getElementById('login-overlay');
        const dailyOverlay = document.getElementById('daily-overlay');
        if (loginOverlay) loginOverlay.classList.add('hidden');
        if (dailyOverlay) dailyOverlay.classList.add('hidden');
        
        // Show app
        const sidebar = document.getElementById('sidebar');
        const mainEl = document.querySelector('.main-content');
        if (sidebar) sidebar.classList.remove('hidden');
        if (mainEl) mainEl.classList.remove('hidden');
        
        // Check if branch-only mode (iframe with #branch=X)
        var _hash = window.location.hash;
        var _branchOnly = (window.self !== window.top) && _hash && _hash.indexOf('branch=') >= 0;
        var _branchParam = _branchOnly ? decodeURIComponent(_hash.split('branch=')[1].split('&')[0]) : null;
        
        if (_branchOnly && _branchParam) {
            // Branch-only mode: skip full data load, just open the branch modal
            updateHeaderDate(state.systemDate);
            try { await fetchData(); } catch(e) { console.warn('fetchData skipped:', e); }
            loadAppUI();
            setTimeout(function() {
                if (typeof openBranchModal === 'function') openBranchModal(_branchParam);
            }, 500);
            return;
        }
        
        // Report mode: #report=DATE&level=LEVEL&type=plan|both
        var _reportMode = _hash && _hash.indexOf('report=') >= 0;
        if (_reportMode) {
            var _reportDate = decodeURIComponent(_hash.split('report=')[1].split('&')[0]);
            var _reportLevel = 'REGION';
            if (_hash.indexOf('level=') >= 0) _reportLevel = _hash.split('level=')[1].split('&')[0];
            var _reportType = 'plan';
            if (_hash.indexOf('type=') >= 0) _reportType = _hash.split('type=')[1].split('&')[0];
            
            // Set state for report
            state.systemDate = _reportDate;
            state.dateFrom = _reportDate;
            state.dateTo = _reportDate;
            state.reportLevel = _reportLevel;
            
            updateHeaderDate(_reportDate);
            try { await fetchData(); } catch(e) { console.warn('Report fetchData:', e); }
            
            // Switch to reports tab and generate
            loadAppUI();
            setTimeout(function() {
                // Switch to reports tab
                if (typeof switchTab === 'function') switchTab('reports');
                // Trigger report generation
                setTimeout(function() {
                    if (_reportType === 'both' && typeof handleGenerateBothReports === 'function') {
                        handleGenerateBothReports();
                    } else if (typeof handleGeneratePlanReport === 'function') {
                        handleGeneratePlanReport();
                    }
                }, 500);
            }, 500);
            return;
        }
        
        // Normal mode
        updateHeaderDate(state.systemDate);
        try {
            await fetchData();
        } catch(e) {
            console.error('Auto-login fetchData error:', e);
        }
        loadAppUI();
        // Auto-open branch modal if hash specifies it
        if (_hash && _hash.indexOf('branch=') >= 0) {
            setTimeout(function() {
                var bp = decodeURIComponent(_hash.split('branch=')[1].split('&')[0]);
                if (bp && typeof openBranchModal === 'function') openBranchModal(bp);
            }, 1000);
        }
        return; // Skip showing the login overlay
    }

    // Init Header Date (Will be updated on selection)
    updateHeaderDate(state.systemDate);

    // Note: We DO NOT fetch data automatically here anymore.
    // We wait for Date Selection in the Overlay.

    // Hide Sidebar/Content initially
    document.getElementById("sidebar").classList.add("hidden");
    document.querySelector(".main-content").classList.add("hidden");

    // Init Toggle State
    const toggle = document.getElementById("viewModeToggle");
    if (toggle) toggle.checked = state.viewMode === 'REVIEW';
    updateViewModeUI();
});

function toggleViewMode() {
    const toggle = document.getElementById("viewModeToggle");
    state.viewMode = toggle.checked ? 'REVIEW' : 'PLAN';
    updateViewModeUI();

    // Re-render
    renderDashboard();
}

function setReportLevel(level) {
    // Restrict report levels based on role
    if (state.role === 'DM' && (level === 'REGION' || level === 'DIVISION')) {
        console.warn('DM users cannot access REGION/DIVISION level reports');
        return;
    }
    if (state.role === 'AM' && (level === 'REGION' || level === 'DIVISION' || level === 'AREA')) {
        console.warn('AM users cannot access higher-level reports');
        return;
    }
    if (state.role === 'BM') {
        console.warn('BM users cannot access reports');
        return;
    }

    state.reportLevel = level;
    // Re-render Reports Tab (since buttons are there)
    if (state.activeTab === 'reports') {
        const buffer = document.createElement("div");
        renderReports(buffer);
        const container = document.getElementById("dashboard-container");
        container.innerHTML = "";
        container.appendChild(buffer);
    }
}

// --- DM HELPER FUNCTIONS ---
// Get the district(s) assigned to the current DM
function getDMDistricts() {
    // V2: returns areas for the current user's division
    if (!state.currentUser || !state.rawData) return [];
    const idxDiv = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'division');
    const idxArea = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'area');
    if (idxDiv === -1 || idxArea === -1) return [];
    const areas = new Set();
    state.rawData.rows.forEach(r => {
        if ((r[idxDiv] || '').trim() === state.currentUser) {
            areas.add(r[idxArea]);
        }
    });
    return Array.from(areas);
}

// Get branches for the current user based on role (cached for performance)
let _dmBranchesCache = null;
let _dmBranchesCacheUser = null;
function getDMBranches() {
    if (!state.currentUser || !state.rawData) return [];
    if (_dmBranchesCache && _dmBranchesCacheUser === state.currentUser + state.role) return _dmBranchesCache;

    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxArea = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'area');
    const idxDiv = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'division');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    if (idxBranch === -1) return [];

    const branches = [];
    state.rawData.rows.forEach(r => {
        const role = state.role;
        if (role === 'CEO' || role === 'RM') {
            branches.push(r[idxBranch]);
        } else if (role === 'DM') {
            // Division manager: match division column
            if (idxDiv !== -1 && (r[idxDiv] || '').trim() === state.currentUser) branches.push(r[idxBranch]);
        } else if (role === 'AM') {
            // Area manager: match area column
            if (idxArea !== -1 && (r[idxArea] || '').trim() === state.currentUser) branches.push(r[idxBranch]);
        } else if (role === 'BM') {
            // BM: match branch name directly
            if ((r[idxBranch] || '').trim() === state.currentUser) branches.push(r[idxBranch]);
        }
    });
    _dmBranchesCache = branches;
    _dmBranchesCacheUser = state.currentUser + state.role;
    return branches;
}

function updateViewModeUI() {
    const body = document.body;
    if (state.viewMode === 'PLAN') {
        body.classList.add('mode-plan');
        document.getElementById('labelPlan').style.color = 'var(--primary)';
        document.getElementById('labelReview').style.color = 'var(--text-secondary)';
    } else {
        body.classList.remove('mode-plan');
        document.getElementById('labelPlan').style.color = 'var(--text-secondary)';
        document.getElementById('labelReview').style.color = 'var(--primary)';
    }
}

function updateToggleVisibility() {
    const toggleContainer = document.getElementById('viewModeToggleContainer');
    if (toggleContainer) {
        // Show only if CEO AND on Dashboard tab
        if (state.role === 'CEO' && state.activeTab === 'dashboard') {
            toggleContainer.classList.remove('hidden');
        } else {
            toggleContainer.classList.add('hidden');
        }
    }
}

function updateHeaderDate(isoDate) {
    const dateObj = new Date(isoDate);
    const dateStr = dateObj.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    document.getElementById("headerDate").textContent = dateStr;

    // Also update the hidden date input for CEO calendar picker
    const dateInput = document.getElementById("ceoDateInput");
    if (dateInput) {
        dateInput.value = isoDate;
    }
}

// CEO Calendar Functions
let calendarDate = new Date(); // Current calendar view month

function toggleDateDropdown() {
    // Only allow CEO to use calendar
    if (state.role !== 'CEO') return;

    const dropdown = document.getElementById('dateRangeDropdown');
    const wasHidden = dropdown.classList.contains('hidden');
    dropdown.classList.toggle('hidden');

    // If opening, render calendar
    if (wasHidden) {
        // Sync calendar view with current selected system date
        if (state.systemDate) {
            calendarDate = new Date(state.systemDate);
        } else {
            calendarDate = new Date(); // Fallback
        }
        renderCalendar();
    }

    // Set default dates if not set
    const dateFrom = document.getElementById('dateFrom');
    const dateTo = document.getElementById('dateTo');
    if (!dateFrom.value) {
        dateFrom.value = state.systemDate;
        dateTo.value = state.systemDate;
    }
}

function navigateCalendar(direction) {
    calendarDate.setMonth(calendarDate.getMonth() + direction);
    renderCalendar();
}

function renderCalendar() {
    const container = document.getElementById('calendarDays');
    const monthYearLabel = document.getElementById('calendarMonthYear');

    const year = calendarDate.getFullYear();
    const month = calendarDate.getMonth();

    // Update month/year label
    monthYearLabel.textContent = calendarDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

    // Get first day of month and number of days
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();

    // Today's date for comparison
    const today = new Date();
    const todayStr = formatDateISO(today);

    let html = '';

    // Previous month's days
    for (let i = firstDay - 1; i >= 0; i--) {
        const day = daysInPrevMonth - i;
        const dateStr = formatDateISO(new Date(year, month - 1, day));
        html += `<div class="calendar-day other-month" onclick="selectCalendarDate('${dateStr}')">${day}</div>`;
    }

    // Current month's days
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = formatDateISO(new Date(year, month, day));
        let classes = 'calendar-day';

        if (dateStr === todayStr) {
            classes += ' today';
        }
        if (dateStr === state.systemDate) {
            classes += ' selected';
        }

        html += `<div class="${classes}" onclick="selectCalendarDate('${dateStr}')">${day}</div>`;
    }

    // Next month's days (fill remaining cells)
    const totalCells = 42; // 6 rows x 7 columns
    const currentCells = firstDay + daysInMonth;
    for (let i = 1; i <= totalCells - currentCells; i++) {
        const dateStr = formatDateISO(new Date(year, month + 1, i));
        html += `<div class="calendar-day other-month" onclick="selectCalendarDate('${dateStr}')">${i}</div>`;
    }

    container.innerHTML = html;
}

// --- CALENDAR DRAG LOGIC ---
function handleCalendarDayDown(dateStr) {
    state.isDragging = true;
    state.dragStartDate = dateStr;
    // Visually start selection
    renderCalendar(dateStr, dateStr);
}

function handleCalendarDayOver(dateStr) {
    if (state.isDragging && state.dragStartDate) {
        // Update visual selection only (don't commit state yet)
        renderCalendar(state.dragStartDate, dateStr);
    }
}

function handleCalendarDayUp(dateStr) {
    if (state.isDragging && state.dragStartDate) {
        state.isDragging = false;

        let start = state.dragStartDate;
        let end = dateStr;

        // Swap if backwards
        if (start > end) {
            [start, end] = [end, start];
        }

        // Apply
        applyDateRangeInternal(start, end, formatRangeLabel(start, end));
        saveDateRange(start, end);

        state.dragStartDate = null;
    }
}

// Global mouse up to cancel drag if outside
document.addEventListener('mouseup', () => {
    if (state.isDragging) {
        state.isDragging = false;
        // Re-render to show actual committed state
        renderCalendar();
    }
});

function renderCalendar(previewStart = null, previewEnd = null) {
    const container = document.getElementById('calendarDays');
    const monthYearLabel = document.getElementById('calendarMonthYear');

    const year = calendarDate.getFullYear();
    const month = calendarDate.getMonth();

    monthYearLabel.textContent = calendarDate.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();
    const todayStr = formatDateISO(new Date());

    // Determine what range to show (State or Drag Preview)
    let currentStart = state.dateFrom;
    let currentEnd = state.dateTo;

    if (previewStart && previewEnd) {
        currentStart = previewStart;
        currentEnd = previewEnd;
        // Validate swap for preview
        if (currentStart > currentEnd) [currentStart, currentEnd] = [currentEnd, currentStart];
    }

    let html = '';

    // Previous month (Disabled for simplicity in drag? Or just standard)
    for (let i = firstDay - 1; i >= 0; i--) {
        const day = daysInPrevMonth - i;
        const dateStr = formatDateISO(new Date(year, month - 1, day));
        // We allow dragging from prev month days too
        html += renderDayHTML(day, dateStr, todayStr, currentStart, currentEnd, 'other-month');
    }

    // Current month
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = formatDateISO(new Date(year, month, day));
        html += renderDayHTML(day, dateStr, todayStr, currentStart, currentEnd, '');
    }

    // Next month
    const totalCells = 42;
    const currentCells = firstDay + daysInMonth;
    for (let i = 1; i <= totalCells - currentCells; i++) {
        const dateStr = formatDateISO(new Date(year, month + 1, i));
        html += renderDayHTML(i, dateStr, todayStr, currentStart, currentEnd, 'other-month');
    }

    container.innerHTML = html;
}

function renderDayHTML(dayNum, dateStr, todayStr, start, end, extraClasses) {
    let classes = `calendar-day ${extraClasses}`;
    if (dateStr === todayStr) classes += ' today';

    // Range Logic
    if (start && end) {
        if (dateStr === start) classes += ' selected range-start';
        if (dateStr === end) classes += ' selected range-end';
        if (dateStr > start && dateStr < end) classes += ' in-range';

        // Single day range
        if (dateStr === start && start === end) classes += ' selected range-start range-end';
    } else if (start && dateStr === start) {
        // Just start selected
        classes += ' selected range-start range-end';
    }

    return `<div class="${classes}" 
        onmousedown="handleCalendarDayDown('${dateStr}')"
        onmouseenter="handleCalendarDayOver('${dateStr}')"
        onmouseup="handleCalendarDayUp('${dateStr}')"
    >${dayNum}</div>`;
}

function formatRangeLabel(from, to) {
    const fromDate = new Date(from);
    const toDate = new Date(to);

    if (from === to) {
        return fromDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    }
    return `${fromDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${toDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;
}

function selectCalendarDate(dateStr) {
    // Clear quick button selection
    document.querySelectorAll('.date-quick-btn').forEach(btn => btn.classList.remove('active'));

    // Format label
    const dateObj = new Date(dateStr);
    const label = dateObj.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

    // Apply single date
    applyDateRangeInternal(dateStr, dateStr, label);
}

// Close dropdown when clicking outside
document.addEventListener('click', function (e) {
    const wrapper = document.querySelector('.date-picker-wrapper');
    const dropdown = document.getElementById('dateRangeDropdown');
    if (wrapper && dropdown && !wrapper.contains(e.target)) {
        dropdown.classList.add('hidden');
    }
});

function selectDatePreset(preset) {
    const now = new Date();
    let fromDate, toDate, label;

    // Remove active state from all buttons
    document.querySelectorAll('.date-quick-btn').forEach(btn => btn.classList.remove('active'));

    switch (preset) {
        case 'today':
            fromDate = toDate = formatDateISO(now);
            label = new Date(fromDate).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
            document.getElementById('btn-today').classList.add('active');
            break;
        case 'month':
            fromDate = formatDateISO(new Date(now.getFullYear(), now.getMonth(), 1));
            toDate = formatDateISO(now);
            label = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            document.getElementById('btn-month').classList.add('active');
            break;
        case 'year':
            fromDate = formatDateISO(new Date(now.getFullYear(), 0, 1));
            toDate = formatDateISO(now);
            label = `Year ${now.getFullYear()}`;
            document.getElementById('btn-year').classList.add('active');
            break;
    }

    // Apply the date range
    applyDateRangeInternal(fromDate, toDate, label);
}

function formatDateISO(date) {
    const offset = date.getTimezoneOffset() * 60000;
    return new Date(date.getTime() - offset).toISOString().split('T')[0];
}

// Get yesterday's date in ISO format (YYYY-MM-DD)
function getYesterdayISO() {
    const today = new Date();
    today.setDate(today.getDate() - 1);
    return formatDateISO(today);
}

// Check if a given ISO date string is today
function isToday(isoDate) {
    return isoDate === formatDateISO(new Date());
}

function validateDateRange() {
    const from = document.getElementById('dateFrom').value;
    const to = document.getElementById('dateTo').value;
    const btn = document.getElementById('dateApplyBtn');

    if (from && to && from <= to) {
        btn.disabled = false;
    } else {
        btn.disabled = true;
    }
}

function applyDateRange() {
    const from = document.getElementById('dateFrom').value;
    const to = document.getElementById('dateTo').value;

    if (!from || !to || from > to) {
        showToast('⚠️ Please select a valid date range', true);
        return;
    }

    // Format label
    const fromDate = new Date(from);
    const toDate = new Date(to);
    let label;

    if (from === to) {
        label = fromDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    } else {
        label = `${fromDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${toDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}`;
    }

    applyDateRangeInternal(from, to, label);
}

async function applyDateRangeInternal(fromDate, toDate, label) {
    // Close dropdown
    document.getElementById('dateRangeDropdown').classList.add('hidden');

    // Store date range in state
    state.dateFrom = fromDate;
    state.dateTo = toDate;
    state.systemDate = toDate; // For compatibility

    // Update header
    document.getElementById('headerDate').textContent = label;

    // Only show full loading if no cache available
    const cacheKey = (state.role || 'CEO') + '_range_' + fromDate + '_' + toDate;
    if (!getCachedData(cacheKey)) {
        setLoading(true, `Loading data for ${label}...`);
    }

    try {
        // Fetch data for date range (handles cache + network internally)
        await fetchDataRange(fromDate, toDate);

        // Final render with fresh network data
        renderDashboard();

        // Show toast
        if (fromDate === toDate) {
            showToast(`📅 Viewing data for ${label}`);
        } else {
            showToast(`📅 Viewing aggregated data: ${label}`);
        }
    } catch (error) {
        console.error('Error fetching date range:', error);
        showToast('⚠️ Error loading date range data', true);
    }
}

// Set Target for Tomorrow (DM Action)
function setTargetForTomorrow() {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);

    const isoDate = tomorrow.toISOString().split('T')[0];
    const label = tomorrow.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

    // helper to pad
    const pad = n => n < 10 ? '0' + n : n;
    // ensure formatted iso
    const y = tomorrow.getFullYear();
    const m = pad(tomorrow.getMonth() + 1);
    const d = pad(tomorrow.getDate());
    const iso = `${y}-${m}-${d}`;

    applyDateRangeInternal(iso, iso, label);

    setTimeout(() => {
        showToast(`🚀 Target Date set to Tomorrow (${label})`);
    }, 600);
}

// Navigate Date Picker by Day (Arrow Buttons)
function navigateDatePicker(delta) {
    // Get current date from state
    const currentDate = state.systemDate ? new Date(state.systemDate) : new Date();

    // Add delta days
    currentDate.setDate(currentDate.getDate() + delta);

    // Format ISO date
    const pad = n => n < 10 ? '0' + n : n;
    const y = currentDate.getFullYear();
    const m = pad(currentDate.getMonth() + 1);
    const d = pad(currentDate.getDate());
    const isoDate = `${y}-${m}-${d}`;

    // Format label
    const label = currentDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

    // Apply the new date
    applyDateRangeInternal(isoDate, isoDate, label);
}

// Fetch data for a date range (aggregates multiple days)
// Core fetch and aggregate logic - Reusable
async function fetchAndAggregateData(fromDate, toDate, retryCount = 0) {
    const MAX_RETRIES = 3;
    const selectCols = 'branch_name,date,region,district,dm_name,ftod_actual,ftod_plan,dpd_1_30_actual,dpd_1_30_plan,dpd_31_60_actual,dpd_31_60_plan,dpd_61_90_actual,dpd_61_90_plan,npa_activation,npa_closure,fy_non_start_acc,fy_non_start_plan,disb_igl_acc,disb_igl_amt,disb_fig_acc,disb_fig_amt,disb_il_acc,disb_il_amt,kyc_igl,kyc_fig,kyc_il';
    const isSingleDate = fromDate === toDate;

    // Build SQL query with parameterized WHERE clause
    let whereClauses = [];
    let params = [];
    let paramIdx = 1;

    if (isSingleDate) {
        whereClauses.push(`date = $${paramIdx}`);
        params.push(fromDate);
        paramIdx++;
    } else {
        whereClauses.push(`date >= $${paramIdx}`);
        params.push(fromDate);
        paramIdx++;
        whereClauses.push(`date <= $${paramIdx}`);
        params.push(toDate);
        paramIdx++;
    }

    // Note: Branch filtering is done client-side via getDMBranches() after fetch.
    // The neonQuery wrapper can't handle array params, so we fetch all and filter.

    const whereSQL = whereClauses.length > 0 ? ' WHERE ' + whereClauses.join(' AND ') : '';
    const sqlPlan = `SELECT ${selectCols} FROM daily_reports${whereSQL}`;
    const sqlAchieve = `SELECT ${selectCols} FROM daily_reports_achievements${whereSQL}`;

    // Add timeout protection (mobile networks need longer)
    let timeoutId;
    const isMobile = /Android|iPhone|iPad|iPod|Opera Mini|IEMobile/i.test(navigator.userAgent);
    const timeoutMs = isMobile ? 60000 : 45000;
    const timeoutPromise = new Promise((_, reject) =>
        timeoutId = setTimeout(() => reject(new Error('Request timeout')), timeoutMs)
    );

    try {
        const [resPlan, resAchieve] = await Promise.race([
            Promise.all([neonQuery(sqlPlan, params), neonQuery(sqlAchieve, params)]),
            timeoutPromise
        ]);
        clearTimeout(timeoutId);

        if (resPlan.error || resAchieve.error) {
            const errorMsg = resPlan.error?.message || resAchieve.error?.message;
            const isNetworkError = errorMsg && (errorMsg.includes('Failed to fetch') || errorMsg.includes('Load failed') || errorMsg.includes('NetworkError') || errorMsg.includes('timeout'));
            if (isNetworkError && retryCount < MAX_RETRIES) {
                const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
                console.warn(`Range fetch error, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
                showToast(`Connection slow, retrying... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
                await new Promise(resolve => setTimeout(resolve, delay));
                return fetchAndAggregateData(fromDate, toDate, retryCount + 1);
            }
            throw new Error(errorMsg);
        }

        // Aggregate data by branch
        const aggregatedPlans = aggregateDataByBranch(resPlan.data || []);
        const aggregatedAchievements = aggregateDataByBranch(resAchieve.data || []);

        // Create Branch Details Map
        const branchDetails = {};

        aggregatedPlans.forEach(row => {
            const br = row.branch_name;
            if (!branchDetails[br]) branchDetails[br] = {};
            branchDetails[br].target = row;
        });

        aggregatedAchievements.forEach(row => {
            const br = row.branch_name;
            if (!branchDetails[br]) branchDetails[br] = {};
            branchDetails[br].achievement = row;
        });

        // Backfill missing achievement values with plan data when needed
        mergeAchievementsWithPlan(branchDetails);

        return branchDetails;
    } catch (err) {
        clearTimeout(timeoutId);
        const isNetworkError = err.message && (
            err.message.includes('Failed to fetch') ||
            err.message.includes('Load failed') ||
            err.message.includes('timeout') ||
            err.message.includes('NetworkError')
        );
        if (isNetworkError && retryCount < MAX_RETRIES) {
            const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
            console.warn(`Range fetch timeout, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
            showToast(`Connection slow, retrying... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
            await new Promise(resolve => setTimeout(resolve, delay));
            return fetchAndAggregateData(fromDate, toDate, retryCount + 1);
        }
        throw err;
    }
}

// Fetch data for a date range (aggregates multiple days)
async function fetchDataRange(fromDate, toDate) {
    // DEDUPLICATION: Abort any in-flight fetch when a new range fetch starts
    if (state._fetchController) {
        console.log("fetchDataRange: Aborting previous in-flight request");
        state._fetchController.abort();
    }
    state._fetchController = new AbortController();
    const abortSignal = state._fetchController.signal;

    const cacheKey = (state.role || 'CEO') + '_range_' + fromDate + '_' + toDate;

    // Show cached data instantly (stale-while-revalidate)
    const cached = getCachedData(cacheKey);
    if (cached) {
        state.branchDetails = cached;
        renderDashboard();
        setLoading(true, 'Refreshing...');
    }

    try {
        const branchDetails = await fetchAndAggregateData(fromDate, toDate);
        // Check abort before applying results
        if (abortSignal.aborted) return;
        state.branchDetails = branchDetails;
        setCachedData(cacheKey, branchDetails);
        setLoading(false);
    } catch (error) {
        if (abortSignal.aborted) return; // Silently ignore aborted requests
        console.error('Fetch range error:', error);
        setLoading(false);
        // If we showed cached data, don't re-throw — user already sees data
        if (cached) {
            showToast('Using cached data — refresh failed', 'warning');
            return;
        }
        throw error;
    }
}

// Helper to aggregate data by branch (sum numeric fields)
function aggregateDataByBranch(rows) {
    const byBranch = {};
    const metaKeys = new Set(['id', 'branch_name', 'date', 'region', 'district', 'dm_name', 'created_at']);

    rows.forEach(row => {
        const branchName = row.branch_name;
        if (!branchName) return;

        if (!byBranch[branchName]) {
            byBranch[branchName] = { ...row };
        } else {
            Object.keys(row).forEach(key => {
                if (metaKeys.has(key)) return;

                const incomingVal = row[key];
                const numericVal = incomingVal === null || incomingVal === undefined || incomingVal === '' ? null : Number(incomingVal);

                // Sum only when the value is numeric; keep nulls as-is so we can detect missing data later
                if (numericVal !== null && !isNaN(numericVal)) {
                    const existing = byBranch[branchName][key];
                    const existingNum = existing === null || existing === undefined || existing === '' ? 0 : Number(existing) || 0;
                    byBranch[branchName][key] = existingNum + numericVal;
                } else if (!(key in byBranch[branchName])) {
                    byBranch[branchName][key] = incomingVal ?? null;
                }
            });
            // Keep non-numeric fields from latest date
            if (row.date > byBranch[branchName].date) {
                byBranch[branchName].date = row.date;
                byBranch[branchName].dm_name = row.dm_name;
                byBranch[branchName].region = row.region;
                byBranch[branchName].district = row.district;
            }
        }
    });

    // Normalize numeric-looking values to numbers, leave null/undefined intact
    Object.values(byBranch).forEach(row => {
        Object.keys(row).forEach(key => {
            if (metaKeys.has(key)) return;
            const val = row[key];
            if (val === null || val === undefined || val === '') {
                row[key] = null;
                return;
            }
            if (!isNaN(Number(val))) {
                row[key] = Number(val);
            }
        });
    });

    return Object.values(byBranch);
}

// Merge achievement data with plan data when Neon achievement rows miss values
// IMPORTANT: Only merges when there is ACTUAL achievement data in Neon, not empty objects
// NOTE: Only merge fields that have DIFFERENT names between plan and achievement tables.
// Fields with SAME names (npa_activation, npa_closure, disb_*, kyc_*) should NOT be merged
// because they are independent values in both tables.
function mergeAchievementsWithPlan(branchDetails) {
    // Only include fields that have corresponding plan counterparts with DIFFERENT names
    // e.g., ftod_actual (achievement) -> ftod_plan (plan)
    // Exclude: npa_activation, npa_closure, disb_*, kyc_* as they are independent values
    const achievementToMerge = [
        'ftod_actual',
        'dpd_1_30_actual',
        'dpd_31_60_actual',
        'dpd_61_90_actual',
        'fy_non_start_acc'
    ];

    // Mapping from achievement field to corresponding plan field
    const achievementToPlanMap = {
        'ftod_actual': 'ftod_plan',
        'dpd_1_30_actual': 'dpd_1_30_plan',
        'dpd_31_60_actual': 'dpd_31_60_plan',
        'dpd_61_90_actual': 'dpd_61_90_plan',
        'fy_non_start_acc': 'fy_non_start_plan'
    };

    // Original full list for reference (used for hasRealAchievementData check)
    const achievementFields = [
        'ftod_actual',
        'dpd_1_30_actual',
        'dpd_31_60_actual',
        'dpd_61_90_actual',
        'npa_activation',
        'npa_closure',
        'fy_non_start_acc',
        'disb_igl_acc',
        'disb_igl_amt',
        'disb_fig_acc',
        'disb_fig_amt',
        'disb_il_acc',
        'disb_il_amt',
        'kyc_igl',
        'kyc_fig',
        'kyc_il'
    ];

    Object.keys(branchDetails).forEach(branchName => {
        const entry = branchDetails[branchName] || {};
        const plan = entry.target || {};
        const achievement = entry.achievement; // Don't default to empty object here!

        // Only process if there is ACTUAL achievement data from Neon
        // Check if entry.achievement exists and has at least one meaningful field 
        // (e.g., id, branch_name from Neon, or any non-null achievement field)
        const hasRealAchievementData = achievement && (
            achievement.id ||
            achievement.branch_name ||
            achievementFields.some(field => {
                const val = achievement[field];
                return val !== null && val !== undefined && val !== '' && val !== 0;
            })
        );

        if (!hasRealAchievementData) {
            // No actual achievement data - don't create a fake achievement object
            return;
        }

        const merged = { ...achievement };

        // Only merge fields that have different names between plan and achievement tables
        // Use the mapping to get the correct plan field name
        achievementToMerge.forEach(achField => {
            const achVal = achievement ? achievement[achField] : null;
            const hasAchVal = achVal !== null && achVal !== undefined && achVal !== '';
            if (!hasAchVal) {
                // Get the corresponding plan field name from the mapping
                const planField = achievementToPlanMap[achField];
                const planVal = plan ? plan[planField] : null;
                if (planVal !== null && planVal !== undefined && planVal !== '') {
                    merged[achField] = planVal;
                }
            } else if (typeof achVal === 'string' && achVal.trim() !== '' && !isNaN(Number(achVal))) {
                merged[achField] = Number(achVal);
            }
        });

        // Carry forward metadata so report rows remain populated
        merged.branch_name = achievement.branch_name || plan.branch_name || branchName;
        merged.region = achievement.region || plan.region || '';
        merged.district = achievement.district || plan.district || '';
        merged.dm_name = achievement.dm_name || achievement.dm || plan.dm_name || plan.dm || '';
        merged.date = achievement.date || plan.date;

        entry.achievement = merged;
        branchDetails[branchName] = entry;
    });
}

// --- SUPABASE ACTIONS ---
async function fetchData(retryCount = 0, skipRender = false) {
    const MAX_RETRIES = 3;
    const targetDate = state.systemDate;
    const cacheKey = (state.role || 'CEO') + '_' + targetDate;

    // DEDUPLICATION: Abort any in-flight fetch when a new one starts (retryCount === 0 means fresh call)
    if (retryCount === 0) {
        if (state._fetchController) {
            console.log("fetchData: Aborting previous in-flight request");
            state._fetchController.abort();
        }
        state._fetchController = new AbortController();
    }
    const abortSignal = state._fetchController?.signal;

    console.log("fetchData: Starting for date " + targetDate + (retryCount > 0 ? ` (retry ${retryCount}/${MAX_RETRIES})` : ''));

    // Show cached data instantly while we fetch fresh data from network
    const hasCachedData = retryCount === 0 && !!getCachedData(cacheKey);
    if (retryCount === 0) {
        if (hasCachedData) {
            console.log("fetchData: Showing cached data for " + cacheKey);
            state.branchDetails = getCachedData(cacheKey);
            if (!skipRender) renderDashboard();
            setLoading(false); // Don't block UI — user already sees data
        } else {
            setLoading(true, 'Loading data...');
        }
    } else {
        setLoading(true, `Retrying... (${retryCount}/${MAX_RETRIES})`);
    }

    try {
        // Check if this request was already aborted (superseded by a newer call)
        if (abortSignal?.aborted) {
            console.log("fetchData: Aborted before fetch (superseded)");
            return;
        }

        // Parallel Fetch with timeout for reliability
        // Mobile networks (especially Android on cellular) need longer timeouts
        let timeoutId;
        const isMobile = /Android|iPhone|iPad|iPod|Opera Mini|IEMobile/i.test(navigator.userAgent);
        const timeoutMs = isMobile ? 45000 : 30000; // 45s mobile, 30s desktop
        const timeout = new Promise((_, reject) =>
            timeoutId = setTimeout(() => {
                console.error("fetchData: Timeout triggered after " + timeoutMs + "ms");
                reject(new Error('Request timeout'));
            }, timeoutMs)
        );

        // Also reject on abort signal (newer fetch supersedes this one)
        const abortPromise = new Promise((_, reject) => {
            if (abortSignal) {
                abortSignal.addEventListener('abort', () => {
                    reject(new Error('Fetch aborted: superseded by newer request'));
                }, { once: true });
            }
        });

        // DM optimization: only fetch branches assigned to this DM
        const selectCols = 'branch_name,date,region,district,dm_name,ftod_actual,ftod_plan,dpd_1_30_actual,dpd_1_30_plan,dpd_31_60_actual,dpd_31_60_plan,dpd_61_90_actual,dpd_61_90_plan,npa_activation,npa_closure,fy_non_start_acc,fy_non_start_plan,disb_igl_acc,disb_igl_amt,disb_fig_acc,disb_fig_amt,disb_il_acc,disb_il_amt,kyc_igl,kyc_fig,kyc_il';
        let whereClauses = ['date = $1'];
        let queryParams = [targetDate];
        // Branch filtering is done client-side after fetch (neonQuery can't handle array params)

        const whereSQL = ' WHERE ' + whereClauses.join(' AND ');
        const p1 = neonQuery(`SELECT ${selectCols} FROM daily_reports${whereSQL}`, queryParams);
        const p2 = neonQuery(`SELECT ${selectCols} FROM daily_reports_achievements${whereSQL}`, queryParams);

        console.log("fetchData: Awaiting Promise.race");
        const [resPlan, resAchieve] = await Promise.race([
            Promise.all([p1, p2]),
            timeout,
            abortPromise
        ]);
        clearTimeout(timeoutId); // Clear timeout on success

        // Check abort after fetch completes (another call may have started during await)
        if (abortSignal?.aborted) {
            console.log("fetchData: Aborted after fetch (superseded)");
            return;
        }

        console.log("fetchData: Promise.race resolved", { resPlan, resAchieve });

        if (resPlan.error || resAchieve.error) {
            const errorMsg = resPlan.error?.message || resAchieve.error?.message || 'Unknown error';
            const isNetworkError = errorMsg.includes('Failed to fetch') || errorMsg.includes('Load failed') || errorMsg.includes('NetworkError') || errorMsg.includes('timeout');

            // If user already sees cached data, don't retry — just warn
            if (hasCachedData) {
                console.warn('fetchData: Network error but cache shown, skipping retries');
                showToast('Showing cached data. Could not refresh from server.', 'warning');
                return;
            }

            if (isNetworkError && retryCount < MAX_RETRIES) {
                const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
                console.warn(`Fetch network error, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
                clearTimeout(timeoutId);
                await new Promise(resolve => setTimeout(resolve, delay));
                // Check abort before retrying
                if (abortSignal?.aborted) return;
                return fetchData(retryCount + 1, skipRender);
            }

            console.error("Neon Load Error:", errorMsg);
            showToast('Failed to load data: ' + errorMsg, 'alert');
            return;
        }

        // MERGE Logic
        state.branchDetails = {};

        // Map Plans
        (resPlan.data || []).forEach(row => {
            if (!state.branchDetails[row.branch_name]) state.branchDetails[row.branch_name] = {};
            state.branchDetails[row.branch_name].target = row;
        });

        // Map Achievements
        (resAchieve.data || []).forEach(row => {
            if (!state.branchDetails[row.branch_name]) state.branchDetails[row.branch_name] = {};
            state.branchDetails[row.branch_name].achievement = row;
        });

        // Ensure achievement data is filled even if Neon achievement rows are sparse
        mergeAchievementsWithPlan(state.branchDetails);

        // Cache fresh data for instant load on next visit
        setCachedData(cacheKey, state.branchDetails);

        const targetCount = resPlan.data?.length || 0;
        const achieveCount = resAchieve.data?.length || 0;
        console.log(`fetchData: Processed ${targetCount} plans, ${achieveCount} achievements`);

        if (!skipRender) {
            renderDashboard();
            console.log("fetchData: renderDashboard completed");
        }
    } catch (err) {
        // Silently ignore aborted requests — they were intentionally cancelled
        if (err.message?.includes('aborted') || err.message?.includes('superseded') || abortSignal?.aborted) {
            console.log("fetchData: Request was aborted (superseded by newer request)");
            return;
        }

        console.error("Fetch Error:", err);

        // Retry on network/timeout errors — but skip retries if cache already shown
        const isNetworkError = err.message && (
            err.message.includes('Failed to fetch') ||
            err.message.includes('Load failed') ||
            err.message.includes('timeout') ||
            err.message.includes('NetworkError')
        );

        if (hasCachedData) {
            console.warn('fetchData: Timeout but cache shown, skipping retries');
            showToast('Showing cached data. Could not refresh from server.', 'warning');
            return;
        }

        if (isNetworkError && retryCount < MAX_RETRIES) {
            const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
            console.warn(`Fetch timeout, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
            showToast(`Connection slow, retrying... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
            await new Promise(resolve => setTimeout(resolve, delay));
            // Check abort before retrying
            if (abortSignal?.aborted) return;
            return fetchData(retryCount + 1, skipRender);
        }

        showToast('Connection error. Please check your network and try again.', 'alert');
    } finally {
        console.log("fetchData: Entering finally, calling setLoading(false)");
        setLoading(false);
    }
}

// Check if yesterday's achievement exists for a branch
async function checkYesterdayAchievementExists(branchName) {
    const yesterdayDate = getYesterdayISO();
    try {
        const result = await neonQuery(
            'SELECT COUNT(*)::int AS cnt FROM daily_reports_achievements WHERE date = $1 AND branch_name = $2',
            [yesterdayDate, branchName]
        );
        if (result.error) return true; // Fail open
        return (result.data[0]?.cnt || 0) > 0;
    } catch (err) {
        return true; // Fail open
    }
}

// Check if yesterday's plan exists for a branch
async function checkYesterdayPlanExists(branchName) {
    const yesterdayDate = getYesterdayISO();
    try {
        const result = await neonQuery(
            'SELECT COUNT(*)::int AS cnt FROM daily_reports WHERE date = $1 AND branch_name = $2',
            [yesterdayDate, branchName]
        );
        if (result.error) return false;
        return (result.data[0]?.cnt || 0) > 0;
    } catch (err) {
        return false;
    }
}


// --- REALTIME SUBSCRIPTION ---
// Throttled render to prevent UI thrashing during rapid updates
const throttledRender = throttle(() => {
    if (state.activeTab === 'hierarchy') {
        renderDashboard();
    }
}, 500);

function cleanupRealtimeSubscriptions() {
    // Clear polling interval
    if (state._pollingInterval) {
        clearInterval(state._pollingInterval);
        state._pollingInterval = null;
    }
    state.realtimeChannels = [];
}

function setupRealtime() {
    // Clean up existing polling first
    cleanupRealtimeSubscriptions();

    // Poll for updates every 30 seconds (replaces Neon realtime)
    state._pollingInterval = setInterval(async () => {
        try {
            await fetchData(0, false);
        } catch (err) {
            console.warn('Polling refresh error:', err);
        }
    }, 30000);
}

async function saveData(branchName, branchData, table, retryCount = 0) {
    const MAX_RETRIES = 3;
    const dateToSave = state.systemDate; // Expected YYYY-MM-DD

    // Get V2 meta info: [branch, area, division, region]
    const branchRow = state.rawData.rows.find(r => r[0] === branchName);
    const district = branchRow ? branchRow[1] : "";   // area → stored in district column
    const dm = branchRow ? branchRow[2] : "";          // division → stored in dm_name column
    const region = branchRow ? branchRow[3] : "";

    // Helper: Convert string to number or null (fixes empty string "" vs 0 vs null)
    const toNum = (val) => {
        if (val === "" || val === null || val === undefined) return null;
        const n = Number(val);
        return isNaN(n) ? null : n;
    };

    // Flatten payload with Type Conversion
    const payload = {
        date: dateToSave,
        branch_name: branchName,
        region: region,
        district: district,
        dm_name: dm,

        // Map fields directly with conversion
        ftod_actual: toNum(branchData.ftod_actual),
        ftod_plan: toNum(branchData.ftod_plan),
        dpd_1_30_actual: toNum(branchData.dpd_1_30_actual),
        dpd_1_30_plan: toNum(branchData.dpd_1_30_plan),
        dpd_31_60_actual: toNum(branchData.dpd_31_60_actual),
        dpd_31_60_plan: toNum(branchData.dpd_31_60_plan),
        dpd_61_90_actual: toNum(branchData.dpd_61_90_actual),
        dpd_61_90_plan: toNum(branchData.dpd_61_90_plan),
        npa_activation: toNum(branchData.npa_activation),
        npa_closure: toNum(branchData.npa_closure),
        fy_non_start_acc: toNum(branchData.fy_non_start_acc),
        fy_non_start_plan: toNum(branchData.fy_non_start_plan),

        disb_igl_acc: toNum(branchData.disb_igl_acc),
        disb_igl_amt: toNum(branchData.disb_igl_amt),
        disb_fig_acc: toNum(branchData.disb_fig_acc),
        disb_fig_amt: toNum(branchData.disb_fig_amt),
        disb_il_acc: toNum(branchData.disb_il_acc),
        disb_il_amt: toNum(branchData.disb_il_amt),
        kyc_igl: toNum(branchData.kyc_igl),
        kyc_fig: toNum(branchData.kyc_fig),
        kyc_il: toNum(branchData.kyc_il)
    };

    const columns = Object.keys(payload);
    const values = Object.values(payload);
    const placeholders = columns.map((_, i) => `$${i + 1}`);
    // Quote column names to preserve mixed-case (e.g. dpd_1_30_actual)
    const q = c => `"${c}"`;
    const updateSet = columns
        .filter(c => c !== 'date' && c !== 'branch_name')
        .map(c => `${q(c)} = EXCLUDED.${q(c)}`)
        .join(', ');

    const sql = `INSERT INTO ${table} (${columns.map(q).join(', ')})
VALUES (${placeholders.join(', ')})
ON CONFLICT (date, branch_name) DO UPDATE SET ${updateSet}
RETURNING *`;

    try {
        const result = await neonQuery(sql, values);

        if (result.error) {
            const errorMsg = result.error.message || String(result.error);
            const isNetworkError = errorMsg && (
                errorMsg.includes('Failed to fetch') ||
                errorMsg.includes('Load failed') ||
                errorMsg.includes('NetworkError') ||
                errorMsg.includes('timeout') ||
                errorMsg.includes('ERR_CONNECTION')
            );

            if (isNetworkError && retryCount < MAX_RETRIES) {
                const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
                console.warn(`Save timeout for ${branchName}, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
                showToast(`Connection slow, retrying... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
                await new Promise(resolve => setTimeout(resolve, delay));
                return saveData(branchName, branchData, table, retryCount + 1);
            }

            console.error("Save Error:", result.error);
            if (isNetworkError) {
                alert(`Save Failed: Network connection timed out after ${MAX_RETRIES} retries.\nPlease check your internet connection and try again.`);
            } else {
                alert(`Save Failed: ${errorMsg}\nHint: Check column types.`);
            }
            return { success: false, error: result.error };
        }
        return { success: true, row: (result.data && result.data[0]) ? result.data[0] : payload };
    } catch (err) {
        const isNetworkError = err.message && (
            err.message.includes('Failed to fetch') ||
            err.message.includes('Load failed') ||
            err.message.includes('NetworkError') ||
            err.message.includes('timeout')
        );

        if (isNetworkError && retryCount < MAX_RETRIES) {
            const delay = Math.min(2000 * Math.pow(2, retryCount), 10000);
            console.warn(`Save network error for ${branchName}, retrying in ${delay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})...`);
            showToast(`Connection slow, retrying... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
            await new Promise(resolve => setTimeout(resolve, delay));
            return saveData(branchName, branchData, table, retryCount + 1);
        }

        console.error("Save Error:", err);
        alert(`Save Failed: Network connection error.\nPlease check your internet connection and try again.`);
        return { success: false, error: err };
    }
}

// --- THEME ---
function toggleTheme() {
    document.body.classList.toggle("dark-mode");
}

function toggleSidebar() {
    const sidebar = document.getElementById("sidebar");
    const overlay = document.getElementById("sidebarOverlay");
    sidebar.classList.toggle("open");
    overlay.classList.toggle("visible");
}

// --- ACTIONS ---
function handleLogin(role) {
    let user = "";
    if (role === 'CEO') {
        const btn = document.getElementById("login-btn-ceo");
        if (btn) btn.innerHTML = 'Logging in...';
        user = "CEO";
        state.role = 'CEO';
    } else if (role === 'RM' || role === 'DM' || role === 'AM' || role === 'BM' || role === 'OTHER') {
        var emp = window._drSelectedEmployee;
        if (!emp) return alert("Please search and select an employee.");
        var mappedRole = (role === 'OTHER') ? 'BM' : role;
        var locationMap = { RM: emp.region_name, DM: emp.division_name, AM: emp.area_name, BM: emp.branch_name };
        user = locationMap[mappedRole] || emp.branch_name || '';
        if (!user) return alert("Employee has no location assigned for this role.");
        state.role = mappedRole;
        localStorage.setItem('nlpl_last_dm', user);
        localStorage.setItem('nlpl_last_role', mappedRole);
        localStorage.setItem('nlpl_last_emp', emp.emp_id);
    } else {
        return alert("Please select a role.");
    }
    state.currentUser = user;

    // CEO & RM: Skip date overlay, auto-select today
    if (role === 'CEO' || role === 'RM') {
        if (role === 'RM') state.role = 'CEO'; // RM sees everything like CEO
        autoLoginCEO();
        return;
    }

    // DM/AM/BM: SHOW INTERACTIVE DATE SELECTION OVERLAY
    const dailyOverlay = document.getElementById("daily-overlay");
    const loginOverlay = document.getElementById("login-overlay");

    // 1. Show Selection Overlay
    dailyOverlay.classList.remove("hidden");
    dailyOverlay.style.background = "white"; // Solid background

    // Hide Login Overlay
    loginOverlay.style.opacity = '0';
    setTimeout(() => loginOverlay.classList.add("hidden"), 500);

    // 2. Start Clock and Populate Buttons
    startOverlayClock();
    populateOverlayDates();
}

// CEO Quick Login - Auto-select today's date
async function autoLoginCEO() {
    console.log("autoLoginCEO: Started");
    const loginOverlay = document.getElementById("login-overlay");

    // Hide login overlay immediately
    loginOverlay.style.opacity = '0';
    setTimeout(() => loginOverlay.classList.add("hidden"), 300);

    // Set today's date
    const now = new Date();
    const offsetDate = new Date(now.getTime() - (now.getTimezoneOffset() * 60000));
    const isoDate = offsetDate.toISOString().split('T')[0];
    state.systemDate = isoDate;
    console.log("autoLoginCEO: Date set to " + isoDate);

    // Fetch Data for today
    try {
        console.log("autoLoginCEO: Calling fetchData");
        await fetchData();
        console.log("autoLoginCEO: fetchData returned");
    } catch (e) {
        console.error("autoLoginCEO: fetchData threw", e);
    }

    // Update Header Display
    updateHeaderDate(isoDate);

    // Load App UI
    console.log("autoLoginCEO: Calling loadAppUI");
    loadAppUI();
    console.log("autoLoginCEO: loadAppUI returned");
}


let clockInterval;
function startOverlayClock() {
    const timeEl = document.getElementById("overlay-time");
    const dateEl = document.getElementById("overlay-date");

    const update = () => {
        const now = new Date();
        timeEl.textContent = now.toLocaleTimeString('en-US', { hour12: false });
        dateEl.textContent = now.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
    };
    update();
    clockInterval = setInterval(update, 1000);
}

function populateOverlayDates() {
    const now = new Date();

    // Helper for formatting "12 Jan"
    const fmt = (d) => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });

    // Today
    document.getElementById("btn-date-0").textContent = fmt(now);

    // Tomorrow
    const tom = new Date(now); tom.setDate(tom.getDate() + 1);
    document.getElementById("btn-date-1").textContent = fmt(tom);

    // Day After
    const dat = new Date(now); dat.setDate(dat.getDate() + 2);
    document.getElementById("btn-date-2").textContent = fmt(dat);
}

async function selectSystemDate(offset) {
    clearInterval(clockInterval);

    // Calculate Plan Date
    const targetDateObj = new Date();
    targetDateObj.setDate(targetDateObj.getDate() + offset);

    // Fix: Use local time YYYY-MM-DD instead of ISO (which is UTC)
    const offsetDate = new Date(targetDateObj.getTime() - (targetDateObj.getTimezoneOffset() * 60000));
    const isoDate = offsetDate.toISOString().split('T')[0];

    state.systemDate = isoDate;

    // Fetch Data for this Date (skip render — loadAppUI will trigger it)
    await fetchData(0, true);

    // Update Header Display
    updateHeaderDate(isoDate);

    // Transition to App — load immediately, don't wait for fade animation
    const dailyOverlay = document.getElementById("daily-overlay");
    dailyOverlay.style.opacity = '0';
    dailyOverlay.classList.add("hidden");
    loadAppUI();
}

function loadAppUI() {
    // Show Main Content
    document.getElementById("sidebar").classList.remove("hidden");
    document.querySelector(".main-content").classList.remove("hidden");

    // Update Header Info
    const role = state.role;
    const user = state.currentUser;
    document.getElementById("headerName").textContent = user;
    var roleLabels = { CEO: 'Headquarters', RM: 'Regional Manager', DM: 'Division Manager', AM: 'Area Manager', BM: 'Branch Manager' };
    document.getElementById("headerRole").textContent = roleLabels[role] || role;
    document.getElementById("headerAvatar").textContent = user.charAt(0);

    // Restrict navigation based on role
    if (role !== 'CEO') {
        document.getElementById('nav-dashboard').classList.add('hidden');
        document.getElementById('nav-reports').classList.remove('hidden');
        // Default report level based on role
        state.reportLevel = role === 'DM' ? 'AREA' : 'BRANCH';

        // Hide date picker and search for DM - they cannot use filters
        // const datePicker = document.querySelector('.date-picker-wrapper');
        // if (datePicker) datePicker.style.display = 'none';
        const searchContainer = document.querySelector('.search-container');
        if (searchContainer) searchContainer.style.display = 'none';

        // Auto switch to hierarchy
        switchTab('hierarchy');
    } else {
        document.getElementById('nav-dashboard').classList.remove('hidden');
        document.getElementById('nav-reports').classList.remove('hidden');
        renderDashboard();
    }
    updateToggleVisibility();
}

function handleLogout() {
    location.reload();
}

function switchTab(tab) {
    state.activeTab = tab;
    document.querySelectorAll(".nav-item").forEach(el => el.classList.remove("active"));

    if (tab === 'dashboard') {
        document.getElementById("nav-dashboard").classList.add("active");
        document.getElementById("breadcrumbs").innerHTML = 'Home / <span>Dashboard</span>';
    }
    if (tab === 'hierarchy') {
        document.getElementById("nav-hierarchy").classList.add("active");
        document.getElementById("breadcrumbs").innerHTML = 'Home / <span>Detailed View</span>';
        state.viewStack = []; // Reset drill-down
    }
    if (tab === 'reports') {
        document.getElementById("nav-reports").classList.add("active");
        document.getElementById("breadcrumbs").innerHTML = 'Home / <span>Reports</span>';
    }

    // Auto-close sidebar on mobile if open
    if (window.innerWidth <= 768) {
        document.getElementById("sidebar").classList.remove("open");
        document.getElementById("sidebarOverlay").classList.remove("visible");
    }

    renderDashboard();
    updateToggleVisibility();
}

let toastTimeout = null;
function showToast(msg, icon = 'info') {
    const toast = document.getElementById("toast");
    const toastMsg = document.getElementById("toastMsg");
    if (!toast || !toastMsg) return;

    // Clear any existing timeout to prevent stacking
    if (toastTimeout) {
        clearTimeout(toastTimeout);
        toast.classList.remove("visible");
    }

    toastMsg.textContent = msg;
    toast.classList.add("visible");

    toastTimeout = setTimeout(() => {
        toast.classList.remove("visible");
        toastTimeout = null;
    }, 3000);
}

// --- DRILL DOWN ACTIONS ---
function pushStack(name, data) {
    state.viewStack.push({ name, data });
    renderDashboard();
}

function popStack() {
    state.viewStack.pop();
    renderDashboard();
}

// --- SEED / DELETE DATA ---



// --- RENDERERS ---
function renderDashboard() {
    const container = document.getElementById("dashboard-container");

    // 1. Safe Render Buffer
    const buffer = document.createElement("div");
    buffer.className = "dashboard-grid fade-enter";

    try {
        // 1. DASHBOARD TAB
        if (state.activeTab === 'dashboard') {
            if (state.role === 'CEO') {
                const stats = calculateCEOStats();

                if (state.viewMode === 'PLAN') {
                    renderCEOPlanDashboard(stats, buffer);
                } else {
                    renderCEOReviewDashboard(stats, buffer);
                }

                // --- COMMIT TO DOM ---
                container.innerHTML = "";
                container.appendChild(buffer);

                // Render Charts (Hooks depend on mode)
                if (state.viewMode === 'REVIEW') {
                    renderPortfolioDonut(stats.portfolioHealth, document.getElementById("portfolioDonutHook"));
                    renderRegionalTable(stats.regionalBreakdown, document.getElementById("regionalTableHook"));
                } else {
                    // Plan Mode Charts if any
                    renderRegionalPlanTable(stats.regionalBreakdown, document.getElementById("regionalPlanTableHook"));
                }
            } else {
                const tableCard = document.createElement("div");
                tableCard.className = "table-card grid-full";
                tableCard.innerHTML = `
                            <div class="chart-header">
                                <div class="chart-title">Branch Performance Plans</div>
                                <div style="display:flex; gap:10px;">
                                    <button class="btn btn-outline" onclick="setTargetForTomorrow()" style="width:auto; padding:8px 16px;">Set Target for Tomorrow</button>
                                    <button class="btn btn-primary" onclick="savePlans()" style="width:auto; padding:8px 16px;">Save Changes</button>
                                </div>
                            </div>
                            <table>
                                <thead><tr><th>District</th><th>Branch</th><th>Plan Input</th><th>Status</th></tr></thead>
                                <tbody id="dmTableHook"></tbody>
                            </table>`;

                buffer.appendChild(tableCard);
                container.innerHTML = "";
                container.appendChild(buffer);

                renderDMTable(document.getElementById("dmTableHook"));
            }
        }
        // 3. REPORTS TAB (CEO and DM with role-based restrictions)
        else if (state.activeTab === 'reports') {
            renderReports(buffer);
            container.innerHTML = "";
            container.appendChild(buffer);

            // Initialize Report Charts/Tables if needed
            // For now, renderReports handles static content or simple tables
        }
        // 2. DETAILED VIEW / DRILL DOWN
        else if (state.activeTab === 'hierarchy') {
            const wrapper = document.createElement("div");
            wrapper.className = "table-card grid-full";

            // Determine Current Level
            let currentData = state.fullTree;
            let currentTitle = "All Regions";
            let currentType = "Region";

            // DM SPECIFIC OVERRIDE
            if (state.role === 'DM') {
                const dmNode = getDMNode();
                if (dmNode) {
                    currentData = dmNode;
                    currentTitle = "My Districts";
                    currentType = "District"; // Root for DM is District list
                }
            }

            if (state.viewStack.length > 0) {
                const top = state.viewStack[state.viewStack.length - 1];
                currentData = top.data;
                currentTitle = top.name;

                if (state.role === 'DM') {
                    if (state.viewStack.length === 1) currentType = "Branch";
                } else {
                    // CEO Logic
                    if (state.viewStack.length === 1) currentType = "DM";
                    if (state.viewStack.length === 2) currentType = "District";
                    if (state.viewStack.length === 3) currentType = "Branch";
                }
            }

            // --- HIERARCHY SUMMARY (CEO Only, Non-Leaf) ---
            let summaryHTML = '';
            // Only show summary if not at leaf (Branch) level and is CEO
            if (state.role === 'CEO' && !(currentData instanceof Set)) {
                const allBranches = getAllBranchesFromNode(currentData);
                const aggStats = calculateAggregateStatsForBranches(allBranches);
                summaryHTML = renderHierarchySummaryCard(aggStats, currentTitle);
            }

            // --- DM SUMMARY (District Level) ---
            if (state.role === 'DM' && state.viewStack.length === 1 && currentData instanceof Set) {
                // We are in District View (showing list of branches)
                const branchList = Array.from(currentData);
                const aggStats = calculateAggregateStatsForBranches(branchList);
                summaryHTML = renderHierarchySummaryCard(aggStats, currentTitle);
            }
            // --- DM ROOT SUMMARY (All Districts) ---
            if (state.role === 'DM' && state.viewStack.length === 0) {
                // We are in Root View (showing list of Districts)
                // Need to aggregate ALL branches from these districts
                const allBranches = getAllBranchesFromNode(currentData);
                const aggStats = calculateAggregateStatsForBranches(allBranches);
                summaryHTML = renderHierarchySummaryCard(aggStats, currentTitle);
            }

            // Header with Back Button
            wrapper.innerHTML = `
                        <div class="chart-header" style="justify-content:flex-start; gap:16px;">
                            ${state.viewStack.length > 0 ?
                    `<button class="btn btn-outline" onclick="popStack()" style="width:auto; padding:8px 12px;">
                                    <svg class="icon" viewBox="0 0 24 24"><line x1="19" y1="12" x2="5" y2="12"></line><polyline points="12 19 5 12 12 5"></polyline></svg>
                                 </button>`
                    : ''}
                            <div class="chart-title">${currentTitle}</div>
                        </div>
                        
                        <!-- SUMMARY SECTION -->
                        ${summaryHTML}
                        
                        <div id="drillDownGrid" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap:16px; padding-top:10px;"></div>
                    `;

            buffer.appendChild(wrapper);
            container.innerHTML = "";
            container.appendChild(buffer);

            const ddGrid = document.getElementById("drillDownGrid");

            // Render Items
            if (currentData instanceof Set) {
                // Leaf Nodes (Branches)
                currentData.forEach(name => {
                    ddGrid.appendChild(createDrillDownCard(name, "Branch", null, true));
                });
            } else {
                Object.keys(currentData).sort().forEach(key => {
                    ddGrid.appendChild(createDrillDownCard(key, currentType, currentData[key], false));
                });
            }
        }
    } catch (err) {
        console.error("Render Error:", err);
        showToast("Error rendering dashboard. Please refresh.", "alert");
    }
}

function renderReports(buffer) {
    // DM-specific rendering with restrictions
    if (state.role === 'DM') {
        renderDMReports(buffer);
        return;
    }

    // CEO Full Access UI
    const pct = calculateTotalCollectionPercentage();
    buffer.innerHTML = `
        <div class="chart-card grid-full">
            <div class="chart-header">
                <div class="chart-title">📊 Custom Report Builder</div>
                <div style="font-size:12px; color:var(--text-secondary);">Select the data points you want to include in your report.</div>
            </div>

            <!-- 1. DATE SELECTION & LEVEL -->
            <div style="padding: 16px; border-bottom: 1px solid var(--border-color);">
                <div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:16px;">
                    <div>
                        <div style="font-size:14px; font-weight:600; margin-bottom:12px;">1. Select Date Range</div>
                        <div style="display:flex; gap:12px; align-items:center;">
                            <div style="display:flex; gap:8px;">
                                <button class="btn btn-outline" onclick="setReportDate(-1)">Yesterday</button>
                                <button class="btn btn-primary" onclick="setReportDate(0)">Today</button>
                                <button class="btn btn-outline" onclick="setReportDate(1)">Tomorrow</button>
                            </div>
                            <div style="width:1px; height:24px; background:var(--border-color); margin:0 8px;"></div>
                            <input type="date" id="reportDateInput" value="${state.systemDate}"
                                onchange="state.systemDate = this.value; updateHeaderDate(this.value); debouncedFetchData();"
                                style="padding:8px; border:1px solid var(--border-color); border-radius:6px; background:var(--bg-body); color:var(--text-primary);">
                        </div>
                    </div>

                    <div>
                        <div style="font-size:14px; font-weight:600; margin-bottom:12px;">2. Report Level</div>
                        <div style="display:flex; gap:8px;">
                            <button class="btn ${state.reportLevel === 'REGION' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('REGION')">Region Level</button>
                            <button class="btn ${state.reportLevel === 'DIVISION' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('DIVISION')">Division Level</button>
                            <button class="btn ${state.reportLevel === 'AREA' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('AREA')">Area Level</button>
                            <button class="btn ${state.reportLevel === 'BRANCH' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('BRANCH')">Branch Level</button>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 2. DATA SELECTION (REMOVED) -->
             <div style="padding: 16px; display:none;">
                <!-- Options removed as per request -->
             </div>

            <!-- 3. ACTIONS -->
            <div style="padding: 16px; border-top: 1px solid var(--border-color);">
                <div style="display:flex; justify-content:center; align-items:center; flex-wrap:wrap; gap:12px;">
                    <button class="btn btn-outline" onclick="handleGeneratePlanReport()" style="padding:12px 24px; border-color: var(--primary); color: var(--primary); font-size: 15px; font-weight: 600;">
                        <svg class="icon" viewBox="0 0 24 24" style="stroke: currentColor; width: 18px; height: 18px;"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg>
                        Generate Plan Report
                    </button>
                    <button class="btn btn-primary" onclick="handleGenerateBothReports()" style="padding:12px 32px; background: linear-gradient(135deg, #6366F1, #4F46E5); color: white; border: none; font-size: 16px; font-weight: 600; box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3);">
                        <svg class="icon" viewBox="0 0 24 24" style="stroke: white; width: 20px; height: 20px;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        Generate Plan & Achievement Reports ${pct > 0 ? `(${pct}%)` : ''}
                    </button>
                </div>
            </div>
        </div>
        <style>
            .checkbox-row { display: flex; align-items: center; gap: 8px; font-size: 13px; margin-bottom: 6px; cursor: pointer; }
            .checkbox-row input[type="checkbox"] { accent-color: var(--primary); width: 16px; height: 16px; }
        </style>
    `;
}

// DM Restricted Reports View
function renderDMReports(buffer) {
    const dmDistricts = getDMDistricts();
    const dmBranches = getDMBranches();
    const districtDisplay = dmDistricts.length > 0 ? dmDistricts.join(', ') : 'Your District';
    const branchCount = dmBranches.length;

    // Ensure DM cannot have REGION or DIVISION level set
    if (state.reportLevel === 'REGION' || state.reportLevel === 'DIVISION') {
        state.reportLevel = 'AREA';
    }

    const pct = calculateTotalCollectionPercentage();
    buffer.innerHTML = `
        <div class="chart-card grid-full">
            <div class="chart-header">
                <div class="chart-title">📊 My Reports</div>
                <div style="font-size:12px; color:var(--text-secondary);">View reports for your assigned district and branches.</div>
            </div>

            <!-- DM INFO BANNER -->
            <div style="padding: 16px; background: linear-gradient(135deg, rgba(99, 102, 241, 0.1), rgba(79, 70, 229, 0.05)); border-bottom: 1px solid var(--border-color);">
                <div style="display:flex; align-items:center; gap:12px;">
                    <div style="width:40px; height:40px; background:var(--primary-accent); border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-weight:600;">
                        ${state.currentUser ? state.currentUser.charAt(0) : 'D'}
                    </div>
                    <div>
                        <div style="font-weight:600; color:var(--text-primary);">${state.currentUser || 'Division Manager'}</div>
                        <div style="font-size:12px; color:var(--text-secondary);">District: ${districtDisplay} • ${branchCount} Branch${branchCount !== 1 ? 'es' : ''}</div>
                    </div>
                </div>
            </div>

            <!-- REPORT LEVEL (DM sees Area + Branch) -->
            <div style="padding: 16px; border-bottom: 1px solid var(--border-color);">
                <div style="font-size:14px; font-weight:600; margin-bottom:12px;">Report View</div>
                <div style="display:flex; gap:8px;">
                    <button class="btn ${state.reportLevel === 'AREA' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('AREA')">
                        📍 Area View
                    </button>
                    <button class="btn ${state.reportLevel === 'BRANCH' ? 'btn-primary' : 'btn-outline'}" onclick="setReportLevel('BRANCH')">
                        🏢 Branch View
                    </button>
                </div>
                <div style="margin-top:12px; padding:10px; background:var(--bg-body); border-radius:8px; font-size:12px; color:var(--text-secondary);">
                    ${state.reportLevel === 'AREA'
            ? `<strong>Area View:</strong> Showing aggregated data for <strong>${districtDisplay}</strong>`
            : `<strong>Branch View:</strong> Showing ${branchCount} branch${branchCount !== 1 ? 'es' : ''} under your division`
        }
                </div>
            </div>

            <!-- ACTIONS -->
            <div style="padding: 16px; border-top: 1px solid var(--border-color);">
                <div style="display:flex; justify-content:center; align-items:center; flex-wrap:wrap; gap:12px;">
                    <button class="btn btn-outline" onclick="handleGeneratePlanReport()" style="padding:12px 24px; border-color: var(--primary); color: var(--primary); font-size: 15px; font-weight: 600;">
                        <svg class="icon" viewBox="0 0 24 24" style="stroke: currentColor; width: 18px; height: 18px;"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg>
                        Generate Plan Report
                    </button>
                    <button class="btn btn-primary" onclick="handleGenerateBothReports()" style="padding:12px 32px; background: linear-gradient(135deg, #6366F1, #4F46E5); color: white; border: none; font-size: 16px; font-weight: 600; box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3);">
                        <svg class="icon" viewBox="0 0 24 24" style="stroke: white; width: 20px; height: 20px;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                        Generate Plan & Achievement Reports ${pct > 0 ? `(${pct}%)` : ''}
                    </button>
                </div>
            </div>
        </div>
    `;
}

function setReportDate(offset) {
    const d = new Date();
    d.setDate(d.getDate() + offset);
    // Adjust for timezone offset to keep date correct
    const offsetDate = new Date(d.getTime() - (d.getTimezoneOffset() * 60000));
    const isoDate = offsetDate.toISOString().split('T')[0];

    // Update global state
    state.systemDate = isoDate;

    // FIX: Update range to match single selected date
    state.dateFrom = isoDate;
    state.dateTo = isoDate;
    saveDateRange(isoDate, isoDate); // Persist if needed


    // Update inputs
    const input = document.getElementById("reportDateInput");
    if (input) input.value = isoDate;

    // Update header
    updateHeaderDate(isoDate);

    // Refresh UI to reflect date change (highlight button etc - optional)
    // Ideally we might want to re-fetch data if it wasn't pre-loaded, but for now we assume data changes with date
    // Trigger data refresh if needed:
    debouncedFetchData(); // Re-fetch for the new date (debounced to prevent rapid-fire)
}

// --- COLOR HELPER FOR REPORTS ---
function getCellColor(achievement, plan) {
    if (!plan || plan === 0) return '#FFFFFF'; // White for no plan
    const percentage = (achievement / plan) * 100;

    if (percentage >= 100) return '#D1FAE5'; // Green - Met or exceeded
    if (percentage >= 50) return '#FEF3C7';   // Yellow - 50-99%
    return '#FECACA';                          // Light red - Below 50%
}

// --- REPORT AGGREGATION HELPER ---
function getReportRows(level) {
    const rows = [];
    // V2 headers: Branch(0), Area(1), Division(2), Region(3)
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxArea = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'area');
    const idxDivision = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'division');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    // Helper to sum objects (only numeric values)
    const sumObjects = (acc, curr) => {
        if (!curr) return acc;
        Object.keys(curr).forEach(key => {
            if (['id', 'branch_name', 'date', 'region', 'district', 'dm_name', 'created_at'].includes(key)) return;
            const val = Number(curr[key]);
            if (!isNaN(val)) {
                acc[key] = (Number(acc[key]) || 0) + val;
            }
        });
        return acc;
    };

    // Role-based filtering
    const roleBranches = (state.role !== 'CEO') ? getDMBranches() : null;

    if (level === 'BRANCH') {
        state.rawData.rows.forEach(r => {
            const name = r[idxBranch];
            if (roleBranches && !roleBranches.includes(name)) return;
            const entry = state.branchDetails[name] || {};
            rows.push({
                name: name,
                region: r[idxRegion],
                district: r[idxArea],
                dm: r[idxDivision],
                target: entry.target || {},
                achievement: entry.achievement || {}
            });
        });
        return rows;
    }

    // Aggregation for REGION / DIVISION / AREA
    const groups = {};

    state.rawData.rows.forEach(r => {
        const branchName = r[idxBranch];
        const region = r[idxRegion] || "Unknown";
        const division = r[idxDivision] || "Unknown";
        const area = r[idxArea] || "Unknown";

        if (roleBranches && !roleBranches.includes(branchName)) return;

        let key;
        if (level === 'REGION') {
            key = region;
        } else if (level === 'DIVISION') {
            key = division;
        } else if (level === 'AREA') {
            key = area;
        } else if (level === 'DISTRICT') {
            // Legacy fallback — treat as AREA
            key = area;
        }

        if (!key) return;

        if (!groups[key]) {
            groups[key] = {
                name: key,
                region: region,
                district: level === 'AREA' || level === 'DISTRICT' ? key : (level === 'DIVISION' ? key : 'All'),
                dm: division,
                target: {},
                achievement: {}
            };
        }

        const entry = state.branchDetails[branchName];
        if (entry) {
            if (entry.target) groups[key].target = sumObjects(groups[key].target, entry.target);
            if (entry.achievement) groups[key].achievement = sumObjects(groups[key].achievement, entry.achievement);
        }
    });

    return Object.values(groups).sort((a, b) => a.name.localeCompare(b.name));
}

// --- REPORT GENERATION ---

// --- REPORT PREVIEW MODAL ---
function showReportPreviewModal(htmlContent, title, filename) {
    const modal = document.getElementById('reportPreviewModal');
    const body = document.getElementById('reportPreviewBody');
    const titleEl = document.getElementById('reportPreviewTitle');
    const downloadBtn = document.getElementById('btnCopyReport');

    if (!modal || !body || !titleEl || !downloadBtn) {
        console.error("Report Preview Modal elements not found");
        return;
    }

    // Set Content
    titleEl.textContent = title;
    body.innerHTML = htmlContent;

    // Set Copy Action
    downloadBtn.onclick = async () => {
        // Start animation and indicate copying
        downloadBtn.textContent = "Copying...";
        downloadBtn.disabled = true;
        downloadBtn.classList.add('btn-success-anim');

        // Attempt to copy image (mobile & desktop)
        const success = await copyTableImageToClipboard(htmlContent);

        if (success) {
            downloadBtn.textContent = "Copied! ✅";
        } else {
            downloadBtn.textContent = "Copy";
        }
        // Reset after a short period
        setTimeout(() => {
            downloadBtn.textContent = "Copy";
            downloadBtn.disabled = false;
            downloadBtn.classList.remove('btn-success-anim');
        }, 2000);
    };

    // Set Share Action
    const shareBtn = document.getElementById('btnShareReport');
    if (shareBtn) {
        shareBtn.onclick = async () => {
            shareBtn.disabled = true;
            shareBtn.textContent = "Processing...";
            await shareReportImage(htmlContent);
            shareBtn.textContent = "Share Image";
            shareBtn.disabled = false;
        };
    }

    // Set Excel Action
    const excelBtn = document.getElementById('btnExcelReport');
    if (excelBtn) {
        excelBtn.onclick = () => {
            downloadReportAsExcel(title, filename);
        };
    }

    // Show
    modal.classList.add('visible');
    document.body.style.overflow = 'hidden'; // Prevent background scrolling
}

function openReportInNewTab() {
    const body = document.getElementById('reportPreviewBody');
    if (!body) return;

    const htmlContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Report Preview</title>
            <style>
                body { margin: 20px; font-family: Arial, sans-serif; background: white; }
                table { border-collapse: collapse; width: 100%; }
                /* Add basic styles to ensure it looks good */
            </style>
        </head>
        <body>
            ${body.innerHTML}
        </body>
        </html>
    `;

    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    window.open(url, '_blank');
}

function closeReportPreviewModal() {
    const modal = document.getElementById('reportPreviewModal');
    if (modal) {
        modal.classList.remove('visible');
        document.body.style.overflow = '';
    }
}

async function downloadReportAsExcel(title, filename) {
    const table = document.getElementById('reportPreviewBody').querySelector('table');
    if (!table) {
        showToast("No report table found to export", "alert");
        return;
    }

    try {
        showToast("Preparing Excel...", "info");
        await loadScript('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
        // Create workbook from table
        const wb = XLSX.utils.table_to_book(table, { sheet: "Report" });

        // Generate Excel file
        XLSX.writeFile(wb, (filename || 'Report') + '.xlsx');

        showToast("Report downloaded as Excel! 📊");
    } catch (error) {
        console.error("Excel export failed:", error);
        showToast("Failed to export Excel. Please try again.", "alert");
    }
}

// Generate Only Plan Report
async function handleGeneratePlanReport() {
    // 1. Validate
    if (!state.branchDetails) state.branchDetails = {};

    const level = state.reportLevel; // REGION, DISTRICT, BRANCH
    const dateStr = state.systemDate;
    const dateDisplay = formatDateForDisplay(dateStr);

    setLoading(true, "Generating Plan Report...");

    try {
        // GENERATE PLAN REPORT PREVIEW
        const type = 'PLAN';
        const isPlan = true;
        const reportTitle = `${type} – ${level} – ${dateDisplay}`;
        const filename = `${type.toLowerCase()}_${level.toLowerCase()}_${dateStr}.png`;

        // 1. Get Aggregated Data
        const rows = getAggregatedReportData(level, isPlan);

        // 2. Generate HTML
        const html = generateReportHTML(reportTitle, level, rows, isPlan);

        // 3. Show Preview Modal
        showReportPreviewModal(html, reportTitle, filename);
        setLoading(false);
    } catch (e) {
        console.error("Report Generation Error:", e);
        showToast("Error generating report.", "alert");
        setLoading(false);
    }
}

// Generate Both Plan & Achievement Reports (Combined Side-by-Side)
async function handleGenerateBothReports() {
    // 1. Validate
    if (!state.branchDetails) state.branchDetails = {};

    const level = state.reportLevel; // REGION, DISTRICT, BRANCH
    const dateStr = state.systemDate;
    const dateDisplay = formatDateForDisplay(dateStr);

    setLoading(true, "Generating Combined Report...");

    try {
        // 1. Get Data for both
        const planRows = getAggregatedReportData(level, true);
        const achieveRows = getAggregatedReportData(level, false);

        // 2. Generate Combined HTML
        const reportTitle = `Plan vs Achievement – ${level} – ${dateDisplay}`;
        const html = generateCombinedReportHTML(reportTitle, level, planRows, achieveRows);

        // 3. Show Preview Modal
        // plan_vs_achievement_region_2023-10-25.png
        const filename = `plan_vs_achievement_${level.toLowerCase()}_${dateStr}.png`;

        showReportPreviewModal(html, reportTitle, filename);
        setLoading(false);

    } catch (e) {
        console.error("Report Generation Error:", e);
        showToast("Error generating report.", "alert");
        setLoading(false);
    }
}

// Format date for display (DD-MM-YYYY)
function formatDateForDisplay(dateStr) {
    const d = new Date(dateStr);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}-${month}-${year}`;
}

// Get data filtered by Type (Plan/Achievement) and Aggregated by Level
function getAggregatedReportData(level, isPlan) {
    const rows = [];
    // V2 headers: Branch(0), Area(1), Division(2), Region(3)
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxArea = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'area');
    const idxDivision = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'division');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    // Helper to sum objects (only numeric values)
    const sumObjects = (acc, curr) => {
        if (!curr) return acc;
        Object.keys(curr).forEach(key => {
            if (['id', 'branch_name', 'date', 'region', 'district', 'dm_name', 'created_at'].includes(key)) return;
            const val = Number(curr[key]);
            if (!isNaN(val)) {
                acc[key] = (Number(acc[key]) || 0) + val;
            }
        });
        return acc;
    };

    // Role-based filtering
    const roleBranches = (state.role !== 'CEO') ? getDMBranches() : null;

    // 1. BRANCH LEVEL (No Aggregation, just Filter)
    if (level === 'BRANCH') {
        state.rawData.rows.forEach(r => {
            const name = r[idxBranch];
            const region = r[idxRegion];
            const area = r[idxArea];

            if (roleBranches && !roleBranches.includes(name)) return;

            const entry = state.branchDetails[name] || {};
            const data = isPlan ? (entry.target || {}) : (entry.achievement || {});

            rows.push({
                name: name,
                region: region,
                district: area,
                data: data
            });
        });
        return rows.sort((a, b) => a.name.localeCompare(b.name));
    }

    // 2. REGION / DIVISION / AREA LEVEL (Aggregation)
    const groups = {};

    state.rawData.rows.forEach(r => {
        const branchName = r[idxBranch];
        const region = r[idxRegion] || "Unknown";
        const division = r[idxDivision] || "Unknown";
        const area = r[idxArea] || "Unknown";

        if (roleBranches && !roleBranches.includes(branchName)) return;

        let key;
        if (level === 'REGION') {
            key = region;
        } else if (level === 'DIVISION') {
            key = division;
        } else if (level === 'AREA' || level === 'DISTRICT') {
            key = area;
        }

        if (!key) return;

        if (!groups[key]) {
            groups[key] = {
                name: key,
                region: region,
                district: level === 'AREA' || level === 'DISTRICT' ? key : (level === 'DIVISION' ? key : 'All'),
                data: {}
            };
        }

        const entry = state.branchDetails[branchName];
        if (entry) {
            const source = isPlan ? entry.target : entry.achievement;
            if (source) {
                groups[key].data = sumObjects(groups[key].data, source);
            }
        }
    });

    return Object.values(groups).sort((a, b) => a.name.localeCompare(b.name));
}

// Generate Combined HTML Table string (Side-by-Side)
function generateCombinedReportHTML(title, level, planRows, achieveRows) {
    // Colors
    const colors = {
        white: '#FFFFFF',
        header: '#4682B4',
        planBg: '#F0F8FF',      // AliceBlue for Plan section header
        achieveBg: '#F0FFF0',   // HoneyDew for Achievement section header

        // Data Columns
        ftod: '#E0F2FE',
        slipped: '#DCFCE7',
        pnpa: '#FCE7F3',
        npa: '#FEF9C3',
        fy2526: '#F3F4F6',
        disb: '#FFEDD5',
        kyc: '#F3E8FF',

        border: '#000000'
    };

    const firstColHeader = level === 'BRANCH' ? 'BRANCH' : level === 'AREA' ? 'AREA' : level === 'DIVISION' ? 'DIVISION' : level === 'DISTRICT' ? 'AREA' : 'REGION';
    const fmt = n => n === 0 ? '-' : n.toLocaleString('en-IN');

    // 1. Merge Rows by Name
    const mergedMap = new Map();
    planRows.forEach(r => mergedMap.set(r.name, { name: r.name, plan: r.data || {}, achieve: {} }));
    achieveRows.forEach(r => {
        if (!mergedMap.has(r.name)) {
            mergedMap.set(r.name, { name: r.name, plan: {}, achieve: r.data || {} });
        } else {
            mergedMap.get(r.name).achieve = r.data || {};
        }
    });

    // Sort by name
    const sortedRows = Array.from(mergedMap.values()).sort((a, b) => a.name.localeCompare(b.name));

    // Totals
    const tPlan = { ftod: 0, dpd130: 0, dpd3160: 0, dpd6190: 0, npaAct: 0, npaClose: 0, ns: 0, disbIglAcc: 0, disbIglAmt: 0, disbFigAcc: 0, disbFigAmt: 0, disbIlAcc: 0, disbIlAmt: 0, kycIgl: 0, kycFig: 0, kycIl: 0 };
    const tAch = { ftod: 0, dpd130: 0, dpd3160: 0, dpd6190: 0, npaAct: 0, npaClose: 0, ns: 0, disbIglAcc: 0, disbIglAmt: 0, disbFigAcc: 0, disbFigAmt: 0, disbIlAcc: 0, disbIlAmt: 0, kycIgl: 0, kycFig: 0, kycIl: 0 };

    let bodyRows = '';
    let totalRowAchievementSum = 0;
    let rowCountForAvg = 0;
    let rowIndex = 0;

    sortedRows.forEach(row => {
        const p = row.plan;
        const a = row.achieve;

        // Alternating row color for better readability
        const rowBgColor = rowIndex % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
        rowIndex++;

        // --- EXTRACT VALUES ---

        // PLAN Values
        const p_ftod = parseInt(p.ftod_plan) || 0;
        const p_dpd130 = parseInt(p.dpd_1_30_plan) || 0;
        const p_dpd3160 = parseInt(p.dpd_31_60_plan) || 0;
        const p_dpd6190 = parseInt(p.dpd_61_90_plan) || 0;
        const p_npaAct = parseInt(p.npa_activation) || 0;
        const p_npaClose = parseInt(p.npa_closure) || 0;
        const p_ns = parseInt(p.fy_non_start_plan) || 0;
        const p_disbIglAcc = parseInt(p.disb_igl_acc) || 0;
        const p_disbIglAmt = parseInt(p.disb_igl_amt) || 0;
        const p_disbFigAcc = parseInt(p.disb_fig_acc) || 0;
        const p_disbFigAmt = parseInt(p.disb_fig_amt) || 0;
        const p_disbIlAcc = parseInt(p.disb_il_acc) || 0;
        const p_disbIlAmt = parseInt(p.disb_il_amt) || 0;
        const p_kycIgl = parseInt(p.kyc_igl) || 0;
        const p_kycFig = parseInt(p.kyc_fig) || 0;
        const p_kycIl = parseInt(p.kyc_il) || 0;

        // ACHIEVE Values
        const a_ftod = parseInt(a.ftod_actual) || 0;
        const a_dpd130 = parseInt(a.dpd_1_30_actual) || 0;
        const a_dpd3160 = parseInt(a.dpd_31_60_actual) || 0;
        const a_dpd6190 = parseInt(a.dpd_61_90_actual) || 0;
        const a_npaAct = parseInt(a.npa_activation) || 0;
        const a_npaClose = parseInt(a.npa_closure) || 0;
        const a_ns = parseInt(a.fy_non_start_acc) || 0;
        const a_disbIglAcc = parseInt(a.disb_igl_acc) || 0;
        const a_disbIglAmt = parseInt(a.disb_igl_amt) || 0;
        const a_disbFigAcc = parseInt(a.disb_fig_acc) || 0;
        const a_disbFigAmt = parseInt(a.disb_fig_amt) || 0;
        const a_disbIlAcc = parseInt(a.disb_il_acc) || 0;
        const a_disbIlAmt = parseInt(a.disb_il_amt) || 0;
        const a_kycIgl = parseInt(a.kyc_igl) || 0;
        const a_kycFig = parseInt(a.kyc_fig) || 0;
        const a_kycIl = parseInt(a.kyc_il) || 0;

        // Update Totals
        tPlan.ftod += p_ftod; tPlan.dpd130 += p_dpd130; tPlan.dpd3160 += p_dpd3160; tPlan.dpd6190 += p_dpd6190;
        tPlan.npaAct += p_npaAct; tPlan.npaClose += p_npaClose;
        tPlan.ns += p_ns;
        tPlan.disbIglAcc += p_disbIglAcc; tPlan.disbIglAmt += p_disbIglAmt;
        tPlan.disbFigAcc += p_disbFigAcc; tPlan.disbFigAmt += p_disbFigAmt;
        tPlan.disbIlAcc += p_disbIlAcc; tPlan.disbIlAmt += p_disbIlAmt;
        tPlan.kycIgl += p_kycIgl; tPlan.kycFig += p_kycFig; tPlan.kycIl += p_kycIl;

        tAch.ftod += a_ftod; tAch.dpd130 += a_dpd130; tAch.dpd3160 += a_dpd3160; tAch.dpd6190 += a_dpd6190;
        tAch.npaAct += a_npaAct; tAch.npaClose += a_npaClose;
        tAch.ns += a_ns;
        tAch.disbIglAcc += a_disbIglAcc; tAch.disbIglAmt += a_disbIglAmt;
        tAch.disbFigAcc += a_disbFigAcc; tAch.disbFigAmt += a_disbFigAmt;
        tAch.disbIlAcc += a_disbIlAcc; tAch.disbIlAmt += a_disbIlAmt;
        tAch.kycIgl += a_kycIgl; tAch.kycFig += a_kycFig; tAch.kycIl += a_kycIl;

        // Calc Row Achievement % for Avg
        const rowCollPlan = p_ftod + p_dpd130 + p_dpd3160 + p_dpd6190;
        const rowCollAch = a_ftod + a_dpd130 + a_dpd3160 + a_dpd6190;
        const rowPct = rowCollPlan > 0 ? (rowCollAch / rowCollPlan) * 100 : 0;

        if (rowCollPlan > 0) {
            totalRowAchievementSum += rowPct;
            rowCountForAvg++;
        }

        bodyRows += `
            <tr style="font-size: 10px; background: ${rowBgColor};">
                <td style="background: ${rowBgColor}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: left; font-weight: 600;">${row.name}</td>

                <!-- PLAN DATA -->
                <td style="background: ${colors.ftod}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_ftod)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_dpd130)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_dpd3160)}</td>
                <td style="background: ${colors.pnpa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_dpd6190)}</td>
                <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_npaAct)}</td>
                <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_npaClose)}</td>
                <td style="background: ${colors.fy2526}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_ns)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_disbIglAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(p_disbIglAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_disbFigAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(p_disbFigAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_disbIlAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(p_disbIlAmt)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_kycIgl)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_kycFig)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(p_kycIl)}</td>

                <!-- SEPARATOR -->
                <td style="background: #000; width: 2px; padding: 0;"></td>

                <!-- ACHIEVEMENT DATA -->
                <td style="background: ${colors.ftod}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_ftod)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_dpd130)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_dpd3160)}</td>
                <td style="background: ${colors.pnpa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_dpd6190)}</td>
                <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_npaAct)}</td>
                <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_npaClose)}</td>
                <td style="background: ${colors.fy2526}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_ns)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_disbIglAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(a_disbIglAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_disbFigAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(a_disbFigAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_disbIlAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(a_disbIlAmt)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_kycIgl)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_kycFig)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(a_kycIl)}</td>
            </tr>
        `;
    });

    // --- EXECUTIVE SUMMARY METRICS ---
    const totalCollPlan = tPlan.ftod + tPlan.dpd130 + tPlan.dpd3160 + tPlan.dpd6190;
    const totalCollAch = tAch.ftod + tAch.dpd130 + tAch.dpd3160 + tAch.dpd6190;
    const totalCollPct = totalCollPlan > 0 ? Math.round((totalCollAch / totalCollPlan) * 100) : 0;

    const totalDisbAmt = tAch.disbIglAmt + tAch.disbIlAmt;
    const netNpaChange = tAch.npaAct - tAch.npaClose; // Net = Activation - Closure
    const avgAchPct = rowCountForAvg > 0 ? Math.round(totalRowAchievementSum / rowCountForAvg) : 0;

    const netNpaLabel = netNpaChange > 0 ? `+${netNpaChange}` : `${netNpaChange}`;
    const netNpaColor = netNpaChange > 0 ? '#EF4444' : '#10B981'; // Red if increased, Green if decreased

    const summaryHTML = `
        <div style="display:flex; justify-content:space-around; margin-bottom:20px; font-family:Arial, sans-serif;">
            <div style="border:1px solid #ddd; border-radius:8px; padding:16px; width:22%; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size:12px; color:#666; font-weight:600; text-transform:uppercase;">Total Collection</div>
                <div style="font-size:24px; font-weight:700; color:#4F46E5; margin-top:8px;">${totalCollPct}%</div>
                <div style="font-size:10px; color:#999; margin-top:4px;">${fmt(totalCollAch)} / ${fmt(totalCollPlan)}</div>
            </div>
            <div style="border:1px solid #ddd; border-radius:8px; padding:16px; width:22%; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size:12px; color:#666; font-weight:600; text-transform:uppercase;">Total Disbursement</div>
                <div style="font-size:24px; font-weight:700; color:#10B981; margin-top:8px;">₹${(totalDisbAmt / 10000000).toFixed(2)}Cr</div>
                <div style="font-size:10px; color:#999; margin-top:4px;">IGL & IL</div>
            </div>
            <div style="border:1px solid #ddd; border-radius:8px; padding:16px; width:22%; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size:12px; color:#666; font-weight:600; text-transform:uppercase;">Net NPA Change</div>
                <div style="font-size:24px; font-weight:700; color:${netNpaColor}; margin-top:8px;">${netNpaLabel}</div>
                <div style="font-size:10px; color:#999; margin-top:4px;">Act: ${tAch.npaAct} | Close: ${tAch.npaClose}</div>
            </div>
            <div style="border:1px solid #ddd; border-radius:8px; padding:16px; width:22%; text-align:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size:12px; color:#666; font-weight:600; text-transform:uppercase;">Avg Achievement</div>
                <div style="font-size:24px; font-weight:700; color:#F59E0B; margin-top:8px;">${avgAchPct}%</div>
                <div style="font-size:10px; color:#999; margin-top:4px;">Per Branch/Row</div>
            </div>
        </div>
    `;

    // Grand Total Row
    const totalRow = `
        <tr style="font-weight: bold; font-size: 10px; background: #FFFF00;">
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: left;">Grand Total</td>

            <!-- PLAN TOTALS -->
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.ftod)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.dpd130)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.dpd3160)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.dpd6190)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.npaAct)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.npaClose)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.ns)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.disbIglAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tPlan.disbIglAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.disbFigAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tPlan.disbFigAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.disbIlAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tPlan.disbIlAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.kycIgl)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.kycFig)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tPlan.kycIl)}</td>

             <!-- SEPARATOR -->
            <td style="background: #000;"></td>

            <!-- ACHIEVE TOTALS -->
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.ftod)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.dpd130)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.dpd3160)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.dpd6190)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.npaAct)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.npaClose)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.ns)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.disbIglAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tAch.disbIglAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.disbFigAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tAch.disbFigAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.disbIlAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: right;">${fmt(tAch.disbIlAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.kycIgl)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.kycFig)}</td>
            <td style="background: #FFFF00; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center;">${fmt(tAch.kycIl)}</td>
        </tr>
    `;

    // Category Headers Row (merged cells for grouped columns)
    const categoryHeaders = `
        <td style="background: ${colors.ftod}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">FTOD</td>
        <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">1-30 DPD</td>
        <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">31-60 DPD</td>
        <td style="background: ${colors.pnpa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">61-90 DPD</td>
        <td colspan="2" style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">NPA Accounts</td>
        <td style="background: ${colors.fy2526}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">FY NS</td>
        <td colspan="6" style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">Disbursement</td>
        <td colspan="3" style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:9px; font-weight:bold;">KYC Sourcing</td>
    `;

    // Sub-column Headers Row (matching Set Target data entry names)
    const subHeaders = `
        <td style="background: ${colors.ftod}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Accounts</td>
        <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Accounts</td>
        <td style="background: ${colors.slipped}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Accounts</td>
        <td style="background: ${colors.pnpa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Accounts</td>
        <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Activation</td>
        <td style="background: ${colors.npa}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">Closure</td>
        <td style="background: ${colors.fy2526}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">NS Acc</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IGL Acc</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IGL Amt</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">FIG Acc</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">FIG Amt</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IL Acc</td>
        <td style="background: ${colors.disb}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IL Amt</td>
        <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IGL</td>
        <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">FIG</td>
        <td style="background: ${colors.kyc}; border: 1px solid ${colors.border}; padding: 3px 6px; text-align: center; font-size:8px;">IL</td>
    `;

    return `
        <div style="background:white; padding:20px;">
            <!-- Title -->
            <div style="text-align: center; font-weight: bold; font-size: 18px; padding: 12px; font-family:Arial, sans-serif; margin-bottom:20px;">
                ${title}
            </div>

            <!-- EXECUTIVE SUMMARY -->
            ${summaryHTML}

            <!-- MAIN TABLE -->
            <table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; width: 100%; background:white;">
            <!-- Row 1: PLAN / ACHIEVEMENT Headers -->
            <tr style="font-weight: bold; font-size: 11px;">
                <td rowspan="3" style="background: ${colors.white}; border: 1px solid ${colors.border}; padding: 4px 8px; text-align: center; vertical-align:middle; width: 120px;">${firstColHeader}</td>

                <!-- PLAN HEADER -->
                <td colspan="13" style="background: ${colors.planBg}; border: 1px solid ${colors.border}; padding: 6px; text-align: center; border-bottom: 2px solid ${colors.border}; color:black;">
                    PLAN
                </td>

                <td rowspan="3" style="background: #000; width: 2px;"></td>

                <!-- ACHIEVEMENT HEADER -->
                <td colspan="13" style="background: ${colors.achieveBg}; border: 1px solid ${colors.border}; padding: 6px; text-align: center; border-bottom: 2px solid ${colors.border}; color:black;">
                    ACHIEVEMENT
                </td>
            </tr>

            <!-- Row 2: Category Headers (merged cells) -->
            <tr style="font-weight: bold;">
                ${categoryHeaders}
                ${categoryHeaders}
            </tr>

            <!-- Row 3: Sub-column Headers -->
            <tr style="font-weight: bold;">
                ${subHeaders}
                ${subHeaders}
            </tr>

            ${bodyRows}
            ${totalRow}
            </table>
        </div>
    `;
}

// Generate HTML Table string (Original Single)
function generateReportHTML(title, level, rows, isPlan) {
    // Colors
    const colors = {
        title: '#000000',
        ftod: '#87CEEB',      // Light blue
        slipped: '#90EE90',   // Light green  
        pnpa: '#FFB6C1',      // Light pink
        npa: '#FFFACD',       // Light yellow
        fy2526: '#E0E0E0',    // Light gray
        disb: '#FFDAB9',      // Peach
        kyc: '#DDA0DD',       // Light purple
        white: '#FFFFFF',
        header: '#4682B4'     // Steel blue for main header
    };

    const firstColHeader = level === 'BRANCH' ? 'BRANCH' : level === 'AREA' ? 'AREA' : level === 'DIVISION' ? 'DIVISION' : level === 'DISTRICT' ? 'AREA' : 'REGION';

    // Labels based on Type
    const suffix = isPlan ? 'Plan' : 'Actual';
    const npaLabel = isPlan ? 'Plan' : 'Actual'; // Assuming Plan table has NPA columns? Code logic suggests yes.

    // Data Field Mappings
    // Determine which field to read from the 'data' object
    // Plan uses _plan suffix mostly, Achievement uses _actual or _acc
    const f = (fieldPlan, fieldAchieve) => isPlan ? fieldPlan : fieldAchieve;

    // Build Rows & Totals
    const totals = {
        ftod: 0, dpd130: 0, dpd3160: 0, dpd6190: 0,
        npaAct: 0, npaClose: 0,
        ns: 0,
        disbIglAcc: 0, disbIglAmt: 0, disbFigAcc: 0, disbFigAmt: 0, disbIlAcc: 0, disbIlAmt: 0,
        kycIgl: 0, kycFig: 0, kycIl: 0
    };

    let bodyRows = '';
    const fmt = n => n === 0 ? '-' : n.toLocaleString('en-IN');
    let rowIndex = 0;

    rows.forEach(row => {
        const d = row.data || {};

        // Alternating row color for better readability
        const rowBgColor = rowIndex % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
        rowIndex++;

        // Extract values using dynamic mapping
        const ftod = parseInt(d[f('ftod_plan', 'ftod_actual')]) || 0;
        const dpd130 = parseInt(d[f('dpd_1_30_plan', 'dpd_1_30_actual')]) || 0;
        const dpd3160 = parseInt(d[f('dpd_31_60_plan', 'dpd_31_60_actual')]) || 0;
        const dpd6190 = parseInt(d[f('dpd_61_90_plan', 'dpd_61_90_actual')]) || 0;

        const npaAct = parseInt(d['npa_activation']) || 0;
        const npaClose = parseInt(d['npa_closure']) || 0;

        const ns = parseInt(d[f('fy_non_start_plan', 'fy_non_start_acc')]) || 0;

        const disbIglAcc = parseInt(d['disb_igl_acc']) || 0;
        const disbIglAmt = parseInt(d['disb_igl_amt']) || 0;
        const disbFigAcc = parseInt(d['disb_fig_acc']) || 0;
        const disbFigAmt = parseInt(d['disb_fig_amt']) || 0;
        const disbIlAcc = parseInt(d['disb_il_acc']) || 0;
        const disbIlAmt = parseInt(d['disb_il_amt']) || 0;

        const kycIgl = parseInt(d['kyc_igl']) || 0;
        const kycFig = parseInt(d['kyc_fig']) || 0;
        const kycIl = parseInt(d['kyc_il']) || 0;

        // Sum Totals
        totals.ftod += ftod; totals.dpd130 += dpd130; totals.dpd3160 += dpd3160; totals.dpd6190 += dpd6190;
        totals.npaAct += npaAct; totals.npaClose += npaClose;
        totals.ns += ns;
        totals.disbIglAcc += disbIglAcc; totals.disbIglAmt += disbIglAmt;
        totals.disbFigAcc += disbFigAcc; totals.disbFigAmt += disbFigAmt;
        totals.disbIlAcc += disbIlAcc; totals.disbIlAmt += disbIlAmt;
        totals.kycIgl += kycIgl; totals.kycFig += kycFig; totals.kycIl += kycIl;

        bodyRows += `
            <tr style="font-size: 10px; background: ${rowBgColor};">
                <td style="background: ${rowBgColor}; border: 1px solid #000; padding: 3px 6px; text-align: left; font-weight: 600;">${row.name}</td>
                <td style="background: ${colors.ftod}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(ftod)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(dpd130)}</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(dpd3160)}</td>
                <td style="background: ${colors.pnpa}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(dpd6190)}</td>
                <td style="background: ${colors.npa}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(npaAct)}</td>
                <td style="background: ${colors.npa}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(npaClose)}</td>
                <td style="background: ${colors.fy2526}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(ns)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(disbIglAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(disbIglAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(disbFigAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(disbFigAmt)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(disbIlAcc)}</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(disbIlAmt)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(kycIgl)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(kycFig)}</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(kycIl)}</td>
            </tr>
        `;
    });

    // Grand Total Row
    const totalRow = `
        <tr style="font-weight: bold; font-size: 10px; background: #FFFF00;">
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: left;">Grand Total</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.ftod)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.dpd130)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.dpd3160)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.dpd6190)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.npaAct)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.npaClose)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.ns)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.disbIglAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(totals.disbIglAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.disbFigAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(totals.disbFigAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.disbIlAcc)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: right;">${fmt(totals.disbIlAmt)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.kycIgl)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.kycFig)}</td>
            <td style="background: #FFFF00; border: 1px solid #000; padding: 3px 6px; text-align: center;">${fmt(totals.kycIl)}</td>
        </tr>
    `;

    return `
        <table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; width: auto; background:white;">
            <!-- Title -->
            <tr>
                <td colspan="17" style="text-align: center; font-weight: bold; font-size: 14px; padding: 8px; background: ${colors.white}; border: 1px solid #000; color:black;">
                    ${title}
                </td>
            </tr>
            <!-- Header Row 1: Category Headers (merged cells) -->
            <tr style="font-weight: bold; font-size: 10px;">
                <td rowspan="2" style="background: ${colors.white}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">${firstColHeader}</td>
                <td style="background: ${colors.ftod}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">FTOD</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">1-30 DPD</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">31-60 DPD</td>
                <td style="background: ${colors.pnpa}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">61-90 DPD</td>
                <td colspan="2" style="background: ${colors.npa}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">NPA Accounts</td>
                <td style="background: ${colors.fy2526}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">FY NS</td>
                <td colspan="6" style="background: ${colors.disb}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">Disbursement</td>
                <td colspan="3" style="background: ${colors.kyc}; border: 1px solid #000; padding: 4px 8px; text-align: center; color:black;">KYC Sourcing</td>
            </tr>
            <!-- Header Row 2: Sub-column Headers (matching Set Target) -->
            <tr style="font-weight: bold; font-size: 9px;">
                <td style="background: ${colors.ftod}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">${suffix}</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">${suffix}</td>
                <td style="background: ${colors.slipped}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">${suffix}</td>
                <td style="background: ${colors.pnpa}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">${suffix}</td>
                <td style="background: ${colors.npa}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">Activation</td>
                <td style="background: ${colors.npa}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">Closure</td>
                <td style="background: ${colors.fy2526}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">NS Acc</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IGL Acc</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IGL Amt</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">FIG Acc</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">FIG Amt</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IL Acc</td>
                <td style="background: ${colors.disb}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IL Amt</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IGL</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">FIG</td>
                <td style="background: ${colors.kyc}; border: 1px solid #000; padding: 3px 6px; text-align: center; color:black;">IL</td>
            </tr>
            ${bodyRows}
            ${totalRow}
        </table>
    `;
}

// --- CLIPBOARD HELPERS (Mobile Safe) ---
// Comprehensive mobile-compatible clipboard implementation
// Fixes for: iOS Safari, Android Chrome, Desktop browsers

function isMobileDevice() {
    return window.innerWidth < 768 || /Mobi|Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
}

function isIOS() {
    return /iPhone|iPad|iPod/i.test(navigator.userAgent) ||
        (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}

function supportsAsyncClipboard() {
    return !!(navigator.clipboard && typeof navigator.clipboard.write === 'function');
}

function supportsClipboardItems() {
    return typeof ClipboardItem !== 'undefined';
}

async function getClipboardPermission() {
    if (!navigator.permissions || !navigator.permissions.query) return 'granted'; // Assume granted if can't check
    try {
        const status = await navigator.permissions.query({ name: 'clipboard-write' });
        return status.state || 'granted';
    } catch {
        return 'granted'; // Permissions API may not support clipboard-write query
    }
}

// Robust legacy copy using multiple fallback methods
function legacyCopyText(text) {
    // Method 1: Textarea approach (works on most browsers)
    try {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.setAttribute('readonly', '');
        // Critical: Position must be visible for iOS but off-screen for others
        textarea.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 2em;
            height: 2em;
            padding: 0;
            border: none;
            outline: none;
            box-shadow: none;
            background: transparent;
            font-size: 16px;
            opacity: 0;
            z-index: 999999;
        `;

        document.body.appendChild(textarea);

        if (isIOS()) {
            // iOS specific: Need to use setSelectionRange and contentEditable trick
            textarea.contentEditable = true;
            textarea.readOnly = false;

            const range = document.createRange();
            range.selectNodeContents(textarea);

            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);

            textarea.setSelectionRange(0, 999999);
        } else {
            textarea.select();
            textarea.setSelectionRange(0, textarea.value.length);
        }

        const success = document.execCommand('copy');
        document.body.removeChild(textarea);

        if (success) return true;
    } catch (e) {
        console.warn('Textarea copy failed:', e);
    }

    // Method 2: ContentEditable div approach (iOS fallback)
    try {
        const div = document.createElement('div');
        div.contentEditable = true;
        div.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100px;
            height: 100px;
            opacity: 0;
            z-index: 999999;
            white-space: pre-wrap;
        `;
        div.textContent = text;
        document.body.appendChild(div);

        div.focus();

        const range = document.createRange();
        range.selectNodeContents(div);

        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        const success = document.execCommand('copy');
        document.body.removeChild(div);

        if (success) return true;
    } catch (e) {
        console.warn('ContentEditable copy failed:', e);
    }

    return false;
}

// Helper: Extract plain text from HTML table
function extractTableText(tableHTML) {
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = tableHTML;
    return tempDiv.innerText || tempDiv.textContent || '';
}

// Helper: Generate Blob from Table HTML
async function generateTableImageBlob(tableHTML) {
    const container = document.createElement('div');
    container.style.cssText = `
        position: fixed;
        left: -9999px;
        top: 0;
        background: white;
        padding: 20px;
        z-index: 9999;
        font-family: sans-serif;
    `;

    const styledTableHTML = `
        <style>
            table { border-collapse: collapse; width: 100%; color: #000; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f0f0f0; font-weight: bold; }
        </style>
        ${tableHTML}
    `;

    container.innerHTML = `<div style="display:inline-block; width:max-content; background:white;">${styledTableHTML}</div>`;
    document.body.appendChild(container);

    try {
        await new Promise(resolve => setTimeout(resolve, 150));
        const contentEl = container.firstElementChild;
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
        const isMobile = isMobileDevice();
        const scale = isMobile ? 1.5 : 2;
        const canvas = await html2canvas(contentEl, {
            backgroundColor: '#ffffff',
            scale: scale,
            logging: false,
            useCORS: true,
            allowTaint: true
        });
        document.body.removeChild(container);
        return new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    } catch (e) {
        console.error("Image generation error:", e);
        if (document.body.contains(container)) {
            document.body.removeChild(container);
        }
        return null;
    }
}

// Function to Share Report Image using Web Share API
async function shareReportImage(tableHTML) {
    if (!navigator.canShare) {
        showToast("Sharing not supported on this browser", "alert");
        return;
    }
    // Show loading toast
    showToast("Generating image for sharing...", "info");

    try {
        const blob = await generateTableImageBlob(tableHTML);
        if (!blob) throw new Error("Failed to generate image");

        const file = new File([blob], "daily_report.png", { type: "image/png" });
        const shareData = {
            files: [file],
            title: 'Daily Report',
            text: 'Here is the daily report.'
        };

        if (navigator.canShare(shareData)) {
            await navigator.share(shareData);
            showToast("Shared successfully! 🚀", "success");
        } else {
            showToast("Device doesn't support image sharing", "alert");
        }
    } catch (err) {
        console.warn("Share failed:", err);
        if (err.name !== 'AbortError') {
            showToast("Share failed. Try Copy instead.", "alert");
        }
    }
}

// Main copy function - tries image copy on ALL devices (mobile + desktop)
async function copyTableImageToClipboard(tableHTML) {
    const isMobile = isMobileDevice();
    const canWriteImages = supportsAsyncClipboard() && supportsClipboardItems();

    // If clipboard API not available, go to fallback
    if (!canWriteImages) {
        return await copyTableAsTextFallback(tableHTML);
    }

    const blob = await generateTableImageBlob(tableHTML);

    if (!blob) {
        console.warn("Canvas blob generation failed, trying HTML fallback...");
        return await copyTableAsTextFallback(tableHTML);
    }

    try {
        const item = new ClipboardItem({ 'image/png': blob });
        await navigator.clipboard.write([item]);
        showToast("Report image copied! 📷", "success");
        return true;
    } catch (err) {
        // Mobile browsers block image clipboard write - download the image instead
        console.warn("Image copy failed, downloading image...", err);

        if (isMobileDevice()) {
            // Download image for mobile
            try {
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'report_' + Date.now() + '.png';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
                showToast("Image downloaded! Share from Photos 📱", "success");
                return true;
            } catch (downloadErr) {
                console.warn("Download failed too:", downloadErr);
                return await copyTableAsTextFallback(tableHTML);
            }
        } else {
            return await copyTableAsTextFallback(tableHTML);
        }
    }
}

// Fallback: Try HTML copy first (preserves table formatting), then plain text
async function copyTableAsTextFallback(tableHTML) {
    const canWriteHtml = supportsAsyncClipboard() && supportsClipboardItems();

    // TIER 1: Try copying as HTML (preserves table formatting when pasted)
    if (canWriteHtml && window.isSecureContext) {
        try {
            const htmlBlob = new Blob([tableHTML], { type: 'text/html' });
            const textBlob = new Blob([extractTableText(tableHTML)], { type: 'text/plain' });

            // Write both HTML and plain text - apps will use the best format they support
            const item = new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            });
            await navigator.clipboard.write([item]);
            showToast("Table copied with formatting! 📋", "success");
            return true;
        } catch (err) {
            console.warn("HTML clipboard write failed:", err.message);
            // Continue to text fallback
        }
    }

    // Extract plain text from HTML
    const text = extractTableText(tableHTML);

    // TIER 1: Try modern Clipboard API (works best on desktop and newer mobile)
    if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
        try {
            await navigator.clipboard.writeText(text);
            showToast("Table copied! 📋", "success");
            return true;
        } catch (err) {
            console.warn("Clipboard API writeText failed:", err.message);
            // Continue to fallbacks - don't return false yet
        }
    }

    // TIER 2: Legacy execCommand copy (sync, works within user gesture)
    const legacySuccess = legacyCopyText(text);
    if (legacySuccess) {
        showToast("Table copied! 📋", "success");
        return true;
    }

    // TIER 3: Prompt user to manually copy
    // This is a last resort for very restricted browsers
    try {
        // Create a visible prompt with the text for manual copy
        const modal = document.createElement('div');
        modal.style.cssText = `
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0,0,0,0.8);
            z-index: 999999;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        `;
        modal.innerHTML = `
            <div style="background: white; border-radius: 12px; padding: 24px; max-width: 90%; max-height: 80%; overflow: auto;">
                <h3 style="margin: 0 0 16px 0; color: #333;">Copy Table Data</h3>
                <p style="color: #666; margin-bottom: 16px;">Automatic copy failed. Please select all text below and copy manually:</p>
                <textarea id="manualCopyArea" style="width: 100%; height: 200px; padding: 12px; border: 1px solid #ddd; border-radius: 8px; font-family: monospace; font-size: 12px;" readonly>${text}</textarea>
                <div style="display: flex; gap: 12px; margin-top: 16px;">
                    <button id="selectAllBtn" style="flex: 1; padding: 12px; background: #4F46E5; color: white; border: none; border-radius: 8px; cursor: pointer; font-weight: 600;">Select All</button>
                    <button id="closeManualCopy" style="flex: 1; padding: 12px; background: #e5e5e5; color: #333; border: none; border-radius: 8px; cursor: pointer; font-weight: 600;">Close</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);

        const textarea = document.getElementById('manualCopyArea');
        document.getElementById('selectAllBtn').onclick = () => {
            textarea.select();
            textarea.setSelectionRange(0, textarea.value.length);
        };
        document.getElementById('closeManualCopy').onclick = () => {
            document.body.removeChild(modal);
        };
        modal.onclick = (e) => {
            if (e.target === modal) document.body.removeChild(modal);
        };

        // Auto-select the text
        setTimeout(() => {
            textarea.focus();
            textarea.select();
        }, 100);

        showToast("Please copy manually from the dialog", "alert");
        return false;
    } catch (e) {
        console.error("Manual copy fallback failed:", e);
        showToast("Copy failed. Please take a screenshot.", "alert");
        return false;
    }
}

// --- DOWNLOAD ACHIEVEMENT PLAN REPORT (With Comparison & Colors) ---
// --- DOWNLOAD ACHIEVEMENT PLAN REPORT (With Comparison & Colors) ---
async function downloadAchievementPlanReportDirectly() {
    // 1. Get Range from State
    const fromDate = state.dateFrom;
    const toDate = state.dateTo;

    // 2. Validate
    if (!fromDate || !toDate) {
        showToast('⚠️ No date range selected', true);
        return;
    }

    // 3. Confirm with user via Toast or just proceed?
    // Requirement says: "Query Neon for ALL rows between startDate and endDate"
    // "Add loading state"

    const label = formatRangeLabel(fromDate, toDate);
    setLoading(true, `Generating Report for ${label}...`);

    try {
        // Fetch aggregated data for report
        const reportData = await fetchAndAggregateData(fromDate, toDate);

        // Generate Check
        const count = Object.keys(reportData).length;
        if (count === 0) {
            showToast('ℹ️ No data found for this range', false);
            setLoading(false);
            return;
        }

        // Apply aggregation if needed (using temporary state logic)
        // Since getReportRows uses state.branchDetails and state.rawData,
        // we need to temporarily inject the fetched range data into branchDetails to use getReportRows logic
        // OR refactor getReportRows. Refactoring is cleaner but passing data source is safer.

        // Let's manually reuse getReportRows logic here for the range data
        // BUT fetchAndAggregateData returns a MAP branch->details.
        // getReportRows expects state.branchDetails to be that map.
        // So we can mock state.branchDetails for getReportRows if we want to reuse it exactly.

        const originalDetails = state.branchDetails;
        state.branchDetails = reportData; // temporarily swap for aggregation

        let reportRows = [];
        try {
            reportRows = getReportRows(state.reportLevel);
        } finally {
            state.branchDetails = originalDetails; // restore
        }

        await generateAndDownloadReport(reportRows, fromDate, toDate);

        setLoading(false);
        showToast(`✅ Report Downloaded: ${label}`);
    } catch (error) {
        console.error("Report Generation Error:", error);
        setLoading(false);
        showToast('❌ Failed to generate report', true);
    }
}





async function generateAndDownloadReport(reportRows, fromDate, toDate) {
    // Legacy support or remove if unused, but updating to use new function might be better if needed.
    // For now, keeping as is but be aware it calls missing function 'convertTableToPNG' if strict.
    // Since we removed convertTableToPNG, we should update this or remove usage.
    // Based on user request, this flow might not be needed or should use copy too.
    // Let's stub it or map to copy for now to prevent errors.
    console.warn("Direct download deprecated in favor of Copy.");
}



// --- NEW RENDERERS ---

function renderCompletionAlert(stats, mode) {
    const total = stats.totalBranches;
    let completed = 0;
    let incomplete = 0;
    let label = '';
    let alertColor = 'var(--primary)';

    if (mode === 'PLAN') {
        completed = total - stats.noDataBranches; // Using logic where 'noData' means no plan in this context roughly? 
        // Wait, logic in calculateCEOStats: 
        // if (hasPlan) regionStats[region].plansSet++;
        // if (hasPlan && !hasAchieve) onlyPlans++;
        // if (!hasPlan && !hasAchieve) noDataBranches++;
        // So for Plan Mode, "Completed" = Total - branches_with_no_plan.
        // Actually stats.plansRate is based on (total - noDataBranches) / total roughly?
        // Let's look at calculateCEOStats again.
        // plansRate = Math.round(((totalBranches - noDataBranches) / totalBranches) * 100).
        // BUT wait, noDataBranches is incremented if (!hasPlan && !hasAchieve).
        // If I have Plan but no Achieve, it's NOT noDataBranches.
        // So "Branches with Plan" = Total - "Branches with neither".
        // Let's rely on stats.plansRate logic for count.
        completed = stats.onlyPlans + stats.completedBranches; // This covers all who have plans (whether they have achievement or not)
        // Wait, stats.completedBranches is (hasPlan && hasAchieve).
        // stats.onlyPlans is (hasPlan && !hasAchieve).
        // So yes, sum is all with Plans.

        incomplete = total - completed;
        label = 'Target Planning Status';
        if (incomplete > 0) alertColor = '#F59E0B'; // Orange if pending
    } else {
        // REVIEW MODE
        completed = stats.completedBranches;
        incomplete = total - completed;
        label = 'Daily Reporting Status';
        if (incomplete > 0) alertColor = '#EF4444'; // Red if pending
    }

    // Don't show if all good? User specifically asked "show how many completed and how many incomplete". 
    // So always show it.

    return `
        <div class="alert-banner" onclick="openDetailModal('completion')" style="grid-column: 1 / -1; background: ${alertColor}15; border: 1px solid ${alertColor}; border-radius: 12px; padding: 12px 20px; display: flex; align-items: center; justify-content: space-between; cursor: pointer; margin-bottom: 8px;">
            <div style="display: flex; align-items: center; gap: 12px;">
                 <div style="background: ${alertColor}; color: white; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center;">
                    <svg class="icon" viewBox="0 0 24 24" style="width: 18px; height: 18px;"><polyline points="20 6 9 17 4 12"></polyline></svg>
                 </div>
                 <div>
                    <div style="font-weight: 700; color: var(--text-primary); font-size: 15px;">${label}</div>
                    <div style="font-size: 13px; color: var(--text-secondary);">
                        <span style="font-weight: 600; color: var(--success);">${completed} Completed</span> 
                        <span style="margin: 0 6px;">•</span>
                        <span style="font-weight: 600; color: ${incomplete > 0 ? '#EF4444' : 'var(--text-secondary)'};">${incomplete} Pending</span>
                    </div>
                 </div>
            </div>
            <div style="display: flex; align-items: center; gap: 8px;">
                <span style="font-size: 12px; font-weight: 600; color: ${alertColor};">View Details</span>
                <svg class="icon" viewBox="0 0 24 24" style="width: 16px; height: 16px; color: ${alertColor};"><polyline points="9 18 15 12 9 6"></polyline></svg>
            </div>
        </div>
    `;
}

function renderCEOPlanDashboard(stats, buffer) {
    // ALERT BANNER
    buffer.innerHTML += renderCompletionAlert(stats, 'PLAN');

    // ROW 1: PLAN METRICS
    buffer.innerHTML += `
        <div class="metric-card">
            <div class="metric-header">
                <div class="metric-icon" style="background:var(--primary-light); color:var(--primary-accent);">
                    <svg class="icon" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                </div>
                <div class="metric-badge ${stats.plansRate >= 100 ? 'badge-up' : ''}" style="${stats.plansRate < 100 ? 'background:#FEF3C7; color:#D97706;' : ''}">
                    ${stats.plansRate}% Set
                </div>
            </div>
            <div>
                <div class="metric-title">Plan Compliance</div>
                <div class="metric-value">${stats.plansRate}%</div>
                <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">${stats.totalBranches - stats.noDataBranches} of ${stats.totalBranches} Branches</div>
            </div>
        </div>

        <div class="metric-card">
            <div class="metric-header">
                <div class="metric-icon" style="background:#EEF2FF; color:#6366F1;">
                    <svg class="icon" viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path></svg>
                </div>
            </div>
            <div>
                <div class="metric-title">Total Collection Target</div>
                <div class="metric-value">${stats.totalCollectionPlan.toLocaleString()}</div>
                <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Accounts (FTOD + 1-30 + 31-60 + 61-90 DPD)</div>
            </div>
        </div>

        <div class="metric-card">
            <div class="metric-header">
                <div class="metric-icon" style="background:#D1FAE5; color:#059669;">
                    <svg class="icon" viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" rx="2"></rect><line x1="2" y1="10" x2="22" y2="10"></line></svg>
                </div>
            </div>
            <div>
                <div class="metric-title">Disbursement Target</div>
                <div class="metric-value">₹${(stats.totalDisbursementPlan / 10000000).toFixed(2)} Cr</div>
                <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Total Amount Plan</div>
            </div>
        </div>
        
        <div class="metric-card clickable" onclick="showLastMonthSummary()" style="cursor:pointer; transition: transform 0.2s;">
            <div class="metric-header">
                <div class="metric-icon" style="background:#F3E8FF; color:#9333EA;">
                    <svg class="icon" viewBox="0 0 24 24">
                        <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
                        <line x1="16" y1="2" x2="16" y2="6"></line>
                        <line x1="8" y1="2" x2="8" y2="6"></line>
                        <line x1="3" y1="10" x2="21" y2="10"></line>
                    </svg>
                </div>
            </div>
            <div>
                <div class="metric-title">Last Month Summary</div>
                <div class="metric-value" style="font-size: 16px; margin-top: 4px;">View Report →</div>
                <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Performance Analysis</div>
            </div>
        </div>
    `;

    // ROW 2: DETAILED BREAKDOWN
    buffer.innerHTML += `
        <div class="chart-card grid-full">
            <div class="chart-header">
                <div class="chart-title">🎯 Collection Targets Breakdown</div>
            </div>
             <div style="display:flex; flex-direction:column; gap:16px;">
                ${renderPlanBar('FTOD Accounts', stats.ftodPlan, stats.totalCollectionPlan, '#6366F1')}
                ${renderPlanBar('1-30 DPD', stats.dpd130Plan, stats.totalCollectionPlan, '#F59E0B')}
                ${renderPlanBar('31-60 DPD', stats.dpd3160Plan, stats.totalCollectionPlan, '#FB923C')}
                ${renderPlanBar('61-90 DPD', stats.dpd6190Plan, stats.totalCollectionPlan, '#EF4444')}
                ${renderPlanBar('NPA Accounts', stats.npaTotalPlan, stats.totalCollectionPlan, '#EC4899')}
            </div>
        </div>

        <div class="chart-card grid-half">
            <div class="chart-header">
                <div class="chart-title">👤 KYC Sourcing Plan</div>
            </div>
             <div style="display:flex; flex-direction:column; gap:16px;">
                ${renderPlanBar('IGL', stats.kycIglPlan, stats.kycTotal, '#F59E0B')}
                ${renderPlanBar('FIG', stats.kycFigPlan, stats.kycTotal, '#6366F1')}
                ${renderPlanBar('IL', stats.kycIlPlan, stats.kycTotal, '#3B82F6')}
            </div>
        </div>

        <div class="chart-card grid-half">
            <div class="chart-header">
                <div class="chart-title">💰 Disbursement Plan Breakdown</div>
            </div>
            
            <div style="display:flex; flex-direction:column; gap:20px;">
                <!-- By Amount -->
                <div>
                    <div style="font-size:12px; font-weight:600; color:var(--text-secondary); margin-bottom:12px; text-transform:uppercase; letter-spacing:0.5px;">By Amount (₹ ${(stats.totalDisbursementPlan / 10000000).toFixed(2)} Cr)</div>
                    <div style="display:flex; flex-direction:column; gap:12px;">
                        ${renderPlanBar('IGL Amount', stats.disbIglAmtPlan, stats.totalDisbursementPlan, '#10B981', true)}
                        ${renderPlanBar('FIG Amount', stats.disbFigAmtPlan, stats.totalDisbursementPlan, '#6366F1', true)}
                        ${renderPlanBar('IL Amount', stats.disbIlAmtPlan, stats.totalDisbursementPlan, '#8B5CF6', true)}
                    </div>
                </div>

                <!-- By Accounts -->
                <div>
                    <div style="font-size:12px; font-weight:600; color:var(--text-secondary); margin-bottom:12px; text-transform:uppercase; letter-spacing:0.5px; padding-top:12px; border-top:1px dashed var(--border-color);">By Accounts (${stats.totalDisbAccPlan.toLocaleString()})</div>
                    <div style="display:flex; flex-direction:column; gap:12px;">
                        ${renderPlanBar('IGL Accounts', stats.disbIglAccPlan, stats.totalDisbAccPlan, '#34D399')}
                        ${renderPlanBar('FIG Accounts', stats.disbFigAccPlan, stats.totalDisbAccPlan, '#818CF8')}
                        ${renderPlanBar('IL Accounts', stats.disbIlAccPlan, stats.totalDisbAccPlan, '#A78BFA')}
                    </div>
                </div>
            </div>
        </div>
    `;

    // ROW 3: REGIONAL TABLE
    buffer.innerHTML += `
        <div class="table-card grid-full">
            <div class="chart-header">
                <div class="chart-title">🌍 Regional Plan Status</div>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>Region</th>
                        <th>Total Branches</th>
                        <th>Plans Set</th>
                        <th>Compliance %</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody id="regionalPlanTableHook"></tbody>
            </table>
        </div>
    `;
}

function renderCEOReviewDashboard(stats, buffer) {
    // ALERT BANNER
    buffer.innerHTML += renderCompletionAlert(stats, 'REVIEW');

    // ============ ROW 1: EXECUTIVE SUMMARY (4 Cards) ============
    buffer.innerHTML += `
                <div class="metric-card clickable-metric" onclick="openDetailModal('completion')">
                    <div class="metric-header">
                        <div class="metric-icon" style="background: linear-gradient(135deg, #10B981 0%, #059669 100%); color: white;">
                            <svg class="icon" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                        </div>
                        <div class="metric-badge ${stats.completionRate >= 80 ? 'badge-up' : ''}" style="${stats.completionRate < 80 ? 'background: #FEF3C7; color: #D97706;' : ''}">
                            ${stats.completionRate >= 80 ? '✓ On Track' : '⚠ Pending'}
                        </div>
                    </div>
                    <div>
                        <div class="metric-title">Report Completion</div>
                        <div class="metric-value">${stats.completionRate}%</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">${stats.completedBranches} of ${stats.totalBranches} branches</div>
                    </div>
                </div>

                <div class="metric-card clickable-metric" onclick="openDetailModal('branches')">
                    <div class="metric-header">
                        <div class="metric-icon" style="background: linear-gradient(135deg, #6366F1 0%, #4F46E5 100%); color: white;">
                            <svg class="icon" viewBox="0 0 24 24"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path></svg>
                        </div>
                    </div>
                    <div>
                        <div class="metric-title">Active Branches</div>
                        <div class="metric-value">${stats.totalBranches}</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">${Object.keys(stats.regionStats).length} Regions</div>
                    </div>
                </div>

                <div class="metric-card clickable-metric" onclick="openDetailModal('achievement')">
                    <div class="metric-header">
                        <div class="metric-icon" style="background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%); color: white;">
                            <svg class="icon" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline></svg>
                        </div>
                        <div class="metric-badge badge-up">
                            <svg class="icon" style="width:12px;height:12px;" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline></svg>
                            ${stats.avgAchievementPct}% Avg
                        </div>
                    </div>
                    <div>
                        <div class="metric-title">Overall Achievement</div>
                        <div class="metric-value">${stats.avgAchievementPct}%</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Across all metrics</div>
                    </div>
                </div>

                <div class="metric-card clickable-metric" onclick="openDetailModal('collection')">
                    <div class="metric-header">
                        <div class="metric-icon" style="background: linear-gradient(135deg, #8B5CF6 0%, #7C3AED 100%); color: white;">
                            <svg class="icon" viewBox="0 0 24 24"><line x1="12" y1="1" x2="12" y2="23"></line><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"></path></svg>
                        </div>
                    </div>
                    <div>
                        <div class="metric-title">Collection Achievement</div>
                        <div class="metric-value">₹${((stats.ftodAchieve + stats.dpd130Achieve + stats.dpd3160Achieve + stats.dpd6190Achieve) / 100000).toFixed(1)}L</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">vs Plan: ₹${(stats.totalCollectionPlan / 100000).toFixed(1)}L</div>
                    </div>
                </div>
            `;

    // ============ ROW 2: COLLECTION PERFORMANCE ============
    buffer.innerHTML += `
                <div class="chart-card grid-full">
                    <div class="chart-header">
                        <div class="chart-title">📊 Collection Performance</div>
                        <div style="font-size:11px; color:var(--text-secondary);">Plan vs Achievement</div>
                    </div>
                    <div class="collection-grid">
                        <div class="clickable-section" onclick="openDetailModal('ftod')">${renderCollectionMetric('FTOD', stats.ftodPlan, stats.ftodAchieve)}</div>
                        <div class="clickable-section" onclick="openDetailModal('dpd_1_30')">${renderCollectionMetric('1-30 DPD', stats.dpd130Plan, stats.dpd130Achieve)}</div>
                        <div class="clickable-section" onclick="openDetailModal('dpd_31_60')">${renderCollectionMetric('31-60 DPD', stats.dpd3160Plan, stats.dpd3160Achieve)}</div>
                        <div class="clickable-section" onclick="openDetailModal('dpd_61_90')">${renderCollectionMetric('61-90 DPD', stats.dpd6190Plan, stats.dpd6190Achieve)}</div>
                        <div class="clickable-section" onclick="openDetailModal('npa')">${renderNPAMovement(stats.npaActivation, stats.npaClosure)}</div>
                    </div>
                </div>
            `;

    // ============ ROW 3: PORTFOLIO HEALTH + DISBURSEMENT ============
    buffer.innerHTML += `
                <div class="chart-card grid-half clickable-section" onclick="openDetailModal('portfolio')">
                    <div class="chart-header">
                        <div class="chart-title">🏦 Portfolio Health</div>
                        <div style="font-size:11px; color:var(--text-secondary);">Click for details</div>
                    </div>
                    <div class="bar-chart-container" style="align-items:center; justify-content:center; flex-direction:row;" id="portfolioDonutHook"></div>
                </div>

                <div class="chart-card grid-half">
                    <div class="chart-header">
                        <div class="chart-title">💰 Disbursement Summary</div>
                        <div style="font-size:18px; font-weight:700; color:var(--success);">₹${(stats.totalDisbursementAchieve / 10000000).toFixed(2)} Cr</div>
                    </div>
                    <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:16px; padding:16px 0;">

                        <div class="clickable-section" onclick="openDetailModal('disb_igl')" style="background:var(--bg-body); padding:16px; border-radius:12px; text-align:center; cursor:pointer;">
                            <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px;">IGL</div>
                            <div style="font-size:20px; font-weight:700; color:#10B981;">${stats.disbIglAccAchieve}</div>
                            <div style="font-size:11px; color:var(--text-secondary);">₹${(stats.disbIglAmtAchieve / 100000).toFixed(1)}L</div>
                             <div style="font-size:10px; color:var(--text-secondary); margin-top:4px;">Plan: ₹${(stats.disbIglAmtPlan / 100000).toFixed(1)}L</div>
                        </div>
                        <div class="clickable-section" onclick="openDetailModal('disb_fig')" style="background:var(--bg-body); padding:16px; border-radius:12px; text-align:center; cursor:pointer;">
                            <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px;">FIG</div>
                            <div style="font-size:20px; font-weight:700; color:#6366F1;">${stats.disbFigAccAchieve}</div>
                            <div style="font-size:11px; color:var(--text-secondary);">₹${(stats.disbFigAmtAchieve / 100000).toFixed(1)}L</div>
                             <div style="font-size:10px; color:var(--text-secondary); margin-top:4px;">Plan: ₹${(stats.disbFigAmtPlan / 100000).toFixed(1)}L</div>
                        </div>
                        <div class="clickable-section" onclick="openDetailModal('disb_il')" style="background:var(--bg-body); padding:16px; border-radius:12px; text-align:center; cursor:pointer;">
                            <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px;">IL</div>
                            <div style="font-size:20px; font-weight:700; color:#F59E0B;">${stats.disbIlAccAchieve}</div>
                            <div style="font-size:11px; color:var(--text-secondary);">₹${(stats.disbIlAmtAchieve / 100000).toFixed(1)}L</div>
                             <div style="font-size:10px; color:var(--text-secondary); margin-top:4px;">Plan: ₹${(stats.disbIlAmtPlan / 100000).toFixed(1)}L</div>
                        </div>
                        <div class="clickable-section" onclick="openDetailModal('disbursement')" style="background:var(--primary-light); padding:16px; border-radius:12px; text-align:center; cursor:pointer; grid-column: span 3;">
                            <div style="font-size:11px; color:var(--primary-accent); margin-bottom:4px;">Total Accounts</div>
                            <div style="font-size:20px; font-weight:700; color:var(--primary-accent);">${stats.disbIglAccAchieve + stats.disbFigAccAchieve + stats.disbIlAccAchieve}</div>
                            <div style="font-size:11px; color:var(--text-secondary);">Plan: ${stats.disbIglAccPlan + stats.disbFigAccPlan + stats.disbIlAccPlan}</div>
                        </div>
                    </div>
                </div>
            `;

    // ============ ROW 4: REGIONAL PERFORMANCE ============
    buffer.innerHTML += `
                <div class="table-card grid-full">
                    <div class="chart-header">
                        <div class="chart-title">🌍 Regional Performance</div>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>Region</th>
                                <th>Branches</th>
                                <th>Completed</th>
                                <th>Avg Achievement</th>
                                <th>Status</th>
                            </tr>
                        </thead>
                        <tbody id="regionalTableHook"></tbody>
                    </table>
                </div>
            `;

    // ============ ROW 5: QUICK INSIGHTS (Review Only) ============
    buffer.innerHTML += `
                <div class="chart-card grid-half">
                    <div class="chart-header">
                        <div class="chart-title">⚡ Pending Actions</div>
                    </div>
                    <div style="display:flex; flex-direction:column; gap:12px; padding:12px 0;">
                        <div class="clickable-row" onclick="openDetailModal('pending')" style="display:flex; justify-content:space-between; align-items:center; padding:12px 16px; background:${stats.onlyPlans > 0 ? '#FEF3C7' : 'var(--bg-body)'}; border-radius:10px; cursor:pointer;">
                            <div style="display:flex; align-items:center; gap:10px;">
                                <span style="font-size:18px;">🎯</span>
                                <span style="font-weight:500;">Branches with only Plans</span>
                            </div>
                            <span style="font-size:18px; font-weight:700; color:${stats.onlyPlans > 0 ? '#D97706' : 'var(--success)'};">${stats.onlyPlans} →</span>
                        </div>
                        <div class="clickable-row" onclick="openDetailModal('nodata')" style="display:flex; justify-content:space-between; align-items:center; padding:12px 16px; background:${stats.noDataBranches > 0 ? '#FEE2E2' : 'var(--bg-body)'}; border-radius:10px; cursor:pointer;">
                            <div style="display:flex; align-items:center; gap:10px;">
                                <span style="font-size:18px;">⏳</span>
                                <span style="font-weight:500;">Branches with No Data</span>
                            </div>
                            <span style="font-size:18px; font-weight:700; color:${stats.noDataBranches > 0 ? '#DC2626' : 'var(--success)'};">${stats.noDataBranches} →</span>
                        </div>
                    </div>
                </div>

                <div class="chart-card grid-half clickable-section" onclick="openDetailModal('kyc')">
                    <div class="chart-header">
                        <div class="chart-title">📋 KYC Sourcing Summary</div>
                        <div style="font-size:14px; font-weight:700; color:var(--primary-accent);">${stats.kycTotal} Total · Click for details</div>
                    </div>
                    <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:12px; padding:16px 0;">
                        <div style="background:linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%); padding:20px 16px; border-radius:12px; text-align:center;">
                            <div style="font-size:24px; font-weight:700; color:#6366F1;">${stats.kycIgl}</div>
                            <div style="font-size:11px; color:#4F46E5; margin-top:4px;">IGL</div>
                        </div>
                        <div style="background:linear-gradient(135deg, #D1FAE5 0%, #A7F3D0 100%); padding:20px 16px; border-radius:12px; text-align:center;">
                            <div style="font-size:24px; font-weight:700; color:#059669;">${stats.kycFig}</div>
                            <div style="font-size:11px; color:#047857; margin-top:4px;">FIG</div>
                        </div>
                        <div style="background:linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%); padding:20px 16px; border-radius:12px; text-align:center;">
                            <div style="font-size:24px; font-weight:700; color:#D97706;">${stats.kycIl}</div>
                            <div style="font-size:11px; color:#B45309; margin-top:4px;">IL</div>
                        </div>
                    </div>
                </div>
            `;
}

function renderPlanBar(label, value, total, color, isMoney = false) {
    const pct = total > 0 ? Math.round((value / total) * 100) : 0;
    const displayValue = isMoney ? `₹${(value / 100000).toFixed(1)}L` : value.toLocaleString();

    return `
        <div>
            <div style="display:flex; justify-content:space-between; font-size:12px; font-weight:500; color:var(--text-primary); margin-bottom:6px;">
                <span>${label}</span>
                <span>${isMoney ? displayValue.replace('₹', 'Amt ') : displayValue}</span>
            </div>
            <div style="height:8px; background:var(--bg-body); border-radius:4px; overflow:hidden;">
                <div style="height:100%; width:${pct}%; background:${color}; border-radius:4px;"></div>
            </div>
            <div style="text-align:right; font-size:10px; color:${color}; margin-top:2px; font-weight:600;">${pct}%</div>
        </div>
    `;
}

function renderRegionalPlanTable(data, tbody) {
    if (!tbody || !data) return;

    Object.keys(data).sort().forEach(region => {
        const r = data[region];
        const pct = Math.round((r.plansSet / r.branches) * 100) || 0;

        let statusColor = '#10B981';
        let statusText = 'Completed';
        if (pct < 100) { statusColor = '#F59E0B'; statusText = 'Pending'; }
        if (pct < 50) { statusColor = '#EF4444'; statusText = 'Lagging'; }

        tbody.innerHTML += `
            <tr class="clickable-row">
                <td><span style="font-weight:600;">${region}</span></td>
                <td>${r.branches}</td>
                <td>${r.plansSet}</td>
                <td>
                    <span style="font-weight:600; color:${statusColor};">${pct}%</span>
                </td>
                <td><span class="status-pill" style="background:${statusColor}20; color:${statusColor};">${statusText}</span></td>
            </tr>`;
    });
}

// --- HIERARCHY SUMMARY HELPERS ---

function getAllBranchesFromNode(node) {
    let branches = [];
    if (node instanceof Set) {
        node.forEach(b => branches.push(b));
    } else {
        Object.values(node).forEach(child => {
            branches = branches.concat(getAllBranchesFromNode(child));
        });
    }
    return branches;
}

function calculateAggregateStatsForBranches(branchNames) {
    let stats = {
        branches: 0,
        plansSet: 0,
        completed: 0,
        ftodPlan: 0, ftodAchieve: 0,
        dpd130Plan: 0, dpd130Achieve: 0,
        dpd3160Plan: 0, dpd3160Achieve: 0,
        dpd6190Plan: 0, dpd6190Achieve: 0,
        collPlan: 0, collAchieve: 0,
        disbPlan: 0, disbAchieve: 0,
        disbIglAmtPlan: 0, disbIglAmtAchieve: 0,
        disbFigAmtPlan: 0, disbFigAmtAchieve: 0,
        disbIlAmtPlan: 0, disbIlAmtAchieve: 0,
        kyc: 0
    };

    const safeInt = (v) => parseInt(v) || 0;
    const safeFloat = (v) => parseFloat(v) || 0;

    branchNames.forEach(br => {
        stats.branches++;
        const entry = state.branchDetails[br];
        if (!entry) return;

        const hasPlan = !!entry.target;
        const hasAchieve = !!entry.achievement;

        if (hasPlan) stats.plansSet++;
        if (hasPlan && hasAchieve) stats.completed++;

        // Plan Data
        if (hasPlan) {
            const t = entry.target;
            const ftod = safeInt(t.ftod_plan);
            const dpd130 = safeInt(t.dpd_1_30_plan);
            const dpd3160 = safeInt(t.dpd_31_60_plan);
            const dpd6190 = safeInt(t.dpd_61_90_plan);

            stats.ftodPlan += ftod;
            stats.dpd130Plan += dpd130;
            stats.dpd3160Plan += dpd3160;
            stats.dpd6190Plan += dpd6190;
            stats.collPlan += (ftod + dpd130 + dpd3160 + dpd6190);

            const iglAmt = safeFloat(t.disb_igl_amt);
            const figAmt = safeFloat(t.disb_fig_amt);
            const ilAmt = safeFloat(t.disb_il_amt);
            stats.disbIglAmtPlan += iglAmt;
            stats.disbFigAmtPlan += figAmt;
            stats.disbIlAmtPlan += ilAmt;
            stats.disbPlan += (iglAmt + figAmt + ilAmt);
        }

        // Achievement Data
        if (hasAchieve) {
            const a = entry.achievement;
            const ftod = safeInt(a.ftod_actual);
            const dpd130 = safeInt(a.dpd_1_30_actual);
            const dpd3160 = safeInt(a.dpd_31_60_actual);
            const dpd6190 = safeInt(a.dpd_61_90_actual);

            stats.ftodAchieve += ftod;
            stats.dpd130Achieve += dpd130;
            stats.dpd3160Achieve += dpd3160;
            stats.dpd6190Achieve += dpd6190;
            stats.collAchieve += (ftod + dpd130 + dpd3160 + dpd6190);

            const iglAmt = safeFloat(a.disb_igl_amt);
            const figAmt = safeFloat(a.disb_fig_amt);
            const ilAmt = safeFloat(a.disb_il_amt);
            stats.disbIglAmtAchieve += iglAmt;
            stats.disbFigAmtAchieve += figAmt;
            stats.disbIlAmtAchieve += ilAmt;
            stats.disbAchieve += (iglAmt + figAmt + ilAmt);

            stats.kyc += (safeInt(a.kyc_igl) + safeInt(a.kyc_fig) + safeInt(a.kyc_il));
        }
    });

    return stats;
}

// --- DM SUMMARY STATS ---
function calculateDMSummaryStats(branchNames) {
    const safeInt = (v) => parseInt(v) || 0;
    const safeFloat = (v) => parseFloat(v) || 0;

    let stats = {
        disbAcc: 0, disbAmt: 0,
        ftodPlan: 0, ftodAchieve: 0,
        dpd130Plan: 0, dpd130Achieve: 0,
        dpd3160Plan: 0, dpd3160Achieve: 0,
        dpd6190Plan: 0, dpd6190Achieve: 0,
        npaClosePlan: 0, npaCloseAchieve: 0
    };

    branchNames.forEach(br => {
        const entry = state.branchDetails[br];
        if (!entry) return;

        // TARGETS
        if (entry.target) {
            const t = entry.target;
            stats.ftodPlan += safeInt(t.ftod_plan);
            stats.dpd130Plan += safeInt(t.dpd_1_30_plan);
            stats.dpd3160Plan += safeInt(t.dpd_31_60_plan);
            stats.dpd6190Plan += safeInt(t.dpd_61_90_plan);
            stats.npaClosePlan += safeInt(t.npa_closure);
        }

        // ACTUALS
        if (entry.achievement) {
            const a = entry.achievement;
            stats.disbAcc += (safeInt(a.disb_igl_acc) + safeInt(a.disb_fig_acc) + safeInt(a.disb_il_acc));
            stats.disbAmt += (safeFloat(a.disb_igl_amt) + safeFloat(a.disb_fig_amt) + safeFloat(a.disb_il_amt));

            stats.ftodAchieve += safeInt(a.ftod_actual);
            stats.dpd130Achieve += safeInt(a.dpd_1_30_actual);
            stats.dpd3160Achieve += safeInt(a.dpd_31_60_actual);
            stats.dpd6190Achieve += safeInt(a.dpd_61_90_actual);
            stats.npaCloseAchieve += safeInt(a.npa_closure);
        }
    });

    return stats;
}

function renderDMSummaryCards(stats) {
    return ""; // Summary cards removed as per user request
}

function renderHierarchySummaryCard(stats, title) {
    const collPct = stats.collPlan > 0 ? Math.round((stats.collAchieve / stats.collPlan) * 100) : 0;
    const disbPct = stats.disbPlan > 0 ? Math.round((stats.disbAchieve / stats.disbPlan) * 100) : 0;
    const disbPlanCr = (stats.disbPlan / 10000000).toFixed(2);
    const disbAchieveCr = (stats.disbAchieve / 10000000).toFixed(2);

    // Helpers for subsets
    const subRow = (label, ach, pln, isCr = false) => {
        const p = isCr ? (pln / 10000000).toFixed(2) + 'Cr' : pln.toLocaleString();
        const a = isCr ? (ach / 10000000).toFixed(2) + 'Cr' : ach.toLocaleString();
        const pct = pln > 0 ? Math.round((ach / pln) * 100) : 0;
        const col = pct >= 100 ? '#10B981' : (pct >= 80 ? '#F59E0B' : '#EF4444');
        return `
            <div style="display:flex; justify-content:space-between; font-size:10px; color:var(--text-secondary); margin-top:2px;">
                <span>${label}</span>
                <span>
                    <span style="color:${col}; font-weight:600;">${a}</span> / ${p}
                </span>
            </div>
        `;
    };

    return `
        <div style="background:var(--bg-body); border:1px solid var(--border-color); border-radius:12px; padding:16px; margin-bottom:20px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid var(--border-color); padding-bottom:8px;">
                <div style="font-weight:700; color:var(--text-primary); font-size:14px;">SUMMARY: ${title}</div>
                <div style="font-size:12px; color:var(--text-secondary);">
                    <span style="font-weight:600; color:var(--primary);">${stats.completed}</span> / ${stats.branches} Completed
                </div>
            </div>
            
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:16px;">
                <!-- Collection -->
                <div>
                    <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px; font-weight:600;">TOTAL COLLECTION</div>
                    <div style="display:flex; align-items:baseline; gap:6px;">
                        <span style="font-weight:700; font-size:16px;">${stats.collAchieve.toLocaleString()}</span>
                        <span style="font-size:11px; color:var(--text-secondary);">/ ${stats.collPlan.toLocaleString()}</span>
                    </div>
                    <div style="height:4px; background:var(--border-color); border-radius:2px; margin-top:6px; overflow:hidden;">
                        <div style="height:100%; width:${collPct}%; background:${collPct >= 100 ? '#10B981' : '#F59E0B'};"></div>
                    </div>
                    <div style="margin-top:6px; padding-top:4px; border-top:1px dashed var(--border-color);">
                        ${subRow("FTOD Accounts", stats.ftodAchieve, stats.ftodPlan)}
                        ${subRow("1-30 DPD", stats.dpd130Achieve, stats.dpd130Plan)}
                        ${subRow("31-60 DPD", stats.dpd3160Achieve, stats.dpd3160Plan)}
                        ${subRow("61-90 DPD", stats.dpd6190Achieve, stats.dpd6190Plan)}
                    </div>
                </div>

                <!-- Disbursement -->
                 <div>
                    <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px; font-weight:600;">TOTAL DISBURSEMENT</div>
                    <div style="display:flex; align-items:baseline; gap:6px;">
                        <span style="font-weight:700; font-size:16px;">Amt ${disbAchieveCr}Cr</span>
                        <span style="font-size:11px; color:var(--text-secondary);">/ ${disbPlanCr}Cr</span>
                    </div>
                    <div style="height:4px; background:var(--border-color); border-radius:2px; margin-top:6px; overflow:hidden;">
                        <div style="height:100%; width:${disbPct}%; background:${disbPct >= 100 ? '#10B981' : '#8B5CF6'};"></div>
                    </div>
                    <div style="margin-top:6px; padding-top:4px; border-top:1px dashed var(--border-color);">
                        ${subRow("IGL (Amt)", stats.disbIglAmtAchieve, stats.disbIglAmtPlan, true)}
                        ${subRow("FIG (Amt)", stats.disbFigAmtAchieve, stats.disbFigAmtPlan, true)}
                        ${subRow("IL (Amt)", stats.disbIlAmtAchieve, stats.disbIlAmtPlan, true)}
                    </div>
                </div>

                <!-- KYC -->
                <div>
                     <div style="font-size:11px; color:var(--text-secondary); margin-bottom:4px; font-weight:600;">TOTAL KYC</div>
                     <div style="font-weight:700; font-size:16px;">${stats.kyc.toLocaleString()}</div>
                     <div style="font-size:10px; color:var(--text-secondary); margin-top:2px;">Sourcing Generated</div>
                </div>
            </div>
        </div>
    `;
}

// --- COMPONENTS ---
function createMetricCard(title, value, iconType, isSuccess = false) {
    const icons = {
        "trending-up": `<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline>`,
        "activity": `<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>`,
        "bar-chart": `<line x1="12" y1="20" x2="12" y2="10"></line><line x1="18" y1="20" x2="18" y2="4"></line><line x1="6" y1="20" x2="6" y2="16"></line>`,
        "award": `<circle cx="12" cy="8" r="7"></circle><polyline points="8.21 13.89 7 23 12 20 17 23 15.79 13.88"></polyline>`
    };

    return `
            <div class="metric-card">
                <div class="metric-header">
                    <div class="metric-icon">
                        <svg class="icon" viewBox="0 0 24 24">${icons[iconType]}</svg>
                    </div>
                     ${isSuccess ? '<div class="metric-badge badge-up"><svg class="icon" style="width:12px;height:12px;" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline></svg> +12%</div>' : ''}
                </div>
                <div>
                    <div class="metric-title">${title}</div>
                    <div class="metric-value">${value}</div>
                </div>
            </div>`;
}

function createDrillDownCard(name, type, children, isLeaf) {
    const card = document.createElement("div");
    card.style.cssText = `
        background: var(--bg-body); border-radius: 12px; padding: 16px;
        cursor: pointer; transition: background 0.2s; border: 1px solid var(--border-color);
        display: flex; flex-direction: column; gap: 8px;
    `;

    // Calculate Average Percentage instead of total
    let displayValue = '--';
    let valueColor = 'var(--text-secondary)';

    if (isLeaf) {
        const avgPct = calculateBranchAveragePercentage(name);
        if (avgPct !== null) {
            displayValue = avgPct + '%';
            // Color coding based on percentage
            if (avgPct >= 100) valueColor = 'var(--success)';
            else if (avgPct >= 50) valueColor = '#F59E0B'; // Warning yellow
            else valueColor = '#EF4444'; // Red
        }
    } else {
        const avgPct = calculateAveragePercentage(type, name, children);
        if (avgPct !== null) {
            displayValue = avgPct + '%';
            // Color coding based on percentage
            if (avgPct >= 100) valueColor = 'var(--success)';
            else if (avgPct >= 50) valueColor = '#F59E0B'; // Warning yellow
            else valueColor = '#EF4444'; // Red
        }
    }

    card.innerHTML = `
        <div style="font-size:12px; color:var(--text-secondary); text-transform:uppercase; font-weight:600;">${type}</div>
        <div style="font-size:15px; font-weight:600; color:var(--text-primary);">${name}</div>
        <div style="font-size:18px; font-weight:700; color:${valueColor}; margin-top:auto;">${displayValue}</div>
    `;

    if (!isLeaf) {
        card.onclick = () => pushStack(name, children);
        card.onmouseover = () => card.style.background = "var(--primary-light)";
        card.onmouseout = () => card.style.background = "var(--bg-body)";
    } else {
        // LEAF NODE (BRANCH) - 3 STATE LOGIC
        const branchEntry = state.branchDetails[name];
        const hasPlan = branchEntry && branchEntry.target;
        const hasAchieve = branchEntry && branchEntry.achievement;

        // STATE 3: FULL GREEN (DONE)
        if (hasPlan && hasAchieve) {
            card.innerHTML += `<div style="margin-top:4px; font-size:11px; color:var(--success); display:flex; align-items:center; gap:4px;">
                        <svg class="icon" style="width:12px;height:12px;" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"></polyline></svg> Completed
                     </div>`;
            card.style.border = "1px solid var(--success-light)";
            card.style.background = "var(--success-light)";
            card.querySelector("div:nth-child(2)").style.color = "var(--success)"; // Text Green
        }
        // STATE 2: GREEN OUTLINE (PENDING ACHIEVEMENT)
        else if (hasPlan && !hasAchieve) {
            card.innerHTML += `<div style="margin-top:4px; font-size:11px; color:var(--success); display:flex; align-items:center; gap:4px;">
                        <svg class="icon" style="width:12px;height:12px;" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle></svg> Set Achievement
                     </div>`;
            card.style.border = "2px solid var(--success)"; // Thicker Border
            card.style.background = "var(--bg-card)";
        }
        // STATE 1: DEFAULT (SET TARGET)
        else {
            card.innerHTML += `<div style="margin-top:4px; font-size:11px; color:var(--warning); display:flex; align-items:center; gap:4px;">
                        <svg class="icon" style="width:12px;height:12px;" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle></svg> Set Plan
                     </div>`;
            card.style.border = "1px solid var(--border-color)";
            // card.style.background stays default
        }

        card.onclick = () => openBranchModal(name);
        card.onmouseover = () => { if (!(hasPlan && hasAchieve)) card.style.background = "var(--primary-light)"; }
        card.onmouseout = () => { if (!(hasPlan && hasAchieve)) card.style.background = "var(--bg-body)"; }
        card.style.cursor = "pointer";
    }

    return card;
}

// --- VIEW SUMMARY WITH PERCENTAGES ---
function createViewSummary(targetData, achieveData) {
    // Helper to safely get number value
    const getVal = (data, field) => {
        if (!data) return 0;
        // Map IL fields to JL database columns
        // Columns now match form field IDs directly
        return Number(data[field]) || 0;
    };

    // Calculate percentage with color coding
    const calcPercent = (target, achieve) => {
        if (target === 0) return { pct: achieve > 0 ? '∞' : '-', color: '#6B7280' };
        const pct = Math.round((achieve / target) * 100);
        let color = '#10B981'; // Green for ≥100%
        if (pct < 100) color = '#F59E0B'; // Yellow for 50-99%
        if (pct < 50) color = '#EF4444'; // Red for <50%
        return { pct: pct + '%', color };
    };

    // Create metric row HTML
    // isActual: When true, don't show achievement value or percentage for "Actual" account rows
    const metricRow = (label, targetField, achieveField, isActual = false) => {
        const t = getVal(targetData, targetField);
        const a = isActual ? '-' : getVal(achieveData, achieveField || targetField);
        const { pct, color } = isActual ? { pct: '-', color: '#6B7280' } : calcPercent(t, a);
        const progressWidth = t === 0 ? 0 : Math.min((a / t) * 100, 100);

        return `
                    <div style="display:flex; justify-content:space-between; align-items:center; padding:12px 0; border-bottom:1px solid var(--border-color);">
                        <div style="flex:2; font-weight:500; color:var(--text-primary);">${label}</div>
                        <div style="flex:1; text-align:center; color:var(--text-secondary);">${t}</div>
                        <div style="flex:1; text-align:center; font-weight:600; color:var(--text-primary);">${a}</div>
                        <div style="flex:1; text-align:center;">
                            <span style="background:${color}20; color:${color}; padding:4px 12px; border-radius:20px; font-weight:700; font-size:13px;">${pct}</span>
                        </div>
                    </div>
                `;
    };

    // Create section HTML
    const section = (title, rows) => `
                <div style="background:var(--bg-card); border-radius:12px; padding:20px; margin-bottom:16px; box-shadow:var(--card-shadow);">
                    <div style="font-size:14px; font-weight:700; color:var(--primary-accent); text-transform:uppercase; margin-bottom:16px; letter-spacing:0.5px;">${title}</div>
                    <div style="display:flex; justify-content:space-between; padding:8px 0; margin-bottom:8px; border-bottom:2px solid var(--border-color);">
                        <div style="flex:2; font-size:11px; color:var(--text-secondary); font-weight:600;">METRIC</div>
                        <div style="flex:1; text-align:center; font-size:11px; color:var(--text-secondary); font-weight:600;">TARGET</div>
                        <div style="flex:1; text-align:center; font-size:11px; color:var(--text-secondary); font-weight:600;">ACHIEVED</div>
                        <div style="flex:1; text-align:center; font-size:11px; color:var(--text-secondary); font-weight:600;">%</div>
                    </div>
                    ${rows}
                </div>
            `;

    // Build HTML
    let html = `
                <div style="padding:20px;">
                    <div style="background:linear-gradient(135deg, #10B981 0%, #059669 100%); color:white; padding:16px 24px; border-radius:12px; margin-bottom:24px; text-align:center;">
                        <div style="font-size:14px; opacity:0.9;">Report Status</div>
                        <div style="font-size:20px; font-weight:700;">✓ Completed - Plan & Achievement Saved</div>
                    </div>

                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
                        <div>
                            ${section('FTOD - Accounts',
        metricRow('Actual FTOD', 'ftod_actual', null, true) +
        metricRow('Collection Plan', 'ftod_plan')
    )}
                            
                            ${section('1-30 DPD',
        metricRow('Actual Account', 'dpd_1_30_actual', null, true) +
        metricRow('Collection Plan', 'dpd_1_30_plan')
    )}

                            ${section('31-60 DPD',
        metricRow('Actual Account', 'dpd_31_60_actual', null, true) +
        metricRow('Collection Plan', 'dpd_31_60_plan')
    )}

                            ${section('61-90 DPD',
        metricRow('Actual Account', 'dpd_61_90_actual', null, true) +
        metricRow('Collection Plan', 'dpd_61_90_plan')
    )}
                            
                            ${section('NPA Accounts',
        metricRow('NPA Activation', 'npa_activation') +
        metricRow('NPA Closure', 'npa_closure')
    )}
                        </div>
                        <div>
                            ${section('FY Non-Starter Accounts',
        metricRow('Non-Starter Accounts', 'fy_non_start_acc', null, true) +
        metricRow('Non-Starter Plan', 'fy_non_start_plan')
    )}
                            
                            ${section('Disbursement',

        metricRow('IGL (Acc)', 'disb_igl_acc') +
        metricRow('IGL (Amt)', 'disb_igl_amt') +
        metricRow('FIG (Acc)', 'disb_fig_acc') +
        metricRow('FIG (Amt)', 'disb_fig_amt') +
        metricRow('IL (Acc)', 'disb_il_acc') +
        metricRow('IL (Amt)', 'disb_il_amt')
    )}

                            ${section('KYC Sourcing',
        metricRow('IGL', 'kyc_igl') +
        metricRow('FIG', 'kyc_fig') +
        metricRow('IL', 'kyc_il')
    )}
                        </div>
                    </div>
                </div>
            `;

    return html;
}

// --- BRANCH DETAILS MODAL LOGIC ---
let currentEditingBranch = null;

// Track current modal state for use in save function
let currentModalState = 'TARGET';

// Track pending yesterday redirect state
let pendingYesterdayRedirect = { branchName: null, yesterdayDate: null };

function showYesterdayReminderModal(branchName) {
    const yesterdayDate = getYesterdayISO();
    pendingYesterdayRedirect = { branchName, yesterdayDate };

    // Set branch name in modal
    document.getElementById('reminderBranchName').textContent = branchName;

    // Format and set yesterday's date
    const dateObj = new Date(yesterdayDate + 'T00:00:00');
    const formattedDate = dateObj.toLocaleDateString('en-US', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    document.getElementById('reminderYesterdayDate').textContent = formattedDate;

    // Show modal
    document.getElementById('yesterdayReminderModal').classList.add('visible');
}

function closeYesterdayReminderModal() {
    document.getElementById('yesterdayReminderModal').classList.remove('visible');
    pendingYesterdayRedirect = { branchName: null, yesterdayDate: null };
}

async function redirectToYesterdayAchievement() {
    const { branchName, yesterdayDate } = pendingYesterdayRedirect;
    if (!branchName || !yesterdayDate) return;

    // Close the reminder modal
    closeYesterdayReminderModal();

    // Change systemDate to yesterday
    state.systemDate = yesterdayDate;

    // Update UI to reflect the date change
    const dateObj = new Date(yesterdayDate + 'T00:00:00');
    const formattedDate = dateObj.toLocaleDateString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        year: 'numeric'
    });
    document.getElementById('headerDate').textContent = formattedDate;

    // Show loading and fetch yesterday's data
    setLoading(true);
    await fetchData();
    setLoading(false);

    // Open branch modal - will show ACHIEVEMENT mode since yesterday has plan but no achievement
    openBranchModal(branchName);
}

// Check if monthly actual fields have already been set for this branch in the current month
async function checkMonthlyActuals(branchName, systemDate) {
    try {
        const result = await neonQuery(
            `SELECT dpd_1_30_actual, dpd_31_60_actual, dpd_61_90_actual, date
             FROM daily_reports
             WHERE branch_name = $1
               AND date >= date_trunc('month', $2::date)
               AND date < date_trunc('month', $2::date) + interval '1 month'
               AND date != $2
               AND (dpd_1_30_actual IS NOT NULL OR dpd_31_60_actual IS NOT NULL OR dpd_61_90_actual IS NOT NULL)
             ORDER BY date ASC LIMIT 1`,
            [branchName, systemDate]
        );
        if (result && result.data && result.data.length > 0) {
            return result.data[0];
        }
        // Also check tagged template result format
        if (result && Array.isArray(result) && result.length > 0) {
            return result[0];
        }
        return null;
    } catch (e) {
        console.warn('Monthly actuals check failed:', e);
        return null;
    }
}

async function openBranchModal(branchName) {
    currentEditingBranch = branchName;
    document.getElementById("modalBranchTitle").textContent = `${branchName} - Details`;

    // GET STATE
    const branchEntry = state.branchDetails[branchName] || {};
    const targetData = branchEntry.target;
    const achieveData = branchEntry.achievement;

    // 1. STATE DETECTION
    currentModalState = 'TARGET'; // Default: State 1
    if (targetData && !achieveData) currentModalState = 'ACHIEVEMENT'; // State 2
    if (targetData && achieveData) currentModalState = 'VIEW'; // State 3

    // Local alias for use in this function
    let reportState = currentModalState;

    // Helper to update labels dynamically
    const updateLabel = (id, text) => {
        const el = document.getElementById(id);
        if (el && el.previousElementSibling) el.previousElementSibling.textContent = text;
    };

    // Helper to toggle visibility
    const toggleRow = (id, visible) => {
        const el = document.getElementById(id);
        if (el && el.parentElement) {
            el.parentElement.style.display = visible ? '' : 'none';
        }
    };

    // Update dynamic titles
    const fySectionEl = document.getElementById('fySectionTitle');
    if (fySectionEl) {
        fySectionEl.textContent = 'FY Non-Starter Accounts';
    }

    if (reportState === 'ACHIEVEMENT') {
        updateLabel('ftod_plan', 'Achievement');
        updateLabel('dpd_1_30_plan', 'Achievement');
        updateLabel('dpd_31_60_plan', 'Achievement');
        updateLabel('dpd_61_90_plan', 'Achievement');

        // Specific labels for Non-Starter
        updateLabel('fy_non_start_plan', 'NON-STARTER Achievement');

        // Update Section Titles
        const lblDisb = document.getElementById('lbl_disb');
        if (lblDisb) lblDisb.textContent = "Disbursement Achievement";

        const lblKyc = document.getElementById('lbl_kyc');
        if (lblKyc) lblKyc.textContent = "KYC SOURCING Achievement";

        // Hide Actuals and set to 0
        ['ftod_actual', 'dpd_1_30_actual', 'dpd_31_60_actual', 'dpd_61_90_actual', 'fy_non_start_acc'].forEach(id => {
            toggleRow(id, false);
            const el = document.getElementById(id);
            if (el) el.value = "0";
        });
    } else {
        updateLabel('ftod_plan', 'FTOD Collection Plan');
        updateLabel('dpd_1_30_plan', 'Collection Plan');
        updateLabel('dpd_31_60_plan', 'Collection Plan');
        updateLabel('dpd_61_90_plan', 'Collection Plan');
        updateLabel('fy_non_start_plan', 'COLLECTION PLAN');

        // Reset Section Titles
        const lblDisb = document.getElementById('lbl_disb');
        if (lblDisb) lblDisb.textContent = "Disbursement Targets";

        const lblKyc = document.getElementById('lbl_kyc');
        if (lblKyc) lblKyc.textContent = "KYC SOURCING";

        // Show Actuals
        ['ftod_actual', 'dpd_1_30_actual', 'dpd_31_60_actual', 'dpd_61_90_actual', 'fy_non_start_acc'].forEach(id => {
            toggleRow(id, true);
        });
    }

    // Helper lists
    const planFields = [
        'ftod_plan', 'dpd_1_30_plan', 'dpd_31_60_plan', 'dpd_61_90_plan', 'fy_non_start_plan'
    ];
    const achieveFields = [
        'ftod_actual', 'dpd_1_30_actual', 'dpd_31_60_actual', 'dpd_61_90_actual', 'npa_activation', 'npa_closure',
        'fy_non_start_acc',
        'disb_igl_acc', 'disb_igl_amt',
        'disb_fig_acc', 'disb_fig_amt',
        'disb_il_acc', 'disb_il_amt',
        'kyc_igl', 'kyc_fig', 'kyc_il'
    ];

    const modalContainer = document.querySelector(".modal-container");
    const footer = document.querySelector(".modal-footer");

    // Clean banners
    const existingBanner = document.querySelector(".read-only-banner");
    if (existingBanner) existingBanner.remove();

    // Clean up any VIEW mode elements
    const existingSummary = document.getElementById("view-summary-container");
    if (existingSummary) existingSummary.remove();

    // Reset panels visibility (in case they were hidden by VIEW mode)
    const leftPanel = document.querySelector(".modal-body .left-panel");
    const rightPanel = document.querySelector(".modal-body .right-panel");
    if (leftPanel) leftPanel.style.display = '';
    if (rightPanel) rightPanel.style.display = '';

    modalContainer.classList.remove("read-only-mode"); // Reset
    footer.style.display = 'flex'; // Reset footer

    // 2. APPLY STATE UI
    if (reportState === 'TARGET') {
        // STATE 1: SET TARGET
        // ALL UNLOCKED
        const banner = document.createElement("div");
        banner.className = "read-only-banner";
        banner.style.background = "#FFC10720"; // Light Yellow
        banner.style.color = "#D97706";
        banner.style.border = "1px solid #D97706";
        banner.textContent = "Step 1: Set Daily Plans";
        document.querySelector(".modal-body").prepend(banner);

        planFields.forEach(id => {
            const el = document.getElementById(id);
            el.value = "";
            el.disabled = false;
        });
        achieveFields.forEach(id => {
            const el = document.getElementById(id);
            el.value = "";
            el.disabled = false; // UNLOCKED
        });

        // Once-per-month lock: check if actual fields were already set this month on a different date
        checkMonthlyActuals(branchName, state.systemDate).then(existingRow => {
            if (!existingRow) return;
            const lockedFields = ['dpd_1_30_actual', 'dpd_31_60_actual', 'dpd_61_90_actual'];
            const lockedDate = existingRow.date;
            lockedFields.forEach(fieldId => {
                const val = existingRow[fieldId];
                if (val !== null && val !== undefined) {
                    const el = document.getElementById(fieldId);
                    if (el) {
                        el.value = val;
                        el.disabled = true;
                        el.style.backgroundColor = '#FFFDE7';
                        el.style.border = '2px solid #F59E0B';
                        el.style.cursor = 'not-allowed';
                        el.title = `Already set for this month (on ${lockedDate})`;
                    }
                }
            });
        });
    }
    else if (reportState === 'ACHIEVEMENT') {
        // STATE 2: SET ACHIEVEMENT
        // Check if viewing TODAY and if current time is before 6 PM
        const todayISO = formatDateISO(new Date());
        const isViewingToday = state.systemDate === todayISO;
        const currentHour = new Date().getHours();
        const isAfter6PM = currentHour >= 18;

        // Only block if TODAY's achievement and before 6 PM
        if (isViewingToday && !isAfter6PM) {
            // Block achievement entry before 6 PM for today only
            const banner = document.createElement("div");
            banner.className = "read-only-banner";
            banner.style.background = "#FEF3C7"; // Light amber
            banner.style.color = "#D97706";
            banner.style.border = "1px solid #F59E0B";
            banner.style.padding = "12px 16px";
            banner.style.borderRadius = "8px";
            banner.style.textAlign = "center";
            banner.style.fontSize = "14px";
            banner.style.fontWeight = "500";
            banner.innerHTML = `<span style="font-size:18px;">⏰</span> You can enter today's achievement after <strong>6 PM</strong>`;
            document.querySelector(".modal-body").prepend(banner);

            // Disable all inputs
            const allInputs = document.querySelectorAll('.modal-body input');
            allInputs.forEach(input => {
                if (input.type !== 'checkbox') {
                    input.disabled = true;
                    input.style.backgroundColor = "#f3f4f6";
                    input.style.cursor = "not-allowed";
                }
            });

            // Hide the save button in footer
            footer.style.display = 'none';
            return;
        }

        // ALL UNLOCKED - Allow entry (either after 6 PM for today, or any time for past dates)
        const banner = document.createElement("div");
        banner.className = "read-only-banner";
        banner.style.background = "#D1FAE5"; // Light Green
        banner.style.color = "#059669";
        banner.textContent = "Step 2: Enter Achievements (Plans Editable)";
        document.querySelector(".modal-body").prepend(banner);

        // Define mappings
        const pairs = {
            'ftod_actual': 'ftod_plan',
            'dpd_1_30_actual': 'dpd_1_30_plan',
            'dpd_31_60_actual': 'dpd_31_60_plan',
            'dpd_61_90_actual': 'dpd_61_90_plan',
            'fy_non_start_acc': 'fy_non_start_plan'
        };

        // DYNAMICALLY HANDLE ALL NON-PLAN INPUTS
        const allInputs = document.querySelectorAll('.modal-body input');

        allInputs.forEach(input => {

            const id = input.id;
            if (!id || input.type === 'checkbox') return;

            // UNIVERSAL LOGIC: Clear Value, Show Placeholder
            input.value = "";
            input.disabled = false;

            // Determine Plan Key - map form fields to database columns
            let targetKey = pairs[id] || id;
            // Map form field to database column for IL fields
            // Columns now match form field IDs directly

            if (targetData && targetData[targetKey] !== undefined && targetData[targetKey] !== null && targetData[targetKey] !== "") {
                input.placeholder = `Plan: ${targetData[targetKey]}`;
            } else {
                input.placeholder = "Enter Achievement...";
            }
        });

        // Disbursement fields (disb_igl_acc, disb_igl_amt, disb_fig_acc, disb_fig_amt, disb_il_acc, disb_il_amt) are now unblocked
        // and can be edited in achievement mode
    }
    else {
        // STATE 3: VIEW ONLY - Beautiful UI with Percentage Calculations
        modalContainer.classList.add("read-only-mode");

        // Hide left and right panels (the form sections)
        const leftPanel = document.querySelector(".modal-body .left-panel");
        const rightPanel = document.querySelector(".modal-body .right-panel");
        if (leftPanel) leftPanel.style.display = 'none';
        if (rightPanel) rightPanel.style.display = 'none';

        // Remove any existing summary container
        const existingSummary = document.getElementById("view-summary-container");
        if (existingSummary) existingSummary.remove();

        // Create beautiful summary view
        const summaryHTML = createViewSummary(targetData, achieveData);
        const summaryContainer = document.createElement("div");
        summaryContainer.id = "view-summary-container";
        summaryContainer.style.gridColumn = "1 / -1"; // Span full width
        summaryContainer.innerHTML = summaryHTML;
        document.querySelector(".modal-body").appendChild(summaryContainer);
    }



    // Inject Test Connection
    const btnTestId = "btn-test-conn";
    if (!document.getElementById(btnTestId)) {
        const btn = document.createElement("button");
        btn.id = btnTestId;
        btn.className = "btn btn-outline";
        btn.style.width = "auto";
        btn.style.fontSize = "10px";
        btn.style.padding = "4px 8px";
        btn.innerText = "Test Connection";
        btn.onclick = testConnection;
        // Prepend to footer or right panel? Footer is best
        footer.insertBefore(btn, footer.firstChild);
    }

    // Add Indian number formatting to amount fields
    const amountFields = ['ftod_plan', 'dpd_1_30_plan', 'dpd_31_60_plan', 'dpd_61_90_plan',
                          'fy_non_start_plan', 'disb_igl_amt', 'disb_fig_amt', 'disb_il_amt'];
    amountFields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.oninput = function() {
                const cursorPos = this.selectionStart;
                const oldLength = this.value.length;
                const raw = this.value.replace(/,/g, '');
                this.value = formatIndianNumber(raw);
                const newLength = this.value.length;
                const newPos = cursorPos + (newLength - oldLength);
                this.setSelectionRange(newPos, newPos);
            };
        }
    });

    document.getElementById("branchModal").classList.add("visible");
}

function closeBranchModal() {
    document.getElementById("branchModal").classList.remove("visible");
    currentEditingBranch = null;
}


async function saveBranchDetails(andNext) {
    if (!currentEditingBranch) return;

    // Use the tracked modal state instead of re-calculating from state.branchDetails
    // This ensures the correct table is used even if state hasn't been refreshed
    const saveMode = currentModalState;


    // Don't save in VIEW mode
    if (saveMode === 'VIEW') {
        showToast("View mode - no changes to save", "alert");
        return;
    }

    const data = {};
    // Fields to capture based on mode
    // If Mode TARGET: Capture Plan, Zero Actuals
    // If Mode ACHIEVEMENT: Capture Actuals, Keep Plan from Plan Record

    const allFields = [
        'ftod_actual', 'ftod_plan', 'dpd_1_30_actual', 'dpd_1_30_plan',
        'dpd_31_60_actual', 'dpd_31_60_plan', 'dpd_61_90_actual', 'dpd_61_90_plan',
        'npa_activation', 'npa_closure',
        'fy_non_start_acc', 'fy_non_start_plan',
        'disb_igl_acc', 'disb_igl_amt',
        'disb_fig_acc', 'disb_fig_amt',
        'disb_il_acc', 'disb_il_amt',
        'kyc_igl', 'kyc_fig', 'kyc_il'
    ];

    // Reset any error styling from previous attempts
    allFields.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.classList.remove("input-error");
    });

    // 1. Validate for empty fields - check all visible/enabled inputs
    for (const id of allFields) {
        const el = document.getElementById(id);
        if (el) {
            // Check if the input is visible and not disabled/readonly
            const isVisible = el.offsetParent !== null;
            const isEnabled = !el.disabled && !el.readOnly;

            // Check if empty (trim to handle whitespace-only values)
            const value = el.value.trim();

            if (isVisible && isEnabled && value === '') {
                // Add error styling
                el.classList.add("input-error");
                // Focus on the empty field
                el.focus();
                // Scroll the input into view if needed
                el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                // Show error message
                showToast("Enter an amount; if it does not apply, type 0 instead of leaving it empty.", "alert");
                return; // Stop save process
            }
        }
    }

    // 2. Gather Data from Form
    // Note: Even locked fields return values, which is good for copying Plans in Achievement Mode
    for (const id of allFields) {
        const el = document.getElementById(id);
        if (el) {
            // No longer requiring all fields - allow empty values
            // Empty values will be sent as empty string and converted to null/0 by saveData
            // Strip commas from formatted numbers before saving
            data[id] = unformatNumber(el.value);
        }
    }


    // 2. Prepare Payload based on Mode
    // If ACHIEVEMENT mode, we need to ensure PLAN values are preserved from the Plan record
    // However, since we populate them into the inputs (locked), reading the input value `data[id]` works perfectly.

    // 3. Save
    const table = saveMode === 'TARGET' ? 'daily_reports' : 'daily_reports_achievements';

    // Persist
    const result = await saveData(currentEditingBranch, data, table);

    if (!result || !result.success) {
        // Toast already shown in saveData
        return;
    }

    // Use DB-returned row; columns now match form field IDs directly
    const savedRow = result.row || {};
    const normalized = {
        ...savedRow,
        branch_name: currentEditingBranch,
        date: state.systemDate
    };

    // Optimistic Update
    if (!state.branchDetails[currentEditingBranch]) state.branchDetails[currentEditingBranch] = {};

    if (saveMode === 'TARGET') state.branchDetails[currentEditingBranch].target = normalized;
    else state.branchDetails[currentEditingBranch].achievement = normalized;

    showToast(`Details Saved (${saveMode === 'TARGET' ? 'Plan' : 'Achievement'})`, "check");

    if (andNext) {
        loadNextBranch();
    } else {
        closeBranchModal();
    }
    // Refresh view to show checkmark
    try {
        if (state.activeTab === 'hierarchy') renderDashboard();
    } catch (e) {
        console.error("Error updating dashboard:", e);
        showToast("Saved, but view refresh failed.", "alert");
    }
}

async function testConnection() {
    showToast("Testing connection...", "info");
    try {
        await neonQuery('SELECT 1');
        showToast("Connection Successful!", "check");
    } catch (err) {
        showToast("Connection Failed: " + err.message, "alert");
        console.error("Connection Test Error:", err);
    }
}

function loadNextBranch() {
    // Get DM's branches only
    let dmBranches = [];

    if (state.role === 'DM') {
        const idxDM = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === "dm name");
        // Get rows for this DM ONLY
        const dmRows = state.rawData.rows.filter(r => (r[idxDM] || "").trim() === state.currentUser);
        // Extract branch names and sort
        dmBranches = dmRows.map(r => r[0]).filter(Boolean).sort();
    } else {
        // CEO sees all
        dmBranches = state.rawData.rows.map(r => r[0]).filter(Boolean).sort();
    }

    // Check completion status for DM's branches
    const checkCompletionStatus = () => {
        let branchesWithoutTarget = [];
        let branchesWithoutAchievement = [];

        dmBranches.forEach(branch => {
            const branchEntry = state.branchDetails[branch] || {};
            if (!branchEntry.target) {
                branchesWithoutTarget.push(branch);
            } else if (!branchEntry.achievement) {
                branchesWithoutAchievement.push(branch);
            }
        });

        return { branchesWithoutTarget, branchesWithoutAchievement };
    };

    const { branchesWithoutTarget, branchesWithoutAchievement } = checkCompletionStatus();

    // Determine what we were just saving
    const justSavedMode = currentModalState;

    // Find the next branch based on what we just saved
    if (justSavedMode === 'TARGET') {
        // We just saved a target - find next branch without target
        if (branchesWithoutTarget.length > 0) {
            // Find next branch after current one that needs target
            const currentIdx = dmBranches.indexOf(currentEditingBranch);
            let nextBranch = null;

            // Look for next branch without target after current position
            for (let i = currentIdx + 1; i < dmBranches.length; i++) {
                if (branchesWithoutTarget.includes(dmBranches[i])) {
                    nextBranch = dmBranches[i];
                    break;
                }
            }

            // If not found, wrap around to beginning
            if (!nextBranch) {
                for (let i = 0; i < currentIdx; i++) {
                    if (branchesWithoutTarget.includes(dmBranches[i])) {
                        nextBranch = dmBranches[i];
                        break;
                    }
                }
            }

            if (nextBranch) {
                openBranchModal(nextBranch);
                return;
            }
        }

        // All targets are done - prompt for achievements
        closeBranchModal();
        if (branchesWithoutAchievement.length > 0) {
            showCompletionModal('targets', branchesWithoutAchievement[0]);
        } else {
            showCompletionModal('all');
        }
        return;
    }
    else if (justSavedMode === 'ACHIEVEMENT') {
        // We just saved an achievement - find next branch without achievement
        if (branchesWithoutAchievement.length > 0) {
            // Find next branch after current one that needs achievement
            const currentIdx = dmBranches.indexOf(currentEditingBranch);
            let nextBranch = null;

            // Look for next branch without achievement after current position
            for (let i = currentIdx + 1; i < dmBranches.length; i++) {
                if (branchesWithoutAchievement.includes(dmBranches[i])) {
                    nextBranch = dmBranches[i];
                    break;
                }
            }

            // If not found, wrap around to beginning
            if (!nextBranch) {
                for (let i = 0; i < currentIdx; i++) {
                    if (branchesWithoutAchievement.includes(dmBranches[i])) {
                        nextBranch = dmBranches[i];
                        break;
                    }
                }
            }

            if (nextBranch) {
                openBranchModal(nextBranch);
                return;
            }
        }

        // All achievements are done
        closeBranchModal();
        showCompletionModal('all');
        return;
    }

    // Fallback: just go to next sequential branch
    const idx = dmBranches.indexOf(currentEditingBranch);
    if (idx !== -1 && idx < dmBranches.length - 1) {
        const nextBranch = dmBranches[idx + 1];
        openBranchModal(nextBranch);
    } else {
        showToast("Reached end of your branch list", "alert");
        closeBranchModal();
    }
}

// Show completion modal for targets/achievements done
function showCompletionModal(type, firstBranchForAchievement = null) {
    // Create overlay
    const overlay = document.createElement('div');
    overlay.id = 'completionOverlay';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0, 0, 0, 0.7);
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10000;
        animation: fadeIn 0.3s ease;
    `;

    let content = '';

    if (type === 'targets') {
        // All targets done, prompt for achievements
        content = `
            <div style="background: var(--bg-card); border-radius: 20px; padding: 40px; text-align: center; max-width: 450px; box-shadow: 0 20px 60px rgba(0,0,0,0.3);">
                <div style="width: 80px; height: 80px; background: linear-gradient(135deg, #10B981, #059669); border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 24px;">
                    <svg viewBox="0 0 24 24" style="width: 40px; height: 40px; stroke: white; fill: none; stroke-width: 2.5;">
                        <polyline points="20 6 9 17 4 12"></polyline>
                    </svg>
                </div>
                <h2 style="color: var(--text-primary); font-size: 24px; margin-bottom: 12px; font-weight: 700;">All Targets Uploaded! 🎯</h2>
                <p style="color: var(--text-secondary); font-size: 16px; margin-bottom: 32px; line-height: 1.6;">
                    Great work! All your branch targets have been set successfully.<br>
                    <strong style="color: var(--text-primary);">Would you like to update achievements now?</strong>
                </p>
                <div style="display: flex; gap: 16px; justify-content: center;">
                    <button onclick="closeCompletionModal()" style="padding: 14px 28px; border-radius: 12px; border: 2px solid var(--border-color); background: transparent; color: var(--text-primary); font-weight: 600; cursor: pointer; font-size: 15px; transition: all 0.2s;">
                        Maybe Later
                    </button>
                    <button onclick="startAchievementEntry('${firstBranchForAchievement}')" style="padding: 14px 28px; border-radius: 12px; border: none; background: linear-gradient(135deg, #6366F1, #4F46E5); color: white; font-weight: 600; cursor: pointer; font-size: 15px; box-shadow: 0 4px 15px rgba(99, 102, 241, 0.4); transition: all 0.2s;">
                        Yes, Update Achievements →
                    </button>
                </div>
            </div>
        `;
    } else {
        // All done!
        content = `
            <div style="background: var(--bg-card); border-radius: 20px; padding: 40px; text-align: center; max-width: 450px; box-shadow: 0 20px 60px rgba(0,0,0,0.3);">
                <div style="width: 80px; height: 80px; background: linear-gradient(135deg, #10B981, #059669); border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 24px; animation: pulse 2s infinite;">
                    <svg viewBox="0 0 24 24" style="width: 40px; height: 40px; stroke: white; fill: none; stroke-width: 2.5;">
                        <polyline points="20 6 9 17 4 12"></polyline>
                    </svg>
                </div>
                <h2 style="color: var(--text-primary); font-size: 28px; margin-bottom: 12px; font-weight: 700;">All Done! 🎉</h2>
                <p style="color: var(--text-secondary); font-size: 18px; margin-bottom: 12px; line-height: 1.6;">
                    All targets and achievements have been uploaded successfully.
                </p>
                <p style="color: #10B981; font-size: 20px; font-weight: 600; margin-bottom: 32px;">
                    Thank you for your hard work! 💪
                </p>
                <button onclick="closeCompletionModal()" style="padding: 14px 36px; border-radius: 12px; border: none; background: linear-gradient(135deg, #10B981, #059669); color: white; font-weight: 600; cursor: pointer; font-size: 16px; box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4); transition: all 0.2s;">
                    Close
                </button>
            </div>
        `;
    }

    overlay.innerHTML = content;
    document.body.appendChild(overlay);

    // Add animation keyframes if not present
    if (!document.getElementById('completionAnimations')) {
        const style = document.createElement('style');
        style.id = 'completionAnimations';
        style.textContent = `
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }
            @keyframes pulse {
                0%, 100% { transform: scale(1); }
                50% { transform: scale(1.05); }
            }
        `;
        document.head.appendChild(style);
    }
}

// Close completion modal
function closeCompletionModal() {
    const overlay = document.getElementById('completionOverlay');
    if (overlay) {
        overlay.style.opacity = '0';
        setTimeout(() => overlay.remove(), 300);
    }
}

// Start achievement entry from completion modal
function startAchievementEntry(branchName) {
    closeCompletionModal();
    if (branchName) {
        setTimeout(() => openBranchModal(branchName), 300);
    }
}

// Debounced search for better performance during rapid typing
const debouncedSearch = debounce((query) => {
    _handleSearchInternal(query);
}, 150);

function handleSearch(query) {
    // For very short queries, respond immediately to clear results
    if (!query || query.length < 2) {
        const resultsDiv = document.getElementById("searchResults");
        if (resultsDiv) resultsDiv.classList.remove("visible");
        return;
    }
    debouncedSearch(query);
}

// Internal search function (called via debounce)
function _handleSearchInternal(query) {
    const resultsDiv = document.getElementById("searchResults");
    if (!resultsDiv || !query || query.length < 2) {
        if (resultsDiv) resultsDiv.classList.remove("visible");
        return;
    }

    const searchLower = query.toLowerCase();
    const branches = state.rawData.rows
        .map(r => ({ name: r[0], district: r[1] }))
        .filter(b => b.name && b.name.toLowerCase().includes(searchLower))
        .slice(0, 10);

    if (branches.length === 0) {
        resultsDiv.classList.remove("visible");
        return;
    }

    // Use DocumentFragment for better performance
    const fragment = document.createDocumentFragment();
    branches.forEach(match => {
        const branchEntry = state.branchDetails[match.name];
        const hasPlan = branchEntry && branchEntry.target;
        const hasAchieve = branchEntry && branchEntry.achievement;

        const item = document.createElement("div");
        item.className = "search-item";

        let statusIcon = '';
        if (hasPlan && hasAchieve) {
            statusIcon = '<svg class="icon" style="width:14px;height:14px;color:var(--success);" viewBox="0 0 24 24"><polyline points="20 6 9 17 4 12"></polyline></svg>';
        } else if (hasPlan) {
            statusIcon = '<svg class="icon" style="width:14px;height:14px;color:var(--warning);" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle></svg>';
        }

        item.innerHTML = `
                    <div>
                        <div style="font-weight:600">${match.name}</div>
                        <div style="font-size:11px; color:var(--text-secondary);">${match.district || ''}</div>
                    </div>
                    ${statusIcon}
                `;
        item.onclick = () => {
            openBranchModal(match.name);
            resultsDiv.classList.remove("visible");
            const searchInput = document.querySelector(".search-input");
            if (searchInput) searchInput.value = "";
        };
        fragment.appendChild(item);
    });

    resultsDiv.innerHTML = "";
    resultsDiv.appendChild(fragment);
    resultsDiv.classList.add("visible");
}

// Close search when clicking outside
document.addEventListener('click', (e) => {
    const searchContainer = document.querySelector('.search-container');
    if (searchContainer && !searchContainer.contains(e.target)) {
        document.getElementById("searchResults").classList.remove("visible");
    }
});


function renderLineChart(data, container) {
    const width = 400;
    const height = 200;
    const padding = 20;

    const keys = Object.keys(data);
    const values = Object.values(data);
    const max = Math.max(...values, 1);

    // Generate Points
    const points = values.map((val, idx) => {
        const x = (idx / (values.length - 1)) * (width - 2 * padding) + padding;
        const y = height - ((val / max) * (height - 2 * padding)) - padding;
        return `${x},${y}`;
    }).join(" ");

    // SVG
    container.innerHTML = `
                <svg width="100%" height="100%" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
                    <!-- Grid Lines -->
                    <line x1="${padding}" y1="${height - padding}" x2="${width - padding}" y2="${height - padding}" stroke="var(--border-color)" stroke-width="1" />
                    <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${height - padding}" stroke="var(--border-color)" stroke-width="1" />
                    
                    <!-- Path -->
                    <polyline points="${points}" fill="none" stroke="var(--primary-accent)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
                    
                    <!-- Points -->
                    ${values.map((val, idx) => {
        const x = (idx / (values.length - 1)) * (width - 2 * padding) + padding;
        const y = height - ((val / max) * (height - 2 * padding)) - padding;
        return `<circle cx="${x}" cy="${y}" r="4" fill="var(--bg-card)" stroke="var(--primary-accent)" stroke-width="2" />`;
    }).join('')}
                </svg>
                <div style="display:flex; justify-content:space-between; width:100%; margin-top:10px; font-size:10px; color:var(--text-secondary);">
                    ${keys.map(k => `<span>${k.substring(0, 3)}</span>`).join('')}
                </div>
            `;
    container.style.flexDirection = "column";
}

function renderDonutChart(data, container) {
    const total = Object.values(data).reduce((a, b) => a + b, 0) || 1;
    let currentAngle = 0;
    const colors = ['#6366F1', '#10B981', '#F59E0B']; // Primary, Success, Warning

    // SVG Slices
    const slices = Object.entries(data).map(([key, val], idx) => {
        const pct = val / total;
        const angle = pct * 360;

        return `${colors[idx]} ${currentAngle}deg ${(currentAngle += angle)}deg`;
    });

    const gradient = `conic-gradient(${slices.join(', ')})`;

    container.innerHTML = `
               <div style="position:relative; width:180px; height:180px; border-radius:50%; background:${gradient}; display:flex; align-items:center; justify-content:center;">
                   <div style="width:120px; height:120px; background:var(--bg-card); border-radius:50%; display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:10;">
                       <span style="font-size:12px; color:var(--text-secondary);">Total</span>
                       <span style="font-size:16px; font-weight:700;">₹${(total / 100000).toFixed(1)}L</span>
                   </div>
               </div>
               <div style="margin-left:20px; display:flex; flex-direction:column; gap:8px;">
                   ${Object.entries(data).map(([k, v], i) => `
                        <div style="display:flex; align-items:center; gap:8px; font-size:12px;">
                            <span style="width:10px; height:10px; background:${colors[i]}; border-radius:2px;"></span>
                            <span style="color:var(--text-secondary);">${k}</span>
                            <span style="font-weight:600;">${Math.round((v / total) * 100)}%</span>
                        </div>
                   `).join('')}
               </div>
            `;
    container.style.flexDirection = "row";
}

function renderBarChart(data, container) {
    const max = Math.max(...Object.values(data), 1);
    Object.keys(data).forEach(label => {
        const val = data[label];
        const pct = (val / max) * 100;
        container.innerHTML += `
                    <div class="bar-group">
                        <div class="bar" style="height:${pct}%"></div>
                        <div class="bar-label">${label.substring(0, 3)}</div>
                    </div>`;
    });
}

// ============ NEW CEO DASHBOARD HELPER FUNCTIONS ============

function renderCollectionMetric(label, target, achieve) {
    const pct = target > 0 ? Math.round((achieve / target) * 100) : 0;
    const progressWidth = Math.min(pct, 100);

    let color = '#10B981'; // Green
    if (pct < 100) color = '#F59E0B'; // Yellow
    if (pct < 50) color = '#EF4444'; // Red

    return `
                <div style="background:var(--bg-body); padding:20px; border-radius:12px;">
                    <div style="font-size:12px; color:var(--text-secondary); font-weight:600; margin-bottom:12px;">${label}</div>
                    <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:8px;">
                        <span style="font-size:24px; font-weight:700; color:var(--text-primary);">${achieve.toLocaleString()}</span>
                        <span style="font-size:12px; color:var(--text-secondary);">/ ${target.toLocaleString()}</span>
                    </div>
                    <div style="height:8px; background:var(--border-color); border-radius:4px; overflow:hidden;">
                        <div style="height:100%; width:${progressWidth}%; background:${color}; border-radius:4px; transition:width 0.5s ease;"></div>
                    </div>
                    <div style="text-align:right; margin-top:6px;">
                        <span style="font-size:13px; font-weight:700; color:${color};">${pct}%</span>
                    </div>
                </div>
            `;
}

function renderNPAMovement(activation, closure) {
    const net = activation - closure;
    const isPositive = net > 0;
    const color = isPositive ? '#EF4444' : '#10B981'; // Red for increase, Green for decrease

    return `
                <div style="background:var(--bg-body); padding:20px; border-radius:12px;">
                    <div style="font-size:12px; color:var(--text-secondary); font-weight:600; margin-bottom:12px;">NPA Movement</div>
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:12px;">
                        <div style="text-align:center;">
                            <div style="font-size:18px; font-weight:700; color:#EF4444;">${activation}</div>
                            <div style="font-size:10px; color:var(--text-secondary);">Activated</div>
                        </div>
                        <div style="text-align:center;">
                            <div style="font-size:18px; font-weight:700; color:#10B981;">${closure}</div>
                            <div style="font-size:10px; color:var(--text-secondary);">Closed</div>
                        </div>
                    </div>
                    <div style="background:${color}15; padding:8px; border-radius:8px; text-align:center;">
                        <span style="font-size:14px; font-weight:700; color:${color};">
                            ${isPositive ? '↑' : '↓'} Net: ${Math.abs(net)}
                        </span>
                    </div>
                </div>
            `;
}

function renderPortfolioDonut(data, container) {
    if (!container) return;

    const total = (data.healthy || 0) + (data.dpd130 || 0) + (data.dpd3160 || 0) + (data.dpd6190 || 0);
    if (total === 0) {
        container.innerHTML = '<div style="color:var(--text-secondary); text-align:center;">No data available</div>';
        return;
    }

    const colors = { healthy: '#10B981', dpd130: '#F59E0B', dpd3160: '#FB923C', dpd6190: '#EF4444' };
    const labels = { healthy: 'Healthy', dpd130: '1-30 DPD', dpd3160: '31-60 DPD', dpd6190: '61-90 DPD' };

    let currentAngle = 0;
    const slices = Object.entries(data).map(([key, val]) => {
        const pct = val / total;
        const angle = pct * 360;
        const result = `${colors[key]} ${currentAngle}deg ${currentAngle + angle}deg`;
        currentAngle += angle;
        return result;
    });

    const gradient = `conic-gradient(${slices.join(', ')})`;

    container.innerHTML = `
               <div style="position:relative; width:160px; height:160px; border-radius:50%; background:${gradient}; display:flex; align-items:center; justify-content:center;">
                   <div style="width:100px; height:100px; background:var(--bg-card); border-radius:50%; display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:10;">
                       <span style="font-size:11px; color:var(--text-secondary);">Total</span>
                       <span style="font-size:18px; font-weight:700;">${total.toLocaleString()}</span>
                   </div>
               </div>
               <div style="margin-left:24px; display:flex; flex-direction:column; gap:12px;">
                   ${Object.entries(data).map(([k, v]) => `
                        <div style="display:flex; align-items:center; gap:10px;">
                            <span style="width:12px; height:12px; background:${colors[k]}; border-radius:3px;"></span>
                            <div>
                                <div style="font-size:11px; color:var(--text-secondary);">${labels[k]}</div>
                                <div style="font-size:16px; font-weight:700; color:var(--text-primary);">${v.toLocaleString()}</div>
                            </div>
                            <span style="font-size:12px; font-weight:600; color:${colors[k]}; margin-left:auto;">${Math.round((v / total) * 100)}%</span>
                        </div>
                   `).join('')}
               </div>
            `;
    container.style.flexDirection = "row";
}

function renderRegionalTable(data, tbody) {
    if (!tbody || !data) return;

    Object.keys(data).sort().forEach(region => {
        const r = data[region];
        const completionPct = r.branches > 0 ? Math.round((r.completed / r.branches) * 100) : 0;

        let statusColor = '#10B981';
        let statusText = 'On Track';
        if (completionPct < 80) { statusColor = '#F59E0B'; statusText = 'In Progress'; }
        if (completionPct < 50) { statusColor = '#EF4444'; statusText = 'Needs Attention'; }

        let achieveColor = '#10B981';
        if (r.avgPct < 100) achieveColor = '#F59E0B';
        if (r.avgPct < 50) achieveColor = '#EF4444';

        tbody.innerHTML += `
                    <tr class="clickable-row" onclick="openDetailModal('region', '${region}')" style="cursor:pointer;">
                        <td><span style="font-weight:600;">${region}</span> <span style="font-size:10px; color:var(--text-secondary);">→ Click to view</span></td>
                        <td>${r.branches}</td>
                        <td>${r.completed} <span style="color:var(--text-secondary);">(${completionPct}%)</span></td>
                        <td>
                            <span style="background:${achieveColor}20; color:${achieveColor}; padding:4px 12px; border-radius:20px; font-weight:600; font-size:12px;">
                                ${r.avgPct}%
                            </span>
                        </td>
                        <td><span class="status-pill" style="background:${statusColor}20; color:${statusColor};">${statusText}</span></td>
                    </tr>`;
    });
}

function renderLeaderboard(data, tbody) {
    const sorted = Object.entries(data).sort((a, b) => b[1] - a[1]).slice(0, 5);
    sorted.forEach((item, idx) => {
        tbody.innerHTML += `
                    <tr>
                        <td>#${idx + 1}</td>
                        <td><span style="font-weight:600">${item[0]}</span></td>
                        <td>${item[1].toLocaleString()}</td>
                        <td><span class="status-pill status-active">Top Tier</span></td>
                    </tr>`;
    });
}

function renderDMTable(tbody) {
    const idxDM = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === "dm name");
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === "branch");
    const idxDistrict = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === "district");

    const rows = state.rawData.rows.filter(r => (r[idxDM] || "").trim() === state.currentUser);

    rows.forEach(row => {
        const branch = row[idxBranch];
        const district = row[idxDistrict];
        const val = state.targetData[branch] || 0;

        tbody.innerHTML += `
                <tr>
                    <td>${district}</td>
                    <td>${branch}</td>
                    <td>
                        <input type="number" data-branch="${branch}" value="${val}" style="padding:8px; border:1px solid #ddd; border-radius:6px; width:100px;">
                    </td>
                    <td><span class="status-pill status-pending">Pending</span></td>
                </tr>`;
    });
}

function buildHierarchy(headers, rows) {
    // V2: Region → Division → Area → Branch
    const tree = {};
    rows.forEach(r => {
        const [branch, area, division, region] = r;
        if (!region) return;
        if (!tree[region]) tree[region] = {};
        if (!tree[region][division]) tree[region][division] = {};
        if (!tree[region][division][area]) tree[region][division][area] = new Set();
        tree[region][division][area].add(branch);
    });
    return tree;
}

function calculateTotal(type, name, children) {
    if (children instanceof Set) {
        let sum = 0;
        children.forEach(branch => sum += (state.targetData[branch] || 0));
        return sum;
    }
    let sum = 0;
    Object.keys(children).forEach(k => sum += calculateTotal(getNextType(type), k, children[k]));
    return sum;
}

// Calculate average percentage for a single branch based on all metrics
function calculateBranchAveragePercentage(branchName) {
    const branchEntry = state.branchDetails[branchName];
    if (!branchEntry || !branchEntry.target || !branchEntry.achievement) {
        return null; // No data available
    }

    const targetData = branchEntry.target;
    const achieveData = branchEntry.achievement;

    // All fields that have target/achievement pairs
    const fields = [
        'ftod_actual', 'ftod_plan', 'dpd_1_30_actual', 'dpd_1_30_plan',
        'dpd_31_60_actual', 'dpd_31_60_plan', 'dpd_61_90_actual', 'dpd_61_90_plan',
        'npa_activation', 'npa_closure',
        'fy_non_start_acc', 'fy_non_start_plan',
        'disb_igl_acc', 'disb_igl_amt', 'disb_fig_acc', 'disb_fig_amt',
        'disb_il_acc', 'disb_il_amt', 'kyc_igl', 'kyc_fig', 'kyc_il'
    ];

    let validPercentages = [];

    fields.forEach(field => {
        const target = Number(targetData[field]) || 0;
        const achieve = Number(achieveData[field]) || 0;

        if (target > 0) {
            const pct = (achieve / target) * 100;
            validPercentages.push(pct);
        }
    });

    if (validPercentages.length === 0) return 0;

    const avgPct = validPercentages.reduce((sum, p) => sum + p, 0) / validPercentages.length;
    return Math.round(avgPct);
}

// Calculate average percentage for hierarchy nodes (aggregated from children)
function calculateAveragePercentage(type, name, children) {
    if (children instanceof Set) {
        // Leaf level - aggregate all branch percentages
        let validBranches = [];
        children.forEach(branch => {
            const pct = calculateBranchAveragePercentage(branch);
            if (pct !== null) {
                validBranches.push(pct);
            }
        });

        if (validBranches.length === 0) return null;
        return Math.round(validBranches.reduce((sum, p) => sum + p, 0) / validBranches.length);
    }

    // Non-leaf - aggregate from children
    let validChildren = [];
    Object.keys(children).forEach(k => {
        const pct = calculateAveragePercentage(getNextType(type), k, children[k]);
        if (pct !== null) {
            validChildren.push(pct);
        }
    });

    if (validChildren.length === 0) return null;
    return Math.round(validChildren.reduce((sum, p) => sum + p, 0) / validChildren.length);
}

function getNextType(t) { return t === "Region" ? "Division" : t === "Division" ? "Area" : "Branch"; }

function parseTSV(text) {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    const rows = lines.slice(1).map(line => line.split("\t").map(c => c.trim()));
    return { headers: lines[0].split("\t"), rows };
}

function populateDMDropdown(rows, headers) {
    // V2: Build location lists for validation
    window._drLocationData = {
        RM: Array.from(new Set(rows.map(r => r[3]).filter(Boolean))).sort(),
        DM: Array.from(new Set(rows.map(r => r[2]).filter(Boolean))).sort(),
        AM: Array.from(new Set(rows.map(r => r[1]).filter(Boolean))).sort(),
        BM: Array.from(new Set(rows.map(r => r[0]).filter(Boolean))).sort()
    };
    initRolePopup();
    initEmpIdLogin();
}

// ===== ROLE POPUP =====
function initRolePopup() {
    window._drSelectedRole = null;
    window._drSelectedEmployee = null;

    window._drOpenRolePopup = function() {
        document.getElementById('drRolePopup').style.display = '';
    };
    window._drCloseRolePopup = function() {
        document.getElementById('drRolePopup').style.display = 'none';
    };

    // Attach click handlers to popup items
    document.querySelectorAll('.dr-popup-item').forEach(function(item) {
        item.addEventListener('click', function() {
            var role = this.getAttribute('data-role');
            var label = this.getAttribute('data-label');

            window._drCloseRolePopup();

            // CEO → direct login
            if (role === 'CEO') {
                handleLogin('CEO');
                return;
            }

            // Store selected role
            window._drSelectedRole = role;
            window._drSelectedEmployee = null;

            // Show search section
            var wrap = document.getElementById('drEmpSearchWrap');
            wrap.style.display = '';

            // Show role chip
            var chip = document.getElementById('drRoleChip');
            chip.innerHTML = label + ' <span class="dr-chip-x" onclick="window._drResetRole(event)">&times;</span>';

            // Reset search
            document.getElementById('drEmpSearch').value = '';
            document.getElementById('drSearchResults').className = 'dr-search-results';
            document.getElementById('drSelectedEmp').style.display = 'none';
            document.getElementById('login-btn-dm').style.display = 'none';

            // Focus search
            setTimeout(function(){ document.getElementById('drEmpSearch').focus(); }, 100);
        });
    });

    window._drResetRole = function(e) {
        if (e) e.stopPropagation();
        window._drSelectedRole = null;
        window._drSelectedEmployee = null;
        document.getElementById('drEmpSearchWrap').style.display = 'none';
    };
}

// ===== EMP ID TYPEAHEAD =====
function initEmpIdLogin() {
    var searchInput = document.getElementById('drEmpSearch');
    var resultsDiv = document.getElementById('drSearchResults');
    if (!searchInput || !resultsDiv) return;

    var timer = null;
    var activeIdx = -1;

    searchInput.addEventListener('input', function() {
        clearTimeout(timer);
        var q = this.value.trim();
        if (q.length < 2) {
            resultsDiv.className = 'dr-search-results';
            resultsDiv.innerHTML = '';
            return;
        }
        timer = setTimeout(function() { doSearch(q); }, 300);
    });

    // Keyboard navigation
    searchInput.addEventListener('keydown', function(e) {
        var rows = resultsDiv.querySelectorAll('.dr-emp-row');
        if (!rows.length) return;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            activeIdx = Math.min(activeIdx + 1, rows.length - 1);
            updateActive(rows);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            activeIdx = Math.max(activeIdx - 1, 0);
            updateActive(rows);
        } else if (e.key === 'Enter' && activeIdx >= 0 && rows[activeIdx]) {
            e.preventDefault();
            rows[activeIdx].click();
        }
    });

    function updateActive(rows) {
        rows.forEach(function(r, i) { r.classList.toggle('active', i === activeIdx); });
        if (rows[activeIdx]) rows[activeIdx].scrollIntoView({ block: 'nearest' });
    }

    function doSearch(q) {
        var role = window._drSelectedRole;
        var url = '/api/employees/search?q=' + encodeURIComponent(q);
        // Map popup roles to DB role values for filtering
        if (role && role !== 'OTHER') {
            url += '&role=' + encodeURIComponent(role);
        }
        fetch(url).then(function(r) { return r.json(); }).then(function(data) {
            activeIdx = -1;
            if (!data || !data.length) {
                resultsDiv.innerHTML = '<div class="dr-search-empty">No employees found</div>';
                resultsDiv.className = 'dr-search-results open';
                return;
            }
            var html = '';
            data.slice(0, 30).forEach(function(emp) {
                html += '<div class="dr-emp-row" data-emp=\'' + JSON.stringify(emp).replace(/'/g, '&#39;') + '\'>'
                    + '<div class="dr-emp-id">' + esc(emp.emp_id) + ' — ' + esc(emp.full_name) + '</div>'
                    + '<div class="dr-emp-name">' + esc(emp.designation || '') + ' &middot; ' + esc(emp.branch_name || '') + '</div>'
                    + '</div>';
            });
            resultsDiv.innerHTML = html;
            resultsDiv.className = 'dr-search-results open';

            // Click handlers
            resultsDiv.querySelectorAll('.dr-emp-row').forEach(function(row) {
                row.addEventListener('click', function() {
                    var emp = JSON.parse(this.getAttribute('data-emp'));
                    selectEmployee(emp);
                });
            });
        }).catch(function(err) {
            console.warn('Employee search failed:', err);
            resultsDiv.innerHTML = '<div class="dr-search-empty">Search failed</div>';
            resultsDiv.className = 'dr-search-results open';
        });
    }

    function esc(s) { return String(s || '').replace(/[&<>"']/g, function(c) { return '&#' + c.charCodeAt(0) + ';'; }); }

    function selectEmployee(emp) {
        window._drSelectedEmployee = emp;

        // Hide search results
        resultsDiv.className = 'dr-search-results';
        searchInput.value = '';

        // Show selected chip
        var chip = document.getElementById('drSelectedEmp');
        chip.innerHTML = '<span>' + esc(emp.emp_id) + ' — ' + esc(emp.full_name) + '</span>'
            + '<span class="dr-clear" onclick="window._drClearEmployee()">&times;</span>';
        chip.style.display = '';

        // Show login button
        var loginBtn = document.getElementById('login-btn-dm');
        var roleLabel = { RM: 'Regional Manager', DM: 'Division Manager', AM: 'Area Manager', BM: 'Branch Manager', OTHER: 'Employee' };
        loginBtn.textContent = 'Login as ' + (roleLabel[window._drSelectedRole] || window._drSelectedRole);
        loginBtn.style.display = '';

        // Also set dmSelect value for backward compat
        var locationMap = { RM: emp.region_name, DM: emp.division_name, AM: emp.area_name, BM: emp.branch_name, OTHER: emp.branch_name };
        var loc = locationMap[window._drSelectedRole] || emp.branch_name || '';
        document.getElementById('dmSelect').value = loc;
    }

    window._drClearEmployee = function() {
        window._drSelectedEmployee = null;
        document.getElementById('drSelectedEmp').style.display = 'none';
        document.getElementById('login-btn-dm').style.display = 'none';
        searchInput.focus();
    };

    // Close dropdown when clicking outside
    document.addEventListener('click', function(e) {
        if (!searchInput.contains(e.target) && !resultsDiv.contains(e.target)) {
            resultsDiv.className = 'dr-search-results';
        }
    });
}

function getDMNode() {
    // Find the tree node for the current user based on role
    // RM: full tree, DM: division node, AM: area node, BM: single branch
    if (!state.currentUser) return null;
    for (const region in state.fullTree) {
        // DM level: match division
        if (state.fullTree[region][state.currentUser]) {
            return state.fullTree[region][state.currentUser];
        }
        // AM level: look inside divisions for area match
        if (state.role === 'AM') {
            for (const div in state.fullTree[region]) {
                if (state.fullTree[region][div][state.currentUser]) {
                    const node = {};
                    node[state.currentUser] = state.fullTree[region][div][state.currentUser];
                    return node;
                }
            }
        }
        // BM: find branch in any area
        if (state.role === 'BM') {
            for (const div in state.fullTree[region]) {
                for (const area in state.fullTree[region][div]) {
                    if (state.fullTree[region][div][area].has && state.fullTree[region][div][area].has(state.currentUser)) {
                        const node = {};
                        node[area] = new Set([state.currentUser]);
                        return node;
                    }
                }
            }
        }
    }
    return null;
}

function savePlans() {
    const inputs = document.querySelectorAll("input[data-branch]");
    inputs.forEach(i => {
        state.targetData[i.dataset.branch] = parseInt(i.value) || 0;
    });
    localStorage.setItem('dosa_targets', JSON.stringify(state.targetData));
    showToast("Plans Saved Successfully!", "check");
}

function calculateCEOStats() {
    let totalBranches = 0;
    let completedBranches = 0;
    let pendingAchievement = 0;
    let onlyPlans = 0; // Branches with only plans
    let noDataBranches = 0;

    // Collection Metrics
    let ftodPlan = 0, ftodAchieve = 0;
    let dpd130Plan = 0, dpd130Achieve = 0;
    let dpd3160Plan = 0, dpd3160Achieve = 0;
    let dpd6190Plan = 0, dpd6190Achieve = 0;
    let npaActivation = 0, npaClosure = 0; // Actuals
    let npaActivationPlan = 0, npaClosurePlan = 0; // Plans

    // Disbursement
    let disbIglAccPlan = 0, disbIglAmtPlan = 0; // Plans from daily_reports
    let disbFigAccPlan = 0, disbFigAmtPlan = 0;
    let disbIlAccPlan = 0, disbIlAmtPlan = 0;

    let disbIglAccAchieve = 0, disbIglAmtAchieve = 0; // Actuals from daily_reports_achievements
    let disbFigAccAchieve = 0, disbFigAmtAchieve = 0;
    let disbIlAccAchieve = 0, disbIlAmtAchieve = 0;

    let disbSancAcc = 0, disbSancAmt = 0; // Sanction Pending (if applicable)

    // KYC
    let kycIgl = 0, kycFig = 0, kycIl = 0; // Actuals
    let kycIglPlan = 0, kycFigPlan = 0, kycIlPlan = 0; // Plans

    // Portfolio Health
    let portfolioHealth = { healthy: 0, dpd130: 0, dpd3160: 0, dpd6190: 0 };

    // Regional breakdown
    let regionStats = {};
    let regionalBreakdown = {};

    // Achievement percentages
    let allAchievementPcts = [];

    const safeInt = (v) => parseInt(v) || 0;
    const safeFloat = (v) => parseFloat(v) || 0;

    state.rawData.rows.forEach(row => {
        const branch = row[0];
        const region = row[3];

        if (branch) {
            totalBranches++;

            // Initialize region
            if (!regionStats[region]) regionStats[region] = { branches: 0, completed: 0, plansSet: 0, totalPct: 0, count: 0 };
            regionStats[region].branches++;

            const branchEntry = state.branchDetails[branch];
            const hasPlan = branchEntry && branchEntry.target;
            const hasAchieve = branchEntry && branchEntry.achievement;

            if (hasPlan && hasAchieve) completedBranches++;
            if (hasPlan) regionStats[region].plansSet++;
            if (hasPlan && !hasAchieve) onlyPlans++;
            if (hasPlan && hasAchieve) regionStats[region].completed++;
            if (!hasPlan && !hasAchieve) noDataBranches++;


            // --- AGGREGATE PLAN DATA ---
            if (hasPlan) {
                const t = branchEntry.target;
                ftodPlan += safeInt(t.ftod_plan);
                dpd130Plan += safeInt(t.dpd_1_30_plan);
                dpd3160Plan += safeInt(t.dpd_31_60_plan);
                dpd6190Plan += safeInt(t.dpd_61_90_plan);
                npaActivationPlan += safeInt(t.npa_activation);
                npaClosurePlan += safeInt(t.npa_closure);

                // Disbursement Plans (from daily_reports)
                disbIglAccPlan += safeInt(t.disb_igl_acc);
                disbIglAmtPlan += safeFloat(t.disb_igl_amt);
                disbFigAccPlan += safeInt(t.disb_fig_acc);
                disbFigAmtPlan += safeFloat(t.disb_fig_amt);
                disbIlAccPlan += safeInt(t.disb_il_acc);
                disbIlAmtPlan += safeFloat(t.disb_il_amt);

                // KYC Plans
                kycIglPlan += safeInt(t.kyc_igl);
                kycFigPlan += safeInt(t.kyc_fig);
                kycIlPlan += safeInt(t.kyc_il);
            }

            // --- AGGREGATE ACHIEVEMENT DATA ---
            if (hasAchieve) {
                const a = branchEntry.achievement;

                ftodAchieve += safeInt(a.ftod_actual);
                dpd130Achieve += safeInt(a.dpd_1_30_actual);
                dpd3160Achieve += safeInt(a.dpd_31_60_actual);
                dpd6190Achieve += safeInt(a.dpd_61_90_actual);
                npaActivation += safeInt(a.npa_activation);
                npaClosure += safeInt(a.npa_closure);

                // Disbursement Actuals
                disbIglAccAchieve += safeInt(a.disb_igl_acc);
                disbIglAmtAchieve += safeFloat(a.disb_igl_amt);
                disbFigAccAchieve += safeInt(a.disb_fig_acc);
                disbFigAmtAchieve += safeFloat(a.disb_fig_amt);
                disbIlAccAchieve += safeInt(a.disb_il_acc);
                disbIlAmtAchieve += safeFloat(a.disb_il_amt);

                // KYC
                kycIgl += safeInt(a.kyc_igl);
                kycFig += safeInt(a.kyc_fig);
                kycIl += safeInt(a.kyc_il);

                // Portfolio Health (Actuals)
                portfolioHealth.healthy += safeInt(a.ftod_actual);
                portfolioHealth.dpd130 += safeInt(a.dpd_1_30_actual);
                portfolioHealth.dpd3160 += safeInt(a.dpd_31_60_actual);
                portfolioHealth.dpd6190 += safeInt(a.dpd_61_90_actual);
            }

            // --- CALCULATE PERCENTAGES (Only if both exist) ---
            if (hasPlan && hasAchieve) {
                const branchPct = calculateBranchAveragePercentage(branch);
                if (branchPct !== null && branchPct > 0) {
                    allAchievementPcts.push(branchPct);
                    regionStats[region].totalPct += branchPct;
                    regionStats[region].count++;
                }
            }
        }
    });

    // Build regional breakdown for table
    Object.keys(regionStats).sort().forEach(region => {
        const r = regionStats[region];
        regionalBreakdown[region] = {
            branches: r.branches,
            plansSet: r.plansSet,
            completed: r.completed,
            avgPct: r.count > 0 ? Math.round(r.totalPct / r.count) : 0
        };
    });

    // Calculate overall averages
    const completionRate = totalBranches > 0 ? Math.round((completedBranches / totalBranches) * 100) : 0;
    const plansRate = totalBranches > 0 ? Math.round(((totalBranches - noDataBranches) / totalBranches) * 100) : 0;

    const avgAchievementPct = allAchievementPcts.length > 0
        ? Math.round(allAchievementPcts.reduce((a, b) => a + b, 0) / allAchievementPcts.length)
        : 0;

    const totalCollectionPlan = ftodPlan + dpd130Plan + dpd3160Plan + dpd6190Plan;

    // NPA Totals
    const npaTotalPlan = npaActivationPlan + npaClosurePlan;

    // Disbursement Totals
    const totalDisbursementPlan = disbIglAmtPlan + disbFigAmtPlan + disbIlAmtPlan;
    const totalDisbursementAchieve = disbIglAmtAchieve + disbFigAmtAchieve + disbIlAmtAchieve;

    // Total Disbursement Accounts
    const totalDisbAccPlan = disbIglAccPlan + disbFigAccPlan + disbIlAccPlan;
    const totalDisbAccAchieve = disbIglAccAchieve + disbFigAccAchieve + disbIlAccAchieve;

    // KYC Total (Plan)
    const kycTotal = kycIglPlan + kycFigPlan + kycIlPlan;

    return {
        totalBranches,
        completedBranches,
        onlyPlans,
        noDataBranches,
        completionRate,
        plansRate,
        avgAchievementPct,
        totalCollectionPlan,

        // Collection
        ftodPlan, ftodAchieve,
        dpd130Plan, dpd130Achieve,
        dpd3160Plan, dpd3160Achieve,
        dpd6190Plan, dpd6190Achieve,
        npaActivation, npaClosure,
        npaActivationPlan, npaClosurePlan,
        npaTotalPlan,

        // Disbursement
        totalDisbursementPlan, totalDisbursementAchieve,
        totalDisbAccPlan, totalDisbAccAchieve,
        disbIglAccPlan, disbIglAmtPlan,
        disbFigAccPlan, disbFigAmtPlan,
        disbIlAccPlan, disbIlAmtPlan,
        disbIglAccAchieve, disbIglAmtAchieve,
        disbFigAccAchieve, disbFigAmtAchieve,
        disbIlAccAchieve, disbIlAmtAchieve,

        // KYC
        kycIgl, kycFig, kycIl, kycTotal,
        kycIglPlan, kycFigPlan, kycIlPlan,

        // Portfolio
        portfolioHealth,

        // Regional
        regionStats,
        regionalBreakdown
    };
}

// ============ CEO DETAIL MODAL FUNCTIONS ============

// Add Escape key listener for closing detail modal
document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
        const detailModal = document.getElementById('detailModal');
        if (detailModal && detailModal.classList.contains('visible')) {
            closeDetailModal();
        }
    }
});

function openDetailModal(type, data = null) {
    const modal = document.getElementById('detailModal');
    const body = document.getElementById('detailModalBody');
    const titleText = document.getElementById('detailTitleText');
    const titleIcon = document.getElementById('detailTitleIcon');
    const subtitle = document.getElementById('detailSubtitle');

    // Set title based on type
    const config = {
        'completion': { icon: '✅', title: 'Completion Rate Analysis', subtitle: 'Branch-wise completion status' },
        'branches': { icon: '🏢', title: 'Active Branches', subtitle: 'All branches by region' },
        'achievement': { icon: '📈', title: 'Achievement Analysis', subtitle: 'Performance ranking' },
        'collection': { icon: '💰', title: 'Collection Plan Breakdown', subtitle: 'FTOD + 1-30 + 31-60 + 61-90 DPD' },
        'ftod': { icon: '🎯', title: 'FTOD Collection Details', subtitle: 'First Time Overdue' },
        'dpd_1_30': { icon: '⚠️', title: '1-30 DPD Details', subtitle: '1-30 Days Past Due accounts' },
        'dpd_31_60': { icon: '⚠️', title: '31-60 DPD Details', subtitle: '31-60 Days Past Due accounts' },
        'dpd_61_90': { icon: '📊', title: '61-90 DPD Details', subtitle: '61-90 Days Past Due accounts' },
        'npa': { icon: '🔴', title: 'NPA Movement Details', subtitle: 'Activation vs Closure' },
        'portfolio': { icon: '📉', title: 'Portfolio Health Details', subtitle: 'Account distribution' },
        'disbursement': { icon: '💳', title: 'Disbursement Details', subtitle: 'Product-wise breakdown' },
        'disb_igl': { icon: '🌱', title: 'IGL Details', subtitle: 'Individual Group loans' },
        'disb_fig': { icon: '🌾', title: 'FIG Details', subtitle: 'Financial Inclusion Group loans' },
        'disb_il': { icon: '🏠', title: 'IL Details', subtitle: 'Individual loans' },
        'region': { icon: '🌍', title: `Region: ${data || 'All'}`, subtitle: 'Regional performance' },
        'pending': { icon: '⏳', title: 'Pending Achievement', subtitle: 'Branches with only targets' },
        'nodata': { icon: '❌', title: 'No Data Branches', subtitle: 'Branches without reports' },
        'completed': { icon: '✓', title: 'Fully Completed', subtitle: 'Branches with full data' },
        'kyc': { icon: '📋', title: 'KYC Sourcing Details', subtitle: 'Customer acquisition' }
    };

    const cfg = config[type] || { icon: '📊', title: 'Details', subtitle: '' };
    titleIcon.textContent = cfg.icon;
    titleText.textContent = cfg.title;
    subtitle.textContent = cfg.subtitle;

    // Render content based on type
    body.innerHTML = renderDetailContent(type, data);

    // Show modal
    modal.classList.add('visible');
}

function closeDetailModal() {
    document.getElementById('detailModal').classList.remove('visible');
}

const mockDashboardData = {
    "(All)": {
        paid: "90.44", paidSub: "280.6",
        par: "1,499", parSub: "7,318",
        res: "32.2%",
        eff: "20.5%",
        buckets: [
            { label: "1 to 30 days", books: "714", paid: "29.16", paidPerc: 44.7, resPerc: 67.6 },
            { label: "31 to 60 days", books: "645", paid: "24.14", paidPerc: 40.1, resPerc: 59.2 },
            { label: "61 to 90 days", books: "812", paid: "29.53", paidPerc: 46.7, resPerc: 58.9 },
            { label: "NPA", books: "5,147", paid: "7.62", paidPerc: 14.9, resPerc: 5.1 }
        ],
        bubbles: [
            { x: 20, y: 85, val: 85, label: "JabalpurCluster", color: "bubble-pink" },
            { x: 60, y: 70, val: 70, label: "IndoreCluster", color: "bubble-orange" },
            { x: 90, y: 30, val: 30, label: "SalemCluster", color: "bubble-yellow" },
            { x: 15, y: 70, val: "", label: "DehradunCluster", color: "bubble-pink" },
            { x: 50, y: 65, val: "", label: "AshtaCluster", color: "bubble-pink" },
            { x: 55, y: 50, val: "", label: "HaryanaCluster", color: "bubble-orange" }
        ]
    },
    "Region 1": {
        paid: "45.12", paidSub: "120.4",
        par: "850", parSub: "3,100",
        res: "38.5%",
        eff: "24.1%",
        buckets: [
            { label: "1 to 30 days", books: "350", paid: "15.00", paidPerc: 50.0, resPerc: 72.0 },
            { label: "31 to 60 days", books: "300", paid: "12.00", paidPerc: 45.0, resPerc: 65.0 },
            { label: "61 to 90 days", books: "400", paid: "14.00", paidPerc: 48.0, resPerc: 62.0 },
            { label: "NPA", books: "2,500", paid: "4.12", paidPerc: 18.0, resPerc: 8.0 }
        ],
        bubbles: [
            { x: 30, y: 80, val: 80, label: "DelhiCluster", color: "bubble-pink" },
            { x: 25, y: 60, val: 60, label: "JaipurCluster", color: "bubble-orange" },
            { x: 40, y: 40, val: 40, label: "LucknowCluster", color: "bubble-yellow" }
        ]
    },
    "Region 2": {
        paid: "30.15", paidSub: "100.2",
        par: "450", parSub: "2,500",
        res: "28.1%",
        eff: "18.4%",
        buckets: [
            { label: "1 to 30 days", books: "250", paid: "10.00", paidPerc: 40.0, resPerc: 60.0 },
            { label: "31 to 60 days", books: "200", paid: "8.00", paidPerc: 35.0, resPerc: 55.0 },
            { label: "61 to 90 days", books: "300", paid: "10.00", paidPerc: 42.0, resPerc: 52.0 },
            { label: "NPA", books: "1,800", paid: "2.15", paidPerc: 12.0, resPerc: 4.0 }
        ],
        bubbles: [
            { x: 80, y: 30, val: 30, label: "ChennaiCluster", color: "bubble-yellow" },
            { x: 70, y: 40, val: 40, label: "BangaloreCluster", color: "bubble-orange" },
            { x: 85, y: 20, val: 20, label: "KeralaCluster", color: "bubble-pink" }
        ]
    }
};

function showLastMonthSummary() {
    const modal = document.getElementById('lastMonthSummaryModal');
    const body = document.getElementById('lastMonthSummaryBody');
    if (modal && body) {
        modal.classList.add('visible');
        renderLastMonthDashboard(body, "(All)");
    }
}

function handleSummaryFilterChange(event, container) {
    const selectedRegion = event.target.value;
    renderLastMonthDashboard(container, selectedRegion);
}

function renderLastMonthDashboard(container, region) {
    const data = mockDashboardData[region] || mockDashboardData["(All)"];

    // Generate Bubbles HTML
    const bubblesHtml = data.bubbles.map(b => `
        <div class="scatter-bubble ${b.color} animate-enter" 
             style="left: ${b.x}%; bottom: ${b.y}%; width: ${b.val ? 30 + (b.val / 5) : 20}px; height: ${b.val ? 30 + (b.val / 5) : 20}px;">
            ${b.val}
            <div class="chart-tooltip">
                <strong>${b.label}</strong><br>
                NPA: ${b.y}%<br>
                Eff: ${b.x}%
            </div>
        </div>
        <div class="scatter-label animate-enter" style="left: ${b.x}%; bottom: ${b.y}%; transform: translate(-50%, -${b.val ? 130 : 50}%);">
            ${b.label}
        </div>
    `).join('');

    // Generate Bucket Table Rows HTML
    const bucketRowsHtml = data.buckets.map(row => `
        <tr>
            <td class="s-bucket-col">${row.label}</td>
            <td>${row.books}</td>
            <td>
                <div class="s-progress-wrapper" style="gap:6px;">
                    <span>${row.paid}</span>
                    <div class="s-progress-track" style="width:50px; height:10px;">
                        <div class="s-progress-fill fill-gray" style="width:${row.paidPerc}%;"></div>
                    </div>
                    <span class="s-val-label" style="font-size:10px;">${row.paidPerc}</span>
                </div>
            </td>
            <td>
                 <div class="s-progress-wrapper" style="gap:6px;">
                    <div class="s-progress-track" style="height:10px;">
                        <div class="s-progress-fill fill-teal" style="width:${row.resPerc}%;"></div>
                    </div>
                    <span class="s-val-label" style="font-size:10px;">${row.resPerc}%</span>
                </div>
            </td>
        </tr>
    `).join('');

    container.innerHTML = `
        <div class="summary-dashboard" style="gap:16px; padding:20px;">
            <!-- Filter Bar -->
            <div class="summary-filter-bar" style="padding: 10px 16px;">
                <div style="flex:1;">
                    <label style="display:block; font-size:11px; color:#6B7280; margin-bottom:4px;">Region</label>
                    <select class="s-select" id="regionSelect" style="padding:6px 10px;">
                        <option value="(All)" ${region === '(All)' ? 'selected' : ''}>(All)</option>
                        <option value="Region 1" ${region === 'Region 1' ? 'selected' : ''}>Region 1</option>
                        <option value="Region 2" ${region === 'Region 2' ? 'selected' : ''}>Region 2</option>
                    </select>
                </div>
                <div style="flex:1;">
                    <label style="display:block; font-size:11px; color:#6B7280; margin-bottom:4px;">District</label>
                    <select class="s-select" style="padding:6px 10px;">
                        <option>(All)</option>
                        <option>Indore</option>
                        <option>Bhopal</option>
                        <option>Chennai</option>
                        <option>Jabalpur</option>
                    </select>
                </div>
            </div>

            <!-- Metrics Row -->
            <div class="summary-metrics-row animate-enter delay-1" style="gap:16px;">
                <!-- Metric 1 -->
                <div class="s-metric-card" style="padding:16px; gap:12px;">
                    <div class="s-metric-icon" style="width:40px; height:40px; font-size:18px;">
                        <svg class="icon" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle><path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8"></path><path d="M12 18V6"></path></svg>
                    </div>
                    <div class="s-metric-content">
                        <div class="s-metric-value" style="font-size:18px;">${data.paid} <span style="font-size:12px; font-weight:500;">(in Cr.)</span></div>
                        <div class="s-metric-label" style="font-size:11px;">Amount Paid</div>
                        <div class="s-metric-sub" style="font-size:10px;">${data.paidSub} <span style="color:#6B7280;">Opening POS</span></div>
                    </div>
                </div>

                <!-- Metric 2 -->
                <div class="s-metric-card" style="padding:16px; gap:12px;">
                    <div class="s-metric-icon" style="width:40px; height:40px; font-size:18px;">
                         <svg class="icon" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                    </div>
                    <div class="s-metric-content">
                         <div class="s-metric-value" style="font-size:18px;">${data.par}</div>
                         <div class="s-metric-label" style="font-size:11px;">PAR Resolution</div>
                         <div class="s-metric-sub" style="font-size:10px;">${data.parSub} <span style="color:#6B7280;">(# of Cases)</span></div>
                    </div>
                </div>

                <!-- Metric 3 -->
                <div class="s-metric-card" style="padding:16px; gap:12px;">
                    <div class="s-metric-icon" style="width:40px; height:40px; font-size:18px;">
                         <svg class="icon" viewBox="0 0 24 24"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path></svg>
                    </div>
                    <div class="s-metric-content">
                         <div class="s-metric-value" style="font-size:18px;">${data.res} <span style="font-size:12px; color:#10B981;">($)</span></div>
                         <div class="s-metric-label" style="font-size:11px;">Resolution %</div>
                    </div>
                </div>

                <!-- Metric 4 -->
                <div class="s-metric-card" style="padding:16px; gap:12px;">
                    <div class="s-metric-icon" style="width:40px; height:40px; font-size:18px;">
                        <svg class="icon" viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"></path><line x1="3" y1="6" x2="21" y2="6"></line><path d="M16 10a4 4 0 0 1-8 0"></path></svg>
                    </div>
                    <div class="s-metric-content">
                         <div class="s-metric-value" style="font-size:18px;">${data.eff} <span style="font-size:12px; color:#3B82F6;">(#)</span></div>
                         <div class="s-metric-label" style="font-size:11px;">Efficiency %</div>
                    </div>
                </div>
            </div>

            <!-- Table 1 -->
            <div class="s-widget animate-enter" style="padding:16px; margin-bottom: 16px;">
                 <div class="s-widget-header" style="font-size:13px; margin-bottom:12px;">Table 1</div>
                 <div class="s-table-container">
                    <table class="s-table" style="font-size:11px;">
                        <thead>
                            <tr>
                                <th></th>
                                <th></th>
                                <th></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td colspan="3" style="text-align:center; padding: 20px; color: var(--text-secondary);">
                                    
                                </td>
                            </tr>
                        </tbody>
                    </table>
                 </div>
            </div>

            <!-- Content Grid -->
            <div class="summary-details-grid" style="gap:16px;">
                
                <!-- Widget 1: Bucket Wise Performance -->
                <div class="s-widget animate-enter delay-2" style="padding:16px;">
                    <div class="s-widget-header" style="font-size:13px; margin-bottom:12px;">Bucket Wise Performance</div>
                    <div class="s-table-container">
                        <table class="s-table" style="font-size:11px;">
                            <thead>
                                <tr>
                                    <th>Opening Bucket</th>
                                    <th>Opening Books</th>
                                    <th>Total Amount Paid</th>
                                    <th>% Resolution</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${bucketRowsHtml}
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Widget 2: Bucket-Status wise Opening POS -->
                <div class="s-widget animate-enter delay-2" style="padding:16px;">
                    <div class="s-widget-header" style="font-size:13px; margin-bottom:12px;">Bucket-Status wise Opening POS (in Crs.)</div>
                     <div class="s-table-container">
                        <table class="s-table" style="font-size:11px;">
                            <thead>
                                <tr>
                                    <th rowspan="2">Opening Bucket</th>
                                    <th colspan="3" style="text-align:center;">Sub Status</th>
                                    <th rowspan="2">Bucket Mix</th>
                                </tr>
                                <tr>
                                    <th style="font-size:10px;">Forward Flow</th>
                                    <th style="font-size:10px;">Normalized</th>
                                    <th style="font-size:10px;">Rolled Back</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td class="s-bucket-col">1 to 30 days</td>
                                    <td><span class="badge-red">15.6</span></td>
                                    <td><span class="badge-green">1.3</span></td>
                                    <td><span class="badge-gray">27.9</span></td>
                                    <td>
                                        <div style="display:flex; height:14px; width:100%; border-radius:2px; overflow:hidden;">
                                            <div style="background:#9CA3AF; width:62%"></div>
                                            <div style="background:#EF4444; width:38%"></div>
                                        </div>
                                        <div style="display:flex; justify-content:space-between; font-size:9px; color:#6B7280; margin-top:2px;">
                                            <span>62.3%</span><span>34.8%</span>
                                        </div>
                                    </td>
                                </tr>
                                <!-- Keeps other static rows for simplicity in demo or can assume generic -->
                                <tr>
                                    <td class="s-bucket-col">31 to 60 days</td>
                                    <td><span class="badge-red">16.0</span></td>
                                    <td><span class="badge-green">0.8</span></td>
                                    <td><span class="badge-gray">23.1</span></td>
                                    <td>
                                         <div style="display:flex; height:14px; width:100%; border-radius:2px; overflow:hidden;">
                                            <div style="background:#9CA3AF; width:57%"></div>
                                            <div style="background:#EF4444; width:43%"></div>
                                        </div>
                                         <div style="display:flex; justify-content:space-between; font-size:9px; color:#6B7280; margin-top:2px;">
                                            <span>57.4%</span><span>39.9%</span>
                                        </div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Widget 3: State wise Performance -->
                 <div class="s-widget animate-enter delay-3" style="padding:16px;">
                    <div class="s-widget-header" style="font-size:13px; margin-bottom:12px;">State wise Performance</div>
                    <div class="s-table-container">
                        <table class="s-table" style="font-size:11px;">
                            <thead>
                                <tr>
                                    <th>State</th>
                                    <th>Opening Books</th>
                                    <th>Total Amount Paid</th>
                                    <th>% Resolution</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td class="s-bucket-col">Tamil Nadu</td>
                                    <td>1,808</td>
                                    <td>32.97</td>
                                    <td>
                                         <div class="s-progress-wrapper" style="gap:6px;">
                                            <div class="s-progress-track" style="height:10px;"><div class="s-progress-fill fill-teal" style="width:34%;"></div></div>
                                            <span class="s-val-label" style="font-size:10px;">34.6%</span>
                                        </div>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="s-bucket-col">Maharashtra</td>
                                    <td>1,016</td>
                                    <td>11.93</td>
                                    <td>
                                         <div class="s-progress-wrapper" style="gap:6px;">
                                            <div class="s-progress-track" style="height:10px;"><div class="s-progress-fill fill-teal" style="width:30%;"></div></div>
                                            <span class="s-val-label" style="font-size:10px;">30.4%</span>
                                        </div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                 <!-- Widget 4: Bubble Chart -->
                <div class="s-widget animate-enter delay-3" style="padding:16px; min-height:300px;">
                    <div class="s-widget-header" style="font-size:13px; margin-bottom:12px;">Cluster Wise Performance</div>
                    
                    <div class="scatter-chart-container">
                        <!-- Y-Axis Label -->
                        <div class="scatter-axis-label-y">NPA %</div>
                        <!-- X-Axis Label -->
                        <div class="scatter-axis-label-x">Selected Benchmark X Axis : 50</div>

                        <!-- Bubbles -->
                        ${bubblesHtml}
                    </div>
                </div>

            </div>
        </div>
    `;

    // Attach Event Listener
    const select = document.getElementById('regionSelect');
    if (select) {
        select.addEventListener('change', (e) => handleSummaryFilterChange(e, container));
    }
}

function closeLastMonthSummary() {
    const modal = document.getElementById('lastMonthSummaryModal');
    if (modal) modal.classList.remove('visible');
}

// Helper function to open branch modal from detail view
// Closes detail modal first so branch modal appears on top
function openBranchFromDetail(branchName) {
    closeDetailModal();
    // Small delay to ensure detail modal is closed before opening branch modal
    setTimeout(() => {
        openBranchModal(branchName);
    }, 150);
}

function renderDetailContent(type, data) {
    const stats = calculateCEOStats();

    switch (type) {
        case 'completion': return renderCompletionDetail(stats);
        case 'branches': return renderBranchesDetail(stats);
        case 'achievement': return renderAchievementDetail(stats);
        case 'collection': return renderCollectionDetail(stats);
        case 'ftod': return renderMetricDetail('ftod', stats);
        case 'dpd_1_30': return renderMetricDetail('dpd_1_30', stats);
        case 'dpd_31_60': return renderMetricDetail('dpd_31_60', stats);
        case 'dpd_61_90': return renderMetricDetail('dpd_61_90', stats);
        case 'npa': return renderMetricDetail('npa', stats);
        case 'portfolio': return renderPortfolioDetailView(data, stats);
        case 'disbursement': return renderDisbursementDetailView(stats);
        case 'disb_sanction': return renderDisbProductDetail('sanction', stats);
        case 'disb_igl': return renderDisbProductDetail('igl', stats);
        case 'disb_fig': return renderDisbProductDetail('fig', stats);
        case 'disb_il': return renderDisbProductDetail('il', stats);
        case 'region': return renderRegionDetailView(data, stats);
        case 'pending': return renderInsightDetail('pending', stats);
        case 'nodata': return renderInsightDetail('nodata', stats);
        case 'completed': return renderInsightDetail('completed', stats);
        case 'kyc': return renderKYCDetail(stats);
        default: return '<p>No data available</p>';
    }
}

// --- Completion Rate Detail ---
function renderCompletionDetail(stats) {
    const branches = getBranchListByStatus();

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.completedBranches}</div>
                        <div class="detail-stat-label">Completed</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#F59E0B;">${stats.pendingAchievement}</div>
                        <div class="detail-stat-label">Pending Achievement</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#EF4444;">${stats.noDataBranches}</div>
                        <div class="detail-stat-label">No Data</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:var(--primary-accent);">${stats.completionRate}%</div>
                        <div class="detail-stat-label">Overall Rate</div>
                    </div>
                </div>
                
                <div class="detail-filter-bar">
                    <button class="detail-filter-btn active" onclick="filterBranchList('all')">All (${stats.totalBranches})</button>
                    <button class="detail-filter-btn" onclick="filterBranchList('completed')">✅ Completed (${stats.completedBranches})</button>
                    <button class="detail-filter-btn" onclick="filterBranchList('pending')">⏳ Pending (${stats.pendingAchievement})</button>
                    <button class="detail-filter-btn" onclick="filterBranchList('nodata')">❌ No Data (${stats.noDataBranches})</button>
                </div>
                
                <div class="detail-branch-list" id="branchListContainer">
                    <div class="detail-branch-header">
                        <div>Branch</div>
                        <div>Region</div>
                        <div>DM</div>
                        <div>Status</div>
                        <div>Achievement %</div>
                        <div>Action</div>
                    </div>
                    ${branches.map(b => `
                        <div class="detail-branch-row" data-status="${b.status}">
                            <div>
                                <div class="detail-branch-name">${b.name}</div>
                                <div class="detail-branch-region">${b.district}</div>
                            </div>
                            <div>${b.region}</div>
                            <div>${b.dm}</div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.statusColor}20; color:${b.statusColor};">
                                    ${b.statusLabel}
                                </span>
                            </div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.pctColor}20; color:${b.pctColor};">
                                    ${b.pct}
                                </span>
                            </div>
                            <div>
                                <button onclick="openBranchFromDetail('${b.name}')" style="padding:6px 12px; border-radius:6px; border:1px solid var(--border-color); background:var(--bg-card); cursor:pointer; font-size:11px;">
                                    View
                                </button>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Branches by Region Detail ---
function renderBranchesDetail(stats) {
    const regionData = getRegionBreakdown();

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${stats.totalBranches}</div>
                        <div class="detail-stat-label">Total Branches</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${Object.keys(stats.regionStats).length}</div>
                        <div class="detail-stat-label">Regions</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${getUniqueDMs()}</div>
                        <div class="detail-stat-label">Division Managers</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${getUniqueDistricts()}</div>
                        <div class="detail-stat-label">Areas</div>
                    </div>
                </div>
                
                <div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap:20px;">
                    ${Object.keys(regionData).sort().map(region => `
                        <div class="detail-stat-card clickable-section" onclick="openDetailModal('region', '${region}')" style="text-align:left; cursor:pointer;">
                            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                                <div style="font-size:16px; font-weight:700; color:var(--text-primary);">🌍 ${region}</div>
                                <span style="font-size:12px; color:var(--text-secondary);">${regionData[region].branches} branches</span>
                            </div>
                            <div style="height:6px; background:var(--border-color); border-radius:3px; overflow:hidden;">
                                <div style="height:100%; width:${regionData[region].completionPct}%; background:var(--primary-accent);"></div>
                            </div>
                            <div style="display:flex; justify-content:space-between; margin-top:8px; font-size:11px; color:var(--text-secondary);">
                                <span>${regionData[region].completed} completed</span>
                                <span style="font-weight:600; color:var(--primary-accent);">${regionData[region].completionPct}%</span>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Achievement Ranking Detail ---
function renderAchievementDetail(stats) {
    const rankedBranches = getRankedBranches();

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.avgAchievementPct}%</div>
                        <div class="detail-stat-label">Average Achievement</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${rankedBranches.filter(b => b.pct >= 100).length}</div>
                        <div class="detail-stat-label">Above 100%</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#F59E0B;">${rankedBranches.filter(b => b.pct >= 50 && b.pct < 100).length}</div>
                        <div class="detail-stat-label">50-99%</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#EF4444;">${rankedBranches.filter(b => b.pct > 0 && b.pct < 50).length}</div>
                        <div class="detail-stat-label">Below 50%</div>
                    </div>
                </div>
                
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:24px;">
                    <div class="detail-branch-list">
                        <div style="padding:16px 24px; background:linear-gradient(135deg, #10B981 0%, #059669 100%); color:white; font-weight:700;">
                            🏆 Top 10 Performers
                        </div>
                        ${rankedBranches.slice(0, 10).map((b, i) => `
                            <div class="detail-branch-row" style="grid-template-columns: 40px 2fr 1fr;">
                                <div style="font-size:18px; font-weight:700; color:${i < 3 ? '#F59E0B' : 'var(--text-secondary)'};">#${i + 1}</div>
                                <div>
                                    <div class="detail-branch-name">${b.name}</div>
                                    <div class="detail-branch-region">${b.region}</div>
                                </div>
                                <div>
                                    <span class="detail-pct-badge" style="background:#10B98120; color:#10B981;">${b.pct}%</span>
                                </div>
                            </div>
                        `).join('')}
                    </div>
                    
                    <div class="detail-branch-list">
                        <div style="padding:16px 24px; background:linear-gradient(135deg, #EF4444 0%, #DC2626 100%); color:white; font-weight:700;">
                            ⚠️ Bottom 10 - Need Attention
                        </div>
                        ${rankedBranches.slice(-10).reverse().map((b, i) => `
                            <div class="detail-branch-row" style="grid-template-columns: 40px 2fr 1fr;">
                                <div style="font-size:14px; color:#EF4444;">📉</div>
                                <div>
                                    <div class="detail-branch-name">${b.name}</div>
                                    <div class="detail-branch-region">${b.region}</div>
                                </div>
                                <div>
                                    <span class="detail-pct-badge" style="background:#EF444420; color:#EF4444;">${b.pct}%</span>
                                </div>
                            </div>
                        `).join('')}
                    </div>
                </div>
            `;
}

// --- Collection Plan Detail ---
function renderCollectionDetail(stats) {
    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('ftod')">
                        <div class="detail-stat-value" style="color:#6366F1;">₹${(stats.ftodPlan / 100000).toFixed(1)}L</div>
                        <div class="detail-stat-label">FTOD Plan</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Click for details</div>
                    </div>
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('dpd_1_30')">
                        <div class="detail-stat-value" style="color:#F59E0B;">₹${(stats.dpd130Plan / 100000).toFixed(1)}L</div>
                        <div class="detail-stat-label">1-30 DPD Plan</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Click for details</div>
                    </div>
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('dpd_31_60')">
                        <div class="detail-stat-value" style="color:#FB923C;">₹${(stats.dpd3160Plan / 100000).toFixed(1)}L</div>
                        <div class="detail-stat-label">31-60 DPD Plan</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Click for details</div>
                    </div>
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('dpd_61_90')">
                        <div class="detail-stat-value" style="color:#EF4444;">₹${(stats.dpd6190Plan / 100000).toFixed(1)}L</div>
                        <div class="detail-stat-label">61-90 DPD Plan</div>
                        <div style="font-size:11px; color:var(--text-secondary); margin-top:4px;">Click for details</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:var(--primary-accent);">₹${(stats.totalCollectionPlan / 100000).toFixed(1)}L</div>
                        <div class="detail-stat-label">Total Plan</div>
                    </div>
                </div>

                ${renderMetricBranchTable('collection', stats)}
            `;
}

// --- Single Metric Detail (FTOD, DPD buckets, NPA) ---
function renderMetricDetail(metricType, stats) {
    const config = {
        'ftod': { label: 'FTOD', targetField: 'ftod_plan', achieveField: 'ftod_actual', target: stats.ftodPlan, achieve: stats.ftodAchieve },
        'dpd_1_30': { label: '1-30 DPD', targetField: 'dpd_1_30_plan', achieveField: 'dpd_1_30_actual', target: stats.dpd130Plan, achieve: stats.dpd130Achieve },
        'dpd_31_60': { label: '31-60 DPD', targetField: 'dpd_31_60_plan', achieveField: 'dpd_31_60_actual', target: stats.dpd3160Plan, achieve: stats.dpd3160Achieve },
        'dpd_61_90': { label: '61-90 DPD', targetField: 'dpd_61_90_plan', achieveField: 'dpd_61_90_actual', target: stats.dpd6190Plan, achieve: stats.dpd6190Achieve },
        'npa': { label: 'NPA Movement', targetField: null, achieveField: null, activation: stats.npaActivation, closure: stats.npaClosure }
    };

    const cfg = config[metricType];

    if (metricType === 'npa') {
        return renderNPADetail(stats);
    }

    const pct = cfg.target > 0 ? Math.round((cfg.achieve / cfg.target) * 100) : 0;
    const pctColor = pct >= 100 ? '#10B981' : pct >= 50 ? '#F59E0B' : '#EF4444';

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${cfg.target.toLocaleString()}</div>
                        <div class="detail-stat-label">Total Plan</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:${pctColor};">${cfg.achieve.toLocaleString()}</div>
                        <div class="detail-stat-label">Total Achievement</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:${pctColor};">${pct}%</div>
                        <div class="detail-stat-label">Achievement Rate</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:${cfg.target - cfg.achieve > 0 ? '#EF4444' : '#10B981'};">${Math.abs(cfg.target - cfg.achieve).toLocaleString()}</div>
                        <div class="detail-stat-label">${cfg.target - cfg.achieve > 0 ? 'Gap' : 'Surplus'}</div>
                    </div>
                </div>
                
                ${renderMetricBranchTable(metricType, stats)}
            `;
}

// --- NPA Detail ---
function renderNPADetail(stats) {
    const net = stats.npaActivation - stats.npaClosure;
    const netColor = net > 0 ? '#EF4444' : '#10B981';

    const branchNPA = getBranchNPAData();

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#EF4444;">${stats.npaActivation}</div>
                        <div class="detail-stat-label">NPA Activated</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.npaClosure}</div>
                        <div class="detail-stat-label">NPA Closed</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:${netColor};">${net > 0 ? '+' : ''}${net}</div>
                        <div class="detail-stat-label">Net Movement</div>
                    </div>
                    <div class="detail-stat-card">
                        <div style="font-size:40px;">${net > 0 ? '📈' : '📉'}</div>
                        <div class="detail-stat-label">${net > 0 ? 'Increasing' : 'Decreasing'}</div>
                    </div>
                </div>
                
                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>Region</div>
                        <div style="color:#EF4444;">Activated</div>
                        <div style="color:#10B981;">Closed</div>
                        <div>Net</div>
                    </div>
                    ${branchNPA.map(b => `
                        <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.region}</div>
                            <div style="font-weight:600; color:#EF4444;">${b.activation}</div>
                            <div style="font-weight:600; color:#10B981;">${b.closure}</div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.net > 0 ? '#EF4444' : '#10B981'}20; color:${b.net > 0 ? '#EF4444' : '#10B981'};">
                                    ${b.net > 0 ? '+' : ''}${b.net}
                                </span>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Portfolio Detail View ---
function renderPortfolioDetailView(category, stats) {
    const portfolioData = getPortfolioBreakdown();
    const total = stats.portfolioHealth.healthy + stats.portfolioHealth.dpd130 + stats.portfolioHealth.dpd3160 + stats.portfolioHealth.dpd6190;

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.portfolioHealth.healthy}</div>
                        <div class="detail-stat-label">Healthy (FTOD)</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#F59E0B;">${stats.portfolioHealth.dpd130}</div>
                        <div class="detail-stat-label">1-30 DPD</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#FB923C;">${stats.portfolioHealth.dpd3160}</div>
                        <div class="detail-stat-label">31-60 DPD</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#EF4444;">${stats.portfolioHealth.dpd6190}</div>
                        <div class="detail-stat-label">61-90 DPD</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${total}</div>
                        <div class="detail-stat-label">Total Accounts</div>
                    </div>
                </div>

                <div class="detail-filter-bar">
                    <button class="detail-filter-btn active" onclick="filterPortfolio('all')">All</button>
                    <button class="detail-filter-btn" onclick="filterPortfolio('healthy')" style="border-color:#10B981; color:#10B981;">Healthy</button>
                    <button class="detail-filter-btn" onclick="filterPortfolio('dpd130')" style="border-color:#F59E0B; color:#F59E0B;">1-30 DPD</button>
                    <button class="detail-filter-btn" onclick="filterPortfolio('dpd3160')" style="border-color:#FB923C; color:#FB923C;">31-60 DPD</button>
                    <button class="detail-filter-btn" onclick="filterPortfolio('dpd6190')" style="border-color:#EF4444; color:#EF4444;">61-90 DPD</button>
                </div>

                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>Region</div>
                        <div style="color:#10B981;">Healthy</div>
                        <div style="color:#F59E0B;">1-30 DPD</div>
                        <div style="color:#FB923C;">31-60 DPD</div>
                        <div style="color:#EF4444;">61-90 DPD</div>
                    </div>
                    ${portfolioData.map(b => `
                        <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr 1fr;">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.region}</div>
                            <div style="font-weight:600; color:#10B981;">${b.healthy}</div>
                            <div style="font-weight:600; color:#F59E0B;">${b.dpd130}</div>
                            <div style="font-weight:600; color:#FB923C;">${b.dpd3160}</div>
                            <div style="font-weight:600; color:#EF4444;">${b.dpd6190}</div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Disbursement Detail View ---
function renderDisbursementDetailView(stats) {
    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('disb_igl')">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.disbIglAccAchieve}</div>
                        <div class="detail-stat-label">IGL</div>
                        <div style="font-size:12px; color:var(--text-secondary);">Amt ${(stats.disbIglAmtAchieve / 100000).toFixed(1)}L</div>
                    </div>
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('disb_fig')">
                        <div class="detail-stat-value" style="color:#6366F1;">${stats.disbFigAccAchieve}</div>
                        <div class="detail-stat-label">FIG</div>
                        <div style="font-size:12px; color:var(--text-secondary);">Amt ${(stats.disbFigAmtAchieve / 100000).toFixed(1)}L</div>
                    </div>
                    <div class="detail-stat-card clickable-section" onclick="openDetailModal('disb_il')">
                        <div class="detail-stat-value" style="color:#F59E0B;">${stats.disbIlAccAchieve}</div>
                        <div class="detail-stat-label">IL</div>
                        <div style="font-size:12px; color:var(--text-secondary);">Amt ${(stats.disbIlAmtAchieve / 100000).toFixed(1)}L</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:var(--primary-accent);">Amt ${(stats.totalDisbursementAchieve / 10000000).toFixed(2)}Cr</div>
                        <div class="detail-stat-label">Total Disbursement</div>
                    </div>
                </div>

                ${renderDisbursementBranchTable(stats)}
            `;
}

// --- Disbursement Product Detail ---
function renderDisbProductDetail(product, stats) {
    const config = {
        'igl': { accField: 'disb_igl_acc', amtField: 'disb_igl_amt', label: 'IGL' },
        'fig': { accField: 'disb_fig_acc', amtField: 'disb_fig_amt', label: 'FIG' },
        'il': { accField: 'disb_il_acc', amtField: 'disb_il_amt', label: 'IL' }
    };

    const cfg = config[product];
    const branchData = getDisbursementByProduct(cfg.accField, cfg.amtField);

    return `
                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>Region</div>
                        <div>Accounts</div>
                        <div>Amount (₹)</div>
                    </div>
                    ${branchData.map(b => `
                        <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr;">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.region}</div>
                            <div style="font-weight:600;">${b.accounts}</div>
                            <div style="font-weight:600;">₹${(b.amount / 100000).toFixed(2)}L</div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Region Detail View ---
function renderRegionDetailView(regionName, stats) {
    const branches = getBranchesByRegion(regionName);
    const regionStats = stats.regionalBreakdown[regionName] || { branches: 0, completed: 0, avgPct: 0 };

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${regionStats.branches}</div>
                        <div class="detail-stat-label">Total Branches</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${regionStats.completed}</div>
                        <div class="detail-stat-label">Completed</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#F59E0B;">${regionStats.branches - regionStats.completed}</div>
                        <div class="detail-stat-label">Pending</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:var(--primary-accent);">${regionStats.avgPct}%</div>
                        <div class="detail-stat-label">Avg Achievement</div>
                    </div>
                </div>
                
                <div class="detail-branch-list">
                    <div class="detail-branch-header">
                        <div>Branch</div>
                        <div>District</div>
                        <div>DM</div>
                        <div>Status</div>
                        <div>Achievement</div>
                        <div>Action</div>
                    </div>
                    ${branches.map(b => `
                        <div class="detail-branch-row">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.district}</div>
                            <div>${b.dm}</div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.statusColor}20; color:${b.statusColor};">
                                    ${b.statusLabel}
                                </span>
                            </div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.pctColor}20; color:${b.pctColor};">
                                    ${b.pct}
                                </span>
                            </div>
                            <div>
                                <button onclick="openBranchFromDetail('${b.name}')" style="padding:6px 12px; border-radius:6px; border:1px solid var(--border-color); background:var(--bg-card); cursor:pointer; font-size:11px;">
                                    View
                                </button>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- Insight Detail (Pending/NoData/Completed) ---
function renderInsightDetail(insightType, stats) {
    const branches = getBranchListByStatus();
    let filtered = [];

    if (insightType === 'pending') {
        filtered = branches.filter(b => b.status === 'pending');
    } else if (insightType === 'nodata') {
        filtered = branches.filter(b => b.status === 'nodata');
    } else {
        filtered = branches.filter(b => b.status === 'completed');
    }

    return `
                <div style="margin-bottom:20px; font-size:14px; color:var(--text-secondary);">
                    Found <strong>${filtered.length}</strong> branches matching this criteria
                </div>
                
                <div class="detail-branch-list">
                    <div class="detail-branch-header">
                        <div>Branch</div>
                        <div>District</div>
                        <div>DM</div>
                        <div>Region</div>
                        <div>Status</div>
                        <div>Action</div>
                    </div>
                    ${filtered.map(b => `
                        <div class="detail-branch-row">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.district}</div>
                            <div>${b.dm}</div>
                            <div>${b.region}</div>
                            <div>
                                <span class="detail-pct-badge" style="background:${b.statusColor}20; color:${b.statusColor};">
                                    ${b.statusLabel}
                                </span>
                            </div>
                            <div>
                                <button onclick="openBranchFromDetail('${b.name}')" style="padding:6px 12px; border-radius:6px; border:1px solid var(--border-color); background:var(--bg-card); cursor:pointer; font-size:11px;">
                                    ${insightType === 'completed' ? 'View' : 'Enter Data'}
                                </button>
                            </div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// --- KYC Detail ---
function renderKYCDetail(stats) {
    const kycData = getKYCBreakdown();

    return `
                <div class="detail-summary-row">
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#6366F1;">${stats.kycIgl}</div>
                        <div class="detail-stat-label">IGL</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#10B981;">${stats.kycFig}</div>
                        <div class="detail-stat-label">FIG</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value" style="color:#F59E0B;">${stats.kycIl}</div>
                        <div class="detail-stat-label">IL</div>
                    </div>
                    <div class="detail-stat-card">
                        <div class="detail-stat-value">${stats.kycTotal}</div>
                        <div class="detail-stat-label">Total KYC</div>
                    </div>
                </div>

                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>Region</div>
                        <div style="color:#6366F1;">IGL</div>
                        <div style="color:#10B981;">FIG</div>
                        <div style="color:#F59E0B;">IL</div>
                    </div>
                    ${kycData.map(b => `
                        <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                            <div class="detail-branch-name">${b.name}</div>
                            <div>${b.region}</div>
                            <div style="font-weight:600; color:#6366F1;">${b.igl}</div>
                            <div style="font-weight:600; color:#10B981;">${b.fig}</div>
                            <div style="font-weight:600; color:#F59E0B;">${b.il}</div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// ============ HELPER FUNCTIONS FOR DETAIL VIEWS ============

function getBranchListByStatus() {
    const branches = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxDistrict = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'district');
    const idxDM = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'dm name');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        if (!name) return;

        const entry = state.branchDetails[name];
        const hasPlan = entry && entry.target;
        const hasAchieve = entry && entry.achievement;

        let status = 'nodata';
        let statusLabel = 'No Data';
        let statusColor = '#EF4444';

        if (hasPlan && hasAchieve) {
            status = 'completed';
            statusLabel = 'Completed';
            statusColor = '#10B981';
        } else if (hasPlan) {
            status = 'pending';
            statusLabel = 'Pending';
            statusColor = '#F59E0B';
        }

        const pctVal = calculateBranchAveragePercentage(name);
        const pct = pctVal !== null ? pctVal + '%' : '--';
        let pctColor = '#6B7280';
        if (pctVal !== null) {
            pctColor = pctVal >= 100 ? '#10B981' : pctVal >= 50 ? '#F59E0B' : '#EF4444';
        }

        branches.push({
            name,
            district: row[idxDistrict] || '',
            dm: row[idxDM] || '',
            region: row[idxRegion] || '',
            status,
            statusLabel,
            statusColor,
            pct,
            pctColor,
            pctVal: pctVal || 0
        });
    });

    return branches;
}

function getRegionBreakdown() {
    const regions = {};
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const branch = row[idxBranch];
        const region = row[idxRegion];
        if (!branch || !region) return;

        if (!regions[region]) regions[region] = { branches: 0, completed: 0 };
        regions[region].branches++;

        const entry = state.branchDetails[branch];
        if (entry && entry.target && entry.achievement) {
            regions[region].completed++;
        }
    });

    Object.keys(regions).forEach(r => {
        regions[r].completionPct = regions[r].branches > 0
            ? Math.round((regions[r].completed / regions[r].branches) * 100)
            : 0;
    });

    return regions;
}

function getUniqueDMs() {
    const idx = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'dm name');
    const dms = new Set(state.rawData.rows.map(r => r[idx]).filter(Boolean));
    return dms.size;
}

function getUniqueDistricts() {
    const idx = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'district');
    const districts = new Set(state.rawData.rows.map(r => r[idx]).filter(Boolean));
    return districts.size;
}

function getRankedBranches() {
    const branches = getBranchListByStatus();
    return branches.filter(b => b.status === 'completed' && b.pctVal > 0).sort((a, b) => b.pctVal - a.pctVal);
}

function getBranchNPAData() {
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.achievement) {
            const activation = parseInt(entry.achievement.npa_activation) || 0;
            const closure = parseInt(entry.achievement.npa_closure) || 0;
            data.push({ name, region, activation, closure, net: activation - closure });
        }
    });

    return data.sort((a, b) => b.net - a.net);
}

function getPortfolioBreakdown() {
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.achievement) {
            data.push({
                name,
                region,
                healthy: parseInt(entry.achievement.ftod_actual) || 0,
                dpd130: parseInt(entry.achievement.dpd_1_30_actual) || 0,
                dpd3160: parseInt(entry.achievement.dpd_31_60_actual) || 0,
                dpd6190: parseInt(entry.achievement.dpd_61_90_actual) || 0
            });
        }
    });

    return data;
}

function getDisbursementByProduct(accField, amtField) {
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.achievement) {
            const accounts = parseInt(entry.achievement[accField]) || 0;
            const amount = parseInt(entry.achievement[amtField]) || 0;
            if (accounts > 0 || amount > 0) {
                data.push({ name, region, accounts, amount });
            }
        }
    });

    return data.sort((a, b) => b.amount - a.amount);
}

function getBranchesByRegion(regionName) {
    const branches = getBranchListByStatus();
    return branches.filter(b => b.region === regionName);
}

function getKYCBreakdown() {
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.achievement) {
            data.push({
                name,
                region,
                igl: parseInt(entry.achievement.kyc_igl) || 0,
                fig: parseInt(entry.achievement.kyc_fig) || 0,
                il: parseInt(entry.achievement.kyc_il) || 0
            });
        }
    });

    return data;
}

function renderMetricBranchTable(metricType, stats) {
    const config = {
        'ftod': { targetField: 'ftod_plan', achieveField: 'ftod_actual', label: 'FTOD' },
        'dpd_1_30': { targetField: 'dpd_1_30_plan', achieveField: 'dpd_1_30_actual', label: '1-30 DPD' },
        'dpd_31_60': { targetField: 'dpd_31_60_plan', achieveField: 'dpd_31_60_actual', label: '31-60 DPD' },
        'dpd_61_90': { targetField: 'dpd_61_90_plan', achieveField: 'dpd_61_90_actual', label: '61-90 DPD' },
        'collection': { targetField: null, achieveField: null, label: 'All' }
    };

    const cfg = config[metricType] || config['collection'];
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.target && entry.achievement) {
            let target, achieve;
            if (cfg.targetField) {
                target = parseInt(entry.target[cfg.targetField]) || 0;
                achieve = parseInt(entry.achievement[cfg.achieveField]) || 0;
            } else {
                target = (parseInt(entry.target.ftod_plan) || 0) + (parseInt(entry.target.dpd_1_30_plan) || 0) + (parseInt(entry.target.dpd_31_60_plan) || 0) + (parseInt(entry.target.dpd_61_90_plan) || 0);
                achieve = (parseInt(entry.achievement.ftod_actual) || 0) + (parseInt(entry.achievement.dpd_1_30_actual) || 0) + (parseInt(entry.achievement.dpd_31_60_actual) || 0) + (parseInt(entry.achievement.dpd_61_90_actual) || 0);
            }

            const pct = target > 0 ? Math.round((achieve / target) * 100) : 0;
            data.push({ name, region, target, achieve, pct });
        }
    });

    data.sort((a, b) => a.pct - b.pct);

    return `
                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>Region</div>
                        <div>Plan</div>
                        <div>Achievement</div>
                        <div>%</div>
                    </div>
                    ${data.map(b => {
        const pctColor = b.pct >= 100 ? '#10B981' : b.pct >= 50 ? '#F59E0B' : '#EF4444';
        return `
                            <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                                <div class="detail-branch-name">${b.name}</div>
                                <div>${b.region}</div>
                                <div>${b.target.toLocaleString()}</div>
                                <div style="font-weight:600;">${b.achieve.toLocaleString()}</div>
                                <div>
                                    <span class="detail-pct-badge" style="background:${pctColor}20; color:${pctColor};">${b.pct}%</span>
                                </div>
                            </div>
                        `;
    }).join('')}
                </div>
            `;
}

function renderDisbursementBranchTable(stats) {
    const data = [];
    const idxBranch = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'branch');
    const idxRegion = state.rawData.headers.findIndex(h => h.trim().toLowerCase() === 'region');

    state.rawData.rows.forEach(row => {
        const name = row[idxBranch];
        const region = row[idxRegion];
        if (!name) return;

        const entry = state.branchDetails[name];
        if (entry && entry.achievement) {
            const a = entry.achievement;
            const total = (parseInt(a.disb_igl_amt) || 0) + (parseInt(a.disb_fig_amt) || 0) + (parseInt(a.disb_il_amt) || 0);
            if (total > 0) {
                data.push({
                    name,
                    region,
                    igl: (parseInt(a.disb_igl_amt) || 0),
                    fig: (parseInt(a.disb_fig_amt) || 0),
                    il: (parseInt(a.disb_il_amt) || 0),
                    total
                });
            }
        }
    });

    data.sort((a, b) => b.total - a.total);

    return `
                <div class="detail-branch-list">
                    <div class="detail-branch-header" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                        <div>Branch</div>
                        <div>IGL (₹L)</div>
                        <div>FIG (₹L)</div>
                        <div>IL (₹L)</div>
                        <div>Total (₹L)</div>
                    </div>
                    ${data.map(b => `
                        <div class="detail-branch-row" style="grid-template-columns: 2fr 1fr 1fr 1fr 1fr;">
                            <div>
                                <div class="detail-branch-name">${b.name}</div>
                                <div class="detail-branch-region">${b.region}</div>
                            </div>
                            <div>${(b.igl / 100000).toFixed(1)}</div>
                            <div>${(b.fig / 100000).toFixed(1)}</div>
                            <div>${(b.il / 100000).toFixed(1)}</div>
                            <div style="font-weight:700; color:var(--primary-accent);">${(b.total / 100000).toFixed(1)}</div>
                        </div>
                    `).join('')}
                </div>
            `;
}

// Filter functions for detail views
function filterBranchList(filter) {
    document.querySelectorAll('.detail-filter-btn').forEach(btn => btn.classList.remove('active'));
    event.target.classList.add('active');

    document.querySelectorAll('#branchListContainer .detail-branch-row').forEach(row => {
        const status = row.dataset.status;
        if (filter === 'all' || status === filter) {
            row.style.display = 'grid';
        } else {
            row.style.display = 'none';
        }
    });
}

function filterPortfolio(filter) {
    document.querySelectorAll('.detail-filter-btn').forEach(btn => btn.classList.remove('active'));
    event.target.classList.add('active');
    // For now just visual feedback, could implement actual filtering
}

// --- INPUT VALIDATION ---
function setupInputValidators() {
    const pairs = [
        { plan: 'ftod_plan', actual: 'ftod_actual', name: 'FTOD Collection Plan' },
        { plan: 'dpd_1_30_plan', actual: 'dpd_1_30_actual', name: '1-30 DPD Collection Plan' },
        { plan: 'dpd_31_60_plan', actual: 'dpd_31_60_actual', name: '31-60 DPD Collection Plan' },
        { plan: 'dpd_61_90_plan', actual: 'dpd_61_90_actual', name: '61-90 DPD Collection Plan' },
        { plan: 'fy_non_start_plan', actual: 'fy_non_start_acc', name: 'Non-Starter Collection Plan' }
    ];

    pairs.forEach(pair => {
        const planInput = document.getElementById(pair.plan);
        const actualInput = document.getElementById(pair.actual);

        if (planInput && actualInput) {
            // Validate when plan input changes
            planInput.addEventListener('input', function () {
                // SKIP VALIDATION IN ACHIEVEMENT MODE (Actuals are 0/hidden)
                if (typeof currentModalState !== 'undefined' && currentModalState === 'ACHIEVEMENT') return;

                // Remove non-numeric characters for safety if needed, or just parse
                const planVal = parseInt(this.value) || 0;
                const actualVal = parseInt(actualInput.value) || 0;

                if (planVal > actualVal) {
                    showToast(`${pair.name} cannot be more than Actual Account (${actualVal})`, "alert");
                    this.value = actualVal;
                }
            });

            // Re-validate when actual input changes
            actualInput.addEventListener('input', function () {
                const planVal = parseInt(planInput.value) || 0;
                const actualVal = parseInt(this.value) || 0;
                if (planVal > actualVal) {
                    planInput.value = actualVal;
                }
            });
        }
    });
}

// Initialize validators
document.addEventListener('DOMContentLoaded', setupInputValidators);
// Also call immediately in case DOM is already loaded (script is at end of body)
setupInputValidators();
// --- PROFILE MODAL ---
function toggleProfileModal() {
    const modal = document.getElementById('profileModal');
    if (!modal) return;

    // Toggle visibility
    if (modal.classList.contains('visible')) {
        modal.classList.remove('visible');
    } else {
        modal.classList.add('visible');

        // Update content dynamically
        // Note: state.role and state.currentUser are set during login
        const role = state.role === 'CEO' ? 'Administrator' : 'Division Manager';
        // Use currentUser if available, or fallback to 'User'
        const name = state.currentUser === 'CEO' ? 'CEO User' : (state.currentUser || 'User');

        const modalName = document.getElementById('modalName');
        const modalRole = document.getElementById('modalRole');
        const headerAvatar = document.getElementById('headerAvatar');
        const modalAvatar = document.getElementById('modalAvatar');
        const modalPhone = document.getElementById('modalPhone');

        if (modalName) modalName.textContent = name;
        if (modalRole) modalRole.textContent = role;
        // Copy the avatar text (initials) from the header to the modal
        if (modalAvatar && headerAvatar) modalAvatar.textContent = headerAvatar.textContent;

        // Phone number logic
        if (modalPhone) {
            if (state.role === 'CEO') {
                modalPhone.textContent = "8792320623"; // Default for CEO/Support
            } else {
                // Find user in rawData
                // [0]=Branch, [1]=District, [2]=DM Name, [3]=Region, [4]=Phone
                const dmRow = state.rawData.rows.find(row => row[2] === name);
                const phone = (dmRow && dmRow[4]) ? dmRow[4] : "Not Available";
                modalPhone.textContent = phone;
            }
        }
    }
}

// Close profile modal on outside click
window.addEventListener('click', function (e) {
    const modal = document.getElementById('profileModal');
    const profileBtn = document.querySelector('.user-profile');

    // If click is outside modal container AND not on the toggle button
    if (modal && modal.classList.contains('visible')) {
        const modalContainer = modal.querySelector('.modal-container');
        // Ensure we are clicking on the overlay (modal itself) or outside the container
        // Actually, the structure is overlay > container. 
        // If we click on overlay (which has id profileModal), it means we are outside container.
        if (e.target === modal) {
            modal.classList.remove('visible');
        }
    }
});

