require('dotenv').config();
const express = require("express");
const { Pool, Client } = require("pg");
const cors = require("cors");
const http = require("http");
const { Server } = require("socket.io");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const XLSX = require("xlsx");

const UPLOAD_DIR = path.join(__dirname, "..", "data");

const rateLimit = require('express-rate-limit');

const app = express();
app.set('trust proxy', 1);
app.use(cors());
app.use(express.json({limit: '10mb'}));

// Rate limiters
const uploadLimiter = rateLimit({ windowMs: 60 * 1000, max: 5, message: { error: 'Too many uploads. Try again in a minute.' } });
const aiLimiter = rateLimit({ windowMs: 60 * 1000, max: 20, message: { error: 'Too many AI requests. Slow down.' } });

// AI access gate. FO is the only role excluded — directors / area / branch
// managers / ops execs all keep access. Override via env var if the policy
// changes (comma-separated, case-insensitive).
const AI_BLOCKED_ROLES = new Set(
  String(process.env.AI_BLOCKED_ROLES || 'FO')
    .split(',')
    .map((s) => s.trim().toUpperCase())
    .filter(Boolean)
);
function _readAiSessionRole(req) {
  // Voice paths use multipart, chat paths use JSON. role lives in body
  // for both, plus a couple of legacy headers from the mobile client.
  const fromBody = req.body && (req.body.role || req.body.session && req.body.session.role);
  const fromHdr = req.headers['x-nlpl-role'];
  return String(fromBody || fromHdr || '').trim().toUpperCase();
}
function requireAiAccess(req, res, next) {
  const role = _readAiSessionRole(req);
  if (role && AI_BLOCKED_ROLES.has(role)) {
    return res.status(403).json({
      error: 'role_not_allowed',
      message: 'AI Assistant is available for Branch Managers and above. Field Officers do not have access.',
      role,
    });
  }
  next();
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    if (!file.originalname.match(/\.xlsx?$/i)) {
      return cb(new Error("Only .xlsx/.xls files allowed"));
    }
    cb(null, true);
  }
});

const dbConfig = {
  host: process.env.PGHOST || "127.0.0.1",
  user: process.env.PGUSER || "Raghunandan1157",
  password: process.env.PGPASSWORD || "raghu",
  database: process.env.PGDATABASE || "postgres",
  port: Number(process.env.PGPORT) || 5432,
};

const pool = new Pool({ ...dbConfig, max: 10 });

// Database pool cache for multi-database support
const poolCache = {};
function getPool(dbName) {
  if (!dbName || dbName === 'postgres') return pool;
  if (!poolCache[dbName]) {
    poolCache[dbName] = new Pool({
      host: process.env.PGHOST || '127.0.0.1',
      port: Number(process.env.PGPORT) || 5432,
      user: process.env.PGUSER || 'Raghunandan1157',
      password: process.env.PGPASSWORD || 'raghu',
      database: dbName,
      max: 5
    });
  }
  return poolCache[dbName];
}

// Graceful pool error handling
pool.on("error", (err) => {
  console.error("Unexpected pool error:", err.message);
});


// ========== DISTRICT_NORMALIZE ==========
const DISTRICT_NORMALIZE = {
  'BENGALORE -RURAL': 'BANGALORE RURAL',
  'BENGALORE -URBAN': 'BANGALORE URBAN',
  'BENGALORE-RURAL': 'BANGALORE RURAL',
  'BENGALORE-URBAN': 'BANGALORE URBAN',
  'CHIKKABALLAPUR': 'CHIKKABALLAPURA',
  'KALBURGI': 'KALABURGI',
};
const REGION_NORMALIZE = {
  'KALBURGI': 'KALABURAGI',
  'KALBURAGI': 'KALABURAGI',
};
function normalizeRegion(name) {
  var n = (name || '').trim();
  return REGION_NORMALIZE[n] || n;
}
function normalizeDistrict(name) {
  var n = (name || '').trim();
  return DISTRICT_NORMALIZE[n] || n;
}

// ========== EOD METRIC HEADER MAP ==========
// EOD report sheets (IGL/FIG/VVY/*_FY) have 29 metric columns including 4
// "1-90 Demand/Collection" + "1-90 Demand/Collection Amt" buckets that are
// not part of the DB schema. A positional 25-col read shifts indexes 8-24
// into the wrong DB columns (e.g. npa_cases got "1-90 Demand", every amount
// column was off by two). Helpers below map each DB column by header name
// and fall back to the original positional layout with a warning if the
// header row is missing or unrecognisable. Same fix pattern as 91e068f
// (POS column header map).
const EOD_METRIC_HEADERS = [
  ['regular_demand',          'regular demand'],
  ['regular_collection',      'regular collection'],
  ['demand_1_30',             '1-30 demand'],
  ['collection_1_30',         '1-30 collection'],
  ['demand_31_60',            '31-60 demand'],
  ['collection_31_60',        '31-60 collection'],
  ['pnpa_demand',             'pnpa demand'],
  ['pnpa_collection',         'pnpa collection'],
  ['npa_cases',               'npa cases'],
  ['npa_act_acc',             'npa act acc'],
  ['npa_act_amt',             'npa act amt'],
  ['npa_clo_acc',             'npa clo acc'],
  ['npa_clo_amt',             'npa clo amt'],
  ['on_date_demand',          'on-date demand'],
  ['on_date_collection',      'on-date collection'],
  ['regular_demand_amt',      'regular demand amt'],
  ['regular_collection_amt',  'regular collection amt'],
  ['demand_1_30_amt',         '1-30 demand amt'],
  ['collection_1_30_amt',     '1-30 collection amt'],
  ['demand_31_60_amt',        '31-60 demand amt'],
  ['collection_31_60_amt',    '31-60 collection amt'],
  ['pnpa_demand_amt',         'pnpa demand amt'],
  ['pnpa_collection_amt',     'pnpa collection amt'],
  ['on_date_demand_amt',      'on-date demand amt'],
  ['on_date_collection_amt',  'on-date collection amt'],
];
const EOD_METRIC_DB_COLS = EOD_METRIC_HEADERS.map(p => p[0]);
function _normEodHeader(h) {
  return String(h || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function buildEodMetricColMap(headerRow) {
  const map = {};
  if (!Array.isArray(headerRow)) return map;
  const idx = {};
  for (let i = 0; i < headerRow.length; i++) {
    const k = _normEodHeader(headerRow[i]);
    if (k && !(k in idx)) idx[k] = i;
  }
  for (const [dbCol, headerText] of EOD_METRIC_HEADERS) {
    if (headerText in idx) map[dbCol] = idx[headerText];
  }
  return map;
}
function readEodMetrics(row, colMap, useHeaderMap, fallbackStart) {
  return EOD_METRIC_DB_COLS.map((dbCol, i) => {
    const colIdx = useHeaderMap && colMap[dbCol] !== undefined
      ? colMap[dbCol]
      : (fallbackStart + i);
    const raw = row[colIdx];
    if (raw == null || raw === '') return 0;
    const num = Number(raw);
    return Number.isFinite(num) ? num : 0;
  });
}

// ========== BRANCH V2 MAP (ESAF hierarchy: state → division → area) ==========
const BRANCH_V2_MAP = {
  'AFZALPUR': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'AJJAMPURA': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'ALAND': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'ALMEL': {state: 'KALBURGI', division: 'KALBURGI', area: 'INDI'},
  'ATHANI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'AURAD': {state: 'KALBURGI', division: 'BIDAR', area: 'BIDAR'},
  'BADAMI': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'BAGALKOT': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BAGALKOT'},
  'BAGEPALLI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'CHIKBALLAPURA'},
  'BAILHONGAL': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'BALLARI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'BANGARPET': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'BASAVAKALYAN': {state: 'KALBURGI', division: 'BIDAR', area: 'HUMNABAD'},
  'BELAGAVI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'BETHAMANGALA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'BHALKI': {state: 'KALBURGI', division: 'BIDAR', area: 'BIDAR'},
  'BIDAR': {state: 'KALBURGI', division: 'BIDAR', area: 'BIDAR'},
  'BIDAR-2': {state: 'KALBURGI', division: 'BIDAR', area: 'BIDAR'},
  'BILAGI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BAGALKOT'},
  'BUDWAL': {state: 'AP', division: 'KADAPPA', area: 'KADAPA'},
  'CHADCHAN': {state: 'KALBURGI', division: 'KALBURGI', area: 'INDI'},
  'CHALLAKERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'CHANDAPURA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'BANGALORE URBAN'},
  'CHANNAGIRI': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'CHIKBALLAPURA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'CHIKBALLAPURA'},
  'CHIKKAMAGALURU': {state: 'TUMKUR', division: 'TUMKUR', area: 'CHIKKAMAGALURU'},
  'CHIKKANAYAKANAHALLI': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'CHIKKODI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'CHINCHOLI': {state: 'KALBURGI', division: 'BIDAR', area: 'SEDAM'},
  'CHINTAMANI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'CHIKBALLAPURA'},
  'CHITRADURGA': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'CORPORATE OFFICE': {state: 'CORPORATE OFFICE', division: 'CORPORATE OFFICE', area: 'CORPORATE OFFICE'},
  'DABUSPET': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'DAVANAGERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
  'DEVADURGA': {state: 'KALBURGI', division: 'KALBURGI', area: 'SHAHAPUR'},
  'DEVANAHALLI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'DHARMAVARAM': {state: 'AP', division: 'KADAPPA', area: 'KADAPA'},
  'DHARWAD': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'DODDABALLAPURA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'GADAG': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'GADWAL': {state: 'TS', division: 'SANGAREDDY', area: 'MAHABUB NAGAR'},
  'GAJENDRAGAD': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'GANGAVATHI': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'GOKAK': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'GOWRIBIDANUR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'GUBBI': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'HAGARIBOMMANAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'HARAPANAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'HARIHARA': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
  'HEAD OFFICE': {state: 'HEAD OFFICE', division: 'HEAD OFFICE', area: 'HEAD OFFICE'},
  'HEBBAL': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'BANGALORE URBAN'},
  'HIRIYUR': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'HOLAKERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'HONNALI': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
  'HOSADURGA': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'CHITRADURGA'},
  'HOSPET': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'HUBLI': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'HUBLI-2': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'HULIYAR': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'HULSOOR': {state: 'KALBURGI', division: 'BIDAR', area: 'HUMNABAD'},
  'HUMNABAD': {state: 'KALBURGI', division: 'BIDAR', area: 'HUMNABAD'},
  'HUNGUND': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'HUVENAHADAGALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'INDI': {state: 'KALBURGI', division: 'KALBURGI', area: 'INDI'},
  'J P NAGAR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'BANGALORE URBAN'},
  'JAGALORE': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'JAMAKHANDI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BAGALKOT'},
  'JEVARGI': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'KADAPA': {state: 'AP', division: 'KADAPPA', area: 'KADAPA'},
  'KADIRI': {state: 'AP', division: 'KADAPPA', area: 'KADAPA'},
  'KADUR': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'KALABURAGI': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'KALAGI': {state: 'KALBURGI', division: 'BIDAR', area: 'SEDAM'},
  'KALBURGI-2': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'KALGHATGI': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'KAMALAPURA': {state: 'KALBURGI', division: 'BIDAR', area: 'HUMNABAD'},
  'KENGERI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'BANGALORE URBAN'},
  'KHANAHOSAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'KITTUR': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'KODANGAL': {state: 'TS', division: 'SANGAREDDY', area: 'SANGAREDDY'},
  'KODANGAL(VIKARABAD)': {state: 'TS', division: 'SANGAREDDY', area: 'SANGAREDDY'},
  'KOLAR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'KOPPAL': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'KORATAGERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'KOTTURU': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'KUDATHINI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'KUDLIGI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'KUNIGAL': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'KUSHTAGI': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'LAXMESHWAR': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'LINGSUGUR': {state: 'KALBURGI', division: 'KALBURGI', area: 'LINGSUGUR'},
  'LOKAPUR': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BAGALKOT'},
  'MADHUGIRI': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'MAHABUB NAGAR': {state: 'TS', division: 'SANGAREDDY', area: 'MAHABUB NAGAR'},
  'MALUR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'MANVI': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'MARIKAL': {state: 'TS', division: 'SANGAREDDY', area: 'MAHABUB NAGAR'},
  'MUDALAGI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'MUDDEBIHAL': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'MUDIGERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'CHIKKAMAGALURU'},
  'MUNDARAGI': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'NARAGUNDA': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'NARAYANKHED': {state: 'TS', division: 'SANGAREDDY', area: 'SANGAREDDY'},
  'NIPPANI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'NR PURA': {state: 'TUMKUR', division: 'TUMKUR', area: 'CHIKKAMAGALURU'},
  'PANCHANHALLI': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'RAICHUR': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'RAMDURGA': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'SANDURU': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'SANGAREDDY': {state: 'TS', division: 'SANGAREDDY', area: 'SANGAREDDY'},
  'SANTHEBENNURU': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
  'SEDAM': {state: 'KALBURGI', division: 'BIDAR', area: 'SEDAM'},
  'SHAHAPUR': {state: 'KALBURGI', division: 'KALBURGI', area: 'SHAHAPUR'},
  'SINDAGI': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'SINDHNUR': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'SIRA': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'SIRUGUPPA': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'SIRWAR': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'SRINIVASPURA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'CHIKBALLAPURA'},
  'TALIKOTI': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'TANDUR': {state: 'TS', division: 'SANGAREDDY', area: 'MAHABUB NAGAR'},
  'TARIKERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'TIKOTA': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'TIPTUR': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'TUMKUR': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'TUREVEKERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'VIJAYAPUR': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'YADGIR': {state: 'KALBURGI', division: 'BIDAR', area: 'SHAHAPUR'},
  'YARAGATTI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'ZAHEERABAD': {state: 'TS', division: 'SANGAREDDY', area: 'SANGAREDDY'}
};

// ========== V2 TABLE POPULATION HELPERS ==========

async function populateV2Tables(client) {
  await client.query("BEGIN");
  try {
    const { rows } = await client.query(`
      SELECT b.branch_name, e.emp_id, e.officer_name,
             ep.product_type_id,
             ep.regular_demand, ep.regular_collection,
             ep.demand_1_30, ep.collection_1_30,
             ep.demand_31_60, ep.collection_31_60,
             ep.pnpa_demand, ep.pnpa_collection,
             ep.npa_cases, ep.npa_act_acc, ep.npa_act_amt,
             ep.npa_clo_acc, ep.npa_clo_amt,
             ep.on_date_demand, ep.on_date_collection,
             ep.regular_demand_amt, ep.regular_collection_amt,
             ep.demand_1_30_amt, ep.collection_1_30_amt,
             ep.demand_31_60_amt, ep.collection_31_60_amt,
             ep.pnpa_demand_amt, ep.pnpa_collection_amt,
             ep.on_date_demand_amt, ep.on_date_collection_amt
      FROM employee_performance ep
      JOIN employees e ON ep.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
    `);

    await client.query(`TRUNCATE v2_employee_performance, v2_employees, v2_branches, v2_areas, v2_divisions, v2_states RESTART IDENTITY CASCADE`);

    const statesMap = {}, divisionsMap = {}, areasMap = {}, branchesMap = {};
    const employeesInserted = new Set();

    for (const row of rows) {
      const key = (row.branch_name || '').trim().toUpperCase();
      const info = BRANCH_V2_MAP[key] || {state: 'UNMAPPED', division: 'UNMAPPED', area: 'UNMAPPED'};

      if (!statesMap[info.state]) {
        const r = await client.query('INSERT INTO v2_states (state_name) VALUES ($1) RETURNING state_id', [info.state]);
        statesMap[info.state] = r.rows[0].state_id;
      }
      const stateId = statesMap[info.state];

      const divKey = stateId + '|' + info.division;
      if (!divisionsMap[divKey]) {
        const r = await client.query('INSERT INTO v2_divisions (division_name, state_id) VALUES ($1,$2) RETURNING division_id', [info.division, stateId]);
        divisionsMap[divKey] = r.rows[0].division_id;
      }
      const divId = divisionsMap[divKey];

      const areaKey = divId + '|' + info.area;
      if (!areasMap[areaKey]) {
        const r = await client.query('INSERT INTO v2_areas (area_name, division_id) VALUES ($1,$2) RETURNING area_id', [info.area, divId]);
        areasMap[areaKey] = r.rows[0].area_id;
      }
      const areaId = areasMap[areaKey];

      const branchKey = areaId + '|' + key;
      if (!branchesMap[branchKey]) {
        const r = await client.query('INSERT INTO v2_branches (branch_name, area_id) VALUES ($1,$2) RETURNING branch_id', [row.branch_name.trim() || 'UNKNOWN', areaId]);
        branchesMap[branchKey] = r.rows[0].branch_id;
      }
      const branchId = branchesMap[branchKey];

      if (!employeesInserted.has(row.emp_id)) {
        await client.query(
          'INSERT INTO v2_employees (emp_id, officer_name, branch_id) VALUES ($1,$2,$3) ON CONFLICT (emp_id) DO NOTHING',
          [row.emp_id, row.officer_name, branchId]
        );
        employeesInserted.add(row.emp_id);
      }

      await client.query(
        `INSERT INTO v2_employee_performance (emp_id, product_type_id,
          regular_demand, regular_collection, demand_1_30, collection_1_30,
          demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
          npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
          on_date_demand, on_date_collection,
          regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
          demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
          on_date_demand_amt, on_date_collection_amt)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)
        ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
          regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
          demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
          demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
          pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
          npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
          npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
          on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
          regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
          demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
          demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
          pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
          on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`,
        [row.emp_id, row.product_type_id,
         row.regular_demand, row.regular_collection, row.demand_1_30, row.collection_1_30,
         row.demand_31_60, row.collection_31_60, row.pnpa_demand, row.pnpa_collection,
         row.npa_cases, row.npa_act_acc, row.npa_act_amt, row.npa_clo_acc, row.npa_clo_amt,
         row.on_date_demand, row.on_date_collection,
         row.regular_demand_amt, row.regular_collection_amt, row.demand_1_30_amt, row.collection_1_30_amt,
         row.demand_31_60_amt, row.collection_31_60_amt, row.pnpa_demand_amt, row.pnpa_collection_amt,
         row.on_date_demand_amt, row.on_date_collection_amt]
      );
    }

    await client.query("COMMIT");
    console.log(`V2 tables populated: ${Object.keys(statesMap).length} states, ${Object.keys(divisionsMap).length} divisions, ${Object.keys(areasMap).length} areas, ${Object.keys(branchesMap).length} branches, ${employeesInserted.size} employees`);
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    throw err;
  }
}

async function populateV2PortfolioTables(client) {
  await client.query("BEGIN");
  try {
    const { rows } = await client.query(`
      SELECT pp.month_id, pp.emp_id, pp.product_type_id,
             pp.regular_demand, pp.regular_collection,
             pp.demand_1_30, pp.collection_1_30,
             pp.demand_31_60, pp.collection_31_60,
             pp.pnpa_demand, pp.pnpa_collection,
             pp.npa_cases, pp.npa_act_acc, pp.npa_act_amt,
             pp.npa_clo_acc, pp.npa_clo_amt,
             pp.on_date_demand, pp.on_date_collection,
             pp.regular_demand_amt, pp.regular_collection_amt,
             pp.demand_1_30_amt, pp.collection_1_30_amt,
             pp.demand_31_60_amt, pp.collection_31_60_amt,
             pp.pnpa_demand_amt, pp.pnpa_collection_amt,
             pp.on_date_demand_amt, pp.on_date_collection_amt,
             pp.regular_pos, pp.sma0_pos, pp.sma1_pos,
             pp.pnpa_pos, pp.npa_pos, pp.total_pos,
             e.officer_name, b.branch_name
      FROM portfolio_performance pp
      JOIN employees e ON pp.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
    `);

    await client.query(`
      TRUNCATE v2_employee_pos, v2_branch_pos, v2_fy_performance,
               v2_portfolio_performance, v2_employees, v2_branches,
               v2_areas, v2_divisions, v2_states
      RESTART IDENTITY CASCADE
    `);

    // Build unique employee set from all portfolio data
    const uniqueEmployees = {};
    for (const row of rows) {
      if (!uniqueEmployees[row.emp_id]) {
        uniqueEmployees[row.emp_id] = {officerName: row.officer_name, branchName: row.branch_name};
      }
    }

    const statesMap = {}, divisionsMap = {}, areasMap = {}, branchesMap = {};
    const employeesInserted = new Set();

    for (const [empId, {officerName, branchName}] of Object.entries(uniqueEmployees)) {
      const key = (branchName || '').trim().toUpperCase();
      const info = BRANCH_V2_MAP[key] || {state: 'UNMAPPED', division: 'UNMAPPED', area: 'UNMAPPED'};

      if (!statesMap[info.state]) {
        const r = await client.query('INSERT INTO v2_states (state_name) VALUES ($1) RETURNING state_id', [info.state]);
        statesMap[info.state] = r.rows[0].state_id;
      }
      const stateId = statesMap[info.state];

      const divKey = stateId + '|' + info.division;
      if (!divisionsMap[divKey]) {
        const r = await client.query('INSERT INTO v2_divisions (division_name, state_id) VALUES ($1,$2) RETURNING division_id', [info.division, stateId]);
        divisionsMap[divKey] = r.rows[0].division_id;
      }
      const divId = divisionsMap[divKey];

      const areaKey = divId + '|' + info.area;
      if (!areasMap[areaKey]) {
        const r = await client.query('INSERT INTO v2_areas (area_name, division_id) VALUES ($1,$2) RETURNING area_id', [info.area, divId]);
        areasMap[areaKey] = r.rows[0].area_id;
      }
      const areaId = areasMap[areaKey];

      const branchKey = areaId + '|' + key;
      if (!branchesMap[branchKey]) {
        const r = await client.query('INSERT INTO v2_branches (branch_name, area_id) VALUES ($1,$2) RETURNING branch_id', [(branchName || '').trim() || 'UNKNOWN', areaId]);
        branchesMap[branchKey] = r.rows[0].branch_id;
      }
      const branchId = branchesMap[branchKey];

      if (!employeesInserted.has(empId)) {
        await client.query(
          'INSERT INTO v2_employees (emp_id, officer_name, branch_id) VALUES ($1,$2,$3) ON CONFLICT (emp_id) DO NOTHING',
          [empId, officerName, branchId]
        );
        employeesInserted.add(empId);
      }
    }

    // Copy all portfolio_performance to v2_portfolio_performance
    for (const row of rows) {
      await client.query(
        `INSERT INTO v2_portfolio_performance (month_id, emp_id, product_type_id,
          regular_demand, regular_collection, demand_1_30, collection_1_30,
          demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
          npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
          on_date_demand, on_date_collection,
          regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
          demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
          on_date_demand_amt, on_date_collection_amt,
          regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33,$34)
        ON CONFLICT (month_id, emp_id, product_type_id) DO UPDATE SET
          regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
          demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
          demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
          pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
          npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
          npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
          on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
          regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
          demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
          demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
          pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
          on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt,
          regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos, sma1_pos=EXCLUDED.sma1_pos,
          pnpa_pos=EXCLUDED.pnpa_pos, npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`,
        [row.month_id, row.emp_id, row.product_type_id,
         row.regular_demand, row.regular_collection, row.demand_1_30, row.collection_1_30,
         row.demand_31_60, row.collection_31_60, row.pnpa_demand, row.pnpa_collection,
         row.npa_cases, row.npa_act_acc, row.npa_act_amt, row.npa_clo_acc, row.npa_clo_amt,
         row.on_date_demand, row.on_date_collection,
         row.regular_demand_amt, row.regular_collection_amt, row.demand_1_30_amt, row.collection_1_30_amt,
         row.demand_31_60_amt, row.collection_31_60_amt, row.pnpa_demand_amt, row.pnpa_collection_amt,
         row.on_date_demand_amt, row.on_date_collection_amt,
         row.regular_pos, row.sma0_pos, row.sma1_pos, row.pnpa_pos, row.npa_pos, row.total_pos]
      );
    }

    // Sync fy_performance → v2_fy_performance
    try {
      const { rows: fyRows } = await client.query(`
        SELECT fp.emp_id, fp.product_type_id,
               fp.regular_demand, fp.regular_collection, fp.demand_1_30, fp.collection_1_30,
               fp.demand_31_60, fp.collection_31_60, fp.pnpa_demand, fp.pnpa_collection,
               fp.npa_cases, fp.npa_act_acc, fp.npa_act_amt, fp.npa_clo_acc, fp.npa_clo_amt,
               fp.on_date_demand, fp.on_date_collection,
               fp.regular_demand_amt, fp.regular_collection_amt, fp.demand_1_30_amt, fp.collection_1_30_amt,
               fp.demand_31_60_amt, fp.collection_31_60_amt, fp.pnpa_demand_amt, fp.pnpa_collection_amt,
               fp.on_date_demand_amt, fp.on_date_collection_amt,
               fp.regular_pos, fp.sma0_pos, fp.sma1_pos, fp.pnpa_pos, fp.npa_pos, fp.total_pos
        FROM fy_performance fp WHERE fp.emp_id IN (SELECT emp_id FROM v2_employees)
      `);
      for (const fyRow of fyRows) {
        await client.query(
          `INSERT INTO v2_fy_performance (emp_id, product_type_id,
            regular_demand, regular_collection, demand_1_30, collection_1_30,
            demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt,
            regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33)
          ON CONFLICT (emp_id, product_type_id) DO NOTHING`,
          [fyRow.emp_id, fyRow.product_type_id,
           fyRow.regular_demand, fyRow.regular_collection, fyRow.demand_1_30, fyRow.collection_1_30,
           fyRow.demand_31_60, fyRow.collection_31_60, fyRow.pnpa_demand, fyRow.pnpa_collection,
           fyRow.npa_cases, fyRow.npa_act_acc, fyRow.npa_act_amt, fyRow.npa_clo_acc, fyRow.npa_clo_amt,
           fyRow.on_date_demand, fyRow.on_date_collection,
           fyRow.regular_demand_amt, fyRow.regular_collection_amt, fyRow.demand_1_30_amt, fyRow.collection_1_30_amt,
           fyRow.demand_31_60_amt, fyRow.collection_31_60_amt, fyRow.pnpa_demand_amt, fyRow.pnpa_collection_amt,
           fyRow.on_date_demand_amt, fyRow.on_date_collection_amt,
           fyRow.regular_pos, fyRow.sma0_pos, fyRow.sma1_pos, fyRow.pnpa_pos, fyRow.npa_pos, fyRow.total_pos]
        );
      }
    } catch (fyErr) {
      console.error("v2_fy_performance sync warning:", fyErr.message);
    }

    // Sync branch_pos → v2_branch_pos
    try {
      const { rows: bpRows } = await client.query(`
        SELECT bp.month_id, bp.branch_name, bp.product_name,
               bp.regular_pos, bp.sma0_pos, bp.sma1_pos, bp.pnpa_pos, bp.npa_pos, bp.total_pos
        FROM branch_pos bp
      `);
      for (const bpRow of bpRows) {
        const key = (bpRow.branch_name || '').trim().toUpperCase();
        const info = BRANCH_V2_MAP[key] || {state: 'UNMAPPED', division: 'UNMAPPED', area: 'UNMAPPED'};
        await client.query(
          `INSERT INTO v2_branch_pos (month_id, state_name, division_name, area_name, branch_name, product_name,
            regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
          ON CONFLICT (month_id, state_name, division_name, area_name, branch_name, product_name) DO NOTHING`,
          [bpRow.month_id, info.state, info.division, info.area, bpRow.branch_name, bpRow.product_name,
           bpRow.regular_pos, bpRow.sma0_pos, bpRow.sma1_pos, bpRow.pnpa_pos, bpRow.npa_pos, bpRow.total_pos]
        );
      }
    } catch (bpErr) {
      console.error("v2_branch_pos sync warning:", bpErr.message);
    }

    // Sync employee_pos → v2_employee_pos
    try {
      const { rows: epPosRows } = await client.query(`
        SELECT ep.month_id, ep.emp_id, ep.regular_pos, ep.sma0_pos, ep.sma1_pos,
               ep.pnpa_pos, ep.npa_pos, ep.total_pos
        FROM employee_pos ep WHERE ep.emp_id IN (SELECT emp_id FROM v2_employees)
      `);
      for (const epRow of epPosRows) {
        await client.query(
          `INSERT INTO v2_employee_pos (month_id, emp_id, regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
          ON CONFLICT (month_id, emp_id) DO NOTHING`,
          [epRow.month_id, epRow.emp_id, epRow.regular_pos, epRow.sma0_pos, epRow.sma1_pos,
           epRow.pnpa_pos, epRow.npa_pos, epRow.total_pos]
        );
      }
    } catch (epErr) {
      console.error("v2_employee_pos sync warning:", epErr.message);
    }

    await client.query("COMMIT");
    console.log(`V2 portfolio tables populated: ${Object.keys(statesMap).length} states, ${Object.keys(divisionsMap).length} divisions, ${Object.keys(areasMap).length} areas, ${Object.keys(branchesMap).length} branches, ${employeesInserted.size} employees, ${rows.length} performance rows`);
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    throw err;
  }
}

app.use(express.static(path.join(__dirname, "..")));

// ========== UPLOAD LOCK ==========
// Prevent concurrent uploads from corrupting schema
let _uploadInProgress = false;

app.post("/api/upload", uploadLimiter, upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  if (_uploadInProgress) {
    return res.status(409).json({ error: "Another upload is in progress. Try again shortly." });
  }
  _uploadInProgress = true;

  const client = await pool.connect();
  try {
    // Parse Excel with error handling
    let wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: "buffer" });
    } catch (parseErr) {
      return res.status(400).json({ error: "Invalid Excel file: " + parseErr.message });
    }

    const sheetNames = wb.SheetNames;
    if (!sheetNames || !sheetNames.length) {
      return res.status(400).json({ error: "Excel file has no sheets" });
    }

    // Check for duplicate sheet names
    const uniqueSheets = [...new Set(sheetNames)];
    if (uniqueSheets.length !== sheetNames.length) {
      return res.status(400).json({ error: "Excel file has duplicate sheet names" });
    }

    await client.query("BEGIN");

    // Lock to prevent concurrent schema changes
    await client.query("LOCK TABLE pg_catalog.pg_class IN ACCESS EXCLUSIVE MODE");

    // Drop and recreate tables
    await client.query("DROP TABLE IF EXISTS employee_performance");
    await client.query("DROP TABLE IF EXISTS employees");
    await client.query("DROP TABLE IF EXISTS branches");
    await client.query("DROP TABLE IF EXISTS districts");
    await client.query("DROP TABLE IF EXISTS regions");
    // product_types: drop FK constraints from all referencing tables, then drop+recreate
    await client.query("ALTER TABLE IF EXISTS daily_performance DROP CONSTRAINT IF EXISTS daily_performance_product_type_id_fkey");
    await client.query("ALTER TABLE IF EXISTS hourly_performance DROP CONSTRAINT IF EXISTS hourly_performance_product_type_id_fkey");
    await client.query("ALTER TABLE IF EXISTS v2_employee_performance DROP CONSTRAINT IF EXISTS v2_employee_performance_product_type_id_fkey");
    await client.query("DROP TABLE IF EXISTS product_types");

    await client.query(`CREATE TABLE regions (region_id SERIAL PRIMARY KEY, region_name VARCHAR(100) NOT NULL UNIQUE)`);
    await client.query(`CREATE TABLE districts (district_id SERIAL PRIMARY KEY, district_name VARCHAR(100) NOT NULL, region_id INT NOT NULL REFERENCES regions(region_id), UNIQUE(district_name, region_id))`);
    await client.query(`CREATE TABLE branches (branch_id SERIAL PRIMARY KEY, branch_name VARCHAR(100) NOT NULL, district_id INT NOT NULL REFERENCES districts(district_id), UNIQUE(branch_name, district_id))`);
    await client.query(`CREATE TABLE product_types (product_type_id SERIAL PRIMARY KEY, product_type_name VARCHAR(10) NOT NULL UNIQUE)`);
    await client.query(`CREATE TABLE employees (emp_id VARCHAR(10) PRIMARY KEY, officer_name VARCHAR(150) NOT NULL, branch_id INT NOT NULL REFERENCES branches(branch_id))`);
    await client.query(`CREATE TABLE employee_performance (
      performance_id SERIAL PRIMARY KEY,
      emp_id VARCHAR(10) NOT NULL,
      product_type_id INT NOT NULL,
      regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
      demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
      demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
      pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
      npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
      npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
      regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
      demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
      demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
      pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
      UNIQUE(emp_id, product_type_id)
    )`);

    // Recreate indexes
    await client.query("CREATE INDEX idx_employees_branch_id ON employees(branch_id)");
    await client.query("CREATE INDEX idx_ep_emp_id ON employee_performance(emp_id)");
    await client.query("CREATE INDEX idx_ep_product_type_id ON employee_performance(product_type_id)");
    await client.query("CREATE INDEX idx_branches_district_id ON branches(district_id)");
    await client.query("CREATE INDEX idx_districts_region_id ON districts(region_id)");

    // Insert product types (rename VVY -> IL)
    const ptRename = { "VVY": "IL" };
    const ptMap = {};
    for (let i = 0; i < sheetNames.length; i++) {
      const ptName = ptRename[sheetNames[i]] || sheetNames[i];
      const r = await client.query("INSERT INTO product_types (product_type_name) VALUES ($1) RETURNING product_type_id", [ptName]);
      ptMap[sheetNames[i]] = r.rows[0].product_type_id;
    }

    // Parse all sheets
    const regions = {};
    const districts = {};
    const branches = {};
    const employees = {};
    let regionId = 0, districtId = 0, branchId = 0;
    let skippedRows = 0;

    for (const sheetName of sheetNames) {
      const ws = wb.Sheets[sheetName];
      if (!ws) { skippedRows++; continue; }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Auto-detect column layout from header row
      // Old format (5 cols): Region, District, Branch, Emp ID, Officer Name, ...metrics
      // New format (6 cols): Region, Division, Area, Branch, Emp ID, Officer Name, ...metrics
      const header = rows[0] || [];
      let colRegion = 0, colDistrict = 1, colBranch = 2, colEmpId = 3, colOfficer = 4, colMetricsStart = 5;
      const headerStr = header.map(h => String(h || "").toLowerCase().trim());
      const empIdIdx = headerStr.findIndex(h => h === "emp id" || h === "empid" || h === "emp_id");
      if (empIdIdx >= 0) {
        colEmpId = empIdIdx;
        colOfficer = empIdIdx + 1;
        colMetricsStart = empIdIdx + 2;
        // Work backwards: branch is right before emp id, district before that, region before that
        colBranch = empIdIdx - 1;
        colDistrict = empIdIdx - 2;
        colRegion = empIdIdx >= 4 ? 0 : 0; // Region is always first
      }

      // Build header-name → column-index map for metrics. Falls back to the
      // positional layout (colMetricsStart + i) when the header row is
      // missing or under-mapped (<20 of 25 cols recognised).
      const metricColMap = buildEodMetricColMap(header);
      const metricMappedCount = Object.keys(metricColMap).length;
      const useMetricHeaderMap = metricMappedCount >= 20;
      if (!useMetricHeaderMap) {
        const missing = EOD_METRIC_DB_COLS.filter(c => metricColMap[c] === undefined);
        console.warn(`/api/upload sheet "${sheetName}" metric header missing/ambiguous (${metricMappedCount}/25), falling back to positional for: ${missing.join(',')}`);
      }

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[colRegion] || !row[colEmpId]) { skippedRows++; continue; }

        const regionName = normalizeRegion(String(row[colRegion]));
        const districtName = String(row[colDistrict] || "").trim();
        const branchName = String(row[colBranch] || "").trim();
        const empId = String(row[colEmpId]).trim();
        const officerName = String(row[colOfficer] || "").trim();

        if (!regionName || !empId) { skippedRows++; continue; }

        // Region
        if (!regions[regionName]) {
          regionId++;
          await client.query("INSERT INTO regions (region_id, region_name) VALUES ($1, $2)", [regionId, regionName]);
          regions[regionName] = regionId;
        }

        // District
        const distKey = districtName + "|" + regionName;
        if (!districts[distKey]) {
          districtId++;
          await client.query("INSERT INTO districts (district_id, district_name, region_id) VALUES ($1, $2, $3)", [districtId, districtName, regions[regionName]]);
          districts[distKey] = districtId;
        }

        // Branch
        const brKey = branchName + "|" + districtName + "|" + regionName;
        if (!branches[brKey]) {
          branchId++;
          await client.query("INSERT INTO branches (branch_id, branch_name, district_id) VALUES ($1, $2, $3)", [branchId, branchName, districts[distKey]]);
          branches[brKey] = branchId;
        }

        // Employee — first occurrence wins
        if (!employees[empId]) {
          await client.query("INSERT INTO employees (emp_id, officer_name, branch_id) VALUES ($1, $2, $3)", [empId, officerName, branches[brKey]]);
          employees[empId] = true;
        }

        // Performance metrics — map each DB column by Excel header name
        // (or fall back to positional read when the header row is unusable).
        const metrics = readEodMetrics(row, metricColMap, useMetricHeaderMap, colMetricsStart);

        await client.query(
          `INSERT INTO employee_performance (emp_id, product_type_id,
            regular_demand, regular_collection, demand_1_30, collection_1_30,
            demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)
          ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
            regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
            demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
            demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
            pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
            npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
            npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
            on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
            regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
            demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
            demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
            pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
            on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`,
          [empId, ptMap[sheetName], ...metrics]
        );
      }
    }

    // Grant access
    await client.query('GRANT ALL ON ALL TABLES IN SCHEMA public TO "Raghunandan1157"');
    await client.query('GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO "Raghunandan1157"');

    await client.query("COMMIT");

    // After EOD upload: ensure ALL product type exists and reset hourly_performance
    // so the Hourly tab remains functional (avoids broken JOIN on stale product_type_id)
    try {
      await pool.query("INSERT INTO product_types (product_type_name) VALUES ('ALL') ON CONFLICT (product_type_name) DO NOTHING");
      await pool.query("DROP TABLE IF EXISTS hourly_performance");
      await pool.query(`CREATE TABLE hourly_performance (
        performance_id SERIAL PRIMARY KEY,
        emp_id VARCHAR(10) NOT NULL,
        product_type_id INT NOT NULL,
        regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
        demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
        demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
        pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
        npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
        npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
        on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
        regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
        demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
        demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
        pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
        on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
        UNIQUE(emp_id, product_type_id)
      )`);
      await pool.query("CREATE INDEX idx_hp_emp_id ON hourly_performance(emp_id)");
      await pool.query('GRANT ALL ON ALL TABLES IN SCHEMA public TO "Raghunandan1157"');
      await pool.query('GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO "Raghunandan1157"');
      console.log("hourly_performance reset after EOD upload");
    } catch (hrErr) {
      console.error("hourly_performance reset warning:", hrErr.message);
    }

    // Save the uploaded file to disk so frontend can download it
    try {
      if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
      fs.writeFileSync(path.join(UPLOAD_DIR, "report.xlsx"), req.file.buffer);
    } catch (fsErr) {
      console.error("File save warning:", fsErr.message);
    }

    const empCount = Object.keys(employees).length;
    const perfCount = (await pool.query("SELECT count(*) FROM employee_performance")).rows[0].count;

    // Auto-populate v2 tables (non-blocking: failures log but don't affect upload response)
    try {
      await populateV2Tables(client);
    } catch (v2Err) {
      console.error("V2 table population error (non-fatal):", v2Err.message);
    }

    res.json({
      success: true,
      sheets: sheetNames,
      regions: Object.keys(regions).length,
      districts: Object.keys(districts).length,
      branches: Object.keys(branches).length,
      employees: empCount,
      performance: parseInt(perfCount),
      skipped_rows: skippedRows,
    });
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Upload error:", err);
    res.status(500).json({ error: "Upload failed: " + err.message });
  } finally {
    _uploadInProgress = false;
    client.release();
  }
});

// ========== NPA ACTIVATION API ==========
let _npaUploadInProgress = false;
let _npaAmountOverflowCount = 0;
const NPA_NUMERIC_MAX = 999999999999.99; // limit for NUMERIC(14,2)

function _npaHeaderKey(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Normalized source header → DB column. Unmapped columns are dropped.
const NPA_HEADER_MAP = {
  identifier: "identifier",
  designation: "designation",
  memberid: "member_id",
  accountid: "account_id",
  mobile: "mobile",
  groupname: "group_name",
  productid: "product_id",
  loanamount: "loan_amount",
  loanoustanding: "loan_outstanding",
  odamount: "od_amount",
  dpddays: "dpd_bucket",
  type: "task_type",
  employee: "employee",
  title: "title",
  started: "started_at",
  completed: "completed_on",
  timetaken: "time_taken",
  priority: "priority",
  status: "status",
  presentstate: "present_state",
  customer: "customer",
  customeraddress: "customer_address",
  taskaddress: "task_address",
  checkin: "check_in_address",
  checkout: "check_out_address",
  taskcompletelatlng: "task_latlng",
  followup: "follow_up",
  nextfollowuptime: "next_follow_up_at",
  followupcomment: "follow_up_comment",
  customerloanid: "customer_loan_id",
  memberalternativecontactnumber: "alt_contact",
  memberattendence: "member_attendance",
  memberphoto: "member_photo",
  paymentstatus: "payment_status",
  promisetopaydateandtime: "promise_to_pay_at",
  paymentmode: "payment_mode",
  amount: "amount",
  receiptnumber: "receipt_number",
  utrnumber: "utr_number",
  reasonfornotcollected: "reason_not_collected",
  ifmemberabsenthousephoto: "absent_house_photo",
  transactionscreenshot: "transaction_screenshot",
  loanstatus: "loan_status",
  collectionreport: "collection_report",
  worklocation: "branch_name",
};

const NPA_BOOL_COLS = new Set(["follow_up","member_photo","absent_house_photo","transaction_screenshot"]);
const NPA_NUMERIC_COLS = new Set(["loan_amount","loan_outstanding","od_amount","amount","collection_report"]);
const NPA_BIGINT_COLS = new Set(["account_id","customer_loan_id"]);
const NPA_DATE_COLS = new Set(["completed_on"]);
const NPA_TS_COLS = new Set(["started_at","next_follow_up_at","promise_to_pay_at"]);

function _npaParseBool(v) {
  if (v == null) return null;
  const s = String(v).trim().toLowerCase();
  if (!s) return null;
  if (s === "yes" || s === "true" || s === "1" || s === "y") return true;
  if (s === "no" || s === "false" || s === "0" || s === "n") return false;
  return null;
}
function _npaParseNum(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s || s.toUpperCase() === "NA" || s === "-") return null;
  const n = Number(s.replace(/,/g, ""));
  if (!Number.isFinite(n)) return null;
  if (Math.abs(n) > NPA_NUMERIC_MAX) { _npaAmountOverflowCount++; return null; }
  return n;
}
function _npaParseBigInt(v) {
  if (v == null) return null;
  if (typeof v === "number" && Number.isFinite(v)) return String(Math.trunc(v));
  const s = String(v).trim();
  if (!s || s.toUpperCase() === "NA") return null;
  const cleaned = s.replace(/,/g, "").split(".")[0];
  if (!/^-?\d+$/.test(cleaned)) return null;
  return cleaned;
}
function _npaParseDate(v) {
  if (v == null) return null;
  if (typeof v === "number" && Number.isFinite(v)) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
  }
  const s = String(v).trim();
  if (!s) return null;
  // try ISO / common formats first
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0,10);
  // try dd-mm-yyyy or dd/mm/yyyy
  const m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
  if (m) {
    const yy = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${yy}-${String(m[2]).padStart(2,"0")}-${String(m[1]).padStart(2,"0")}`;
  }
  return null;
}
function _npaParseTs(v) {
  if (v == null) return null;
  if (typeof v === "number" && Number.isFinite(v)) {
    const dc = XLSX.SSF.parse_date_code(v);
    if (dc) return new Date(Date.UTC(dc.y, dc.m-1, dc.d, dc.H||0, dc.M||0, Math.floor(dc.S||0))).toISOString();
  }
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

app.post("/api/npa/upload", uploadLimiter, upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });
  if (_npaUploadInProgress) return res.status(409).json({ error: "Another NPA upload in progress. Try again shortly." });
  _npaUploadInProgress = true;

  const client = await pool.connect();
  try {
    let wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: "buffer", cellDates: false });
    } catch (e) {
      return res.status(400).json({ error: "Invalid Excel: " + e.message });
    }
    const sheetName = wb.SheetNames.find(n => /NPA Collection Task/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) return res.status(400).json({ error: "No sheet found" });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    if (!rows.length) return res.status(400).json({ error: "Empty sheet" });

    const header = rows[0] || [];
    const colMap = header.map(h => NPA_HEADER_MAP[_npaHeaderKey(h)] || null);

    const dbCols = [
      "run_id","branch_name","identifier","designation","member_id","account_id","mobile","group_name","product_id",
      "loan_amount","loan_outstanding","od_amount","dpd_bucket","task_type","employee","title","started_at","completed_on",
      "time_taken","priority","status","present_state","customer","customer_address","task_address","check_in_address",
      "check_out_address","task_latlng","follow_up","next_follow_up_at","follow_up_comment","customer_loan_id","alt_contact",
      "member_attendance","member_photo","payment_status","promise_to_pay_at","payment_mode","amount","receipt_number",
      "utr_number","reason_not_collected","absent_house_photo","transaction_screenshot","loan_status","collection_report"
    ];

    // Parse + shape rows in memory first, then TX + bulk insert
    const parsed = [];
    let npaCount = 0;
    let maxCompleted = null;
    const branchesSeen = new Set();
    _npaAmountOverflowCount = 0;

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.length) continue;
      const rec = {};
      for (let c = 0; c < row.length; c++) {
        const target = colMap[c];
        if (!target) continue;
        const raw = row[c];
        let val;
        if (target === "branch_name") {
          val = String(raw || "").trim().toUpperCase() || null;
        } else if (NPA_BOOL_COLS.has(target)) {
          val = _npaParseBool(raw);
        } else if (NPA_BIGINT_COLS.has(target)) {
          val = _npaParseBigInt(raw);
        } else if (NPA_NUMERIC_COLS.has(target)) {
          val = _npaParseNum(raw);
        } else if (NPA_DATE_COLS.has(target)) {
          val = _npaParseDate(raw);
        } else if (NPA_TS_COLS.has(target)) {
          val = _npaParseTs(raw);
        } else {
          val = (raw == null || String(raw).trim() === "") ? null : String(raw);
        }
        rec[target] = val;
      }
      if (!rec.branch_name || !rec.account_id) continue;
      branchesSeen.add(rec.branch_name);
      if ((rec.loan_status || "").toUpperCase() === "NPA") npaCount++;
      if (rec.completed_on && (!maxCompleted || rec.completed_on > maxCompleted)) maxCompleted = rec.completed_on;
      parsed.push(rec);
    }

    const summary = {
      final_rows: parsed.length,
      branches_seen: branchesSeen.size,
      npa_seen: npaCount,
      sheet: sheetName,
      amount_overflow_skipped: _npaAmountOverflowCount,
    };

    await client.query("BEGIN");
    const runRes = await client.query(
      `INSERT INTO npa_activation_runs (source_filename, report_date, row_count, npa_count, summary)
       VALUES ($1, $2, $3, $4, $5::jsonb) RETURNING run_id`,
      [req.file.originalname, maxCompleted, parsed.length, npaCount, JSON.stringify(summary)]
    );
    const runId = runRes.rows[0].run_id;

    const colList = dbCols.join(", ");
    const BATCH = 200;
    for (let i = 0; i < parsed.length; i += BATCH) {
      const chunk = parsed.slice(i, i + BATCH);
      const values = [];
      const params = [];
      let p = 1;
      for (const rec of chunk) {
        const ph = dbCols.map(col => {
          if (col === "run_id") { params.push(runId); return `$${p++}`; }
          const v = rec[col] === undefined ? null : rec[col];
          params.push(v);
          return `$${p++}`;
        });
        values.push(`(${ph.join(",")})`);
      }
      await client.query(
        `INSERT INTO npa_activation_rows (${colList}) VALUES ${values.join(",")}`,
        params
      );
    }

    await client.query("COMMIT");
    res.json({ run_id: runId, row_count: parsed.length, npa_count: npaCount, report_date: maxCompleted, summary });
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("NPA upload error:", err);
    res.status(500).json({ error: "NPA upload failed: " + err.message });
  } finally {
    _npaUploadInProgress = false;
    client.release();
  }
});

app.get("/api/npa/runs", async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT run_id, uploaded_at, uploaded_by, source_filename, report_date,
              row_count, npa_count, summary, notes
         FROM npa_activation_runs
         ORDER BY uploaded_at DESC LIMIT 50`
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

const _NPA_HIER_JOIN = `
  FROM npa_activation_rows nar
  LEFT JOIN v2_branches  b  ON UPPER(b.branch_name) = UPPER(nar.branch_name)
  LEFT JOIN v2_areas     a  ON b.area_id = a.area_id
  LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id
  LEFT JOIN v2_states    s  ON dv.state_id = s.state_id`;

const _NPA_METRICS = `
    COUNT(*) FILTER (WHERE nar.loan_status = 'NPA')::int AS npa_count,
    COALESCE(SUM(nar.collection_report), 0)::numeric AS collection_sum,
    COUNT(DISTINCT nar.branch_name)::int AS branch_count,
    COUNT(DISTINCT nar.account_id)::int AS account_count,
    COUNT(*)::int AS total_rows`;

function _npaBuildWhere(q) {
  const where = [];
  const params = [];
  let i = 1;
  if (q.run_id) { where.push(`nar.run_id = $${i++}`); params.push(Number(q.run_id)); }
  if (q.region) { where.push(`UPPER(TRIM(s.state_name)) = UPPER(TRIM($${i++}))`); params.push(q.region); }
  if (q.division) { where.push(`UPPER(TRIM(dv.division_name)) = UPPER(TRIM($${i++}))`); params.push(q.division); }
  if (q.area) { where.push(`UPPER(TRIM(a.area_name)) = UPPER(TRIM($${i++}))`); params.push(q.area); }
  if (q.branch) { where.push(`UPPER(TRIM(nar.branch_name)) = UPPER(TRIM($${i++}))`); params.push(q.branch); }
  if (q.loan_status) { where.push(`nar.loan_status = $${i++}`); params.push(q.loan_status); }
  if (q.from) { where.push(`nar.completed_on >= $${i++}`); params.push(q.from); }
  if (q.to) { where.push(`nar.completed_on <= $${i++}`); params.push(q.to); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/npa/by-region", async (req, res) => {
  try {
    const { clause, params } = _npaBuildWhere(req.query);
    const sql = `SELECT COALESCE(s.state_name,'(unmapped)') AS region, ${_NPA_METRICS}
                 ${_NPA_HIER_JOIN} ${clause}
                 GROUP BY COALESCE(s.state_name,'(unmapped)')
                 ORDER BY region`;
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/npa/by-division", async (req, res) => {
  try {
    const { clause, params } = _npaBuildWhere(req.query);
    const sql = `SELECT COALESCE(s.state_name,'(unmapped)') AS region,
                        COALESCE(dv.division_name,'(unmapped)') AS division, ${_NPA_METRICS}
                 ${_NPA_HIER_JOIN} ${clause}
                 GROUP BY COALESCE(s.state_name,'(unmapped)'), COALESCE(dv.division_name,'(unmapped)')
                 ORDER BY region, division`;
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/npa/by-area", async (req, res) => {
  try {
    const { clause, params } = _npaBuildWhere(req.query);
    const sql = `SELECT COALESCE(s.state_name,'(unmapped)') AS region,
                        COALESCE(dv.division_name,'(unmapped)') AS division,
                        COALESCE(a.area_name,'(unmapped)') AS area, ${_NPA_METRICS}
                 ${_NPA_HIER_JOIN} ${clause}
                 GROUP BY COALESCE(s.state_name,'(unmapped)'), COALESCE(dv.division_name,'(unmapped)'), COALESCE(a.area_name,'(unmapped)')
                 ORDER BY region, division, area`;
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/npa/by-branch", async (req, res) => {
  try {
    const { clause, params } = _npaBuildWhere(req.query);
    const sql = `SELECT COALESCE(s.state_name,'(unmapped)') AS region,
                        COALESCE(dv.division_name,'(unmapped)') AS division,
                        COALESCE(a.area_name,'(unmapped)') AS area,
                        nar.branch_name AS branch, ${_NPA_METRICS}
                 ${_NPA_HIER_JOIN} ${clause}
                 GROUP BY COALESCE(s.state_name,'(unmapped)'), COALESCE(dv.division_name,'(unmapped)'), COALESCE(a.area_name,'(unmapped)'), nar.branch_name
                 ORDER BY region, division, area, branch`;
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/npa/rows", async (req, res) => {
  try {
    const { clause, params } = _npaBuildWhere(req.query);
    const limit = Math.min(Number(req.query.limit) || 200, 2000);
    const sql = `SELECT nar.* ${_NPA_HIER_JOIN} ${clause}
                 ORDER BY nar.row_id ASC LIMIT ${limit}`;
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ========== SERVE REPORT FILE ==========
app.get("/api/report.xlsx", (req, res) => {
  const filePath = path.join(UPLOAD_DIR, "report.xlsx");
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: "No report uploaded yet" });
  }
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.sendFile(filePath);
});

// ========== EMPLOYEES API ==========
const EMP_BASE = `
  SELECT e.emp_id, e.officer_name AS name, b.branch_name AS branch,
         d.district_name AS district, r.region_name AS region
  FROM employees e
  JOIN branches b ON e.branch_id = b.branch_id
  JOIN districts d ON b.district_id = d.district_id
  JOIN regions r ON d.region_id = r.region_id`;

// Use employee_master if available, fallback to old employees table
async function hasEmployeeMaster() {
  try {
    const r = await pool.query("SELECT count(*) FROM employee_master");
    return parseInt(r.rows[0].count) > 0;
  } catch { return false; }
}

app.get("/api/employees", async (req, res) => {
  try {
    const useMaster = await hasEmployeeMaster();
    const { q } = req.query;

    if (useMaster) {
      let where = [], params = [];
      if (q) {
        const search = "%" + q + "%";
        where.push("(em.full_name ILIKE $1 OR em.emp_id ILIKE $1 OR em.branch_name ILIKE $1 OR em.area_name ILIKE $1 OR em.region_name ILIKE $1 OR em.division_name ILIKE $1 OR em.role ILIKE $1 OR em.mobile ILIKE $1)");
        params.push(search);
      }
      const clause = where.length ? " WHERE " + where.join(" AND ") : "";
      const result = await pool.query(
        `SELECT em.emp_id, em.full_name AS name, em.role, em.designation,
                em.branch_name AS branch, em.area_name AS area, em.division_name AS division,
                em.region_name AS region, em.mobile,
                em.branch_name AS location
         FROM employee_master em${clause} ORDER BY em.full_name`, params
      );
      res.json(result.rows);
    } else {
      let result;
      if (q) {
        const search = "%" + q + "%";
        result = await pool.query(
          EMP_BASE + ` WHERE e.officer_name ILIKE $1 OR e.emp_id ILIKE $1
            OR b.branch_name ILIKE $1 OR d.district_name ILIKE $1 OR r.region_name ILIKE $1
          ORDER BY e.officer_name`, [search]
        );
      } else {
        result = await pool.query(EMP_BASE + " ORDER BY e.officer_name");
      }
      res.json(result.rows);
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/employees/count", async (req, res) => {
  try {
    const useMaster = await hasEmployeeMaster();
    const table = useMaster ? "employee_master" : "employees";
    const result = await pool.query("SELECT count(*) FROM " + table);
    res.json({ count: parseInt(result.rows[0].count) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Working count per role (FO/BM/AM/DM/RM), optionally scoped to a region/division/area/branch.
app.get("/api/employees/working-counts", async (req, res) => {
  try {
    const where = ["status='Working'"];
    const params = [];
    let idx = 1;
    if (req.query.region || req.query.state) {
      where.push("TRIM(region_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(req.query.region || req.query.state);
    }
    if (req.query.division) {
      where.push("TRIM(division_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(req.query.division);
    }
    if (req.query.area || req.query.district) {
      where.push("TRIM(area_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(req.query.area || req.query.district);
    }
    if (req.query.branch) {
      where.push("UPPER(branch_name)=UPPER($" + (idx++) + ")");
      params.push(req.query.branch);
    }
    const sql = "SELECT UPPER(role) AS role, COUNT(*)::int AS cnt FROM employee_master WHERE " + where.join(" AND ") + " GROUP BY UPPER(role)";
    const r = await pool.query(sql, params);
    const out = { FO: 0, BM: 0, AM: 0, DM: 0, RM: 0 };
    for (const row of r.rows) if (out.hasOwnProperty(row.role)) out[row.role] = row.cnt;
    res.json(out);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ========== COLLECTION DATA API ==========
function buildCollectionQuery(groupBy, groupCol) {
  return `
    SELECT ${groupCol},
      SUM(ep.regular_demand)::int AS regular_demand, SUM(ep.regular_collection)::int AS regular_collection,
      SUM(ep.demand_1_30)::int AS demand_1_30, SUM(ep.collection_1_30)::int AS collection_1_30,
      SUM(ep.demand_31_60)::int AS demand_31_60, SUM(ep.collection_31_60)::int AS collection_31_60,
      SUM(ep.pnpa_demand)::int AS pnpa_demand, SUM(ep.pnpa_collection)::int AS pnpa_collection,
      SUM(ep.npa_cases)::int AS npa_cases, SUM(ep.npa_act_acc)::int AS npa_act_acc, SUM(ep.npa_act_amt) AS npa_act_amt,
      SUM(ep.npa_clo_acc)::int AS npa_clo_acc, SUM(ep.npa_clo_amt) AS npa_clo_amt,
      SUM(ep.on_date_demand)::int AS on_date_demand, SUM(ep.on_date_collection)::int AS on_date_collection,
      SUM(ep.regular_demand_amt) AS regular_demand_amt, SUM(ep.regular_collection_amt) AS regular_collection_amt,
      SUM(ep.demand_1_30_amt) AS demand_1_30_amt, SUM(ep.collection_1_30_amt) AS collection_1_30_amt,
      SUM(ep.demand_31_60_amt) AS demand_31_60_amt, SUM(ep.collection_31_60_amt) AS collection_31_60_amt,
      SUM(ep.pnpa_demand_amt) AS pnpa_demand_amt, SUM(ep.pnpa_collection_amt) AS pnpa_collection_amt,
      SUM(ep.on_date_demand_amt) AS on_date_demand_amt, SUM(ep.on_date_collection_amt) AS on_date_collection_amt
    FROM employee_performance ep
    JOIN product_types pt ON ep.product_type_id = pt.product_type_id
    JOIN employees e ON ep.emp_id = e.emp_id
    JOIN branches b ON e.branch_id = b.branch_id
    JOIN districts d ON b.district_id = d.district_id
    JOIN regions r ON d.region_id = r.region_id`;
}

function buildWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  // Hierarchy filters — resolve via employee_master by emp_id (same key as grouping)
  if (filters.region || filters.state) {
    where.push(`ep.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.region || filters.state);
  }
  if (filters.division) {
    where.push(`ep.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.division);
  }
  if (filters.district || filters.area) {
    where.push(`ep.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.district || filters.area);
  }
  if (filters.branch) {
    where.push(`ep.emp_id IN (SELECT emp_id FROM employee_master WHERE UPPER(branch_name) = UPPER($${idx++}))`);
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push(`ep.emp_id = $${idx++}`); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/collection/summary", async (req, res) => {
  try {
    if (req.query.date) {
      const base = buildDailyQuery(null, "");
      const { clause, params } = buildDailyWhere(req.query);
      const sql = base.replace("SELECT ,", "SELECT ") + clause;
      const result = await pool.query(sql, params);
      return res.json(result.rows[0] || {});
    }
    const base = buildCollectionQuery(null, "");
    const { clause, params } = buildWhere(req.query);
    // Remove leading comma from SELECT when no groupBy
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/collection/by-region", async (req, res) => {
  try {
    // Old structure: group by branches.old_region (5 regions)
    var useOld = req.query.structure === 'old';
    var regionCol = useOld ? "b.old_region AS region_name" : "r.region_name";
    var groupCol = useOld ? "b.old_region" : "r.region_name";
    var nullFilter = useOld ? " AND b.old_region IS NOT NULL" : " AND r.region_name IS NOT NULL";

    if (req.query.date) {
      const base = buildDailyQuery(regionCol);
      const { clause, params } = buildDailyWhere(req.query);
      const result = await pool.query(base + clause + nullFilter + " GROUP BY " + groupCol + " ORDER BY " + groupCol, params);
      return res.json(result.rows);
    }
    const base = buildCollectionQuery(groupCol, regionCol);
    const { clause, params } = buildWhere(req.query);
    const result = await pool.query(base + clause + nullFilter + " GROUP BY " + groupCol + " ORDER BY " + groupCol, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/collection/by-district", async (req, res) => {
  try {
    if (req.query.date) {
      const base = buildDailyQuery("d.district_name, r.region_name", "d.district_name, r.region_name");
      const { clause, params } = buildDailyWhere(req.query);
      const result = await pool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
      return res.json(result.rows);
    }
    const base = buildCollectionQuery("d.district_name, r.region_name", "d.district_name, r.region_name");
    const { clause, params } = buildWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/collection/by-branch", async (req, res) => {
  try {
    if (req.query.date) {
      const base = buildDailyQuery("b.branch_name, d.district_name", "b.branch_name, d.district_name");
      const { clause, params } = buildDailyWhere(req.query);
      const result = await pool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
      return res.json(result.rows);
    }
    const base = buildCollectionQuery("b.branch_name, d.district_name", "b.branch_name, d.district_name");
    const { clause, params } = buildWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/collection/by-employee", async (req, res) => {
  try {
    const base = buildCollectionQuery("e.emp_id, e.officer_name, b.branch_name", "e.emp_id, e.officer_name AS name, b.branch_name");
    const { clause, params } = buildWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/product_types", async (req, res) => {
  try {
    const result = await pool.query("SELECT product_type_id, product_type_name FROM product_types ORDER BY product_type_id");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/regions", async (req, res) => {
  try {
    const result = await pool.query("SELECT region_id, region_name FROM regions ORDER BY region_name");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/performance/:emp_id", async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT ep.*, pt.product_type_name
       FROM employee_performance ep
       JOIN product_types pt ON ep.product_type_id = pt.product_type_id
       WHERE ep.emp_id = $1`, [req.params.emp_id]
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ========== WebSocket with safe reconnect ==========
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

let _pgListener = null;
let _listenerRetries = 0;
const MAX_LISTENER_RETRIES = 10;

async function startPgListener() {
  // Close existing listener if any
  if (_pgListener) {
    try { await _pgListener.end(); } catch (e) {}
    _pgListener = null;
  }

  if (_listenerRetries >= MAX_LISTENER_RETRIES) {
    console.error("PG listener: max retries reached, giving up");
    return;
  }

  try {
    _pgListener = new Client(dbConfig);
    await _pgListener.connect();
    await _pgListener.query("LISTEN table_change");
    _listenerRetries = 0; // Reset on successful connect
    _pgListener.on("notification", (msg) => {
      io.emit("db_change", { table: msg.payload });
    });
    _pgListener.on("error", (err) => {
      console.error("PG listener error:", err.message);
      _listenerRetries++;
      setTimeout(startPgListener, 3000 * Math.min(_listenerRetries, 5));
    });
    console.log("PostgreSQL LISTEN active");
  } catch (err) {
    console.error("PG listener connect failed:", err.message);
    _listenerRetries++;
    setTimeout(startPgListener, 3000 * Math.min(_listenerRetries, 5));
  }
}

// Global error handlers
process.on("uncaughtException", (err) => {
  console.error("Uncaught exception:", err.message);
});
process.on("unhandledRejection", (err) => {
  console.error("Unhandled rejection:", err);
});


// ========== AWS Panel API Endpoints ==========

// Whitelisted tables for direct query
const AWS_PANEL_TABLES = ['regions', 'districts', 'branches', 'employees', 'employee_performance', 'product_types', 'hourly_performance', 'months', 'portfolio_performance', 'employee_master', 'daily_performance', 'v2_employee_performance', 'v2_employees', 'v2_branches', 'v2_areas', 'v2_divisions', 'v2_states', 'daily_reports', 'daily_reports_achievements'];

// Whitelisted base directories for file browsing
const AWS_PANEL_DIRS = ['/home/ec2-user/Coll_Db', '/var/www/html/coll-db'];

// Dashboard auth middleware
const DASHBOARD_TOKEN = process.env.DASHBOARD_TOKEN || 'colldb-admin-2024';
function dashboardAuth(req, res, next) {
  const token = req.headers['x-dashboard-token'] || req.query.token;
  if (token !== DASHBOARD_TOKEN) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}
app.use('/api/aws-panel', dashboardAuth);

// SQL rate limiter
const sqlRateLimit = {};
function sqlRateLimiter(req, res, next) {
  const ip = req.ip;
  const now = Date.now();
  if (!sqlRateLimit[ip]) sqlRateLimit[ip] = [];
  sqlRateLimit[ip] = sqlRateLimit[ip].filter(t => now - t < 60000);
  if (sqlRateLimit[ip].length >= 30) {
    return res.status(429).json({ error: 'Rate limit exceeded. Max 30 queries per minute.' });
  }
  sqlRateLimit[ip].push(now);
  next();
}

// 0a. GET /api/aws-panel/databases — list all databases
app.get('/api/aws-panel/databases', async (req, res) => {
  try {
    const result = await pool.query("SELECT datname FROM pg_database WHERE datistemplate=false ORDER BY datname");
    res.json({ databases: result.rows.map(r => r.datname) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// 0b. POST /api/aws-panel/databases — create new database
app.post('/api/aws-panel/databases', async (req, res) => {
  const { name } = req.body;
  if (!name || !/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(name)) {
    return res.status(400).json({ error: 'Invalid database name' });
  }
  try {
    await pool.query('CREATE DATABASE ' + name);
    res.json({ success: true, database: name });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// 1a. GET /api/aws-panel/tables — list all public tables
app.get('/api/aws-panel/tables', async (req, res) => {
  try {
    const result = await getPool(req.query.db).query("SELECT tablename FROM pg_tables WHERE schemaname='public' ORDER BY tablename");
    res.json({ tables: result.rows.map(r => r.tablename) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 1b. GET /api/aws-panel/tables/:tableName — query a whitelisted table
app.get('/api/aws-panel/tables/:tableName', async (req, res) => {
  const tableName = req.params.tableName;
  if (!AWS_PANEL_TABLES.includes(tableName)) {
    return res.status(403).json({ error: 'Table not allowed: ' + tableName });
  }
  try {
    const p = getPool(req.query.db);
    const limit = req.query.limit ? Math.min(parseInt(req.query.limit), 10000) : 10000;
    const offset = parseInt(req.query.offset) || 0;
    const result = await p.query('SELECT * FROM ' + tableName + ' LIMIT ' + limit + ' OFFSET ' + offset);
    const columns = result.fields.map(f => f.name);
    res.json({ columns, rows: result.rows, count: result.rowCount });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 2. POST /api/aws-panel/sql — execute read-only SQL
app.post('/api/aws-panel/sql', sqlRateLimiter, async (req, res) => {
  const sql = req.body.sql || req.body.query;
  const query = (sql || '').trim();
  if (!query) {
    return res.status(400).json({ error: 'No query provided' });
  }
  // Block destructive schema changes (allow DML: SELECT, INSERT, UPDATE, DELETE)
  const forbidden = /^(DROP|TRUNCATE|GRANT|REVOKE)/i;
  if (forbidden.test(query)) {
    return res.status(403).json({ error: 'Destructive schema changes are not allowed' });
  }
  try {
    const start = Date.now();
    const result = await getPool(req.query.db).query({ text: query, statement_timeout: 5000 });
    const time_ms = Date.now() - start;
    const columns = result.fields ? result.fields.map(f => f.name) : [];
    res.json({ columns, rows: result.rows, rowCount: result.rowCount, time_ms });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 3a. GET /api/aws-panel/files — list directory contents
app.get('/api/aws-panel/files', (req, res) => {
  const reqPath = req.query.path || '/home/ec2-user/Coll_Db';
  const resolved = path.resolve(reqPath);
  const allowed = AWS_PANEL_DIRS.some(base => resolved.startsWith(base));
  if (!allowed) {
    return res.status(403).json({ error: 'Path not allowed: ' + resolved });
  }
  try {
    const entries = fs.readdirSync(resolved).map(name => {
      try {
        const stat = fs.statSync(path.join(resolved, name));
        return { name, type: stat.isDirectory() ? 'dir' : 'file', size: stat.size, modified: stat.mtime };
      } catch (e) {
        return { name, type: 'unknown', size: 0, modified: null };
      }
    });
    res.json({ path: resolved, entries });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 3b. GET /api/aws-panel/files/download — serve a file
app.get('/api/aws-panel/files/download', (req, res) => {
  const reqPath = req.query.path || '';
  const resolved = path.resolve(reqPath);
  const allowed = AWS_PANEL_DIRS.some(base => resolved.startsWith(base));
  if (!allowed) {
    return res.status(403).json({ error: 'Path not allowed: ' + resolved });
  }
  try {
    if (!fs.existsSync(resolved) || fs.statSync(resolved).isDirectory()) {
      return res.status(404).json({ error: 'File not found or is a directory' });
    }
    res.download(resolved);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// 4. GET /api/aws-panel/status — server status
app.get('/api/aws-panel/status', async (req, res) => {
  let db_connected = false;
  let tables = [];
  try {
    const r = await pool.query("SELECT tablename FROM pg_tables WHERE schemaname='public' ORDER BY tablename");
    db_connected = true;
    tables = r.rows.map(r => r.tablename);
  } catch (e) {}
  res.json({
    uptime: process.uptime() + 's',
    node_version: process.version,
    db_connected,
    tables
  });
});


// ========== GitHub Proxy (for AWS Panel) ==========
const https = require('https');

app.get("/api/aws-panel/github/repos/:username", async (req, res) => {
  try {
    const username = req.params.username.replace(/[^a-zA-Z0-9_-]/g, '');
    const url = `https://api.github.com/users/${username}/repos?sort=updated&per_page=20`;

    const data = await new Promise((resolve, reject) => {
      https.get(url, { headers: { 'User-Agent': 'AWSPanel', 'Accept': 'application/json' } }, (resp) => {
        let body = '';
        resp.on('data', chunk => body += chunk);
        resp.on('end', () => {
          try { resolve(JSON.parse(body)); } catch(e) { reject(e); }
        });
      }).on('error', reject);
    });

    const repos = Array.isArray(data) ? data.map(r => ({
      name: r.name,
      full_name: r.full_name,
      description: r.description,
      language: r.language,
      updated_at: r.updated_at,
      html_url: r.html_url,
      default_branch: r.default_branch,
      stargazers_count: r.stargazers_count,
      private: r.private
    })) : [];

    res.json({ repos });
  } catch (err) {
    res.status(500).json({ error: "GitHub API error: " + err.message });
  }
});

app.get("/api/aws-panel/github/commits/:owner/:repo", async (req, res) => {
  try {
    const owner = req.params.owner.replace(/[^a-zA-Z0-9_-]/g, '');
    const repo = req.params.repo.replace(/[^a-zA-Z0-9_-]/g, '');
    const url = `https://api.github.com/repos/${owner}/${repo}/commits?per_page=10`;

    const data = await new Promise((resolve, reject) => {
      https.get(url, { headers: { 'User-Agent': 'AWSPanel', 'Accept': 'application/json' } }, (resp) => {
        let body = '';
        resp.on('data', chunk => body += chunk);
        resp.on('end', () => {
          try { resolve(JSON.parse(body)); } catch(e) { reject(e); }
        });
      }).on('error', reject);
    });

    const commits = Array.isArray(data) ? data.map(c => ({
      sha: c.sha ? c.sha.substring(0, 7) : '',
      message: c.commit ? c.commit.message.split('\n')[0] : '',
      author: c.commit && c.commit.author ? c.commit.author.name : '',
      date: c.commit && c.commit.author ? c.commit.author.date : '',
      html_url: c.html_url
    })) : [];

    res.json({ commits });
  } catch (err) {
    res.status(500).json({ error: "GitHub API error: " + err.message });
  }
});



// ========== Schema / Definition APIs (for AWS Panel) ==========

// GET table definition (CREATE TABLE DDL)
app.get("/api/aws-panel/definition/:tableName", async (req, res) => {
  const tbl = req.params.tableName;
  if (!AWS_PANEL_TABLES.includes(tbl)) return res.status(400).json({ error: "Table not allowed" });

  try {
    const p = getPool(req.query.db);
    const colResult = await p.query(`
      SELECT c.column_name, c.data_type, c.is_nullable, c.column_default,
             c.character_maximum_length, c.numeric_precision, c.numeric_scale,
             CASE WHEN pk.column_name IS NOT NULL THEN true ELSE false END as is_primary_key,
             CASE WHEN u.column_name IS NOT NULL THEN true ELSE false END as is_unique
      FROM information_schema.columns c
      LEFT JOIN (
        SELECT ku.column_name FROM information_schema.table_constraints tc
        JOIN information_schema.key_column_usage ku ON tc.constraint_name = ku.constraint_name
        WHERE tc.table_name = $1 AND tc.constraint_type = 'PRIMARY KEY'
      ) pk ON c.column_name = pk.column_name
      LEFT JOIN (
        SELECT ku.column_name FROM information_schema.table_constraints tc
        JOIN information_schema.key_column_usage ku ON tc.constraint_name = ku.constraint_name
        WHERE tc.table_name = $1 AND tc.constraint_type = 'UNIQUE'
      ) u ON c.column_name = u.column_name
      WHERE c.table_name = $1 AND c.table_schema = 'public'
      ORDER BY c.ordinal_position
    `, [tbl]);

    const fkResult = await p.query(`
      SELECT kcu.column_name, ccu.table_name AS foreign_table, ccu.column_name AS foreign_column
      FROM information_schema.table_constraints tc
      JOIN information_schema.key_column_usage kcu ON tc.constraint_name = kcu.constraint_name
      JOIN information_schema.constraint_column_usage ccu ON tc.constraint_name = ccu.constraint_name
      WHERE tc.table_name = $1 AND tc.constraint_type = 'FOREIGN KEY'
    `, [tbl]);

    const countResult = await p.query('SELECT COUNT(*) as count FROM ' + tbl);

    const idxResult = await p.query(`
      SELECT indexname, indexdef FROM pg_indexes WHERE tablename = $1
    `, [tbl]);

    res.json({
      table: tbl,
      columns: colResult.rows,
      foreign_keys: fkResult.rows,
      indexes: idxResult.rows,
      row_count: parseInt(countResult.rows[0].count)
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET schema overview (all tables with columns for ERD visualizer)
app.get("/api/aws-panel/schema", async (req, res) => {
  try {
    const sp = getPool(req.query.db);
    const tables = await sp.query(`
      SELECT t.table_name,
        (SELECT COUNT(*) FROM information_schema.columns c WHERE c.table_name = t.table_name AND c.table_schema = 'public') as column_count
      FROM information_schema.tables t
      WHERE t.table_schema = 'public' AND t.table_type = 'BASE TABLE'
      ORDER BY t.table_name
    `);

    const schema = [];
    for (const t of tables.rows) {
      const cols = await sp.query(`
        SELECT c.column_name, c.data_type, c.is_nullable,
          CASE WHEN pk.column_name IS NOT NULL THEN true ELSE false END as is_primary_key
        FROM information_schema.columns c
        LEFT JOIN (
          SELECT ku.column_name FROM information_schema.table_constraints tc
          JOIN information_schema.key_column_usage ku ON tc.constraint_name = ku.constraint_name
          WHERE tc.table_name = $1 AND tc.constraint_type = 'PRIMARY KEY'
        ) pk ON c.column_name = pk.column_name
        WHERE c.table_name = $1 AND c.table_schema = 'public'
        ORDER BY c.ordinal_position
      `, [t.table_name]);

      const fks = await sp.query(`
        SELECT kcu.column_name, ccu.table_name AS foreign_table, ccu.column_name AS foreign_column
        FROM information_schema.table_constraints tc
        JOIN information_schema.key_column_usage kcu ON tc.constraint_name = kcu.constraint_name
        JOIN information_schema.constraint_column_usage ccu ON tc.constraint_name = ccu.constraint_name
        WHERE tc.table_name = $1 AND tc.constraint_type = 'FOREIGN KEY'
      `, [t.table_name]);

      schema.push({
        table_name: t.table_name,
        columns: cols.rows,
        foreign_keys: fks.rows
      });
    }

    res.json({ schema });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PM2 logs endpoint
app.get('/api/aws-panel/logs', async (req, res) => {
  const { execSync } = require('child_process');
  try {
    const lines = parseInt(req.query.lines) || 50;
    const logs = execSync(`pm2 logs coll-db --lines ${Math.min(lines, 200)} --nostream 2>&1`, { timeout: 5000 }).toString();
    res.json({ logs: logs.split('\n') });
  } catch (e) {
    res.json({ logs: ['Error reading logs: ' + e.message] });
  }
});

// Deployments endpoint (git log)
app.get('/api/aws-panel/deployments', async (req, res) => {
  const { execSync } = require('child_process');
  try {
    const gitLog = execSync('cd /home/ec2-user/Coll_Db && git log --oneline -10 --format="%H|%s|%ar"', { timeout: 5000 }).toString();
    const commits = gitLog.trim().split('\n').map(line => {
      const [hash, message, time] = line.split('|');
      return { hash, message, time };
    });
    res.json({ deployments: commits });
  } catch (e) {
    res.json({ deployments: [] });
  }
});


// ========== Deploy Endpoint ==========
app.post("/api/aws-panel/deploy", dashboardAuth, async (req, res) => {
  const { repo, projectName } = req.body;
  if (!repo || !projectName) return res.status(400).json({ error: "repo and projectName required" });

  const { execSync } = require("child_process");
  const deployDir = "/var/www/html/coll-db/projects/" + projectName;
  const steps = [];

  try {
    steps.push({ step: "Cloning repository...", status: "running" });
    execSync("sudo mkdir -p /var/www/html/coll-db/projects");
    execSync("sudo rm -rf " + deployDir);
    execSync("git clone https://github.com/" + repo + ".git " + deployDir, { timeout: 30000 });
    steps.push({ step: "Repository cloned", status: "done" });

    const hasPackageJson = fs.existsSync(deployDir + "/package.json");
    if (hasPackageJson) {
      steps.push({ step: "Installing dependencies...", status: "running" });
      try {
        execSync("cd " + deployDir + " && npm install --production 2>&1", { timeout: 120000 });
        steps.push({ step: "Dependencies installed", status: "done" });
      } catch(e) {
        steps.push({ step: "npm install skipped (non-critical)", status: "warn" });
      }
    }

    execSync("sudo chown -R apache:apache " + deployDir);
    execSync("sudo chmod -R 755 " + deployDir);
    steps.push({ step: "Permissions configured", status: "done" });

    const url = "http://52.66.163.52/projects/" + projectName;
    steps.push({ step: "Deployment complete!", status: "done", url: url });

    const deploymentsFile = "/home/ec2-user/Coll_Db/data/deployments.json";
    let deployments = [];
    try { deployments = JSON.parse(fs.readFileSync(deploymentsFile, "utf8")); } catch(e) {}
    deployments.unshift({
      projectName: projectName,
      repo: repo,
      url: url,
      time: new Date().toISOString(),
      status: "Ready"
    });
    fs.writeFileSync(deploymentsFile, JSON.stringify(deployments, null, 2));

    res.json({ success: true, steps: steps, url: url, projectName: projectName });
  } catch (e) {
    steps.push({ step: "Error: " + e.message, status: "error" });
    res.status(500).json({ success: false, steps: steps, error: e.message });
  }
});

// ========== Projects List Endpoint ==========
app.get("/api/aws-panel/projects", dashboardAuth, async (req, res) => {
  const deploymentsFile = "/home/ec2-user/Coll_Db/data/deployments.json";
  let deployments = [];
  try { deployments = JSON.parse(fs.readFileSync(deploymentsFile, "utf8")); } catch(e) {}

  const projects = [{
    projectName: "Coll_Db",
    repo: "Raghunandan1157/Coll_Db",
    url: "http://52.66.163.52",
    time: new Date().toISOString(),
    status: "Ready"
  }].concat(deployments);

  res.json({ projects: projects });
});


// Delete project endpoint
app.delete('/api/aws-panel/projects/:projectName', dashboardAuth, async (req, res) => {
  const { projectName } = req.params;
  if (!projectName || projectName === 'Coll_Db') {
    return res.status(400).json({ error: 'Cannot delete the main project' });
  }
  const { execSync } = require('child_process');
  try {
    // Remove project files
    const deployDir = '/var/www/html/coll-db/projects/' + projectName;
    execSync('sudo rm -rf ' + deployDir);
    
    // Remove from deployments.json
    const deploymentsFile = '/home/ec2-user/Coll_Db/data/deployments.json';
    let deployments = [];
    try { deployments = JSON.parse(fs.readFileSync(deploymentsFile, 'utf8')); } catch(e) {}
    deployments = deployments.filter(d => d.projectName !== projectName);
    fs.writeFileSync(deploymentsFile, JSON.stringify(deployments, null, 2));
    
    res.json({ success: true, message: projectName + ' deleted' });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});
// Env vars managementapp.get("/api/aws-panel/env/:projectName", dashboardAuth, (req, res) => {  const envFile = "/home/ec2-user/Coll_Db/data/env-" + req.params.projectName + ".json";  try {    const vars = JSON.parse(fs.readFileSync(envFile, "utf8"));    res.json({ vars });  } catch(e) {    res.json({ vars: [] });  }});app.put("/api/aws-panel/env/:projectName", dashboardAuth, (req, res) => {  const envFile = "/home/ec2-user/Coll_Db/data/env-" + req.params.projectName + ".json";  try {    fs.writeFileSync(envFile, JSON.stringify(req.body.vars || [], null, 2));    res.json({ success: true });  } catch(e) {    res.status(500).json({ error: e.message });  }});
// ========== PORTFOLIO (month-wise, separate DB) ==========
const portfolioPool = new Pool({
  host: '127.0.0.1', port: 5432,
  user: 'Raghunandan1157', password: 'raghu',
  database: 'portfolio_month_wise', max: 5
});
portfolioPool.on('error', (err) => console.error('Portfolio pool error:', err.message));

// ========== FY 25-26 AGGREGATE (auto-recomputed after portfolio upload) ==========
async function recomputeFY() {
  const client = await portfolioPool.connect();
  try {
    // Create fy_performance table if not exists (same structure as employee_performance)
    await client.query(`CREATE TABLE IF NOT EXISTS fy_performance (
      performance_id SERIAL PRIMARY KEY,
      emp_id VARCHAR(10) NOT NULL,
      product_type_id INT NOT NULL,
      regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
      demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
      demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
      pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
      npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
      npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
      regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
      demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
      demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
      pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
      UNIQUE(emp_id, product_type_id)
    )`);

    // Use the LATEST month's data as FY snapshot (portfolio is point-in-time, not cumulative)
    await client.query("TRUNCATE fy_performance");

    // Find the latest month by sort_order
    const latestMonth = await client.query("SELECT month_id, month_label FROM months ORDER BY sort_order DESC LIMIT 1");
    if (!latestMonth.rows.length) {
      console.log("FY recompute: no months found, skipping");
      return;
    }
    const latestMonthId = latestMonth.rows[0].month_id;
    const latestLabel = latestMonth.rows[0].month_label;
    console.log("FY recompute: using latest month " + latestLabel + " (id=" + latestMonthId + ")");

    await client.query(`INSERT INTO fy_performance (emp_id, product_type_id,
      regular_demand, regular_collection, demand_1_30, collection_1_30,
      demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
      npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
      on_date_demand, on_date_collection,
      regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
      demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
      on_date_demand_amt, on_date_collection_amt,
      regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
    SELECT pp.emp_id, pp.product_type_id,
      pp.regular_demand, pp.regular_collection,
      pp.demand_1_30, pp.collection_1_30,
      pp.demand_31_60, pp.collection_31_60,
      pp.pnpa_demand, pp.pnpa_collection,
      pp.npa_cases, pp.npa_act_acc, pp.npa_act_amt,
      pp.npa_clo_acc, pp.npa_clo_amt,
      pp.on_date_demand, pp.on_date_collection,
      pp.regular_demand_amt, pp.regular_collection_amt,
      pp.demand_1_30_amt, pp.collection_1_30_amt,
      pp.demand_31_60_amt, pp.collection_31_60_amt,
      pp.pnpa_demand_amt, pp.pnpa_collection_amt,
      pp.on_date_demand_amt, pp.on_date_collection_amt
    FROM portfolio_performance pp
    WHERE pp.month_id = ` + latestMonthId + `
    ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
      regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
      demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
      demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
      pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
      npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
      npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
      on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
      regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
      demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
      demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
      pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
      on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt,
      regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos, sma1_pos=EXCLUDED.sma1_pos,
      pnpa_pos=EXCLUDED.pnpa_pos, npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`);

    const count = await client.query("SELECT count(*) FROM fy_performance");
    console.log("FY 25-26 recomputed: " + count.rows[0].count + " records");
  } catch (err) {
    console.error("recomputeFY error:", err.message);
  } finally {
    client.release();
  }
}


// List available months
app.get("/api/portfolio/months", async (req, res) => {
  try {
    const result = await portfolioPool.query("SELECT month_id, month_label, sort_order FROM months ORDER BY sort_order");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// List product types in portfolio DB
app.get("/api/portfolio/product_types", async (req, res) => {
  try {
    const result = await portfolioPool.query("SELECT product_type_id, product_type_name FROM product_types ORDER BY product_type_id");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

function buildPortfolioQuery(groupCol) {
  return `SELECT ${groupCol ? groupCol + ',' : ''}
    SUM(pp.regular_demand)::int AS regular_demand, SUM(pp.regular_collection)::int AS regular_collection,
    SUM(pp.regular_pos) AS regular_pos, SUM(pp.sma0_pos) AS sma0_pos, SUM(pp.sma1_pos) AS sma1_pos,
    SUM(pp.pnpa_pos) AS pnpa_pos, SUM(pp.npa_pos) AS npa_pos, SUM(pp.total_pos) AS total_pos,
    SUM(pp.demand_1_30)::int AS demand_1_30, SUM(pp.collection_1_30)::int AS collection_1_30,
    SUM(pp.demand_31_60)::int AS demand_31_60, SUM(pp.collection_31_60)::int AS collection_31_60,
    SUM(pp.pnpa_demand)::int AS pnpa_demand, SUM(pp.pnpa_collection)::int AS pnpa_collection,
    SUM(pp.npa_cases)::int AS npa_cases, SUM(pp.npa_act_acc)::int AS npa_act_acc, SUM(pp.npa_act_amt) AS npa_act_amt,
    SUM(pp.npa_clo_acc)::int AS npa_clo_acc, SUM(pp.npa_clo_amt) AS npa_clo_amt,
    SUM(pp.on_date_demand)::int AS on_date_demand, SUM(pp.on_date_collection)::int AS on_date_collection,
    SUM(pp.regular_demand_amt) AS regular_demand_amt, SUM(pp.regular_collection_amt) AS regular_collection_amt,
    SUM(pp.demand_1_30_amt) AS demand_1_30_amt, SUM(pp.collection_1_30_amt) AS collection_1_30_amt,
    SUM(pp.demand_31_60_amt) AS demand_31_60_amt, SUM(pp.collection_31_60_amt) AS collection_31_60_amt,
    SUM(pp.pnpa_demand_amt) AS pnpa_demand_amt, SUM(pp.pnpa_collection_amt) AS pnpa_collection_amt,
    SUM(pp.on_date_demand_amt) AS on_date_demand_amt, SUM(pp.on_date_collection_amt) AS on_date_collection_amt
  FROM portfolio_performance pp
  JOIN months m ON pp.month_id = m.month_id
  JOIN product_types pt ON pp.product_type_id = pt.product_type_id
  JOIN employees e ON pp.emp_id = e.emp_id
  JOIN branches b ON e.branch_id = b.branch_id
  JOIN districts d ON b.district_id = d.district_id
  JOIN regions r ON d.region_id = r.region_id`;
}

// Old-structure portfolio — portfolio DB has regions/districts/branches
// tables. Use direct case-insensitive matching (no employee_master here).
function buildPortfolioWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.month) { where.push(`m.month_label = $${idx++}`); params.push(filters.month); }
  if (filters.product_type && filters.product_type !== 'All') {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.region || filters.state) {
    where.push(`(TRIM(r.region_name) ILIKE TRIM($${idx}) OR TRIM(d.district_name) ILIKE TRIM($${idx}))`);
    params.push(filters.region || filters.state); idx++;
  }
  if (filters.district || filters.area) {
    where.push(`TRIM(d.district_name) ILIKE TRIM($${idx++})`);
    params.push(filters.district || filters.area);
  }
  if (filters.branch) {
    where.push(`TRIM(b.branch_name) ILIKE TRIM($${idx++})`);
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push(`pp.emp_id = $${idx++}`); params.push(filters.emp_id); }
  return { clause: where.length ? ' WHERE ' + where.join(' AND ') : '', params };
}

app.get("/api/portfolio/summary", async (req, res) => {
  try {
    const base = buildPortfolioQuery(null);
    const { clause, params } = buildPortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/portfolio/by-region", async (req, res) => {
  try {
    const base = buildPortfolioQuery("r.region_name");
    const { clause, params } = buildPortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/portfolio/by-district", async (req, res) => {
  try {
    const base = buildPortfolioQuery("d.district_name, r.region_name");
    const { clause, params } = buildPortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/portfolio/by-branch", async (req, res) => {
  try {
    const base = buildPortfolioQuery("b.branch_name, d.district_name");
    const { clause, params } = buildPortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/portfolio/by-employee", async (req, res) => {
  try {
    const base = buildPortfolioQuery("e.emp_id, e.officer_name, b.branch_name");
    const { clause, params } = buildPortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ========== FY 25-26 API ENDPOINTS ==========
function buildFYQuery(groupBy, groupCol) {
  return `
    SELECT ${groupCol},
      SUM(fp.regular_demand)::int AS regular_demand, SUM(fp.regular_collection)::int AS regular_collection,
      SUM(fp.regular_pos) AS regular_pos, SUM(fp.sma0_pos) AS sma0_pos, SUM(fp.sma1_pos) AS sma1_pos,
      SUM(fp.pnpa_pos) AS pnpa_pos, SUM(fp.npa_pos) AS npa_pos, SUM(fp.total_pos) AS total_pos,
      SUM(fp.demand_1_30)::int AS demand_1_30, SUM(fp.collection_1_30)::int AS collection_1_30,
      SUM(fp.demand_31_60)::int AS demand_31_60, SUM(fp.collection_31_60)::int AS collection_31_60,
      SUM(fp.pnpa_demand)::int AS pnpa_demand, SUM(fp.pnpa_collection)::int AS pnpa_collection,
      SUM(fp.npa_cases)::int AS npa_cases, SUM(fp.npa_act_acc)::int AS npa_act_acc, SUM(fp.npa_act_amt) AS npa_act_amt,
      SUM(fp.npa_clo_acc)::int AS npa_clo_acc, SUM(fp.npa_clo_amt) AS npa_clo_amt,
      SUM(fp.on_date_demand)::int AS on_date_demand, SUM(fp.on_date_collection)::int AS on_date_collection,
      SUM(fp.regular_demand_amt) AS regular_demand_amt, SUM(fp.regular_collection_amt) AS regular_collection_amt,
      SUM(fp.demand_1_30_amt) AS demand_1_30_amt, SUM(fp.collection_1_30_amt) AS collection_1_30_amt,
      SUM(fp.demand_31_60_amt) AS demand_31_60_amt, SUM(fp.collection_31_60_amt) AS collection_31_60_amt,
      SUM(fp.pnpa_demand_amt) AS pnpa_demand_amt, SUM(fp.pnpa_collection_amt) AS pnpa_collection_amt,
      SUM(fp.on_date_demand_amt) AS on_date_demand_amt, SUM(fp.on_date_collection_amt) AS on_date_collection_amt
    FROM fy_performance fp
    JOIN product_types pt ON fp.product_type_id = pt.product_type_id
    JOIN employees e ON fp.emp_id = e.emp_id
    JOIN branches b ON e.branch_id = b.branch_id
    JOIN districts d ON b.district_id = d.district_id
    JOIN regions r ON d.region_id = r.region_id`;
}

function buildFYWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push("pt.product_type_name = $" + idx++); params.push(filters.product_type);
  }
  if (filters.region) { where.push("r.region_name = $" + idx++); params.push(filters.region); }
  if (filters.district) { where.push("d.district_name = $" + idx++); params.push(filters.district); }
  if (filters.branch) { where.push("b.branch_name = $" + idx++); params.push(filters.branch); }
  if (filters.emp_id) { where.push("fp.emp_id = $" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params: params };
}

app.get("/api/fy/summary", async (req, res) => {
  try {
    const base = buildFYQuery(null, "");
    const { clause, params } = buildFYWhere(req.query);
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await portfolioPool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/fy/by-region", async (req, res) => {
  try {
    const base = buildFYQuery("r.region_name", "r.region_name");
    const { clause, params } = buildFYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/fy/by-district", async (req, res) => {
  try {
    const base = buildFYQuery("d.district_name, r.region_name", "d.district_name, r.region_name");
    const { clause, params } = buildFYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/fy/by-branch", async (req, res) => {
  try {
    const base = buildFYQuery("b.branch_name, d.district_name", "b.branch_name, d.district_name");
    const { clause, params } = buildFYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/fy/by-employee", async (req, res) => {
  try {
    const base = buildFYQuery("e.emp_id, e.officer_name, b.branch_name", "e.emp_id, e.officer_name AS name, b.branch_name");
    const { clause, params } = buildFYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});




// ========== Branch POS API ==========
app.get("/api/portfolio/pos-summary", async (req, res) => {
  try {
    const { region, district, branch, month } = req.query;
    let where = [];
    let params = [];
    let idx = 1;
    if (month) {
      where.push("month_id=(SELECT month_id FROM months WHERE month_label=$" + idx++ + ")");
      params.push(month);
    }
    if (region) { where.push("region_name=$" + idx++); params.push(region); }
    if (district) { where.push("district_name=$" + idx++); params.push(district); }
    if (branch) { where.push("branch_name=$" + idx++); params.push(branch); }
    var product_type = req.query.product_type;
    if (product_type && product_type !== 'All') { where.push("product_name=$" + idx++); params.push(product_type); }
    const clause = where.length ? " WHERE " + where.join(" AND ") : "";
    const result = await portfolioPool.query(
      "SELECT SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM branch_pos" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/portfolio/pos-by-region", async (req, res) => {
  try {
    const { month } = req.query;
    var where2 = [];
    var params2 = [];
    var pidx = 1;
    if (month) { where2.push("month_id=(SELECT month_id FROM months WHERE month_label=$" + pidx++ + ")"); params2.push(month); }
    var pt = req.query.product_type;
    if (pt && pt !== 'All') { where2.push("product_name=$" + pidx++); params2.push(pt); }
    var whereStr = where2.length ? " WHERE " + where2.join(" AND ") : "";
    const result = await portfolioPool.query(
      "SELECT region_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM branch_pos" + whereStr + " GROUP BY region_name ORDER BY region_name",
      params2
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/portfolio/pos-by-district", async (req, res) => {
  try {
    const { region, month } = req.query;
    let where = [];
    if (month) where.push("month_id=(SELECT month_id FROM months WHERE month_label='" + month.replace(/'/g,'') + "')");
    if (region) where.push("region_name=$1");
    let whereClause = where.length ? " WHERE " + where.join(" AND ") : "";
    const result = await portfolioPool.query(
      "SELECT district_name, region_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM branch_pos" + whereClause + " GROUP BY district_name, region_name ORDER BY district_name",
      region ? [region] : []
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/portfolio/pos-by-branch", async (req, res) => {
  try {
    const { district, month } = req.query;
    let where = [];
    if (month) where.push("month_id=(SELECT month_id FROM months WHERE month_label='" + month.replace(/'/g,'') + "')");
    if (district) where.push("district_name=$1");
    let whereClause = where.length ? " WHERE " + where.join(" AND ") : "";
    const result = await portfolioPool.query(
      "SELECT branch_name, district_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM branch_pos" + whereClause + " GROUP BY branch_name, district_name ORDER BY branch_name",
      district ? [district] : []
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ========== Employee POS API ==========
app.get("/api/portfolio/emp-pos", async (req, res) => {
  try {
    const { month, emp_id, region, district, branch } = req.query;
    let where = [];
    let params = [];
    let idx = 1;
    if (month) { where.push("ep.month_id=(SELECT month_id FROM months WHERE month_label=$" + idx++ + ")"); params.push(month); }
    if (emp_id) { where.push("ep.emp_id=$" + idx++); params.push(emp_id); }
    if (region) { where.push("r.region_name=$" + idx++); params.push(region); }
    if (district) { where.push("d.district_name=$" + idx++); params.push(district); }
    if (branch) { where.push("b.branch_name=$" + idx++); params.push(branch); }
    var clause = where.length ? " WHERE " + where.join(" AND ") : "";
    var result = await portfolioPool.query(
      "SELECT ep.emp_id, e.officer_name AS name, b.branch_name, " +
      "ep.regular_pos, ep.sma0_pos, ep.sma1_pos, ep.pnpa_pos, ep.npa_pos, ep.total_pos " +
      "FROM employee_pos ep " +
      "JOIN employees e ON ep.emp_id = e.emp_id " +
      "JOIN branches b ON e.branch_id = b.branch_id " +
      "JOIN districts d ON b.district_id = d.district_id " +
      "JOIN regions r ON d.region_id = r.region_id" +
      clause + " ORDER BY ep.total_pos DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/portfolio/emp-pos-summary", async (req, res) => {
  try {
    const { month, region, district, branch } = req.query;
    let where = [];
    let params = [];
    let idx = 1;
    if (month) { where.push("ep.month_id=(SELECT month_id FROM months WHERE month_label=$" + idx++ + ")"); params.push(month); }
    if (region) { where.push("r.region_name=$" + idx++); params.push(region); }
    if (district) { where.push("d.district_name=$" + idx++); params.push(district); }
    if (branch) { where.push("b.branch_name=$" + idx++); params.push(branch); }
    var clause = where.length ? " WHERE " + where.join(" AND ") : "";
    var result = await portfolioPool.query(
      "SELECT SUM(ep.regular_pos) AS regular_pos, SUM(ep.sma0_pos) AS sma0_pos, " +
      "SUM(ep.sma1_pos) AS sma1_pos, SUM(ep.pnpa_pos) AS pnpa_pos, " +
      "SUM(ep.npa_pos) AS npa_pos, SUM(ep.total_pos) AS total_pos " +
      "FROM employee_pos ep " +
      "JOIN employees e ON ep.emp_id = e.emp_id " +
      "JOIN branches b ON e.branch_id = b.branch_id " +
      "JOIN districts d ON b.district_id = d.district_id " +
      "JOIN regions r ON d.region_id = r.region_id" +
      clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});



// ========== DAILY PERFORMANCE (date-wise EOD data) ==========

// List available dates
app.get("/api/daily/dates", async (req, res) => {
  try {
    const result = await pool.query("SELECT DISTINCT report_date FROM daily_performance ORDER BY report_date DESC");
    res.json(result.rows.map(r => r.report_date));
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Check if date has data
app.get("/api/daily/check-date", async (req, res) => {
  try {
    const { date } = req.query;
    if (!date) return res.status(400).json({ error: "date required" });
    const result = await pool.query("SELECT COUNT(*) as count FROM daily_performance WHERE report_date=$1", [date]);
    res.json({ date, hasData: parseInt(result.rows[0].count) > 0, count: parseInt(result.rows[0].count) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Daily collection queries (same as collection but filtered by date)
function buildDailyQuery(groupCol, useOld) {
  // When useOld=false (default): join through V2 hierarchy tables.
  // When useOld=true: fall back to legacy v1 tables.
  var needsHierarchy = true;
  if (useOld === undefined) useOld = (groupCol && !/^(em\.|s\.|dv\.|a\.)/.test(groupCol));
  var joins = needsHierarchy
    ? (useOld
      ? ` LEFT JOIN employees e ON dp.emp_id = e.emp_id
          LEFT JOIN branches b ON e.branch_id = b.branch_id
          LEFT JOIN districts d ON b.district_id = d.district_id
          LEFT JOIN regions r ON d.region_id = r.region_id`
      : ` LEFT JOIN v2_employees e ON dp.emp_id = e.emp_id
          LEFT JOIN v2_branches b ON e.branch_id = b.branch_id
          LEFT JOIN v2_areas a ON b.area_id = a.area_id
          LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id
          LEFT JOIN v2_states s ON dv.state_id = s.state_id`)
    : '';
  return `SELECT ${groupCol ? groupCol + ',' : ''}
    SUM(dp.regular_demand)::int AS regular_demand, SUM(dp.regular_collection)::int AS regular_collection,
    SUM(dp.demand_1_30)::int AS demand_1_30, SUM(dp.collection_1_30)::int AS collection_1_30,
    SUM(dp.demand_31_60)::int AS demand_31_60, SUM(dp.collection_31_60)::int AS collection_31_60,
    SUM(dp.pnpa_demand)::int AS pnpa_demand, SUM(dp.pnpa_collection)::int AS pnpa_collection,
    SUM(dp.npa_cases)::int AS npa_cases, SUM(dp.npa_act_acc)::int AS npa_act_acc, SUM(dp.npa_act_amt) AS npa_act_amt,
    SUM(dp.npa_clo_acc)::int AS npa_clo_acc, SUM(dp.npa_clo_amt) AS npa_clo_amt,
    SUM(dp.on_date_demand)::int AS on_date_demand, SUM(dp.on_date_collection)::int AS on_date_collection,
    SUM(dp.regular_demand_amt) AS regular_demand_amt, SUM(dp.regular_collection_amt) AS regular_collection_amt,
    SUM(dp.demand_1_30_amt) AS demand_1_30_amt, SUM(dp.collection_1_30_amt) AS collection_1_30_amt,
    SUM(dp.demand_31_60_amt) AS demand_31_60_amt, SUM(dp.collection_31_60_amt) AS collection_31_60_amt,
    SUM(dp.pnpa_demand_amt) AS pnpa_demand_amt, SUM(dp.pnpa_collection_amt) AS pnpa_collection_amt,
    SUM(dp.on_date_demand_amt) AS on_date_demand_amt, SUM(dp.on_date_collection_amt) AS on_date_collection_amt
  FROM daily_performance dp` + joins;
}

function buildDailyWhere(filters, useOld) {
  var where = [];
  var params = [];
  var idx = 1;
  if (useOld === undefined) useOld = (filters.structure === 'old');
  if (filters.date) { where.push("dp.report_date=$" + idx++); params.push(filters.date); }
  if (filters.date_from) { where.push("dp.report_date >= $" + idx++); params.push(filters.date_from); }
  if (filters.date_to)   { where.push("dp.report_date <= $" + idx++); params.push(filters.date_to); }
  // Scope filter: 'fy' = FY products only, 'oa' or default = OverAll products only
  // OA product_type_ids: 1(IGL), 2(FIG), 3(IL)  |  FY: 4(IGL_FY), 5(FIG_FY), 6(VVY_FY)
  if (filters.scope === 'fy') {
    where.push("dp.product_type_id IN (4,5,6)");
  } else if (filters.scope === 'oa' || filters.date || filters.date_from || filters.date_to) {
    // Default to OA scope to avoid mixing OA+FY
    where.push("dp.product_type_id IN (1,2,3)");
  }
  if (filters.product_type && filters.product_type !== "All") {
    where.push("dp.product_type_id IN (SELECT product_type_id FROM product_types WHERE product_type_name=$" + idx++ + ")"); params.push(filters.product_type);
  }
  // Hierarchy filters — V2 direct-column when new, employee_master subquery fallback when old
  if (useOld) {
    if (filters.region || filters.state) {
      where.push("dp.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("dp.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("dp.emp_id IN (SELECT emp_id FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.district || filters.area);
    }
    if (filters.branch) {
      where.push("dp.emp_id IN (SELECT emp_id FROM employee_master WHERE UPPER(branch_name) = UPPER($" + (idx++) + "))");
      params.push(filters.branch);
    }
  } else {
    if (filters.region || filters.state) {
      where.push("TRIM(s.state_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("TRIM(dv.division_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("TRIM(a.area_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(filters.district || filters.area);
    }
    if (filters.branch) {
      where.push("UPPER(b.branch_name) = UPPER($" + (idx++) + ")");
      params.push(filters.branch);
    }
  }
  if (filters.emp_id) { where.push("dp.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/daily/summary", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const base = buildDailyQuery(null, useOld);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-region", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var regionCol = useOld ? "r.region_name" : "s.state_name";
    var groupCol = useOld ? "r.region_name" : "s.state_name";
    var nullFilter = useOld ? " AND r.region_name IS NOT NULL" : " AND s.state_name IS NOT NULL";
    const base = buildDailyQuery(regionCol, useOld);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const result = await pool.query(base + clause + nullFilter + " GROUP BY " + groupCol + " ORDER BY " + groupCol, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-district", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var col = useOld ? "d.district_name, r.region_name" : "a.area_name AS district_name, dv.division_name";
    var gcol = useOld ? "d.district_name, r.region_name" : "a.area_name, dv.division_name";
    const base = buildDailyQuery(col, useOld);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const result = await pool.query(base + clause + " GROUP BY " + gcol + " ORDER BY " + (useOld ? "d.district_name" : "a.area_name"), params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-branch", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var col = useOld ? "b.branch_name, d.district_name" : "b.branch_name, a.area_name";
    var gcol = useOld ? "b.branch_name, d.district_name" : "b.branch_name, a.area_name";
    const base = buildDailyQuery(col, useOld);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const result = await pool.query(base + clause + " GROUP BY " + gcol + " ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-employee", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var col = useOld ? "e.emp_id, e.officer_name AS name, b.branch_name" : "e.emp_id, e.officer_name AS name, b.branch_name";
    var gcol = useOld ? "e.emp_id, e.officer_name, b.branch_name" : "e.emp_id, e.officer_name, b.branch_name";
    const base = buildDailyQuery(col, useOld);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const result = await pool.query(base + clause + " GROUP BY " + gcol + " ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Hierarchy endpoints — V2 direct joins default, employee_master fallback with ?structure=old
app.get("/api/daily/by-area", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const base = useOld
      ? buildDailyQuery("em.area_name, em.division_name", true) + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id"
      : buildDailyQuery("a.area_name, dv.division_name", false);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const extra = (clause ? " AND " : " WHERE ") + (useOld ? "em.area_name IS NOT NULL" : "a.area_name IS NOT NULL");
    var gcol = useOld ? "em.area_name, em.division_name" : "a.area_name, dv.division_name";
    const result = await pool.query(base + clause + extra + " GROUP BY " + gcol + " ORDER BY " + (useOld ? "em.area_name" : "a.area_name"), params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-division", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const base = useOld
      ? buildDailyQuery("em.division_name, em.region_name", true) + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id"
      : buildDailyQuery("dv.division_name, s.state_name AS region_name", false);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const extra = (clause ? " AND " : " WHERE ") + (useOld ? "em.division_name IS NOT NULL" : "dv.division_name IS NOT NULL");
    var gcol = useOld ? "em.division_name, em.region_name" : "dv.division_name, s.state_name";
    const result = await pool.query(base + clause + extra + " GROUP BY " + gcol + " ORDER BY " + (useOld ? "em.division_name" : "dv.division_name"), params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-state", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const base = useOld
      ? buildDailyQuery("em.region_name AS state_name", true) + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id"
      : buildDailyQuery("s.state_name", false);
    const { clause, params } = buildDailyWhere(req.query, useOld);
    const extra = (clause ? " AND " : " WHERE ") + (useOld ? "em.region_name IS NOT NULL" : "s.state_name IS NOT NULL");
    var gcol = useOld ? "em.region_name" : "s.state_name";
    const result = await pool.query(base + clause + extra + " GROUP BY " + gcol + " ORDER BY " + (useOld ? "em.region_name" : "s.state_name"), params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Upload daily data
app.post("/api/upload-daily", uploadLimiter, upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });
  const reportDate = (req.body.date || '').trim();
  if (!reportDate) return res.status(400).json({ error: "date is required (YYYY-MM-DD)" });

  const client = await pool.connect();
  try {
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    if (!wb.SheetNames.length) return res.status(400).json({ error: "Excel has no sheets" });

    await client.query("BEGIN");
    // Delete existing data for this date
    await client.query("DELETE FROM daily_performance WHERE report_date=$1", [reportDate]);

    const ptRename = { VVY: "IL" };
    let inserted = 0, skipped = 0;

    for (const sheetName of wb.SheetNames) {
      if (sheetName.endsWith('_FY') || sheetName === 'POS' || sheetName === 'EMP_POS') continue;
      const ptName = ptRename[sheetName] || sheetName;
      // Get or create product type
      let ptRes = await client.query("SELECT product_type_id FROM product_types WHERE product_type_name=$1", [ptName]);
      if (!ptRes.rows.length) {
        ptRes = await client.query("INSERT INTO product_types (product_type_name) VALUES ($1) ON CONFLICT (product_type_name) DO UPDATE SET product_type_name=EXCLUDED.product_type_name RETURNING product_type_id", [ptName]);
      }
      const ptId = ptRes.rows[0].product_type_id;

      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Auto-detect column layout from header (same logic as /api/upload)
      const hdr = (rows[0] || []).map(h => String(h || "").toLowerCase().trim());
      let dEmpId = 3, dMetrics = 5;
      const eidx = hdr.findIndex(h => h === "emp id" || h === "empid" || h === "emp_id");
      if (eidx >= 0) { dEmpId = eidx; dMetrics = eidx + 2; }

      // Header-name → column-index map for metrics. Avoids the off-by-2
      // shift caused by 4 extra "1-90" cols in the EOD report.
      const dMetricColMap = buildEodMetricColMap(rows[0] || []);
      const dMetricMappedCount = Object.keys(dMetricColMap).length;
      const useDMetricHeaderMap = dMetricMappedCount >= 20;
      if (!useDMetricHeaderMap) {
        const missing = EOD_METRIC_DB_COLS.filter(c => dMetricColMap[c] === undefined);
        console.warn(`/api/upload-daily sheet "${sheetName}" metric header missing/ambiguous (${dMetricMappedCount}/25), falling back to positional for: ${missing.join(',')}`);
      }

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0] || !row[dEmpId]) { skipped++; continue; }
        const empId = String(row[dEmpId]).trim();
        if (!empId) { skipped++; continue; }

        const metrics = readEodMetrics(row, dMetricColMap, useDMetricHeaderMap, dMetrics);

        await client.query(
          `INSERT INTO daily_performance (report_date, emp_id, product_type_id,
            regular_demand, regular_collection, demand_1_30, collection_1_30,
            demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28)
          ON CONFLICT (report_date, emp_id, product_type_id) DO UPDATE SET
            regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
            demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
            demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
            pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
            npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
            npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
            on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
            regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
            demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
            demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
            pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
            on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`,
          [reportDate, empId, ptId, ...metrics]
        );
        inserted++;
      }
    }

    await client.query("COMMIT");
    res.json({ success: true, date: reportDate, inserted, skipped });
  } catch(err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Daily upload error:", err);
    res.status(500).json({ error: "Upload failed: " + err.message });
  } finally {
    client.release();
  }
});




// ========== DISBURSEMENT API ==========
// CTE that exposes monthly disbursement rows. Months present in `disbursement` come
// from there; months only available in `disbursement_daily` (e.g. the current
// month before the monthly rollup upload) are aggregated on the fly so month-mode
// queries don't return empty for the freshest month.
const DISB_CTE = `WITH d AS (
  SELECT db_month, region_name, district_name, branch_name, emp_id,
         COALESCE(officer_name,'') AS officer_name, product_name,
         disb_count, disb_amount
  FROM disbursement
  UNION ALL
  SELECT to_char(disb_date,'Mon-YY') AS db_month,
         COALESCE(region_name,'') AS region_name,
         COALESCE(district_name,'') AS district_name,
         COALESCE(branch_name,'') AS branch_name,
         COALESCE(emp_id,'') AS emp_id,
         COALESCE(officer_name,'') AS officer_name,
         product_name,
         SUM(disb_count)::int AS disb_count,
         SUM(disb_amount)::numeric AS disb_amount
  FROM disbursement_daily
  WHERE to_char(disb_date,'Mon-YY') NOT IN (SELECT DISTINCT db_month FROM disbursement)
  GROUP BY to_char(disb_date,'Mon-YY'), region_name, district_name, branch_name, emp_id, officer_name, product_name
) `;

app.get("/api/disbursement/months", async (req, res) => {
  try {
    const result = await pool.query("SELECT DISTINCT db_month FROM disbursement");
    var months = result.rows.map(r => r.db_month);
    var monthIdx = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
    function rank(label) {
      var parts = String(label||'').split('-');
      var mi = monthIdx[parts[0]] || 0;
      var yr = parseInt(parts[1], 10) || 0;
      return yr * 100 + mi;
    }
    // Union with any disbursement_daily months not yet rolled up (so Apr-26 shows even before monthly load).
    const daily = await pool.query(
      "SELECT DISTINCT to_char(disb_date, 'Mon-YY') AS db_month FROM disbursement_daily"
    );
    daily.rows.forEach(function(r) { if (r.db_month && months.indexOf(r.db_month) === -1) months.push(r.db_month); });
    months.sort(function(a,b) { return rank(a) - rank(b); });
    res.json(months);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Route all hierarchy filters through employee_master via branch_name.
// Case-insensitive, whitespace-trimmed. Works for region/division/area/branch
// regardless of casing differences between disbursement table and employee_master.
function buildDisbWhere(filters, useOld) {
  var where = [];
  var params = [];
  var idx = 1;
  if (useOld === undefined) useOld = (filters.structure === 'old');
  if (filters.month) { where.push("d.db_month=$" + idx++); params.push(filters.month); }
  if (filters.product_name && filters.product_name !== 'All') {
    where.push("d.product_name=$" + idx++); params.push(filters.product_name);
  }
  if (useOld) {
    if (filters.region || filters.state) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.district || filters.area);
    }
  } else {
    if (filters.region || filters.state) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "JOIN v2_divisions dv ON a.division_id = dv.division_id " +
        "JOIN v2_states s ON dv.state_id = s.state_id " +
        "WHERE TRIM(s.state_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "JOIN v2_divisions dv ON a.division_id = dv.division_id " +
        "WHERE TRIM(dv.division_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "WHERE TRIM(a.area_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.district || filters.area);
    }
  }
  if (filters.branch) {
    where.push("UPPER(d.branch_name) = UPPER($" + (idx++) + ")");
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push("d.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/disbursement/summary", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-product", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.product_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.product_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-region", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-district", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.district_name, d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.district_name, d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-branch", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.branch_name, d.district_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.branch_name, d.district_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-employee", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.emp_id, d.officer_name AS name, d.branch_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.emp_id, d.officer_name, d.branch_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// New-structure endpoints (join V2 hierarchy via branch_name for direct lookups)
app.get("/api/disbursement/by-state", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    if (useOld) {
      const extra = (clause ? " AND " : " WHERE ") + "em.region_name IS NOT NULL";
      const result = await pool.query(
        DISB_CTE + "SELECT em.region_name AS state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + extra + " GROUP BY em.region_name ORDER BY em.region_name", params
      );
      return res.json(result.rows);
    }
    const extra = (clause ? " AND " : " WHERE ") + "s.state_name IS NOT NULL";
    const sql = DISB_CTE +
      "SELECT s.state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id " +
      "LEFT JOIN v2_states s ON dv.state_id = s.state_id" +
      clause + extra + " GROUP BY s.state_name ORDER BY s.state_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-division", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    if (useOld) {
      const extra = (clause ? " AND " : " WHERE ") + "em.division_name IS NOT NULL";
      const result = await pool.query(
        DISB_CTE + "SELECT em.division_name, em.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + extra + " GROUP BY em.division_name, em.region_name ORDER BY total_amount DESC", params
      );
      return res.json(result.rows);
    }
    const extra = (clause ? " AND " : " WHERE ") + "dv.division_name IS NOT NULL";
    const sql = DISB_CTE +
      "SELECT dv.division_name, s.state_name AS region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id " +
      "LEFT JOIN v2_states s ON dv.state_id = s.state_id" +
      clause + extra + " GROUP BY dv.division_name, s.state_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-area", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    if (useOld) {
      const extra = (clause ? " AND " : " WHERE ") + "em.area_name IS NOT NULL";
      const result = await pool.query(
        DISB_CTE + "SELECT em.area_name, em.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + extra + " GROUP BY em.area_name, em.division_name ORDER BY total_amount DESC", params
      );
      return res.json(result.rows);
    }
    const extra = (clause ? " AND " : " WHERE ") + "a.area_name IS NOT NULL";
    const sql = DISB_CTE +
      "SELECT a.area_name, dv.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id" +
      clause + extra + " GROUP BY a.area_name, dv.division_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-month", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbWhere(req.query, useOld);
    const result = await pool.query(
      DISB_CTE + "SELECT d.db_month, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.db_month", params
    );
    // Sort by FY order
    var monthOrder = {'Apr':1,'May':2,'Jun':3,'Jul':4,'Aug':5,'Sep':6,'Oct':7,'Nov':8,'Dec':9,'Jan':10,'Feb':11,'Mar':12};
    result.rows.sort(function(a,b) {
      var ma = a.db_month.split('-')[0], mb = b.db_month.split('-')[0];
      return (monthOrder[ma]||0) - (monthOrder[mb]||0);
    });
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ========== DISBURSEMENT DAILY API (date-grain) ==========
function buildDisbDailyWhere(filters, useOld) {
  var where = [];
  var params = [];
  var idx = 1;
  if (useOld === undefined) useOld = (filters.structure === 'old');
  if (filters.from && filters.to) {
    where.push("d.disb_date BETWEEN $" + idx++ + " AND $" + idx++);
    params.push(filters.from, filters.to);
  } else if (filters.date) { where.push("d.disb_date=$" + idx++); params.push(filters.date); }
  if (filters.product_name && filters.product_name !== 'All') {
    where.push("d.product_name=$" + idx++); params.push(filters.product_name);
  }
  if (useOld) {
    if (filters.region || filters.state) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($" + (idx++) + "))");
      params.push(filters.district || filters.area);
    }
  } else {
    if (filters.region || filters.state) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "JOIN v2_divisions dv ON a.division_id = dv.division_id " +
        "JOIN v2_states s ON dv.state_id = s.state_id " +
        "WHERE TRIM(s.state_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.region || filters.state);
    }
    if (filters.division) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "JOIN v2_divisions dv ON a.division_id = dv.division_id " +
        "WHERE TRIM(dv.division_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.division);
    }
    if (filters.district || filters.area) {
      where.push("UPPER(d.branch_name) IN (" +
        "SELECT UPPER(b.branch_name) FROM v2_branches b " +
        "JOIN v2_areas a ON b.area_id = a.area_id " +
        "WHERE TRIM(a.area_name) ILIKE TRIM($" + (idx++) + ")" +
      ")");
      params.push(filters.district || filters.area);
    }
  }
  if (filters.branch) {
    where.push("UPPER(d.branch_name) = UPPER($" + (idx++) + ")");
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push("d.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/disbursement/daily/dates", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    // Allow scope filtering but ignore `date`/`from`/`to` themselves for the list
    var q = Object.assign({}, req.query); delete q.date; delete q.from; delete q.to;
    const { clause, params } = buildDisbDailyWhere(q, useOld);
    const result = await pool.query(
      "SELECT DISTINCT d.disb_date FROM disbursement_daily d" + clause + " ORDER BY d.disb_date DESC", params
    );
    res.json(result.rows.map(function(r) {
      var dt = r.disb_date;
      if (dt instanceof Date) {
        var y = dt.getFullYear(), m = String(dt.getMonth()+1).padStart(2,'0'), day = String(dt.getDate()).padStart(2,'0');
        return y + '-' + m + '-' + day;
      }
      return dt;
    }));
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/summary", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-product", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT d.product_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.product_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-region", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-district", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT d.district_name, d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.district_name, d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-branch", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT d.branch_name, d.district_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.branch_name, d.district_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-employee", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    const result = await pool.query(
      "SELECT d.emp_id, d.officer_name AS name, d.branch_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.emp_id, d.officer_name, d.branch_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Daily disbursement aggregated by state. V2 direct joins default; ?structure=old for employee_master fallback.
app.get("/api/disbursement/daily/by-state", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    if (useOld) {
      const sql = "SELECT em.region_name AS state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + (clause ? " AND " : " WHERE ") + "em.region_name IS NOT NULL " +
        "GROUP BY em.region_name ORDER BY em.region_name";
      const result = await pool.query(sql, params);
      return res.json(result.rows);
    }
    const sql =
      "SELECT s.state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id " +
      "LEFT JOIN v2_states s ON dv.state_id = s.state_id" +
      clause + (clause ? " AND " : " WHERE ") + "s.state_name IS NOT NULL " +
      "GROUP BY s.state_name ORDER BY s.state_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Daily disbursement by area. V2 direct joins default; ?structure=old for employee_master fallback.
app.get("/api/disbursement/daily/by-area", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    if (useOld) {
      const sql = "SELECT em.area_name, em.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + (clause ? " AND " : " WHERE ") + "em.area_name IS NOT NULL " +
        "GROUP BY em.area_name, em.division_name ORDER BY total_amount DESC";
      const result = await pool.query(sql, params);
      return res.json(result.rows);
    }
    const sql =
      "SELECT a.area_name, dv.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id" +
      clause + (clause ? " AND " : " WHERE ") + "a.area_name IS NOT NULL " +
      "GROUP BY a.area_name, dv.division_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-division", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    const { clause, params } = buildDisbDailyWhere(req.query, useOld);
    if (useOld) {
      const sql = "SELECT em.division_name, em.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
        "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
        clause + (clause ? " AND " : " WHERE ") + "em.division_name IS NOT NULL " +
        "GROUP BY em.division_name, em.region_name ORDER BY total_amount DESC";
      const result = await pool.query(sql, params);
      return res.json(result.rows);
    }
    const sql =
      "SELECT dv.division_name, s.state_name AS region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d " +
      "LEFT JOIN v2_branches b ON UPPER(d.branch_name) = UPPER(b.branch_name) " +
      "LEFT JOIN v2_areas a ON b.area_id = a.area_id " +
      "LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id " +
      "LEFT JOIN v2_states s ON dv.state_id = s.state_id" +
      clause + (clause ? " AND " : " WHERE ") + "dv.division_name IS NOT NULL " +
      "GROUP BY dv.division_name, s.state_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-date-range", async (req, res) => {
  try {
    const from = req.query.from, to = req.query.to;
    if (!from || !to) return res.status(400).json({ error: "from and to are required (YYYY-MM-DD)" });

    var useOld = req.query.structure === 'old';
    // Reuse scope filters — strip date so it doesn't collide with range
    var q = Object.assign({}, req.query); delete q.date; delete q.from; delete q.to;
    const { clause, params } = buildDisbDailyWhere(q, useOld);

    // Prepend the range predicates; shift existing $N by 2
    var shifted = clause.replace(/\$(\d+)/g, function(_, n) { return "$" + (parseInt(n,10) + 2); });
    var rangeClause = " WHERE d.disb_date >= $1 AND d.disb_date <= $2"
      + (shifted ? shifted.replace(/^ WHERE /, " AND ") : "");
    var allParams = [from, to].concat(params);

    const result = await pool.query(
      "SELECT d.disb_date, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount)::numeric AS total_amount FROM disbursement_daily d"
      + rangeClause + " GROUP BY d.disb_date ORDER BY d.disb_date ASC",
      allParams
    );
    res.json(result.rows.map(function(r) {
      var dt = r.disb_date;
      if (dt instanceof Date) dt = dt.toISOString().substring(0,10);
      return { disb_date: dt, total_count: r.total_count, total_amount: r.total_amount };
    }));
  } catch(e) { res.status(500).json({ error: e.message }); }
});


const PORT = 3000;

// ========== PORTFOLIO UPLOAD (admin) ==========
app.post("/api/portfolio/upload", dashboardAuth, upload.single("file"), async (req, res) => {
  const monthLabel = (req.body.month || '').trim().toUpperCase();
  if (!monthLabel) return res.status(400).json({ error: "month is required (e.g. MAR)" });
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  const client = await portfolioPool.connect();
  try {
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    if (!wb.SheetNames.length) return res.status(400).json({ error: "Excel has no sheets" });

    await client.query("BEGIN");

    // Get or create month
    let monthRes = await client.query("SELECT month_id FROM months WHERE month_label=$1", [monthLabel]);
    let monthId;
    if (monthRes.rows.length) {
      monthId = monthRes.rows[0].month_id;
      // Clear old data for this month
      await client.query("DELETE FROM portfolio_performance WHERE month_id=$1", [monthId]);
    } else {
      // Auto-assign next sort_order
      const maxSort = await client.query("SELECT COALESCE(MAX(sort_order),0)+1 AS next FROM months");
      monthRes = await client.query("INSERT INTO months (month_label, sort_order) VALUES ($1,$2) RETURNING month_id",
        [monthLabel, maxSort.rows[0].next]);
      monthId = monthRes.rows[0].month_id;
    }

    const ptRename = { VVY: "IL" };
    const regions = {}, districts = {}, branches = {}, employees = {}, ptCache = {};

    // Preload existing lookup data
    (await client.query("SELECT region_id, region_name FROM regions")).rows.forEach(r => regions[r.region_name] = r.region_id);
    (await client.query("SELECT district_id, district_name, region_id FROM districts")).rows.forEach(r => districts[r.district_name + '|' + r.region_id] = r.district_id);
    (await client.query("SELECT branch_id, branch_name, district_id FROM branches")).rows.forEach(r => branches[r.branch_name + '|' + r.district_id] = r.branch_id);
    (await client.query("SELECT emp_id FROM employees")).rows.forEach(r => employees[r.emp_id] = true);
    (await client.query("SELECT product_type_id, product_type_name FROM product_types")).rows.forEach(r => ptCache[r.product_type_name] = r.product_type_id);

    let inserted = 0, skipped = 0;

    for (const sheetName of wb.SheetNames) {
      if (sheetName.endsWith("_FY") || sheetName === "POS" || sheetName === "EMP_POS") continue; // Skip FY sheets — parsed separately
      const ptName = ptRename[sheetName] || sheetName;
      if (!ptCache[ptName]) {
        const pr = await client.query("INSERT INTO product_types (product_type_name) VALUES ($1) ON CONFLICT (product_type_name) DO UPDATE SET product_type_name=EXCLUDED.product_type_name RETURNING product_type_id", [ptName]);
        ptCache[ptName] = pr.rows[0].product_type_id;
      }
      const ptId = ptCache[ptName];

      const ws = wb.Sheets[sheetName];
      if (!ws) continue;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0] || !row[3]) { skipped++; continue; }

        const regionName = normalizeRegion(String(row[0]));
        const districtName = normalizeDistrict(String(row[1] || ''));
        const branchName = String(row[2] || '').trim();
        const empId = String(row[3]).trim();
        const officerName = String(row[4] || '').trim();
        if (!regionName || !empId) { skipped++; continue; }

        // Region
        if (!regions[regionName]) {
          const rr = await client.query("INSERT INTO regions (region_name) VALUES ($1) ON CONFLICT (region_name) DO UPDATE SET region_name=EXCLUDED.region_name RETURNING region_id", [regionName]);
          regions[regionName] = rr.rows[0].region_id;
        }
        const rid = regions[regionName];

        // District
        const dKey = districtName + '|' + rid;
        if (!districts[dKey]) {
          const dr = await client.query("INSERT INTO districts (district_name, region_id) VALUES ($1,$2) ON CONFLICT (district_name, region_id) DO UPDATE SET district_name=EXCLUDED.district_name RETURNING district_id", [districtName, rid]);
          districts[dKey] = dr.rows[0].district_id;
        }
        const did = districts[dKey];

        // Branch
        const bKey = branchName + '|' + did;
        if (!branches[bKey]) {
          const br = await client.query("INSERT INTO branches (branch_name, district_id) VALUES ($1,$2) ON CONFLICT (branch_name, district_id) DO UPDATE SET branch_name=EXCLUDED.branch_name RETURNING branch_id", [branchName, did]);
          branches[bKey] = br.rows[0].branch_id;
        }
        const bid = branches[bKey];

        // Employee
        if (!employees[empId]) {
          await client.query("INSERT INTO employees (emp_id, officer_name, branch_id) VALUES ($1,$2,$3) ON CONFLICT (emp_id) DO UPDATE SET officer_name=EXCLUDED.officer_name, branch_id=EXCLUDED.branch_id", [empId, officerName, bid]);
          employees[empId] = true;
        }

        // Metrics (cols 5-29)
        const metrics = [];
        for (let c = 5; c < 30; c++) {
          const raw = row[c];
          if (raw == null || raw === '') { metrics.push(0); continue; }
          const num = Number(raw);
          metrics.push(Number.isFinite(num) ? num : 0);
        }
        while (metrics.length < 25) metrics.push(0);
        // POS columns (cols 30-35): Regular_POS, SMA0_POS, SMA1_POS, PNPA_POS, NPA_POS, Total_POS
        const posMetrics = [];
        for (let c = 30; c < 36; c++) {
          const raw = row[c];
          if (raw == null || raw === '') { posMetrics.push(0); continue; }
          const num = Number(raw);
          posMetrics.push(Number.isFinite(num) ? num : 0);
        }
        while (posMetrics.length < 6) posMetrics.push(0);

        await client.query(
          `INSERT INTO portfolio_performance (month_id, emp_id, product_type_id,
            regular_demand, regular_collection, demand_1_30, collection_1_30,
            demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
            npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
            on_date_demand, on_date_collection,
            regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
            demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
            on_date_demand_amt, on_date_collection_amt,
            regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33,$34)
          ON CONFLICT (month_id, emp_id, product_type_id) DO UPDATE SET
            regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
            demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
            demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
            pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
            npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
            npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
            on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
            regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
            demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
            demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
            pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
            on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt,
            regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos, sma1_pos=EXCLUDED.sma1_pos,
            pnpa_pos=EXCLUDED.pnpa_pos, npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`,
          [monthId, empId, ptId, ...metrics, ...posMetrics]
        );
        inserted++;
      }
    }

    await client.query("COMMIT");
    // Parse _FY sheets into fy_performance table (if present)
    const fySheets = wb.SheetNames.filter(s => s.endsWith('_FY'));
    let fyInserted = 0;
    if (fySheets.length > 0) {
      console.log("Found _FY sheets:", fySheets.join(', '));
      try {
        await client.query("TRUNCATE fy_performance");
        for (const sheetName of fySheets) {
          const baseName = sheetName.replace('_FY', '');
          const ptName = ({ VVY: "IL" })[baseName] || baseName;
          if (!ptCache[ptName]) {
            const pr = await client.query("INSERT INTO product_types (product_type_name) VALUES ($1) ON CONFLICT (product_type_name) DO UPDATE SET product_type_name=EXCLUDED.product_type_name RETURNING product_type_id", [ptName]);
            ptCache[ptName] = pr.rows[0].product_type_id;
          }
          const ptId = ptCache[ptName];
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
          for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !row[0] || !row[3]) continue;
            const empId = String(row[3]).trim();
            if (!empId || !employees[empId]) continue;
            const metrics = [];
            for (let c = 5; c < 30; c++) {
              const raw = row[c];
              if (raw == null || raw === '') { metrics.push(0); continue; }
              const num = Number(raw);
              metrics.push(Number.isFinite(num) ? num : 0);
            }
            while (metrics.length < 25) metrics.push(0);
            // Read POS cols (30-35)
            const fyPosMetrics = [];
            for (let pc = 30; pc < 36; pc++) {
              const raw = row[pc];
              if (raw == null || raw === '') { fyPosMetrics.push(0); continue; }
              const num = Number(raw);
              fyPosMetrics.push(Number.isFinite(num) ? num : 0);
            }
            while (fyPosMetrics.length < 6) fyPosMetrics.push(0);
            await client.query(
              `INSERT INTO fy_performance (emp_id, product_type_id,
                regular_demand, regular_collection, demand_1_30, collection_1_30,
                demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
                npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
                on_date_demand, on_date_collection,
                regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
                demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
                on_date_demand_amt, on_date_collection_amt,
                regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
              VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27,$28,$29,$30,$31,$32,$33)
              ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
                regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
                demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
                demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
                pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
                npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
                npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
                on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
                regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
                demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
                demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
                pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
                on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt,
                regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos, sma1_pos=EXCLUDED.sma1_pos,
                pnpa_pos=EXCLUDED.pnpa_pos, npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`,
              [empId, ptId, ...metrics, ...fyPosMetrics]
            );
            fyInserted++;
          }
        }
        console.log("FY sheets parsed: " + fyInserted + " records into fy_performance");
      } catch (fyErr) {
        console.error("FY sheet parse error:", fyErr.message);
      }
    } else {
      // No _FY sheets — fall back to recompute from latest month
      try { await recomputeFY(); } catch (fyErr) { console.error("FY recompute after upload:", fyErr.message); }
    }

    // Parse POS sheet into branch_pos table
    console.log("POS check: sheets=" + JSON.stringify(wb.SheetNames));
    // Map header text to canonical POS bucket key. Accepts variants like
    // "Regular POS", "SMA-0", "SMA 1", "PNPA", "NPA POS", "Total".
    function detectPosKey(header) {
      var h = String(header || '').trim().toLowerCase().replace(/[\s_\-]+/g, '');
      if (!h) return null;
      if (h.indexOf('total') !== -1) return 'total_pos';
      if (h.indexOf('regular') !== -1 || h === 'reg' || h === 'regpos') return 'regular_pos';
      if (h.indexOf('sma0') !== -1 || h.indexOf('sma00') !== -1) return 'sma0_pos';
      if (h.indexOf('sma1') !== -1) return 'sma1_pos';
      if (h.indexOf('sma2') !== -1 || h.indexOf('pnpa') !== -1) return 'pnpa_pos';
      if (h.indexOf('npa') !== -1) return 'npa_pos';
      return null;
    }
    function buildPosColMap(headerRow, expectedKeys) {
      var map = {};
      if (!Array.isArray(headerRow)) return map;
      for (var i = 0; i < headerRow.length; i++) {
        var k = detectPosKey(headerRow[i]);
        if (k && expectedKeys.indexOf(k) !== -1 && map[k] === undefined) map[k] = i;
      }
      return map;
    }
    if (wb.SheetNames.includes('POS')) {
      try {
        await portfolioPool.query("DELETE FROM branch_pos WHERE month_id=$1", [monthId]);
        const posWs = wb.Sheets['POS'];
        const posRows = XLSX.utils.sheet_to_json(posWs, { header: 1 });
        const posKeys = ['regular_pos','sma0_pos','sma1_pos','pnpa_pos','npa_pos','total_pos'];
        const posColMap = buildPosColMap(posRows[0] || [], posKeys);
        const posMappedCount = Object.keys(posColMap).length;
        const posUseHeaderMap = posMappedCount >= 5;
        if (!posUseHeaderMap) {
          const missing = posKeys.filter(k => posColMap[k] === undefined);
          console.warn("POS header missing/ambiguous, falling back to positional for: " + missing.join(','));
        }
        let posInserted = 0;
        for (let r = 1; r < posRows.length; r++) {
          const row = posRows[r];
          if (!row || !row[0]) continue;
          const region = String(row[0]).trim();
          const district = String(row[1] || '').trim();
          const branch = String(row[2] || '').trim();
          const productName = String(row[3] || 'ALL').trim();
          const vals = posKeys.map(function(k, i) {
            var colIdx = posUseHeaderMap && posColMap[k] !== undefined ? posColMap[k] : (4 + i);
            var v = Number(row[colIdx]);
            return Number.isFinite(v) ? v : 0;
          });
          await portfolioPool.query(
            `INSERT INTO branch_pos (month_id, region_name, district_name, branch_name, product_name,
              regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
            ON CONFLICT (month_id, region_name, district_name, branch_name, product_name) DO UPDATE SET
              regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos,
              sma1_pos=EXCLUDED.sma1_pos, pnpa_pos=EXCLUDED.pnpa_pos,
              npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`,
            [monthId, region, district, branch, productName, ...vals]
          );
          posInserted++;
        }
        console.log("POS sheet: " + posInserted + " branches into branch_pos (headerMap=" + posUseHeaderMap + ")");
      } catch(posErr) {
        console.error("POS sheet parse error:", posErr.message);
      }
    }

    // Parse EMP_POS sheet into employee_pos table
    if (wb.SheetNames.includes('EMP_POS')) {
      try {
        await portfolioPool.query("DELETE FROM employee_pos WHERE month_id=$1", [monthId]);
        const empPosWs = wb.Sheets['EMP_POS'];
        const empPosRows = XLSX.utils.sheet_to_json(empPosWs, { header: 1 });
        const empKeys = ['regular_pos','sma0_pos','sma1_pos','pnpa_pos','npa_pos','total_pos'];
        const empColMap = buildPosColMap(empPosRows[0] || [], empKeys);
        const empMappedCount = Object.keys(empColMap).length;
        const empUseHeaderMap = empMappedCount >= 5;
        if (!empUseHeaderMap) {
          const missing = empKeys.filter(k => empColMap[k] === undefined);
          console.warn("EMP_POS header missing/ambiguous, falling back to positional for: " + missing.join(','));
        }
        let empPosInserted = 0;
        for (let r = 1; r < empPosRows.length; r++) {
          const row = empPosRows[r];
          if (!row || !row[0]) continue;
          const empId = String(row[0]).trim();
          if (!empId) continue;
          const vals = empKeys.map(function(k, i) {
            var colIdx = empUseHeaderMap && empColMap[k] !== undefined ? empColMap[k] : (1 + i);
            var v = Number(row[colIdx]);
            return Number.isFinite(v) ? v : 0;
          });
          await portfolioPool.query(
            `INSERT INTO employee_pos (month_id, emp_id, regular_pos, sma0_pos, sma1_pos, pnpa_pos, npa_pos, total_pos)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
            ON CONFLICT (month_id, emp_id) DO UPDATE SET
              regular_pos=EXCLUDED.regular_pos, sma0_pos=EXCLUDED.sma0_pos, sma1_pos=EXCLUDED.sma1_pos,
              pnpa_pos=EXCLUDED.pnpa_pos, npa_pos=EXCLUDED.npa_pos, total_pos=EXCLUDED.total_pos`,
            [monthId, empId, ...vals]
          );
          empPosInserted++;
        }
        console.log("EMP_POS sheet: " + empPosInserted + " employees into employee_pos (headerMap=" + empUseHeaderMap + ")");
      } catch(e) {
        console.error("EMP_POS parse error:", e.message);
      }
    }
    // Auto-populate v2 portfolio tables (non-blocking: failures log but don't affect upload response)
    try {
      await populateV2PortfolioTables(client);
    } catch (v2Err) {
      console.error("V2 portfolio table population error (non-fatal):", v2Err.message);
    }

    res.json({ success: true, month: monthLabel, sheets: wb.SheetNames, inserted, skipped, fyInserted });
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Portfolio upload error:", err);
    res.status(500).json({ error: "Upload failed: " + err.message });
  } finally {
    client.release();
  }
});
// ========== V2 COLLECTION API (State → Division → Area → Branch → Employee) ==========

const V2_JOIN = `
    FROM v2_employee_performance ep
    JOIN product_types pt ON ep.product_type_id = pt.product_type_id
    JOIN v2_employees e ON ep.emp_id = e.emp_id
    JOIN v2_branches b ON e.branch_id = b.branch_id
    JOIN v2_areas a ON b.area_id = a.area_id
    JOIN v2_divisions dv ON a.division_id = dv.division_id
    JOIN v2_states s ON dv.state_id = s.state_id`;

const V2_SELECT_METRICS = `
      SUM(ep.regular_demand)::int AS regular_demand, SUM(ep.regular_collection)::int AS regular_collection,
      SUM(ep.demand_1_30)::int AS demand_1_30, SUM(ep.collection_1_30)::int AS collection_1_30,
      SUM(ep.demand_31_60)::int AS demand_31_60, SUM(ep.collection_31_60)::int AS collection_31_60,
      SUM(ep.pnpa_demand)::int AS pnpa_demand, SUM(ep.pnpa_collection)::int AS pnpa_collection,
      SUM(ep.npa_cases)::int AS npa_cases, SUM(ep.npa_act_acc)::int AS npa_act_acc, SUM(ep.npa_act_amt) AS npa_act_amt,
      SUM(ep.npa_clo_acc)::int AS npa_clo_acc, SUM(ep.npa_clo_amt) AS npa_clo_amt,
      SUM(ep.on_date_demand)::int AS on_date_demand, SUM(ep.on_date_collection)::int AS on_date_collection,
      SUM(ep.regular_demand_amt) AS regular_demand_amt, SUM(ep.regular_collection_amt) AS regular_collection_amt,
      SUM(ep.demand_1_30_amt) AS demand_1_30_amt, SUM(ep.collection_1_30_amt) AS collection_1_30_amt,
      SUM(ep.demand_31_60_amt) AS demand_31_60_amt, SUM(ep.collection_31_60_amt) AS collection_31_60_amt,
      SUM(ep.pnpa_demand_amt) AS pnpa_demand_amt, SUM(ep.pnpa_collection_amt) AS pnpa_collection_amt,
      SUM(ep.on_date_demand_amt) AS on_date_demand_amt, SUM(ep.on_date_collection_amt) AS on_date_collection_amt`;

function buildV2CollectionQuery(groupCol) {
  return `SELECT ${groupCol ? groupCol + "," : ""} ${V2_SELECT_METRICS} ${V2_JOIN}`;
}

// Resolve a hierarchy filter (state, division, area, branch) directly on V2 table
// aliases (s, dv, a, b) instead of fragile employee_master subqueries.
function buildV2Where(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  // Hierarchy filters — filter directly on V2 table aliases
  if (filters.state || filters.region) {
    where.push(`s.state_name ILIKE TRIM($${idx++})`);
    params.push(filters.state || filters.region);
  }
  if (filters.division) {
    where.push(`dv.division_name ILIKE TRIM($${idx++})`);
    params.push(filters.division);
  }
  if (filters.area || filters.district) {
    where.push(`a.area_name ILIKE TRIM($${idx++})`);
    params.push(filters.area || filters.district);
  }
  if (filters.branch) {
    where.push(`UPPER(b.branch_name) = UPPER($${idx++})`);
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push(`e.emp_id = $${idx++}`); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/v2/collection/summary", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery(null) + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/collection/by-state", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery("s.state_name") + clause + " GROUP BY s.state_name ORDER BY s.state_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/collection/by-division", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery("dv.division_name, s.state_name") + clause + " GROUP BY dv.division_name, s.state_name ORDER BY dv.division_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/collection/by-area", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery("a.area_name, dv.division_name") + clause + " GROUP BY a.area_name, dv.division_name ORDER BY a.area_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/collection/by-branch", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery("b.branch_name, a.area_name") + clause + " GROUP BY b.branch_name, a.area_name ORDER BY b.branch_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/collection/by-employee", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2CollectionQuery("e.emp_id, e.officer_name AS name, b.branch_name") + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ========== V2 HOURLY API ==========

const V2_HOURLY_JOIN = `
    FROM v2_employee_performance ep
    JOIN product_types pt ON ep.product_type_id = pt.product_type_id
    JOIN v2_employees e ON ep.emp_id = e.emp_id
    JOIN v2_branches b ON e.branch_id = b.branch_id
    JOIN v2_areas a ON b.area_id = a.area_id
    JOIN v2_divisions dv ON a.division_id = dv.division_id
    JOIN v2_states s ON dv.state_id = s.state_id`;

function buildV2HourlyQuery(groupCol) {
  return `SELECT ${groupCol ? groupCol + "," : ""} ${V2_SELECT_METRICS} ${V2_HOURLY_JOIN}`;
}

app.get("/api/v2/hourly/summary", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery(null) + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/hourly/by-state", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery("s.state_name") + clause + " GROUP BY s.state_name ORDER BY s.state_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/hourly/by-division", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery("dv.division_name, s.state_name") + clause + " GROUP BY dv.division_name, s.state_name ORDER BY dv.division_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/hourly/by-area", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery("a.area_name, dv.division_name") + clause + " GROUP BY a.area_name, dv.division_name ORDER BY a.area_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/hourly/by-branch", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery("b.branch_name, a.area_name") + clause + " GROUP BY b.branch_name, a.area_name ORDER BY b.branch_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/hourly/by-employee", async (req, res) => {
  try {
    const { clause, params } = buildV2Where(req.query);
    const sql = buildV2HourlyQuery("e.emp_id, e.officer_name AS name, b.branch_name") + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ========== V2 HIERARCHY LOOKUP ENDPOINTS ==========

app.get("/api/v2/states", async (req, res) => {
  try {
    const result = await pool.query("SELECT * FROM v2_states ORDER BY state_name");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/divisions", async (req, res) => {
  try {
    const where = [];
    const params = [];
    if (req.query.state) {
      where.push(`s.state_name = $1`);
      params.push(req.query.state);
    }
    const sql = `SELECT dv.division_id, dv.division_name, s.state_name
      FROM v2_divisions dv JOIN v2_states s ON dv.state_id = s.state_id`
      + (where.length ? " WHERE " + where.join(" AND ") : "")
      + " ORDER BY dv.division_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/areas", async (req, res) => {
  try {
    const where = [];
    const params = [];
    if (req.query.division) {
      where.push(`dv.division_name = $1`);
      params.push(req.query.division);
    }
    const sql = `SELECT a.area_id, a.area_name, dv.division_name
      FROM v2_areas a JOIN v2_divisions dv ON a.division_id = dv.division_id`
      + (where.length ? " WHERE " + where.join(" AND ") : "")
      + " ORDER BY a.area_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// ========== V2 PORTFOLIO API ENDPOINTS ==========

app.get("/api/v2/portfolio/months", async (req, res) => {
  try {
    const result = await portfolioPool.query("SELECT month_label, sort_order FROM months ORDER BY sort_order");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/product_types", async (req, res) => {
  try {
    const result = await portfolioPool.query("SELECT product_type_name FROM product_types ORDER BY product_type_id");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

function buildV2PortfolioQuery(groupCol) {
  return `SELECT ${groupCol ? groupCol + ',' : ''}
    SUM(pp.regular_demand)::int AS regular_demand, SUM(pp.regular_collection)::int AS regular_collection,
    SUM(pp.regular_pos) AS regular_pos, SUM(pp.sma0_pos) AS sma0_pos, SUM(pp.sma1_pos) AS sma1_pos,
    SUM(pp.pnpa_pos) AS pnpa_pos, SUM(pp.npa_pos) AS npa_pos, SUM(pp.total_pos) AS total_pos,
    SUM(pp.demand_1_30)::int AS demand_1_30, SUM(pp.collection_1_30)::int AS collection_1_30,
    SUM(pp.demand_31_60)::int AS demand_31_60, SUM(pp.collection_31_60)::int AS collection_31_60,
    SUM(pp.pnpa_demand)::int AS pnpa_demand, SUM(pp.pnpa_collection)::int AS pnpa_collection,
    SUM(pp.npa_cases)::int AS npa_cases, SUM(pp.npa_act_acc)::int AS npa_act_acc, SUM(pp.npa_act_amt) AS npa_act_amt,
    SUM(pp.npa_clo_acc)::int AS npa_clo_acc, SUM(pp.npa_clo_amt) AS npa_clo_amt,
    SUM(pp.on_date_demand)::int AS on_date_demand, SUM(pp.on_date_collection)::int AS on_date_collection,
    SUM(pp.regular_demand_amt) AS regular_demand_amt, SUM(pp.regular_collection_amt) AS regular_collection_amt,
    SUM(pp.demand_1_30_amt) AS demand_1_30_amt, SUM(pp.collection_1_30_amt) AS collection_1_30_amt,
    SUM(pp.demand_31_60_amt) AS demand_31_60_amt, SUM(pp.collection_31_60_amt) AS collection_31_60_amt,
    SUM(pp.pnpa_demand_amt) AS pnpa_demand_amt, SUM(pp.pnpa_collection_amt) AS pnpa_collection_amt,
    SUM(pp.on_date_demand_amt) AS on_date_demand_amt, SUM(pp.on_date_collection_amt) AS on_date_collection_amt
  FROM v2_portfolio_performance pp
  JOIN product_types pt ON pp.product_type_id = pt.product_type_id
  JOIN months m ON pp.month_id = m.month_id
  JOIN v2_employees e ON pp.emp_id = e.emp_id
  JOIN v2_branches b ON e.branch_id = b.branch_id
  JOIN v2_areas a ON b.area_id = a.area_id
  JOIN v2_divisions dv ON a.division_id = dv.division_id
  JOIN v2_states s ON dv.state_id = s.state_id`;
}

// Portfolio DB has its own v2_* tables without employee_master — use direct
// case-insensitive matching. For region/state, also try division_name since
// KA sub-region names (Chitradurga, Dharwad, etc.) live as divisions in the
// portfolio DB, not as distinct states.
function buildV2PortfolioWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.month) { where.push(`m.month_label = $${idx++}`); params.push(filters.month); }
  if (filters.product_type && filters.product_type !== 'All') {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.state || filters.region) {
    where.push(`(TRIM(s.state_name) ILIKE TRIM($${idx}) OR TRIM(dv.division_name) ILIKE TRIM($${idx}))`);
    params.push(filters.state || filters.region); idx++;
  }
  if (filters.division) {
    where.push(`TRIM(dv.division_name) ILIKE TRIM($${idx++})`);
    params.push(filters.division);
  }
  if (filters.area || filters.district) {
    where.push(`TRIM(a.area_name) ILIKE TRIM($${idx++})`);
    params.push(filters.area || filters.district);
  }
  if (filters.branch) {
    where.push(`TRIM(b.branch_name) ILIKE TRIM($${idx++})`);
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push(`pp.emp_id = $${idx++}`); params.push(filters.emp_id); }
  return { clause: where.length ? ' WHERE ' + where.join(' AND ') : '', params };
}

app.get("/api/v2/portfolio/summary", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery(null);
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/by-state", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery("s.state_name");
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY s.state_name ORDER BY s.state_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/by-division", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery("dv.division_name, s.state_name");
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY dv.division_name, s.state_name ORDER BY dv.division_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/by-area", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery("a.area_name, dv.division_name");
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY a.area_name, dv.division_name ORDER BY a.area_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/by-branch", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery("b.branch_name, a.area_name");
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY b.branch_name, a.area_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/by-employee", async (req, res) => {
  try {
    const base = buildV2PortfolioQuery("e.emp_id, e.officer_name, b.branch_name");
    const { clause, params } = buildV2PortfolioWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ========== V2 Branch POS API ==========

function buildV2PosWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.month) {
    where.push(`month_id=(SELECT month_id FROM months WHERE month_label=$${idx++})`);
    params.push(filters.month);
  }
  if (filters.state) { where.push(`state_name=$${idx++}`); params.push(filters.state); }
  if (filters.division) { where.push(`division_name=$${idx++}`); params.push(filters.division); }
  if (filters.area) { where.push(`area_name=$${idx++}`); params.push(filters.area); }
  if (filters.branch) { where.push(`branch_name=$${idx++}`); params.push(filters.branch); }
  return { clause: where.length ? ' WHERE ' + where.join(' AND ') : '', params };
}

app.get("/api/v2/portfolio/pos-summary", async (req, res) => {
  try {
    const { clause, params } = buildV2PosWhere(req.query);
    const result = await portfolioPool.query(
      "SELECT SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM v2_branch_pos" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/pos-by-state", async (req, res) => {
  try {
    const { clause, params } = buildV2PosWhere(req.query);
    const result = await portfolioPool.query(
      "SELECT state_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM v2_branch_pos" + clause + " GROUP BY state_name ORDER BY state_name", params
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/pos-by-division", async (req, res) => {
  try {
    const { clause, params } = buildV2PosWhere(req.query);
    const result = await portfolioPool.query(
      "SELECT division_name, state_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM v2_branch_pos" + clause + " GROUP BY division_name, state_name ORDER BY division_name", params
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/pos-by-area", async (req, res) => {
  try {
    const { clause, params } = buildV2PosWhere(req.query);
    const result = await portfolioPool.query(
      "SELECT area_name, division_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM v2_branch_pos" + clause + " GROUP BY area_name, division_name ORDER BY area_name", params
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/portfolio/pos-by-branch", async (req, res) => {
  try {
    const { clause, params } = buildV2PosWhere(req.query);
    const result = await portfolioPool.query(
      "SELECT branch_name, area_name, SUM(regular_pos) AS regular_pos, SUM(sma0_pos) AS sma0_pos, SUM(sma1_pos) AS sma1_pos, SUM(pnpa_pos) AS pnpa_pos, SUM(npa_pos) AS npa_pos, SUM(total_pos) AS total_pos FROM v2_branch_pos" + clause + " GROUP BY branch_name, area_name ORDER BY branch_name", params
    );
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ========== V2 FY API ENDPOINTS ==========

function buildV2FYQuery(groupCol) {
  return `SELECT ${groupCol ? groupCol + ',' : ''}
    SUM(fp.regular_demand)::int AS regular_demand, SUM(fp.regular_collection)::int AS regular_collection,
    SUM(fp.regular_pos) AS regular_pos, SUM(fp.sma0_pos) AS sma0_pos, SUM(fp.sma1_pos) AS sma1_pos,
    SUM(fp.pnpa_pos) AS pnpa_pos, SUM(fp.npa_pos) AS npa_pos, SUM(fp.total_pos) AS total_pos,
    SUM(fp.demand_1_30)::int AS demand_1_30, SUM(fp.collection_1_30)::int AS collection_1_30,
    SUM(fp.demand_31_60)::int AS demand_31_60, SUM(fp.collection_31_60)::int AS collection_31_60,
    SUM(fp.pnpa_demand)::int AS pnpa_demand, SUM(fp.pnpa_collection)::int AS pnpa_collection,
    SUM(fp.npa_cases)::int AS npa_cases, SUM(fp.npa_act_acc)::int AS npa_act_acc, SUM(fp.npa_act_amt) AS npa_act_amt,
    SUM(fp.npa_clo_acc)::int AS npa_clo_acc, SUM(fp.npa_clo_amt) AS npa_clo_amt,
    SUM(fp.on_date_demand)::int AS on_date_demand, SUM(fp.on_date_collection)::int AS on_date_collection,
    SUM(fp.regular_demand_amt) AS regular_demand_amt, SUM(fp.regular_collection_amt) AS regular_collection_amt,
    SUM(fp.demand_1_30_amt) AS demand_1_30_amt, SUM(fp.collection_1_30_amt) AS collection_1_30_amt,
    SUM(fp.demand_31_60_amt) AS demand_31_60_amt, SUM(fp.collection_31_60_amt) AS collection_31_60_amt,
    SUM(fp.pnpa_demand_amt) AS pnpa_demand_amt, SUM(fp.pnpa_collection_amt) AS pnpa_collection_amt,
    SUM(fp.on_date_demand_amt) AS on_date_demand_amt, SUM(fp.on_date_collection_amt) AS on_date_collection_amt
  FROM v2_fy_performance fp
  JOIN product_types pt ON fp.product_type_id = pt.product_type_id
  JOIN v2_employees e ON fp.emp_id = e.emp_id
  JOIN v2_branches b ON e.branch_id = b.branch_id
  JOIN v2_areas a ON b.area_id = a.area_id
  JOIN v2_divisions dv ON a.division_id = dv.division_id
  JOIN v2_states s ON dv.state_id = s.state_id`;
}

function buildV2FYWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.product_type && filters.product_type !== 'All') {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.state) { where.push(`s.state_name = $${idx++}`); params.push(filters.state); }
  if (filters.division) { where.push(`dv.division_name = $${idx++}`); params.push(filters.division); }
  if (filters.area) { where.push(`a.area_name = $${idx++}`); params.push(filters.area); }
  if (filters.branch) { where.push(`b.branch_name = $${idx++}`); params.push(filters.branch); }
  if (filters.emp_id) { where.push(`fp.emp_id = $${idx++}`); params.push(filters.emp_id); }
  return { clause: where.length ? ' WHERE ' + where.join(' AND ') : '', params };
}

app.get("/api/v2/fy/summary", async (req, res) => {
  try {
    const base = buildV2FYQuery(null);
    const { clause, params } = buildV2FYWhere(req.query);
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await portfolioPool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/fy/by-state", async (req, res) => {
  try {
    const base = buildV2FYQuery("s.state_name");
    const { clause, params } = buildV2FYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY s.state_name ORDER BY s.state_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/fy/by-division", async (req, res) => {
  try {
    const base = buildV2FYQuery("dv.division_name, s.state_name");
    const { clause, params } = buildV2FYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY dv.division_name, s.state_name ORDER BY dv.division_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/fy/by-area", async (req, res) => {
  try {
    const base = buildV2FYQuery("a.area_name, dv.division_name");
    const { clause, params } = buildV2FYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY a.area_name, dv.division_name ORDER BY a.area_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/fy/by-branch", async (req, res) => {
  try {
    const base = buildV2FYQuery("b.branch_name, a.area_name");
    const { clause, params } = buildV2FYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY b.branch_name, a.area_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/v2/fy/by-employee", async (req, res) => {
  try {
    const base = buildV2FYQuery("e.emp_id, e.officer_name, b.branch_name");
    const { clause, params } = buildV2FYWhere(req.query);
    const result = await portfolioPool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ========== BULK DAILY UPLOAD (JSON array via POST body) ==========
app.post("/api/bulk-daily", async (req, res) => {
  const client = await pool.connect();
  try {
    const { date, rows, del } = req.body;
    if (!date || !rows || !rows.length) return res.status(400).json({error: 'Need date and rows[]'});

    await client.query("BEGIN");
    if (del) await client.query("DELETE FROM daily_performance WHERE report_date=$1", [date]);

    // Build single bulk INSERT with multi-row VALUES (1 query instead of N)
    const valueParts = [];
    const params = [];
    let pi = 1;
    for (const r of rows) {
      if (!r.e) continue;
      const m = r.m || [];
      while (m.length < 25) m.push(0);
      const ph = [];
      for (let k = 0; k < 28; k++) ph.push('$' + pi++);
      valueParts.push('(' + ph.join(',') + ')');
      const v = m.map((x, i) => {
        const n = parseFloat(x) || 0;
        return (i === 12 || i === 14 || i >= 17) ? n : Math.round(n);
      });
      params.push(date, r.e, parseInt(r.p), ...v);
    }

    if (valueParts.length > 0) {
      const cols = `report_date, emp_id, product_type_id,
          regular_demand, regular_collection, demand_1_30, collection_1_30,
          demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
          npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
          on_date_demand, on_date_collection,
          regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
          demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
          on_date_demand_amt, on_date_collection_amt`;
      const upsert = `ON CONFLICT (report_date, emp_id, product_type_id) DO UPDATE SET
          regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
          demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
          demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
          pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
          npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
          npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
          on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
          regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
          demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
          demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
          pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
          on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`;
      await client.query(`INSERT INTO daily_performance (${cols}) VALUES ${valueParts.join(',')} ${upsert}`, params);
    }

    await client.query("COMMIT");
    res.json({success: true, inserted: valueParts.length, date});
  } catch(err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Bulk daily error:", err.message);
    res.status(500).json({error: err.message});
  } finally {
    client.release();
  }
});

// ========== EMPLOYEE MASTER API ==========
app.post("/api/upload-master", async (req, res) => {
  const client = await pool.connect();
  try {
    const { rows, del: shouldDelete } = req.body;
    if (!rows || !rows.length) return res.status(400).json({error: 'No rows'});
    await client.query("BEGIN");
    if (shouldDelete) await client.query("DELETE FROM employee_master");

    // Bulk INSERT with multi-row VALUES
    const valParts = [], params = [];
    let pi = 1;
    for (const r of rows) {
      if (!r.emp_id) continue;
      const ph = [];
      for (let k = 0; k < 15; k++) ph.push('$' + pi++);
      valParts.push('(' + ph.join(',') + ')');
      params.push(r.emp_id, r.full_name||null, r.role||null, r.designation||null, r.branch_name||null, r.area_name||null, r.area_manager||null, r.division_name||null, r.division_manager||null, r.region_name||null, r.mobile||null, r.date_of_joining||null, r.reporting_officer_id||null, r.reporting_officer_name||null, r.status||'Working');
    }
    if (valParts.length > 0) {
      await client.query(
        `INSERT INTO employee_master (emp_id, full_name, role, designation, branch_name, area_name, area_manager, division_name, division_manager, region_name, mobile, date_of_joining, reporting_officer_id, reporting_officer_name, status)
         VALUES ${valParts.join(',')}
         ON CONFLICT (emp_id) DO UPDATE SET full_name=EXCLUDED.full_name, role=EXCLUDED.role, designation=EXCLUDED.designation, branch_name=EXCLUDED.branch_name, area_name=EXCLUDED.area_name, area_manager=EXCLUDED.area_manager, division_name=EXCLUDED.division_name, division_manager=EXCLUDED.division_manager, region_name=EXCLUDED.region_name, mobile=EXCLUDED.mobile, date_of_joining=EXCLUDED.date_of_joining, reporting_officer_id=EXCLUDED.reporting_officer_id, reporting_officer_name=EXCLUDED.reporting_officer_name, status=EXCLUDED.status`,
        params
      );
    }
    await client.query("COMMIT");
    res.json({success: true, inserted: valParts.length});
  } catch(err) {
    await client.query("ROLLBACK").catch(() => {});
    res.status(500).json({error: err.message});
  } finally { client.release(); }
});

app.get("/api/employees/search", async (req, res) => {
  try {
    const { q, branch, area, region, role } = req.query;
    let where = [], params = [], idx = 1;
    if (q) { where.push(`(em.full_name ILIKE $${idx} OR em.emp_id ILIKE $${idx})`); params.push('%'+q+'%'); idx++; }
    if (branch) { where.push(`em.branch_name = $${idx++}`); params.push(branch); }
    if (area) { where.push(`em.area_name = $${idx++}`); params.push(area); }
    if (region) { where.push(`em.region_name = $${idx++}`); params.push(region); }
    if (role) { where.push(`em.role = $${idx++}`); params.push(role); }
    const clause = where.length ? ' WHERE ' + where.join(' AND ') : '';
    const result = await pool.query(
      `SELECT em.* FROM employee_master em${clause} ORDER BY em.full_name LIMIT 200`, params
    );
    res.json(result.rows);
  } catch(err) { res.status(500).json({error: err.message}); }
});

app.get("/api/employees/master/:id", async (req, res) => {
  try {
    const result = await pool.query("SELECT * FROM employee_master WHERE emp_id=$1", [req.params.id]);
    res.json(result.rows[0] || {});
  } catch(err) { res.status(500).json({error: err.message}); }
});

// ========== STAFF UPLOAD (Excel → employee_master) ==========
app.post("/api/upload-staff", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  const client = await pool.connect();
  try {
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    // Find "Working" sheet
    const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes('working')) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) return res.status(400).json({ error: "No Working sheet found" });

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (rows.length < 3) return res.status(400).json({ error: "File has no data rows" });

    // Row 0 = column numbers, Row 1 = headers, Row 2+ = data
    const headers = rows[1].map(h => String(h || '').trim());

    // Find column indices
    function findCol(patterns) {
      for (const p of patterns) {
        const idx = headers.findIndex(h => h.toLowerCase().includes(p.toLowerCase()));
        if (idx >= 0) return idx;
      }
      return -1;
    }

    const colMap = {
      emp_id: findCol(['NMEmpId', 'EMP ID', 'Emp Id']),
      full_name: findCol(['Name(AsperAadhar)', 'Name', 'As Per Aadhaar']),
      role: findCol(['Role', 'Position']),
      designation: findCol(['Designation']),
      branch: findCol(['Branch']),
      area: findCol(['Area Name']),
      area_mgr: findCol(['Area Manager']),
      division: findCol(['Division Name']),
      division_mgr: findCol(['Division Manager']),
      region: findCol(['Region Name', 'Region']),
      mobile: findCol(['PersonalMobile', 'Personal Mobile']),
      doj: findCol(['Date of Joining']),
      rep_id: findCol(['ReportingOfficerEMPID', 'Reporting Officer EMP']),
      rep_name: findCol(['RepotingOfficerName', 'Reporting Officer Name', 'Repoting Officer']),
    };

    if (colMap.emp_id < 0) return res.status(400).json({ error: "Cannot find Employee ID column" });

    await client.query("BEGIN");

    let inserted = 0, skipped = 0;
    const seen = new Set();

    for (let r = 2; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;

      const empId = String(row[colMap.emp_id] || '').trim();
      if (!empId || seen.has(empId)) { skipped++; continue; }
      seen.add(empId);

      const get = (col) => col >= 0 && row[col] != null ? String(row[col]).trim() : null;

      // Parse date of joining
      let doj = null;
      if (colMap.doj >= 0 && row[colMap.doj]) {
        const raw = row[colMap.doj];
        if (typeof raw === 'number' && raw > 1000) {
          // Excel serial date
          const d = new Date((raw - 25569) * 86400 * 1000);
          doj = d.toISOString().substring(0, 10);
        } else {
          try { doj = new Date(raw).toISOString().substring(0, 10); } catch(e) {}
        }
      }

      await client.query(
        `INSERT INTO employee_master (emp_id, full_name, role, designation, branch_name, area_name, area_manager, division_name, division_manager, region_name, mobile, date_of_joining, reporting_officer_id, reporting_officer_name, status)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,'Working')
         ON CONFLICT (emp_id) DO UPDATE SET
           full_name=COALESCE(EXCLUDED.full_name, employee_master.full_name),
           role=COALESCE(EXCLUDED.role, employee_master.role),
           designation=COALESCE(EXCLUDED.designation, employee_master.designation),
           branch_name=COALESCE(EXCLUDED.branch_name, employee_master.branch_name),
           area_name=COALESCE(EXCLUDED.area_name, employee_master.area_name),
           area_manager=COALESCE(EXCLUDED.area_manager, employee_master.area_manager),
           division_name=COALESCE(EXCLUDED.division_name, employee_master.division_name),
           division_manager=COALESCE(EXCLUDED.division_manager, employee_master.division_manager),
           region_name=COALESCE(EXCLUDED.region_name, employee_master.region_name),
           mobile=COALESCE(EXCLUDED.mobile, employee_master.mobile),
           date_of_joining=COALESCE(EXCLUDED.date_of_joining, employee_master.date_of_joining),
           reporting_officer_id=COALESCE(EXCLUDED.reporting_officer_id, employee_master.reporting_officer_id),
           reporting_officer_name=COALESCE(EXCLUDED.reporting_officer_name, employee_master.reporting_officer_name),
           status='Working'`,
        [empId, get(colMap.full_name), get(colMap.role), get(colMap.designation),
         get(colMap.branch), get(colMap.area), get(colMap.area_mgr),
         get(colMap.division), get(colMap.division_mgr), get(colMap.region),
         get(colMap.mobile), doj, get(colMap.rep_id), get(colMap.rep_name)]
      );
      inserted++;
    }

    await client.query("COMMIT");

    // Get total count
    const countResult = await pool.query("SELECT count(*) FROM employee_master");
    const total = parseInt(countResult.rows[0].count);

    // Save file to disk
    try {
      const uploadDir = path.join(__dirname, "..", "data");
      if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
      fs.writeFileSync(path.join(uploadDir, "staff_latest.xlsx"), req.file.buffer);
    } catch(e) {}

    res.json({ success: true, inserted, skipped, total });
  } catch(err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Staff upload error:", err);
    res.status(500).json({ error: err.message });
  } finally {
    client.release();
  }
});

// Employee name map (empId -> full name from master)
app.get("/api/employees/names", async (req, res) => {
  try {
    const result = await pool.query("SELECT emp_id, full_name FROM employee_master");
    var map = {};
    for (const r of result.rows) map[r.emp_id] = r.full_name;
    res.json(map);
  } catch(err) { res.json({}); }
});

// ========== COMPARISON API ==========
app.get("/api/comparison", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var joins = useOld
      ? ` LEFT JOIN employees e ON dp.emp_id = e.emp_id
      LEFT JOIN branches b ON e.branch_id = b.branch_id
      LEFT JOIN districts d ON b.district_id = d.district_id
      LEFT JOIN regions r ON d.region_id = r.region_id`
      : ` LEFT JOIN v2_employees e ON dp.emp_id = e.emp_id
      LEFT JOIN v2_branches b ON e.branch_id = b.branch_id
      LEFT JOIN v2_areas a ON b.area_id = a.area_id
      LEFT JOIN v2_divisions dv ON a.division_id = dv.division_id
      LEFT JOIN v2_states s ON dv.state_id = s.state_id`;
    const base = `SELECT dp.report_date,
        SUM(dp.regular_demand)::int AS regular_demand,
        SUM(dp.regular_collection)::int AS regular_collection,
        SUM(dp.demand_1_30)::int AS demand_1_30,
        SUM(dp.collection_1_30)::int AS collection_1_30,
        SUM(dp.demand_31_60)::int AS demand_31_60,
        SUM(dp.collection_31_60)::int AS collection_31_60,
        SUM(dp.pnpa_demand)::int AS pnpa_demand,
        SUM(dp.pnpa_collection)::int AS pnpa_collection,
        SUM(dp.npa_cases)::int AS npa_cases,
        SUM(dp.npa_act_acc)::int AS npa_act_acc,
        SUM(dp.npa_act_amt) AS npa_act_amt
      FROM daily_performance dp` + joins;
    const { clause, params } = buildDailyWhere({...req.query, date: undefined, scope: req.query.scope || 'oa'}, useOld);
    const sql = base + clause + " GROUP BY dp.report_date ORDER BY dp.report_date";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


server.listen(PORT, "0.0.0.0", async () => {
  await startPgListener();
  console.log("Server running on http://0.0.0.0:" + PORT);
  // Recompute FY aggregate on startup
  recomputeFY().catch(err => console.error("FY startup recompute:", err.message));
});

// ========== STEP 1: Create hourly_performance table on startup ==========
(async function initHourlyTable() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS hourly_performance (
      performance_id SERIAL PRIMARY KEY,
      emp_id VARCHAR(10) NOT NULL,
      product_type_id INT NOT NULL,
      regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
      demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
      demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
      pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
      npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
      npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
      regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
      demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
      demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
      pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
      UNIQUE(emp_id, product_type_id)
    )`);
    await pool.query("CREATE INDEX IF NOT EXISTS idx_hp_emp_id ON hourly_performance(emp_id)");
    console.log("hourly_performance table ready");
  } catch (err) {
    // Table creation may fail if employees/product_types dont exist yet — thats OK
    console.log("hourly_performance init skipped (tables may not exist yet):", err.message);
  }
})();


// ========== POS Columns Migration ==========
(async function addPOSColumns() {
  try {
    // Add POS columns to portfolio_performance
    const posCols = ['regular_pos', 'sma0_pos', 'sma1_pos', 'pnpa_pos', 'npa_pos', 'total_pos'];
    for (const col of posCols) {
      try {
        await portfolioPool.query('ALTER TABLE portfolio_performance ADD COLUMN ' + col + ' DECIMAL(15,2) DEFAULT 0');
      } catch(e) { /* column already exists */ }
      try {
        await portfolioPool.query('ALTER TABLE fy_performance ADD COLUMN ' + col + ' DECIMAL(15,2) DEFAULT 0');
      } catch(e) { /* column already exists */ }
    }
    console.log("POS columns migration complete");

// ========== Branch POS table ==========
(async function initBranchPos() {
  try {
    await portfolioPool.query(`CREATE TABLE IF NOT EXISTS branch_pos (
      pos_id SERIAL PRIMARY KEY,
      region_name VARCHAR(100) NOT NULL,
      district_name VARCHAR(100) NOT NULL,
      branch_name VARCHAR(100) NOT NULL,
      regular_pos DECIMAL(15,2) DEFAULT 0,
      sma0_pos DECIMAL(15,2) DEFAULT 0,
      sma1_pos DECIMAL(15,2) DEFAULT 0,
      pnpa_pos DECIMAL(15,2) DEFAULT 0,
      npa_pos DECIMAL(15,2) DEFAULT 0,
      total_pos DECIMAL(15,2) DEFAULT 0,
      UNIQUE(region_name, district_name, branch_name)
    )`);
    console.log("branch_pos table ready");
  } catch(e) { console.log("branch_pos init:", e.message); }
})();

  } catch(err) {
    console.log("POS migration note:", err.message);
  }
})();

// ========== POS Columns Migration ==========
(async function addPOSColumns() {
  try {
    // Add POS columns to portfolio_performance
    const posCols = ['regular_pos', 'sma0_pos', 'sma1_pos', 'pnpa_pos', 'npa_pos', 'total_pos'];
    for (const col of posCols) {
      try {
        await portfolioPool.query('ALTER TABLE portfolio_performance ADD COLUMN ' + col + ' DECIMAL(15,2) DEFAULT 0');
      } catch(e) { /* column already exists */ }
      try {
        await portfolioPool.query('ALTER TABLE fy_performance ADD COLUMN ' + col + ' DECIMAL(15,2) DEFAULT 0');
      } catch(e) { /* column already exists */ }
    }
    console.log("POS columns migration complete");
  } catch(err) {
    console.log("POS migration note:", err.message);
  }
})();




// ========== STEP 2: POST /api/upload-hourly (Quick Report + Employee Report parser) ==========
let _hourlyUploadInProgress = false;

app.post("/api/upload-hourly", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });

  if (_hourlyUploadInProgress) {
    return res.status(409).json({ error: "Another hourly upload is in progress. Try again shortly." });
  }
  _hourlyUploadInProgress = true;

  const client = await pool.connect();
  try {
    // Check that employees table exists and has data
    let empCheck;
    try {
      empCheck = await client.query("SELECT count(*) FROM employees");
    } catch (e) {
      return res.status(400).json({ error: "Please upload EOD data first" });
    }
    if (parseInt(empCheck.rows[0].count) === 0) {
      return res.status(400).json({ error: "Please upload EOD data first" });
    }

    // Parse Excel
    let wb;
    try {
      wb = XLSX.read(req.file.buffer, { type: "buffer" });
    } catch (parseErr) {
      return res.status(400).json({ error: "Invalid Excel file: " + parseErr.message });
    }

    const sheetNames = wb.SheetNames;
    if (!sheetNames || !sheetNames.length) {
      return res.status(400).json({ error: "Excel file has no sheets" });
    }

    // Detect format: Quick Report has sheets like "OverAll", "OverAll_On-Date"
    // Employee Report has product-type sheets like "IGL", "FIG", "VVY"
    const isQuickReport = sheetNames.includes("OverAll");

    if (!isQuickReport) {
      // ---- LEGACY: Employee Report format (original parser) ----
      await client.query("BEGIN");
      await client.query("DROP TABLE IF EXISTS hourly_performance CASCADE");
      await client.query(`CREATE TABLE hourly_performance (
        performance_id SERIAL PRIMARY KEY,
        emp_id VARCHAR(10) NOT NULL,
        product_type_id INT NOT NULL,
        regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
        demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
        demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
        pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
        npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
        npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
        on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
        regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
        demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
        demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
        pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
        on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
        UNIQUE(emp_id, product_type_id)
      )`);
      await client.query("CREATE INDEX idx_hp_emp_id ON hourly_performance(emp_id)");

      const ptRows = await client.query("SELECT product_type_id, product_type_name FROM product_types");
      const ptMap = {};
      for (const row of ptRows.rows) { ptMap[row.product_type_name] = row.product_type_id; }
      const ptRename = { "VVY": "IL" };

      const empRows = await client.query("SELECT emp_id FROM employees");
      const empSet = new Set(empRows.rows.map(r => r.emp_id));

      let insertedCount = 0;
      let skippedRows = 0;

      for (const sheetName of sheetNames) {
        const ptName = ptRename[sheetName] || sheetName;
        const ptId = ptMap[ptName];
        if (!ptId) { skippedRows++; continue; }
        const ws = wb.Sheets[sheetName];
        if (!ws) { skippedRows++; continue; }
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

        // Header-name → column-index map for metrics. Excludes the 4 extra
        // "1-90 Demand/Collection" buckets and avoids the off-by-2 positional
        // shift. Falls back to positional (col 5 + i) only if the header row
        // is missing/unrecognisable. Same fix as /api/upload.
        const hMetricColMap = buildEodMetricColMap(rows[0] || []);
        const hMetricMappedCount = Object.keys(hMetricColMap).length;
        const useHMetricHeaderMap = hMetricMappedCount >= 20;
        if (!useHMetricHeaderMap) {
          const missing = EOD_METRIC_DB_COLS.filter(c => hMetricColMap[c] === undefined);
          console.warn(`/api/upload-hourly (legacy) sheet "${sheetName}" metric header missing/ambiguous (${hMetricMappedCount}/25), falling back to positional for: ${missing.join(',')}`);
        }

        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row || !row[3]) { skippedRows++; continue; }
          const empId = String(row[3]).trim();
          if (!empId || !empSet.has(empId)) { skippedRows++; continue; }
          const metrics = readEodMetrics(row, hMetricColMap, useHMetricHeaderMap, 5);
          await client.query(
            `INSERT INTO hourly_performance (emp_id, product_type_id,
              regular_demand, regular_collection, demand_1_30, collection_1_30,
              demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
              npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
              on_date_demand, on_date_collection,
              regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
              demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
              on_date_demand_amt, on_date_collection_amt)
            VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)
            ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
              regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
              demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
              demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
              pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
              npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
              npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
              on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
              regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
              demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
              demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
              pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
              on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`,
            [empId, ptId, ...metrics]
          );
          insertedCount++;
        }
      }

      await client.query('GRANT ALL ON ALL TABLES IN SCHEMA public TO "Raghunandan1157"');
      await client.query('GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO "Raghunandan1157"');
      await client.query("COMMIT");

      try {
        if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
        fs.writeFileSync(path.join(UPLOAD_DIR, "hourly_report.xlsx"), req.file.buffer);
      } catch (fsErr) { console.error("Hourly file save warning:", fsErr.message); }

      res.json({ success: true, employees: empSet.size, performance: insertedCount });
      return;
    }

    // ---- NEW: Quick Report format parser ----
    console.log("upload-hourly: Detected Quick Report format, sheets:", JSON.stringify(sheetNames));

    // 1. Parse OverAll sheet - Branch+Officer section
    const wsOverAll = wb.Sheets["OverAll"];
    if (!wsOverAll) {
      return res.status(400).json({ error: "OverAll sheet not found in Quick Report" });
    }
    const overAllRows = XLSX.utils.sheet_to_json(wsOverAll, { header: 1 });

    // Find the Branch+Officer section by looking for "EMP ID" in column 0
    let officerStartRow = -1;
    for (let r = 0; r < overAllRows.length; r++) {
      const row = overAllRows[r];
      if (row && row[0] && String(row[0]).trim() === "EMP ID") {
        officerStartRow = r;
        break;
      }
    }

    if (officerStartRow === -1) {
      return res.status(400).json({ error: "Could not find Branch+Officer section (no EMP ID header row)" });
    }
    console.log("upload-hourly: Found EMP ID header at row " + officerStartRow);

    // Data starts after: EMP ID header, sub-header 1, sub-header 2, blank row
    const dataStartRow = officerStartRow + 4;

    // 2. Parse On-Date sheet for on_date_demand and on_date_collection
    const onDateMap = {};
    const wsOnDate = wb.Sheets["OverAll_On-Date"];
    if (wsOnDate) {
      const onDateRows = XLSX.utils.sheet_to_json(wsOnDate, { header: 1 });
      let onDateOfficerStart = -1;
      for (let r = 0; r < onDateRows.length; r++) {
        if (onDateRows[r] && onDateRows[r][0] && String(onDateRows[r][0]).trim() === "EMP ID") {
          onDateOfficerStart = r;
          break;
        }
      }
      if (onDateOfficerStart >= 0) {
        for (let r = onDateOfficerStart + 4; r < onDateRows.length; r++) {
          const row = onDateRows[r];
          if (!row || !row[0]) continue;
          const empId = String(row[0]).trim();
          if (!empId || empId === "EMP ID") continue;
          const demand = (row[2] != null && row[2] !== "-") ? Number(row[2]) : 0;
          const collection = (row[3] != null && row[3] !== "-") ? Number(row[3]) : 0;
          onDateMap[empId] = {
            demand: Number.isFinite(demand) ? demand : 0,
            collection: Number.isFinite(collection) ? collection : 0
          };
        }
      }
      console.log("upload-hourly: Parsed on-date data for " + Object.keys(onDateMap).length + " employees");
    }

    // 3. Extract employee metrics from Branch+Officer section
    // Quick Report column layout (0-indexed) — verified against the actual
    // OverAll sheet header (Quick_Report_Latest.xlsx). The report DOES include
    // a "1-90 DPD" group between PNPA and NPA, so NPA sits at cols 22-24, NOT
    // 18-20. Reading 18-20 corrupts npa_cases with "1-90 DPD Demand", etc.
    //   0=EMP ID  1=Name  2=Reg Demand  3=Reg Collection  4=FTOD  5=Coll%
    //   6=130 Demand  7=130 Collection  8=130 Balance  9=130 Coll%
    //   10=3160 Demand  11=3160 Collection  12=3160 Balance  13=3160 Coll%
    //   14=PNPA Demand  15=PNPA Collection  16=PNPA Balance  17=PNPA Coll%
    //   18=1-90 Demand  19=1-90 Collection  20=1-90 Balance  21=1-90 Coll%
    //   22=NPA Cases  23=NPA 90+ Acc  24=NPA 90+ Amt
    const employeeMetrics = {};

    const safeNum = (v) => {
      if (v == null || v === "" || v === "-") return 0;
      const n = Number(v);
      return Number.isFinite(n) ? n : 0;
    };

    for (let r = dataStartRow; r < overAllRows.length; r++) {
      const row = overAllRows[r];
      if (!row) continue;

      const col0 = row[0] != null ? String(row[0]).trim() : "";

      // Skip branch rows (col0 empty) and Grand Total
      if (!col0 || col0 === "Grand Total") continue;
      // Skip any nested header rows
      if (col0 === "EMP ID") continue;

      const empId = col0;
      const onDate = onDateMap[empId] || { demand: 0, collection: 0 };

      employeeMetrics[empId] = {
        regular_demand: safeNum(row[2]),
        regular_collection: safeNum(row[3]),
        demand_1_30: safeNum(row[6]),
        collection_1_30: safeNum(row[7]),
        demand_31_60: safeNum(row[10]),
        collection_31_60: safeNum(row[11]),
        pnpa_demand: safeNum(row[14]),
        pnpa_collection: safeNum(row[15]),
        npa_cases: safeNum(row[22]),
        npa_act_acc: safeNum(row[23]),
        npa_act_amt: safeNum(row[24]),
        npa_clo_acc: 0,
        npa_clo_amt: 0,
        on_date_demand: onDate.demand,
        on_date_collection: onDate.collection,
        regular_demand_amt: 0, regular_collection_amt: 0,
        demand_1_30_amt: 0, collection_1_30_amt: 0,
        demand_31_60_amt: 0, collection_31_60_amt: 0,
        pnpa_demand_amt: 0, pnpa_collection_amt: 0,
        on_date_demand_amt: 0, on_date_collection_amt: 0
      };
    }

    const empIds = Object.keys(employeeMetrics);
    console.log("upload-hourly: Parsed " + empIds.length + " employees from Quick Report");

    if (empIds.length === 0) {
      return res.status(400).json({ error: "No employee data found in Quick Report" });
    }

    // 4. Insert into hourly_performance
    await client.query("BEGIN");

    await client.query("DROP TABLE IF EXISTS hourly_performance CASCADE");
    await client.query(`CREATE TABLE hourly_performance (
      performance_id SERIAL PRIMARY KEY,
      emp_id VARCHAR(10) NOT NULL,
      product_type_id INT NOT NULL,
      regular_demand INT DEFAULT 0, regular_collection INT DEFAULT 0,
      demand_1_30 INT DEFAULT 0, collection_1_30 INT DEFAULT 0,
      demand_31_60 INT DEFAULT 0, collection_31_60 INT DEFAULT 0,
      pnpa_demand INT DEFAULT 0, pnpa_collection INT DEFAULT 0,
      npa_cases INT DEFAULT 0, npa_act_acc INT DEFAULT 0, npa_act_amt DECIMAL(15,2) DEFAULT 0,
      npa_clo_acc INT DEFAULT 0, npa_clo_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand INT DEFAULT 0, on_date_collection INT DEFAULT 0,
      regular_demand_amt DECIMAL(15,2) DEFAULT 0, regular_collection_amt DECIMAL(15,2) DEFAULT 0,
      demand_1_30_amt DECIMAL(15,2) DEFAULT 0, collection_1_30_amt DECIMAL(15,2) DEFAULT 0,
      demand_31_60_amt DECIMAL(15,2) DEFAULT 0, collection_31_60_amt DECIMAL(15,2) DEFAULT 0,
      pnpa_demand_amt DECIMAL(15,2) DEFAULT 0, pnpa_collection_amt DECIMAL(15,2) DEFAULT 0,
      on_date_demand_amt DECIMAL(15,2) DEFAULT 0, on_date_collection_amt DECIMAL(15,2) DEFAULT 0,
      UNIQUE(emp_id, product_type_id)
    )`);
    await client.query("CREATE INDEX idx_hp_emp_id ON hourly_performance(emp_id)");

    // Ensure "ALL" product type exists for combined data
    let allPtId;
    const ptCheck = await client.query("SELECT product_type_id FROM product_types WHERE product_type_name = 'ALL'");
    if (ptCheck.rows.length > 0) {
      allPtId = ptCheck.rows[0].product_type_id;
    } else {
      const ptInsert = await client.query("INSERT INTO product_types (product_type_name) VALUES ('ALL') RETURNING product_type_id");
      allPtId = ptInsert.rows[0].product_type_id;
    }

    // Build employee lookup from existing EOD data
    const empRows = await client.query("SELECT emp_id FROM employees");
    const empSet = new Set(empRows.rows.map(r => r.emp_id));

    let insertedCount = 0;
    let skippedCount = 0;

    for (const empId of empIds) {
      if (!empSet.has(empId)) {
        skippedCount++;
        continue;
      }

      const m = employeeMetrics[empId];
      await client.query(
        `INSERT INTO hourly_performance (emp_id, product_type_id,
          regular_demand, regular_collection, demand_1_30, collection_1_30,
          demand_31_60, collection_31_60, pnpa_demand, pnpa_collection,
          npa_cases, npa_act_acc, npa_act_amt, npa_clo_acc, npa_clo_amt,
          on_date_demand, on_date_collection,
          regular_demand_amt, regular_collection_amt, demand_1_30_amt, collection_1_30_amt,
          demand_31_60_amt, collection_31_60_amt, pnpa_demand_amt, pnpa_collection_amt,
          on_date_demand_amt, on_date_collection_amt)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,$25,$26,$27)
        ON CONFLICT (emp_id, product_type_id) DO UPDATE SET
          regular_demand=EXCLUDED.regular_demand, regular_collection=EXCLUDED.regular_collection,
          demand_1_30=EXCLUDED.demand_1_30, collection_1_30=EXCLUDED.collection_1_30,
          demand_31_60=EXCLUDED.demand_31_60, collection_31_60=EXCLUDED.collection_31_60,
          pnpa_demand=EXCLUDED.pnpa_demand, pnpa_collection=EXCLUDED.pnpa_collection,
          npa_cases=EXCLUDED.npa_cases, npa_act_acc=EXCLUDED.npa_act_acc, npa_act_amt=EXCLUDED.npa_act_amt,
          npa_clo_acc=EXCLUDED.npa_clo_acc, npa_clo_amt=EXCLUDED.npa_clo_amt,
          on_date_demand=EXCLUDED.on_date_demand, on_date_collection=EXCLUDED.on_date_collection,
          regular_demand_amt=EXCLUDED.regular_demand_amt, regular_collection_amt=EXCLUDED.regular_collection_amt,
          demand_1_30_amt=EXCLUDED.demand_1_30_amt, collection_1_30_amt=EXCLUDED.collection_1_30_amt,
          demand_31_60_amt=EXCLUDED.demand_31_60_amt, collection_31_60_amt=EXCLUDED.collection_31_60_amt,
          pnpa_demand_amt=EXCLUDED.pnpa_demand_amt, pnpa_collection_amt=EXCLUDED.pnpa_collection_amt,
          on_date_demand_amt=EXCLUDED.on_date_demand_amt, on_date_collection_amt=EXCLUDED.on_date_collection_amt`,
        [empId, allPtId,
          m.regular_demand, m.regular_collection, m.demand_1_30, m.collection_1_30,
          m.demand_31_60, m.collection_31_60, m.pnpa_demand, m.pnpa_collection,
          m.npa_cases, m.npa_act_acc, m.npa_act_amt, m.npa_clo_acc, m.npa_clo_amt,
          m.on_date_demand, m.on_date_collection,
          m.regular_demand_amt, m.regular_collection_amt, m.demand_1_30_amt, m.collection_1_30_amt,
          m.demand_31_60_amt, m.collection_31_60_amt, m.pnpa_demand_amt, m.pnpa_collection_amt,
          m.on_date_demand_amt, m.on_date_collection_amt]
      );
      insertedCount++;
    }

    await client.query('GRANT ALL ON ALL TABLES IN SCHEMA public TO "Raghunandan1157"');
    await client.query('GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO "Raghunandan1157"');
    await client.query("COMMIT");

    // Save the uploaded file
    try {
      if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
      fs.writeFileSync(path.join(UPLOAD_DIR, "hourly_report.xlsx"), req.file.buffer);
    } catch (fsErr) { console.error("Hourly file save warning:", fsErr.message); }

    console.log("upload-hourly: Quick Report imported - " + insertedCount + " employees, " + skippedCount + " skipped (not in EOD)");
    res.json({ success: true, employees: insertedCount, performance: insertedCount, skipped: skippedCount, format: "quick_report" });

  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Hourly upload error:", err);
    res.status(500).json({ error: "Hourly upload failed: " + err.message });
  } finally {
    _hourlyUploadInProgress = false;
    client.release();
  }
});

// ========== STEP 3: Hourly API Routes ==========
function buildHourlyQuery(groupBy, groupCol) {
  return `
    SELECT ${groupCol},
      SUM(ep.regular_demand)::int AS regular_demand, SUM(ep.regular_collection)::int AS regular_collection,
      SUM(ep.demand_1_30)::int AS demand_1_30, SUM(ep.collection_1_30)::int AS collection_1_30,
      SUM(ep.demand_31_60)::int AS demand_31_60, SUM(ep.collection_31_60)::int AS collection_31_60,
      SUM(ep.pnpa_demand)::int AS pnpa_demand, SUM(ep.pnpa_collection)::int AS pnpa_collection,
      SUM(ep.npa_cases)::int AS npa_cases, SUM(ep.npa_act_acc)::int AS npa_act_acc, SUM(ep.npa_act_amt) AS npa_act_amt,
      SUM(ep.npa_clo_acc)::int AS npa_clo_acc, SUM(ep.npa_clo_amt) AS npa_clo_amt,
      SUM(ep.on_date_demand)::int AS on_date_demand, SUM(ep.on_date_collection)::int AS on_date_collection,
      SUM(ep.regular_demand_amt) AS regular_demand_amt, SUM(ep.regular_collection_amt) AS regular_collection_amt,
      SUM(ep.demand_1_30_amt) AS demand_1_30_amt, SUM(ep.collection_1_30_amt) AS collection_1_30_amt,
      SUM(ep.demand_31_60_amt) AS demand_31_60_amt, SUM(ep.collection_31_60_amt) AS collection_31_60_amt,
      SUM(ep.pnpa_demand_amt) AS pnpa_demand_amt, SUM(ep.pnpa_collection_amt) AS pnpa_collection_amt,
      SUM(ep.on_date_demand_amt) AS on_date_demand_amt, SUM(ep.on_date_collection_amt) AS on_date_collection_amt
    FROM hourly_performance ep
    JOIN product_types pt ON ep.product_type_id = pt.product_type_id
    JOIN employees e ON ep.emp_id = e.emp_id
    JOIN branches b ON e.branch_id = b.branch_id
    JOIN districts d ON b.district_id = d.district_id
    JOIN regions r ON d.region_id = r.region_id`;
}

function buildHourlyWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push("pt.product_type_name = $" + idx++); params.push(filters.product_type);
  }
  // Hierarchy filters — resolve via employee_master
  if (filters.region || filters.state) {
    where.push("UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($" + (idx++) + "))");
    params.push(filters.region || filters.state);
  }
  if (filters.division) {
    where.push("UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($" + (idx++) + "))");
    params.push(filters.division);
  }
  if (filters.district || filters.area) {
    where.push("UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($" + (idx++) + "))");
    params.push(filters.district || filters.area);
  }
  if (filters.branch) {
    where.push("UPPER(b.branch_name) = UPPER($" + (idx++) + ")");
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push("ep.emp_id = $" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params: params };
}

app.get("/api/hourly/summary", async (req, res) => {
  try {
    const base = buildHourlyQuery(null, "");
    const { clause, params } = buildHourlyWhere(req.query);
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/hourly/by-region", async (req, res) => {
  try {
    const base = buildHourlyQuery("r.region_name", "r.region_name");
    const { clause, params } = buildHourlyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/hourly/by-district", async (req, res) => {
  try {
    const base = buildHourlyQuery("d.district_name, r.region_name", "d.district_name, r.region_name");
    const { clause, params } = buildHourlyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/hourly/by-branch", async (req, res) => {
  try {
    const base = buildHourlyQuery("b.branch_name, d.district_name", "b.branch_name, d.district_name");
    const { clause, params } = buildHourlyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/hourly/by-employee", async (req, res) => {
  try {
    const base = buildHourlyQuery("e.emp_id, e.officer_name, b.branch_name", "e.emp_id, e.officer_name AS name, b.branch_name");
    const { clause, params } = buildHourlyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ========== DAILY PLAN / DAILY REPORTS API ==========

// Branch hierarchy from employee_master (V2 structure)
// Dedupes by branch_name — some branches appear under multiple rows in
// employee_master with slight hierarchy variations; keep the first one.
app.get("/api/daily-plan/branches", async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT branch_name, area_name, division_name, region_name
      FROM (
        SELECT DISTINCT ON (branch_name)
          branch_name, area_name, division_name, region_name
        FROM employee_master
        WHERE branch_name IS NOT NULL AND branch_name != ''
          AND status = 'Working'
          AND region_name NOT IN ('Corporate Office', 'Head Office')
        ORDER BY branch_name, region_name, division_name, area_name
      ) s
      ORDER BY region_name, division_name, area_name, branch_name
    `);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

const DAILY_PLAN_COLS = 'branch_name,date,region,district,dm_name,ftod_actual,ftod_plan,dpd_1_30_actual,dpd_1_30_plan,dpd_31_60_actual,dpd_31_60_plan,dpd_61_90_actual,dpd_61_90_plan,npa_activation,npa_closure,fy_non_start_acc,fy_non_start_plan,disb_igl_acc,disb_igl_amt,disb_fig_acc,disb_fig_amt,disb_il_acc,disb_il_amt,kyc_igl,kyc_fig,kyc_il';

// GET /api/daily-plan/reports?from=DATE&to=DATE&branch=NAME&region=NAME
app.get("/api/daily-plan/reports", async (req, res) => {
  try {
    const where = []; const params = []; let idx = 1;
    if (req.query.from) { where.push("date >= $" + idx++); params.push(req.query.from); }
    if (req.query.to) { where.push("date <= $" + idx++); params.push(req.query.to); }
    if (req.query.branch) { where.push("UPPER(branch_name) = UPPER($" + idx++ + ")"); params.push(req.query.branch); }
    if (req.query.region) { where.push("region = $" + idx++); params.push(req.query.region); }
    if (req.query.district) { where.push("district = $" + idx++); params.push(req.query.district); }
    if (req.query.dm_name) { where.push("dm_name = $" + idx++); params.push(req.query.dm_name); }
    const sql = "SELECT id," + DAILY_PLAN_COLS + ",created_at FROM daily_reports" + (where.length ? " WHERE " + where.join(" AND ") : "") + " ORDER BY date, branch_name";
    const result = await pool.query(sql, params);
    res.json({ data: result.rows, count: result.rowCount });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/daily-plan/achievements?from=DATE&to=DATE&...
app.get("/api/daily-plan/achievements", async (req, res) => {
  try {
    const where = []; const params = []; let idx = 1;
    if (req.query.from) { where.push("date >= $" + idx++); params.push(req.query.from); }
    if (req.query.to) { where.push("date <= $" + idx++); params.push(req.query.to); }
    if (req.query.branch) { where.push("UPPER(branch_name) = UPPER($" + idx++ + ")"); params.push(req.query.branch); }
    if (req.query.region) { where.push("region = $" + idx++); params.push(req.query.region); }
    if (req.query.district) { where.push("district = $" + idx++); params.push(req.query.district); }
    if (req.query.dm_name) { where.push("dm_name = $" + idx++); params.push(req.query.dm_name); }
    const sql = "SELECT id," + DAILY_PLAN_COLS + ",created_at FROM daily_reports_achievements" + (where.length ? " WHERE " + where.join(" AND ") : "") + " ORDER BY date, branch_name";
    const result = await pool.query(sql, params);
    res.json({ data: result.rows, count: result.rowCount });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/daily-plan/check-chain?branch=NAME&date=ISO
// Chain-of-consistency rule: if previous-day plan exists but its achievement is missing,
// block entry for the requested date. Returns {blocked, missingDate}.
app.get("/api/daily-plan/check-chain", async (req, res) => {
  try {
    const { branch, date } = req.query;
    if (!branch || !date) return res.json({ blocked: false });
    const prev = await pool.query(`
      SELECT date FROM daily_reports
      WHERE UPPER(branch_name) = UPPER($1) AND date < $2::date
      ORDER BY date DESC LIMIT 1`, [branch, date]);
    if (!prev.rows.length) return res.json({ blocked: false });
    const prevDate = prev.rows[0].date;
    const ach = await pool.query(`
      SELECT 1 FROM daily_reports_achievements
      WHERE UPPER(branch_name) = UPPER($1) AND date = $2::date LIMIT 1`, [branch, prevDate]);
    if (ach.rows.length > 0) return res.json({ blocked: false });
    res.json({ blocked: true, missingDate: prevDate });
  } catch(e) { res.status(500).json({ error: e.message, blocked: false }); }
});

// GET /api/daily-plan/monthly-actuals?branch=NAME&date=ISO
// Returns the earliest row in the current month (excluding today) where any DPD actual is set.
// Used to lock DPD actual fields once set per month.
app.get("/api/daily-plan/monthly-actuals", async (req, res) => {
  try {
    const { branch, date } = req.query;
    if (!branch || !date) return res.json({ data: null });
    const sql = `
      SELECT dpd_1_30_actual, dpd_31_60_actual, dpd_61_90_actual, date
      FROM daily_reports
      WHERE UPPER(branch_name) = UPPER($1)
        AND date >= date_trunc('month', $2::date)
        AND date <  date_trunc('month', $2::date) + interval '1 month'
        AND (
          (dpd_1_30_actual IS NOT NULL AND dpd_1_30_actual > 0)
          OR (dpd_31_60_actual IS NOT NULL AND dpd_31_60_actual > 0)
          OR (dpd_61_90_actual IS NOT NULL AND dpd_61_90_actual > 0)
        )
      ORDER BY date ASC LIMIT 1`;
    const result = await pool.query(sql, [branch, date]);
    res.json({ data: result.rows[0] || null });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/daily-plan/exists?date=DATE&branch=NAME&table=reports|achievements
app.get("/api/daily-plan/exists", async (req, res) => {
  try {
    const table = req.query.table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    const result = await pool.query("SELECT COUNT(*)::int AS cnt FROM " + table + " WHERE date=$1 AND UPPER(branch_name)=UPPER($2)", [req.query.date, req.query.branch]);
    res.json({ exists: result.rows[0].cnt > 0 });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/daily-plan/pending-branches?table=reports|achievements&date=ISO&role=CEO|RM|DM|AM&location=NAME
// Returns branches (scoped by role/location) that have NO row in the chosen table for that date.
// Each row includes BM contact from employee_master (role='BM'); null if no BM exists.
app.get("/api/daily-plan/pending-branches", async (req, res) => {
  try {
    const { table, date, role, location } = req.query;
    if (!table || !date || !role) return res.status(400).json({ error: "Missing table, date, or role" });
    if (role !== 'CEO' && !location) return res.status(400).json({ error: "location required for non-CEO roles" });
    const targetTable = table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    const sql = `
      WITH branch_list AS (
        SELECT DISTINCT branch_name, area_name, division_name, region_name
        FROM employee_master
        WHERE status = 'Working'
          AND branch_name IS NOT NULL AND branch_name <> ''
          AND (
            $2 = 'CEO'
            OR ($2 IN ('RM','SM')  AND UPPER(TRIM(region_name))   = UPPER(TRIM($3)))
            OR ($2 IN ('DM','DvM') AND UPPER(TRIM(division_name)) = UPPER(TRIM($3)))
            OR ($2 = 'AM'          AND UPPER(TRIM(area_name))     = UPPER(TRIM($3)))
          )
      ),
      bm AS (
        SELECT DISTINCT ON (UPPER(branch_name)) UPPER(branch_name) AS branch_key,
               full_name, mobile, emp_id
        FROM employee_master
        WHERE role = 'BM' AND status = 'Working'
          AND branch_name IS NOT NULL AND branch_name <> ''
        ORDER BY UPPER(branch_name), emp_id
      )
      SELECT bl.branch_name AS branch,
             bl.area_name   AS area,
             bl.division_name AS division,
             bl.region_name AS region,
             bm.full_name   AS bm_name,
             bm.mobile      AS bm_phone
      FROM branch_list bl
      LEFT JOIN bm ON bm.branch_key = UPPER(bl.branch_name)
      WHERE NOT EXISTS (
        SELECT 1 FROM ${targetTable} dr
        WHERE UPPER(dr.branch_name) = UPPER(bl.branch_name)
          AND dr.date = $1::date
      )
      ORDER BY bl.branch_name ASC`;
    const result = await pool.query(sql, [date, role, location || '']);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Helper: enforce AM tenancy on a daily-plan write.
// Returns { ok: true } or { ok: false, status, error } so callers can short-circuit.
// Opt-in: only runs when X-User-Role === 'AM'. Other roles (or missing header)
// fall through unchecked — preserves backwards compatibility with BM/CEO/etc. callers.
async function _checkDailyPlanTenancy(req, branchName) {
  const role = (req.headers['x-user-role'] || '').trim();
  if (role !== 'AM') return { ok: true };
  const loc = (req.headers['x-user-location'] || '').trim();
  if (!loc) {
    return { ok: false, status: 400, error: "X-User-Location header required for AM role" };
  }
  try {
    const areaLookup = await pool.query(
      `SELECT area_name FROM employee_master
       WHERE UPPER(TRIM(branch_name)) = UPPER(TRIM($1))
         AND status = 'Working'
         AND area_name IS NOT NULL AND area_name <> ''
       LIMIT 1`,
      [branchName]
    );
    if (!areaLookup.rows.length) {
      return { ok: false, status: 403, error: "Branch not in your area" };
    }
    const branchArea = (areaLookup.rows[0].area_name || '').trim().toUpperCase();
    const userLoc = loc.toUpperCase();
    if (branchArea !== userLoc) {
      return { ok: false, status: 403, error: "Branch not in your area" };
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, status: 500, error: "Tenancy check failed: " + e.message };
  }
}

// POST /api/daily-plan/save — Save daily report (plan or achievement)
// Body: { table: 'plans'|'achievements', data: {...} }
//   Also accepts { table, row: {...} } as an alias for `data` (new AM-edit UI).
// Optional headers: X-User-Role, X-User-Location — if role === 'AM', server
//   verifies the branch belongs to the caller's area before writing.
app.post("/api/daily-plan/save", async (req, res) => {
  try {
    const { table } = req.body;
    const data = req.body.data || req.body.row;
    if (!data || !data.date || !data.branch_name) {
      return res.status(400).json({ error: "Missing date or branch_name" });
    }
    const targetTable = table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';

    // Tenancy guard: AM callers may only edit branches in their own area.
    const gate = await _checkDailyPlanTenancy(req, data.branch_name);
    if (!gate.ok) return res.status(gate.status).json({ error: gate.error });

    // Auto-populate region/district from branches hierarchy if not provided
    if (!data.region || !data.district) {
      try {
        const lookup = await pool.query(
          `SELECT r.region_name, d.district_name FROM branches b
           JOIN districts d ON b.district_id = d.district_id
           JOIN regions r ON d.region_id = r.region_id
           WHERE UPPER(b.branch_name) = UPPER($1) LIMIT 1`, [data.branch_name]);
        if (lookup.rows.length) {
          if (!data.region) data.region = lookup.rows[0].region_name;
          if (!data.district) data.district = lookup.rows[0].district_name;
        }
      } catch(e) { /* lookup optional */ }
    }

    const cols = DAILY_PLAN_COLS.split(',');
    const values = cols.map(c => data[c] != null ? data[c] : 0);
    const placeholders = cols.map((_, i) => "$" + (i + 1));

    // Exclude region/district from ON CONFLICT UPDATE so existing values aren't overwritten with empty
    const updateCols = cols.filter(c => c !== 'branch_name' && c !== 'date' && c !== 'region' && c !== 'district');
    const sql = "INSERT INTO " + targetTable + " (" + cols.join(",") + ") VALUES (" + placeholders.join(",") + ") ON CONFLICT (branch_name, date) DO UPDATE SET " + updateCols.map(c => c + "=EXCLUDED." + c).join(",") + ", region = CASE WHEN EXCLUDED.region = '' OR EXCLUDED.region = '0' THEN " + targetTable + ".region ELSE EXCLUDED.region END, district = CASE WHEN EXCLUDED.district = '' OR EXCLUDED.district = '0' THEN " + targetTable + ".district ELSE EXCLUDED.district END";

    await pool.query(sql, values);
    res.json({ success: true, table: targetTable, branch: data.branch_name, date: data.date });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/daily-plan/bulk-save — Save multiple branches at once
// Optional headers: X-User-Role, X-User-Location — if role === 'AM', every
//   branch in the batch is verified against the caller's area; any mismatch
//   aborts the whole transaction.
app.post("/api/daily-plan/bulk-save", async (req, res) => {
  const client = await pool.connect();
  try {
    const { table, rows } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: "No rows" });
    const targetTable = table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    const cols = DAILY_PLAN_COLS.split(',');

    // Tenancy guard: AM callers may only edit branches in their own area.
    // Check every branch up-front so we fail fast before starting the txn.
    const role = (req.headers['x-user-role'] || '').trim();
    if (role === 'AM') {
      for (const data of rows) {
        if (!data || !data.branch_name) continue;
        const gate = await _checkDailyPlanTenancy(req, data.branch_name);
        if (!gate.ok) return res.status(gate.status).json({ error: gate.error + " (" + data.branch_name + ")" });
      }
    }

    await client.query("BEGIN");
    let saved = 0;
    for (const data of rows) {
      if (!data.date || !data.branch_name) continue;
      const values = cols.map(c => data[c] != null ? data[c] : 0);
      const placeholders = cols.map((_, i) => "$" + (i + 1));
      const sql = "INSERT INTO " + targetTable + " (" + cols.join(",") + ") VALUES (" + placeholders.join(",") + ") ON CONFLICT (branch_name, date) DO UPDATE SET " + cols.filter(c => c !== 'branch_name' && c !== 'date').map(c => c + "=EXCLUDED." + c).join(",");
      await client.query(sql, values);
      saved++;
    }
    await client.query("COMMIT");
    res.json({ success: true, saved });
  } catch(e) {
    await client.query("ROLLBACK").catch(() => {});
    res.status(500).json({ error: e.message });
  } finally { client.release(); }
});

// ==================== AI CHAT ENDPOINT ====================
// OpenRouter API key + models are now used client-side (see employee.html)

// Helper: run a read-only query safely (with timeout + row limit)
// AsyncLocalStorage so voice/stream endpoints can capture every SQL run by
// AI tool calls without threading a logger through every callsite.
const { AsyncLocalStorage } = require('async_hooks');
const aiToolLogStore = new AsyncLocalStorage();

async function safeQuery(sql, params, maxRows) {
  const client = await pool.connect();
  const log = aiToolLogStore.getStore();
  const t0 = Date.now();
  try {
    await client.query("SET statement_timeout = '5000'"); // 5s max
    const result = await client.query(sql, params);
    const rows = result.rows.slice(0, maxRows || 50);
    if (log) log.push({ sql, params: params || [], rowCount: rows.length, ms: Date.now() - t0 });
    return rows;
  } catch (e) {
    if (log) log.push({ sql, params: params || [], error: e.message, ms: Date.now() - t0 });
    return { error: e.message };
  } finally {
    client.release();
  }
}

// Build context snapshot for the AI
async function buildDataContext(session) {
  const ctx = {};
  const role = (session.role || '').toUpperCase();
  const loc = (session.location || '').trim();
  const isCeo = !role || role === 'CEO' || !loc;

  // -- Time anchors for the AI's "now" awareness ---------------------------
  const now = new Date();
  const yyyyMmDd = (d) => d.toISOString().slice(0, 10);
  const fyStartYear = now.getMonth() + 1 >= 4 ? now.getFullYear() : now.getFullYear() - 1;
  ctx.now = yyyyMmDd(now);
  ctx.fyStart = fyStartYear + '-04-01';

  // Latest dates available — coerce Date object to ISO YYYY-MM-DD.
  const dates = await safeQuery("SELECT MAX(report_date) as latest FROM daily_performance", [], 1);
  const _latestRaw = (Array.isArray(dates) && dates[0]?.latest) || null;
  ctx.latestDate = _latestRaw
    ? (_latestRaw instanceof Date
        ? _latestRaw.toISOString().slice(0, 10)
        : String(_latestRaw).slice(0, 10))
    : 'unknown';

  // Build a branch-name IN clause via employee_master — single source of truth
  let branchFilter = '';
  let params = [];
  let emCol = null;
  if (!isCeo) {
    if (role === 'RM' || role === 'SM') emCol = 'region_name';
    else if (role === 'DM' || role === 'DVM') emCol = 'division_name';
    else if (role === 'AM') emCol = 'area_name';
    else if (role === 'BM' || role === 'FO') emCol = 'branch_name';
    if (emCol) {
      branchFilter = `WHERE UPPER(b.branch_name) IN (
        SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(${emCol}) ILIKE TRIM($1)
      )`;
      params = [loc];
    }
  }

  // Scoped overall summary (rolling totals from employee_performance)
  const overall = await safeQuery(`
    SELECT SUM(ep.regular_demand) as total_rd, SUM(ep.regular_collection) as total_rc,
           SUM(ep.demand_1_30) as sma0_d, SUM(ep.collection_1_30) as sma0_c,
           SUM(ep.demand_31_60) as sma1_d, SUM(ep.collection_31_60) as sma1_c,
           SUM(ep.pnpa_demand) as pnpa_d, SUM(ep.pnpa_collection) as pnpa_c,
           SUM(ep.npa_cases) as npa_cases, SUM(ep.npa_act_acc) as npa_act
    FROM employee_performance ep
    JOIN employees e ON ep.emp_id = e.emp_id
    JOIN branches b ON e.branch_id = b.branch_id
    ${branchFilter}`, params, 1);
  ctx.summary = (Array.isArray(overall) && overall[0]) || {};
  ctx.scope = isCeo ? 'all' : `${role} — ${loc}`;

  // Scoped branch-level breakdown (top 20 by demand)
  ctx.branches = await safeQuery(`
    SELECT b.branch_name, SUM(ep.regular_demand) as rd, SUM(ep.regular_collection) as rc,
           SUM(ep.npa_cases) as npa_cases
    FROM employee_performance ep
    JOIN employees e ON ep.emp_id = e.emp_id
    JOIN branches b ON e.branch_id = b.branch_id
    ${branchFilter}
    GROUP BY b.branch_name ORDER BY rd DESC LIMIT 20`, params, 20);
  if (!Array.isArray(ctx.branches)) ctx.branches = [];

  // ---------- Monthly history (last 12 months, daily_performance) ---------
  // Used for "last month vs this month", FY-to-date, and trend questions.
  ctx.monthly = await safeQuery(`
    SELECT to_char(date_trunc('month', dp.report_date), 'YYYY-MM') AS month,
           SUM(dp.regular_demand)::bigint AS total_demand,
           SUM(dp.regular_collection)::bigint AS total_collection,
           SUM(dp.npa_cases)::int AS npa_count,
           SUM(dp.npa_act_amt)::numeric AS npa_amount
      FROM daily_performance dp
      JOIN employees e ON dp.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
      ${branchFilter ? branchFilter + ' AND' : 'WHERE'} dp.report_date >= (CURRENT_DATE - INTERVAL '12 months')
     GROUP BY date_trunc('month', dp.report_date)
     ORDER BY month DESC`, params, 12);
  if (!Array.isArray(ctx.monthly)) ctx.monthly = [];
  // Add collection_pct so the model doesn't need to divide.
  ctx.monthly = ctx.monthly.map((r) => {
    const d = Number(r.total_demand) || 0;
    const c = Number(r.total_collection) || 0;
    return { ...r, collection_pct: d > 0 ? Number(((c / d) * 100).toFixed(2)) : 0 };
  });

  // ---------- Daily last-30 (for short-term trend) ------------------------
  ctx.dailyLast30 = await safeQuery(`
    SELECT to_char(dp.report_date, 'YYYY-MM-DD') AS date,
           SUM(dp.regular_demand)::bigint AS demand,
           SUM(dp.regular_collection)::bigint AS collection
      FROM daily_performance dp
      JOIN employees e ON dp.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
      ${branchFilter ? branchFilter + ' AND' : 'WHERE'} dp.report_date >= (CURRENT_DATE - INTERVAL '30 days')
     GROUP BY dp.report_date
     ORDER BY dp.report_date DESC`, params, 30);
  if (!Array.isArray(ctx.dailyLast30)) ctx.dailyLast30 = [];
  ctx.dailyLast30 = ctx.dailyLast30.map((r) => {
    const d = Number(r.demand) || 0;
    const c = Number(r.collection) || 0;
    return { ...r, collection_pct: d > 0 ? Number(((c / d) * 100).toFixed(2)) : 0 };
  });

  // ---------- NPA trajectory (last 6 months) ------------------------------
  ctx.npaTrend = await safeQuery(`
    SELECT to_char(date_trunc('month', dp.report_date), 'YYYY-MM') AS month,
           SUM(dp.npa_cases)::int AS npa_cases,
           SUM(dp.npa_act_acc)::int AS npa_act_acc,
           SUM(dp.npa_act_amt)::numeric AS npa_act_amt
      FROM daily_performance dp
      JOIN employees e ON dp.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
      ${branchFilter ? branchFilter + ' AND' : 'WHERE'} dp.report_date >= (CURRENT_DATE - INTERVAL '6 months')
     GROUP BY date_trunc('month', dp.report_date)
     ORDER BY month DESC`, params, 6);
  if (!Array.isArray(ctx.npaTrend)) ctx.npaTrend = [];

  // ---------- Disbursement: latest 2 months (back-compat) + 12-month trend
  let disbFilter = '';
  if (!isCeo && emCol_for_disb(role)) {
    disbFilter = `WHERE UPPER(d.branch_name) IN (
      SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(${emCol_for_disb(role)}) ILIKE TRIM($1)
    )`;
  }
  ctx.disbursement = await safeQuery(`${DISB_CTE}
    SELECT d.db_month, SUM(d.disb_count) as total_count, SUM(d.disb_amount) as total_amount
    FROM d ${disbFilter}
    GROUP BY d.db_month ORDER BY d.db_month DESC LIMIT 2`, disbFilter ? [loc] : [], 2);
  if (!Array.isArray(ctx.disbursement)) ctx.disbursement = [];

  ctx.disbursementMonthly = await safeQuery(`${DISB_CTE}
    SELECT d.db_month AS month, SUM(d.disb_count)::int AS count, SUM(d.disb_amount)::numeric AS amount
    FROM d ${disbFilter}
    GROUP BY d.db_month ORDER BY d.db_month DESC LIMIT 12`, disbFilter ? [loc] : [], 12);
  if (!Array.isArray(ctx.disbursementMonthly)) ctx.disbursementMonthly = [];

  // ---------- Scope members (branch list + counts in user's scope) -------
  if (isCeo) {
    const totalsRes = await safeQuery(
      "SELECT COUNT(DISTINCT branch_name)::int AS branch_count, COUNT(*)::int AS emp_count FROM employee_master WHERE status='Working'",
      [], 1
    );
    ctx.scopeMembers = {
      scope: 'all',
      branchCount: (Array.isArray(totalsRes) && totalsRes[0]?.branch_count) || 0,
      employeeCount: (Array.isArray(totalsRes) && totalsRes[0]?.emp_count) || 0,
      // Don't dump every branch name for CEO — too noisy. The branches array
      // already has the top 20 by demand.
      branches: [],
    };
  } else if (emCol) {
    const memberRes = await safeQuery(
      `SELECT branch_name, COUNT(*)::int AS emp_count
         FROM employee_master
        WHERE status='Working' AND TRIM(${emCol}) ILIKE TRIM($1)
        GROUP BY branch_name
        ORDER BY branch_name`,
      [loc], 100
    );
    const list = Array.isArray(memberRes) ? memberRes : [];
    ctx.scopeMembers = {
      scope: `${role} — ${loc}`,
      branchCount: list.length,
      employeeCount: list.reduce((acc, r) => acc + (Number(r.emp_count) || 0), 0),
      branches: list.slice(0, 50).map((r) => r.branch_name),
    };
  } else {
    ctx.scopeMembers = { scope: 'unknown', branchCount: 0, employeeCount: 0, branches: [] };
  }

  // ---------- Employee directory (scoped, capped at 500) ----------------
  // For "who is X" / "what's the role of Y" questions. Pulls from
  // employee_master, scoped via the same emCol/loc as everything else.
  {
    let where = "status = 'Working'";
    const empParams = [];
    if (!isCeo && emCol) {
      empParams.push(loc);
      where += ` AND TRIM(${emCol}) ILIKE TRIM($1)`;
    }
    const totalRes = await safeQuery(
      `SELECT COUNT(*)::int AS n FROM employee_master WHERE ${where}`,
      empParams, 1
    );
    const total = (Array.isArray(totalRes) && totalRes[0]?.n) || 0;
    const empRes = await safeQuery(
      `SELECT emp_id, full_name AS name, role, branch_name AS branch,
              area_name AS area, division_name AS division,
              region_name AS region, mobile, status
         FROM employee_master
        WHERE ${where}
        ORDER BY full_name
        LIMIT 500`,
      empParams, 500
    );
    ctx.employees = Array.isArray(empRes) ? empRes : [];
    ctx.employeesTotal = total;
    ctx.employeesTruncated = total > ctx.employees.length;
  }

  // ---------- Employee performance leaderboard (top 50 in scope) ---------
  // FY-to-date totals, used for "top performers" / "who's behind" queries.
  ctx.employeePerformance = await safeQuery(`
    SELECT em.emp_id, em.full_name AS name, em.branch_name AS branch,
           SUM(dp.regular_demand)::bigint AS total_demand,
           SUM(dp.regular_collection)::bigint AS total_collection,
           SUM(dp.npa_cases)::int AS npa_count
      FROM daily_performance dp
      JOIN employees e ON dp.emp_id = e.emp_id
      JOIN branches b ON e.branch_id = b.branch_id
      JOIN employee_master em ON em.emp_id = dp.emp_id
      ${branchFilter ? branchFilter + ' AND' : 'WHERE'} dp.report_date >= $${params.length + 1}
     GROUP BY em.emp_id, em.full_name, em.branch_name
     ORDER BY total_demand DESC
     LIMIT 50`, [...params, ctx.fyStart], 50);
  if (!Array.isArray(ctx.employeePerformance)) ctx.employeePerformance = [];
  ctx.employeePerformance = ctx.employeePerformance.map((r) => {
    const d = Number(r.total_demand) || 0;
    const c = Number(r.total_collection) || 0;
    return { ...r, collection_pct: d > 0 ? Number(((c / d) * 100).toFixed(2)) : 0 };
  });

  // ---------- Per-branch detail (capped at 100 in scope) ----------------
  // {branch_name, region, division, area, employee_count, latest_demand,
  //  latest_collection, npa}. Numbers come from employee_performance (current
  //  rolling totals) joined to employee_master for hierarchy + headcount.
  ctx.branchDetail = await safeQuery(`
    WITH em_agg AS (
      SELECT branch_name,
             MAX(region_name) AS region,
             MAX(division_name) AS division,
             MAX(area_name) AS area,
             COUNT(*) FILTER (WHERE status='Working')::int AS employee_count
        FROM employee_master
       GROUP BY branch_name
    ),
    perf AS (
      SELECT b.branch_name,
             SUM(ep.regular_demand)::bigint AS latest_demand,
             SUM(ep.regular_collection)::bigint AS latest_collection,
             SUM(ep.npa_cases)::int AS npa
        FROM employee_performance ep
        JOIN employees e ON ep.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
        ${branchFilter}
       GROUP BY b.branch_name
    )
    SELECT p.branch_name,
           em_agg.region, em_agg.division, em_agg.area,
           COALESCE(em_agg.employee_count, 0) AS employee_count,
           p.latest_demand, p.latest_collection, p.npa
      FROM perf p
      LEFT JOIN em_agg ON UPPER(em_agg.branch_name) = UPPER(p.branch_name)
     ORDER BY p.latest_demand DESC NULLS LAST
     LIMIT 100`, params, 100);
  if (!Array.isArray(ctx.branchDetail)) ctx.branchDetail = [];

  // ---------- Daily plan + achievement (last 14 days, scope-respecting) ----
  // Aggregates `daily_reports` (PLAN) and `daily_reports_achievements` (ACTUAL)
  // per date over the user's scope. Filters on branch_name via the same
  // employee_master/emCol pattern used elsewhere. CEO sees all branches.
  // Result: ctx.dailyPlan (array DESC) + ctx.latest (single most-recent row)
  // + ctx.latestByBranch (top branches for the latest date — for parity with
  // the online tool's "top 5 branches" answer pattern).
  {
    let drBranchFilter = '';
    let drParams = [];
    if (!isCeo && emCol) {
      drBranchFilter = `AND UPPER(branch_name) IN (
        SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(${emCol}) ILIKE TRIM($1)
      )`;
      drParams = [loc];
    }
    const planSql = `
      SELECT date,
             SUM(COALESCE(ftod_plan, 0))::bigint AS ftod_plan,
             SUM(COALESCE(dpd_1_30_plan, 0))::bigint AS dpd_1_30_plan,
             SUM(COALESCE(dpd_31_60_plan, 0))::bigint AS dpd_31_60_plan,
             SUM(COALESCE(dpd_61_90_plan, 0))::bigint AS dpd_61_90_plan,
             SUM(COALESCE(fy_non_start_plan, 0))::bigint AS fy_non_start_plan
        FROM daily_reports
       WHERE date >= (CURRENT_DATE - INTERVAL '14 days')
         ${drBranchFilter}
       GROUP BY date`;
    const actSql = `
      SELECT date,
             SUM(COALESCE(ftod_actual, 0))::bigint AS ftod_actual,
             SUM(COALESCE(dpd_1_30_actual, 0))::bigint AS dpd_1_30_actual,
             SUM(COALESCE(dpd_31_60_actual, 0))::bigint AS dpd_31_60_actual,
             SUM(COALESCE(dpd_61_90_actual, 0))::bigint AS dpd_61_90_actual,
             SUM(COALESCE(fy_non_start_acc, 0))::bigint AS fy_non_start_acc,
             SUM(COALESCE(npa_activation, 0))::bigint AS npa_activation,
             SUM(COALESCE(npa_closure, 0))::bigint AS npa_closure,
             SUM(COALESCE(disb_igl_acc, 0))::bigint AS disb_igl_acc,
             SUM(COALESCE(disb_igl_amt, 0))::numeric AS disb_igl_amt,
             SUM(COALESCE(disb_fig_acc, 0))::bigint AS disb_fig_acc,
             SUM(COALESCE(disb_fig_amt, 0))::numeric AS disb_fig_amt,
             SUM(COALESCE(disb_il_acc, 0))::bigint AS disb_il_acc,
             SUM(COALESCE(disb_il_amt, 0))::numeric AS disb_il_amt,
             SUM(COALESCE(kyc_igl, 0))::bigint AS kyc_igl,
             SUM(COALESCE(kyc_fig, 0))::bigint AS kyc_fig,
             SUM(COALESCE(kyc_il, 0))::bigint AS kyc_il
        FROM daily_reports_achievements
       WHERE date >= (CURRENT_DATE - INTERVAL '14 days')
         ${drBranchFilter}
       GROUP BY date`;
    const planRows = await safeQuery(planSql, drParams, 14);
    const actRows = await safeQuery(actSql, drParams, 14);
    const planByDate = new Map();
    const actByDate = new Map();
    // pg returns DATE columns as JS Date objects — convert to ISO YYYY-MM-DD.
    const _toIsoDate = (d) => {
      if (!d) return '';
      if (d instanceof Date) return d.toISOString().slice(0, 10);
      const s = String(d);
      // Already ISO?
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      const parsed = new Date(s);
      return isNaN(parsed.getTime()) ? s.slice(0, 10) : parsed.toISOString().slice(0, 10);
    };
    if (Array.isArray(planRows)) for (const r of planRows) planByDate.set(_toIsoDate(r.date), r);
    if (Array.isArray(actRows)) for (const r of actRows) actByDate.set(_toIsoDate(r.date), r);
    const allDates = Array.from(new Set([...planByDate.keys(), ...actByDate.keys()]))
      .filter(Boolean)
      .sort()
      .reverse();
    ctx.dailyPlan = allDates.slice(0, 14).map((d) => {
      const p = planByDate.get(d) || {};
      const a = actByDate.get(d) || {};
      return {
        date: d,
        ftod_plan: Number(p.ftod_plan) || 0,
        ftod_actual: Number(a.ftod_actual) || 0,
        dpd_1_30_plan: Number(p.dpd_1_30_plan) || 0,
        dpd_1_30_actual: Number(a.dpd_1_30_actual) || 0,
        dpd_31_60_plan: Number(p.dpd_31_60_plan) || 0,
        dpd_31_60_actual: Number(a.dpd_31_60_actual) || 0,
        dpd_61_90_plan: Number(p.dpd_61_90_plan) || 0,
        dpd_61_90_actual: Number(a.dpd_61_90_actual) || 0,
        fy_non_start_plan: Number(p.fy_non_start_plan) || 0,
        fy_non_start_acc: Number(a.fy_non_start_acc) || 0,
        npa_activation: Number(a.npa_activation) || 0,
        npa_closure: Number(a.npa_closure) || 0,
        disb_igl_acc: Number(a.disb_igl_acc) || 0,
        disb_igl_amt: Number(a.disb_igl_amt) || 0,
        disb_fig_acc: Number(a.disb_fig_acc) || 0,
        disb_fig_amt: Number(a.disb_fig_amt) || 0,
        disb_il_acc: Number(a.disb_il_acc) || 0,
        disb_il_amt: Number(a.disb_il_amt) || 0,
        kyc_igl: Number(a.kyc_igl) || 0,
        kyc_fig: Number(a.kyc_fig) || 0,
        kyc_il: Number(a.kyc_il) || 0,
      };
    });
    ctx.latest = ctx.dailyPlan[0] || null;

    // Per-branch breakdown for the latest date that has achievement data
    // (today's row may exist on the plan side but be empty on actuals).
    // Supports "top 5 branches by FTOD/DPD/etc" answers without a server
    // roundtrip. Capped at 30 rows.
    const latestAchDateSql = `
      SELECT MAX(date)::text AS d
        FROM daily_reports_achievements
       WHERE COALESCE(ftod_actual, 0) > 0
         ${drBranchFilter}`;
    const latestAchDateRes = await safeQuery(latestAchDateSql, drParams, 1);
    const lbbDate =
      Array.isArray(latestAchDateRes) && latestAchDateRes[0]?.d
        ? String(latestAchDateRes[0].d).slice(0, 10)
        : null;
    if (lbbDate) {
      const byBranchSql = `
        SELECT branch_name,
               SUM(COALESCE(ftod_actual, 0))::bigint AS ftod_actual,
               SUM(COALESCE(ftod_plan, 0))::bigint AS ftod_plan,
               SUM(COALESCE(dpd_1_30_actual, 0))::bigint AS dpd_1_30_actual,
               SUM(COALESCE(dpd_1_30_plan, 0))::bigint AS dpd_1_30_plan,
               SUM(COALESCE(dpd_31_60_actual, 0))::bigint AS dpd_31_60_actual,
               SUM(COALESCE(dpd_31_60_plan, 0))::bigint AS dpd_31_60_plan,
               SUM(COALESCE(dpd_61_90_actual, 0))::bigint AS dpd_61_90_actual,
               SUM(COALESCE(dpd_61_90_plan, 0))::bigint AS dpd_61_90_plan,
               SUM(COALESCE(npa_activation, 0))::bigint AS npa_activation,
               SUM(COALESCE(disb_igl_acc, 0)
                 + COALESCE(disb_fig_acc, 0)
                 + COALESCE(disb_il_acc, 0))::bigint AS disb_total_acc
          FROM daily_reports_achievements
         WHERE date = $${drParams.length + 1}
           ${drBranchFilter}
         GROUP BY branch_name
         ORDER BY ftod_actual DESC NULLS LAST
         LIMIT 30`;
      const lbb = await safeQuery(byBranchSql, [...drParams, lbbDate], 30);
      ctx.latestByBranch = Array.isArray(lbb) ? lbb : [];
      ctx.latestByBranchDate = lbbDate;
    } else {
      ctx.latestByBranch = [];
      ctx.latestByBranchDate = null;
    }
  }

  // ---------- Per-array truncation flags ---------------------------------
  // Lets the model know when a list was capped so it can say "showing top N"
  // instead of pretending the full set is here.
  ctx.truncated = {
    branches: Array.isArray(ctx.branches) && ctx.branches.length >= 20,
    branchDetail: Array.isArray(ctx.branchDetail) && ctx.branchDetail.length >= 100,
    employees: !!ctx.employeesTruncated,
    employeePerformance: Array.isArray(ctx.employeePerformance) && ctx.employeePerformance.length >= 50,
    monthly: false, // 12-month cap; rarely truncated.
    dailyLast30: false, // 30-day cap; matches intent.
    npaTrend: false,
    disbursement: false,
    disbursementMonthly: false,
    dailyPlan: false, // 14-day cap; matches intent.
    latestByBranch: Array.isArray(ctx.latestByBranch) && ctx.latestByBranch.length >= 30,
    scopeMembers:
      !!ctx.scopeMembers &&
      Array.isArray(ctx.scopeMembers.branches) &&
      ctx.scopeMembers.branches.length >= 50,
  };

  // ---------- generated_at (model-visible temporal anchor) ---------------
  ctx.generated_at = new Date().toISOString();

  // ---------- Pre-rendered summary_text (mobile prompt drop-in) -----------
  // The mobile on-device prompt has a tight char budget. summary_text is the
  // server's best one-paragraph summary of the snapshot — mobile uses it
  // verbatim instead of trying to JSON-stringify the whole ctx.
  ctx.summary_text = renderSnapshotSummary(ctx);

  return ctx;
}

// Pre-renders the snapshot for the mobile on-device prompt. Format mirrors
// the parity bar set by the online Mistral path (today total + 7-day trend +
// top-N branches per metric) so offline answers match online ones.
// Markdown sections; ~2-3 KB; well under the 6 KB on-device budget.
function renderSnapshotSummary(ctx) {
  const fmt = (n) => {
    const v = Number(n) || 0;
    if (Math.abs(v) >= 1e7) return (v / 1e7).toFixed(2) + ' Cr';
    if (Math.abs(v) >= 1e5) return (v / 1e5).toFixed(2) + ' L';
    return v.toLocaleString('en-IN');
  };
  const fmtAmt = (n) => '₹' + fmt(n);
  const pct = (a, p) => {
    const av = Number(a) || 0;
    const pv = Number(p) || 0;
    if (pv <= 0) return null;
    return Number(((av / pv) * 100).toFixed(1));
  };
  const md = [];
  const scope = ctx.scope || 'all';
  md.push(`DATA AS OF ${ctx.latestDate || ctx.now} (SCOPE: ${scope})`);
  if (ctx.generated_at) md.push(`Snapshot generated: ${ctx.generated_at}`);
  if (ctx.scopeMembers) {
    md.push(
      `Coverage: ${ctx.scopeMembers.branchCount} branch(es), ${ctx.scopeMembers.employeeCount} active employee(s).`
    );
  }

  // ---------- FTOD ----------
  const dp = Array.isArray(ctx.dailyPlan) ? ctx.dailyPlan : [];
  const lbb = Array.isArray(ctx.latestByBranch) ? ctx.latestByBranch : [];
  // The by-branch slice may be from a past date if today's actuals are empty.
  const lbbDate = ctx.latestByBranchDate || (ctx.latest && ctx.latest.date) || '';
  const lbbLabel = lbbDate && ctx.latest && lbbDate !== ctx.latest.date
    ? `latest with actuals (${lbbDate})`
    : 'today';
  if (ctx.latest) {
    const L = ctx.latest;
    md.push('');
    md.push(
      `FTOD: today ${L.date} plan ${fmt(L.ftod_plan)}, actual ${fmt(L.ftod_actual)}` +
        (pct(L.ftod_actual, L.ftod_plan) != null ? ` (${pct(L.ftod_actual, L.ftod_plan)}%)` : '') +
        '.'
    );
    if (dp.length > 1) {
      const trend = dp
        .slice(0, 7)
        .reverse()
        .map((r) => `${(r.date || '').slice(5)}: ${r.ftod_actual}`)
        .join(', ');
      md.push(`FTOD 7-day trend [${trend}].`);
    }
    if (lbb.length) {
      const top = lbb
        .slice()
        .sort((x, y) => Number(y.ftod_actual) - Number(x.ftod_actual))
        .slice(0, 5)
        .map((r) => `${r.branch_name} ${fmt(r.ftod_actual)}`)
        .join(', ');
      md.push(`FTOD top branches (${lbbLabel}): ${top}.`);
    }

    // ---------- DPD bands ----------
    md.push('');
    md.push(
      `DPD 1-30: today plan ${fmt(L.dpd_1_30_plan)}, actual ${fmt(L.dpd_1_30_actual)}` +
        (pct(L.dpd_1_30_actual, L.dpd_1_30_plan) != null
          ? ` (${pct(L.dpd_1_30_actual, L.dpd_1_30_plan)}%)`
          : '') +
        '.'
    );
    if (lbb.length) {
      const top = lbb
        .slice()
        .sort((x, y) => Number(y.dpd_1_30_actual) - Number(x.dpd_1_30_actual))
        .slice(0, 5)
        .map((r) => `${r.branch_name} ${fmt(r.dpd_1_30_actual)}`)
        .join(', ');
      md.push(`DPD 1-30 top branches (${lbbLabel}): ${top}.`);
    }
    md.push(
      `DPD 31-60: today plan ${fmt(L.dpd_31_60_plan)}, actual ${fmt(L.dpd_31_60_actual)}` +
        (pct(L.dpd_31_60_actual, L.dpd_31_60_plan) != null
          ? ` (${pct(L.dpd_31_60_actual, L.dpd_31_60_plan)}%)`
          : '') +
        '.'
    );
    if (lbb.length) {
      const top = lbb
        .slice()
        .sort((x, y) => Number(y.dpd_31_60_actual) - Number(x.dpd_31_60_actual))
        .slice(0, 5)
        .map((r) => `${r.branch_name} ${fmt(r.dpd_31_60_actual)}`)
        .join(', ');
      md.push(`DPD 31-60 top branches (${lbbLabel}): ${top}.`);
    }
    md.push(
      `DPD 61-90: today plan ${fmt(L.dpd_61_90_plan)}, actual ${fmt(L.dpd_61_90_actual)}` +
        (pct(L.dpd_61_90_actual, L.dpd_61_90_plan) != null
          ? ` (${pct(L.dpd_61_90_actual, L.dpd_61_90_plan)}%)`
          : '') +
        '.'
    );
    if (lbb.length) {
      const top = lbb
        .slice()
        .sort((x, y) => Number(y.dpd_61_90_actual) - Number(x.dpd_61_90_actual))
        .slice(0, 5)
        .map((r) => `${r.branch_name} ${fmt(r.dpd_61_90_actual)}`)
        .join(', ');
      md.push(`DPD 61-90 top branches (${lbbLabel}): ${top}.`);
    }
  }

  // ---------- Collection (rolling + monthly) ----------
  if (ctx.summary && Object.keys(ctx.summary).length) {
    md.push('');
    const s = ctx.summary;
    const rd = Number(s.total_rd) || 0;
    const rc = Number(s.total_rc) || 0;
    const collPct = pct(rc, rd);
    md.push(
      `COLLECTION rolling totals: Demand ${fmtAmt(rd)}, Collection ${fmtAmt(rc)}` +
        (collPct != null ? ` (${collPct}%)` : '') +
        `. NPA cases ${fmt(s.npa_cases)}, NPA activated ${fmt(s.npa_act)}.`
    );
  }
  if (Array.isArray(ctx.monthly) && ctx.monthly.length) {
    const recent = ctx.monthly.slice(0, 4);
    const monthLine = recent
      .map(
        (m) =>
          `${m.month}: collection ${fmtAmt(m.total_collection)}/demand ${fmtAmt(
            m.total_demand
          )} (${m.collection_pct}%)`
      )
      .join('; ');
    md.push(`Recent months — ${monthLine}.`);
    if (ctx.monthly.length >= 2) {
      const cur = Number(ctx.monthly[0].total_collection) || 0;
      const prev = Number(ctx.monthly[1].total_collection) || 0;
      const delta = cur - prev;
      const pctDelta = prev > 0 ? Number(((delta / prev) * 100).toFixed(1)) : null;
      md.push(
        `MoM collection delta (${ctx.monthly[1].month} → ${ctx.monthly[0].month}): ${
          delta >= 0 ? '+' : ''
        }${fmtAmt(delta)}` + (pctDelta != null ? ` (${pctDelta >= 0 ? '+' : ''}${pctDelta}%)` : '') + '.'
      );
    }
  }

  // ---------- Disbursement ----------
  if (ctx.latest) {
    const L = ctx.latest;
    md.push('');
    md.push(
      `DISBURSEMENT today ${L.date}: IGL ${fmt(L.disb_igl_acc)} accs / ${fmtAmt(L.disb_igl_amt)}, ` +
        `FIG ${fmt(L.disb_fig_acc)} accs / ${fmtAmt(L.disb_fig_amt)}, ` +
        `IL ${fmt(L.disb_il_acc)} accs / ${fmtAmt(L.disb_il_amt)}.` +
        ` KYC today: IGL ${fmt(L.kyc_igl)}, FIG ${fmt(L.kyc_fig)}, IL ${fmt(L.kyc_il)}.`
    );
  }
  if (Array.isArray(ctx.disbursementMonthly) && ctx.disbursementMonthly.length) {
    const d = ctx.disbursementMonthly.slice(0, 3);
    const line = d
      .map((m) => `${m.month}: ${fmt(m.count)} accs / ${fmtAmt(m.amount)}`)
      .join('; ');
    md.push(`Recent disb months — ${line}.`);
  }

  // ---------- Employees (top performers) ----------
  if (Array.isArray(ctx.employeePerformance) && ctx.employeePerformance.length) {
    md.push('');
    const top = ctx.employeePerformance
      .slice()
      .sort((a, b) => (Number(b.total_collection) || 0) - (Number(a.total_collection) || 0))
      .slice(0, 5)
      .map(
        (e) =>
          `${e.name || e.emp_id} (${e.branch || '-'}) ${fmtAmt(e.total_collection)} / ${fmtAmt(
            e.total_demand
          )} (${e.collection_pct}%)`
      )
      .join('; ');
    md.push(`EMPLOYEES top 5 by FY-to-date collection: ${top}.`);
  }

  // ---------- Truncation notice ----------
  if (ctx.truncated) {
    const t = ctx.truncated;
    const flagged = Object.keys(t).filter((k) => t[k]);
    if (flagged.length) {
      md.push('');
      md.push(
        `NOTE: the following lists were capped (full set has more rows): ${flagged.join(
          ', '
        )}. Say "showing top N" rather than implying you see every row.`
      );
    }
  }

  return md.join('\n');
}

// Compact reference of the underlying schema. Used by /api/ai-snapshot so
// offline mobile clients carry the same DB cheatsheet as the AI prompt.
// Pruned to tables actually populated in ctx — tables the snapshot does NOT
// surface (hourly_performance, npa_activation_runs, employee_locations,
// chat_*) were dropped to stop the model expecting them and hallucinating
// values. See offline-ai-trace audit P1 #4.
const AI_SCHEMA_CHEATSHEET = {
  branches: 'branch_id, branch_name, district_id',
  employees: 'emp_id, full_name, role, branch_id',
  employee_master: 'emp_id, full_name, role, designation, branch_name, area_name, area_manager, division_name, division_manager, region_name, mobile, status (HR roster, single source of truth for hierarchy)',
  employee_performance: 'emp_id, regular_demand, regular_collection, demand_1_30, collection_1_30, demand_31_60, collection_31_60, pnpa_demand, pnpa_collection, npa_cases, npa_act_acc, npa_act_amt, ... (rolling totals; NO date column)',
  daily_performance: 'emp_id, report_date, same metric columns as employee_performance (historical daily snapshots)',
  disbursement: 'db_month, region_name, district_name, branch_name, emp_id, officer_name, product_name, disb_count, disb_amount (monthly aggregate)',
  disbursement_daily: 'disb_date, region_name, district_name, branch_name, emp_id, officer_name, product_name, disb_count, disb_amount (daily — UNIONed via DISB_CTE for months not yet rolled into disbursement)',
  daily_reports: 'branch_name, date, region, district, dm_name, ftod_plan, dpd_1_30_plan, dpd_31_60_plan, dpd_61_90_plan, fy_non_start_plan, ... (PLAN side of daily reports)',
  daily_reports_achievements: 'branch_name, date, region, district, dm_name, ftod_actual, dpd_1_30_actual, dpd_31_60_actual, dpd_61_90_actual, npa_activation, npa_closure, fy_non_start_acc, disb_igl_acc/amt, disb_fig_acc/amt, disb_il_acc/amt, kyc_igl, kyc_fig, kyc_il (ACTUAL side of daily reports)'
};

const AI_SNAPSHOT_VERSION = 2;

// Bump when AI prompt or tool surface changes — invalidates aiReplyCache keys.
const AI_REPLY_CACHE_VERSION = 'v8-single-entity-employee-series';

// Mistral function-calling tool definitions. The model picks which to call
// when the bundled ctx isn't enough. Every tool maps to a parameterised SQL
// query in dispatchAiTool() below, scope-respecting via _scopeWhere().
const AI_TOOLS_SPEC = [
  {
    type: 'function',
    function: {
      name: 'find_employee',
      description: 'Look up employees by name, mobile, or emp_id. Substring + trigram fuzzy + word-similarity fallback so misheard names ("AP Shivraj" / "EPI Shivraj" → "A P Shivaraj") still resolve. Returns role, branch, area, division, region, mobile, status. Pass location_hint when the user names a branch / area / district / region alongside the person — it narrows ambiguous matches. If multiple matches still come back, list them and ask which one the user means.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Name or mobile or emp_id. Misspellings tolerated via fuzzy match.' },
          location_hint: { type: 'string', description: 'Optional: branch / area / district / region the user mentioned alongside the name. Used to disambiguate when multiple employees share a name (e.g. query="Shivraj", location_hint="Raichur").' },
          limit: { type: 'integer', description: 'Max rows. Default 25, max 200.' }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'employee_performance',
      description: 'Collection/demand/NPA TOTALS (single aggregate row) for ONE employee over a date range. Default range = current FY-to-date. Use for "<person> performance / <person> totals / how has <person> done this FY". For a per-day TREND/series (chart), use employee_collection_series instead. NEVER substitute period_performance(group_by="employee", branch_name=...) for this — that returns the FULL branch leaderboard, not the named person.',
      parameters: {
        type: 'object',
        properties: {
          emp_id: { type: 'string' },
          start_date: { type: 'string', description: 'YYYY-MM-DD. Default = current FY start (Apr 1).' },
          end_date: { type: 'string', description: 'YYYY-MM-DD. Default = today.' }
        },
        required: ['emp_id']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'employee_collection_series',
      description: 'Per-DAY collection / demand / collection % / NPA series for ONE employee over a date range. Returns one row per report_date (the trend chart input). Use for "show <person> trend / <person> performance over time / drill on <person> / day-by-day for <person> / how is <person> doing day to day". Default range = today minus 30 days → today. NEVER substitute period_performance(group_by="employee", branch_name=...) or top_performers(branch_name=...) for this — those return the FULL branch leaderboard (every employee), not the single named person. ALWAYS canonicalise the name with find_employee first and pass the resulting emp_id.',
      parameters: {
        type: 'object',
        properties: {
          emp_id: { type: 'string', description: 'Canonical emp_id from find_employee.' },
          start_date: { type: 'string', description: 'YYYY-MM-DD. Default = today minus 30 days.' },
          end_date: { type: 'string', description: 'YYYY-MM-DD. Default = today.' }
        },
        required: ['emp_id']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'find_branch',
      description: 'Look up branches by name. Returns branch + region/division/area + headcount + current rolling perf (demand/collection/NPA). If multiple branches match, do not choose one silently; list the matches and ask which branch the user means.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Substring of branch_name (ILIKE).' },
          limit: { type: 'integer' }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'period_performance',
      description: 'Aggregate collection/demand/NPA from daily_performance over a date range. Group by day, month, branch, or employee. Optional filters: branch_name / area_name / division_name / region_name (one or more).',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD' },
          end_date: { type: 'string', description: 'YYYY-MM-DD' },
          group_by: { type: 'string', enum: ['day', 'month', 'branch', 'employee'], description: 'Default month.' },
          branch_name: { type: 'string' },
          area_name: { type: 'string' },
          division_name: { type: 'string' },
          region_name: { type: 'string' },
          limit: { type: 'integer' }
        },
        required: ['start_date', 'end_date']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'top_performers',
      description: 'Leaderboard of EMPLOYEES (not branches) by metric over a date range. For top branches, use period_performance(group_by="branch") instead. Filters: branch_name / area_name / division_name / region_name / role narrow the leaderboard to a sub-population.',
      parameters: {
        type: 'object',
        properties: {
          metric: { type: 'string', enum: ['collection', 'demand', 'npa_cases'] },
          start_date: { type: 'string' },
          end_date: { type: 'string' },
          branch_name: { type: 'string' },
          area_name: { type: 'string' },
          division_name: { type: 'string' },
          region_name: { type: 'string' },
          role: { type: 'string', description: 'Filter by exact role (FO, BM, etc.)' },
          limit: { type: 'integer' }
        },
        required: ['metric', 'start_date', 'end_date']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'disbursement_query',
      description: 'Disbursement count + amount over a date range. group_by determines the row breakdown — pick it from the user\'s words: "product wise / by product" → product; "branch wise" → branch; "day by day / Apr 7 vs 8" → day; "monthly trend" → month; "by FO / by employee / per officer" → employee. If the user asks for TWO breakdowns ("FO wise AND product wise"), CALL THIS TOOL TWICE with different group_by values and present both. Optional filters: branch_name / region_name.',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD.' },
          end_date: { type: 'string' },
          group_by: {
            type: 'string',
            enum: ['day', 'month', 'branch', 'product', 'employee'],
            description: 'PICK FROM USER WORDS: product-wise → "product", branch-wise → "branch", day-by-day → "day", monthly → "month", by FO/employee → "employee". When user asks for multiple, call the tool once per breakdown.'
          },
          branch_name: { type: 'string' },
          region_name: { type: 'string' },
          limit: { type: 'integer' }
        },
        required: ['start_date', 'end_date']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_hierarchy',
      description: 'List children of a hierarchy node from employee_master. Examples: list_hierarchy(level="branch", parent_level="area", parent_name="X") → branches under area X. list_hierarchy(level="region") → all regions in scope. Default limit is 500 — there are ~129 branches and ~8 regions, so the default returns everything. NEVER quote the row count of this tool as a metric value.',
      parameters: {
        type: 'object',
        properties: {
          level: { type: 'string', enum: ['region', 'division', 'area', 'branch'] },
          parent_level: { type: 'string', enum: ['region', 'division', 'area'] },
          parent_name: { type: 'string' },
          limit: { type: 'integer' }
        },
        required: ['level']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'list_employees',
      description: 'Roster query — list employees by branch / area / division / region / role. Use for "list everyone in <branch>", "show all FOs in <region>", "who works in <branch>". Returns name, emp_id, role, branch, area, division, region, mobile, status. Use this NOT find_employee for "list all" queries — find_employee is for searching a single named person.',
      parameters: {
        type: 'object',
        properties: {
          branch_name: { type: 'string' },
          area_name: { type: 'string' },
          division_name: { type: 'string' },
          region_name: { type: 'string' },
          role: { type: 'string', description: 'Filter by exact role (BM, FO, AM, etc.)' },
          status: { type: 'string', enum: ['Working', 'Resigned', 'all'], description: 'Default Working.' },
          limit: { type: 'integer', description: 'Default 100, max 500.' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'headcount',
      description: 'Total Working employees in scope, optionally broken down by role / region / division / area / branch. Examples: headcount() → {total: 1285}. headcount(group_by="role") → [{role:"FO",count:791}, …]. headcount(role="BM") → [{role:"BM", count:122}].',
      parameters: {
        type: 'object',
        properties: {
          group_by: { type: 'string', enum: ['role', 'region', 'division', 'area', 'branch'] },
          role: { type: 'string', description: 'Filter by exact role (BM, FO, AM, RM, DM, etc.)' },
          region_name: { type: 'string' },
          branch_name: { type: 'string' },
          limit: { type: 'integer' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'npa_summary',
      description: 'NPA snapshot rollup over a date range. WARNING: npa_cases and npa_act_acc/amt come from daily_performance which stores ROLLING SNAPSHOTS (the same outstanding case is repeated every day). SUM over a multi-day range therefore inflates the total by ~days. **For outstanding NPA cases, pass start_date == end_date (latest report_date) — that returns the true snapshot count.** For NPA activation / closure RATES over a period, use daily_reports_query(table="achievements", metrics="npa") instead — daily_reports_achievements.npa_activation and .npa_closure are per-day DELTAS that sum correctly.',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string' },
          end_date: { type: 'string' },
          group_by: { type: 'string', enum: ['month', 'branch', 'employee'] },
          branch_name: { type: 'string' },
          region_name: { type: 'string' },
          limit: { type: 'integer' }
        },
        required: ['start_date', 'end_date']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'daily_reports_query',
      description: 'Branch-level daily plan + achievement data. Source for FTOD (First Time Over Due), DPD buckets (1-30, 31-60, 61-90), NPA activation/closure, disbursement plan/achievement (IGL/FIG/IL count + amount), KYC counts. Each row = one branch on one date. Use for ANY question mentioning FTOD, DPD bucket, NPA closure, disbursement plan-vs-actual, KYC.',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD. For a single date, use same value for start_date and end_date.' },
          end_date: { type: 'string', description: 'YYYY-MM-DD' },
          branch_name: { type: 'string' },
          region_name: { type: 'string' },
          district_name: { type: 'string' },
          table: { type: 'string', enum: ['plan', 'achievements', 'both'], description: 'Default both — daily_reports (plan) and daily_reports_achievements (actual).' },
          metrics: { type: 'string', description: 'Comma list to project. Available: ftod, dpd_1_30, dpd_31_60, dpd_61_90, npa, disb, kyc, fy_non_start, all. Default all.' },
          limit: { type: 'integer' }
        },
        required: ['start_date', 'end_date']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'branch_summary',
      description: 'ONE-SHOT branch dashboard — combines headcount + role mix, collection (today + MTD + FYTD with %), NPA (cases / activation amount / closure amount), disbursement (today + MTD count + amount), FTOD plan-vs-actual, and DPD bucket (1-30 / 31-60) plan-vs-actual into a SINGLE result. Use for "how is my branch doing today", "branch health check", "give me a summary of <branch>", "end-of-day rollup". Replaces 5+ chained tool calls. For BM/ABM/BOE the branch is auto-resolved from their session — they should call branch_summary() with no args. For CEO/RM/DM/AM, branch_name is REQUIRED (this is a single-branch report, not a region rollup).',
      parameters: {
        type: 'object',
        properties: {
          branch_name: { type: 'string', description: 'Optional for branch-bound roles (auto-resolved from session). Required for CEO/RM/DM/AM.' },
          date: { type: 'string', description: 'YYYY-MM-DD as-of date. Default = latest report_date with data for this branch.' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'hourly_collection',
      description: 'Returns the CURRENT live snapshot — single point in time, NOT a historical series. For day-by-day trends use period_performance(group_by=\'day\'). hourly_performance is one live snapshot (no date, no hour buckets, no time series); UPSERTed by intra-day uploads, reset after EOD. Use this tool ONLY for "collection right now / live collection / how much collected today so far / intra-day / show top FOs right now". NEVER use it for "yesterday\'s hourly" or any historical question — that data does not exist. Returns current demand + collection (counts AND amounts) with collection %, optionally grouped by branch / region / employee. Branch-bound roles see only their own branch.',
      parameters: {
        type: 'object',
        properties: {
          branch_name: { type: 'string' },
          region_name: { type: 'string' },
          group_by: { type: 'string', enum: ['branch', 'region', 'employee', 'none'], description: 'Default: "branch" when caller has multi-branch scope; "none" (single aggregate) for branch-bound roles or when branch_name is passed. Use "employee" for "top FOs right now / live employee leaderboard".' },
          limit: { type: 'integer' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'collection_drilldown',
      description: 'Single-call answer to "why is collection low for my branch / drill down on collection". Returns 3 worst FOs (top_3_underperformers, sorted ascending by collection %), 3 best FOs (bottom_3_underperformers, sorted descending — naming preserved per spec), DPD bucket plan-vs-actual (1-30 / 31-60 / 61-90), NPA today (count + amount), and FTOD gap today (actual / plan / gap). Use AFTER branch_summary when the user asks "why" or "drill down". For BM/ABM/BOE branch_name auto-resolves from session; CEO/RM/DM/AM must pass branch_name.',
      parameters: {
        type: 'object',
        properties: {
          branch_name: { type: 'string', description: 'Optional for branch-bound roles. Required for CEO/RM/DM/AM.' },
          date: { type: 'string', description: 'YYYY-MM-DD as-of date. Default = latest report_date with data for this branch.' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'period_compare',
      description: 'Compute a SINGLE metric across two date windows and return both values + delta_abs + delta_pct. Use for "MoM / WoW / vs last month / this week vs last week / vs yesterday / vs same week last year". Server-side math eliminates the model\'s frequent count-vs-amount mix-ups and broken pct-of-pct arithmetic. metric ∈ {collection, demand, collection_pct, npa_amount, disb_amount, disb_count, ftod}. scope ∈ {all, branch, region, division, area}. scope_value required when scope != "all".',
      parameters: {
        type: 'object',
        properties: {
          metric: { type: 'string', enum: ['collection', 'demand', 'collection_pct', 'npa_amount', 'disb_amount', 'disb_count', 'ftod'] },
          scope: { type: 'string', enum: ['all', 'branch', 'region', 'division', 'area'], description: 'Default "all".' },
          scope_value: { type: 'string', description: 'Required when scope != "all". E.g. branch name, region name.' },
          period_a_start: { type: 'string', description: 'YYYY-MM-DD' },
          period_a_end:   { type: 'string', description: 'YYYY-MM-DD' },
          period_b_start: { type: 'string', description: 'YYYY-MM-DD' },
          period_b_end:   { type: 'string', description: 'YYYY-MM-DD' }
        },
        required: ['metric', 'period_a_start', 'period_a_end', 'period_b_start', 'period_b_end']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'plan_compliance',
      description: 'Lists branches that did NOT file daily_reports (plan) and/or daily_reports_achievements (achievement) for the given date. Use for "which branches missed plan today / who didn\'t file / branches without daily report / plan compliance". Returns expected_branches (from employee_master), filed_plan, filed_achievement, missing_plan, missing_achievement. CEO sees all; RM/DM/AM see their scope; BM/ABM/BOE rejected (single-branch — compliance check is meaningless).',
      parameters: {
        type: 'object',
        properties: {
          date: { type: 'string', description: 'YYYY-MM-DD. Default = today.' },
          scope: { type: 'string', enum: ['all', 'region', 'division', 'area'], description: 'Default "all".' },
          scope_value: { type: 'string', description: 'Required when scope != "all".' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'resolve_date_range',
      description: 'Deterministically converts relative or ambiguous date expressions into ISO YYYY-MM-DD ranges using Asia/Kolkata. Use before any data tool when the user says today, yesterday, this month, last month, this week, last week, FYTD, a bare month, or a spoken/slash date like 11/04.',
      parameters: {
        type: 'object',
        properties: {
          expression: { type: 'string', description: 'Date phrase to resolve, e.g. today, yesterday, this month, last month, April 11, 11/04, FYTD.' },
          comparison_expression: { type: 'string', description: 'Optional second phrase for comparisons, e.g. last month.' },
          anchor_date: { type: 'string', description: 'Optional YYYY-MM-DD override for today. Defaults to current Asia/Kolkata date.' }
        },
        required: ['expression']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'sql_describe',
      description: 'Returns the schema cheatsheet of all tables/columns the AI can query.',
      parameters: { type: 'object', properties: {} }
    }
  }
];

// Build " AND UPPER(<alias>) IN (SELECT ... FROM employee_master WHERE ...)"
// for the current session. Mutates `params` to append $loc. CEO/empty session → ''.
// Hard scope guard at the dispatch layer. _scopeWhere already adds an
// intersection filter to every SQL query, but a Branch Manager asking
// about another branch silently gets 0 rows — confusing. Detect that
// upfront and return an explicit error so the model can tell the user
// "you can only access your own branch".
function _scopeViolation(session, args) {
  const role = String((session && session.role) || '').toUpperCase();
  const loc = String((session && session.location) || '').trim();
  if (!role || role === 'CEO' || role === 'DIRECTOR' || !loc) return null;
  const norm = (s) => String(s || '').trim().toLowerCase();
  if (role === 'BM' || role === 'ABM' || role === 'BOE') {
    // Branch-bound roles can only see their own branch.
    if (args.branch_name && norm(args.branch_name) !== norm(loc)) {
      return `${role} can only access ${loc} branch. Asking about ${args.branch_name} is not allowed.`;
    }
    if (args.region_name || args.area_name || args.division_name) {
      return `${role} can only access their own branch (${loc}). Region/division/area filters are not allowed for this role.`;
    }
  }
  return null;
}

function _scopeWhere(session, alias, params) {
  const role = String((session && session.role) || '').toUpperCase();
  const loc = String((session && session.location) || '').trim();
  if (!role || role === 'CEO' || !loc) return '';
  let emCol = null;
  if (role === 'RM' || role === 'SM') emCol = 'region_name';
  else if (role === 'DM' || role === 'DVM') emCol = 'division_name';
  else if (role === 'AM') emCol = 'area_name';
  else if (role === 'BM' || role === 'FO') emCol = 'branch_name';
  if (!emCol) return '';
  params.push(loc);
  return ` AND UPPER(${alias}) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(${emCol}) ILIKE TRIM($${params.length}))`;
}

async function resolveCanonicalBranchName(input, session) {
  const q = String(input || '').trim();
  if (!q) return { error: 'branch_name required' };

  async function run(whereClause, params) {
    const scoped = _scopeWhere(session, 'branch_name', params);
    return safeQuery(
      `SELECT branch_name,
              MAX(region_name) AS region,
              MAX(division_name) AS division,
              MAX(area_name) AS area,
              COUNT(*) FILTER (WHERE status='Working')::int AS employee_count
         FROM employee_master
        WHERE ${whereClause}
          ${scoped}
        GROUP BY branch_name
        ORDER BY branch_name
        LIMIT 10`,
      params,
      10
    );
  }

  let params = [`%${q}%`];
  let rows = await run('branch_name ILIKE $1', params);
  if (Array.isArray(rows) && rows.length === 0 && q.length >= 4) {
    params = [q];
    rows = await run('similarity(branch_name, $1) > 0.5', params);
  }
  if (!Array.isArray(rows)) return rows;
  if (!rows.length) {
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    if ((role === 'BM' || role === 'ABM' || role === 'BOE') && loc) {
      return { error: 'scope_violation', message: `${role} can only access ${loc} branch. Asking about "${q}" is not allowed.` };
    }
    return { error: 'branch_not_found', message: `No branch matched "${q}". Call find_branch or ask the user to confirm the branch name.` };
  }
  const exact = rows.filter((r) => normalizeLookupValue(r.branch_name) === normalizeLookupValue(q));
  if (exact.length === 1) return { ok: true, branch_name: exact[0].branch_name };
  if (rows.length === 1) return { ok: true, branch_name: rows[0].branch_name };
  return {
    error: 'ambiguous_branch',
    message: `Multiple branches matched "${q}". Ask the user which branch they mean before querying performance data.`,
    matches: rows
  };
}

function normalizeLookupValue(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeDigits(value) {
  return String(value || '').replace(/\D/g, '');
}

function isExactEmployeeLookup(row, query) {
  const q = normalizeLookupValue(query);
  const digits = normalizeDigits(query);
  return normalizeLookupValue(row && row.emp_id) === q ||
    (!!digits && normalizeDigits(row && row.mobile) === digits);
}

function isExactNameMatch(value, query) {
  return normalizeLookupValue(value) === normalizeLookupValue(query);
}

function containsKannadaScript(value) {
  return /[\u0C80-\u0CFF]/.test(String(value || ''));
}

function normalizeKannadaNumbersForAi(value) {
  let out = String(value == null ? '' : value);
  const digitMap = {
    '೦': '0', '೧': '1', '೨': '2', '೩': '3', '೪': '4',
    '೫': '5', '೬': '6', '೭': '7', '೮': '8', '೯': '9'
  };
  out = out.replace(/[೦-೯]/g, (d) => digitMap[d] || d);
  if (!containsKannadaScript(out)) return out;

  const replacements = [
    ['ಕೋಟಿಗಳು', 'crore'], ['ಕೋಟಿ', 'crore'],
    ['ಲಕ್ಷಗಳು', 'lakh'], ['ಲಕ್ಷ', 'lakh'],
    ['ಸಾವಿರಗಳು', 'thousand'], ['ಸಾವಿರ', 'thousand'],
    ['ನೂರು', 'hundred'],
    ['ರೂಪಾಯಿಗಳು', 'rupees'], ['ರೂಪಾಯಿ', 'rupees'], ['ರೂ.', 'rupees'],
    ['ಪ್ರತಿಶತ', 'percent'], ['ಶೇಕಡಾ', 'percent'],
    ['ಶೂನ್ಯ', 'zero'], ['ಒಂದು', 'one'], ['ಎರಡು', 'two'], ['ಮೂರು', 'three'],
    ['ನಾಲ್ಕು', 'four'], ['ಐದು', 'five'], ['ಆರು', 'six'], ['ಏಳು', 'seven'],
    ['ಎಂಟು', 'eight'], ['ಒಂಬತ್ತು', 'nine'], ['ಹತ್ತು', 'ten'],
    ['ಹನ್ನೊಂದು', 'eleven'], ['ಹನ್ನೆರಡು', 'twelve'], ['ಹದಿಮೂರು', 'thirteen'],
    ['ಹದಿನಾಲ್ಕು', 'fourteen'], ['ಹದಿನೈದು', 'fifteen'], ['ಹದಿನಾರು', 'sixteen'],
    ['ಹದಿನೇಳು', 'seventeen'], ['ಹದಿನೆಂಟು', 'eighteen'], ['ಹತ್ತೊಂಬತ್ತು', 'nineteen'],
    ['ಇಪ್ಪತ್ತು', 'twenty'], ['ಮೂವತ್ತು', 'thirty'], ['ನಲವತ್ತು', 'forty'],
    ['ಐವತ್ತು', 'fifty'], ['ಅರವತ್ತು', 'sixty'], ['ಎಪ್ಪತ್ತು', 'seventy'],
    ['ಎಂಭತ್ತು', 'eighty'], ['ತೊಂಬತ್ತು', 'ninety']
  ];

  for (const [kn, en] of replacements) {
    const escaped = kn.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    out = out.replace(new RegExp('(^|[^\\u0C80-\\u0CFF])' + escaped + '(?=$|[^\\u0C80-\\u0CFF])', 'g'), '$1' + en);
  }
  return out;
}

function makeUtcDate(y, m, d) {
  return new Date(Date.UTC(y, m - 1, d));
}

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function addDays(d, n) {
  return makeUtcDate(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate() + n);
}

function monthRange(y, m) {
  return {
    start: makeUtcDate(y, m, 1),
    end: makeUtcDate(y, m + 1, 0),
  };
}

function indiaTodayDate(anchorDate) {
  if (anchorDate && /^\d{4}-\d{2}-\d{2}$/.test(String(anchorDate))) {
    const [y, m, d] = String(anchorDate).split('-').map(Number);
    return makeUtcDate(y, m, d);
  }
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(new Date()).reduce((acc, p) => {
    if (p.type !== 'literal') acc[p.type] = p.value;
    return acc;
  }, {});
  return makeUtcDate(Number(parts.year), Number(parts.month), Number(parts.day));
}

function resolveDateRangeExpression(expression, anchorDate) {
  const today = indiaTodayDate(anchorDate);
  const fyStartYear = today.getUTCMonth() + 1 >= 4 ? today.getUTCFullYear() : today.getUTCFullYear() - 1;
  let raw = String(expression || '').trim();
  let expr = raw.toLowerCase()
    .replace(/,/g, ' ')
    .replace(/\b(\d{1,2})(st|nd|rd|th)\b/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
  if (!expr) expr = 'today';

  const monthNames = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3,
    apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7,
    aug: 8, august: 8, sep: 9, sept: 9, september: 9,
    oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12
  };
  const pack = (start, end, label) => ({
    start_date: isoDate(start),
    end_date: isoDate(end),
    label,
    anchor_date: isoDate(today),
    timezone: 'Asia/Kolkata'
  });

  if (/^\d{4}-\d{2}-\d{2}$/.test(expr)) {
    const [y, m, d] = expr.split('-').map(Number);
    const dt = makeUtcDate(y, m, d);
    return pack(dt, dt, raw || expr);
  }
  if (expr === 'today') return pack(today, today, 'today');
  if (expr === 'yesterday') {
    const d = addDays(today, -1);
    return pack(d, d, 'yesterday');
  }
  if (expr === 'this month' || expr === 'current month' || expr === 'mtd') {
    return pack(makeUtcDate(today.getUTCFullYear(), today.getUTCMonth() + 1, 1), today, 'this month');
  }
  if (expr === 'last month' || expr === 'previous month') {
    const r = monthRange(today.getUTCFullYear(), today.getUTCMonth());
    return pack(r.start, r.end, 'last month');
  }
  if (expr === 'fytd' || expr === 'fy to date' || expr === 'this fy' || expr === 'current fy' || expr === 'this year') {
    return pack(makeUtcDate(fyStartYear, 4, 1), today, 'FY-to-date');
  }
  if (expr === 'this week' || expr === 'current week') {
    const day = today.getUTCDay() || 7;
    return pack(addDays(today, 1 - day), today, 'this week');
  }
  if (expr === 'last week' || expr === 'previous week') {
    const day = today.getUTCDay() || 7;
    const end = addDays(today, -day);
    return pack(addDays(end, -6), end, 'last week');
  }

  const slash = expr.match(/^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/);
  if (slash) {
    const dd = Number(slash[1]);
    const mm = Number(slash[2]);
    let yy = slash[3] ? Number(slash[3]) : (mm >= 4 ? fyStartYear : fyStartYear + 1);
    if (yy < 100) yy += 2000;
    const dt = makeUtcDate(yy, mm, dd);
    return pack(dt, dt, raw);
  }

  const monthDay = expr.match(/\b([a-z]+)\s+(\d{1,2})\b/) || expr.match(/\b(\d{1,2})\s+([a-z]+)\b/);
  if (monthDay) {
    const firstIsMonth = monthNames[monthDay[1]];
    const mm = firstIsMonth || monthNames[monthDay[2]];
    const dd = Number(firstIsMonth ? monthDay[2] : monthDay[1]);
    if (mm && dd) {
      const yy = mm >= 4 ? fyStartYear : fyStartYear + 1;
      const dt = makeUtcDate(yy, mm, dd);
      return pack(dt, dt, raw);
    }
  }

  const monthOnly = monthNames[expr];
  if (monthOnly) {
    const yy = monthOnly >= 4 ? fyStartYear : fyStartYear + 1;
    const r = monthRange(yy, monthOnly);
    return pack(r.start, r.end, raw);
  }

  return {
    error: 'unresolved_date_expression',
    message: `Could not resolve "${raw}". Ask for a concrete date or period.`,
    anchor_date: isoDate(today),
    timezone: 'Asia/Kolkata'
  };
}

async function dispatchAiTool(name, args, session) {
  args = args || {};
  const lim = Math.min(Math.max(parseInt(args.limit, 10) || 25, 1), 200);

  if (name === 'sql_describe') return AI_SCHEMA_CHEATSHEET;

  if (name === 'resolve_date_range') {
    let expression = String(args.expression || 'today').trim();
    let comparisonExpression = String(args.comparison_expression || '').trim();
    if (!comparisonExpression && /\b(vs|versus)\b/i.test(expression)) {
      const parts = expression.split(/\b(?:vs|versus)\b/i).map((s) => s.trim()).filter(Boolean);
      if (parts.length >= 2) {
        expression = parts[0];
        comparisonExpression = parts.slice(1).join(' ');
      }
    }
    const primary = resolveDateRangeExpression(expression, args.anchor_date);
    if (comparisonExpression && !primary.error) {
      const comparison = resolveDateRangeExpression(comparisonExpression, args.anchor_date);
      if (!comparison.error) {
        primary.comparison_start_date = comparison.start_date;
        primary.comparison_end_date = comparison.end_date;
        primary.comparison_label = comparison.label;
      } else {
        primary.comparison_error = comparison;
      }
    }
    return primary;
  }

  if (args.branch_name && name !== 'find_branch') {
    const resolved = await resolveCanonicalBranchName(args.branch_name, session);
    if (!resolved || resolved.error) return resolved;
    args = { ...args, branch_name: resolved.branch_name };
  }
  if (name === 'period_compare' && String(args.scope || '').toLowerCase() === 'branch' && args.scope_value) {
    const resolved = await resolveCanonicalBranchName(args.scope_value, session);
    if (!resolved || resolved.error) return resolved;
    args = { ...args, scope_value: resolved.branch_name };
  }

  // Hard scope guard — BM/ABM/BOE may only query their own branch.
  // Block before any DB call so the model sees a clear refusal it can relay.
  const violation = _scopeViolation(session, args);
  if (violation) {
    return { error: 'scope_violation', message: violation };
  }

  if (name === 'find_employee') {
    const q = String(args.query || '').trim();
    if (!q) return { error: 'query required' };

    const buildSql = (whereClause, scopeParams) => `
      SELECT em.emp_id, em.full_name AS name, em.role, em.designation,
             em.branch_name AS branch, em.area_name AS area,
             em.division_name AS division, em.region_name AS region,
             em.mobile, em.status
        FROM employee_master em
       WHERE ${whereClause}
         ${_scopeWhere(session, 'em.branch_name', scopeParams)}
       ORDER BY (em.status='Working') DESC, em.full_name
       LIMIT ${lim}`;

    // Pass 1: substring ILIKE on name + exact mobile/emp_id.
    let params = [`%${q}%`, q];
    let rows = await safeQuery(buildSql(`(em.full_name ILIKE $1 OR em.mobile = $2 OR em.emp_id = $2)`, params), params, lim);

    // Pass 2: trigram fuzzy on full_name when ILIKE returned nothing.
    // Whisper drops letters / collapses initials ("AP Shivraj" vs the real
    // "A P Shivaraj") and even mangles initials ("EPI Shivraj"). Use BOTH
    // word_similarity (matches any embedded token, e.g. just "Shivraj"
    // against "A P Shivaraj") and full-string similarity. word_similarity
    // is the more useful metric for messy multi-word names — pulls every
    // candidate that contains a near-match for any word in the query.
    if (Array.isArray(rows) && rows.length === 0 && q.length >= 4) {
      params = [q];
      rows = await safeQuery(
        buildSql(
          `(word_similarity($1, em.full_name) > 0.4 OR similarity(em.full_name, $1) > 0.3)`,
          params
        ),
        params, lim
      );
    }
    // Pass 3: optional location hint. If the user provided a branch /
    // area / division / region in the query (the LLM passes it in via
    // the `location_hint` arg), use that to narrow ambiguous matches.
    if (Array.isArray(rows) && rows.length > 1 && args.location_hint) {
      const hint = String(args.location_hint).trim();
      if (hint) {
        const filtered = rows.filter(function (r) {
          const fields = [r.branch, r.area, r.division, r.region].map(function (x) { return String(x || '').toLowerCase(); });
          return fields.some(function (f) { return f.indexOf(hint.toLowerCase()) >= 0; });
        });
        if (filtered.length) rows = filtered;
      }
    }

    if (!Array.isArray(rows)) return rows;

    const exactIdentifierMatches = rows.filter((r) => isExactEmployeeLookup(r, q));
    if (rows.length > 1 && exactIdentifierMatches.length !== 1) {
      return {
        ambiguous: true,
        count: rows.length,
        instruction: 'Multiple employees match this lookup. List the top matches with name, emp_id, role, branch, division, and region, then ask which one the user means. Do not answer as if only one employee matched. Do not include phone numbers in the disambiguation list.',
        matches: rows.map((r) => ({
          emp_id: r.emp_id,
          name: r.name,
          role: r.role,
          designation: r.designation,
          branch: r.branch,
          area: r.area,
          division: r.division,
          region: r.region,
          status: r.status
        }))
      };
    }
    if (exactIdentifierMatches.length === 1) return exactIdentifierMatches;
    return rows;
  }

  if (name === 'employee_performance') {
    const empId = String(args.emp_id || '').trim();
    if (!empId) return { error: 'emp_id required' };
    const now = new Date();
    const fy = now.getMonth() + 1 >= 4 ? now.getFullYear() : now.getFullYear() - 1;
    const startDate = args.start_date || `${fy}-04-01`;
    const endDate = args.end_date || now.toISOString().slice(0, 10);
    const params = [empId, startDate, endDate];
    const scopeClause = _scopeWhere(session, 'em.branch_name', params);
    const sql = `
      SELECT em.emp_id, em.full_name AS name, em.role, em.branch_name AS branch,
             em.area_name AS area, em.division_name AS division,
             em.region_name AS region, em.mobile, em.status,
             COALESCE(SUM(dp.regular_demand), 0)::bigint AS demand,
             COALESCE(SUM(dp.regular_collection), 0)::bigint AS collection,
             COALESCE(SUM(dp.npa_cases), 0)::int AS npa_cases,
             COALESCE(SUM(dp.npa_act_amt), 0)::numeric AS npa_amount,
             COUNT(DISTINCT dp.report_date)::int AS days_reported
        FROM employee_master em
        LEFT JOIN daily_performance dp
          ON dp.emp_id = em.emp_id
         AND dp.report_date BETWEEN $2 AND $3
       WHERE em.emp_id = $1
         ${scopeClause}
       GROUP BY em.emp_id, em.full_name, em.role, em.branch_name,
                em.area_name, em.division_name, em.region_name, em.mobile, em.status`;
    return await safeQuery(sql, params, 1);
  }

  if (name === 'employee_collection_series') {
    const empId = String(args.emp_id || '').trim();
    if (!empId) return { error: 'emp_id required' };
    const now = new Date();
    const todayIso = now.toISOString().slice(0, 10);
    const dflt = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const startDate = args.start_date || dflt.toISOString().slice(0, 10);
    const endDate = args.end_date || todayIso;
    const params = [empId, startDate, endDate];
    const scopeClause = _scopeWhere(session, 'em.branch_name', params);

    // 1) Identity row first — confirms the emp_id is in scope, returns
    //    branch/role/etc. so the model can label the chart even when the
    //    series is empty (no daily_performance rows for the window).
    const idSql = `
      SELECT em.emp_id, em.full_name AS name, em.role,
             em.branch_name AS branch, em.area_name AS area,
             em.division_name AS division, em.region_name AS region,
             em.mobile, em.status
        FROM employee_master em
       WHERE em.emp_id = $1
         ${scopeClause}
       LIMIT 1`;
    const idRows = await safeQuery(idSql, params, 1);
    if (!Array.isArray(idRows)) return idRows; // pass-through {error}
    if (!idRows.length) {
      return { error: 'employee_not_found_or_out_of_scope', message: `emp_id ${empId} not found in your scope. Re-run find_employee for the canonical emp_id.` };
    }

    // 2) Per-day series — strict INNER JOIN so 0 rows = "no daily activity in window"
    const seriesSql = `
      SELECT to_char(dp.report_date, 'YYYY-MM-DD') AS date,
             COALESCE(dp.regular_demand, 0)::bigint AS demand,
             COALESCE(dp.regular_collection, 0)::bigint AS collection,
             CASE WHEN COALESCE(dp.regular_demand, 0) > 0
                  THEN ROUND((dp.regular_collection::numeric / dp.regular_demand) * 100, 1)
                  ELSE 0
             END AS collection_pct,
             COALESCE(dp.npa_cases, 0)::int AS npa_cases,
             COALESCE(dp.npa_act_amt, 0)::numeric AS npa_amount
        FROM daily_performance dp
       WHERE dp.emp_id = $1
         AND dp.report_date BETWEEN $2 AND $3
       ORDER BY dp.report_date ASC
       LIMIT ${lim}`;
    const seriesRows = await safeQuery(seriesSql, params, lim);
    if (!Array.isArray(seriesRows)) return seriesRows;

    // 3) Window totals so the model has a headline number alongside the series.
    let totalDemand = 0, totalCollection = 0;
    for (const r of seriesRows) {
      totalDemand += Number(r.demand) || 0;
      totalCollection += Number(r.collection) || 0;
    }
    const totalPct = totalDemand > 0
      ? Math.round((totalCollection / totalDemand) * 1000) / 10
      : 0;

    return {
      employee: idRows[0],
      start_date: startDate,
      end_date: endDate,
      series: seriesRows,
      totals: {
        demand: totalDemand,
        collection: totalCollection,
        collection_pct: totalPct,
        days_reported: seriesRows.length,
      },
      instruction: 'This is a SINGLE-EMPLOYEE day-by-day series. Narrate the trend and totals for THIS employee only. Do NOT mix in other employees from the same branch. If series is empty, say "no daily activity for <name> in <start_date>..<end_date>" — do NOT fall back to a branch leaderboard.',
    };
  }

  if (name === 'find_branch') {
    const q = String(args.query || '').trim();
    if (!q) return { error: 'query required' };

    // Hard scope check: BM/ABM/BOE looking up a branch other than their own.
    // _scopeViolation only fires on explicit branch_name args; here the
    // user passes the branch via `query`, so do an inline check.
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    if ((role === 'BM' || role === 'ABM' || role === 'BOE') && loc) {
      // Quick fuzzy compare — accept Davanagere/Davangere variants only via
      // similarity check on the BM's actual branch name.
      const lq = q.toLowerCase();
      const lloc = loc.toLowerCase();
      if (lq !== lloc && lloc.indexOf(lq) === -1 && lq.indexOf(lloc) === -1) {
        return {
          error: 'scope_violation',
          message: `${role} can only access ${loc} branch. Looking up "${q}" is not allowed.`,
        };
      }
    }

    // Build the perf + agg SQL parameterised on a candidate branch list
    // so we can run the same shape for ILIKE then for trigram-fuzzy.
    const buildSql = (whereClause, scopeParams) => `
      WITH em_agg AS (
        SELECT em.branch_name,
               MAX(em.region_name) AS region,
               MAX(em.division_name) AS division,
               MAX(em.area_name) AS area,
               COUNT(*) FILTER (WHERE em.status='Working')::int AS employee_count
          FROM employee_master em
         WHERE ${whereClause}
           ${_scopeWhere(session, 'em.branch_name', scopeParams)}
         GROUP BY em.branch_name
      ),
      perf AS (
        SELECT b.branch_name,
               SUM(ep.regular_demand)::bigint AS demand,
               SUM(ep.regular_collection)::bigint AS collection,
               SUM(ep.npa_cases)::int AS npa_cases
          FROM employee_performance ep
          JOIN employees e ON ep.emp_id = e.emp_id
          JOIN branches b ON e.branch_id = b.branch_id
         WHERE UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM em_agg)
         GROUP BY b.branch_name
      )
      SELECT em_agg.branch_name, em_agg.region, em_agg.division, em_agg.area,
             em_agg.employee_count,
             COALESCE(perf.demand, 0) AS demand,
             COALESCE(perf.collection, 0) AS collection,
             COALESCE(perf.npa_cases, 0) AS npa_cases
        FROM em_agg
        LEFT JOIN perf ON UPPER(em_agg.branch_name) = UPPER(perf.branch_name)
       ORDER BY em_agg.branch_name
       LIMIT ${lim}`;

    // Pass 1: substring ILIKE — handles exact and partial matches.
    let params = [`%${q}%`];
    let rows = await safeQuery(buildSql('em.branch_name ILIKE $1', params), params, lim);

    // Pass 2: trigram similarity fallback — handles voice typos
    // ("Davangere" → "Davanagere", "Belgaum" → "Belagavi"). Threshold 0.5
    // — looser was returning bogus matches (e.g. "Bangalore" → "Bangarpet").
    // Skip the fallback for very short queries where typos can match anything.
    if (Array.isArray(rows) && rows.length === 0 && q.length >= 4) {
      params = [q];
      rows = await safeQuery(buildSql(`similarity(em.branch_name, $1) > 0.5`, params), params, lim);
    }

    if (!Array.isArray(rows)) return rows;
    if (rows.length > 1) {
      return {
        ambiguous: true,
        count: rows.length,
        instruction: 'Multiple branches match this lookup (some via fuzzy spelling match — voice transcription often mishears branch names). List the matching branches with region, division, area, and employee_count, then ask which branch the user means. Do not answer as if only one branch matched.',
        matches: rows
      };
    }
    return rows;
  }

  if (name === 'period_performance') {
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const groupBy = ['day', 'month', 'branch', 'employee'].includes(args.group_by) ? args.group_by : 'month';
    const params = [start, end];
    let extra = '';
    if (args.branch_name) { params.push(args.branch_name); extra += ` AND b.branch_name ILIKE $${params.length}`; }
    if (args.area_name) {
      params.push(args.area_name);
      extra += ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${params.length})`;
    }
    if (args.division_name) {
      params.push(args.division_name);
      extra += ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${params.length})`;
    }
    if (args.region_name) {
      params.push(args.region_name);
      extra += ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${params.length})`;
    }
    const scopeClause = _scopeWhere(session, 'b.branch_name', params);
    let selectCols, groupCols, joinEm = '';
    if (groupBy === 'day') {
      selectCols = `to_char(dp.report_date, 'YYYY-MM-DD') AS bucket`;
      groupCols = `dp.report_date`;
    } else if (groupBy === 'month') {
      selectCols = `to_char(date_trunc('month', dp.report_date), 'YYYY-MM') AS bucket`;
      groupCols = `date_trunc('month', dp.report_date)`;
    } else if (groupBy === 'branch') {
      selectCols = `b.branch_name AS bucket`;
      groupCols = `b.branch_name`;
    } else {
      selectCols = `em.full_name AS bucket, dp.emp_id`;
      groupCols = `em.full_name, dp.emp_id`;
      joinEm = `LEFT JOIN employee_master em ON em.emp_id = dp.emp_id`;
    }
    const sql = `
      SELECT ${selectCols},
             SUM(dp.regular_demand)::bigint AS demand,
             SUM(dp.regular_collection)::bigint AS collection,
             SUM(dp.npa_cases)::int AS npa_cases,
             SUM(dp.npa_act_amt)::numeric AS npa_amount
        FROM daily_performance dp
        JOIN employees e ON dp.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
        ${joinEm}
       WHERE dp.report_date BETWEEN $1 AND $2
         ${extra}
         ${scopeClause}
       GROUP BY ${groupCols}
       ORDER BY 1
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
  }

  if (name === 'top_performers') {
    const metric = ['collection', 'demand', 'npa_cases'].includes(args.metric) ? args.metric : 'collection';
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const params = [start, end];
    let extra = '';
    if (args.branch_name)   { params.push(args.branch_name);   extra += ` AND em.branch_name ILIKE $${params.length}`; }
    if (args.area_name)     { params.push(args.area_name);     extra += ` AND em.area_name ILIKE $${params.length}`; }
    if (args.division_name) { params.push(args.division_name); extra += ` AND em.division_name ILIKE $${params.length}`; }
    if (args.region_name)   { params.push(args.region_name);   extra += ` AND em.region_name ILIKE $${params.length}`; }
    if (args.role)          { params.push(args.role);          extra += ` AND em.role ILIKE $${params.length}`; }
    const scopeClause = _scopeWhere(session, 'em.branch_name', params);
    const orderCol = metric === 'collection' ? 'SUM(dp.regular_collection)' :
                     metric === 'demand'     ? 'SUM(dp.regular_demand)' :
                                               'SUM(dp.npa_cases)';
    const sql = `
      SELECT em.emp_id, em.full_name AS name, em.role,
             em.branch_name AS branch, em.area_name AS area,
             em.region_name AS region,
             SUM(dp.regular_demand)::bigint AS demand,
             SUM(dp.regular_collection)::bigint AS collection,
             SUM(dp.npa_cases)::int AS npa_cases
        FROM daily_performance dp
        JOIN employee_master em ON em.emp_id = dp.emp_id
       WHERE dp.report_date BETWEEN $1 AND $2
         ${extra}
         ${scopeClause}
       GROUP BY em.emp_id, em.full_name, em.role, em.branch_name, em.area_name, em.region_name
       ORDER BY ${orderCol} DESC NULLS LAST
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
  }

  if (name === 'disbursement_query') {
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const groupBy = ['month', 'branch', 'product', 'employee', 'day'].includes(args.group_by) ? args.group_by : 'month';
    const params = [start, end];
    let extra = '';
    if (args.branch_name) { params.push(args.branch_name); extra += ` AND d.branch_name ILIKE $${params.length}`; }
    if (args.region_name) { params.push(args.region_name); extra += ` AND d.region_name ILIKE $${params.length}`; }
    const scopeClause = _scopeWhere(session, 'd.branch_name', params);
    let selectCols, groupCols, orderCol;
    if (groupBy === 'month') {
      selectCols = `d.bucket_month AS bucket`;
      groupCols = `d.bucket_month`;
      orderCol = `d.bucket_month`;
    } else if (groupBy === 'day') {
      selectCols = `to_char(d.bucket_day, 'YYYY-MM-DD') AS bucket`;
      groupCols = `d.bucket_day`;
      orderCol = `d.bucket_day`;
    } else if (groupBy === 'branch') {
      selectCols = `d.branch_name AS bucket`;
      groupCols = `d.branch_name`;
      orderCol = `d.branch_name`;
    } else if (groupBy === 'product') {
      selectCols = `d.product_name AS bucket`;
      groupCols = `d.product_name`;
      orderCol = `d.product_name`;
    } else {
      selectCols = `d.officer_name AS bucket, d.emp_id`;
      groupCols = `d.officer_name, d.emp_id`;
      orderCol = `d.officer_name`;
    }
    // Daily-precise CTE: prefer disbursement_daily for any day in the
    // requested range (true day-level totals), then UNION the monthly
    // rollup ONLY for months that have NO daily coverage in the range.
    // The previous DISB_CTE collapsed everything to 'Mon-YY' which made
    // a single-day filter return the whole month's totals — broken
    // semantics for "April 7th vs April 8th" style questions.
    const dailyAwareCte = `
      WITH d AS (
        SELECT disb_date AS bucket_day,
               to_char(disb_date,'Mon-YY') AS bucket_month,
               COALESCE(region_name,'') AS region_name,
               COALESCE(district_name,'') AS district_name,
               COALESCE(branch_name,'') AS branch_name,
               COALESCE(emp_id,'') AS emp_id,
               COALESCE(officer_name,'') AS officer_name,
               product_name,
               disb_count, disb_amount
          FROM disbursement_daily
         WHERE disb_date BETWEEN $1::date AND $2::date
        UNION ALL
        SELECT NULL::date AS bucket_day,
               db_month AS bucket_month,
               region_name, district_name, branch_name, emp_id,
               COALESCE(officer_name,'') AS officer_name,
               product_name,
               disb_count, disb_amount
          FROM disbursement
         WHERE to_date(db_month,'Mon-YY')
               BETWEEN date_trunc('month', $1::date)::date
                   AND date_trunc('month', $2::date)::date
           AND db_month NOT IN (
               SELECT DISTINCT to_char(disb_date,'Mon-YY')
                 FROM disbursement_daily
                WHERE disb_date BETWEEN $1::date AND $2::date
             )
      )`;
    const sql = `${dailyAwareCte}
      SELECT ${selectCols},
             SUM(d.disb_count)::int AS count,
             SUM(d.disb_amount)::numeric AS amount
        FROM d
       WHERE 1=1
         ${extra}
         ${scopeClause}
       GROUP BY ${groupCols}
       ORDER BY ${orderCol}
       LIMIT ${lim}`;
    const rows = await safeQuery(sql, params, lim);
    // Pre-format every money cell so the AI never has to do the
    // rupees → crore division itself. v4-pro and other models have
    // dropped zeros from large amounts (e.g. ₹178.96 Cr → ₹17.90 Cr)
    // when forced to compute. Always quote `amount_str` verbatim.
    if (Array.isArray(rows)) {
      let totalRupees = 0;
      let totalCount = 0;
      for (const r of rows) {
        if (r && r.amount != null) {
          const rupees = Number(r.amount);
          const cr = rupees / 1e7;
          r.amount_rupees = rupees;
          r.amount_cr = Math.round(cr * 100) / 100;
          r.amount_str = '₹' + r.amount_cr.toFixed(2) + ' Cr';
          totalRupees += rupees;
        }
        if (r && r.count != null) totalCount += Number(r.count);
      }
      const totalCr = Math.round((totalRupees / 1e7) * 100) / 100;
      return {
        rows,
        totals: {
          count: totalCount,
          amount_rupees: totalRupees,
          amount_cr: totalCr,
          amount_str: '₹' + totalCr.toFixed(2) + ' Cr',
        },
        formatting_note: 'Each row already has amount_str (formatted ₹X.XX Cr). Quote amount_str EXACTLY — do not divide amount yourself.',
      };
    }
    return rows;
  }

  if (name === 'list_employees') {
    const params = [];
    let where = `1=1`;
    const status = ['Working', 'Resigned', 'all'].includes(args.status) ? args.status : 'Working';
    if (status !== 'all') { params.push(status); where += ` AND status = $${params.length}`; }
    if (args.branch_name)   { params.push(args.branch_name);   where += ` AND branch_name ILIKE $${params.length}`; }
    if (args.area_name)     { params.push(args.area_name);     where += ` AND area_name ILIKE $${params.length}`; }
    if (args.division_name) { params.push(args.division_name); where += ` AND division_name ILIKE $${params.length}`; }
    if (args.region_name)   { params.push(args.region_name);   where += ` AND region_name ILIKE $${params.length}`; }
    if (args.role)          { params.push(args.role);          where += ` AND role ILIKE $${params.length}`; }
    const scopeClause = _scopeWhere(session, 'branch_name', params);
    const eLim = Math.min(Math.max(parseInt(args.limit, 10) || 100, 1), 500);
    const sql = `
      SELECT emp_id, full_name AS name, role, designation,
             branch_name AS branch, area_name AS area,
             division_name AS division, region_name AS region,
             mobile, status
        FROM employee_master
       WHERE ${where}
         ${scopeClause}
       ORDER BY role, full_name
       LIMIT ${eLim}`;
    return await safeQuery(sql, params, eLim);
  }

  if (name === 'headcount') {
    const groupBy = ['role', 'region', 'division', 'area', 'branch'].includes(args.group_by) ? args.group_by : null;
    const params = [];
    let where = `status = 'Working'`;
    if (args.role) { params.push(args.role); where += ` AND role ILIKE $${params.length}`; }
    if (args.region_name) { params.push(args.region_name); where += ` AND region_name ILIKE $${params.length}`; }
    if (args.branch_name) { params.push(args.branch_name); where += ` AND branch_name ILIKE $${params.length}`; }
    const scopeClause = _scopeWhere(session, 'branch_name', params);
    if (groupBy) {
      const col = groupBy === 'role' ? 'role'
        : groupBy === 'region' ? 'region_name'
        : groupBy === 'division' ? 'division_name'
        : groupBy === 'area' ? 'area_name'
        : 'branch_name';
      const sql = `
        SELECT ${col} AS bucket, COUNT(*)::int AS count
          FROM employee_master
         WHERE ${where}
           ${scopeClause}
           AND ${col} IS NOT NULL AND ${col} <> ''
         GROUP BY ${col}
         ORDER BY count DESC, ${col}
         LIMIT ${Math.max(lim, 50)}`;
      return await safeQuery(sql, params, Math.max(lim, 50));
    }
    const sql = `
      SELECT COUNT(*)::int AS total
        FROM employee_master
       WHERE ${where}
         ${scopeClause}`;
    const r = await safeQuery(sql, params, 1);
    if (Array.isArray(r) && r[0]) return { total: r[0].total };
    return r;
  }

  if (name === 'list_hierarchy') {
    const level = ['region', 'division', 'area', 'branch'].includes(args.level) ? args.level : 'branch';
    const parentLevel = ['region', 'division', 'area'].includes(args.parent_level) ? args.parent_level : null;
    const parentName = args.parent_name;
    const params = [];
    let where = `status = 'Working'`;
    if (parentLevel && parentName) {
      params.push(parentName);
      where += ` AND ${parentLevel}_name ILIKE $${params.length}`;
    }
    const scopeClause = _scopeWhere(session, 'branch_name', params);
    const col = level + '_name';
    // Hierarchy is finite (~129 branches, ~8 regions). Bump default limit to
    // 500 so callers don't accidentally truncate. The shared `lim` param
    // defaults to 25 which was way too small here.
    const hLim = Math.max(parseInt(args.limit, 10) || 500, 50);
    const sql = `
      SELECT ${col} AS name,
             COUNT(*)::int AS employee_count,
             COUNT(DISTINCT branch_name)::int AS branch_count
        FROM employee_master
       WHERE ${where}
         ${scopeClause}
         AND ${col} IS NOT NULL AND ${col} <> ''
       GROUP BY ${col}
       ORDER BY ${col}
       LIMIT ${hLim}`;
    return await safeQuery(sql, params, hLim);
  }

  if (name === 'daily_reports_query') {
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const tableArg = ['plan', 'achievements', 'both'].includes(args.table) ? args.table : 'both';
    const tables = tableArg === 'both'
      ? [['daily_reports', 'plan'], ['daily_reports_achievements', 'achievements']]
      : tableArg === 'plan'
        ? [['daily_reports', 'plan']]
        : [['daily_reports_achievements', 'achievements']];
    const allCols = ['ftod_actual','ftod_plan','dpd_1_30_actual','dpd_1_30_plan','dpd_31_60_actual','dpd_31_60_plan','dpd_61_90_actual','dpd_61_90_plan','npa_activation','npa_closure','fy_non_start_acc','fy_non_start_plan','disb_igl_acc','disb_igl_amt','disb_fig_acc','disb_fig_amt','disb_il_acc','disb_il_amt','kyc_igl','kyc_fig','kyc_il'];
    const metricsArg = String(args.metrics || 'all').toLowerCase();
    const wanted = metricsArg === 'all' ? null : metricsArg.split(',').map((s) => s.trim()).filter(Boolean);
    const projectCols = wanted
      ? allCols.filter((c) => wanted.some((w) =>
          (w === 'ftod' && c.startsWith('ftod_')) ||
          (w === 'dpd_1_30' && c.startsWith('dpd_1_30_')) ||
          (w === 'dpd_31_60' && c.startsWith('dpd_31_60_')) ||
          (w === 'dpd_61_90' && c.startsWith('dpd_61_90_')) ||
          (w === 'npa' && (c === 'npa_activation' || c === 'npa_closure')) ||
          (w === 'disb' && c.startsWith('disb_')) ||
          (w === 'kyc' && c.startsWith('kyc_')) ||
          (w === 'fy_non_start' && c.startsWith('fy_non_start_'))
        ))
      : allCols;
    if (!projectCols.length) projectCols.push(...allCols);
    const out = [];
    for (const [t, label] of tables) {
      const params = [start, end];
      let extra = '';
      if (args.branch_name) { params.push(args.branch_name); extra += ` AND branch_name ILIKE $${params.length}`; }
      // Don't filter on dr.region / dr.district — those columns drift in
      // value (e.g. KALABURAGI vs KALBURGI in same table). Resolve the
      // requested region/district to a canonical branch list via
      // employee_master, which is the source of truth, then match by branch.
      if (args.region_name) {
        params.push(args.region_name);
        extra += ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${params.length})`;
      }
      if (args.district_name) {
        params.push(args.district_name);
        extra += ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${params.length} OR area_name ILIKE $${params.length})`;
      }
      const scopeClause = _scopeWhere(session, 'branch_name', params);
      // Daily-reports tables can have one row per branch per day. With ~129
      // branches and a multi-week window, the default 25-row limit was
      // silently truncating. Bump to 500.
      const drLim = Math.min(Math.max(parseInt(args.limit, 10) || 500, 1), 1000);
      const sql = `SELECT date, branch_name, region, district, dm_name, ${projectCols.join(', ')}
                     FROM ${t}
                    WHERE date BETWEEN $1 AND $2
                      ${extra}
                      ${scopeClause}
                    ORDER BY date, branch_name
                    LIMIT ${drLim}`;
      const rows = await safeQuery(sql, params, drLim);
      if (Array.isArray(rows)) out.push({ source: label, count: rows.length, rows });
      else out.push({ source: label, error: rows.error || 'unknown' });
    }
    return out;
  }

  if (name === 'npa_summary') {
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const groupBy = ['month', 'branch', 'employee'].includes(args.group_by) ? args.group_by : 'month';
    const params = [start, end];
    let extra = '';
    if (args.branch_name) { params.push(args.branch_name); extra += ` AND b.branch_name ILIKE $${params.length}`; }
    if (args.region_name) {
      params.push(args.region_name);
      extra += ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${params.length})`;
    }
    const scopeClause = _scopeWhere(session, 'b.branch_name', params);
    let selectCols, groupCols, joinEm = '';
    if (groupBy === 'month') {
      selectCols = `to_char(date_trunc('month', dp.report_date), 'YYYY-MM') AS bucket`;
      groupCols = `date_trunc('month', dp.report_date)`;
    } else if (groupBy === 'branch') {
      selectCols = `b.branch_name AS bucket`;
      groupCols = `b.branch_name`;
    } else {
      selectCols = `em.full_name AS bucket, dp.emp_id`;
      groupCols = `em.full_name, dp.emp_id`;
      joinEm = `LEFT JOIN employee_master em ON em.emp_id = dp.emp_id`;
    }
    const sql = `
      SELECT ${selectCols},
             SUM(dp.npa_cases)::int AS npa_cases,
             SUM(dp.npa_act_acc)::int AS npa_activated,
             SUM(dp.npa_act_amt)::numeric AS npa_amount
        FROM daily_performance dp
        JOIN employees e ON dp.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
        ${joinEm}
       WHERE dp.report_date BETWEEN $1 AND $2
         ${extra}
         ${scopeClause}
       GROUP BY ${groupCols}
       ORDER BY 1
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
  }

  if (name === 'branch_summary') {
    // Resolve target branch. Branch-bound roles fall back to session.location.
    let branchName = String(args.branch_name || '').trim();
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    const branchBound = role === 'BM' || role === 'ABM' || role === 'BOE';
    if (!branchName) {
      if (branchBound && loc) branchName = loc;
      else return { error: 'branch_name required', message: 'branch_summary needs a single branch. Pass branch_name (e.g. "Davanagere"). For region-wide summaries use period_performance(group_by="branch", ...).' };
    }

    // Resolve as-of date: explicit arg, else latest report_date with data for this branch.
    let asOfDate = null;
    if (args.date && /^\d{4}-\d{2}-\d{2}$/.test(String(args.date).trim())) {
      asOfDate = String(args.date).trim();
    }
    if (!asOfDate) {
      const r = await safeQuery(
        `SELECT MAX(dp.report_date) AS d
           FROM daily_performance dp
           JOIN employees e ON dp.emp_id = e.emp_id
           JOIN branches b ON e.branch_id = b.branch_id
          WHERE b.branch_name ILIKE $1`,
        [branchName], 1
      );
      if (Array.isArray(r) && r[0] && r[0].d) {
        const d = r[0].d;
        asOfDate = (d instanceof Date) ? d.toISOString().slice(0, 10) : String(d).slice(0, 10);
      } else {
        asOfDate = new Date().toISOString().slice(0, 10);
      }
    }

    // Month + FY anchors derived from asOfDate (FY = Apr 1 → Mar 31).
    const [yy, mn] = asOfDate.split('-');
    const monthStart = `${yy}-${mn}-01`;
    const fyYear = parseInt(mn, 10) >= 4 ? parseInt(yy, 10) : parseInt(yy, 10) - 1;
    const fyStart = `${fyYear}-04-01`;

    // 1. Headcount + role breakdown (Working only).
    const hcRows = await safeQuery(
      `SELECT role, COUNT(*)::int AS count
         FROM employee_master
        WHERE branch_name ILIKE $1
          AND status = 'Working'
          AND role IS NOT NULL AND role <> ''
        GROUP BY role
        ORDER BY count DESC, role`,
      [branchName], 50
    );
    const byRole = Array.isArray(hcRows) ? hcRows : [];
    const totalHc = byRole.reduce((s, r) => s + Number(r.count || 0), 0);

    // 2. Collection — today / MTD / FYTD (counts) + NPA (MTD aggregates).
    //    daily_performance.regular_demand / regular_collection are COUNT columns.
    const collRows = await safeQuery(
      `SELECT
         SUM(CASE WHEN dp.report_date = $2 THEN dp.regular_demand ELSE 0 END)::bigint     AS today_demand,
         SUM(CASE WHEN dp.report_date = $2 THEN dp.regular_collection ELSE 0 END)::bigint AS today_collection,
         SUM(CASE WHEN dp.report_date BETWEEN $3 AND $2 THEN dp.regular_demand ELSE 0 END)::bigint     AS mtd_demand,
         SUM(CASE WHEN dp.report_date BETWEEN $3 AND $2 THEN dp.regular_collection ELSE 0 END)::bigint AS mtd_collection,
         SUM(CASE WHEN dp.report_date BETWEEN $4 AND $2 THEN dp.regular_demand ELSE 0 END)::bigint     AS fytd_demand,
         SUM(CASE WHEN dp.report_date BETWEEN $4 AND $2 THEN dp.regular_collection ELSE 0 END)::bigint AS fytd_collection,
         SUM(CASE WHEN dp.report_date = $2 THEN dp.npa_cases ELSE 0 END)::int        AS npa_cases,
         SUM(CASE WHEN dp.report_date = $2 THEN dp.npa_act_amt ELSE 0 END)::numeric  AS npa_act_amount,
         SUM(CASE WHEN dp.report_date = $2 THEN dp.npa_clo_amt ELSE 0 END)::numeric  AS npa_clo_amount
       FROM daily_performance dp
       JOIN employees e ON dp.emp_id = e.emp_id
       JOIN branches b ON e.branch_id = b.branch_id
      WHERE b.branch_name ILIKE $1
        AND dp.report_date BETWEEN $4 AND $2`,
      [branchName, asOfDate, monthStart, fyStart], 1
    );
    const c = (Array.isArray(collRows) && collRows[0]) || {};
    const pct = (col, dem) => {
      const cn = Number(col || 0);
      const dn = Number(dem || 0);
      return dn > 0 ? Math.round((cn / dn) * 1000) / 10 : null;
    };

    // 3. Disbursement — today + MTD from disbursement_daily.
    const disbRows = await safeQuery(
      `SELECT
         COALESCE(SUM(CASE WHEN disb_date = $2 THEN disb_count  ELSE 0 END), 0)::int     AS today_count,
         COALESCE(SUM(CASE WHEN disb_date = $2 THEN disb_amount ELSE 0 END), 0)::numeric AS today_amount,
         COALESCE(SUM(CASE WHEN disb_date BETWEEN $3 AND $2 THEN disb_count  ELSE 0 END), 0)::int     AS mtd_count,
         COALESCE(SUM(CASE WHEN disb_date BETWEEN $3 AND $2 THEN disb_amount ELSE 0 END), 0)::numeric AS mtd_amount
       FROM disbursement_daily
      WHERE branch_name ILIKE $1
        AND disb_date BETWEEN $3 AND $2`,
      [branchName, asOfDate, monthStart], 1
    );
    const dd = (Array.isArray(disbRows) && disbRows[0]) || {};

    // 4. FTOD + DPD bucket plan-vs-actual from daily_reports_achievements
    //    (preferred — has both *_actual and *_plan). Fallback to daily_reports
    //    so the BM's plan still surfaces before achievement is filed.
    const drCols = `ftod_actual, ftod_plan,
                    dpd_1_30_actual, dpd_1_30_plan,
                    dpd_31_60_actual, dpd_31_60_plan`;
    let drRows = await safeQuery(
      `SELECT ${drCols} FROM daily_reports_achievements
        WHERE branch_name ILIKE $1 AND date = $2 LIMIT 1`,
      [branchName, asOfDate], 1
    );
    if (!Array.isArray(drRows) || drRows.length === 0) {
      drRows = await safeQuery(
        `SELECT ${drCols} FROM daily_reports
          WHERE branch_name ILIKE $1 AND date = $2 LIMIT 1`,
        [branchName, asOfDate], 1
      );
    }
    const dr = (Array.isArray(drRows) && drRows[0]) || {};
    const numOrNull = (v) => (v === null || v === undefined) ? null : Number(v);
    const ftodActual = numOrNull(dr.ftod_actual);
    const ftodPlan = numOrNull(dr.ftod_plan);

    return {
      branch_name: branchName,
      date: asOfDate,
      headcount: { total: totalHc, by_role: byRole },
      collection_today: {
        demand: Number(c.today_demand || 0),
        collection: Number(c.today_collection || 0),
        pct: pct(c.today_collection, c.today_demand)
      },
      collection_mtd: {
        demand: Number(c.mtd_demand || 0),
        collection: Number(c.mtd_collection || 0),
        pct: pct(c.mtd_collection, c.mtd_demand)
      },
      collection_fytd: {
        demand: Number(c.fytd_demand || 0),
        collection: Number(c.fytd_collection || 0),
        pct: pct(c.fytd_collection, c.fytd_demand)
      },
      npa: {
        cases: Number(c.npa_cases || 0),
        act_amount: Number(c.npa_act_amount || 0),
        closure_amount: Number(c.npa_clo_amount || 0)
      },
      disbursement_today: {
        count: Number(dd.today_count || 0),
        amount: Number(dd.today_amount || 0)
      },
      disbursement_mtd: {
        count: Number(dd.mtd_count || 0),
        amount: Number(dd.mtd_amount || 0)
      },
      ftod_today: {
        actual: ftodActual,
        plan: ftodPlan,
        gap: (ftodActual !== null && ftodPlan !== null) ? (ftodActual - ftodPlan) : null
      },
      dpd_today: {
        dpd_1_30: { actual: numOrNull(dr.dpd_1_30_actual), plan: numOrNull(dr.dpd_1_30_plan) },
        dpd_31_60: { actual: numOrNull(dr.dpd_31_60_actual), plan: numOrNull(dr.dpd_31_60_plan) }
      }
    };
  }

  if (name === 'hourly_collection') {
    // hourly_performance is a SINGLE intra-day snapshot — no date column,
    // no hour column. UPSERTed by intra-day uploads, reset after EOD upload.
    // Time-series ("group by hour with cumulative + delta") is not possible
    // against this schema today. We return the CURRENT live aggregate.
    const groupByArg = ['branch', 'region', 'employee', 'none'].includes(args.group_by) ? args.group_by : null;
    const role = String((session && session.role) || '').toUpperCase();
    const branchBound = role === 'BM' || role === 'ABM' || role === 'BOE';
    const params = [];
    let extra = '';
    if (args.branch_name) { params.push(args.branch_name); extra += ` AND b.branch_name ILIKE $${params.length}`; }
    if (args.region_name) {
      params.push(args.region_name);
      extra += ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${params.length})`;
    }
    const scopeClause = _scopeWhere(session, 'b.branch_name', params);
    // Default grouping: "none" if a single branch is in play (caller passed
    // branch_name OR is branch-bound); else group by branch.
    const effGroup = groupByArg || ((args.branch_name || branchBound) ? 'none' : 'branch');
    let selectCols, groupCols, orderCols, joinEm = '';
    if (effGroup === 'branch') {
      selectCols = `b.branch_name AS bucket`;
      groupCols = `b.branch_name`;
      orderCols = `SUM(hp.regular_collection) DESC NULLS LAST`;
    } else if (effGroup === 'region') {
      selectCols = `r.region_name AS bucket`;
      groupCols = `r.region_name`;
      orderCols = `SUM(hp.regular_collection) DESC NULLS LAST`;
    } else if (effGroup === 'employee') {
      selectCols = `em.full_name AS bucket, hp.emp_id, em.role, b.branch_name AS branch`;
      groupCols = `em.full_name, hp.emp_id, em.role, b.branch_name`;
      orderCols = `SUM(hp.regular_collection) DESC NULLS LAST`;
      joinEm = `LEFT JOIN employee_master em ON em.emp_id = hp.emp_id`;
    } else {
      selectCols = `'live' AS bucket`;
      groupCols = `1`; // group by literal — single-row aggregate
      orderCols = `1`;
    }
    const sql = `
      SELECT ${selectCols},
             SUM(hp.regular_demand)::bigint     AS demand_count,
             SUM(hp.regular_collection)::bigint AS collection_count,
             SUM(hp.regular_demand_amt)::numeric     AS demand_amount,
             SUM(hp.regular_collection_amt)::numeric AS collection_amount
        FROM hourly_performance hp
        JOIN employees e  ON hp.emp_id = e.emp_id
        JOIN branches b   ON e.branch_id = b.branch_id
        JOIN districts d  ON b.district_id = d.district_id
        JOIN regions r    ON d.region_id = r.region_id
        ${joinEm}
       WHERE 1=1
         ${extra}
         ${scopeClause}
       GROUP BY ${groupCols}
       ORDER BY ${orderCols}
       LIMIT ${lim}`;
    const rows = await safeQuery(sql, params, lim);
    if (!Array.isArray(rows)) return rows;
    const decorated = rows.map(function (r) {
      const dem = Number(r.demand_count || 0);
      const col = Number(r.collection_count || 0);
      const p = dem > 0 ? Math.round((col / dem) * 1000) / 10 : null;
      const out = {
        bucket: r.bucket,
        demand_count: dem,
        collection_count: col,
        demand_amount: Number(r.demand_amount || 0),
        collection_amount: Number(r.collection_amount || 0),
        pct: p
      };
      if (effGroup === 'employee') {
        out.emp_id = r.emp_id;
        out.role = r.role;
        out.branch = r.branch;
      }
      return out;
    });
    // as_of: hourly_performance has no upload_time column, so this is the
    // server's view-time. Live snapshot — approximate is fine.
    return {
      snapshot: 'live',
      as_of: new Date().toISOString(),
      note: 'CURRENT live snapshot — single point in time, NOT a historical series. hourly_performance is UPSERTed on each intra-day upload and reset after EOD. For day-by-day trends use period_performance(group_by="day").',
      grouped_by: effGroup,
      rows: decorated
    };
  }

  if (name === 'collection_drilldown') {
    let branchName = String(args.branch_name || '').trim();
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    const branchBound = role === 'BM' || role === 'ABM' || role === 'BOE';
    if (!branchName) {
      if (branchBound && loc) branchName = loc;
      else return { error: 'branch_name required', message: 'collection_drilldown is a single-branch tool. Pass branch_name (e.g. "Davanagere").' };
    }

    // Resolve as-of date — default to MAX(report_date) for this branch.
    let asOfDate = null;
    if (args.date && /^\d{4}-\d{2}-\d{2}$/.test(String(args.date).trim())) {
      asOfDate = String(args.date).trim();
    }
    if (!asOfDate) {
      const r = await safeQuery(
        `SELECT MAX(dp.report_date) AS d
           FROM daily_performance dp
           JOIN employees e ON dp.emp_id = e.emp_id
           JOIN branches b ON e.branch_id = b.branch_id
          WHERE b.branch_name ILIKE $1`,
        [branchName], 1
      );
      const d = (Array.isArray(r) && r[0] && r[0].d) || null;
      asOfDate = d
        ? ((d instanceof Date) ? d.toISOString().slice(0, 10) : String(d).slice(0, 10))
        : new Date().toISOString().slice(0, 10);
    }

    // 1. Per-FO ranking on asOfDate (count cols).
    //    top_3_underperformers = sorted ASC by pct (the 3 worst).
    //    bottom_3_underperformers = sorted DESC by pct (the 3 best — naming
    //    preserved per spec; the field name is intentionally awkward).
    const empSql = `
      SELECT em.emp_id, em.full_name AS name, em.role,
             SUM(dp.regular_demand)::bigint     AS demand,
             SUM(dp.regular_collection)::bigint AS collection
        FROM daily_performance dp
        JOIN employees e ON dp.emp_id = e.emp_id
        JOIN branches b  ON e.branch_id = b.branch_id
        JOIN employee_master em ON em.emp_id = dp.emp_id
       WHERE b.branch_name ILIKE $1
         AND dp.report_date = $2
         AND em.role = 'FO'
       GROUP BY em.emp_id, em.full_name, em.role
      HAVING SUM(dp.regular_demand) > 0
       ORDER BY (SUM(dp.regular_collection)::float / NULLIF(SUM(dp.regular_demand), 0)) ASC NULLS LAST
       LIMIT 200`;
    const empRows = await safeQuery(empSql, [branchName, asOfDate], 200);
    const decorated = (Array.isArray(empRows) ? empRows : []).map(function (r) {
      const dem = Number(r.demand || 0);
      const col = Number(r.collection || 0);
      return {
        emp_id: r.emp_id,
        name: r.name,
        role: r.role,
        demand: dem,
        collection: col,
        pct: dem > 0 ? Math.round((col / dem) * 1000) / 10 : null
      };
    });
    const top3 = decorated.slice(0, 3);                       // 3 worst (lowest pct)
    const bottom3 = decorated.slice().reverse().slice(0, 3);  // 3 best (highest pct)

    // 2. DPD buckets + NPA + FTOD from daily_reports_achievements,
    //    fallback to daily_reports if achievement not yet filed.
    const drCols = `ftod_actual, ftod_plan,
                    dpd_1_30_actual, dpd_1_30_plan,
                    dpd_31_60_actual, dpd_31_60_plan,
                    dpd_61_90_actual, dpd_61_90_plan,
                    npa_activation, npa_closure`;
    let drRows = await safeQuery(
      `SELECT ${drCols} FROM daily_reports_achievements
        WHERE branch_name ILIKE $1 AND date = $2 LIMIT 1`,
      [branchName, asOfDate], 1
    );
    if (!Array.isArray(drRows) || drRows.length === 0) {
      drRows = await safeQuery(
        `SELECT ${drCols} FROM daily_reports
          WHERE branch_name ILIKE $1 AND date = $2 LIMIT 1`,
        [branchName, asOfDate], 1
      );
    }
    const dr = (Array.isArray(drRows) && drRows[0]) || {};
    const numOrNull = (v) => (v === null || v === undefined) ? null : Number(v);
    const ftodActual = numOrNull(dr.ftod_actual);
    const ftodPlan = numOrNull(dr.ftod_plan);

    // 3. NPA today — outstanding case count + activation amount on asOfDate.
    //    Use the single-day snapshot to avoid the rolling-snapshot inflation
    //    that summing over a range would cause.
    const npaRows = await safeQuery(
      `SELECT COALESCE(SUM(dp.npa_cases), 0)::int AS cases,
              COALESCE(SUM(dp.npa_act_amt), 0)::numeric AS amount
         FROM daily_performance dp
         JOIN employees e ON dp.emp_id = e.emp_id
         JOIN branches b ON e.branch_id = b.branch_id
        WHERE b.branch_name ILIKE $1
          AND dp.report_date = $2`,
      [branchName, asOfDate], 1
    );
    const np = (Array.isArray(npaRows) && npaRows[0]) || {};

    return {
      branch_name: branchName,
      date: asOfDate,
      top_3_underperformers: top3,
      bottom_3_underperformers: bottom3,
      dpd_buckets: {
        dpd_1_30:  { actual: numOrNull(dr.dpd_1_30_actual),  plan: numOrNull(dr.dpd_1_30_plan)  },
        dpd_31_60: { actual: numOrNull(dr.dpd_31_60_actual), plan: numOrNull(dr.dpd_31_60_plan) },
        dpd_61_90: { actual: numOrNull(dr.dpd_61_90_actual), plan: numOrNull(dr.dpd_61_90_plan) }
      },
      npa_today: {
        cases: Number(np.cases || 0),
        amount: Number(np.amount || 0)
      },
      ftod_gap_today: {
        actual: ftodActual,
        plan: ftodPlan,
        gap: (ftodActual !== null && ftodPlan !== null) ? (ftodActual - ftodPlan) : null
      }
    };
  }

  if (name === 'period_compare') {
    const allowedMetrics = ['collection', 'demand', 'collection_pct', 'npa_amount', 'disb_amount', 'disb_count', 'ftod'];
    const metric = allowedMetrics.includes(args.metric) ? args.metric : null;
    if (!metric) return { error: 'metric required', message: 'metric must be one of: ' + allowedMetrics.join(', ') };
    const allowedScopes = ['all', 'branch', 'region', 'division', 'area'];
    const scope = allowedScopes.includes(args.scope) ? args.scope : 'all';
    const scopeValue = String(args.scope_value || '').trim();
    if (scope !== 'all' && !scopeValue) return { error: 'scope_value required', message: `scope_value required when scope="${scope}".` };
    const aS = args.period_a_start, aE = args.period_a_end, bS = args.period_b_start, bE = args.period_b_end;
    if (!aS || !aE || !bS || !bE) return { error: 'period_a_start, period_a_end, period_b_start, period_b_end all required' };
    const isoOk = /^\d{4}-\d{2}-\d{2}$/;
    if (!isoOk.test(aS) || !isoOk.test(aE) || !isoOk.test(bS) || !isoOk.test(bE)) {
      return { error: 'invalid date format', message: 'All dates must be YYYY-MM-DD.' };
    }

    // BM/ABM/BOE may only compare their own branch.
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    const branchBound = role === 'BM' || role === 'ABM' || role === 'BOE';
    if (branchBound) {
      if (scope !== 'branch' || scopeValue.toLowerCase() !== loc.toLowerCase()) {
        return { error: 'scope_violation', message: `${role} can only compare their own branch (${loc}). Use scope="branch", scope_value="${loc}".` };
      }
    }

    // Build the value computation for one (start, end). Returns a number.
    async function valueFor(start, end) {
      // Filter clause varies by scope.
      // We always _scopeWhere on the appropriate alias as a safety net.
      const params = [start, end];
      let scopeFilter = '';
      if (scope === 'branch') {
        params.push(scopeValue);
        scopeFilter = ` AND b.branch_name ILIKE $${params.length}`;
      } else if (scope === 'region') {
        params.push(scopeValue);
        scopeFilter = ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${params.length})`;
      } else if (scope === 'division') {
        params.push(scopeValue);
        scopeFilter = ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${params.length})`;
      } else if (scope === 'area') {
        params.push(scopeValue);
        scopeFilter = ` AND UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${params.length})`;
      }
      const sessionScope = _scopeWhere(session, 'b.branch_name', params);

      let sql, alias = 'b.branch_name';
      if (metric === 'collection') {
        sql = `SELECT COALESCE(SUM(dp.regular_collection), 0)::numeric AS v
                 FROM daily_performance dp
                 JOIN employees e ON dp.emp_id = e.emp_id
                 JOIN branches b  ON e.branch_id = b.branch_id
                WHERE dp.report_date BETWEEN $1 AND $2 ${scopeFilter} ${sessionScope}`;
      } else if (metric === 'demand') {
        sql = `SELECT COALESCE(SUM(dp.regular_demand), 0)::numeric AS v
                 FROM daily_performance dp
                 JOIN employees e ON dp.emp_id = e.emp_id
                 JOIN branches b  ON e.branch_id = b.branch_id
                WHERE dp.report_date BETWEEN $1 AND $2 ${scopeFilter} ${sessionScope}`;
      } else if (metric === 'collection_pct') {
        sql = `SELECT CASE WHEN SUM(dp.regular_demand) > 0
                           THEN ROUND(SUM(dp.regular_collection)::numeric / SUM(dp.regular_demand) * 1000) / 10
                           ELSE NULL END AS v
                 FROM daily_performance dp
                 JOIN employees e ON dp.emp_id = e.emp_id
                 JOIN branches b  ON e.branch_id = b.branch_id
                WHERE dp.report_date BETWEEN $1 AND $2 ${scopeFilter} ${sessionScope}`;
      } else if (metric === 'npa_amount') {
        sql = `SELECT COALESCE(SUM(dp.npa_act_amt), 0)::numeric AS v
                 FROM daily_performance dp
                 JOIN employees e ON dp.emp_id = e.emp_id
                 JOIN branches b  ON e.branch_id = b.branch_id
                WHERE dp.report_date BETWEEN $1 AND $2 ${scopeFilter} ${sessionScope}`;
      } else if (metric === 'disb_amount') {
        // disbursement_daily uses disb_date and a flat branch_name column.
        // Replace b.branch_name alias with d.branch_name.
        const dParams = [start, end];
        let dFilter = '';
        if (scope === 'branch') { dParams.push(scopeValue); dFilter = ` AND d.branch_name ILIKE $${dParams.length}`; }
        else if (scope === 'region') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${dParams.length})`; }
        else if (scope === 'division') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${dParams.length})`; }
        else if (scope === 'area') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${dParams.length})`; }
        const dSession = _scopeWhere(session, 'd.branch_name', dParams);
        sql = `SELECT COALESCE(SUM(d.disb_amount), 0)::numeric AS v
                 FROM disbursement_daily d
                WHERE d.disb_date BETWEEN $1 AND $2 ${dFilter} ${dSession}`;
        const r = await safeQuery(sql, dParams, 1);
        return (Array.isArray(r) && r[0] && r[0].v != null) ? Number(r[0].v) : 0;
      } else if (metric === 'disb_count') {
        const dParams = [start, end];
        let dFilter = '';
        if (scope === 'branch') { dParams.push(scopeValue); dFilter = ` AND d.branch_name ILIKE $${dParams.length}`; }
        else if (scope === 'region') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${dParams.length})`; }
        else if (scope === 'division') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${dParams.length})`; }
        else if (scope === 'area') { dParams.push(scopeValue); dFilter = ` AND UPPER(d.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${dParams.length})`; }
        const dSession = _scopeWhere(session, 'd.branch_name', dParams);
        sql = `SELECT COALESCE(SUM(d.disb_count), 0)::numeric AS v
                 FROM disbursement_daily d
                WHERE d.disb_date BETWEEN $1 AND $2 ${dFilter} ${dSession}`;
        const r = await safeQuery(sql, dParams, 1);
        return (Array.isArray(r) && r[0] && r[0].v != null) ? Number(r[0].v) : 0;
      } else if (metric === 'ftod') {
        // daily_reports_achievements has flat branch_name — no JOIN.
        const drParams = [start, end];
        let drFilter = '';
        if (scope === 'branch') { drParams.push(scopeValue); drFilter = ` AND branch_name ILIKE $${drParams.length}`; }
        else if (scope === 'region') { drParams.push(scopeValue); drFilter = ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${drParams.length})`; }
        else if (scope === 'division') { drParams.push(scopeValue); drFilter = ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${drParams.length})`; }
        else if (scope === 'area') { drParams.push(scopeValue); drFilter = ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${drParams.length})`; }
        const drSession = _scopeWhere(session, 'branch_name', drParams);
        sql = `SELECT COALESCE(SUM(ftod_actual), 0)::numeric AS v
                 FROM daily_reports_achievements
                WHERE date BETWEEN $1 AND $2 ${drFilter} ${drSession}`;
        const r = await safeQuery(sql, drParams, 1);
        return (Array.isArray(r) && r[0] && r[0].v != null) ? Number(r[0].v) : 0;
      }
      const r = await safeQuery(sql, params, 1);
      return (Array.isArray(r) && r[0] && r[0].v != null) ? Number(r[0].v) : 0;
    }

    const aVal = await valueFor(aS, aE);
    const bVal = await valueFor(bS, bE);
    const deltaAbs = Number(aVal) - Number(bVal);
    let deltaPct = null;
    if (Number(bVal) !== 0) {
      deltaPct = Math.round(((Number(aVal) - Number(bVal)) / Math.abs(Number(bVal))) * 1000) / 10;
    }
    return {
      metric,
      scope,
      scope_value: scope === 'all' ? null : scopeValue,
      period_a: { start: aS, end: aE, value: Number(aVal) },
      period_b: { start: bS, end: bE, value: Number(bVal) },
      delta_abs: deltaAbs,
      delta_pct: deltaPct
    };
  }

  if (name === 'plan_compliance') {
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    const branchBound = role === 'BM' || role === 'ABM' || role === 'BOE';
    if (branchBound) {
      return { error: 'scope_violation', message: `plan_compliance is a multi-branch tool. ${role} only has one branch (${loc}) — use branch_summary or daily_reports_query instead.` };
    }

    const date = (args.date && /^\d{4}-\d{2}-\d{2}$/.test(String(args.date).trim()))
      ? String(args.date).trim()
      : new Date().toISOString().slice(0, 10);
    const allowedScopes = ['all', 'region', 'division', 'area'];
    const scope = allowedScopes.includes(args.scope) ? args.scope : 'all';
    const scopeValue = String(args.scope_value || '').trim();
    if (scope !== 'all' && !scopeValue) return { error: 'scope_value required', message: `scope_value required when scope="${scope}".` };

    // Build the expected-branches filter (scope arg + session scope).
    const expParams = [];
    let expWhere = `status = 'Working' AND branch_name IS NOT NULL AND branch_name <> ''`;
    if (scope === 'region') { expParams.push(scopeValue); expWhere += ` AND region_name ILIKE $${expParams.length}`; }
    else if (scope === 'division') { expParams.push(scopeValue); expWhere += ` AND division_name ILIKE $${expParams.length}`; }
    else if (scope === 'area') { expParams.push(scopeValue); expWhere += ` AND area_name ILIKE $${expParams.length}`; }
    const expSession = _scopeWhere(session, 'branch_name', expParams);
    const expRows = await safeQuery(
      `SELECT DISTINCT branch_name
         FROM employee_master
        WHERE ${expWhere} ${expSession}
        ORDER BY branch_name`,
      expParams, 1000
    );
    const expected = (Array.isArray(expRows) ? expRows : [])
      .map((r) => String(r.branch_name || '').trim())
      .filter(Boolean);

    // Filed for plan + achievement on `date` — restricted to expected set.
    async function filedFor(table) {
      const fParams = [date];
      let fWhere = `date = $1`;
      if (scope === 'region') { fParams.push(scopeValue); fWhere += ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE region_name ILIKE $${fParams.length})`; }
      else if (scope === 'division') { fParams.push(scopeValue); fWhere += ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE division_name ILIKE $${fParams.length})`; }
      else if (scope === 'area') { fParams.push(scopeValue); fWhere += ` AND UPPER(branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE area_name ILIKE $${fParams.length})`; }
      const fSession = _scopeWhere(session, 'branch_name', fParams);
      const rows = await safeQuery(
        `SELECT DISTINCT branch_name FROM ${table} WHERE ${fWhere} ${fSession} ORDER BY branch_name`,
        fParams, 1000
      );
      return (Array.isArray(rows) ? rows : [])
        .map((r) => String(r.branch_name || '').trim())
        .filter(Boolean);
    }
    const filedPlan = await filedFor('daily_reports');
    const filedAch  = await filedFor('daily_reports_achievements');

    // Set diff — case-insensitive match against expected.
    const upper = (s) => String(s || '').toUpperCase();
    const planSet = new Set(filedPlan.map(upper));
    const achSet  = new Set(filedAch.map(upper));
    const missingPlan = expected.filter((b) => !planSet.has(upper(b))).sort((a, b) => a.localeCompare(b));
    const missingAch  = expected.filter((b) => !achSet.has(upper(b))).sort((a, b) => a.localeCompare(b));

    return {
      date,
      scope,
      scope_value: scope === 'all' ? null : scopeValue,
      expected_count: expected.length,
      filed_plan_count: filedPlan.length,
      filed_achievement_count: filedAch.length,
      missing_plan_count: missingPlan.length,
      missing_achievement_count: missingAch.length,
      expected_branches: expected,
      filed_plan: filedPlan,
      filed_achievement: filedAch,
      missing_plan: missingPlan,
      missing_achievement: missingAch
    };
  }

  return { error: 'unknown_tool: ' + name };
}

function emCol_for_disb(role) {
  role = (role || '').toUpperCase();
  if (role === 'RM' || role === 'SM') return 'region_name';
  if (role === 'DM' || role === 'DVM') return 'division_name';
  if (role === 'AM') return 'area_name';
  if (role === 'BM' || role === 'FO') return 'branch_name';
  return null;
}

// Endpoint: provide data context for client-side AI calls
app.post("/api/ai-context", async (req, res) => {
  try {
    const { session } = req.body;
    const ctx = await buildDataContext(session || {});
    res.json(ctx);
  } catch (e) {
    console.error("AI context error:", e);
    res.status(500).json({ error: "Failed to load data context" });
  }
});

// /api/ai-snapshot — full ctx + schema cheatsheet + version. The mobile app
// pulls this on login and caches locally so the on-device AI assistant has a
// schema map and historical/employee data even when offline.
app.post("/api/ai-snapshot", async (req, res) => {
  try {
    const body = req.body || {};
    const session = {
      role: body.role || (body.session && body.session.role) || '',
      location: body.location || (body.session && body.session.location) || '',
    };
    const ctx = await buildDataContext(session);
    // Surface summary_text + latest + truncated at the top level too — mobile
    // can pull them without traversing into ctx.* (smaller deserialise).
    res.json({
      version: AI_SNAPSHOT_VERSION,
      generated_at: ctx.generated_at || new Date().toISOString(),
      emp_id: body.emp_id || null,
      session,
      schema: AI_SCHEMA_CHEATSHEET,
      summary_text: ctx.summary_text || '',
      latest: ctx.latest || null,
      truncated: ctx.truncated || {},
      ctx,
    });
  } catch (e) {
    console.error("AI snapshot error:", e);
    res.status(500).json({ error: "Failed to build snapshot" });
  }
});

// ========== AI Chat Proxy — Google Gemini ==========
const GEMINI_KEY = process.env.GEMINI_KEY;
const GEMINI_MODELS = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-2.0-flash-lite'];
const MISTRAL_KEY = process.env.MISTRAL_KEY;
const MISTRAL_MODEL = process.env.MISTRAL_MODEL || 'mistral-large-latest';
const DEEPSEEK_KEY = process.env.DEEPSEEK_KEY;
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || 'deepseek-chat';
const crypto = require('crypto');

// ========== OpenAI voice pipeline (STT + LLM + TTS) ==========
const OPENAI_KEY = process.env.OPENAI_API_KEY;
const OPENAI_LLM = process.env.OPENAI_LLM_MODEL || 'gpt-4o';
const OPENAI_STT = process.env.OPENAI_STT_MODEL || 'gpt-4o-mini-transcribe';
const OPENAI_TTS = process.env.OPENAI_TTS_MODEL || 'gpt-4o-mini-tts';
const OPENAI_VOICE = process.env.OPENAI_TTS_VOICE || 'ash';
const DALLE_MODEL = process.env.DALLE_MODEL || 'dall-e-3';
const DALLE_SIZE = process.env.DALLE_SIZE || '1024x1024';
const DALLE_QUALITY = process.env.DALLE_QUALITY || 'standard'; // 'standard' | 'hd'
const DALLE_COST = DALLE_SIZE === '1024x1792' || DALLE_SIZE === '1792x1024'
  ? (DALLE_QUALITY === 'hd' ? 0.120 : 0.080)
  : (DALLE_QUALITY === 'hd' ? 0.080 : 0.040);

const voiceUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 12 * 1024 * 1024 },
});

function _openaiSttCall(audioBuffer, mimetype, originalName, model, biasPrompt) {
  return new Promise((resolve, reject) => {
    const boundary = '----nlplboundary' + Date.now();
    // Pick a sane filename + extension based on the mime so OpenAI can
    // detect the format. Browser MediaRecorder usually emits audio/webm
    // (Opus) — Whisper supports it, but only when the filename ends in
    // a known extension.
    const mt = String(mimetype || 'audio/webm').toLowerCase();
    let ext = 'webm';
    if (mt.indexOf('mp4') >= 0) ext = 'mp4';
    else if (mt.indexOf('mpeg') >= 0 || mt.indexOf('mp3') >= 0) ext = 'mp3';
    else if (mt.indexOf('wav') >= 0) ext = 'wav';
    else if (mt.indexOf('ogg') >= 0) ext = 'ogg';
    else if (mt.indexOf('m4a') >= 0) ext = 'm4a';
    const filename = originalName && /\.[a-z0-9]+$/i.test(originalName) ? originalName : `audio.${ext}`;
    const ctype = mt;
    const head = Buffer.from(
      `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${filename}"\r\nContent-Type: ${ctype}\r\n\r\n`
    );
    // Optional `prompt` field biases Whisper / gpt-4o-mini-transcribe toward
    // domain vocabulary (Indian branch names, employee names). Hard limit
    // 224 tokens per OpenAI docs — we cap at 900 chars to stay safely under.
    let trailer = `\r\n--${boundary}\r\nContent-Disposition: form-data; name="model"\r\n\r\n${model}`;
    if (biasPrompt && typeof biasPrompt === 'string') {
      const trimmed = biasPrompt.length > 900 ? biasPrompt.slice(0, 900) : biasPrompt;
      if (trimmed.trim()) {
        trailer += `\r\n--${boundary}\r\nContent-Disposition: form-data; name="prompt"\r\n\r\n${trimmed}`;
      }
    }
    trailer += `\r\n--${boundary}--\r\n`;
    const tailFile = Buffer.from(trailer);
    const body = Buffer.concat([head, audioBuffer, tailFile]);
    const req = https.request({
      hostname: 'api.openai.com',
      path: '/v1/audio/transcriptions',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_KEY,
        'Content-Type': 'multipart/form-data; boundary=' + boundary,
        'Content-Length': body.length,
      },
    }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        if (resp.statusCode >= 200 && resp.statusCode < 300) {
          try { return resolve((JSON.parse(buf).text || '').trim()); }
          catch (e) { return reject(new Error('stt_parse: ' + e.message)); }
        }
        const err = new Error('stt_http_' + resp.statusCode + ': ' + buf.slice(0, 240));
        err.statusCode = resp.statusCode;
        err.body = buf;
        reject(err);
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => req.destroy(new Error('stt_timeout')));
    req.write(body);
    req.end();
  });
}

// Public STT: try the configured model; if it returns 400 with the
// "audio file might be corrupted" signal (gpt-4o-mini-transcribe does
// this for a lot of valid browser webm), retry with whisper-1 which is
// the most permissive model OpenAI offers.
async function openaiStt(audioBuffer, mimetype, originalName, biasPrompt) {
  try {
    return await _openaiSttCall(audioBuffer, mimetype, originalName, OPENAI_STT, biasPrompt);
  } catch (e) {
    const isAudioReject = e && e.statusCode === 400 && /corrupt|unsupport|invalid/i.test(String(e.body || ''));
    if (isAudioReject && OPENAI_STT !== 'whisper-1') {
      console.warn('STT primary (' + OPENAI_STT + ') rejected audio, retrying with whisper-1');
      return await _openaiSttCall(audioBuffer, mimetype, originalName, 'whisper-1', biasPrompt);
    }
    throw e;
  }
}

// ── STT bias prompt cache ──────────────────────────────────────────────────
// Whisper / gpt-4o-mini-transcribe accept an optional `prompt` field that
// biases recognition toward domain vocabulary. We feed it branch names
// (and later employee first names) so "Davanagere" stops becoming
// "Davangere", etc. Cache per scope for 30 min — branch list changes
// rarely and rebuilding it on every voice request is wasteful.
const sttBiasCache = new Map(); // key: "ROLE|location" → { prompt, ts }
const STT_BIAS_TTL_MS = 30 * 60 * 1000;
const STT_BIAS_MAX = 50;
const STT_BIAS_BUDGET_CHARS = 880;

async function getSttBiasPrompt(session) {
  try {
    const role = String((session && session.role) || '').toUpperCase();
    const loc = String((session && session.location) || '').trim();
    const key = role + '|' + loc;
    const now = Date.now();
    const cached = sttBiasCache.get(key);
    if (cached && now - cached.ts < STT_BIAS_TTL_MS) return cached.prompt;

    let emCol = null;
    if (role === 'RM' || role === 'SM') emCol = 'region_name';
    else if (role === 'DM' || role === 'DVM') emCol = 'division_name';
    else if (role === 'AM') emCol = 'area_name';
    else if (role === 'BM' || role === 'FO') emCol = 'branch_name';

    const inScope = (emCol && loc)
      ? await safeQuery(
          `SELECT branch_name, COUNT(*)::int AS n
             FROM employee_master
            WHERE status='Working' AND TRIM(${emCol}) ILIKE TRIM($1)
            GROUP BY branch_name
            ORDER BY n DESC`, [loc], 200)
      : [];
    const global = await safeQuery(
      `SELECT branch_name, COUNT(*)::int AS n
         FROM employee_master
        WHERE status='Working'
        GROUP BY branch_name
        ORDER BY n DESC
        LIMIT 200`, [], 200);

    // In-scope branches first (highest signal), then fill with global by
    // headcount desc. Dedupe case-insensitively.
    const seen = new Set();
    const ordered = [];
    const push = (rows) => {
      for (const r of (Array.isArray(rows) ? rows : [])) {
        const n = String(r.branch_name || '').trim();
        if (!n) continue;
        const k = n.toUpperCase();
        if (seen.has(k)) continue;
        seen.add(k);
        ordered.push(n);
      }
    };
    push(inScope);
    push(global);

    // Greedy pack under STT_BIAS_BUDGET_CHARS (≈ 220 tokens, safely under
    // the 224-token Whisper prompt limit).
    const parts = [];
    let len = 0;
    for (const name of ordered) {
      const add = (parts.length === 0 ? name.length : name.length + 2);
      if (len + add > STT_BIAS_BUDGET_CHARS) break;
      parts.push(name);
      len += add;
    }
    const prompt = parts.join(', ');

    sttBiasCache.set(key, { prompt, ts: now });
    if (sttBiasCache.size > STT_BIAS_MAX) {
      const oldest = sttBiasCache.keys().next().value;
      if (oldest !== undefined) sttBiasCache.delete(oldest);
    }
    return prompt;
  } catch (e) {
    console.error('getSttBiasPrompt error:', e.message);
    return '';
  }
}

function openaiChat(systemText, userText) {
  const messages = [];
  if (systemText) messages.push({ role: 'system', content: systemText });
  messages.push({ role: 'user', content: userText });
  const payload = JSON.stringify({
    model: OPENAI_LLM,
    messages,
    max_tokens: 400,
    temperature: 0.4,
  });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.openai.com',
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        if (resp.statusCode >= 200 && resp.statusCode < 300) {
          try {
            const j = JSON.parse(buf);
            return resolve(j?.choices?.[0]?.message?.content || '');
          } catch (e) { return reject(new Error('chat_parse: ' + e.message)); }
        }
        reject(new Error('chat_http_' + resp.statusCode + ': ' + buf.slice(0, 240)));
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => req.destroy(new Error('chat_timeout')));
    req.write(payload);
    req.end();
  });
}

function openaiTts(text) {
  const payload = JSON.stringify({
    model: OPENAI_TTS,
    voice: OPENAI_VOICE,
    input: text,
    format: 'mp3',
  });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.openai.com',
      path: '/v1/audio/speech',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, (resp) => {
      const chunks = [];
      resp.on('data', (d) => chunks.push(d));
      resp.on('end', () => {
        const out = Buffer.concat(chunks);
        if (resp.statusCode >= 200 && resp.statusCode < 300) return resolve(out);
        reject(new Error('tts_http_' + resp.statusCode + ': ' + out.toString().slice(0, 240)));
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => req.destroy(new Error('tts_timeout')));
    req.write(payload);
    req.end();
  });
}

// ========== OpenAI Image Generation (DALL-E) ==========
function openaiImageGen(prompt, size, quality) {
  const sz = size || DALLE_SIZE;
  const q = quality || DALLE_QUALITY;
  const payload = JSON.stringify({
    model: DALLE_MODEL,
    prompt,
    n: 1,
    size: sz,
    quality: q,
    response_format: 'url',
  });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'api.openai.com',
      path: '/v1/images/generations',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        if (resp.statusCode >= 200 && resp.statusCode < 300) {
          try {
            const j = JSON.parse(buf);
            const url = j?.data?.[0]?.url;
            if (url) return resolve({ url, revised_prompt: j?.data?.[0]?.revised_prompt || null });
            return reject(new Error('image_gen: no url in response'));
          } catch (e) { return reject(new Error('image_parse: ' + e.message)); }
        }
        reject(new Error('image_http_' + resp.statusCode + ': ' + buf.slice(0, 240)));
      });
    });
    req.on('error', reject);
    req.setTimeout(60000, () => req.destroy(new Error('image_timeout')));
    req.write(payload);
    req.end();
  });
}

// POST /api/image — generate image via DALL-E
// Body: { prompt, size?, quality? }
// Returns: { url, prompt, size, quality, cost, model }
app.post('/api/image', aiLimiter, async (req, res) => {
  if (!OPENAI_KEY) return res.status(503).json({ error: 'ai_not_configured' });

  const { prompt, size, quality } = req.body || {};
  if (!prompt || typeof prompt !== 'string' || !prompt.trim()) {
    return res.status(400).json({ error: 'prompt is required' });
  }

  const sz = size || DALLE_SIZE;
  const q = quality || DALLE_QUALITY;
  const cost = sz === '1024x1792' || sz === '1792x1024'
    ? (q === 'hd' ? 0.120 : 0.080)
    : (q === 'hd' ? 0.080 : 0.040);

  try {
    const result = await openaiImageGen(prompt, sz, q);
    console.log(`[ai-activity] image_gen model=${DALLE_MODEL} size=${sz} quality=${q} cost=$${cost} prompt="${prompt.slice(0, 80)}"`);
    res.json({
      url: result.url,
      prompt,
      revised_prompt: result.revised_prompt,
      size: sz,
      quality: q,
      cost,
      model: DALLE_MODEL,
    });
  } catch (e) {
    console.error('image generation error:', e.message);
    res.status(500).json({ error: 'image_generation_failed', detail: e.message.slice(0, 240) });
  }
});

// POST /api/transcribe — STT-only (no LLM, no TTS). Used by chat-input mic
// so the user can dictate into the textarea without invoking voice cockpit.
app.post('/api/transcribe', voiceUpload.single('audio'), requireAiAccess, async (req, res) => {
  const origin = req.headers.origin || req.headers.referer || '';
  const mobileOrigin = String(req.headers['x-app-origin'] || '').toLowerCase();
  if (origin) {
    if (!origin.includes('navachetanalivelihoods.com') && !origin.includes('localhost')) {
      return res.status(403).json({ error: 'Forbidden' });
    }
  } else if (mobileOrigin !== 'nlpl-mobile') {
    return res.status(403).json({ error: 'Forbidden' });
  }
  if (!OPENAI_KEY) return res.status(503).json({ error: 'openai_not_configured' });
  if (!req.file || !req.file.buffer || !req.file.buffer.length) {
    return res.status(400).json({ error: 'audio_file_required' });
  }
  try {
    const role = String(req.body.role || '').trim();
    const location = String(req.body.location || '').trim();
    const session = (role && location) ? { role, location } : {};
    const biasPrompt = await getSttBiasPrompt(session);
    const transcript = await openaiStt(req.file.buffer, req.file.mimetype, req.file.originalname, biasPrompt);
    res.json({ transcript: transcript || '' });
  } catch (e) {
    console.error('transcribe error:', e.message);
    res.status(500).json({ error: 'transcribe_failed', detail: e.message.slice(0, 240) });
  }
});

// POST /api/voice — multipart upload field "audio". Optional fields: role,
// location. Returns { transcript, reply, audio_b64 (mp3 base64) }.
app.post('/api/voice', voiceUpload.single('audio'), requireAiAccess, async (req, res) => {
  const origin = req.headers.origin || req.headers.referer || '';
  const mobileOrigin = String(req.headers['x-app-origin'] || '').toLowerCase();
  if (origin) {
    if (!origin.includes('navachetanalivelihoods.com') && !origin.includes('localhost')) {
      return res.status(403).json({ error: 'Forbidden' });
    }
  } else if (mobileOrigin !== 'nlpl-mobile') {
    return res.status(403).json({ error: 'Forbidden' });
  }
  if (!OPENAI_KEY) return res.status(503).json({ error: 'openai_not_configured' });
  if (!req.file || !req.file.buffer || !req.file.buffer.length) {
    return res.status(400).json({ error: 'audio_file_required' });
  }
  try {
    const role = String(req.body.role || '').trim();
    const location = String(req.body.location || '').trim();
    const session = (role && location) ? { role, location } : {};
    // Bias Whisper toward in-scope branch names so Indian place names
    // ("Davanagere", "Channagiri") stop getting Anglicised mid-pipeline.
    const biasPrompt = await getSttBiasPrompt(session);
    const transcript = await openaiStt(req.file.buffer, req.file.mimetype, req.file.originalname, biasPrompt);
    if (!transcript) {
      return res.json({ transcript: '', reply: '', audio_b64: null, note: 'no_speech_detected' });
    }
    let systemText = `You are the NLPL Dashboard AI Assistant. LANGUAGE: detect the language the user spoke in (English or Kannada) and reply in the SAME language. If the user spoke Kannada, reply in Kannada (ಕನ್ನಡ script). If the user spoke English, reply in English. If mixed, follow the dominant language. Reply in under 50 words, conversational tone (this will be spoken aloud). In Kannada replies, ALL numbers, percentages, dates, currency values, and number unit words must stay in English digits/English words: "12.34 crore", "5.6 lakh", "45%", "2026-05-01". Do not use Kannada digits or Kannada number words. Quote numbers only from the data below; if missing say "data not available" / "ಲಭ್ಯವಿಲ್ಲ". Never mention the words "snapshot" or "data block".`;
    try {
      const ctx = await buildDataContext(session);
      const summary = ctx.summary_text || JSON.stringify(ctx).slice(0, 6000);
      systemText += `\n\nToday: ${ctx.now}. Scope: ${role && location ? `${role} ${location}` : 'CEO (all)'}.\n\nDATA:\n${summary}`;
    } catch (_) {}
    const reply = normalizeKannadaNumbersForAi(await openaiChat(systemText, transcript));
    let audioB64 = null;
    try {
      const buf = await openaiTts(reply);
      audioB64 = buf.toString('base64');
    } catch (e) {
      console.error('voice tts skipped:', e.message);
    }
    res.json({ transcript, reply, audio_b64: audioB64 });
  } catch (e) {
    console.error('voice error:', e.message);
    res.status(500).json({ error: 'voice_failed', detail: e.message.slice(0, 240) });
  }
});

// POST /api/voice-stream — SSE-style. Same multipart audio input as /api/voice
// but streams: transcript → tool_result events (with SQL + rows) → reply →
// audio. Used by the new voice cockpit overlay so SQL/data appear live.
app.post('/api/voice-stream', aiLimiter, voiceUpload.single('audio'), requireAiAccess, async (req, res) => {
  const origin = req.headers.origin || req.headers.referer || '';
  const mobileOrigin = String(req.headers['x-app-origin'] || '').toLowerCase();
  if (origin) {
    if (!origin.includes('navachetanalivelihoods.com') && !origin.includes('localhost')) {
      return res.status(403).json({ error: 'Forbidden' });
    }
  } else if (mobileOrigin !== 'nlpl-mobile') {
    return res.status(403).json({ error: 'Forbidden' });
  }
  if (!OPENAI_KEY) return res.status(503).json({ error: 'openai_not_configured' });
  if (!req.file || !req.file.buffer || !req.file.buffer.length) {
    return res.status(400).json({ error: 'audio_file_required' });
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders && res.flushHeaders();
  // Disable Nagle on the underlying socket so tiny SSE writes (events +
  // heartbeats) are sent immediately instead of being coalesced.
  try { req.socket.setNoDelay(true); } catch (_) {}

  let closed = false;
  let heartbeat = null;
  const stopHeartbeat = () => {
    if (heartbeat) { clearInterval(heartbeat); heartbeat = null; }
  };
  const finish = () => {
    closed = true;
    stopHeartbeat();
    try { res.end(); } catch (_) {}
  };
  // NOTE: do NOT listen on req.on('close') — for multipart uploads, the
  // request stream emits 'close' as soon as multer finishes parsing the
  // body, well before the response is finalised. That would mis-flag the
  // connection as closed and silence every subsequent SSE write.
  // res.on('close') is the correct "client disconnected" signal here.
  res.on('close', () => { closed = true; stopHeartbeat(); });

  const send = (event, data) => {
    if (closed || res.writableEnded) return;
    try { res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`); } catch (_) {}
  };
  // Heartbeat comments every 250 ms keep the connection alive and push
  // events through Apache mod_proxy_http even during long await steps
  // (STT, tool round-trips, TTS). Padded so the chunk is large enough to
  // cross any small intermediate buffers.
  const HB_PAD = ': ' + 'x'.repeat(2048) + '\n\n';
  heartbeat = setInterval(() => {
    if (closed || res.writableEnded) { stopHeartbeat(); return; }
    try { res.write(HB_PAD); } catch (_) { stopHeartbeat(); }
  }, 250);

  try {
    send('open', {});
    send('stage', { stage: 'stt' });

    // Parse session up-front so the STT call can be biased toward in-scope
    // branch names. getSttBiasPrompt is cached + safe to call every request.
    const role = String(req.body.role || '').trim();
    const location = String(req.body.location || '').trim();
    const session = (role && location) ? { role, location } : {};
    const biasPrompt = await getSttBiasPrompt(session);

    const transcript = await openaiStt(req.file.buffer, req.file.mimetype, req.file.originalname, biasPrompt);
    if (!transcript) {
      send('error', { message: 'No speech detected. Try again.' });
      return finish();
    }
    send('transcript', { text: transcript });
    send('stage', { stage: 'thinking' });

    // Build the same rich tool-aware prompt as the text streaming endpoint,
    // but add a voice-mode rule so the spoken reply stays brief.
    let scopedSystemText = '';
    try {
      const ctx = await buildDataContext(session);
      const scopeLabel = (role && location) ? `${role} for ${location}` : 'CEO (all data)';
      // CEO ctxJson is ~36k tokens — exceeds OpenAI tier-1 30k TPM in a
      // single request. Embed only the compact summary; tools fetch the
      // rest on demand.
      const ctxLite = {
        scope: ctx.scope,
        latestDate: ctx.latestDate,
        fyStart: ctx.fyStart,
        now: ctx.now,
        latest: ctx.latest,
        summary_text: ctx.summary_text,
      };
      const ctxJson = JSON.stringify(ctxLite);
      scopedSystemText = [
        `You are the NLPL Dashboard AI Assistant for ${scopeLabel}.`,
        `Today is ${ctx.now}. FY started ${ctx.fyStart} (April 1 → March 31).`,
        '',
        '## Language',
        '- Detect the language of the transcript and REPLY IN THE SAME LANGUAGE (English or Kannada / ಕನ್ನಡ).',
        '- **DATABASE IS ENGLISH-ONLY.** All identifiers — employee names, branch names, regions, divisions, areas, products — are stored in English (Latin script). Whenever the user speaks a name in Kannada, you MUST transliterate it to English BEFORE passing it to any tool. Examples:',
        '  - ರಘುನಂದನ್ → "Raghunandan"  (find_employee query)',
        '  - ದಾವಣಗೆರೆ → "Davanagere"     (find_branch query)',
        '  - ಬೆಂಗಳೂರು → "Bangalore"      (find_branch query)',
        '  - ಕಲಬುರಗಿ → "Kalaburagi"      (region_name)',
        'Searching the DB with Kannada script will return ZERO rows. Always transliterate the entity name to English before the tool call. The reply text itself stays in Kannada.',
        '- In Kannada replies, ALL numbers, percentages, dates, currency values, and number unit words stay in English digits/English words: "12.34 crore", "5.6 lakh", "45%", "2026-05-01". Do not use Kannada digits or Kannada number words.',
        '',
        '## Voice mode rules',
        '- This reply will be spoken aloud. Keep it under 60 words, conversational, no markdown.',
        '- Lead with the headline number, then ONE supporting sentence.',
        '- The companion screen shows the raw rows you fetch — you do not need to dictate every value, just the highlight + interpretation.',
        '',
        '## Tools available — CALL THESE for any specifics',
        '- resolve_date_range, find_employee, employee_performance, employee_collection_series, find_branch, period_performance, top_performers, disbursement_query, list_hierarchy, list_employees, headcount, npa_summary, daily_reports_query, branch_summary, hourly_collection, collection_drilldown, period_compare, plan_compliance, sql_describe',
        '',
        '## Critical tool routing',
        '- "Show <person> trend / <person> performance over time / drill on <person> / day-by-day for <person> / how is <person> doing day to day" → call `find_employee(query=name, location_hint?)` first to get the canonical emp_id, then `employee_collection_series(emp_id, start_date?, end_date?)` for the per-day series. For TOTAL across a window (single aggregate row) use `employee_performance(emp_id, ...)`. NEVER call `period_performance(group_by="employee", branch_name=...)` or `top_performers(branch_name=...)` for a single named person — those return the WHOLE branch leaderboard, not the named person.',
        '- "How is my branch doing today / how is <branch> doing / branch health / give me a summary of <branch> / end-of-day rollup" → call `branch_summary()` (BM/ABM/BOE pass NO args — auto-resolves; CEO/RM/DM/AM pass branch_name). Returns headcount + collection (today/MTD/FYTD) + NPA + disbursement + FTOD + DPD plan-vs-actual in ONE call. Don\'t chain 5 separate tools for this.',
        '- "Why is collection low / drill down on collection / who\'s underperforming / which FOs are dragging us down" → call `collection_drilldown()`. Returns 3 worst FOs + 3 best FOs + DPD buckets + NPA + FTOD gap in ONE call. Use AFTER branch_summary for the "why" follow-up.',
        '- "MoM / WoW / vs last month / vs yesterday / this week vs last week / vs same period last year / compare X to Y" → call `period_compare(metric, scope, scope_value?, period_a_start, period_a_end, period_b_start, period_b_end)`. ALL 4 dates ISO YYYY-MM-DD. metric ∈ {collection, demand, collection_pct, npa_amount, disb_amount, disb_count, ftod}. scope="branch" + scope_value=<branch> for single-branch MoM. Server computes delta_abs + delta_pct — DO NOT call period_performance twice and do the math yourself (you mix counts vs amounts and miscompute pct-of-pct).',
        '- "Which branches missed plan today / who didn\'t file daily report / plan compliance / branches without daily report" → call `plan_compliance(date?, scope?, scope_value?)`. Date defaults to today (ISO YYYY-MM-DD if passed). Returns missing_plan + missing_achievement lists. CEO/RM/DM/AM only — BM rejected (single branch).',
        '- "Collection right now / live collection / how much collected today so far / hourly trend / intra-day" → call `hourly_collection()`. Returns CURRENT live snapshot (counts + amounts + pct), not a time series.',
        '- "How many employees / staff / BMs / FOs / agents in <scope>?" → call `headcount(group_by=\'role\')` for a role breakdown, or `headcount(role=\'BM\')` for a single role count, or `headcount()` for the overall total. NEVER quote a hierarchy row count or list_hierarchy result count as "number of employees".',
      '- "List all employees / give me everyone / phone numbers of staff in <branch>" / "show roster" → call `list_employees(branch_name=...)`. NOT find_employee — that\'s for searching ONE named person. list_employees returns the full roster (name + role + branch + mobile).',
        '- "Top N employees in <branch> by collection / demand" → call `top_performers(metric, start_date, end_date, branch_name=...)`. NEVER chain list_employees + employee_performance × N — that loses the ranking and produces wrong/zero rows.',
        '- **NPA semantics matter.** `npa_cases` from daily_performance / employee_performance is a ROLLING SNAPSHOT (the same outstanding case is repeated every day). Never SUM npa_cases over a date range — that inflates by ~days. For outstanding cases, query the latest single date. For activation / closure RATES, use `daily_reports_query(table="achievements", metrics="npa")` and SUM npa_activation / npa_closure (those are per-day deltas).',
        '- If today\'s daily_performance is empty (EOD lag — common before evening), retry against `daily_reports_query(table="achievements")` for today\'s achievement numbers before saying "no data".',
      '- "Who is the BM of <branch>" / "who is the manager" → `list_employees(branch_name=X, role=\'BM\')`.',
        '- "Daily collection / collection trend / day-by-day collection" → call `period_performance(group_by=\'day\', start_date, end_date)`. The result has one row per day with demand + collection columns. Average = SUM(collection) / COUNT(days).',
        '- "Daily disbursement" → call `disbursement_query(group_by=\'day\', start_date, end_date)`.',
        '- For list_hierarchy: NEVER report its row count as a money or collection metric. Its rows are entity names with employee_count + branch_count metadata only.',
        '',
        '## Tool guidance',
        '- **CANONICALISE NAMES FIRST.** Voice transcription mishears branch / employee / region names ("Davangere" vs "Davanagere", "AP Shivraj" / "EPI Shivraj" vs "A P Shivaraj"). If the user names a specific branch/employee, your FIRST tool call MUST be `find_branch(query=<spoken name>)` or `find_employee(query=<spoken name>)`. Use the canonical `branch_name` / `emp_id` from the result in every subsequent tool call. Never pass the raw spoken name into disbursement_query / period_performance / daily_reports_query / top_performers — always pass the canonical value find_branch/find_employee returned.',
        '- **WHEN USER NAMES A PERSON + LOCATION** (e.g. "Shivraj from Raichur", "Karthik in Bangalore"), pass BOTH: `find_employee(query="Shivraj", location_hint="Raichur")`. The hint disambiguates common names down to one row.',
        '- find_branch + find_employee both do substring + trigram + word-similarity match — misspellings and mishearings still resolve. Trust the result.',
        '- If find_branch / find_employee returns `ambiguous: true`, ask the user to confirm (list 2-4 candidates with region/branch). Do NOT pick silently.',
        '- If find_employee returns 0 rows for a named person, DO NOT fall back to overall company numbers. Tell the user "no employee matched that name" and ask them to repeat the name or provide an emp_id / mobile / branch hint.',
        '- ANY question with the words disbursement / disb / disbursed / loan disbursed → call disbursement_query. **READ THE USER\'S WORDS to pick group_by**: "product wise / by product / break by product / split by product" → group_by="product". "branch wise / per branch" → group_by="branch". "day by day / daily / on Apr 7 vs 8" → group_by="day". "month by month / monthly trend" → group_by="month". "by employee / by FO / per officer" → group_by="employee". If user combines (e.g. "FO wise AND product wise"), call disbursement_query TWICE — once with group_by="employee" and once with group_by="product" — and present both. Pass the canonical branch_name from find_branch when a branch was named.',
        '- ANY question about FTOD / DPD / KYC / NPA closure / daily plan → call daily_reports_query.',
        '- For collection/demand metrics → call period_performance.',
        '- For FTOD specifically: also fetch demand and collection so the user sees demand + collection + FTOD together (FTOD = demand - collection).',
        '- For comparison questions (vs last month, top N, etc.) → call period_performance or top_performers.',
        '- For any relative, spoken, slash-format, or comparison date phrase ("today", "yesterday", "11/04", "April 11", "this month", "last month"), call `resolve_date_range` first and pass its ISO dates to data tools.',
        '- "this month" → first of current month → today. "last month" → prior calendar month full range.',
        '- If a tool returns 0 rows for the requested period, IMMEDIATELY retry with the previous month, then the previous quarter, until you find data. Then tell the user "no data for {requested period}; latest available is {found period}". Do NOT just say "no data" without trying nearby periods.',
        '- If the at-a-glance JSON below has the answer for a trivial single number, you may quote it directly. For everything else, CALL A TOOL.',
        '- Never say "data not available" without first calling the relevant tool.',
        '',
        '## Hard rules — DO NOT BREAK',
        '- **Single-entity scope.** When the user names ONE person, branch, or region in the question (e.g. "show Manjunatha T N trend", "how is Davanagere doing", "drill on Pavan Kumar K"), the headline answer is scoped to THAT entity ONLY. For a single named PERSON: call `find_employee` → then `employee_collection_series(emp_id, …)` for trend/series, or `employee_performance(emp_id, …)` for totals. NEVER call `period_performance(group_by="employee", branch_name=…)` or `top_performers(branch_name=…)` to answer a single-person question — those return the whole branch leaderboard. For a single named BRANCH: use `branch_summary` / `collection_drilldown`. Listing peers (the rest of the branch, the rest of the region) is allowed ONLY when the user explicitly asks for a ranking, comparison, or "everyone in <branch>".',
        '- **top_performers ranks EMPLOYEES, not branches.** For "top N branches by collection / demand / disbursement / NPA", call `period_performance(group_by=\'branch\', start_date, end_date)` and pick the top N rows yourself. NEVER substitute disbursement_query for collection or vice versa.',
        '- **Never relabel one tool\'s output as a different metric.** disbursement is NOT collection; npa_activation is NOT npa_closure. If the user asks for collection and you only have disbursement, fetch collection.',
        '- **Never fabricate totals.** When a tool returns N rows, your reply must either (a) sum all N rows in the answer, or (b) explicitly say "showing top X of N — total is Y" using the actual sum. If you can\'t compute the total, say "showing first N rows" without a fake total.',
        '- **Cross-branch compare permission depends on the SESSION role above.** CEO / RM / SM / DM / DvM / AM may compare ANY two branches freely — call period_compare with scope="branch" and the two scope_values. Branch-bound roles (BM / ABM / BOE) may ONLY compare their own branch; if the user is BM/ABM/BOE in the session and they ask to compare against any OTHER branch, refuse with: "You can only access <own branch>. Cross-branch comparisons require RM or CEO access." NEVER apply the BM refusal when the session role is CEO/RM/SM/DM/DvM/AM. NEVER call period_compare with own-branch as scope_value for both windows — that silently produces a fake-parity comparison.',
        '- **regular_demand / regular_collection are COUNTS, not money.** Even when narrating branch_summary or period_compare results, format these as plain Indian-comma numbers (e.g. "1,335 collection accounts"), NOT ₹ / L / Cr. Only columns ending in `_amt` and {disb_amount, npa_act_amt, npa_amount, npa_clo_amt} are money. This rule has been ignored in past replies — comply.',
        '- **Multi-employee period queries: ONE call, not N.** When the user asks for a LIST or TABLE of multiple employees with their performance over a period (e.g. "every FO in <branch> with their April collection", "show all my staff and their collection"), call `top_performers(metric=\'collection\', branch_name=<scope>, role=<role-if-named>, start_date, end_date, limit=200)` ONCE — top_performers ALSO accepts role/branch/area/division/region filters and returns the full leaderboard. Or `period_performance(group_by=\'employee\', branch_name=<scope>, start_date, end_date)` if no role filter. NEVER call `list_employees` then loop `employee_performance(emp_id=...)` per person — that anti-pattern wastes round-trips and fragments the output. Reserve `employee_performance` for a SINGLE named individual. Multi-axis requests (per-emp × per-day × per-product × per-DPD-bucket) — pick ONE primary axis, call once, and tell the user the secondary axis needs a follow-up question.',
        '- **Active-filter preamble is BINDING SCOPE, not narrative.** If the latest user message starts with `[Active filter: ...]`, parse the bracket and apply it to every tool call until the next user message overrides it. Three shapes: (a) SINGLE — `[Active filter: employee NL11007 (R Gagan) at Ajjampura branch.]` → pass `emp_id="NL11007"` to employee_performance and pin `branch_name="Ajjampura"` for any branch-scoped follow-up. (b) SET with explicit emp_ids — `[Active filter: 8 employees in Ajjampura branch — emp_ids: NL11007, NL12292, ...]` → call ONE branch-scoped tool (`top_performers(branch_name="Ajjampura", role?, start, end)` or `period_performance(group_by="employee", branch_name="Ajjampura", start, end)`) and IN YOUR REPLY only narrate the rows whose emp_id appears in the bracket list (filter client-side). NEVER loop employee_performance per emp_id. (c) SET-ALL — `[Active filter: all employees in Ajjampura branch.]` → branch-scoped tool ONCE, narrate all rows. Do NOT echo the bracket back in your reply — it is a routing directive, not user-visible content. Reply only to the question that follows the bracket.',
        '',
        '## Units — READ CAREFULLY',
        '**COUNT columns (integers, NOT money — never format as ₹/L/Cr):**',
        '  regular_demand, regular_collection, demand_1_30, collection_1_30, demand_31_60, ..., npa_cases, npa_act_acc, npa_clo_acc, disb_count, ftod, kyc_*, fy_non_start_acc.',
        '**MONETARY columns (rupees — divide by 1,00,000 for L or 1,00,00,000 for Cr):**',
        '  regular_demand_amt, regular_collection_amt, demand_*_amt, npa_act_amt, npa_clo_amt, disb_amount, disb_*_amt.',
        'Rule of thumb: any column ending in `_amt` or `disb_amount` is money. Everything else is a count.',
        'Collection % = regular_collection / regular_demand × 100 (count ratio, dimensionless).',
        '',
        '## Internal context block (do NOT mention)',
        ctxJson,
      ].join('\n');
    } catch (e) {
      console.error('voice-stream context error:', e.message);
    }

    // Active-context preamble (task #15) — frontend sends `context_preamble`
    // as a form field for voice flows since we can't prepend to audio bytes.
    // The system prompts already know how to parse `[Active filter: ...]`
    // brackets at the start of the user message; just stitch them together.
    const _ctxPreamble = String((req.body && req.body.context_preamble) || '').trim();
    const userTurn = _ctxPreamble
      ? (_ctxPreamble + '\n\nUser: ' + transcript)
      : transcript;

    const messages = [
      { role: 'system', content: scopedSystemText },
      { role: 'user', content: userTurn },
    ];

    const onProgress = (ev) => {
      if (closed) return;
      if (ev.type === 'thinking') send('thinking', { round: ev.round });
      else if (ev.type === 'tools') send('tools', { names: ev.names });
      else if (ev.type === 'tool_result') send('tool_result', {
        name: ev.name,
        args: ev.args,
        queries: ev.queries,
        result: ev.result,
      });
    };

    // Provider chain: prefer Mistral (best tool dispatch), then OpenAI.
    // Voice replies are short — cap at 6 tool rounds so a runaway loop
    // doesn't burn the TPM budget.
    // If the client passed an explicit `provider` ('openai'|'mistral'),
    // use ONLY that one — no fallback. Missing key for that provider
    // is sent back as a 503-style SSE error so the UI can prompt the user.
    const VOICE_MAX_ROUNDS = 6;
    const explicitProvider = String((req.body && req.body.provider) || '').toLowerCase().trim();
    const voiceDeepSeekOverride = String((req.body && req.body.deepseek_model) || '').toLowerCase().trim() || undefined;
    const chain = [];
    if (explicitProvider === 'mistral') {
      if (!MISTRAL_KEY) {
        send('error', { message: 'Mistral not configured on this server.', reason: 'provider_unavailable', provider: 'mistral' });
        return finish();
      }
      chain.push({ name: 'mistral', run: () => runMistralWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress) });
    } else if (explicitProvider === 'openai') {
      if (!OPENAI_KEY) {
        send('error', { message: 'OpenAI not configured on this server.', reason: 'provider_unavailable', provider: 'openai' });
        return finish();
      }
      chain.push({ name: 'openai', run: () => runOpenAiWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress) });
    } else if (explicitProvider === 'deepseek') {
      if (!DEEPSEEK_KEY) {
        send('error', { message: 'DeepSeek not configured on this server.', reason: 'provider_unavailable', provider: 'deepseek' });
        return finish();
      }
      chain.push({ name: 'deepseek', run: () => runDeepSeekWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress, voiceDeepSeekOverride) });
    } else {
      if (MISTRAL_KEY)  chain.push({ name: 'mistral',  run: () => runMistralWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress) });
      if (OPENAI_KEY)   chain.push({ name: 'openai',   run: () => runOpenAiWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress) });
      if (DEEPSEEK_KEY) chain.push({ name: 'deepseek', run: () => runDeepSeekWithTools(messages, session, VOICE_MAX_ROUNDS, onProgress, voiceDeepSeekOverride) });
    }

    let result = { ok: false, error: 'no_providers' };
    let usedProvider = null;
    for (let i = 0; i < chain.length; i++) {
      if (closed) return;
      const step = chain[i];
      if (i > 0) send('fallback', { from: chain[i - 1].name, to: step.name });
      result = await step.run();
      if (result.ok && result.text) { usedProvider = step.name; break; }
      console.error('voice-stream: ' + step.name + ' failed:', result.error);
    }

    if (closed) return;

    if (!result.ok || !result.text) {
      send('error', { message: 'AI is briefly busy. Please retry.', reason: result.error || 'all_providers_failed' });
      return finish();
    }

    const reply = normalizeKannadaNumbersForAi(result.text);
    send('reply', { text: reply });

    // Emit `done` BEFORE TTS so the cockpit unblocks immediately — user
    // has the answer rendered; voice is bonus. Late `audio` event still
    // plays if TTS completes within the 15s budget. Prevents the
    // "Synthesising voice…" hang we saw in production when OpenAI TTS
    // is slow on long Indian-English replies.
    send('done', { provider: usedProvider });
    send('stage', { stage: 'tts' });

    let audioB64 = null;
    try {
      const ttsPromise = openaiTts(reply);
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('tts_budget_exceeded_15s')), 15000)
      );
      const buf = await Promise.race([ttsPromise, timeoutPromise]);
      audioB64 = buf.toString('base64');
    } catch (e) {
      console.error('voice-stream tts skipped:', e.message);
    }
    if (audioB64 && !closed) send('audio', { mime: 'audio/mp3', data: audioB64 });

    // 150ms TCP-buffer flush before terminator (avoids
    // ERR_INCOMPLETE_CHUNKED_ENCODING on big audio chunks).
    await new Promise(r => setTimeout(r, 150));
    finish();
  } catch (e) {
    console.error('voice-stream error:', e.message);
    send('error', { message: 'voice_failed', detail: e.message.slice(0, 240) });
    finish();
  }
});

// ── In-memory response cache ───────────────────────────────────────────────────
// Key: SHA256(role | location | lastUserMsg). TTL: 10 min. Hard cap: 500 entries.
const aiReplyCache = new Map();
const AI_CACHE_TTL_MS = 10 * 60 * 1000;
const AI_CACHE_MAX = 500;

function getCacheKey(role, location, lastMsg, provider) {
  return crypto.createHash('sha256')
    .update(
      AI_REPLY_CACHE_VERSION + '|' +
      (provider || '') + '|' +
      (role || '') + '|' +
      (location || '') + '|' +
      (lastMsg || '')
    )
    .digest('hex');
}

function pruneCache() {
  const now = Date.now();
  for (const [k, v] of aiReplyCache) {
    if (now - v.ts > AI_CACHE_TTL_MS) aiReplyCache.delete(k);
  }
  if (aiReplyCache.size > AI_CACHE_MAX) {
    const toDelete = aiReplyCache.size - AI_CACHE_MAX;
    let i = 0;
    for (const k of aiReplyCache.keys()) {
      if (i++ >= toDelete) break;
      aiReplyCache.delete(k);
    }
  }
}

// ── AI provider call helpers ──────────────────────────────────────────────
// Both helpers take the same "OpenAI-format" `messages` array (with a single
// system message merged in) and return { ok, text, error }.

async function callMistralAi(messages, tools) {
  if (!MISTRAL_KEY) return { ok: false, error: 'mistral_not_configured' };
  const body = {
    model: MISTRAL_MODEL,
    messages,
    max_tokens: 1024,
    temperature: 0.3,
  };
  if (tools && tools.length) {
    body.tools = tools;
    body.tool_choice = 'auto';
  }
  const payload = JSON.stringify(body);
  return new Promise((resolve) => {
    const options = {
      hostname: 'api.mistral.ai',
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
        'Authorization': 'Bearer ' + MISTRAL_KEY,
      },
    };
    const r = https.request(options, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        try {
          const j = JSON.parse(buf);
          if (resp.statusCode >= 200 && resp.statusCode < 300) {
            const msg = j?.choices?.[0]?.message;
            if (!msg) return resolve({ ok: false, error: 'mistral_empty', raw: buf.slice(0, 200) });
            return resolve({ ok: true, message: msg, text: msg.content || '' });
          }
          return resolve({
            ok: false,
            error: 'mistral_http_' + resp.statusCode,
            raw: buf.slice(0, 200),
          });
        } catch (e) {
          resolve({ ok: false, error: 'mistral_parse: ' + e.message });
        }
      });
    });
    r.on('error', (e) => resolve({ ok: false, error: 'mistral_net: ' + e.message }));
    r.setTimeout(25000, () => r.destroy(new Error('mistral_timeout')));
    r.write(payload);
    r.end();
  });
}

// Tool-loop runner. Calls Mistral with AI_TOOLS_SPEC. If the model returns
// tool_calls, dispatches each, appends the results, and re-asks. Stops when
// the model returns plain text or after MAX_ROUNDS. Doesn't mutate `messages`.
async function runMistralWithTools(messages, session, maxRounds, onProgress) {
  const convo = messages.slice();
  const max = maxRounds || 12;
  const emit = typeof onProgress === 'function' ? onProgress : null;
  const callSig = new Map(); // signature -> count, to break infinite re-calls
  for (let round = 0; round < max; round++) {
    if (emit) emit({ type: 'thinking', round: round + 1 });
    const r = await callMistralAi(convo, AI_TOOLS_SPEC);
    if (!r.ok) {
      console.error(`[ai-tools] round ${round + 1}/${max} call failed: ${r.error}`);
      return r;
    }
    const msg = r.message || {};
    const calls = Array.isArray(msg.tool_calls) ? msg.tool_calls : null;
    if (!calls || !calls.length) {
      console.log(`[ai-tools] round ${round + 1}/${max} done (text len ${(msg.content || '').length})`);
      return { ok: true, text: normalizeKannadaNumbersForAi(msg.content || '') };
    }
    console.log(`[ai-tools] round ${round + 1}/${max} calling ${calls.length} tool(s): ${calls.map(c => c?.function?.name).join(', ')}`);
    if (emit) emit({ type: 'tools', names: calls.map(c => c?.function?.name).filter(Boolean) });
    // Echo the assistant turn (with tool_calls) back into the conversation.
    convo.push({
      role: 'assistant',
      content: msg.content == null ? null : msg.content,
      tool_calls: calls,
    });
    for (const c of calls) {
      const fnName = c?.function?.name || '';
      let argObj = {};
      try { argObj = JSON.parse(c?.function?.arguments || '{}'); } catch (_) {}
      const sig = fnName + '::' + JSON.stringify(argObj);
      const prev = callSig.get(sig) || 0;
      let toolResult;
      if (prev >= 1) {
        // Already tried this exact call. Synthesize a hard-stop tool result
        // so the model is forced to either change args or give up.
        toolResult = { error: 'duplicate_call_blocked', instruction: 'You already called this tool with these exact arguments. Do NOT retry. Either change the arguments (different date range, different filter) or answer the user with what you have. Do not loop.' };
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: [], result: toolResult });
      } else {
        const sqlLog = [];
        try {
          await aiToolLogStore.run(sqlLog, async () => {
            toolResult = await dispatchAiTool(fnName, argObj, session);
          });
        } catch (e) {
          toolResult = { error: 'tool_threw: ' + e.message };
        }
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: sqlLog, result: toolResult });
      }
      callSig.set(sig, prev + 1);
      convo.push({
        role: 'tool',
        tool_call_id: c.id,
        name: fnName,
        content: JSON.stringify(toolResult).slice(0, 24000),
      });
    }
  }
  return { ok: false, error: 'mistral_tool_loop_exceeded' };
}

// ── OpenAI chat with tool-calling — drop-in shape match for runMistralWithTools.
// OpenAI's chat-completions tool format is identical to Mistral's, so we can
// reuse AI_TOOLS_SPEC and dispatchAiTool unchanged.
async function callOpenAiChatTools(messages, tools) {
  if (!OPENAI_KEY) return { ok: false, error: 'openai_not_configured' };
  const body = {
    model: OPENAI_LLM,
    messages,
    max_tokens: 1024,
    temperature: 0.3,
  };
  if (tools && tools.length) {
    body.tools = tools;
    body.tool_choice = 'auto';
  }
  const payload = JSON.stringify(body);
  return new Promise((resolve) => {
    const r = https.request({
      hostname: 'api.openai.com',
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        try {
          const j = JSON.parse(buf);
          if (resp.statusCode >= 200 && resp.statusCode < 300) {
            const msg = j?.choices?.[0]?.message;
            if (!msg) return resolve({ ok: false, error: 'openai_empty', raw: buf.slice(0, 200) });
            return resolve({ ok: true, message: msg, text: msg.content || '' });
          }
          return resolve({ ok: false, error: 'openai_http_' + resp.statusCode, raw: buf.slice(0, 240) });
        } catch (e) {
          resolve({ ok: false, error: 'openai_parse: ' + e.message });
        }
      });
    });
    r.on('error', (e) => resolve({ ok: false, error: 'openai_net: ' + e.message }));
    r.setTimeout(30000, () => r.destroy(new Error('openai_timeout')));
    r.write(payload);
    r.end();
  });
}

async function runOpenAiWithTools(messages, session, maxRounds, onProgress) {
  const convo = messages.slice();
  const max = maxRounds || 12;
  const emit = typeof onProgress === 'function' ? onProgress : null;
  const callSig = new Map();
  for (let round = 0; round < max; round++) {
    if (emit) emit({ type: 'thinking', round: round + 1 });
    const r = await callOpenAiChatTools(convo, AI_TOOLS_SPEC);
    if (!r.ok) {
      console.error(`[openai-tools] round ${round + 1}/${max} call failed: ${r.error}`);
      return r;
    }
    const msg = r.message || {};
    const calls = Array.isArray(msg.tool_calls) ? msg.tool_calls : null;
    if (!calls || !calls.length) {
      console.log(`[openai-tools] round ${round + 1}/${max} done (text len ${(msg.content || '').length})`);
      return { ok: true, text: normalizeKannadaNumbersForAi(msg.content || '') };
    }
    console.log(`[openai-tools] round ${round + 1}/${max} calling ${calls.length} tool(s): ${calls.map(c => c?.function?.name).join(', ')}`);
    if (emit) emit({ type: 'tools', names: calls.map(c => c?.function?.name).filter(Boolean) });
    convo.push({
      role: 'assistant',
      content: msg.content == null ? null : msg.content,
      tool_calls: calls,
    });
    for (const c of calls) {
      const fnName = c?.function?.name || '';
      let argObj = {};
      try { argObj = JSON.parse(c?.function?.arguments || '{}'); } catch (_) {}
      const sig = fnName + '::' + JSON.stringify(argObj);
      const prev = callSig.get(sig) || 0;
      let toolResult;
      if (prev >= 1) {
        toolResult = { error: 'duplicate_call_blocked', instruction: 'You already called this tool with these exact arguments. Do NOT retry. Either change the arguments or answer the user with what you have. Do not loop.' };
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: [], result: toolResult });
      } else {
        const sqlLog = [];
        try {
          await aiToolLogStore.run(sqlLog, async () => {
            toolResult = await dispatchAiTool(fnName, argObj, session);
          });
        } catch (e) {
          toolResult = { error: 'tool_threw: ' + e.message };
        }
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: sqlLog, result: toolResult });
      }
      callSig.set(sig, prev + 1);
      convo.push({
        role: 'tool',
        tool_call_id: c.id,
        name: fnName,
        content: JSON.stringify(toolResult).slice(0, 24000),
      });
    }
  }
  return { ok: false, error: 'openai_tool_loop_exceeded' };
}

async function callGeminiAi(messages) {
  if (!GEMINI_KEY) return { ok: false, error: 'gemini_not_configured' };
  // Convert OpenAI-style messages to Gemini's contents/systemInstruction format.
  const systemMsg = messages.find((m) => m.role === 'system');
  const chatMessages = messages.filter((m) => m.role !== 'system');
  const contents = chatMessages.map((m) => ({
    role: m.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: m.content }],
  }));
  const systemInstruction = systemMsg
    ? { parts: [{ text: systemMsg.content }] }
    : undefined;
  for (const model of GEMINI_MODELS) {
    try {
      const payload = JSON.stringify({
        contents,
        ...(systemInstruction ? { systemInstruction } : {}),
        generationConfig: { maxOutputTokens: 1024, temperature: 0.3 },
      });
      const result = await new Promise((resolve, reject) => {
        const path = `/v1beta/models/${model}:generateContent?key=${GEMINI_KEY}`;
        const options = {
          hostname: 'generativelanguage.googleapis.com',
          path,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(payload),
          },
        };
        const r = https.request(options, (resp) => {
          let body = '';
          resp.on('data', (d) => (body += d));
          resp.on('end', () => {
            try { resolve(JSON.parse(body)); } catch (e) { reject(e); }
          });
        });
        r.on('error', reject);
        r.setTimeout(20000, () => r.destroy(new Error('gemini_timeout')));
        r.write(payload);
        r.end();
      });
      const text = result.candidates?.[0]?.content?.parts?.[0]?.text;
      if (text) return { ok: true, text: normalizeKannadaNumbersForAi(text), model };
      console.error('Gemini model ' + model + ' empty response:', JSON.stringify(result).slice(0, 200));
    } catch (e) {
      console.error('Gemini model ' + model + ' failed:', e.message);
    }
  }
  return { ok: false, error: 'gemini_all_models_failed' };
}

// ── DeepSeek (OpenAI-compatible API at api.deepseek.com/v1).
// Tool-calling format matches OpenAI exactly — reuse AI_TOOLS_SPEC + dispatchAiTool.
// Allowed override values for per-request DeepSeek model selection.
const DEEPSEEK_MODEL_ALLOWLIST = new Set([
  'deepseek-v4-pro',
  'deepseek-v4-flash',
  'deepseek-chat',
  'deepseek-reasoner',
]);
function _resolveDeepSeekModel(override) {
  if (typeof override === 'string') {
    const v = override.trim().toLowerCase();
    if (v === 'flash') return 'deepseek-v4-flash';
    if (v === 'pro') return 'deepseek-v4-pro';
    if (DEEPSEEK_MODEL_ALLOWLIST.has(v)) return v;
  }
  return DEEPSEEK_MODEL;
}
async function callDeepSeekChatTools(messages, tools, modelOverride) {
  if (!DEEPSEEK_KEY) return { ok: false, error: 'deepseek_not_configured' };
  const body = {
    model: _resolveDeepSeekModel(modelOverride),
    messages,
    max_tokens: 1024,
    temperature: 0.3,
  };
  if (tools && tools.length) {
    body.tools = tools;
    body.tool_choice = 'auto';
  }
  const payload = JSON.stringify(body);
  return new Promise((resolve) => {
    const r = https.request({
      hostname: 'api.deepseek.com',
      path: '/v1/chat/completions',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + DEEPSEEK_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    }, (resp) => {
      let buf = '';
      resp.on('data', (d) => (buf += d));
      resp.on('end', () => {
        try {
          const j = JSON.parse(buf);
          if (resp.statusCode >= 200 && resp.statusCode < 300) {
            const msg = j?.choices?.[0]?.message;
            if (!msg) return resolve({ ok: false, error: 'deepseek_empty', raw: buf.slice(0, 200) });
            return resolve({ ok: true, message: msg, text: msg.content || '' });
          }
          return resolve({ ok: false, error: 'deepseek_http_' + resp.statusCode, raw: buf.slice(0, 240) });
        } catch (e) {
          resolve({ ok: false, error: 'deepseek_parse: ' + e.message });
        }
      });
    });
    r.on('error', (e) => resolve({ ok: false, error: 'deepseek_net: ' + e.message }));
    r.setTimeout(45000, () => r.destroy(new Error('deepseek_timeout')));
    r.write(payload);
    r.end();
  });
}

async function runDeepSeekWithTools(messages, session, maxRounds, onProgress, modelOverride) {
  const convo = messages.slice();
  const max = maxRounds || 12;
  const emit = typeof onProgress === 'function' ? onProgress : null;
  const callSig = new Map();
  for (let round = 0; round < max; round++) {
    if (emit) emit({ type: 'thinking', round: round + 1 });
    let r = await callDeepSeekChatTools(convo, AI_TOOLS_SPEC, modelOverride);
    if (!r.ok && r.error === 'deepseek_http_400' && round > 0) {
      // Retry once after dropping the assistant's reasoning_content + tool_calls
      // and any unmatched tool turns. v4-pro occasionally rejects the echoed
      // chain on the next round — a fresh invocation usually recovers without
      // having to fail the whole stream.
      const rawSnip = r.raw ? ` raw=${String(r.raw).replace(/\s+/g, ' ').slice(0, 240)}` : '';
      console.error(`[deepseek-tools] round ${round + 1}/${max} 400; retrying without reasoning_content${rawSnip}`);
      const cleaned = convo.map(m => {
        if (m && m.role === 'assistant') {
          const c = { role: 'assistant', content: m.content == null ? '' : m.content };
          if (m.tool_calls) c.tool_calls = m.tool_calls;
          return c;
        }
        return m;
      });
      r = await callDeepSeekChatTools(cleaned, AI_TOOLS_SPEC, modelOverride);
    }
    if (!r.ok) {
      const rawSnip = r.raw ? ` raw=${String(r.raw).replace(/\s+/g, ' ').slice(0, 240)}` : '';
      console.error(`[deepseek-tools] round ${round + 1}/${max} call failed: ${r.error}${rawSnip}`);
      return r;
    }
    const msg = r.message || {};
    const calls = Array.isArray(msg.tool_calls) ? msg.tool_calls : null;
    if (!calls || !calls.length) {
      console.log(`[deepseek-tools] round ${round + 1}/${max} done (text len ${(msg.content || '').length})`);
      return { ok: true, text: normalizeKannadaNumbersForAi(msg.content || '') };
    }
    console.log(`[deepseek-tools] round ${round + 1}/${max} calling ${calls.length} tool(s): ${calls.map(c => c?.function?.name).join(', ')}`);
    if (emit) emit({ type: 'tools', names: calls.map(c => c?.function?.name).filter(Boolean) });
    const dsAssistantTurn = {
      role: 'assistant',
      content: msg.content == null ? null : msg.content,
      tool_calls: calls,
    };
    // DeepSeek v4-pro (reasoning mode) rejects the next round with
    // "The reasoning_content in the thinking mode must be passed back to the API."
    // unless we echo reasoning_content on the assistant turn.
    if (msg.reasoning_content) dsAssistantTurn.reasoning_content = msg.reasoning_content;
    convo.push(dsAssistantTurn);
    for (const c of calls) {
      const fnName = c?.function?.name || '';
      let argObj = {};
      try { argObj = JSON.parse(c?.function?.arguments || '{}'); } catch (_) {}
      const sig = fnName + '::' + JSON.stringify(argObj);
      const prev = callSig.get(sig) || 0;
      let toolResult;
      if (prev >= 1) {
        toolResult = { error: 'duplicate_call_blocked', instruction: 'You already called this tool with these exact arguments. Do NOT retry. Either change the arguments or answer the user with what you have. Do not loop.' };
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: [], result: toolResult });
      } else {
        const sqlLog = [];
        try {
          await aiToolLogStore.run(sqlLog, async () => {
            toolResult = await dispatchAiTool(fnName, argObj, session);
          });
        } catch (e) {
          toolResult = { error: 'tool_threw: ' + e.message };
        }
        if (emit) emit({ type: 'tool_result', name: fnName, args: argObj, queries: sqlLog, result: toolResult });
      }
      callSig.set(sig, prev + 1);
      convo.push({
        role: 'tool',
        tool_call_id: c.id,
        name: fnName,
        content: JSON.stringify(toolResult).slice(0, 24000),
      });
    }
  }
  return { ok: false, error: 'deepseek_tool_loop_exceeded' };
}

// GET /api/ai-providers — frontend uses this to render provider chooser
// cards and disable the ones whose env keys are not configured.
app.get('/api/ai-providers', (_req, res) => {
  const all = ['openai', 'mistral', 'gemini', 'deepseek'];
  const haveKey = { openai: !!OPENAI_KEY, mistral: !!MISTRAL_KEY, gemini: !!GEMINI_KEY, deepseek: !!DEEPSEEK_KEY };
  const available = all.filter((p) => haveKey[p]);
  const unavailable = all.filter((p) => !haveKey[p]);
  res.json({ available, unavailable });
});

// /api/ai-activity — real upload/refresh timestamps for the inspector activity
// feed. Uses report_date as the upload proxy (no created_at on these tables);
// also returns row counts so the UI can show "248 branches".
app.get('/api/ai-activity', async (_req, res) => {
  try {
    const [daily, empPerf, dailyReports, dra, branches] = await Promise.all([
      pool.query("SELECT MAX(report_date) AS d FROM daily_performance"),
      pool.query("SELECT MAX(report_date) AS d FROM employee_performance"),
      pool.query("SELECT MAX(date) AS d FROM daily_reports"),
      pool.query("SELECT MAX(date) AS d FROM daily_reports_achievements"),
      pool.query("SELECT COUNT(*)::int AS n FROM branches"),
    ]);
    res.json({
      daily_collection_date: daily.rows[0]?.d || null,
      employee_perf_date:    empPerf.rows[0]?.d || null,
      daily_reports_date:    dailyReports.rows[0]?.d || null,
      achievements_date:     dra.rows[0]?.d || null,
      branch_count:          branches.rows[0]?.n || 0,
      server_now:            new Date().toISOString(),
    });
  } catch (err) {
    console.error('ai-activity error:', err);
    res.status(500).json({ error: err.message });
  }
});

app.post("/api/ai-chat", aiLimiter, requireAiAccess, async (req, res) => {
  // Origin check: allow same-origin web traffic, plus mobile clients which
  // send no Origin/Referer but advertise themselves via x-app-origin.
  const origin = req.headers.origin || req.headers.referer || '';
  const mobileOrigin = String(req.headers['x-app-origin'] || '').toLowerCase();
  if (origin) {
    if (!origin.includes('navachetanalivelihoods.com') && !origin.includes('localhost')) {
      return res.status(403).json({ error: 'Forbidden' });
    }
  } else if (mobileOrigin !== 'nlpl-mobile') {
    // No Origin header AND no mobile marker — likely cURL / scraper.
    return res.status(403).json({ error: 'Forbidden' });
  }

  const { messages, role, location, deepseek_model: deepseekModelOverride } = req.body || {};
  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ error: 'messages array required' });
  }

  // Provider selection: 'mistral' | 'gemini' | 'openai' | 'deepseek'.
  // Default = mistral. Fallback chain runs others on failure.
  const requested = String((req.body && req.body.provider) || 'mistral').toLowerCase();
  const validProviders = ['mistral', 'gemini', 'openai', 'deepseek'];
  const primary = validProviders.includes(requested) ? requested : 'mistral';
  const fallbackOrder = ['mistral', 'gemini', 'openai', 'deepseek'].filter(p => p !== primary);
  const fallback = fallbackOrder[0];

  if (!MISTRAL_KEY && !GEMINI_KEY && !OPENAI_KEY && !DEEPSEEK_KEY) {
    return res.status(503).json({ error: 'AI not configured.' });
  }

  // Fetch role-scoped data and inject into system prompt.
  // session is hoisted because the Mistral tool dispatcher (downstream of this
  // try/catch) needs it to scope every DB tool call.
  const session = (role && location) ? { role, location } : {};
  let scopedSystemText = '';
  try {
    const ctx = await buildDataContext(session);
    const scopeLabel = (role && location) ? `${role} for ${location}` : 'CEO (all data)';
    // Compact JSON to keep payload well under the ~32k Mistral token budget.
    const ctxJson = JSON.stringify(ctx);
    scopedSystemText = [
      '',
      '',
      `You are the NLPL Dashboard AI Assistant for ${scopeLabel}.`,
      `Today is ${ctx.now}. Current FY started ${ctx.fyStart} (April 1 → March 31).`,
      '',
      '## Language',
      '- Detect the language of the latest user message and REPLY IN THE SAME LANGUAGE.',
      '- If the user wrote in Kannada (ಕನ್ನಡ script), reply in Kannada. If in English, reply in English.',
      '- If the user mixed both, follow the dominant language.',
      '- **DATABASE IS ENGLISH-ONLY.** All identifiers (employee names, branch names, regions) are stored in Latin script. When the user names an entity in Kannada, transliterate it to English BEFORE calling any tool: ರಘುನಂದನ್ → "Raghunandan", ದಾವಣಗೆರೆ → "Davanagere", ಬೆಂಗಳೂರು → "Bangalore", ಕಲಬುರಗಿ → "Kalaburagi". Searching the DB with Kannada script returns 0 rows. The reply text stays in Kannada.',
      '- In Kannada replies, ALL numbers, percentages, dates, currency values, and number unit words stay in English digits/English words: "12.34 crore", "5.6 lakh", "45%", "2026-05-01". Do not use Kannada digits or Kannada number words.',
      '- Technical column names (regular_demand, FTOD, NPA, etc.) stay in English even in Kannada replies.',
      '',
      '## Database tables you can query (via tools below)',
      '- branches, employees, employee_master (HR roster — hierarchy: region_name → division_name → area_name → branch_name).',
      '- employee_performance — current rolling totals per employee (no date column).',
      '- daily_performance — historical daily snapshots per employee. Has report_date.',
      '- disbursement / disbursement_daily — monthly + daily loan disbursement (db_month, branch_name, disb_count, disb_amount).',
      '- npa_activation_runs / npa_activation_rows — uploaded NPA action sheets.',
      '- daily_reports / daily_reports_achievements — branch-level daily PLAN + ACHIEVEMENT (per branch per date). Columns: ftod_actual/plan, dpd_1_30_actual/plan, dpd_31_60_actual/plan, dpd_61_90_actual/plan, npa_activation, npa_closure, fy_non_start_acc/plan, disb_igl_acc/amt, disb_fig_acc/amt, disb_il_acc/amt, kyc_igl, kyc_fig, kyc_il.',
      '- hourly_performance — intra-day collection snapshots.',
      '',
      '## Scope',
      'Data is scoped to the user\'s role: CEO=all branches; RM/SM=region; DM/DvM=division; AM=area; BM/ABM/BOE=their own branch only.',
      'Every tool call is automatically scope-filtered. **Do NOT add branch_name / region_name / area_name / division_name filters when the user asks about their own scope** — the server adds them automatically. For Branch Managers, never pass branch_name to tools (they have only one branch). If a tool returns `error: "scope_violation"`, relay the message verbatim to the user — do not retry.',
      '',
      '## Units — READ CAREFULLY',
      '**COUNT columns (integers, NOT money — never format as ₹/L/Cr):**',
      '  regular_demand, regular_collection, demand_1_30, collection_1_30, demand_31_60, ..., npa_cases, npa_act_acc, npa_clo_acc, disb_count, ftod, kyc_*, fy_non_start_acc.',
      '**MONETARY columns (rupees — divide by 1,00,000 for L or 1,00,00,000 for Cr):**',
      '  regular_demand_amt, regular_collection_amt, demand_*_amt, npa_act_amt, npa_clo_amt, disb_amount, disb_*_amt.',
      'Rule of thumb: any column ending in `_amt` or `disb_amount` is money. Everything else is a count.',
      'Collection % = regular_collection / regular_demand × 100 (count ratio, dimensionless).',
      '',
      '## Tools available — CALL THESE for any specifics not in the at-a-glance JSON',
      '- resolve_date_range(expression, comparison_expression?, anchor_date?) — converts relative/spoken/slash dates to ISO YYYY-MM-DD using Asia/Kolkata. Use before data tools for today/yesterday/this month/last month/FYTD/11/04/April 11/comparison dates.',
      '- find_employee(query, limit?) — name/mobile/emp_id substring lookup.',
      '- employee_performance(emp_id, start_date?, end_date?) — ONE employee\'s aggregate TOTALS (single row) over a period (default = FY-to-date). Use for "<person> totals / <person> performance summary".',
      '- employee_collection_series(emp_id, start_date?, end_date?) — ONE employee\'s per-DAY series (one row per report_date with demand/collection/collection_pct/npa). Use for "show <person> trend / <person> day-by-day / drill on <person>". Default window = today minus 30 days. NEVER substitute period_performance(group_by="employee", branch_name=...) or top_performers(branch_name=...) for this — those return the whole branch leaderboard, not the named person.',
      '- find_branch(query, limit?) — branch name lookup with hierarchy + headcount + perf.',
      '- period_performance(start_date, end_date, group_by?, branch_name?, area_name?, division_name?, region_name?, limit?) — collection/demand/NPA over a date range, group by day|month|branch|employee.',
      '- top_performers(metric, start_date, end_date, limit?) — leaderboard by collection|demand|npa_cases.',
      '- disbursement_query(start_date, end_date, group_by?, branch_name?, region_name?, limit?) — disbursement count + amount over range, group by day|month|branch|product|employee. **Use group_by="day" for date-specific questions ("April 7th vs 8th")** so each row is one calendar day; "month" rolls up to month buckets.',
      '- list_hierarchy(level, parent_level?, parent_name?, limit?) — list region|division|area|branch entities, optionally under a parent.',
      '- npa_summary(start_date, end_date, group_by?, branch_name?, region_name?, limit?) — NPA cases + activation amount over range.',
      '- daily_reports_query(start_date, end_date, branch_name?, region_name?, district_name?, table?, metrics?, limit?) — branch-level DAILY PLAN data: FTOD, DPD bucket 1-30/31-60/61-90, NPA activation/closure, disbursement plan vs actual (IGL/FIG/IL), KYC. ALWAYS use this for ANY question mentioning FTOD, DPD, KYC, NPA closure, or disbursement plan-vs-achievement on a specific date or branch.',
      '- branch_summary(branch_name?, date?) — ONE-SHOT branch dashboard: headcount + role mix, collection today/MTD/FYTD with %, NPA, disbursement today/MTD, FTOD plan-vs-actual, DPD bucket plan-vs-actual. Use for "how is my branch doing today / branch summary / health check / EOD rollup". BM/ABM/BOE pass NO args. CEO/RM/DM/AM MUST pass branch_name (single-branch only).',
      '- hourly_collection(branch_name?, region_name?, group_by?) — CURRENT live intra-day collection snapshot (counts + amounts + pct). hourly_performance has no time series, so this returns the latest live totals, NOT an hour-by-hour series.',
      '- collection_drilldown(branch_name?, date?) — Single-branch root-cause drilldown for "why is collection low". Returns top_3_underperformers (3 worst FOs) + bottom_3_underperformers (3 best FOs) + DPD buckets + NPA + FTOD gap. Use after branch_summary.',
      '- period_compare(metric, scope, scope_value?, period_a_start, period_a_end, period_b_start, period_b_end) — Two-window compare with server-computed delta_abs + delta_pct. metric ∈ {collection, demand, collection_pct, npa_amount, disb_amount, disb_count, ftod}. Use for any MoM / WoW / vs-last-month / vs-yesterday question.',
      '- plan_compliance(date?, scope?, scope_value?) — Lists branches that did NOT file daily_reports / daily_reports_achievements for a given date. CEO/RM/DM/AM only.',
      '- sql_describe() — schema cheatsheet refresher.',
      '',
      '## Date range guidance',
      '- For any relative, spoken, slash-format, or comparison date phrase, call `resolve_date_range` first. Then pass only its ISO start_date/end_date values to downstream tools.',
      `- "today" → ${ctx.now}.`,
      `- "this month" → first of current month → ${ctx.now}.`,
      '- "last month" → first to last of prior month.',
      `- "FY-to-date" / "this year" → ${ctx.fyStart} → ${ctx.now}.`,
      '- "last quarter" → prior 3 calendar months.',
      `- A bare month name ("July") → that month in the current FY year (FY runs Apr→Mar). FY start = ${ctx.fyStart}.`,
      '',
      '## Structured answer format (REQUIRED whenever the answer contains numbers)',
      '- ANY answer that yields a list, ranking, leaderboard, per-entity snapshot, comparison, or DIAGNOSTIC drill-down (e.g. "branches below 90% plan attainment", "top 5 FOs", "branches that missed FTOD", "why is <branch> below plan", "drill down on <branch>", "what is happening with <branch>") MUST be presented as a Markdown pipe table with a header row + a separator row (`|---|---|`). Plain bullet lists of metrics + values are forbidden when ≥2 metrics or ≥2 entities are involved.',
      '- For DIAGNOSTIC ("why" / drill-down) answers about ONE entity with several metrics, use a metric-by-metric table. Example for "why is Davanagere below plan": columns `Metric | Plan | Actual | % | Status` with rows like `Collection (Today) | 237 | 178 | 75.1% | Below`, `Collection (MTD) | 521 | 390 | 74.9% | Below`, `NPA Cases | — | 119 | — | —`, `FTOD | 10 | 15 | 150% | Above`. NEVER answer a "why" question as a nested-bullet narrative when a table can carry the metrics.',
      '- A list answer is NEVER allowed to be a single column of names. Include AT LEAST the entity column AND its primary metric column. Example: "branches below 90% plan attainment" → columns `Branch | Plan Target | Actual | Attainment %`. "FOs with low collection" → `FO | Branch | Demand | Collection | Coll %`.',
      '- Right-align numeric columns in the Markdown separator row using `| ---: |` (e.g. `| Branch | Plan Target | Actual | Attainment % |` then `| :--- | ---: | ---: | ---: |`). Left-align text columns with `| :--- |`.',
      '- Output ONLY ONE structured representation per answer. Do NOT also paste a plain-text/ASCII "table" or duplicate the same data as a bullet list — pick the Markdown pipe table and stop.',
      '- After the table, add a short `**Calculation**` section with 1–3 bullets that show the formula AND the raw inputs. Example: "Plan attainment = Actual ÷ Plan × 100" then "Afzalpur: ₹4.2L ÷ ₹5.0L × 100 = 84%". The reader must be able to reproduce the metric from the bullets.',
      '',
      '## How to answer',
      '- Lead with the headline number, then 2-4 supporting bullets.',
      '- Indian number format: lakh comma pattern 12,34,56,789. Prefer "₹X.XX Cr" (1 Cr = 1,00,00,000) or "₹X.XX L" (1 L = 1,00,000) — never raw rupees with broken commas.',
      '- If the user asks only for collection, compare only demand, collection, and collection %. Do not add NPA, disbursement, or other metrics unless asked.',
      '- Label count metrics as counts/accounts/cases and format with plain Indian commas, for example "30,340 cases", not "30.34 L".',
      '- For ANY month-vs-month / period-vs-period / branch-vs-branch comparison, ALWAYS call `period_performance` with explicit start/end dates for each side and compute % change yourself. Do NOT eyeball the at-a-glance monthly array — collection (`monthly[].total_collection`) and disbursement (`disbursementMonthly[].amount`) are different metrics and have been mis-substituted before.',
      '- For ANY specific question (a named employee, a named branch, a date range not pre-aggregated, a comparison, a leaderboard, a hierarchy listing), CALL THE RIGHT TOOL.',
      '- The at-a-glance JSON is fine for trivial single-number lookups in the current period. For everything else, call a tool.',
      '- If the JSON doesn\'t have it, call a tool. Do NOT refuse — the tools cover virtually any DB question in scope.',
      '- For ambiguous questions, ask ONE crisp clarifying question.',
      '- For employee-name lookups: if the user mentioned a branch / area / district / region with the name ("Shivraj from Raichur"), pass it as location_hint to find_employee. If find_employee returns more than one match or `ambiguous: true`, list the matching employees with emp_id + branch/role and ask which one the user means. Never pick one silently for common names like Karthik. Do not list phone numbers while disambiguating multiple people. If find_employee returns 0 rows, do NOT fall back to overall company numbers — tell the user the name didn\'t match and ask them to repeat it.',
      '- For branch/entity lookups: if find_branch returns more than one match or `ambiguous: true`, list the matching branches with region/division/area and ask which one the user means. Never pick one silently for partial branch names.',
      '- Never invent numbers. Quote what tools return.',
      '- **Money formatting — DO NOT do the division yourself.** disbursement_query returns each row with `amount_str` (already formatted, e.g. "₹178.96 Cr") and the wrapper has `totals.amount_str`. Copy `amount_str` VERBATIM into your reply. Do NOT divide `amount` by 10000000 yourself — past replies dropped a zero (₹178.96 Cr → ₹17.90 Cr) and gave wrong totals. If `amount_str` is present, use it exactly as-is. The total row in your table MUST equal `totals.amount_str`, not a re-summed value.',
      '- Never use the words "snapshot" or "provided data" or "JSON below" in your replies — they leak internal plumbing. Just answer with the numbers and a short label.',
      '- If a question mentions FTOD, DPD, KYC, disbursement plan, NPA closure, or "daily plan" → MUST call daily_reports_query. Do NOT say "data not available" without calling the tool first.',
      '- If the user asks "show <person> trend / <person> performance over time / drill on <person> / day-by-day for <person>" → call `find_employee(query=<name>, location_hint?)` to canonicalise, then `employee_collection_series(emp_id, start_date?, end_date?)` for the per-day series. Use `employee_performance(emp_id, …)` only for the aggregate TOTAL across a window. NEVER call `period_performance(group_by="employee", branch_name=...)` or `top_performers(branch_name=...)` to answer a single-named-person question — those return the FULL branch leaderboard, which is what the user did NOT ask for.',
      '- If the user asks "how is my branch doing today / branch summary / give me a summary of <branch> / end-of-day rollup" → call `branch_summary()` (single tool, returns headcount + collection + NPA + disbursement + FTOD + DPD). Do NOT chain 5 separate tool calls.',
      '- If the user asks "why is collection low / drill down / who\'s underperforming / which FOs are dragging us down" → call `collection_drilldown()`. Returns the 3 worst + 3 best FOs + DPD + NPA + FTOD gap in one call. Use AFTER branch_summary for the natural "why" follow-up.',
      '- For any "MoM / WoW / vs last month / vs yesterday / this week vs last week / compare X to Y" question → call `period_compare(metric, scope, scope_value?, period_a_start, period_a_end, period_b_start, period_b_end)`. ALL 4 dates ISO YYYY-MM-DD; metric ∈ {collection, demand, collection_pct, npa_amount, disb_amount, disb_count, ftod}; scope ∈ {all, branch, region, division, area}. Server computes delta_abs + delta_pct — DO NOT call period_performance twice and do the arithmetic yourself.',
      '- For "which branches missed plan today / didn\'t file / plan compliance" → call `plan_compliance(date?, scope?, scope_value?)`. date ISO YYYY-MM-DD (defaults to today). Returns missing_plan + missing_achievement.',
      '- If the user asks "collection right now / live / hourly / intra-day / how much collected today so far" → call `hourly_collection()`. It is a live snapshot, not a time series — answer with the current totals and don\'t pretend to break them down by hour.',
      '- "11th April" / "April 11" / "11/04" all mean the same date — convert to YYYY-MM-DD using the current FY year.',
      '- When a tool returns 0 rows for a specific date, ALWAYS retry the same tool with a wider window (the full month, then the full FY) to find the nearest available date(s). Then tell the user "no data for {requested date}; nearest available is {date} — here it is" and answer with that data. Do NOT just say "data not available" and stop.',
      '',
      '## Hard rules — DO NOT BREAK',
      '- **Single-entity scope.** When the user names ONE person, branch, or region in the question (e.g. "show Manjunatha T N trend", "how is Davanagere doing", "drill on Pavan Kumar K"), the headline answer is scoped to THAT entity ONLY. For a single named PERSON: call `find_employee` → then `employee_collection_series(emp_id, …)` for trend/series, or `employee_performance(emp_id, …)` for totals. NEVER call `period_performance(group_by="employee", branch_name=…)` or `top_performers(branch_name=…)` to answer a single-person question — those return the whole branch leaderboard. For a single named BRANCH: use `branch_summary` / `collection_drilldown`. Listing peers (the rest of the branch, the rest of the region) is allowed ONLY when the user explicitly asks for a ranking, comparison, or "everyone in <branch>".',
      '- **top_performers ranks EMPLOYEES, not branches.** For "top N branches by collection / demand / disbursement / NPA", call `period_performance(group_by=\'branch\', start_date, end_date)` and sort/pick top N yourself. NEVER label disbursement_query rows as collection (or vice versa).',
      '- **Never relabel one tool\'s output as a different metric.** If the user asked for collection and you only fetched disbursement, fetch collection — do not rename columns.',
      '- **Never fabricate totals.** When summarising tool rows, your reply must either (a) sum ALL returned rows accurately, or (b) explicitly say "showing first N of M — total is X" using the actual sum. Never invent a total that doesn\'t match the rows shown.',
      '- **find_branch / find_employee FIRST when user names an entity.** Use the canonical branch_name / emp_id from the result in every downstream tool. The server also canonicalizes branch_name, but you must still resolve ambiguous branches before answering.',
      '- **Cross-branch compare permission depends on the SESSION role above.** CEO / RM / SM / DM / DvM / AM may compare ANY two branches freely — call period_compare with scope="branch" and the two scope_values. Branch-bound roles (BM / ABM / BOE) may ONLY compare their own branch; if the user is BM/ABM/BOE in the session and they ask to compare against any OTHER branch, refuse with: "You can only access <own branch>. Cross-branch comparisons require RM or CEO access." NEVER apply the BM refusal when the session role is CEO/RM/SM/DM/DvM/AM. NEVER call period_compare with own-branch as scope_value for both windows — that silently produces a fake-parity comparison.',
      '- **regular_demand / regular_collection are COUNTS, not money.** Even when narrating branch_summary or period_compare results, format these as plain Indian-comma numbers (e.g. "1,335 collection accounts"), NOT ₹ / L / Cr. Only columns ending in `_amt` and {disb_amount, npa_act_amt, npa_amount, npa_clo_amt} are money.',
      '- **Multi-employee period queries: ONE call, not N.** When the user asks for a LIST or TABLE of multiple employees with their performance over a period (e.g. "every FO in <branch> with their April collection", "show all my staff and their collection"), call `top_performers(metric=\'collection\', branch_name=<scope>, role=<role-if-named>, start_date, end_date, limit=200)` ONCE — top_performers ALSO accepts role/branch/area/division/region filters and returns the full leaderboard. Or `period_performance(group_by=\'employee\', branch_name=<scope>, start_date, end_date)` if no role filter. NEVER call `list_employees` then loop `employee_performance(emp_id=...)` per person — that anti-pattern wastes round-trips and fragments the output. Reserve `employee_performance` for a SINGLE named individual. Multi-axis requests (per-emp × per-day × per-product × per-DPD-bucket) — pick ONE primary axis, call once, and tell the user the secondary axis needs a follow-up question.',
      '- **Active-filter preamble is BINDING SCOPE, not narrative.** If the latest user message starts with `[Active filter: ...]`, parse the bracket and apply it to every tool call until the next user message overrides it. Three shapes: (a) SINGLE — `[Active filter: employee NL11007 (R Gagan) at Ajjampura branch.]` → pass `emp_id="NL11007"` to employee_performance and pin `branch_name="Ajjampura"` for any branch-scoped follow-up. (b) SET with explicit emp_ids — `[Active filter: 8 employees in Ajjampura branch — emp_ids: NL11007, NL12292, ...]` → call ONE branch-scoped tool (`top_performers(branch_name="Ajjampura", role?, start, end)` or `period_performance(group_by="employee", branch_name="Ajjampura", start, end)`) and IN YOUR REPLY only narrate the rows whose emp_id appears in the bracket list (filter client-side). NEVER loop employee_performance per emp_id. (c) SET-ALL — `[Active filter: all employees in Ajjampura branch.]` → branch-scoped tool ONCE, narrate all rows. Do NOT echo the bracket back in your reply — it is a routing directive, not user-visible content. Reply only to the question that follows the bracket.',
      '',
      '## Internal context block (DO NOT mention this in your reply — quote numbers naturally)',
      ctxJson,
    ].join('\n');
  } catch (e) {
    console.error('AI chat scope fetch error:', e.message);
    // non-fatal: proceed without scoped data
  }

  // Build a single OpenAI-format message list with the scoped system prompt
  // merged in. Both Mistral and the Gemini adapter understand this shape.
  const incomingSystem = (messages.find(m => m.role === 'system') || {}).content || '';
  const mergedSystem = (incomingSystem || '') + (scopedSystemText || '');
  const nonSystem = messages.filter(m => m.role !== 'system');
  const mergedMessages = mergedSystem
    ? [{ role: 'system', content: mergedSystem }, ...nonSystem]
    : nonSystem;

  // Cache lookup — provider-aware so a Mistral reply doesn't shadow a Gemini one.
  const lastUserMsg = nonSystem.filter(m => m.role === 'user').slice(-1)[0]?.content || '';
  const cacheKey = getCacheKey(role, location, lastUserMsg, primary);
  pruneCache();
  const cached = aiReplyCache.get(cacheKey);
  if (cached && (Date.now() - cached.ts < AI_CACHE_TTL_MS)) {
    return res.json({
      reply: cached.reply,
      provider: cached.provider,
      cached: true,
    });
  }

  // Try primary, then fallbacks (mistral → gemini → openai → deepseek minus primary).
  const order = [primary, ...fallbackOrder].filter((p, i, a) => a.indexOf(p) === i);
  let lastErr = null;
  for (const which of order) {
    let result;
    if (which === 'mistral') {
      if (!MISTRAL_KEY) { lastErr = 'mistral_not_configured'; continue; }
      result = await runMistralWithTools(mergedMessages, session);
    } else if (which === 'gemini') {
      if (!GEMINI_KEY) { lastErr = 'gemini_not_configured'; continue; }
      result = await callGeminiAi(mergedMessages);
    } else if (which === 'openai') {
      if (!OPENAI_KEY) { lastErr = 'openai_not_configured'; continue; }
      result = await runOpenAiWithTools(mergedMessages, session);
    } else if (which === 'deepseek') {
      if (!DEEPSEEK_KEY) { lastErr = 'deepseek_not_configured'; continue; }
      result = await runDeepSeekWithTools(mergedMessages, session, undefined, undefined, deepseekModelOverride);
    }
    if (result && result.ok && result.text) {
      const providerLabel = which === 'mistral' ? MISTRAL_MODEL
        : which === 'openai' ? OPENAI_LLM
        : which === 'deepseek' ? DEEPSEEK_MODEL
        : (result.model || 'gemini');
      aiReplyCache.set(cacheKey, {
        reply: result.text,
        provider: providerLabel,
        ts: Date.now(),
      });
      return res.json({
        reply: result.text,
        provider: which,
        model: providerLabel,
        ...(which !== primary ? { fallback: true } : {}),
      });
    }
    lastErr = (result && result.error) || lastErr;
    console.error('AI provider ' + which + ' failed:', lastErr);
  }

  res.status(429).json({
    error: 'AI is briefly busy. Please retry in a moment.',
    detail: lastErr,
  });
});

// ── /api/ai-chat-stream — SSE wrapper around the same tool loop. ──────────
// Emits progress events while tools run, then streams the final reply text
// in word-sized chunks for a typing-style UX. Same auth/origin checks as
// /api/ai-chat. Falls back to Gemini if Mistral fails.
app.post("/api/ai-chat-stream", aiLimiter, requireAiAccess, async (req, res) => {
  const origin = req.headers.origin || req.headers.referer || '';
  const mobileOrigin = String(req.headers['x-app-origin'] || '').toLowerCase();
  if (origin) {
    if (!origin.includes('navachetanalivelihoods.com') && !origin.includes('localhost')) {
      return res.status(403).json({ error: 'Forbidden' });
    }
  } else if (mobileOrigin !== 'nlpl-mobile') {
    return res.status(403).json({ error: 'Forbidden' });
  }

  const { messages, role, location, deepseek_model: deepseekModelOverride } = req.body || {};
  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ error: 'messages array required' });
  }

  if (!MISTRAL_KEY && !GEMINI_KEY && !OPENAI_KEY && !DEEPSEEK_KEY) {
    return res.status(503).json({ error: 'AI not configured.' });
  }

  // SSE setup
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders && res.flushHeaders();

  // Disable Nagle so tiny writes flush immediately.
  try { req.socket.setNoDelay(true); } catch (_) {}

  let closed = false;
  let heartbeat = null;
  const stopHeartbeat = () => {
    if (heartbeat) { clearInterval(heartbeat); heartbeat = null; }
  };
  const finish = () => {
    closed = true;
    stopHeartbeat();
    try { res.end(); } catch (_) {}
  };
  // NOTE: do NOT listen on req.on('close') — for multipart uploads, the
  // request stream emits 'close' as soon as multer finishes parsing the
  // body, well before the response is finalised. That would mis-flag the
  // connection as closed and silence every subsequent SSE write.
  // res.on('close') is the correct "client disconnected" signal here.
  res.on('close', () => { closed = true; stopHeartbeat(); });

  const send = (event, data) => {
    if (closed || res.writableEnded) return;
    try {
      res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
    } catch (_) {}
  };
  // Padded heartbeat to defeat intermediate buffers (Apache, browser).
  const HB_PAD = ': ' + 'x'.repeat(2048) + '\n\n';
  heartbeat = setInterval(() => {
    if (closed || res.writableEnded) { stopHeartbeat(); return; }
    try { res.write(HB_PAD); } catch (_) { stopHeartbeat(); }
  }, 250);

  const session = (role && location) ? { role, location } : {};
  let scopedSystemText = '';
  try {
    const ctx = await buildDataContext(session);
    const scopeLabel = (role && location) ? `${role} for ${location}` : 'CEO (all data)';
    // Trim ctxJson — full ctx is ~36k tokens for CEO, exceeds tier-1 30k
    // TPM. Tools fetch specifics; assistant gets summary + headline
    // numbers up front.
    const ctxLite = {
      scope: ctx.scope,
      latestDate: ctx.latestDate,
      fyStart: ctx.fyStart,
      now: ctx.now,
      latest: ctx.latest,
      summary_text: ctx.summary_text,
      monthly: Array.isArray(ctx.monthly) ? ctx.monthly.slice(-6) : undefined,
      disbursementMonthly: Array.isArray(ctx.disbursementMonthly) ? ctx.disbursementMonthly.slice(-6) : undefined,
    };
    const ctxJson = JSON.stringify(ctxLite);
    scopedSystemText = [
      '',
      '',
      `You are the NLPL Dashboard AI Assistant for ${scopeLabel}.`,
      `Today is ${ctx.now}. Current FY started ${ctx.fyStart} (April 1 → March 31).`,
      '',
      '## Language',
      '- Detect the language of the latest user message and REPLY IN THE SAME LANGUAGE.',
      '- If the user wrote in Kannada (ಕನ್ನಡ script), reply in Kannada. If in English, reply in English.',
      '- If the user mixed both, follow the dominant language.',
      '- **DATABASE IS ENGLISH-ONLY.** All identifiers (employee names, branch names, regions) are stored in Latin script. When the user names an entity in Kannada, transliterate it to English BEFORE calling any tool: ರಘುನಂದನ್ → "Raghunandan", ದಾವಣಗೆರೆ → "Davanagere", ಬೆಂಗಳೂರು → "Bangalore", ಕಲಬುರಗಿ → "Kalaburagi". Searching the DB with Kannada script returns 0 rows. The reply text stays in Kannada.',
      '- In Kannada replies, ALL numbers, percentages, dates, currency values, and number unit words stay in English digits/English words: "12.34 crore", "5.6 lakh", "45%", "2026-05-01". Do not use Kannada digits or Kannada number words.',
      '- Technical column names (regular_demand, FTOD, NPA, etc.) stay in English even in Kannada replies.',
      '',
      '## Database tables you can query (via tools below)',
      '- branches, employees, employee_master (HR roster — hierarchy: region_name → division_name → area_name → branch_name).',
      '- employee_performance — current rolling totals per employee (no date column).',
      '- daily_performance — historical daily snapshots per employee. Has report_date.',
      '- disbursement / disbursement_daily — monthly + daily loan disbursement (db_month, branch_name, disb_count, disb_amount).',
      '- npa_activation_runs / npa_activation_rows — uploaded NPA action sheets.',
      '- daily_reports / daily_reports_achievements — branch-level daily PLAN + ACHIEVEMENT (per branch per date). Columns: ftod_actual/plan, dpd_1_30_actual/plan, dpd_31_60_actual/plan, dpd_61_90_actual/plan, npa_activation, npa_closure, fy_non_start_acc/plan, disb_igl_acc/amt, disb_fig_acc/amt, disb_il_acc/amt, kyc_igl, kyc_fig, kyc_il.',
      '- hourly_performance — intra-day collection snapshots.',
      '',
      '## Scope',
      'Data is scoped to the user\'s role: CEO=all branches; RM/SM=region; DM/DvM=division; AM=area; BM/ABM/BOE=their own branch only.',
      'Every tool call is automatically scope-filtered. **Do NOT add branch_name / region_name / area_name / division_name filters when the user asks about their own scope** — the server adds them automatically. For Branch Managers, never pass branch_name to tools (they have only one branch). If a tool returns `error: "scope_violation"`, relay the message verbatim to the user — do not retry.',
      '',
      '## Date format — REQUIRED',
      `- Today is ${ctx.now}. FY started ${ctx.fyStart}.`,
      '- For any relative, spoken, slash-format, or comparison date phrase, call `resolve_date_range` first. Then pass only its ISO start_date/end_date values to downstream tools.',
      '- ALL date arguments to tools MUST be ISO YYYY-MM-DD. Tools reject "April", "this month", "11/04" etc. with "invalid date format" — convert before the call.',
      `- "today" → ${ctx.now}. "yesterday" → ${ctx.now} minus 1.`,
      `- "this month" → first-of-current-month → ${ctx.now}. "last month" → first-to-last of prior calendar month.`,
      '- "this week" → most recent Monday → today. "last week" → prior Monday → prior Sunday.',
      `- "this year" / "FY-to-date" → ${ctx.fyStart} → ${ctx.now}. FY runs Apr 1 → Mar 31.`,
      '- For period_compare, compute BOTH windows yourself (4 dates) and pass all four. Server returns delta_abs + delta_pct.',
      '',
      '## Units — READ CAREFULLY',
      '**COUNT columns (integers, NOT money — never format as ₹/L/Cr):**',
      '  regular_demand, regular_collection, demand_1_30, collection_1_30, demand_31_60, ..., npa_cases, npa_act_acc, npa_clo_acc, disb_count, ftod, kyc_*, fy_non_start_acc.',
      '**MONETARY columns (rupees — divide by 1,00,000 for L or 1,00,00,000 for Cr):**',
      '  regular_demand_amt, regular_collection_amt, demand_*_amt, npa_act_amt, npa_clo_amt, disb_amount, disb_*_amt.',
      'Rule of thumb: any column ending in `_amt` or `disb_amount` is money. Everything else is a count.',
      'Collection % = regular_collection / regular_demand × 100 (count ratio, dimensionless).',
      '',
      '## Tools available — CALL THESE for any specifics not in the at-a-glance JSON',
      '- resolve_date_range, find_employee, employee_performance, employee_collection_series, find_branch, period_performance, top_performers, disbursement_query, list_hierarchy, list_employees, headcount, npa_summary, daily_reports_query, branch_summary, hourly_collection, collection_drilldown, period_compare, plan_compliance, sql_describe',
      '',
      '## Critical tool routing',
      '- "Show <person> trend / <person> performance over time / drill on <person> / day-by-day for <person> / how is <person> doing day to day" → call `find_employee(query=<name>, location_hint?)` first to canonicalise, then `employee_collection_series(emp_id, start_date?, end_date?)` for the per-day series. Use `employee_performance(emp_id, …)` only for the aggregate TOTAL across a window. NEVER call `period_performance(group_by="employee", branch_name=...)` or `top_performers(branch_name=...)` to answer a single-named-person question — those return the FULL branch leaderboard.',
      '- "How is my branch doing today / branch summary / branch health / give me a summary of <branch> / end-of-day rollup" → call `branch_summary()` — ONE tool returns headcount + collection (today/MTD/FYTD) + NPA + disbursement + FTOD + DPD plan-vs-actual. BM/ABM/BOE pass NO args (auto-resolves); CEO/RM/DM/AM MUST pass branch_name. Don\'t chain 5+ separate tool calls for this.',
      '- "Why is collection low / drill down on collection / who\'s underperforming / which FOs are dragging us down" → call `collection_drilldown()`. Returns 3 worst FOs + 3 best FOs + DPD buckets + NPA + FTOD gap. Use AFTER branch_summary for the natural "why" follow-up. Single-branch only — pass branch_name (CEO/RM/DM/AM) or omit it (BM/ABM/BOE auto-resolve).',
      '- "MoM / WoW / vs last month / vs yesterday / this week vs last week / vs same period last year / compare X to Y" → call `period_compare(metric, scope, scope_value?, period_a_start, period_a_end, period_b_start, period_b_end)`. ALL 4 dates ISO YYYY-MM-DD. metric ∈ {collection, demand, collection_pct, npa_amount, disb_amount, disb_count, ftod}. scope ∈ {all, branch, region, division, area} — use scope="branch" + scope_value=<branch> for "compare <branch> this month vs last month". Server computes delta_abs + delta_pct. NEVER call period_performance twice and do the arithmetic yourself.',
      '- "Which branches missed plan today / who didn\'t file daily report / plan compliance / branches without daily report" → call `plan_compliance(date?, scope?, scope_value?)`. Returns missing_plan + missing_achievement lists. CEO/RM/DM/AM only — BM rejected (single branch). date defaults to today; pass scope="region"+scope_value=<region> to narrow.',
      '- "Collection right now / live / how much collected today so far / hourly / intra-day" → call `hourly_collection()`. Returns CURRENT live totals from hourly_performance — NOT a time series, NOT hour-by-hour buckets.',
      '- "How many employees / staff / BMs / FOs / agents in <scope>?" → call `headcount(group_by=\'role\')` for a role breakdown, or `headcount(role=\'BM\')` for a single role count, or `headcount()` for the overall total. NEVER quote a hierarchy row count or list_hierarchy result count as "number of employees".',
      '- "List all employees / give me everyone / phone numbers of staff in <branch>" / "show roster" → call `list_employees(branch_name=...)`. NOT find_employee — that\'s for searching ONE named person. list_employees returns the full roster (name + role + branch + mobile).',
        '- "Top N employees in <branch> by collection / demand" → call `top_performers(metric, start_date, end_date, branch_name=...)`. NEVER chain list_employees + employee_performance × N — that loses the ranking and produces wrong/zero rows.',
        '- **NPA semantics matter.** `npa_cases` from daily_performance / employee_performance is a ROLLING SNAPSHOT (the same outstanding case is repeated every day). Never SUM npa_cases over a date range — that inflates by ~days. For outstanding cases, query the latest single date. For activation / closure RATES, use `daily_reports_query(table="achievements", metrics="npa")` and SUM npa_activation / npa_closure (those are per-day deltas).',
        '- If today\'s daily_performance is empty (EOD lag — common before evening), retry against `daily_reports_query(table="achievements")` for today\'s achievement numbers before saying "no data".',
      '- "Who is the BM of <branch>" / "who is the manager" → `list_employees(branch_name=X, role=\'BM\')`.',
      '- "Daily collection / collection trend / day-by-day collection" → call `period_performance(group_by=\'day\', start_date, end_date)`. One row per day with demand + collection. Average = SUM(collection) / COUNT(days).',
      '- "Daily disbursement" → call `disbursement_query(group_by=\'day\', start_date, end_date)`.',
      '- list_hierarchy returns entity names + headcount metadata only — NEVER report its row count as a money/collection metric.',
      '',
      '## Hard rules — DO NOT BREAK',
      '- **Single-entity scope.** When the user names ONE person, branch, or region in the question (e.g. "show Manjunatha T N trend", "how is Davanagere doing", "drill on Pavan Kumar K"), the headline answer is scoped to THAT entity ONLY. For a single named PERSON: call `find_employee` → then `employee_collection_series(emp_id, …)` for trend/series, or `employee_performance(emp_id, …)` for totals. NEVER call `period_performance(group_by="employee", branch_name=…)` or `top_performers(branch_name=…)` to answer a single-person question — those return the whole branch leaderboard. For a single named BRANCH: use `branch_summary` / `collection_drilldown`. Listing peers (the rest of the branch, the rest of the region) is allowed ONLY when the user explicitly asks for a ranking, comparison, or "everyone in <branch>".',
      '- **Cross-branch compare permission depends on the SESSION role above.** CEO / RM / SM / DM / DvM / AM may compare ANY two branches freely — call period_compare with scope="branch" and the two scope_values. Branch-bound roles (BM / ABM / BOE) may ONLY compare their own branch; if the user is BM/ABM/BOE in the session and they ask to compare against any OTHER branch, refuse with: "You can only access <own branch>. Cross-branch comparisons require RM or CEO access." NEVER apply the BM refusal when the session role is CEO/RM/SM/DM/DvM/AM. NEVER call period_compare with own-branch as scope_value for both windows — that silently produces a fake-parity comparison.',
      '- **regular_demand / regular_collection are COUNTS, not money.** Format as plain Indian-comma numbers (e.g. "1,335 collection accounts"), NOT ₹/L/Cr. Only columns ending in `_amt` and {disb_amount, npa_act_amt, npa_amount, npa_clo_amt} are money.',
      '- **Multi-employee period queries: ONE call, not N.** When the user asks for a LIST or TABLE of multiple employees with their performance over a period (e.g. "every FO in <branch> with their April collection", "show all my staff and their collection"), call `top_performers(metric=\'collection\', branch_name=<scope>, role=<role-if-named>, start_date, end_date, limit=200)` ONCE — top_performers ALSO accepts role/branch/area/division/region filters and returns the full leaderboard. Or `period_performance(group_by=\'employee\', branch_name=<scope>, start_date, end_date)` if no role filter. NEVER call `list_employees` then loop `employee_performance(emp_id=...)` per person — that anti-pattern wastes round-trips and fragments the output. Reserve `employee_performance` for a SINGLE named individual. Multi-axis requests (per-emp × per-day × per-product × per-DPD-bucket) — pick ONE primary axis, call once, and tell the user the secondary axis needs a follow-up question.',
      '- **Active-filter preamble is BINDING SCOPE, not narrative.** If the latest user message starts with `[Active filter: ...]`, parse the bracket and apply it to every tool call until the next user message overrides it. Three shapes: (a) SINGLE — `[Active filter: employee NL11007 (R Gagan) at Ajjampura branch.]` → pass `emp_id="NL11007"` to employee_performance and pin `branch_name="Ajjampura"` for any branch-scoped follow-up. (b) SET with explicit emp_ids — `[Active filter: 8 employees in Ajjampura branch — emp_ids: NL11007, NL12292, ...]` → call ONE branch-scoped tool (`top_performers(branch_name="Ajjampura", role?, start, end)` or `period_performance(group_by="employee", branch_name="Ajjampura", start, end)`) and IN YOUR REPLY only narrate the rows whose emp_id appears in the bracket list (filter client-side). NEVER loop employee_performance per emp_id. (c) SET-ALL — `[Active filter: all employees in Ajjampura branch.]` → branch-scoped tool ONCE, narrate all rows. Do NOT echo the bracket back in your reply — it is a routing directive, not user-visible content. Reply only to the question that follows the bracket.',
      '- **Never relabel one tool\'s output as a different metric.** disbursement is NOT collection; npa_activation is NOT npa_closure.',
      '- **Never fabricate totals.** When summarising tool rows, your reply must either sum ALL returned rows, or explicitly say "showing first N of M" with the actual sum. Never invent a total that doesn\'t match the rows shown.',
      '',
      '## Structured answer format (REQUIRED for any list / ranking / snapshot)',
      '- ANY answer that yields a list, ranking, leaderboard, or per-entity snapshot (e.g. "branches below 90% plan attainment", "top 5 FOs", "branches that missed FTOD") MUST be presented as a Markdown pipe table with a header row + a separator row (`|---|---|`). Plain bullet lists of names are forbidden when a metric exists.',
      '- A list answer is NEVER allowed to be a single column of names. Include AT LEAST the entity column AND its primary metric column. Example: "branches below 90% plan attainment" → columns `Branch | Plan Target | Actual | Attainment %`. "FOs with low collection" → `FO | Branch | Demand | Collection | Coll %`.',
      '- Right-align numeric columns in the Markdown separator row using `| ---: |` (e.g. header `| Branch | Plan Target | Actual | Attainment % |` then separator `| :--- | ---: | ---: | ---: |`). Left-align text columns with `| :--- |`.',
      '- Output ONLY ONE structured representation per answer. Do NOT also paste a plain-text/ASCII "table" or duplicate the same data as a bullet list — pick the Markdown pipe table and stop.',
      '- After the table, add a short `**Calculation**` section with 1–3 bullets that show the formula AND the raw inputs. Example: "Plan attainment = Actual ÷ Plan × 100" then "Afzalpur: ₹4.2L ÷ ₹5.0L × 100 = 84%". The reader must be able to reproduce the metric from the bullets.',
      '',
      '## How to answer (text channel — render Markdown, not voice)',
      '- Use rich Markdown: bold key numbers, use bullet lists, headings (##/###) for sections, tables for comparisons.',
      '- Lead with the headline number, then 2-4 supporting bullets or a small table.',
      '- Indian number format: lakh comma pattern. Prefer "₹X.XX Cr" or "₹X.XX L". In Kannada replies, keep number units in English words, for example "12.34 crore" and "5.6 lakh".',
      '- Label count metrics as counts/accounts/cases.',
      '- For comparisons, output a Markdown table (with right-aligned numeric columns and a Calculation section per the structured answer format above).',
      '- For ambiguous lookups, list options and ask which one.',
      '- Never invent numbers. If JSON doesn\'t have it, call a tool.',
      '- **Money formatting — DO NOT do the division yourself.** disbursement_query returns each row with `amount_str` (already formatted, e.g. "₹178.96 Cr") and the wrapper has `totals.amount_str`. Copy `amount_str` VERBATIM into your reply. Do NOT divide `amount` by 10000000 yourself — past replies dropped a zero (₹178.96 Cr → ₹17.90 Cr) and gave wrong totals. If `amount_str` is present, use it exactly as-is. The total row in your table MUST equal `totals.amount_str`, not a re-summed value.',
      '- Never say "snapshot", "JSON", "provided data".',
      '',
      '## Internal context block (DO NOT mention this in your reply)',
      ctxJson,
    ].join('\n');
  } catch (e) {
    console.error('AI stream scope fetch error:', e.message);
  }

  const incomingSystem = (messages.find(m => m.role === 'system') || {}).content || '';
  const mergedSystem = (incomingSystem || '') + (scopedSystemText || '');
  const nonSystem = messages.filter(m => m.role !== 'system');
  const mergedMessages = mergedSystem
    ? [{ role: 'system', content: mergedSystem }, ...nonSystem]
    : nonSystem;

  // Honor explicit provider ('openai'|'mistral'|'gemini'|'deepseek') from request body —
  // when present, use ONLY that provider (no fallback). Missing key → 503-
  // style SSE error so the UI can disable the card.
  const explicitProvider = String((req.body && req.body.provider) || '').toLowerCase().trim();
  const validProvs = ['mistral', 'gemini', 'openai', 'deepseek'];

  // Pick the primary (for the 'open' event the UI displays before tokens).
  let primaryProv;
  if (validProvs.includes(explicitProvider)) {
    primaryProv = explicitProvider;
  } else {
    primaryProv = MISTRAL_KEY ? 'mistral' : (GEMINI_KEY ? 'gemini' : (OPENAI_KEY ? 'openai' : 'deepseek'));
  }
  const primaryModel = primaryProv === 'mistral' ? MISTRAL_MODEL
    : primaryProv === 'openai' ? OPENAI_LLM
    : primaryProv === 'deepseek' ? DEEPSEEK_MODEL
    : 'gemini';

  // If explicit and that key is missing, fail fast with 503-style SSE.
  if (explicitProvider) {
    const haveKey = { mistral: !!MISTRAL_KEY, gemini: !!GEMINI_KEY, openai: !!OPENAI_KEY, deepseek: !!DEEPSEEK_KEY };
    if (!haveKey[explicitProvider]) {
      send('open', { provider: explicitProvider, model: primaryModel });
      send('error', { message: explicitProvider + ' not configured on this server.', reason: 'provider_unavailable', provider: explicitProvider });
      return finish();
    }
  }

  send('open', { provider: primaryProv, model: primaryModel });

  const onProgress = (ev) => {
    if (closed) return;
    if (ev.type === 'thinking') send('thinking', { round: ev.round });
    else if (ev.type === 'tools') send('tools', { names: ev.names });
    else if (ev.type === 'tool_result') send('tool_result', {
      name: ev.name,
      args: ev.args,
      queries: ev.queries,
      result: ev.result,
    });
  };

  // Build chain. Explicit provider → single-element chain (no fallback).
  // Auto → Mistral → Gemini → OpenAI → DeepSeek, skipping unconfigured ones.
  // Build provider chain. When the user picks a specific provider, that
  // provider runs first — but we still queue the other configured providers
  // as auto-fallback so a transient upstream 400/429 doesn't surface as
  // "AI is briefly busy". Frontend gets a `fallback` SSE event so the user
  // sees the switch in the thinking card.
  const chain = [];
  function addProvider(name) {
    if (chain.find(c => c.name === name)) return;
    if (name === 'mistral'  && MISTRAL_KEY)  chain.push({ name, run: () => runMistralWithTools(mergedMessages, session, undefined, onProgress) });
    if (name === 'openai'   && OPENAI_KEY)   chain.push({ name, run: () => runOpenAiWithTools(mergedMessages, session, undefined, onProgress) });
    if (name === 'deepseek' && DEEPSEEK_KEY) chain.push({ name, run: () => runDeepSeekWithTools(mergedMessages, session, undefined, onProgress, deepseekModelOverride) });
    // Gemini intentionally excluded from automatic fallback — quota / leaked-key issues in prod.
  }
  if (explicitProvider && ['mistral','openai','deepseek'].includes(explicitProvider)) {
    addProvider(explicitProvider);
    addProvider('mistral');
    addProvider('openai');
    addProvider('deepseek');
  } else if (explicitProvider === 'gemini') {
    chain.push({ name: 'gemini',  run: () => callGeminiAi(mergedMessages) });
    addProvider('mistral');
    addProvider('openai');
    addProvider('deepseek');
  } else {
    addProvider('mistral');
    addProvider('openai');
    addProvider('deepseek');
  }

  let result = { ok: false, error: 'no_providers' };
  let usedProvider = null;
  for (let i = 0; i < chain.length; i++) {
    if (closed) return;
    const step = chain[i];
    if (i > 0) send('fallback', { from: chain[i - 1].name, to: step.name });
    result = await step.run();
    if (result.ok && result.text) { usedProvider = step.name; break; }
    console.error('Stream: ' + step.name + ' failed:', result.error);
  }

  if (closed) return;

  if (!result.ok || !result.text) {
    const reason = (result && result.error) ? String(result.error) : 'all_providers_failed';
    send('error', { message: 'AI is briefly busy. Please retry in a moment.', reason });
    return finish();
  }

  const fullText = normalizeKannadaNumbersForAi(result.text);
  // Stream in small chunks. Word-grained for natural typing feel.
  const tokens = fullText.match(/\S+\s*|\s+/g) || [fullText];
  for (const tok of tokens) {
    if (closed) return;
    send('delta', { text: tok });
    // Tiny pacing — total ~30 tokens/sec feel without blocking too long.
    await new Promise(r => setTimeout(r, 12));
  }
  const usedModel = usedProvider === 'mistral' ? MISTRAL_MODEL
    : usedProvider === 'openai' ? OPENAI_LLM
    : usedProvider === 'deepseek' ? DEEPSEEK_MODEL
    : (result.model || 'gemini');
  send('done', {
    provider: usedProvider,
    model: usedModel,
    fallback: usedProvider !== primaryProv,
  });
  finish();
});


// ========== OTP Gate (sensitive roles: CEO / RM / DM) ==========
// Demo at /test.employee.html — sends 4-digit OTP via Vasudev SMS.
// Mobile numbers are user-entered (not looked up from DB).
const otpStore = new Map(); // key: `${mobile}|${role}` -> { code, expiresAt, used }
const SENSITIVE_ROLES = ['CEO', 'RM', 'DM'];
const OTP_TTL_MS = 5 * 60 * 1000;
const OTP_MAX_PER_HOUR = 5;
const OTP_BYPASS_MOBILES = new Set(['6361206965']);

function normalizeMobile(raw) {
  const digits = String(raw || '').replace(/\D/g, '');
  if (digits.length === 10) return digits;
  if (digits.length === 12 && digits.startsWith('91')) return digits.slice(2);
  if (digits.length === 11 && digits.startsWith('0')) return digits.slice(1);
  return null;
}

function gen4DigitOtp() {
  return crypto.randomInt(0, 10000).toString().padStart(4, '0');
}

function sendVasudevSms(mobileTen, otp) {
  const senderId = process.env.SMS_SENDER_ID || 'NVCHTN';
  const payload = JSON.stringify({
    Account: {
      User: process.env.SMS_USER,
      Password: process.env.SMS_PASSWORD,
      SenderId: senderId,
      Channel: '2',
      DCS: '8'
    },
    Messages: [{
      Number: '91' + mobileTen,
      Text: 'Your OTP for Navachetana Microfin Service Private Limited is ' + otp + '. Please DO NOT share this OTP with anyone to keep your data safe.'
    }]
  });
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'portal.vasudevsms.in',
      path: '/api/mt/SendSMS',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
        'APIKey': process.env.SMS_API_KEY
      }
    }, (resp) => {
      let body = '';
      resp.on('data', c => body += c);
      resp.on('end', () => {
        if (resp.statusCode >= 200 && resp.statusCode < 300) resolve(body);
        else reject(new Error('SMS gateway HTTP ' + resp.statusCode + ': ' + body.slice(0, 200)));
      });
    });
    req.on('error', reject);
    req.setTimeout(8000, () => { req.destroy(new Error('SMS gateway timeout')); });
    req.write(payload);
    req.end();
  });
}

async function logOtpAudit(mobile, role, outcome, req) {
  try {
    await pool.query(
      "INSERT INTO otp_audit (mobile, role, outcome, ip, user_agent) VALUES ($1, $2, $3, $4, $5)",
      [mobile, role, outcome, String(req.ip || '').slice(0, 45), String(req.get('user-agent') || '').slice(0, 500)]
    );
  } catch (e) {
    console.error('otp_audit insert failed:', e.message);
  }
}

const otpLimiter = rateLimit({ windowMs: 60 * 1000, max: 20, message: { error: 'Too many OTP requests. Please slow down.' } });

app.post('/api/otp/send', otpLimiter, async (req, res) => {
  try {
    const { mobile, role } = req.body || {};
    if (!SENSITIVE_ROLES.includes(role)) {
      return res.status(400).json({ error: 'OTP only required for CEO, RM, or DM.' });
    }
    const mob = normalizeMobile(mobile);
    if (!mob) {
      return res.status(400).json({ error: 'Enter a valid 10-digit Indian mobile number.' });
    }
    // Whitelisted bypass — skip SMS, return success flag for client to skip code step.
    if (OTP_BYPASS_MOBILES.has(mob)) {
      await logOtpAudit(mob, role, 'bypass', req);
      return res.json({ ok: true, bypass: true });
    }
    const rate = await pool.query(
      "SELECT COUNT(*)::int AS n FROM otp_audit WHERE mobile = $1 AND outcome = 'sent' AND created_at > NOW() - INTERVAL '1 hour'",
      [mob]
    );
    if (rate.rows[0].n >= OTP_MAX_PER_HOUR) {
      return res.status(429).json({ error: 'OTP limit reached (5 per hour). Try again later.' });
    }
    if (!process.env.SMS_API_KEY || !process.env.SMS_USER || !process.env.SMS_PASSWORD) {
      return res.status(503).json({ error: 'SMS service not configured.' });
    }
    const code = gen4DigitOtp();
    otpStore.set(mob + '|' + role, { code, expiresAt: Date.now() + OTP_TTL_MS, used: false });
    try {
      await sendVasudevSms(mob, code);
      await logOtpAudit(mob, role, 'sent', req);
      return res.json({ ok: true });
    } catch (smsErr) {
      console.error('Vasudev SMS send failed:', smsErr.message);
      await logOtpAudit(mob, role, 'send_failed', req);
      return res.status(502).json({ error: 'Could not send SMS. Please try again.' });
    }
  } catch (err) {
    console.error('otp send error:', err);
    res.status(500).json({ error: err.message || 'Server error' });
  }
});

app.post('/api/otp/verify', otpLimiter, async (req, res) => {
  try {
    const { mobile, role, code } = req.body || {};
    if (!SENSITIVE_ROLES.includes(role)) {
      return res.status(400).json({ error: 'Invalid role.' });
    }
    const mob = normalizeMobile(mobile);
    if (!mob || !code) {
      return res.status(400).json({ error: 'Missing fields.' });
    }
    if (OTP_BYPASS_MOBILES.has(mob)) {
      await logOtpAudit(mob, role, 'bypass_verified', req);
      return res.json({ ok: true });
    }
    const key = mob + '|' + role;
    const entry = otpStore.get(key);
    const valid = entry && !entry.used && entry.expiresAt > Date.now() && String(entry.code) === String(code).trim();
    if (!valid) {
      await logOtpAudit(mob, role, 'verify_failed', req);
      return res.status(401).json({ error: 'Invalid or expired OTP.' });
    }
    entry.used = true;
    await logOtpAudit(mob, role, 'verified', req);
    return res.json({ ok: true });
  } catch (err) {
    console.error('otp verify error:', err);
    res.status(500).json({ error: err.message || 'Server error' });
  }
});

app.get('/api/otp/audit', async (req, res) => {
  try {
    const r = await pool.query(
      'SELECT id, mobile, role, outcome, ip, created_at FROM otp_audit ORDER BY created_at DESC LIMIT 200'
    );
    res.json(r.rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ========== LIVE LOCATION TRACKING ==========
// Auto-create employee_locations table + indexes on init.
(async function initEmployeeLocations() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS employee_locations (
      id BIGSERIAL PRIMARY KEY,
      emp_id VARCHAR(20) NOT NULL,
      lat DOUBLE PRECISION NOT NULL,
      lng DOUBLE PRECISION NOT NULL,
      accuracy DOUBLE PRECISION,
      battery_pct INT,
      recorded_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`);
    await pool.query("CREATE INDEX IF NOT EXISTS idx_emp_loc_emp_id ON employee_locations(emp_id)");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_emp_loc_recorded_at ON employee_locations(recorded_at DESC)");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_emp_loc_emp_recorded ON employee_locations(emp_id, recorded_at DESC)");
    console.log("employee_locations table ready");
  } catch (err) {
    console.log("employee_locations init skipped:", err.message);
  }
})();

// POST /api/location/ping — body {mobile, lat, lng, accuracy, battery_pct}
app.post('/api/location/ping', async (req, res) => {
  try {
    const { mobile, lat, lng, accuracy, battery_pct } = req.body || {};
    if (mobile === undefined || mobile === null || lat === undefined || lat === null || lng === undefined || lng === null) {
      return res.status(400).json({ error: 'Missing required fields: mobile, lat, lng' });
    }
    const mob = normalizeMobile(mobile);
    if (!mob) {
      return res.status(400).json({ error: 'Invalid mobile number' });
    }
    const latNum = Number(lat);
    const lngNum = Number(lng);
    if (!Number.isFinite(latNum) || !Number.isFinite(lngNum)) {
      return res.status(400).json({ error: 'Invalid lat/lng' });
    }
    const accNum = (accuracy === undefined || accuracy === null || accuracy === '') ? null : Number(accuracy);
    const battNum = (battery_pct === undefined || battery_pct === null || battery_pct === '') ? null : parseInt(battery_pct, 10);
    const empRes = await pool.query(
      "SELECT emp_id FROM employee_master WHERE REGEXP_REPLACE(COALESCE(mobile,''), '\\D', '', 'g') LIKE $1 LIMIT 1",
      ['%' + mob]
    );
    if (empRes.rows.length === 0) {
      return res.status(404).json({ error: 'Mobile not found in employee_master' });
    }
    const empId = empRes.rows[0].emp_id;
    await pool.query(
      `INSERT INTO employee_locations (emp_id, lat, lng, accuracy, battery_pct, recorded_at)
       VALUES ($1, $2, $3, $4, $5, NOW())`,
      [empId, latNum, lngNum, accNum, (Number.isFinite(battNum) ? battNum : null)]
    );
    res.json({ ok: true, emp_id: empId });
  } catch (err) {
    console.error('location ping error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/location/live?role=&location= — scoped employees + latest position
app.get('/api/location/live', async (req, res) => {
  try {
    const role = String(req.query.role || '').toUpperCase().trim();
    const location = (req.query.location || '').toString().trim();
    if (!role) {
      return res.status(400).json({ error: 'role is required' });
    }
    const where = ["em.status = 'Working'"];
    const params = [];
    let idx = 1;
    if (role === 'CEO') {
      // no scope filter
    } else if (role === 'RM' || role === 'SM') {
      if (!location) return res.status(400).json({ error: 'location required for RM/SM' });
      where.push("TRIM(em.region_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(location);
    } else if (role === 'DM' || role === 'DVM') {
      if (!location) return res.status(400).json({ error: 'location required for DM/DvM' });
      where.push("TRIM(em.division_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(location);
    } else if (role === 'AM') {
      if (!location) return res.status(400).json({ error: 'location required for AM' });
      where.push("TRIM(em.area_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(location);
    } else if (role === 'BM' || role === 'FO') {
      if (!location) return res.status(400).json({ error: 'location required for BM/FO' });
      where.push("TRIM(em.branch_name) ILIKE TRIM($" + (idx++) + ")");
      params.push(location);
    } else {
      return res.status(400).json({ error: 'Unsupported role: ' + role });
    }
    const sql = `
      SELECT em.emp_id, em.full_name, em.mobile, em.role,
             em.branch_name, em.area_name, em.division_name, em.region_name,
             loc.lat, loc.lng, loc.accuracy, loc.battery_pct, loc.recorded_at,
             CASE
               WHEN loc.recorded_at IS NULL THEN 'inactive'
               WHEN loc.recorded_at >= NOW() - INTERVAL '5 minutes' THEN 'live'
               WHEN loc.recorded_at >= NOW() - INTERVAL '7 days' THEN 'last_known'
               ELSE 'inactive'
             END AS status
      FROM employee_master em
      LEFT JOIN LATERAL (
        SELECT lat, lng, accuracy, battery_pct, recorded_at
        FROM employee_locations
        WHERE emp_id = em.emp_id
        ORDER BY recorded_at DESC
        LIMIT 1
      ) loc ON TRUE
      WHERE ${where.join(' AND ')}
      ORDER BY em.full_name`;
    const result = await pool.query(sql, params);
    // Map full_name -> name and rename branch_name/area_name/division_name/region_name
    // for client convenience. Wrap in {employees, count} for forward-compat.
    const employees = result.rows.map(r => ({
      emp_id: r.emp_id,
      name: r.full_name,
      mobile: r.mobile,
      role: r.role,
      branch: r.branch_name,
      area: r.area_name,
      division: r.division_name,
      region: r.region_name,
      lat: r.lat,
      lng: r.lng,
      accuracy: r.accuracy,
      battery_pct: r.battery_pct,
      recorded_at: r.recorded_at,
      status: r.status,
    }));
    res.json({ employees, count: employees.length });
  } catch (err) {
    console.error('location live error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Hourly purge job: delete location pings older than 7 days.
setInterval(async () => {
  try {
    const r = await pool.query("DELETE FROM employee_locations WHERE recorded_at < NOW() - INTERVAL '7 days'");
    if (r.rowCount && r.rowCount > 0) {
      console.log('employee_locations purge: removed ' + r.rowCount + ' rows');
    }
  } catch (err) {
    console.error('employee_locations purge error:', err.message);
  }
}, 60 * 60 * 1000);

// ====================================================================
//  CHAT — tables, endpoints, sockets, 7-day purge
// ====================================================================

// Auto-create chat tables + indexes on init.
(async function initChatTables() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS chat_groups (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      kind TEXT NOT NULL CHECK (kind IN ('auto','custom')),
      scope_type TEXT,
      scope_value TEXT,
      created_by_emp_id TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS chat_group_members (
      group_id TEXT REFERENCES chat_groups(id) ON DELETE CASCADE,
      emp_id TEXT NOT NULL,
      joined_at TIMESTAMPTZ DEFAULT NOW(),
      PRIMARY KEY (group_id, emp_id)
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS chat_messages (
      id BIGSERIAL PRIMARY KEY,
      sender_emp_id TEXT NOT NULL,
      thread_key TEXT NOT NULL,
      body TEXT,
      sent_at TIMESTAMPTZ DEFAULT NOW(),
      read_by_json JSONB DEFAULT '{}'::jsonb
    )`);
    // Idempotency token from the client — dedupe retries of the same logical send.
    await pool.query("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS client_msg_id TEXT");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_chat_msg_thread_sent ON chat_messages(thread_key, sent_at DESC, id DESC)");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_chat_msg_sender ON chat_messages(sender_emp_id)");
    // Unique idempotency per sender — only enforced when client_msg_id is supplied.
    await pool.query("CREATE UNIQUE INDEX IF NOT EXISTS uq_chat_msg_client ON chat_messages(sender_emp_id, client_msg_id) WHERE client_msg_id IS NOT NULL");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_chat_grp_member_emp ON chat_group_members(emp_id)");
    console.log('chat_* tables ready');
  } catch (err) {
    console.log('chat init skipped:', err.message);
  }
})();

// ====================================================================
//  FEEDBACK — employees submit feature requests / complaints; admin reads
// ====================================================================

(async function initFeedbackTable() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS employee_feedback (
      id BIGSERIAL PRIMARY KEY,
      emp_id TEXT,
      emp_name TEXT,
      role TEXT,
      branch TEXT,
      message TEXT NOT NULL,
      status TEXT DEFAULT 'open',
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query("CREATE INDEX IF NOT EXISTS idx_feedback_created ON employee_feedback(created_at DESC)");
    await pool.query("CREATE INDEX IF NOT EXISTS idx_feedback_emp ON employee_feedback(emp_id)");
    console.log('employee_feedback table ready');
  } catch (err) {
    console.log('feedback init skipped:', err.message);
  }
})();

app.post("/api/feedback", async (req, res) => {
  try {
    const body = req.body || {};
    const message = String(body.message || '').trim();
    if (!message) return res.status(400).json({ error: 'message required' });
    if (message.length > 4000) return res.status(400).json({ error: 'message too long (max 4000 chars)' });
    const empId = body.emp_id ? String(body.emp_id).slice(0, 32) : null;
    const empName = body.emp_name ? String(body.emp_name).slice(0, 200) : null;
    const role = body.role ? String(body.role).slice(0, 32) : null;
    const branch = body.branch ? String(body.branch).slice(0, 200) : null;
    const r = await pool.query(
      `INSERT INTO employee_feedback (emp_id, emp_name, role, branch, message)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id, created_at`,
      [empId, empName, role, branch, message]
    );
    res.status(201).json({ id: r.rows[0].id, created_at: r.rows[0].created_at });
  } catch (e) {
    console.error('feedback POST error:', e.message);
    res.status(500).json({ error: 'failed' });
  }
});

app.get("/api/feedback", async (req, res) => {
  try {
    const lim = Math.min(Math.max(parseInt(req.query.limit, 10) || 200, 1), 500);
    const role = String(req.query.role || '').toUpperCase();
    const empId = req.query.emp_id ? String(req.query.emp_id) : null;
    let sql, params;
    if (role === 'CEO' || role === '') {
      sql = `SELECT id, emp_id, emp_name, role, branch, message, status, created_at
               FROM employee_feedback
              ORDER BY created_at DESC
              LIMIT ${lim}`;
      params = [];
    } else if (empId) {
      sql = `SELECT id, emp_id, emp_name, role, branch, message, status, created_at
               FROM employee_feedback
              WHERE emp_id = $1
              ORDER BY created_at DESC
              LIMIT ${lim}`;
      params = [empId];
    } else {
      sql = `SELECT id, emp_id, emp_name, role, branch, message, status, created_at
               FROM employee_feedback
              ORDER BY created_at DESC
              LIMIT ${lim}`;
      params = [];
    }
    const r = await pool.query(sql, params);
    res.json({ count: r.rows.length, items: r.rows });
  } catch (e) {
    console.error('feedback GET error:', e.message);
    res.status(500).json({ error: 'failed' });
  }
});

// ----- Helpers: hierarchy + thread membership ------------------------

async function chatGetEmployee(empId) {
  if (!empId) return null;
  const r = await pool.query(
    "SELECT emp_id, full_name, role, branch_name, area_name, division_name, region_name, status FROM employee_master WHERE emp_id = $1 LIMIT 1",
    [String(empId)]
  );
  return r.rows[0] || null;
}

// Look up the canonical chain for a given level/value.
// Used to compute upstream membership (e.g., "what region does branch X live in?").
async function chatCanonicalChain(level, value) {
  if (!value) return null;
  let col = null;
  if (level === 'branch') col = 'branch_name';
  else if (level === 'area') col = 'area_name';
  else if (level === 'division') col = 'division_name';
  else if (level === 'region') col = 'region_name';
  else return null;
  const r = await pool.query(
    `SELECT region_name, division_name, area_name, branch_name FROM employee_master
       WHERE TRIM(${col}) ILIKE TRIM($1) AND status = 'Working' LIMIT 1`,
    [String(value)]
  );
  return r.rows[0] || null;
}

function chatParseThread(threadKey) {
  // Returns {kind, ...details} or null.
  if (typeof threadKey !== 'string') return null;
  if (threadKey.startsWith('dm:')) {
    const parts = threadKey.split(':');
    if (parts.length !== 3) return null;
    return { kind: 'dm', a: parts[1], b: parts[2] };
  }
  if (threadKey.startsWith('auto:')) {
    const parts = threadKey.split(':');
    if (parts.length < 3) return null;
    return { kind: 'auto', level: parts[1], value: parts.slice(2).join(':') };
  }
  if (threadKey.startsWith('custom:')) {
    return { kind: 'custom', groupId: threadKey };
  }
  return null;
}

// Returns true if emp_id may read/post into this thread.
async function chatIsMember(empId, threadKey) {
  const parsed = chatParseThread(threadKey);
  if (!parsed) return false;
  const emp = await chatGetEmployee(empId);
  if (!emp) return false;
  if (parsed.kind === 'dm') {
    return String(empId) === parsed.a || String(empId) === parsed.b;
  }
  if (parsed.kind === 'auto') {
    if (String(emp.role || '').toUpperCase() === 'CEO') return true;
    const v = parsed.value;
    const norm = (x) => (x == null ? '' : String(x).trim().toUpperCase());
    if (parsed.level === 'region') {
      return norm(emp.region_name) === norm(v);
    }
    if (parsed.level === 'division') {
      if (norm(emp.division_name) === norm(v)) return true;
      const chain = await chatCanonicalChain('division', v);
      if (!chain) return false;
      const role = norm(emp.role);
      return role === 'RM' && norm(emp.region_name) === norm(chain.region_name);
    }
    if (parsed.level === 'area') {
      if (norm(emp.area_name) === norm(v)) return true;
      const chain = await chatCanonicalChain('area', v);
      if (!chain) return false;
      const role = norm(emp.role);
      if (role === 'DM' || role === 'DVM') {
        return norm(emp.division_name) === norm(chain.division_name);
      }
      if (role === 'RM') {
        return norm(emp.region_name) === norm(chain.region_name);
      }
      return false;
    }
    if (parsed.level === 'branch') {
      if (norm(emp.branch_name) === norm(v)) return true;
      const chain = await chatCanonicalChain('branch', v);
      if (!chain) return false;
      const role = norm(emp.role);
      if (role === 'AM') return norm(emp.area_name) === norm(chain.area_name);
      if (role === 'DM' || role === 'DVM') {
        return norm(emp.division_name) === norm(chain.division_name);
      }
      if (role === 'RM') {
        return norm(emp.region_name) === norm(chain.region_name);
      }
      return false;
    }
    return false;
  }
  if (parsed.kind === 'custom') {
    const r = await pool.query(
      'SELECT 1 FROM chat_group_members WHERE group_id = $1 AND emp_id = $2 LIMIT 1',
      [parsed.groupId, String(empId)]
    );
    return r.rows.length > 0;
  }
  return false;
}

// Recipients of a post into thread (including the sender, the caller filters).
async function chatThreadRecipients(threadKey) {
  const parsed = chatParseThread(threadKey);
  if (!parsed) return [];
  if (parsed.kind === 'dm') {
    return [parsed.a, parsed.b];
  }
  if (parsed.kind === 'auto') {
    const v = parsed.value;
    if (parsed.level === 'region') {
      const r = await pool.query(
        "SELECT emp_id FROM employee_master WHERE (TRIM(region_name) ILIKE TRIM($1) AND status='Working') OR UPPER(role) = 'CEO'",
        [v]
      );
      return r.rows.map((x) => x.emp_id);
    }
    if (parsed.level === 'division') {
      const chain = await chatCanonicalChain('division', v);
      const params = [v];
      let extra = '';
      if (chain) {
        params.push(chain.region_name || '');
        extra = " OR (UPPER(role)='RM' AND TRIM(region_name) ILIKE TRIM($2))";
      }
      const r = await pool.query(
        `SELECT emp_id FROM employee_master WHERE
           (TRIM(division_name) ILIKE TRIM($1) AND status='Working')${extra}
           OR UPPER(role)='CEO'`,
        params
      );
      return r.rows.map((x) => x.emp_id);
    }
    if (parsed.level === 'area') {
      const chain = await chatCanonicalChain('area', v);
      const params = [v];
      let extra = '';
      if (chain) {
        params.push(chain.division_name || '', chain.region_name || '');
        extra =
          " OR (UPPER(role) IN ('DM','DVM') AND TRIM(division_name) ILIKE TRIM($2))" +
          " OR (UPPER(role)='RM' AND TRIM(region_name) ILIKE TRIM($3))";
      }
      const r = await pool.query(
        `SELECT emp_id FROM employee_master WHERE
           (TRIM(area_name) ILIKE TRIM($1) AND status='Working')${extra}
           OR UPPER(role)='CEO'`,
        params
      );
      return r.rows.map((x) => x.emp_id);
    }
    if (parsed.level === 'branch') {
      const chain = await chatCanonicalChain('branch', v);
      const params = [v];
      let extra = '';
      if (chain) {
        params.push(
          chain.area_name || '',
          chain.division_name || '',
          chain.region_name || ''
        );
        extra =
          " OR (UPPER(role)='AM' AND TRIM(area_name) ILIKE TRIM($2))" +
          " OR (UPPER(role) IN ('DM','DVM') AND TRIM(division_name) ILIKE TRIM($3))" +
          " OR (UPPER(role)='RM' AND TRIM(region_name) ILIKE TRIM($4))";
      }
      const r = await pool.query(
        `SELECT emp_id FROM employee_master WHERE
           (TRIM(branch_name) ILIKE TRIM($1) AND status='Working')${extra}
           OR UPPER(role)='CEO'`,
        params
      );
      return r.rows.map((x) => x.emp_id);
    }
    return [];
  }
  if (parsed.kind === 'custom') {
    const r = await pool.query(
      'SELECT emp_id FROM chat_group_members WHERE group_id = $1',
      [parsed.groupId]
    );
    return r.rows.map((x) => x.emp_id);
  }
  return [];
}

function chatNanoid(n) {
  return crypto.randomBytes(Math.ceil(n / 2)).toString('hex').slice(0, n);
}

// ----- Endpoints -----------------------------------------------------

// GET /api/chat/scopes?emp_id=ID
app.get('/api/chat/scopes', async (req, res) => {
  try {
    const empId = String(req.query.emp_id || '').trim();
    if (!empId) return res.status(400).json({ error: 'emp_id required' });
    const emp = await chatGetEmployee(empId);
    if (!emp) return res.status(404).json({ error: 'Employee not found' });
    const role = String(emp.role || '').toUpperCase();
    const out = { regions: [], divisions: [], areas: [], branches: [] };
    // Own chain.
    if (emp.region_name) out.regions.push(emp.region_name);
    if (emp.division_name) out.divisions.push(emp.division_name);
    if (emp.area_name) out.areas.push(emp.area_name);
    if (emp.branch_name) out.branches.push(emp.branch_name);
    // Downstream scope.
    let extra = null;
    if (role === 'CEO') {
      extra = await pool.query(
        "SELECT DISTINCT region_name, division_name, area_name, branch_name FROM employee_master WHERE status='Working'"
      );
    } else if (role === 'RM') {
      extra = await pool.query(
        "SELECT DISTINCT region_name, division_name, area_name, branch_name FROM employee_master WHERE status='Working' AND TRIM(region_name) ILIKE TRIM($1)",
        [emp.region_name || '']
      );
    } else if (role === 'DM' || role === 'DVM') {
      extra = await pool.query(
        "SELECT DISTINCT region_name, division_name, area_name, branch_name FROM employee_master WHERE status='Working' AND TRIM(division_name) ILIKE TRIM($1)",
        [emp.division_name || '']
      );
    } else if (role === 'AM') {
      extra = await pool.query(
        "SELECT DISTINCT region_name, division_name, area_name, branch_name FROM employee_master WHERE status='Working' AND TRIM(area_name) ILIKE TRIM($1)",
        [emp.area_name || '']
      );
    }
    if (extra) {
      for (const r of extra.rows) {
        if (r.region_name && !out.regions.includes(r.region_name)) out.regions.push(r.region_name);
        if (r.division_name && !out.divisions.includes(r.division_name)) out.divisions.push(r.division_name);
        if (r.area_name && !out.areas.includes(r.area_name)) out.areas.push(r.area_name);
        if (r.branch_name && !out.branches.includes(r.branch_name)) out.branches.push(r.branch_name);
      }
    }
    out.regions.sort();
    out.divisions.sort();
    out.areas.sort();
    out.branches.sort();
    res.json(out);
  } catch (err) {
    console.error('chat scopes error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/chat/threads?emp_id=ID
// Returns recent threads with last message + unread count + member count.
// Optimized: one window-function query for last message + unread count, batched
// group lookups, member count via chatThreadRecipients only for `auto` kind.
app.get('/api/chat/threads', async (req, res) => {
  try {
    const empId = String(req.query.emp_id || '').trim();
    if (!empId) return res.status(400).json({ error: 'emp_id required' });
    // Single CTE: candidate threads + their last message + unread count.
    const rows = await pool.query(
      `WITH candidates AS (
         SELECT thread_key, MAX(sent_at) AS last_sent_at
           FROM chat_messages
          GROUP BY thread_key
          ORDER BY MAX(sent_at) DESC
          LIMIT 500
       ),
       last_msg AS (
         SELECT DISTINCT ON (m.thread_key)
                m.thread_key, m.id, m.body, m.sent_at, m.sender_emp_id, m.read_by_json
           FROM chat_messages m
           JOIN candidates c ON c.thread_key = m.thread_key
          ORDER BY m.thread_key, m.sent_at DESC, m.id DESC
       ),
       unread AS (
         SELECT m.thread_key, COUNT(*)::int AS n
           FROM chat_messages m
           JOIN candidates c ON c.thread_key = m.thread_key
          WHERE m.sender_emp_id <> $1
            AND m.id > COALESCE((m.read_by_json->>$1)::bigint, 0)
          GROUP BY m.thread_key
       )
       SELECT c.thread_key, c.last_sent_at,
              l.id AS last_id, l.body AS last_body, l.sent_at AS last_at,
              l.sender_emp_id AS last_sender, l.read_by_json AS last_read_by,
              em.full_name AS last_sender_name,
              COALESCE(u.n, 0) AS unread_count
         FROM candidates c
         LEFT JOIN last_msg l ON l.thread_key = c.thread_key
         LEFT JOIN unread u ON u.thread_key = c.thread_key
         LEFT JOIN employee_master em ON em.emp_id = l.sender_emp_id
        ORDER BY c.last_sent_at DESC`,
      [String(empId)]
    );
    // Filter by membership and enrich. Use caches to avoid duplicate lookups
    // across the 100-thread page.
    const dmEmpIds = new Set();
    const customGroupIds = new Set();
    const memberOk = [];
    for (const r of rows.rows) {
      const ok = await chatIsMember(empId, r.thread_key);
      if (!ok) continue;
      memberOk.push(r);
      const parsed = chatParseThread(r.thread_key);
      if (parsed && parsed.kind === 'dm') {
        const otherId = parsed.a === String(empId) ? parsed.b : parsed.a;
        if (otherId) dmEmpIds.add(otherId);
      } else if (parsed && parsed.kind === 'custom') {
        customGroupIds.add(parsed.groupId);
      }
      if (memberOk.length >= 100) break;
    }
    // Batch DM peer name lookup.
    const dmNames = new Map();
    if (dmEmpIds.size > 0) {
      const r = await pool.query(
        'SELECT emp_id, full_name FROM employee_master WHERE emp_id = ANY($1::text[])',
        [Array.from(dmEmpIds)]
      );
      for (const row of r.rows) dmNames.set(row.emp_id, row.full_name);
    }
    // Batch custom group name + member count lookups.
    const groupInfo = new Map();
    if (customGroupIds.size > 0) {
      const ids = Array.from(customGroupIds);
      const gn = await pool.query(
        'SELECT id, name FROM chat_groups WHERE id = ANY($1::text[])',
        [ids]
      );
      const gc = await pool.query(
        `SELECT group_id, COUNT(*)::int AS n
           FROM chat_group_members
          WHERE group_id = ANY($1::text[])
          GROUP BY group_id`,
        [ids]
      );
      for (const row of gn.rows) groupInfo.set(row.id, { name: row.name, count: 0 });
      for (const row of gc.rows) {
        const e = groupInfo.get(row.group_id) || { name: null, count: 0 };
        e.count = Number(row.n) || 0;
        groupInfo.set(row.group_id, e);
      }
    }
    const out = [];
    for (const r of memberOk) {
      const tk = r.thread_key;
      const parsed = chatParseThread(tk);
      let title = tk;
      let kind = 'auto';
      let memberCount = 0;
      if (parsed) {
        kind = parsed.kind;
        if (parsed.kind === 'dm') {
          const otherId = parsed.a === String(empId) ? parsed.b : parsed.a;
          title = dmNames.get(otherId) || otherId;
          memberCount = 2;
        } else if (parsed.kind === 'auto') {
          title = parsed.value;
          const recip = await chatThreadRecipients(tk);
          memberCount = recip.length;
        } else if (parsed.kind === 'custom') {
          const info = groupInfo.get(parsed.groupId);
          title = (info && info.name) || tk;
          memberCount = info ? info.count : 0;
        }
      }
      out.push({
        thread_key: tk,
        title,
        kind,
        last_message: r.last_id
          ? {
              id: Number(r.last_id),
              body: r.last_body,
              sent_at: r.last_at,
              sender_emp_id: r.last_sender,
              sender_name: r.last_sender_name,
            }
          : null,
        unread_count: Number(r.unread_count) || 0,
        member_count: memberCount,
      });
    }
    res.json({ threads: out, count: out.length });
  } catch (err) {
    console.error('chat threads error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/chat/messages?thread_key=K&emp_id=ID&limit=50&before_id=N&before_sent_at=ISO&after_id=N
// Two cursor modes:
//   - Older messages (default): pass `before_id` (+ optional `before_sent_at`).
//     Returns DESC-ordered page with has_more / next_before_* cursor.
//   - Reconnect catch-up: pass `after_id` to fetch every message strictly newer
//     than that anchor in the thread, ASC-ordered. Useful for socket-reconnect
//     replay of a single thread.
// `before_*` and `after_id` are mutually exclusive; if both are given,
// `after_id` wins.
app.get('/api/chat/messages', async (req, res) => {
  try {
    const threadKey = String(req.query.thread_key || '').trim();
    const empId = String(req.query.emp_id || '').trim();
    const limit = Math.min(parseInt(req.query.limit, 10) || 50, 200);
    const beforeId = req.query.before_id ? parseInt(req.query.before_id, 10) : null;
    const beforeSentAt = req.query.before_sent_at
      ? new Date(String(req.query.before_sent_at))
      : null;
    const afterId = req.query.after_id ? parseInt(req.query.after_id, 10) : null;
    if (!threadKey || !empId) {
      return res.status(400).json({ error: 'thread_key and emp_id required' });
    }
    const ok = await chatIsMember(empId, threadKey);
    if (!ok) return res.status(403).json({ error: 'Not a member of this thread' });
    const params = [threadKey];
    let where = 'm.thread_key = $1';
    let order = 'm.sent_at DESC, m.id DESC';
    let mode = 'before';
    if (afterId !== null && !isNaN(afterId)) {
      params.push(afterId);
      where += ` AND m.id > $${params.length}`;
      order = 'm.sent_at ASC, m.id ASC';
      mode = 'after';
    } else if (beforeSentAt && !isNaN(beforeSentAt.getTime()) && beforeId) {
      // Stable composite cursor.
      params.push(beforeSentAt.toISOString(), beforeId);
      where += ` AND (m.sent_at, m.id) < ($${params.length - 1}::timestamptz, $${params.length}::bigint)`;
    } else if (beforeId) {
      params.push(beforeId);
      where += ` AND m.id < $${params.length}`;
    }
    params.push(limit + 1);
    const limitIdx = params.length;
    const sql = `SELECT m.id, m.sender_emp_id, m.thread_key, m.body, m.sent_at,
                        m.read_by_json, m.client_msg_id,
                        em.full_name AS sender_name
                   FROM chat_messages m
                   LEFT JOIN employee_master em ON em.emp_id = m.sender_emp_id
                  WHERE ${where}
                  ORDER BY ${order}
                  LIMIT $${limitIdx}`;
    const r = await pool.query(sql, params);
    const hasMore = r.rows.length > limit;
    const messages = hasMore ? r.rows.slice(0, limit) : r.rows;
    const last = messages[messages.length - 1] || null;
    res.json({
      messages,
      count: messages.length,
      has_more: hasMore,
      mode,
      // Older-page cursor (only meaningful for `mode === 'before'`).
      next_before_id: mode === 'before' && last ? Number(last.id) : null,
      next_before_sent_at: mode === 'before' && last ? last.sent_at : null,
      // Newer-page cursor (only meaningful for `mode === 'after'`).
      next_after_id: mode === 'after' && last ? Number(last.id) : null,
    });
  } catch (err) {
    console.error('chat messages error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ----- Rate limit (token bucket) for chat:send per emp_id -----------
// 1 token/sec sustained, burst of 5. In-memory; resets on restart.
const _chatSendBuckets = new Map();
function chatSendAcquireToken(empId) {
  if (!empId) return false;
  const now = Date.now();
  const RATE_PER_MS = 1 / 1000; // 1 token/sec
  const BURST = 5;
  let b = _chatSendBuckets.get(empId);
  if (!b) {
    b = { tokens: BURST, ts: now };
    _chatSendBuckets.set(empId, b);
  }
  // Refill.
  const elapsed = now - b.ts;
  if (elapsed > 0) {
    b.tokens = Math.min(BURST, b.tokens + elapsed * RATE_PER_MS);
    b.ts = now;
  }
  if (b.tokens >= 1) {
    b.tokens -= 1;
    return true;
  }
  return false;
}
// Periodic GC so the map doesn't grow unbounded.
setInterval(() => {
  const cutoff = Date.now() - 10 * 60 * 1000;
  for (const [k, v] of _chatSendBuckets) {
    if (v.ts < cutoff) _chatSendBuckets.delete(k);
  }
}, 5 * 60 * 1000).unref?.();

// Internal: insert a message + emit socket events. Returns the inserted row.
// Supports idempotency via clientMsgId — same (sender, clientMsgId) returns the
// previously inserted row instead of creating a duplicate.
async function chatPostInternal({ senderEmpId, threadKey, body, clientMsgId }) {
  if (!senderEmpId || !threadKey) {
    const e = new Error('sender_emp_id and thread_key required');
    e.status = 400;
    throw e;
  }
  if (!body || !String(body).trim()) {
    const e = new Error('body required');
    e.status = 400;
    throw e;
  }
  const ok = await chatIsMember(senderEmpId, threadKey);
  if (!ok) {
    const e = new Error('Not a member of this thread');
    e.status = 403;
    throw e;
  }
  const cmid = clientMsgId ? String(clientMsgId).slice(0, 64) : null;
  // Idempotency check — if the client retried with the same client_msg_id,
  // return the existing row without re-broadcasting.
  if (cmid) {
    const existing = await pool.query(
      `SELECT m.id, m.sender_emp_id, m.thread_key, m.body, m.sent_at, m.read_by_json,
              m.client_msg_id, em.full_name AS sender_name
         FROM chat_messages m
         LEFT JOIN employee_master em ON em.emp_id = m.sender_emp_id
        WHERE m.sender_emp_id = $1 AND m.client_msg_id = $2
        LIMIT 1`,
      [String(senderEmpId), cmid]
    );
    if (existing.rows.length > 0) {
      const row = existing.rows[0];
      row.duplicate = true;
      return row;
    }
  }
  const ins = await pool.query(
    `INSERT INTO chat_messages (sender_emp_id, thread_key, body, client_msg_id)
       VALUES ($1, $2, $3, $4)
       RETURNING id, sender_emp_id, thread_key, body, sent_at, read_by_json, client_msg_id`,
    [String(senderEmpId), threadKey, String(body).trim(), cmid]
  );
  const row = ins.rows[0];
  // Sender display name for the broadcast.
  const senderRes = await pool.query(
    'SELECT full_name FROM employee_master WHERE emp_id = $1 LIMIT 1',
    [String(senderEmpId)]
  );
  row.sender_name = senderRes.rows[0]?.full_name || null;
  // Broadcast to recipients (including sender for confirmation).
  const recipients = await chatThreadRecipients(threadKey);
  const seen = new Set();
  for (const rid of recipients) {
    if (!rid) continue;
    if (seen.has(String(rid))) continue;
    seen.add(String(rid));
    io.to('emp:' + String(rid)).emit('chat:new', row);
  }
  return row;
}

// POST /api/chat/send
// Accepts optional `client_msg_id` for idempotent retries. Returns the row
// (with `duplicate: true` if the same client_msg_id was already accepted).
app.post('/api/chat/send', async (req, res) => {
  try {
    const { sender_emp_id, thread_key, body, client_msg_id } = req.body || {};
    if (sender_emp_id && !chatSendAcquireToken(String(sender_emp_id))) {
      return res.status(429).json({ error: 'rate_limited', retry_after_ms: 1000 });
    }
    const row = await chatPostInternal({
      senderEmpId: sender_emp_id,
      threadKey: thread_key,
      body,
      clientMsgId: client_msg_id,
    });
    res.json(row);
  } catch (err) {
    console.error('chat send error:', err);
    res.status(err.status || 500).json({ error: err.message });
  }
});

// Validate emp_id format before using it as a JSONB key — alnum + dash/underscore
// only, max 32 chars. Prevents malformed IDs from corrupting read_by_json.
function chatValidEmpId(s) {
  return typeof s === 'string' && /^[A-Za-z0-9_-]{1,32}$/.test(s);
}

// POST /api/chat/read
app.post('/api/chat/read', async (req, res) => {
  try {
    const { emp_id, thread_key, up_to_message_id } = req.body || {};
    if (!emp_id || !thread_key || up_to_message_id === undefined) {
      return res.status(400).json({ error: 'emp_id, thread_key, up_to_message_id required' });
    }
    if (!chatValidEmpId(String(emp_id))) {
      return res.status(400).json({ error: 'invalid emp_id format' });
    }
    const ok = await chatIsMember(emp_id, thread_key);
    if (!ok) return res.status(403).json({ error: 'Not a member of this thread' });
    // jsonb_set the emp_id key on the latest message — but we want each message
    // up to up_to_message_id to have read_by_json[emp_id] = id.
    await pool.query(
      `UPDATE chat_messages
          SET read_by_json = jsonb_set(read_by_json, ARRAY[$1::text], to_jsonb($2::bigint), true)
        WHERE thread_key = $3
          AND id <= $2
          AND COALESCE((read_by_json->>$1)::bigint, 0) < $2`,
      [String(emp_id), Number(up_to_message_id), String(thread_key)]
    );
    // Notify thread members so other clients can update tick state.
    const recipients = await chatThreadRecipients(thread_key);
    const seen = new Set();
    for (const rid of recipients) {
      if (!rid) continue;
      if (seen.has(String(rid))) continue;
      seen.add(String(rid));
      io.to('emp:' + String(rid)).emit('chat:read', {
        thread_key,
        emp_id: String(emp_id),
        up_to_message_id: Number(up_to_message_id),
      });
    }
    res.json({ ok: true });
  } catch (err) {
    console.error('chat read error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/chat/unread?emp_id=ID
// Lightweight per-thread unread counts so the chatSidebarBadge can update in
// O(1) without re-running the heavy /threads pipeline.
// Response: { total, threads: { [thread_key]: count } }
app.get('/api/chat/unread', async (req, res) => {
  try {
    const empId = String(req.query.emp_id || '').trim();
    if (!empId) return res.status(400).json({ error: 'emp_id required' });
    if (!chatValidEmpId(empId)) {
      return res.status(400).json({ error: 'invalid emp_id format' });
    }
    // Pull recent threads once, filter by membership, then aggregate.
    const recent = await pool.query(
      `SELECT thread_key, MAX(sent_at) AS last_sent_at
         FROM chat_messages
         GROUP BY thread_key
         ORDER BY MAX(sent_at) DESC
         LIMIT 500`
    );
    const myThreads = [];
    for (const r of recent.rows) {
      const ok = await chatIsMember(empId, r.thread_key);
      if (ok) myThreads.push(r.thread_key);
    }
    if (myThreads.length === 0) return res.json({ total: 0, threads: {} });
    const counts = await pool.query(
      `SELECT thread_key, COUNT(*)::int AS n
         FROM chat_messages
        WHERE thread_key = ANY($1::text[])
          AND sender_emp_id <> $2
          AND id > COALESCE((read_by_json->>$2)::bigint, 0)
        GROUP BY thread_key`,
      [myThreads, empId]
    );
    const threads = {};
    let total = 0;
    for (const row of counts.rows) {
      const n = Number(row.n) || 0;
      if (n > 0) {
        threads[row.thread_key] = n;
        total += n;
      }
    }
    res.json({ total, threads });
  } catch (err) {
    console.error('chat unread error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/chat/groups — create custom group.
// Validates that every supplied emp_id exists in employee_master and reports
// any unknown ones in `skipped` instead of silently inserting them.
app.post('/api/chat/groups', async (req, res) => {
  try {
    const { created_by_emp_id, name, member_emp_ids } = req.body || {};
    if (!created_by_emp_id || !name || !Array.isArray(member_emp_ids)) {
      return res.status(400).json({ error: 'created_by_emp_id, name, member_emp_ids[] required' });
    }
    if (!chatValidEmpId(String(created_by_emp_id))) {
      return res.status(400).json({ error: 'invalid created_by_emp_id format' });
    }
    const trimmedName = String(name).trim();
    if (!trimmedName) return res.status(400).json({ error: 'name required' });
    // Dedupe + normalize requested members. Always include the creator.
    const requested = new Set(
      member_emp_ids.map((x) => String(x || '').trim()).filter(Boolean)
    );
    requested.add(String(created_by_emp_id));
    // Verify each emp_id exists.
    const verifyRes = await pool.query(
      "SELECT emp_id FROM employee_master WHERE emp_id = ANY($1::text[])",
      [Array.from(requested)]
    );
    const valid = new Set(verifyRes.rows.map((r) => r.emp_id));
    const skipped = Array.from(requested).filter((e) => !valid.has(e));
    if (!valid.has(String(created_by_emp_id))) {
      return res.status(400).json({ error: 'created_by_emp_id not found in employee_master' });
    }
    const id = 'custom:' + chatNanoid(10);
    await pool.query(
      `INSERT INTO chat_groups (id, name, kind, created_by_emp_id) VALUES ($1, $2, 'custom', $3)`,
      [id, trimmedName, String(created_by_emp_id)]
    );
    for (const mid of valid) {
      await pool.query(
        'INSERT INTO chat_group_members (group_id, emp_id) VALUES ($1, $2) ON CONFLICT DO NOTHING',
        [id, mid]
      );
    }
    res.json({
      id,
      name: trimmedName,
      kind: 'custom',
      created_by_emp_id: String(created_by_emp_id),
      member_emp_ids: Array.from(valid),
      skipped,
    });
  } catch (err) {
    console.error('chat groups create error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/chat/groups/:id/members — add members.
// Returns `added` (count), `added_emp_ids` (the new ones), and `skipped`
// (emp_ids that don't exist in employee_master).
app.post('/api/chat/groups/:id/members', async (req, res) => {
  try {
    const id = String(req.params.id || '').trim();
    const { emp_ids } = req.body || {};
    if (!id || !Array.isArray(emp_ids)) {
      return res.status(400).json({ error: 'group id + emp_ids[] required' });
    }
    const grp = await pool.query('SELECT id FROM chat_groups WHERE id = $1 LIMIT 1', [id]);
    if (grp.rows.length === 0) return res.status(404).json({ error: 'Group not found' });
    const requested = Array.from(
      new Set(emp_ids.map((x) => String(x || '').trim()).filter(Boolean))
    );
    if (requested.length === 0) {
      return res.json({ id, added: 0, added_emp_ids: [], skipped: [] });
    }
    const verifyRes = await pool.query(
      "SELECT emp_id FROM employee_master WHERE emp_id = ANY($1::text[])",
      [requested]
    );
    const valid = new Set(verifyRes.rows.map((r) => r.emp_id));
    const skipped = requested.filter((e) => !valid.has(e));
    const addedIds = [];
    for (const mid of valid) {
      const r = await pool.query(
        'INSERT INTO chat_group_members (group_id, emp_id) VALUES ($1, $2) ON CONFLICT DO NOTHING',
        [id, mid]
      );
      if (r.rowCount > 0) addedIds.push(mid);
    }
    res.json({ id, added: addedIds.length, added_emp_ids: addedIds, skipped });
  } catch (err) {
    console.error('chat group add members error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/chat/employees?q=&emp_id=&limit=50 — search employee_master.
app.get('/api/chat/employees', async (req, res) => {
  try {
    const q = String(req.query.q || '').trim();
    const limit = Math.min(parseInt(req.query.limit, 10) || 50, 200);
    const params = [];
    let where = "status = 'Working'";
    if (q) {
      params.push('%' + q + '%');
      where +=
        " AND (full_name ILIKE $1 OR branch_name ILIKE $1 OR role ILIKE $1 OR emp_id ILIKE $1 OR mobile ILIKE $1)";
    }
    const sql = `SELECT emp_id, full_name, role, branch_name, area_name, division_name, region_name, mobile
                   FROM employee_master
                  WHERE ${where}
                  ORDER BY full_name
                  LIMIT ${limit}`;
    const r = await pool.query(sql, params);
    res.json({ employees: r.rows, count: r.rows.length });
  } catch (err) {
    console.error('chat employees search error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ===================================================================
// AI Chat history (Phase B) — keyed by user_key (phone / emp_id / role).
// Tables: ai_chat_conversations, ai_chat_messages.
// All endpoints require user_key — server only ever returns rows for
// that key, so users see only their own history.
// ===================================================================
function _aiChatUserKey(req) {
  const k = String((req.query && req.query.user_key) || (req.body && req.body.user_key) || '').trim();
  return k && k.length <= 200 ? k : '';
}

// GET /api/ai-chat/conversations?user_key=...
app.get('/api/ai-chat/conversations', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const r = await pool.query(
      `SELECT c.id, c.title, c.pinned, c.provider, c.created_at, c.updated_at,
              (SELECT COUNT(*) FROM ai_chat_messages m WHERE m.conversation_id = c.id) AS msg_count
         FROM ai_chat_conversations c
        WHERE c.user_key = $1
        ORDER BY c.pinned DESC, c.updated_at DESC
        LIMIT 100`,
      [userKey]
    );
    res.json({ conversations: r.rows });
  } catch (err) {
    console.error('ai-chat list error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/ai-chat/conversations  body: {user_key, title?, provider?}
app.post('/api/ai-chat/conversations', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const title = String((req.body && req.body.title) || 'New conversation').slice(0, 200);
    const provider = req.body && req.body.provider ? String(req.body.provider).slice(0, 40) : null;
    const r = await pool.query(
      `INSERT INTO ai_chat_conversations (user_key, title, provider)
       VALUES ($1, $2, $3)
       RETURNING id, title, pinned, provider, created_at, updated_at`,
      [userKey, title, provider]
    );
    res.json({ conversation: r.rows[0] });
  } catch (err) {
    console.error('ai-chat create error:', err);
    res.status(500).json({ error: err.message });
  }
});

// PATCH /api/ai-chat/conversations/:id  body: {user_key, title?, pinned?}
app.patch('/api/ai-chat/conversations/:id', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const id = parseInt(req.params.id, 10);
    if (!id || id < 1) return res.status(400).json({ error: 'invalid id' });
    const sets = [];
    const params = [id, userKey];
    if (req.body && typeof req.body.title === 'string') {
      params.push(req.body.title.slice(0, 200));
      sets.push('title = $' + params.length);
    }
    if (req.body && typeof req.body.pinned === 'boolean') {
      params.push(req.body.pinned);
      sets.push('pinned = $' + params.length);
    }
    if (!sets.length) return res.json({ ok: true });
    sets.push('updated_at = NOW()');
    const r = await pool.query(
      `UPDATE ai_chat_conversations SET ${sets.join(', ')}
        WHERE id = $1 AND user_key = $2
       RETURNING id, title, pinned, updated_at`,
      params
    );
    if (!r.rowCount) return res.status(404).json({ error: 'not found' });
    res.json({ conversation: r.rows[0] });
  } catch (err) {
    console.error('ai-chat patch error:', err);
    res.status(500).json({ error: err.message });
  }
});

// DELETE /api/ai-chat/conversations/:id?user_key=...
app.delete('/api/ai-chat/conversations/:id', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const id = parseInt(req.params.id, 10);
    if (!id || id < 1) return res.status(400).json({ error: 'invalid id' });
    const r = await pool.query(
      `DELETE FROM ai_chat_conversations WHERE id = $1 AND user_key = $2`,
      [id, userKey]
    );
    res.json({ deleted: r.rowCount });
  } catch (err) {
    console.error('ai-chat delete error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/ai-chat/relink  body: {from, to}
// One-shot migration when a user gains a more durable user_key (e.g. emp:X → phone:Y).
// Moves all conversations + messages from `from` to `to`, then deletes the `from` row.
app.post('/api/ai-chat/relink', async (req, res) => {
  try {
    const from = String((req.body && req.body.from) || '').trim();
    const to = String((req.body && req.body.to) || '').trim();
    if (!from || !to) return res.status(400).json({ error: 'from and to required' });
    if (from === to) return res.json({ moved: 0, note: 'noop' });
    const r = await pool.query(
      `UPDATE ai_chat_conversations SET user_key = $1 WHERE user_key = $2`,
      [to, from]
    );
    res.json({ moved: r.rowCount });
  } catch (err) {
    console.error('ai-chat relink error:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/ai-chat/conversations/:id/messages?user_key=...
app.get('/api/ai-chat/conversations/:id/messages', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const id = parseInt(req.params.id, 10);
    if (!id || id < 1) return res.status(400).json({ error: 'invalid id' });
    const own = await pool.query(
      `SELECT id FROM ai_chat_conversations WHERE id = $1 AND user_key = $2 LIMIT 1`,
      [id, userKey]
    );
    if (!own.rowCount) return res.status(404).json({ error: 'not found' });
    const r = await pool.query(
      `SELECT id, role, content, created_at
         FROM ai_chat_messages
        WHERE conversation_id = $1
        ORDER BY id ASC
        LIMIT 1000`,
      [id]
    );
    res.json({ messages: r.rows });
  } catch (err) {
    console.error('ai-chat messages list error:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/ai-chat/conversations/:id/messages  body: {user_key, role, content}
app.post('/api/ai-chat/conversations/:id/messages', async (req, res) => {
  try {
    const userKey = _aiChatUserKey(req);
    if (!userKey) return res.status(400).json({ error: 'user_key required' });
    const id = parseInt(req.params.id, 10);
    if (!id || id < 1) return res.status(400).json({ error: 'invalid id' });
    const role = String((req.body && req.body.role) || '').toLowerCase();
    const content = String((req.body && req.body.content) || '');
    if (!['user', 'assistant', 'system'].includes(role)) {
      return res.status(400).json({ error: 'invalid role' });
    }
    if (!content) return res.status(400).json({ error: 'content required' });
    const own = await pool.query(
      `SELECT id, title FROM ai_chat_conversations WHERE id = $1 AND user_key = $2 LIMIT 1`,
      [id, userKey]
    );
    if (!own.rowCount) return res.status(404).json({ error: 'not found' });
    const insert = await pool.query(
      `INSERT INTO ai_chat_messages (conversation_id, role, content)
       VALUES ($1, $2, $3)
       RETURNING id, role, content, created_at`,
      [id, role, content.slice(0, 16000)]
    );
    // Auto-title from first user message; bump updated_at always.
    if (role === 'user' && (own.rows[0].title === 'New conversation' || !own.rows[0].title)) {
      const t = content.replace(/\s+/g, ' ').trim().slice(0, 80);
      await pool.query(
        `UPDATE ai_chat_conversations SET title = $1, updated_at = NOW() WHERE id = $2`,
        [t || 'New conversation', id]
      );
    } else {
      await pool.query(`UPDATE ai_chat_conversations SET updated_at = NOW() WHERE id = $1`, [id]);
    }
    res.json({ message: insert.rows[0] });
  } catch (err) {
    console.error('ai-chat messages append error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ----- socket.io: chat namespace on the main io server ---------------
io.on('connection', (socket) => {
  // chat:join — emp joins their personal room. Optional `last_seen_ids`
  // (`{ [thread_key]: last_id }`) or `since` ISO timestamp triggers a server-side
  // missed-message replay over `chat:replay`. Caps at 200 messages to bound work.
  socket.on('chat:join', async (data, ack) => {
    try {
      const empId = data && data.emp_id ? String(data.emp_id) : '';
      if (!empId) {
        if (typeof ack === 'function') ack({ ok: false, error: 'emp_id required' });
        return;
      }
      socket.data.empId = empId;
      socket.join('emp:' + empId);
      // Optional missed-message replay.
      const lastSeenIds =
        data && typeof data.last_seen_ids === 'object' && data.last_seen_ids
          ? data.last_seen_ids
          : null;
      const since =
        data && data.since ? new Date(String(data.since)) : null;
      let missed = [];
      if (lastSeenIds) {
        const keys = Object.keys(lastSeenIds).slice(0, 50);
        const ids = keys.map((k) => Number(lastSeenIds[k]) || 0);
        if (keys.length > 0) {
          const r = await pool.query(
            `SELECT m.id, m.sender_emp_id, m.thread_key, m.body, m.sent_at,
                    m.read_by_json, m.client_msg_id,
                    em.full_name AS sender_name
               FROM chat_messages m
               LEFT JOIN employee_master em ON em.emp_id = m.sender_emp_id
               JOIN unnest($1::text[], $2::bigint[]) AS c(tk, last_id)
                 ON c.tk = m.thread_key AND m.id > c.last_id
              ORDER BY m.sent_at ASC, m.id ASC
              LIMIT 200`,
            [keys, ids]
          );
          missed = r.rows;
        }
      } else if (since && !isNaN(since.getTime())) {
        const r = await pool.query(
          `SELECT m.id, m.sender_emp_id, m.thread_key, m.body, m.sent_at,
                  m.read_by_json, m.client_msg_id,
                  em.full_name AS sender_name
             FROM chat_messages m
             LEFT JOIN employee_master em ON em.emp_id = m.sender_emp_id
            WHERE m.sent_at > $1
            ORDER BY m.sent_at ASC, m.id ASC
            LIMIT 200`,
          [since.toISOString()]
        );
        missed = r.rows;
      }
      // Filter by membership before emitting.
      if (missed.length > 0) {
        const memberCache = new Map();
        const out = [];
        for (const m of missed) {
          let isMember = memberCache.get(m.thread_key);
          if (isMember === undefined) {
            isMember = await chatIsMember(empId, m.thread_key);
            memberCache.set(m.thread_key, isMember);
          }
          if (isMember) out.push(m);
        }
        if (out.length > 0) socket.emit('chat:replay', { messages: out, count: out.length });
      }
      if (typeof ack === 'function') ack({ ok: true, replayed: missed.length });
    } catch (e) {
      console.error('chat:join error:', e.message);
      if (typeof ack === 'function') ack({ ok: false, error: e.message });
    }
  });

  socket.on('chat:send', async (data, ack) => {
    try {
      const senderId = data && data.sender_emp_id ? String(data.sender_emp_id) : '';
      if (!senderId) {
        if (typeof ack === 'function') ack({ ok: false, error: 'sender_emp_id required' });
        return;
      }
      if (!chatSendAcquireToken(senderId)) {
        if (typeof ack === 'function') ack({ ok: false, error: 'rate_limited', retry_after_ms: 1000 });
        return;
      }
      const row = await chatPostInternal({
        senderEmpId: senderId,
        threadKey: data && data.thread_key,
        body: data && data.body,
        clientMsgId: data && data.client_msg_id,
      });
      if (typeof ack === 'function') ack({ ok: true, message: row });
    } catch (e) {
      if (typeof ack === 'function') ack({ ok: false, error: e.message });
    }
  });

  socket.on('chat:read', async (data) => {
    try {
      const empId = data && data.emp_id;
      const tk = data && data.thread_key;
      const upTo = data && data.up_to_message_id;
      if (!empId || !tk || upTo === undefined) return;
      if (!chatValidEmpId(String(empId))) return;
      const ok = await chatIsMember(empId, tk);
      if (!ok) return;
      await pool.query(
        `UPDATE chat_messages
            SET read_by_json = jsonb_set(read_by_json, ARRAY[$1::text], to_jsonb($2::bigint), true)
          WHERE thread_key = $3
            AND id <= $2
            AND COALESCE((read_by_json->>$1)::bigint, 0) < $2`,
        [String(empId), Number(upTo), String(tk)]
      );
      const recipients = await chatThreadRecipients(tk);
      const seen = new Set();
      for (const rid of recipients) {
        if (!rid || seen.has(String(rid))) continue;
        seen.add(String(rid));
        io.to('emp:' + String(rid)).emit('chat:read', {
          thread_key: tk,
          emp_id: String(empId),
          up_to_message_id: Number(upTo),
        });
      }
    } catch (e) {
      console.error('chat:read socket error:', e.message);
    }
  });
});

// Hourly purge job: delete chat messages older than 7 days.
setInterval(async () => {
  try {
    const r = await pool.query(
      "DELETE FROM chat_messages WHERE sent_at < NOW() - INTERVAL '7 days'"
    );
    if (r.rowCount && r.rowCount > 0) {
      console.log('chat_messages purge: removed ' + r.rowCount + ' rows');
    }
  } catch (err) {
    console.error('chat_messages purge error:', err.message);
  }
}, 60 * 60 * 1000);
