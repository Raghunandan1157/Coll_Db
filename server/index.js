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
  'BUDWAL': {state: 'ANDRA PRADESH', division: 'ANDRA PRADESH', area: 'KADAPA'},
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
  'DABUSPET': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'DAVANAGERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
  'DEVADURGA': {state: 'KALBURGI', division: 'KALBURGI', area: 'SHAHAPUR'},
  'DEVANAHALLI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'DHARMAVARAM': {state: 'ANDRA PRADESH', division: 'ANDRA PRADESH', area: 'KADAPA'},
  'DHARWAD': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'DODDABALLAPURA': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'GADAG': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'GADWAL': {state: 'TELANGANA', division: 'TELANGANA', area: 'MAHABOOBNAGAR'},
  'GAJENDRAGAD': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'GANGAVATHI': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'GOKAK': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'GOWRIBIDANUR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'DODDABALLAPURA'},
  'GUBBI': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'HAGARIBOMMANAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'HARAPANAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'HARIHARA': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'DAVANAGERE'},
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
  'KADAPA': {state: 'ANDRA PRADESH', division: 'ANDRA PRADESH', area: 'KADAPA'},
  'KADIRI': {state: 'ANDRA PRADESH', division: 'ANDRA PRADESH', area: 'KADAPA'},
  'KADUR': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'KALABURAGI': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'KALAGI': {state: 'KALBURGI', division: 'BIDAR', area: 'SEDAM'},
  'KALBURGI-2': {state: 'KALBURGI', division: 'KALBURGI', area: 'KALBURGI'},
  'KALGHATGI': {state: 'DHARWAD', division: 'HUBLI', area: 'DHARWAD'},
  'KAMALAPURA': {state: 'KALBURGI', division: 'BIDAR', area: 'HUMNABAD'},
  'KENGERI': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'BANGALORE URBAN'},
  'KHANAHOSAHALLI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'KITTUR': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'KODANGAL': {state: 'TELANGANA', division: 'TELANGANA', area: 'SANGAREDDY'},
  'KOLAR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'KOPPAL': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'KORATAGERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'KOTTURU': {state: 'CHITRADURGA', division: 'HOSPET', area: 'KOTTURU'},
  'KUDATHINI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'KUDLIGI': {state: 'CHITRADURGA', division: 'HOSPET', area: 'HOSPET'},
  'KUNIGAL': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'KUSHTAGI': {state: 'DHARWAD', division: 'HUBLI', area: 'KUSHTAGI'},
  'LAXMESHWAR': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'LINGSUGUR': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'LOKAPUR': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BAGALKOT'},
  'MADHUGIRI': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'MAHABUB NAGAR': {state: 'TELANGANA', division: 'TELANGANA', area: 'MAHABOOBNAGAR'},
  'MALUR': {state: 'TUMKUR', division: 'DODDABALLAPURA', area: 'KOLAR'},
  'MANVI': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'MARIKAL': {state: 'TELANGANA', division: 'TELANGANA', area: 'MAHABOOBNAGAR'},
  'MUDALAGI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'MUDDEBIHAL': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'MUDIGERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'CHIKKAMAGALURU'},
  'MUNDARAGI': {state: 'DHARWAD', division: 'HUBLI', area: 'GADAG'},
  'NARAGUNDA': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'NARAYANKHED': {state: 'TELANGANA', division: 'TELANGANA', area: 'SANGAREDDY'},
  'NIPPANI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'CHIKKODI'},
  'NR PURA': {state: 'TUMKUR', division: 'TUMKUR', area: 'CHIKKAMAGALURU'},
  'PANCHANHALLI': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'RAICHUR': {state: 'KALBURGI', division: 'BIDAR', area: 'LINGSUGUR'},
  'RAMDURGA': {state: 'DHARWAD', division: 'HUBLI', area: 'BADAMI'},
  'SANDURU': {state: 'CHITRADURGA', division: 'HOSPET', area: 'BALLARI'},
  'SANGAREDDY': {state: 'TELANGANA', division: 'TELANGANA', area: 'SANGAREDDY'},
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
  'TANDUR': {state: 'TELANGANA', division: 'TELANGANA', area: 'MAHABOOBNAGAR'},
  'TARIKERE': {state: 'CHITRADURGA', division: 'CHITRADURGA', area: 'KADUR'},
  'TIKOTA': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'TIPTUR': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'TUMKUR': {state: 'TUMKUR', division: 'TUMKUR', area: 'TUMKUR'},
  'TUREVEKERE': {state: 'TUMKUR', division: 'TUMKUR', area: 'TIPTUR'},
  'VIJAYAPUR': {state: 'KALBURGI', division: 'KALBURGI', area: 'VIJAYAPUR'},
  'YADGIR': {state: 'KALBURGI', division: 'KALBURGI', area: 'SHAHAPUR'},
  'YARAGATTI': {state: 'DHARWAD', division: 'BELAGAVI', area: 'BELAGAVI'},
  'ZAHEERABAD': {state: 'TELANGANA', division: 'TELANGANA', area: 'SANGAREDDY'}
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

        // Performance metrics — safely parse each value
        const metrics = [];
        for (let c = colMetricsStart; c < colMetricsStart + 25; c++) {
          const raw = row[c];
          if (raw == null || raw === "") { metrics.push(0); continue; }
          const num = Number(raw);
          metrics.push(Number.isFinite(num) ? num : 0);
        }
        // Pad to 25 if fewer columns
        while (metrics.length < 25) metrics.push(0);

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
  // Hierarchy filters — resolve via employee_master
  if (filters.region || filters.state) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.region || filters.state);
  }
  if (filters.division) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.division);
  }
  if (filters.district || filters.area) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.district || filters.area);
  }
  if (filters.branch) {
    where.push(`UPPER(b.branch_name) = UPPER($${idx++})`);
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
function buildDailyQuery(groupCol) {
  // Simple query directly from daily_performance — no hierarchy joins needed
  // The upload-daily stores all data with emp_id and product_type_id
  // For summary (no groupCol): just SUM everything
  // For grouping: LEFT JOIN through hierarchy tables (tolerant of missing employees)
  // Always add hierarchy joins — they're needed for WHERE filters too (branch, district, region)
  var needsHierarchy = true;
  var joins = needsHierarchy
    ? ` LEFT JOIN employees e ON dp.emp_id = e.emp_id
        LEFT JOIN branches b ON e.branch_id = b.branch_id
        LEFT JOIN districts d ON b.district_id = d.district_id
        LEFT JOIN regions r ON d.region_id = r.region_id`
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

function buildDailyWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
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
  if (filters.emp_id) { where.push("dp.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/daily/summary", async (req, res) => {
  try {
    const base = buildDailyQuery(null);
    const { clause, params } = buildDailyWhere(req.query);
    const sql = base.replace("SELECT ,", "SELECT ") + clause;
    const result = await pool.query(sql, params);
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-region", async (req, res) => {
  try {
    var useOld = req.query.structure === 'old';
    var regionCol = useOld ? "b.old_region AS region_name" : "r.region_name";
    var groupCol = useOld ? "b.old_region" : "r.region_name";
    var nullFilter = useOld ? " AND b.old_region IS NOT NULL" : " AND r.region_name IS NOT NULL";
    const base = buildDailyQuery(regionCol);
    const { clause, params } = buildDailyWhere(req.query);
    const result = await pool.query(base + clause + nullFilter + " GROUP BY " + groupCol + " ORDER BY " + groupCol, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-district", async (req, res) => {
  try {
    const base = buildDailyQuery("d.district_name, r.region_name");
    const { clause, params } = buildDailyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY d.district_name, r.region_name ORDER BY d.district_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-branch", async (req, res) => {
  try {
    const base = buildDailyQuery("b.branch_name, d.district_name");
    const { clause, params } = buildDailyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY b.branch_name, d.district_name ORDER BY b.branch_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-employee", async (req, res) => {
  try {
    const base = buildDailyQuery("e.emp_id, e.officer_name AS name, b.branch_name");
    const { clause, params } = buildDailyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY e.emp_id, e.officer_name, b.branch_name ORDER BY e.officer_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// New employee_master-driven hierarchy endpoints (used by Priority Panel)
app.get("/api/daily/by-area", async (req, res) => {
  try {
    const base = buildDailyQuery("em.area_name, em.division_name") + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id";
    const { clause, params } = buildDailyWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.area_name IS NOT NULL";
    const result = await pool.query(base + clause + extra + " GROUP BY em.area_name, em.division_name ORDER BY em.area_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-division", async (req, res) => {
  try {
    const base = buildDailyQuery("em.division_name, em.region_name") + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id";
    const { clause, params } = buildDailyWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.division_name IS NOT NULL";
    const result = await pool.query(base + clause + extra + " GROUP BY em.division_name, em.region_name ORDER BY em.division_name", params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/daily/by-state", async (req, res) => {
  try {
    const base = buildDailyQuery("em.region_name AS state_name") + " LEFT JOIN employee_master em ON dp.emp_id = em.emp_id";
    const { clause, params } = buildDailyWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.region_name IS NOT NULL";
    const result = await pool.query(base + clause + extra + " GROUP BY em.region_name ORDER BY em.region_name", params);
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

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0] || !row[dEmpId]) { skipped++; continue; }
        const empId = String(row[dEmpId]).trim();
        if (!empId) { skipped++; continue; }

        const metrics = [];
        for (let c = dMetrics; c < dMetrics + 25; c++) {
          const raw = row[c];
          if (raw == null || raw === '') { metrics.push(0); continue; }
          const num = Number(raw);
          metrics.push(Number.isFinite(num) ? num : 0);
        }
        while (metrics.length < 25) metrics.push(0);

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
function buildDisbWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
  if (filters.month) { where.push("d.db_month=$" + idx++); params.push(filters.month); }
  if (filters.product_name && filters.product_name !== 'All') {
    where.push("d.product_name=$" + idx++); params.push(filters.product_name);
  }
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
  if (filters.branch) {
    where.push("UPPER(d.branch_name) = UPPER($" + (idx++) + ")");
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push("d.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/disbursement/summary", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-product", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT d.product_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.product_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-region", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-district", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT d.district_name, d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.district_name, d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-branch", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT d.branch_name, d.district_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.branch_name, d.district_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-employee", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      DISB_CTE + "SELECT d.emp_id, d.officer_name AS name, d.branch_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM d" + clause + " GROUP BY d.emp_id, d.officer_name, d.branch_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// New-structure endpoints (join employee_master for state/division/area)
app.get("/api/disbursement/by-state", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.region_name IS NOT NULL";
    const result = await pool.query(
      DISB_CTE + "SELECT em.region_name AS state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + extra + " GROUP BY em.region_name ORDER BY em.region_name", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-division", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.division_name IS NOT NULL";
    const result = await pool.query(
      DISB_CTE + "SELECT em.division_name, em.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + extra + " GROUP BY em.division_name, em.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-area", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const extra = (clause ? " AND " : " WHERE ") + "em.area_name IS NOT NULL";
    const result = await pool.query(
      DISB_CTE + "SELECT em.area_name, em.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + extra + " GROUP BY em.area_name, em.division_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-month", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
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
function buildDisbDailyWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
  if (filters.from && filters.to) {
    where.push("d.disb_date BETWEEN $" + idx++ + " AND $" + idx++);
    params.push(filters.from, filters.to);
  } else if (filters.date) { where.push("d.disb_date=$" + idx++); params.push(filters.date); }
  if (filters.product_name && filters.product_name !== 'All') {
    where.push("d.product_name=$" + idx++); params.push(filters.product_name);
  }
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
  if (filters.branch) {
    where.push("UPPER(d.branch_name) = UPPER($" + (idx++) + ")");
    params.push(filters.branch);
  }
  if (filters.emp_id) { where.push("d.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/disbursement/daily/dates", async (req, res) => {
  try {
    // Allow scope filtering but ignore `date`/`from`/`to` themselves for the list
    var q = Object.assign({}, req.query); delete q.date; delete q.from; delete q.to;
    const { clause, params } = buildDisbDailyWhere(q);
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
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-product", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT d.product_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.product_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-region", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-district", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT d.district_name, d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.district_name, d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-branch", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT d.branch_name, d.district_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.branch_name, d.district_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-employee", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const result = await pool.query(
      "SELECT d.emp_id, d.officer_name AS name, d.branch_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement_daily d" + clause + " GROUP BY d.emp_id, d.officer_name, d.branch_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Daily disbursement aggregated by state (new-structure CEO view).
// disbursement_daily only carries branch_name; we derive state via employee_master.
app.get("/api/disbursement/daily/by-state", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const sql =
      "SELECT em.region_name AS state_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + (clause ? " AND " : " WHERE ") + "em.region_name IS NOT NULL " +
      "GROUP BY em.region_name ORDER BY em.region_name";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Daily disbursement aggregated by area / division (joined via employee_master since
// disbursement_daily only carries region/district/branch labels).
app.get("/api/disbursement/daily/by-area", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const sql =
      "SELECT em.area_name, em.division_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + (clause ? " AND " : " WHERE ") + "em.area_name IS NOT NULL " +
      "GROUP BY em.area_name, em.division_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-division", async (req, res) => {
  try {
    const { clause, params } = buildDisbDailyWhere(req.query);
    const sql =
      "SELECT em.division_name, em.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount " +
      "FROM disbursement_daily d LEFT JOIN employee_master em ON UPPER(d.branch_name)=UPPER(em.branch_name)" +
      clause + (clause ? " AND " : " WHERE ") + "em.division_name IS NOT NULL " +
      "GROUP BY em.division_name, em.region_name ORDER BY total_amount DESC";
    const result = await pool.query(sql, params);
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/daily/by-date-range", async (req, res) => {
  try {
    const from = req.query.from, to = req.query.to;
    if (!from || !to) return res.status(400).json({ error: "from and to are required (YYYY-MM-DD)" });

    // Reuse scope filters — strip date so it doesn't collide with range
    var q = Object.assign({}, req.query); delete q.date; delete q.from; delete q.to;
    const { clause, params } = buildDisbDailyWhere(q);

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
    if (wb.SheetNames.includes('POS')) {
      try {
        await portfolioPool.query("DELETE FROM branch_pos WHERE month_id=$1", [monthId]);
        const posWs = wb.Sheets['POS'];
        const posRows = XLSX.utils.sheet_to_json(posWs, { header: 1 });
        let posInserted = 0;
        for (let r = 1; r < posRows.length; r++) {
          const row = posRows[r];
          if (!row || !row[0]) continue;
          const region = String(row[0]).trim();
          const district = String(row[1] || '').trim();
          const branch = String(row[2] || '').trim();
          const productName = String(row[3] || 'ALL').trim();
          const vals = [];
          for (let c = 4; c < 10; c++) {
            const v = Number(row[c]);
            vals.push(Number.isFinite(v) ? v : 0);
          }
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
        console.log("POS sheet: " + posInserted + " branches inserted into branch_pos");
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
        let empPosInserted = 0;
        for (let r = 1; r < empPosRows.length; r++) {
          const row = empPosRows[r];
          if (!row || !row[0]) continue;
          const empId = String(row[0]).trim();
          if (!empId) continue;
          const vals = [];
          for (let c = 1; c < 7; c++) {
            const v = Number(row[c]);
            vals.push(Number.isFinite(v) ? v : 0);
          }
          while (vals.length < 6) vals.push(0);
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
        console.log("EMP_POS sheet: " + empPosInserted + " employees into employee_pos");
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

// Resolve a hierarchy filter (state/region, division, area, branch) to a set of
// v2_branches via employee_master — single source of truth, case-insensitive.
function buildV2Where(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  // Hierarchy filters — resolve via employee_master
  if (filters.state || filters.region) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(region_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.state || filters.region);
  }
  if (filters.division) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(division_name) ILIKE TRIM($${idx++}))`);
    params.push(filters.division);
  }
  if (filters.area || filters.district) {
    where.push(`UPPER(b.branch_name) IN (SELECT UPPER(branch_name) FROM employee_master WHERE TRIM(area_name) ILIKE TRIM($${idx++}))`);
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
      FROM daily_performance dp
      LEFT JOIN employees e ON dp.emp_id = e.emp_id
      LEFT JOIN branches b ON e.branch_id = b.branch_id
      LEFT JOIN districts d ON b.district_id = d.district_id
      LEFT JOIN regions r ON d.region_id = r.region_id`;
    const { clause, params } = buildDailyWhere({...req.query, date: undefined, scope: req.query.scope || 'oa'});
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
        for (let r = 1; r < rows.length; r++) {
          const row = rows[r];
          if (!row || !row[3]) { skippedRows++; continue; }
          const empId = String(row[3]).trim();
          if (!empId || !empSet.has(empId)) { skippedRows++; continue; }
          const metrics = [];
          for (let c = 5; c < 30; c++) {
            const raw = row[c];
            if (raw == null || raw === "") { metrics.push(0); continue; }
            const num = Number(raw);
            metrics.push(Number.isFinite(num) ? num : 0);
          }
          while (metrics.length < 25) metrics.push(0);
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
    // Quick Report column layout (0-indexed):
    //   0=EMP ID  1=Name  2=Reg Demand  3=Reg Collection  4=FTOD  5=Coll%
    //   6=130 Demand  7=130 Collection  8=130 Balance  9=130 Coll%
    //   10=3160 Demand  11=3160 Collection  12=3160 Balance  13=3160 Coll%
    //   14=PNPA Demand  15=PNPA Collection  16=PNPA Balance  17=PNPA Coll%
    //   18=NPA Cases  19=NPA 90+ Acc  20=NPA 90+ Amt
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
        npa_cases: safeNum(row[18]),
        npa_act_acc: safeNum(row[19]),
        npa_act_amt: safeNum(row[20]),
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
async function safeQuery(sql, params, maxRows) {
  const client = await pool.connect();
  try {
    await client.query("SET statement_timeout = '5000'"); // 5s max
    const result = await client.query(sql, params);
    return result.rows.slice(0, maxRows || 50);
  } catch (e) {
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

  // Latest dates available
  const dates = await safeQuery("SELECT MAX(report_date) as latest FROM daily_performance", [], 1);
  ctx.latestDate = (Array.isArray(dates) && dates[0]?.latest) || 'unknown';

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

  return ctx;
}

// Compact reference of the underlying schema. Used by /api/ai-snapshot so
// offline mobile clients carry the same DB cheatsheet as the AI prompt.
const AI_SCHEMA_CHEATSHEET = {
  branches: 'branch_id, branch_name, district_id',
  employees: 'emp_id, full_name, role, branch_id',
  employee_master: 'emp_id, full_name, role, designation, branch_name, area_name, area_manager, division_name, division_manager, region_name, mobile, status (HR roster, single source of truth for hierarchy)',
  employee_performance: 'emp_id, regular_demand, regular_collection, demand_1_30, collection_1_30, demand_31_60, collection_31_60, pnpa_demand, pnpa_collection, npa_cases, npa_act_acc, npa_act_amt, ... (rolling totals; NO date column)',
  daily_performance: 'emp_id, report_date, same metric columns as employee_performance (historical daily snapshots)',
  disbursement: 'db_month, region_name, district_name, branch_name, emp_id, officer_name, product_name, disb_count, disb_amount (monthly aggregate)',
  disbursement_daily: 'disb_date, region_name, district_name, branch_name, emp_id, officer_name, product_name, disb_count, disb_amount (daily — UNIONed via DISB_CTE for months not yet rolled into disbursement)',
  daily_reports: 'branch-level plan + achievement entries',
  npa_activation_runs: 'run_id, source_filename, report_date, row_count, npa_count',
  hourly_performance: 'intra-day collection snapshots',
  employee_locations: 'emp_id, lat, lng, accuracy, battery_pct, recorded_at (live tracking pings)',
  chat_messages: 'id, sender_emp_id, thread_key, body, sent_at, read_by_json',
  chat_groups: 'id, name, kind (auto|custom), scope_type, scope_value, created_by_emp_id',
  chat_group_members: 'group_id, emp_id (FK to chat_groups)'
};

const AI_SNAPSHOT_VERSION = 1;

// Bump when AI prompt or tool surface changes — invalidates aiReplyCache keys.
const AI_REPLY_CACHE_VERSION = 'v6-ambiguous-entity-lookup';

// Mistral function-calling tool definitions. The model picks which to call
// when the bundled ctx isn't enough. Every tool maps to a parameterised SQL
// query in dispatchAiTool() below, scope-respecting via _scopeWhere().
const AI_TOOLS_SPEC = [
  {
    type: 'function',
    function: {
      name: 'find_employee',
      description: 'Look up employees by name, mobile, or emp_id. Returns matching rows from employee_master with role, branch, area, division, region, mobile, status. Use for "who is X" / "find Y" / "what is Z\'s mobile/branch/role". If multiple matches are returned for a name, do not choose one; list the matches and ask which employee the user means.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Substring of full_name (ILIKE), or exact mobile, or exact emp_id.' },
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
      description: 'Collection/demand/NPA totals for ONE employee over a date range. Default range = current FY-to-date.',
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
      description: 'Leaderboard of employees by metric over a date range.',
      parameters: {
        type: 'object',
        properties: {
          metric: { type: 'string', enum: ['collection', 'demand', 'npa_cases'] },
          start_date: { type: 'string' },
          end_date: { type: 'string' },
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
      description: 'Disbursement count + amount over a date range. Group by month, branch, product, or employee. Optional filters: branch_name / region_name.',
      parameters: {
        type: 'object',
        properties: {
          start_date: { type: 'string', description: 'YYYY-MM-DD; reduced to month boundary.' },
          end_date: { type: 'string' },
          group_by: { type: 'string', enum: ['month', 'branch', 'product', 'employee'] },
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
      description: 'List children of a hierarchy node from employee_master. Examples: list_hierarchy(level="branch", parent_level="area", parent_name="X") → branches under area X. list_hierarchy(level="region") → all regions in scope.',
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
      name: 'npa_summary',
      description: 'NPA cases + activated count + amount over a date range. Group by month, branch, or employee. Optional filters: branch_name / region_name.',
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
      name: 'sql_describe',
      description: 'Returns the schema cheatsheet of all tables/columns the AI can query.',
      parameters: { type: 'object', properties: {} }
    }
  }
];

// Build " AND UPPER(<alias>) IN (SELECT ... FROM employee_master WHERE ...)"
// for the current session. Mutates `params` to append $loc. CEO/empty session → ''.
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

async function dispatchAiTool(name, args, session) {
  args = args || {};
  const lim = Math.min(Math.max(parseInt(args.limit, 10) || 25, 1), 200);

  if (name === 'sql_describe') return AI_SCHEMA_CHEATSHEET;

  if (name === 'find_employee') {
    const q = String(args.query || '').trim();
    if (!q) return { error: 'query required' };
    const params = [`%${q}%`, q];
    const scopeClause = _scopeWhere(session, 'em.branch_name', params);
    const sql = `
      SELECT em.emp_id, em.full_name AS name, em.role, em.designation,
             em.branch_name AS branch, em.area_name AS area,
             em.division_name AS division, em.region_name AS region,
             em.mobile, em.status
        FROM employee_master em
       WHERE (em.full_name ILIKE $1 OR em.mobile = $2 OR em.emp_id = $2)
         ${scopeClause}
       ORDER BY (em.status='Working') DESC, em.full_name
       LIMIT ${lim}`;
    const rows = await safeQuery(sql, params, lim);
    if (!Array.isArray(rows)) return rows;

    const exactMatches = rows.filter((r) => isExactEmployeeLookup(r, q) || isExactNameMatch(r.name, q));
    if (rows.length > 1 && exactMatches.length !== 1) {
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
    if (exactMatches.length === 1) return exactMatches;
    return rows;
  }

  if (name === 'employee_performance') {
    const empId = String(args.emp_id || '').trim();
    if (!empId) return { error: 'emp_id required' };
    const now = new Date();
    const fy = now.getMonth() + 1 >= 4 ? now.getFullYear() : now.getFullYear() - 1;
    const startDate = args.start_date || `${fy}-04-01`;
    const endDate = args.end_date || now.toISOString().slice(0, 10);
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
       GROUP BY em.emp_id, em.full_name, em.role, em.branch_name,
                em.area_name, em.division_name, em.region_name, em.mobile, em.status`;
    return await safeQuery(sql, [empId, startDate, endDate], 1);
  }

  if (name === 'find_branch') {
    const q = String(args.query || '').trim();
    if (!q) return { error: 'query required' };
    const params = [`%${q}%`];
    const scopeClause = _scopeWhere(session, 'em.branch_name', params);
    const sql = `
      WITH em_agg AS (
        SELECT em.branch_name,
               MAX(em.region_name) AS region,
               MAX(em.division_name) AS division,
               MAX(em.area_name) AS area,
               COUNT(*) FILTER (WHERE em.status='Working')::int AS employee_count
          FROM employee_master em
         WHERE em.branch_name ILIKE $1
           ${scopeClause}
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
         WHERE b.branch_name ILIKE $1
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
    const rows = await safeQuery(sql, params, lim);
    if (!Array.isArray(rows)) return rows;
    const exactMatches = rows.filter((r) => isExactNameMatch(r.branch_name, q));
    if (rows.length > 1 && exactMatches.length !== 1) {
      return {
        ambiguous: true,
        count: rows.length,
        instruction: 'Multiple branches match this lookup. List the matching branches with region, division, area, and employee_count, then ask which branch the user means. Do not answer as if only one branch matched.',
        matches: rows
      };
    }
    if (exactMatches.length === 1) return exactMatches;
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
    const scopeClause = _scopeWhere(session, 'b.branch_name', params);
    const orderCol = metric === 'collection' ? 'SUM(dp.regular_collection)' :
                     metric === 'demand'     ? 'SUM(dp.regular_demand)' :
                                               'SUM(dp.npa_cases)';
    const sql = `
      SELECT em.emp_id, em.full_name AS name, em.branch_name AS branch,
             em.region_name AS region,
             SUM(dp.regular_demand)::bigint AS demand,
             SUM(dp.regular_collection)::bigint AS collection,
             SUM(dp.npa_cases)::int AS npa_cases
        FROM daily_performance dp
        JOIN employees e ON dp.emp_id = e.emp_id
        JOIN branches b ON e.branch_id = b.branch_id
        JOIN employee_master em ON em.emp_id = dp.emp_id
       WHERE dp.report_date BETWEEN $1 AND $2
         ${scopeClause}
       GROUP BY em.emp_id, em.full_name, em.branch_name, em.region_name
       ORDER BY ${orderCol} DESC NULLS LAST
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
  }

  if (name === 'disbursement_query') {
    const start = args.start_date, end = args.end_date;
    if (!start || !end) return { error: 'start_date and end_date required' };
    const groupBy = ['month', 'branch', 'product', 'employee'].includes(args.group_by) ? args.group_by : 'month';
    const params = [start, end];
    let extra = '';
    if (args.branch_name) { params.push(args.branch_name); extra += ` AND d.branch_name ILIKE $${params.length}`; }
    if (args.region_name) { params.push(args.region_name); extra += ` AND d.region_name ILIKE $${params.length}`; }
    const scopeClause = _scopeWhere(session, 'd.branch_name', params);
    let selectCols, groupCols;
    if (groupBy === 'month') {
      selectCols = `d.db_month AS bucket`;
      groupCols = `d.db_month`;
    } else if (groupBy === 'branch') {
      selectCols = `d.branch_name AS bucket`;
      groupCols = `d.branch_name`;
    } else if (groupBy === 'product') {
      selectCols = `d.product_name AS bucket`;
      groupCols = `d.product_name`;
    } else {
      selectCols = `d.officer_name AS bucket, d.emp_id`;
      groupCols = `d.officer_name, d.emp_id`;
    }
    const sql = `${DISB_CTE}
      SELECT ${selectCols},
             SUM(d.disb_count)::int AS count,
             SUM(d.disb_amount)::numeric AS amount
        FROM d
       WHERE d.db_month BETWEEN to_char($1::date, 'YYYY-MM') AND to_char($2::date, 'YYYY-MM')
         ${extra}
         ${scopeClause}
       GROUP BY ${groupCols}
       ORDER BY 1
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
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
       LIMIT ${lim}`;
    return await safeQuery(sql, params, lim);
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
      if (args.region_name) { params.push(args.region_name); extra += ` AND region ILIKE $${params.length}`; }
      if (args.district_name) { params.push(args.district_name); extra += ` AND district ILIKE $${params.length}`; }
      const scopeClause = _scopeWhere(session, 'branch_name', params);
      const sql = `SELECT date, branch_name, region, district, dm_name, ${projectCols.join(', ')}
                     FROM ${t}
                    WHERE date BETWEEN $1 AND $2
                      ${extra}
                      ${scopeClause}
                    ORDER BY date, branch_name
                    LIMIT ${lim}`;
      const rows = await safeQuery(sql, params, lim);
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
    res.json({
      version: AI_SNAPSHOT_VERSION,
      generated_at: new Date().toISOString(),
      emp_id: body.emp_id || null,
      session,
      schema: AI_SCHEMA_CHEATSHEET,
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
const MISTRAL_MODEL = 'mistral-small-latest';
const crypto = require('crypto');

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
async function runMistralWithTools(messages, session, maxRounds) {
  const convo = messages.slice();
  const max = maxRounds || 12;
  for (let round = 0; round < max; round++) {
    const r = await callMistralAi(convo, AI_TOOLS_SPEC);
    if (!r.ok) {
      console.error(`[ai-tools] round ${round + 1}/${max} call failed: ${r.error}`);
      return r;
    }
    const msg = r.message || {};
    const calls = Array.isArray(msg.tool_calls) ? msg.tool_calls : null;
    if (!calls || !calls.length) {
      console.log(`[ai-tools] round ${round + 1}/${max} done (text len ${(msg.content || '').length})`);
      return { ok: true, text: msg.content || '' };
    }
    console.log(`[ai-tools] round ${round + 1}/${max} calling ${calls.length} tool(s): ${calls.map(c => c?.function?.name).join(', ')}`);
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
      let toolResult;
      try {
        toolResult = await dispatchAiTool(fnName, argObj, session);
      } catch (e) {
        toolResult = { error: 'tool_threw: ' + e.message };
      }
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
      if (text) return { ok: true, text, model };
      console.error('Gemini model ' + model + ' empty response:', JSON.stringify(result).slice(0, 200));
    } catch (e) {
      console.error('Gemini model ' + model + ' failed:', e.message);
    }
  }
  return { ok: false, error: 'gemini_all_models_failed' };
}

app.post("/api/ai-chat", aiLimiter, async (req, res) => {
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

  const { messages, role, location } = req.body || {};
  if (!messages || !Array.isArray(messages)) {
    return res.status(400).json({ error: 'messages array required' });
  }

  // Provider selection: 'mistral' (default per spec) | 'gemini'. Anything else
  // is normalised to 'mistral'. The other provider is the fallback on failure.
  const requested = String((req.body && req.body.provider) || 'mistral').toLowerCase();
  const primary = requested === 'gemini' ? 'gemini' : 'mistral';
  const fallback = primary === 'mistral' ? 'gemini' : 'mistral';

  if (!MISTRAL_KEY && !GEMINI_KEY) {
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
      'Data is scoped to the user\'s role: CEO=all branches; RM/SM=region; DM/DvM=division; AM=area; BM/FO=branch.',
      'Every tool call is automatically scope-filtered. You do not need to add scope filters yourself.',
      '',
      '## Units',
      'All monetary columns (regular_demand, regular_collection, npa_act_amt, disb_amount) are stored in **rupees** (raw integers, NOT lakhs). When formatting: divide by 1,00,000 for L; by 1,00,00,000 for Cr.',
      'Count columns (regular_demand, regular_collection when used as account counts, npa_cases, npa_act_acc, npa_clo_acc, disb_count, KYC counts, DPD/FTOD counts) are plain counts. Never format count columns as ₹, L, or Cr.',
      '',
      '## Tools available — CALL THESE for any specifics not in the at-a-glance JSON',
      '- find_employee(query, limit?) — name/mobile/emp_id substring lookup.',
      '- employee_performance(emp_id, start_date?, end_date?) — one employee\'s totals over a period (default = FY-to-date).',
      '- find_branch(query, limit?) — branch name lookup with hierarchy + headcount + perf.',
      '- period_performance(start_date, end_date, group_by?, branch_name?, area_name?, division_name?, region_name?, limit?) — collection/demand/NPA over a date range, group by day|month|branch|employee.',
      '- top_performers(metric, start_date, end_date, limit?) — leaderboard by collection|demand|npa_cases.',
      '- disbursement_query(start_date, end_date, group_by?, branch_name?, region_name?, limit?) — disbursement count + amount over range, group by month|branch|product|employee.',
      '- list_hierarchy(level, parent_level?, parent_name?, limit?) — list region|division|area|branch entities, optionally under a parent.',
      '- npa_summary(start_date, end_date, group_by?, branch_name?, region_name?, limit?) — NPA cases + activation amount over range.',
      '- daily_reports_query(start_date, end_date, branch_name?, region_name?, district_name?, table?, metrics?, limit?) — branch-level DAILY PLAN data: FTOD, DPD bucket 1-30/31-60/61-90, NPA activation/closure, disbursement plan vs actual (IGL/FIG/IL), KYC. ALWAYS use this for ANY question mentioning FTOD, DPD, KYC, NPA closure, or disbursement plan-vs-achievement on a specific date or branch.',
      '- sql_describe() — schema cheatsheet refresher.',
      '',
      '## Date range guidance',
      `- "today" → ${ctx.now}.`,
      `- "this month" → first of current month → ${ctx.now}.`,
      '- "last month" → first to last of prior month.',
      `- "FY-to-date" / "this year" → ${ctx.fyStart} → ${ctx.now}.`,
      '- "last quarter" → prior 3 calendar months.',
      `- A bare month name ("July") → that month in the current FY year (FY runs Apr→Mar). FY start = ${ctx.fyStart}.`,
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
      '- For employee-name lookups: if find_employee returns more than one match or `ambiguous: true`, list the matching employees with emp_id + branch/role and ask which one the user means. Never pick one silently for common names like Karthik. Do not list phone numbers while disambiguating multiple people.',
      '- For branch/entity lookups: if find_branch returns more than one match or `ambiguous: true`, list the matching branches with region/division/area and ask which one the user means. Never pick one silently for partial branch names.',
      '- Never invent numbers. Quote what tools return.',
      '- Never use the words "snapshot" or "provided data" or "JSON below" in your replies — they leak internal plumbing. Just answer with the numbers and a short label.',
      '- If a question mentions FTOD, DPD, KYC, disbursement plan, NPA closure, or "daily plan" → MUST call daily_reports_query. Do NOT say "data not available" without calling the tool first.',
      '- "11th April" / "April 11" / "11/04" all mean the same date — convert to YYYY-MM-DD using the current FY year.',
      '- When a tool returns 0 rows for a specific date, ALWAYS retry the same tool with a wider window (the full month, then the full FY) to find the nearest available date(s). Then tell the user "no data for {requested date}; nearest available is {date} — here it is" and answer with that data. Do NOT just say "data not available" and stop.',
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

  // Try primary, fall back to the other on failure.
  const order = [primary, fallback];
  let lastErr = null;
  for (const which of order) {
    const result = which === 'mistral'
      ? await runMistralWithTools(mergedMessages, session)
      : await callGeminiAi(mergedMessages);
    if (result.ok && result.text) {
      const providerLabel = which === 'mistral' ? MISTRAL_MODEL : (result.model || 'gemini');
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
    lastErr = result.error;
    console.error('AI provider ' + which + ' failed:', result.error);
  }

  res.status(429).json({
    error: 'AI is briefly busy. Please retry in a moment.',
    detail: lastErr,
  });
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
