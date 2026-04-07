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

const app = express();
app.use(cors());
app.use(express.json());

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
  host: "127.0.0.1",
  user: "Raghunandan1157",
  password: "raghu",
  database: "postgres",
  port: 5432,
};

const pool = new Pool({ ...dbConfig, max: 10 });

// Database pool cache for multi-database support
const poolCache = {};
function getPool(dbName) {
  if (!dbName || dbName === 'postgres') return pool;
  if (!poolCache[dbName]) {
    poolCache[dbName] = new Pool({
      host: '127.0.0.1',
      port: 5432,
      user: 'Raghunandan1157',
      password: 'raghu',
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

app.post("/api/upload", upload.single("file"), async (req, res) => {
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

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0] || !row[3]) { skippedRows++; continue; }

        const regionName = normalizeRegion(String(row[0]));
        const districtName = String(row[1] || "").trim();
        const branchName = String(row[2] || "").trim();
        const empId = String(row[3]).trim();
        const officerName = String(row[4] || "").trim();

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
        for (let c = 5; c < 30; c++) {
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
      await pool.query("CREATE INDEX idx_hp_product_type_id ON hourly_performance(product_type_id)");
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

app.get("/api/employees", async (req, res) => {
  try {
    const { q } = req.query;
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
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/employees/count", async (req, res) => {
  try {
    const result = await pool.query("SELECT count(*) FROM employees");
    res.json({ count: parseInt(result.rows[0].count) });
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
  if (filters.region) { where.push(`r.region_name = $${idx++}`); params.push(filters.region); }
  if (filters.district) { where.push(`d.district_name = $${idx++}`); params.push(filters.district); }
  if (filters.branch) { where.push(`b.branch_name = $${idx++}`); params.push(filters.branch); }
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
    if (req.query.date) {
      const base = buildDailyQuery("r.region_name", "r.region_name");
      const { clause, params } = buildDailyWhere(req.query);
      const result = await pool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
      return res.json(result.rows);
    }
    const base = buildCollectionQuery("r.region_name", "r.region_name");
    const { clause, params } = buildWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
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
    const result = await pool.query("SELECT * FROM product_types ORDER BY product_type_id");
    res.json(result.rows);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/regions", async (req, res) => {
  try {
    const result = await pool.query("SELECT * FROM regions ORDER BY region_name");
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
const AWS_PANEL_TABLES = ['regions', 'districts', 'branches', 'employees', 'employee_performance', 'product_types', 'hourly_performance', 'months', 'portfolio_performance'];

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
    const result = await getPool(req.query.db).query('SELECT * FROM ' + tableName + ' LIMIT 50');
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
  const ALLOWED = ['regions','districts','branches','employees','employee_performance','product_types','hourly_performance','months','portfolio_performance'];
  const tbl = req.params.tableName;
  if (!ALLOWED.includes(tbl)) return res.status(400).json({ error: "Table not allowed" });

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

function buildPortfolioWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.month) { where.push(`m.month_label = $${idx++}`); params.push(filters.month); }
  if (filters.product_type && filters.product_type !== 'All') {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.region) { where.push(`r.region_name = $${idx++}`); params.push(filters.region); }
  if (filters.district) { where.push(`d.district_name = $${idx++}`); params.push(filters.district); }
  if (filters.branch) { where.push(`b.branch_name = $${idx++}`); params.push(filters.branch); }
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
  if (filters.product_type && filters.product_type !== "All") {
    where.push("dp.product_type_id IN (SELECT product_type_id FROM product_types WHERE product_type_name=$" + idx++ + ")"); params.push(filters.product_type);
  }
  if (filters.region) { where.push("r.region_name=$" + idx++); params.push(filters.region); }
  if (filters.district) { where.push("d.district_name=$" + idx++); params.push(filters.district); }
  if (filters.branch) { where.push("b.branch_name=$" + idx++); params.push(filters.branch); }
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
    const base = buildDailyQuery("r.region_name");
    const { clause, params } = buildDailyWhere(req.query);
    const result = await pool.query(base + clause + " GROUP BY r.region_name ORDER BY r.region_name", params);
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

// Upload daily data
app.post("/api/upload-daily", upload.single("file"), async (req, res) => {
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

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !row[0] || !row[3]) { skipped++; continue; }
        const empId = normalizeRegion(String(row[3]).trim()) === String(row[3]).trim() ? String(row[3]).trim() : String(row[3]).trim();
        if (!empId) { skipped++; continue; }

        const metrics = [];
        for (let c = 5; c < 30; c++) {
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
app.get("/api/disbursement/months", async (req, res) => {
  try {
    const result = await pool.query("SELECT DISTINCT db_month FROM disbursement ORDER BY db_month");
    // Sort by actual date order
    var months = result.rows.map(r => r.db_month);
    var monthOrder = {'Apr':1,'May':2,'Jun':3,'Jul':4,'Aug':5,'Sep':6,'Oct':7,'Nov':8,'Dec':9,'Jan':10,'Feb':11,'Mar':12};
    months.sort(function(a,b) {
      var ma = a.split('-')[0], mb = b.split('-')[0];
      return (monthOrder[ma]||0) - (monthOrder[mb]||0);
    });
    res.json(months);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

function buildDisbWhere(filters) {
  var where = [];
  var params = [];
  var idx = 1;
  if (filters.month) { where.push("d.db_month=$" + idx++); params.push(filters.month); }
  if (filters.product_name && filters.product_name !== 'All') {
    where.push("d.product_name=$" + idx++); params.push(filters.product_name);
  }
  if (filters.region) { where.push("d.region_name=$" + idx++); params.push(filters.region); }
  if (filters.district) { where.push("d.district_name=$" + idx++); params.push(filters.district); }
  if (filters.branch) { where.push("d.branch_name=$" + idx++); params.push(filters.branch); }
  if (filters.emp_id) { where.push("d.emp_id=$" + idx++); params.push(filters.emp_id); }
  return { clause: where.length ? " WHERE " + where.join(" AND ") : "", params };
}

app.get("/api/disbursement/summary", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause, params
    );
    res.json(result.rows[0] || {});
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-product", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.product_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.product_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-region", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-district", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.district_name, d.region_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.district_name, d.region_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-branch", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.branch_name, d.district_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.branch_name, d.district_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-employee", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.emp_id, d.officer_name AS name, d.branch_name, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.emp_id, d.officer_name, d.branch_name ORDER BY total_amount DESC", params
    );
    res.json(result.rows);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/api/disbursement/by-month", async (req, res) => {
  try {
    const { clause, params } = buildDisbWhere(req.query);
    const result = await pool.query(
      "SELECT d.db_month, SUM(d.disb_count)::int AS total_count, SUM(d.disb_amount) AS total_amount FROM disbursement d" + clause + " GROUP BY d.db_month", params
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

function buildV2Where(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.product_type && filters.product_type !== "All") {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.state) { where.push(`s.state_name = $${idx++}`); params.push(filters.state); }
  if (filters.division) { where.push(`dv.division_name = $${idx++}`); params.push(filters.division); }
  if (filters.area) { where.push(`a.area_name = $${idx++}`); params.push(filters.area); }
  if (filters.branch) { where.push(`b.branch_name = $${idx++}`); params.push(filters.branch); }
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

function buildV2PortfolioWhere(filters) {
  const where = [];
  const params = [];
  let idx = 1;
  if (filters.month) { where.push(`m.month_label = $${idx++}`); params.push(filters.month); }
  if (filters.product_type && filters.product_type !== 'All') {
    where.push(`pt.product_type_name = $${idx++}`); params.push(filters.product_type);
  }
  if (filters.state) { where.push(`s.state_name = $${idx++}`); params.push(filters.state); }
  if (filters.division) { where.push(`dv.division_name = $${idx++}`); params.push(filters.division); }
  if (filters.area) { where.push(`a.area_name = $${idx++}`); params.push(filters.area); }
  if (filters.branch) { where.push(`b.branch_name = $${idx++}`); params.push(filters.branch); }
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
app.post("/api/bulk-daily", express.json({limit: '50mb'}), async (req, res) => {
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

// ========== COMPARISON API ==========
app.get("/api/comparison", async (req, res) => {
  try {
    // Get all available dates with their aggregated metrics
    const result = await pool.query(`
      SELECT dp.report_date,
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
      GROUP BY dp.report_date
      ORDER BY dp.report_date
    `);
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
    await pool.query("CREATE INDEX IF NOT EXISTS idx_hp_product_type_id ON hourly_performance(product_type_id)");
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
      await client.query("CREATE INDEX idx_hp_product_type_id ON hourly_performance(product_type_id)");

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
    await client.query("CREATE INDEX idx_hp_product_type_id ON hourly_performance(product_type_id)");

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
  if (filters.region) { where.push("r.region_name = $" + idx++); params.push(filters.region); }
  if (filters.district) { where.push("d.district_name = $" + idx++); params.push(filters.district); }
  if (filters.branch) { where.push("b.branch_name = $" + idx++); params.push(filters.branch); }
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
const DAILY_PLAN_COLS = 'branch_name,date,region,district,dm_name,ftod_actual,ftod_plan,dpd_1_30_actual,dpd_1_30_plan,dpd_31_60_actual,dpd_31_60_plan,dpd_61_90_actual,dpd_61_90_plan,npa_activation,npa_closure,fy_non_start_acc,fy_non_start_plan,disb_igl_acc,disb_igl_amt,disb_fig_acc,disb_fig_amt,disb_il_acc,disb_il_amt,kyc_igl,kyc_fig,kyc_il';

// GET /api/daily-plan/reports?from=DATE&to=DATE&branch=NAME&region=NAME
app.get("/api/daily-plan/reports", async (req, res) => {
  try {
    const where = []; const params = []; let idx = 1;
    if (req.query.from) { where.push("date >= $" + idx++); params.push(req.query.from); }
    if (req.query.to) { where.push("date <= $" + idx++); params.push(req.query.to); }
    if (req.query.branch) { where.push("branch_name = $" + idx++); params.push(req.query.branch); }
    if (req.query.region) { where.push("region = $" + idx++); params.push(req.query.region); }
    if (req.query.district) { where.push("district = $" + idx++); params.push(req.query.district); }
    if (req.query.dm_name) { where.push("dm_name = $" + idx++); params.push(req.query.dm_name); }
    const sql = "SELECT " + DAILY_PLAN_COLS + ",created_at FROM daily_reports" + (where.length ? " WHERE " + where.join(" AND ") : "") + " ORDER BY date, branch_name";
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
    if (req.query.branch) { where.push("branch_name = $" + idx++); params.push(req.query.branch); }
    if (req.query.region) { where.push("region = $" + idx++); params.push(req.query.region); }
    if (req.query.district) { where.push("district = $" + idx++); params.push(req.query.district); }
    if (req.query.dm_name) { where.push("dm_name = $" + idx++); params.push(req.query.dm_name); }
    const sql = "SELECT " + DAILY_PLAN_COLS + ",created_at FROM daily_reports_achievements" + (where.length ? " WHERE " + where.join(" AND ") : "") + " ORDER BY date, branch_name";
    const result = await pool.query(sql, params);
    res.json({ data: result.rows, count: result.rowCount });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// GET /api/daily-plan/exists?date=DATE&branch=NAME&table=reports|achievements
app.get("/api/daily-plan/exists", async (req, res) => {
  try {
    const table = req.query.table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    const result = await pool.query("SELECT COUNT(*)::int AS cnt FROM " + table + " WHERE date=$1 AND branch_name=$2", [req.query.date, req.query.branch]);
    res.json({ exists: result.rows[0].cnt > 0 });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/daily-plan/save — Save daily report (plan or achievement)
app.post("/api/daily-plan/save", async (req, res) => {
  try {
    const { table, data } = req.body;
    if (!data || !data.date || !data.branch_name) {
      return res.status(400).json({ error: "Missing date or branch_name" });
    }
    const targetTable = table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    
    const cols = DAILY_PLAN_COLS.split(',');
    const values = cols.map(c => data[c] != null ? data[c] : 0);
    const placeholders = cols.map((_, i) => "$" + (i + 1));
    
    const sql = "INSERT INTO " + targetTable + " (" + cols.join(",") + ") VALUES (" + placeholders.join(",") + ") ON CONFLICT (branch_name, date) DO UPDATE SET " + cols.filter(c => c !== 'branch_name' && c !== 'date').map(c => c + "=EXCLUDED." + c).join(",");
    
    await pool.query(sql, values);
    res.json({ success: true, table: targetTable, branch: data.branch_name, date: data.date });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST /api/daily-plan/bulk-save — Save multiple branches at once
app.post("/api/daily-plan/bulk-save", async (req, res) => {
  const client = await pool.connect();
  try {
    const { table, rows } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: "No rows" });
    const targetTable = table === 'achievements' ? 'daily_reports_achievements' : 'daily_reports';
    const cols = DAILY_PLAN_COLS.split(',');
    
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

// Serve daily-reports.html
app.get("/daily-reports", (req, res) => {
  res.sendFile(__dirname + "/../daily-reports.html");
});
app.get("/daily-reports.html", (req, res) => {
  res.sendFile(__dirname + "/../daily-reports.html");
});
