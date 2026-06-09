/**
 * ec2Sync.js — Sync the flat EC2 `employee_master` table to a Working roster.
 *
 * Match key: emp_id. employee_master has no FK references, so absent rows are
 * hard-deleted. All kept rows get status='Working'.
 */

const MASTER_COLS = [
  "emp_id", "full_name", "role", "designation", "branch_name", "area_name",
  "area_manager", "division_name", "division_manager", "region_name",
  "mobile", "date_of_joining",
];

function _val(w, field) {
  switch (field) {
    case "emp_id": return w.emp_id;
    case "full_name": return w.name;
    case "branch_name": return w.branch;
    case "area_name": return w.area;
    case "division_name": return w.division;
    case "region_name": return w.region;
    case "date_of_joining": return w.doj;
    default: return w[field] ?? null; // role, designation, area_manager, division_manager, mobile
  }
}

/** Read-only diff of the Working set vs current employee_master. */
async function computeDiff(pool, working) {
  const { rows } = await pool.query("SELECT emp_id FROM employee_master");
  const dbIds = new Set(rows.map((r) => r.emp_id));
  const fileIds = new Set(working.map((w) => w.emp_id));
  const add = working.filter((w) => !dbIds.has(w.emp_id)).map((w) => w.emp_id);
  const update = working.filter((w) => dbIds.has(w.emp_id)).map((w) => w.emp_id);
  const del = [...dbIds].filter((id) => !fileIds.has(id));
  return {
    current: dbIds.size,
    add: add.length,
    update: update.length,
    delete: del.length,
    delete_ids: del.sort(),
    final: fileIds.size,
  };
}

/** Apply the sync in a single transaction. Returns final counts. */
async function apply(pool, working) {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");

    // Upsert every working row.
    const colList = MASTER_COLS.join(",") + ",status";
    const placeholders = [];
    const params = [];
    let i = 1;
    for (const w of working) {
      const slot = MASTER_COLS.map(() => `$${i++}`);
      slot.push("'Working'");
      placeholders.push("(" + slot.join(",") + ")");
      for (const c of MASTER_COLS) params.push(_val(w, c));
    }
    const updateSet = MASTER_COLS.filter((c) => c !== "emp_id")
      .map((c) => `${c}=EXCLUDED.${c}`)
      .concat("status='Working'")
      .join(",");
    await client.query(
      `INSERT INTO employee_master (${colList}) VALUES ${placeholders.join(",")}
       ON CONFLICT (emp_id) DO UPDATE SET ${updateSet}`,
      params
    );

    // Delete anyone not in the working set.
    const ids = working.map((w) => w.emp_id);
    const delRes = await client.query(
      "DELETE FROM employee_master WHERE NOT (emp_id = ANY($1::text[]))",
      [ids]
    );

    const final = await client.query(
      "SELECT count(*)::int AS total, count(*) FILTER (WHERE status='Working')::int AS working FROM employee_master"
    );
    await client.query("COMMIT");
    return {
      deleted: delRes.rowCount,
      total: final.rows[0].total,
      working: final.rows[0].working,
    };
  } catch (err) {
    await client.query("ROLLBACK").catch(() => {});
    throw err;
  } finally {
    client.release();
  }
}

module.exports = { computeDiff, apply };
