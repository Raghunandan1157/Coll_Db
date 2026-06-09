/**
 * staffBackup.js — Snapshot employee_master on EC2 before a destructive sync.
 * Supabase backups are taken inside the hr_apply RPC.
 */

function _stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return (
    d.getFullYear() +
    p(d.getMonth() + 1) +
    p(d.getDate()) +
    "_" +
    p(d.getHours()) +
    p(d.getMinutes()) +
    p(d.getSeconds())
  );
}

/** CREATE TABLE employee_master_bak_<ts> AS SELECT *. Returns the table name. */
async function backupEmployeeMaster(pool) {
  const ts = _stamp();
  const table = `employee_master_bak_${ts}`;
  await pool.query(`CREATE TABLE ${table} AS SELECT * FROM employee_master`);
  const { rows } = await pool.query(`SELECT count(*)::int AS n FROM ${table}`);
  return { table, rows: rows[0].n };
}

module.exports = { backupEmployeeMaster };
