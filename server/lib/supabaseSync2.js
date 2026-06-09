/**
 * supabaseSync2.js — Sync the second Supabase project's flat public.employees
 * roster (NLPL via Readdy / asset-stock app). Upsert-only: add new + refresh
 * details, no deletes (flat table, FK-protected, no status column).
 *
 * Requires SUPABASE2_URL + SUPABASE2_SERVICE_KEY in the environment.
 */

const URL = process.env.SUPABASE2_URL || "";
const KEY = process.env.SUPABASE2_SERVICE_KEY || "";

function isConfigured() {
  return Boolean(URL && KEY);
}

// Only the fields this flat table needs.
function _slim(working) {
  return working.map((w) => ({
    emp_id: w.emp_id,
    name: w.name,
    role: w.role,
    mobile: w.mobile,
    branch: w.branch, // -> location
  }));
}

async function _rpc(fn, working) {
  if (!isConfigured()) throw new Error("Project 2 not configured (SUPABASE2_* missing)");
  const res = await fetch(`${URL}/rest/v1/rpc/${fn}`, {
    method: "POST",
    headers: {
      apikey: KEY,
      Authorization: `Bearer ${KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ p_working: _slim(working) }),
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Project2 ${fn} failed (${res.status}): ${text.slice(0, 400)}`);
  return text ? JSON.parse(text) : null;
}

const preview = (working) => _rpc("hr_preview2", working);
const apply = (working) => _rpc("hr_apply2", working);

module.exports = { preview, apply, isConfigured };
