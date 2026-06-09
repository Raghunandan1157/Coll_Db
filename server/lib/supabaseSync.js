/**
 * supabaseSync.js — Drive the Supabase Grow_With_Me sync via PostgREST RPCs.
 * Requires SUPABASE_URL + SUPABASE_SERVICE_KEY in the environment (server-side).
 */

const SUPABASE_URL = process.env.SUPABASE_URL || "";
const SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || "";
const SCHEMA = "Grow_With_Me";

function _assertConfigured() {
  if (!SUPABASE_URL || !SERVICE_KEY) {
    throw new Error("Supabase not configured (SUPABASE_URL / SUPABASE_SERVICE_KEY missing)");
  }
}

async function _rpc(fn, args) {
  _assertConfigured();
  const res = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fn}`, {
    method: "POST",
    headers: {
      apikey: SERVICE_KEY,
      Authorization: `Bearer ${SERVICE_KEY}`,
      "Content-Type": "application/json",
      "Content-Profile": SCHEMA,
      "Accept-Profile": SCHEMA,
    },
    body: JSON.stringify(args),
  });
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`Supabase ${fn} failed (${res.status}): ${text.slice(0, 500)}`);
  }
  return text ? JSON.parse(text) : null;
}

/** Read-only change plan for the given roster rows. */
function preview(allRows) {
  return _rpc("hr_preview", { p_all: allRows });
}

/** Transactional apply. Returns the summary the RPC produced. */
function apply(allRows) {
  return _rpc("hr_apply", { p_all: allRows });
}

function isConfigured() {
  return Boolean(SUPABASE_URL && SERVICE_KEY);
}

module.exports = { preview, apply, isConfigured };
