/**
 * hrTeam.js — /hr_team staff-sync API.
 *
 *   POST /api/staff-sync/preview  (multipart: file, passcode)  -> change plan + token
 *   POST /api/staff-sync/apply    (json: token, passcode)      -> applies both DBs
 *
 * EC2 employee_master = hard-delete sync. Supabase Grow_With_Me = status-flip +
 * insert via transactional RPCs. Guarded by HR_SYNC_PASSCODE.
 */

const crypto = require("crypto");
const rateLimit = require("express-rate-limit");

const { parseStaffWorkbook } = require("../lib/staffParser");
const ec2Sync = require("../lib/ec2Sync");
const supabaseSync = require("../lib/supabaseSync");
const backup = require("../lib/staffBackup");
const tokens = require("../lib/syncTokens");

function _passcodeOk(provided) {
  const expected = process.env.HR_SYNC_PASSCODE || "";
  if (!expected) return false; // not configured -> deny
  const a = Buffer.from(String(provided || ""));
  const b = Buffer.from(expected);
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function registerHrTeamRoutes(app, pool, upload) {
  const applyLimiter = rateLimit({
    windowMs: 60 * 1000,
    max: 5,
    message: { error: "Too many sync attempts. Wait a minute." },
  });

  // ---- PREVIEW (read-only) ----
  app.post("/api/staff-sync/preview", upload.single("file"), async (req, res) => {
    if (!_passcodeOk(req.body && req.body.passcode)) {
      return res.status(401).json({ error: "Invalid HR passcode" });
    }
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    let parsed;
    try {
      parsed = parseStaffWorkbook(req.file.buffer);
    } catch (err) {
      return res.status(400).json({ error: "Could not parse file: " + err.message });
    }
    if (!parsed.working.length) {
      return res.status(400).json({ error: "No rows with status 'Working' found in the file" });
    }

    try {
      const ec2 = await ec2Sync.computeDiff(pool, parsed.working);
      let supabase = null;
      if (supabaseSync.isConfigured()) {
        supabase = await supabaseSync.preview(parsed.rows);
      }
      const token = tokens.issue({ rows: parsed.rows, working: parsed.working });
      return res.json({
        token,
        sheet: parsed.sheet,
        total_rows: parsed.total,
        working_rows: parsed.working.length,
        ec2,
        supabase,
        supabase_configured: supabaseSync.isConfigured(),
      });
    } catch (err) {
      console.error("[hr-sync] preview error:", err);
      return res.status(500).json({ error: err.message });
    }
  });

  // ---- APPLY (writes) ----
  app.post("/api/staff-sync/apply", applyLimiter, async (req, res) => {
    if (!_passcodeOk(req.body && req.body.passcode)) {
      return res.status(401).json({ error: "Invalid HR passcode" });
    }
    const token = req.body && req.body.token;
    if (!token) return res.status(400).json({ error: "Missing token" });

    const payload = tokens.consume(token);
    if (!payload) {
      return res.status(409).json({ error: "Preview expired or already used. Re-upload and preview again." });
    }

    const result = { ec2: null, supabase: null, backup: null };
    try {
      // 1. EC2 backup (cheap insurance) then sync.
      result.backup = await backup.backupEmployeeMaster(pool);
      result.ec2 = await ec2Sync.apply(pool, payload.working);
    } catch (err) {
      console.error("[hr-sync] EC2 apply error:", err);
      return res.status(500).json({ error: "EC2 sync failed: " + err.message, stage: "ec2", result });
    }

    // 2. Supabase (its own transaction + internal backup).
    if (supabaseSync.isConfigured()) {
      try {
        result.supabase = await supabaseSync.apply(payload.rows);
      } catch (err) {
        console.error("[hr-sync] Supabase apply error:", err);
        return res.status(500).json({
          error: "EC2 synced OK, but Supabase failed: " + err.message,
          stage: "supabase",
          result,
        });
      }
    }

    return res.json({ ok: true, ...result });
  });

  console.log("[hr-sync] /hr_team routes registered" + (supabaseSync.isConfigured() ? " (Supabase on)" : " (Supabase OFF — set SUPABASE_* env)"));
}

module.exports = registerHrTeamRoutes;
