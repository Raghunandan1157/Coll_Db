/**
 * syncTokens.js — Short-lived, single-use tokens that hold a parsed roster
 * between /preview and /apply so a bad file can't be applied blind and a
 * stale plan can't be replayed.
 */

const crypto = require("crypto");

const TTL_MS = 15 * 60 * 1000;
const store = new Map(); // token -> { rows, working, createdAt }

function _sweep() {
  const now = Date.now();
  for (const [t, v] of store) {
    if (now - v.createdAt > TTL_MS) store.delete(t);
  }
}

/** Stash a parsed roster, return a token. */
function issue(payload) {
  _sweep();
  const token = crypto.randomUUID();
  store.set(token, { ...payload, createdAt: Date.now() });
  return token;
}

/** Consume a token once. Returns the payload or null if missing/expired. */
function consume(token) {
  _sweep();
  const v = store.get(token);
  if (!v) return null;
  store.delete(token);
  if (Date.now() - v.createdAt > TTL_MS) return null;
  return v;
}

module.exports = { issue, consume };
