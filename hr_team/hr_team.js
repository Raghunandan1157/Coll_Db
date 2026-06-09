/* /hr_team — staff sync UI. Talks to /api/staff-sync/{preview,apply}. */
(function () {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const STATUS = { 1: "Active", 2: "On leave", 3: "Resigned", 4: "Terminated", 5: "Retired", 6: "Absconded" };

  const state = { passcode: "", file: null, token: null };

  function show(el) { el.hidden = false; }
  function hide(el) { el.hidden = true; }
  function err(el, msg) { el.textContent = msg; el.hidden = !msg; }
  function overlay(on, msg) { $("overlayMsg").textContent = msg || "Working…"; $("overlay").hidden = !on; }

  /* ---------- Gate ---------- */
  $("gateForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const pass = $("passcode").value.trim();
    if (!pass) return;
    state.passcode = pass;
    err($("gateErr"), "");
    hide($("gate"));
    show($("uploadStage"));
  });

  /* ---------- Upload ---------- */
  const drop = $("drop");
  const fileInput = $("fileInput");

  function setFile(f) {
    if (!f) return;
    if (!/\.xlsx?$/i.test(f.name)) { err($("uploadErr"), "Please choose an .xlsx / .xls file."); return; }
    state.file = f;
    $("fileName").textContent = f.name;
    $("previewBtn").disabled = false;
    err($("uploadErr"), "");
  }

  fileInput.addEventListener("change", (e) => setFile(e.target.files[0]));
  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("drag"); })
  );
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("drag"); })
  );
  drop.addEventListener("drop", (e) => { if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]); });

  $("previewBtn").addEventListener("click", runPreview);

  async function runPreview() {
    err($("uploadErr"), "");
    overlay(true, "Parsing file & computing diff…");
    const fd = new FormData();
    fd.append("file", state.file);
    fd.append("passcode", state.passcode);
    try {
      const res = await fetch("/api/staff-sync/preview", { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok) {
        if (res.status === 401) { backToGate("Wrong passcode."); return; }
        throw new Error(data.error || "Preview failed");
      }
      state.token = data.token;
      renderPreview(data);
      hide($("uploadStage"));
      show($("previewStage"));
    } catch (e2) {
      err($("uploadErr"), e2.message);
    } finally {
      overlay(false);
    }
  }

  /* ---------- Preview render ---------- */
  function stat(dot, label, num) {
    return `<li><span class="lbl"><span class="dot ${dot}"></span>${label}</span><span class="num">${num}</span></li>`;
  }

  function renderPreview(d) {
    $("previewMeta").textContent =
      `Sheet "${d.sheet || "?"}" · ${d.total_rows} rows · ${d.working_rows} working`;

    const e = d.ec2 || {};
    $("ec2Stats").innerHTML =
      stat("add", "Add", e.add ?? 0) +
      stat("upd", "Update", e.update ?? 0) +
      stat("del", "Delete", e.delete ?? 0) +
      stat("", "Final rows", e.final ?? 0);

    const s = d.supabase;
    if (!d.supabase_configured || !s) {
      $("sbStats").innerHTML = `<li><span class="lbl">Not configured</span><span class="num">—</span></li>`;
    } else {
      $("sbStats").innerHTML =
        stat("add", "Insert (new)", s.insert_count ?? 0) +
        stat("upd", "Reactivate", s.reactivate_count ?? 0) +
        stat("del", "Deactivate", s.deactivate_count ?? 0);
    }

    // New master-data
    const newRoles = (s && s.new_roles) || [];
    const newDesig = (s && s.new_designations) || [];
    const newBranch = (s && s.new_branches) || [];
    const all = [
      ...newRoles.map((x) => "Role: " + x),
      ...newDesig.map((x) => "Designation: " + x),
      ...newBranch.map((x) => "Branch: " + x),
    ];
    if (all.length) {
      $("masterChips").innerHTML = all.map((x) => `<span class="chip">${esc(x)}</span>`).join("");
      show($("masterWarn"));
    } else {
      hide($("masterWarn"));
    }

    // Delete / deactivate detail
    const delIds = (e.delete_ids || []).map((id) => `<span class="code-pill dl">${esc(id)}</span>`);
    const deact = (s && s.deactivate) || [];
    const deactPills = deact.map(
      (x) => `<span class="code-pill dl">${esc(x.code)} → ${STATUS[x.to_status] || x.to_status}</span>`
    );
    const body = [];
    if (delIds.length) body.push(`<div><b>EC2 delete (${delIds.length})</b><div class="reveal-body">${delIds.join("")}</div></div>`);
    if (deactPills.length) body.push(`<div><b>Supabase deactivate (${deactPills.length})</b><div class="reveal-body">${deactPills.join("")}</div></div>`);
    if (body.length) {
      $("deleteBody").innerHTML = body.join("");
      show($("deleteDetails"));
    } else {
      hide($("deleteDetails"));
    }
  }

  /* ---------- Confirm / apply ---------- */
  $("confirmBtn").addEventListener("click", runApply);
  $("cancelBtn").addEventListener("click", () => location.reload());
  $("againBtn").addEventListener("click", () => location.reload());

  async function runApply() {
    err($("previewErr"), "");
    overlay(true, "Backing up & syncing both databases…");
    try {
      const res = await fetch("/api/staff-sync/apply", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token: state.token, passcode: state.passcode }),
      });
      const data = await res.json();
      if (!res.ok) {
        if (res.status === 409) { err($("previewErr"), data.error + " "); return; }
        if (res.status === 401) { backToGate("Wrong passcode."); return; }
        renderResult(data, false);
        hide($("previewStage")); show($("resultStage"));
        return;
      }
      renderResult(data, true);
      hide($("previewStage"));
      show($("resultStage"));
    } catch (e2) {
      err($("previewErr"), e2.message);
    } finally {
      overlay(false);
    }
  }

  function renderResult(d, ok) {
    const lines = [];
    const banner = ok
      ? `<div class="result-banner ok">Both databases synced successfully.</div>`
      : `<div class="result-banner bad">${esc(d.error || "Sync failed")}</div>`;

    const ec2 = (d.result && d.result.ec2) || d.ec2;
    const sb = (d.result && d.result.supabase) || d.supabase;
    const bak = (d.result && d.result.backup) || d.backup;

    if (ec2) lines.push(line("EC2 — final / working", `${ec2.total} / ${ec2.working}`));
    if (ec2) lines.push(line("EC2 — deleted", ec2.deleted));
    if (sb) {
      lines.push(line("Supabase — inserted", sb.inserted));
      lines.push(line("Supabase — reactivated", sb.reactivated));
      lines.push(line("Supabase — deactivated", sb.deactivated));
      const nm = (sb.new_roles || 0) + (sb.new_designations || 0) + (sb.new_branches || 0);
      lines.push(line("Supabase — new master-data", nm));
    }
    if (bak) lines.push(line("EC2 backup table", bak.table));

    $("resultBody").innerHTML = banner + lines.join("");
  }

  function line(k, v) { return `<div class="result-line"><span>${esc(k)}</span><span class="v">${esc(String(v))}</span></div>`; }
  function backToGate(msg) {
    state.passcode = "";
    ["uploadStage", "previewStage", "resultStage"].forEach((i) => hide($(i)));
    show($("gate"));
    err($("gateErr"), msg);
  }
  function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
})();
