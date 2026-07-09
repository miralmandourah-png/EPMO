// Strategic Health Check Platform — frontend.
//
// Loads the schema + saved state from the local backend, renders 14 editable
// tabs, autosaves every change to the local JSON store, and drives the Excel
// import/export flows. All state lives on this machine; nothing is sent to any
// third party.

let SCHEMA = { sections: [] };
let state = { fields: {}, rows: {}, config: {} };
let activeTab = null;
let saveTimer = null;
let templateConfigured = false;

function updateExportButton() {
  const btn = document.getElementById("btnExport");
  btn.title = templateConfigured
    ? "Export the branded deck from your template"
    : "Upload a .pptx template first (Data ▾ → Upload PPTX template)";
  btn.style.opacity = templateConfigured ? "1" : "0.6";
}

const STATUS_COLOR = {
  "On-track": "green", "Overachieved": "green",
  "Cautious": "amber", "Watch": "amber",
  "Critical": "red", "At-risk": "red",
  "Not scored": "grey",
};

// ---- helpers -------------------------------------------------------------
function el(tag, attrs = {}, kids = []) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") n.className = v;
    else if (k === "html") n.innerHTML = v;
    else if (k.startsWith("on") && typeof v === "function") n.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined) n.setAttribute(k, v);
  }
  for (const kid of [].concat(kids)) if (kid != null) n.appendChild(typeof kid === "string" ? document.createTextNode(kid) : kid);
  return n;
}
const getF = (key) => (state.fields[key] ?? "");
function setF(key, val) {
  if (val === "" || val === null || val === undefined) delete state.fields[key];
  else state.fields[key] = val;
  scheduleSave();
}
const rowCount = (tid) => Math.max(0, parseInt(state.rows[tid] ?? 0, 10));

function toast(msg, kind = "") {
  const t = document.getElementById("toast");
  t.textContent = msg; t.className = "toast " + kind; t.hidden = false;
  clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), 4200);
}

// ---- autosave ------------------------------------------------------------
function scheduleSave() {
  setSaveState("saving");
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flushSave, 600);
}
async function flushSave() {
  try {
    const res = await fetch("/api/state", {
      method: "PUT", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ fields: state.fields, rows: state.rows }),
    });
    if (!res.ok) throw new Error(await res.text());
    setSaveState("saved");
  } catch (e) { setSaveState("error"); toast("Autosave failed: " + e.message, "error"); }
}
function setSaveState(s) {
  const n = document.getElementById("saveState");
  if (s === "saving") { n.textContent = "Saving…"; n.className = "save-state saving"; }
  else if (s === "error") { n.textContent = "Save failed"; n.className = "save-state"; }
  else { n.textContent = "All changes saved"; n.className = "save-state"; }
}

// ---- computed columns ----------------------------------------------------
function computeCell(expr, rowVals) {
  const m = expr.match(/^(\w+)\s*([+\-])\s*(\w+)$/);
  if (!m) return "";
  const a = parseFloat(rowVals[m[1]]) || 0, b = parseFloat(rowVals[m[3]]) || 0;
  const r = m[2] === "+" ? a + b : a - b;
  return Number.isInteger(r) ? String(r) : r.toFixed(1);
}

// ---- rendering: fields ---------------------------------------------------
function renderFieldsBlock(block) {
  const grid = el("div", { class: "field-grid" });
  for (const f of block.fields) {
    const wide = f.type === "textarea";
    const wrap = el("div", { class: "field" + (wide ? " wide" : "") });
    wrap.appendChild(el("label", {}, f.label));
    let input;
    if (f.type === "textarea") {
      input = el("textarea", {}); input.value = getF(f.key);
    } else if (f.type === "select") {
      input = el("select", {}, (f.options || []).map((o) => el("option", { value: o }, o)));
      input.value = getF(f.key) || (f.options ? f.options[0] : "");
    } else {
      input = el("input", { type: f.type === "number" || f.type === "pct" ? "number" : "text",
        step: "any", placeholder: f.type === "pct" ? "%" : "" });
      input.value = getF(f.key);
    }
    input.addEventListener("input", () => setF(f.key, input.value));
    wrap.appendChild(input);
    grid.appendChild(wrap);
  }
  return el("div", { class: "block" }, [el("h3", {}, block.title), grid]);
}

// ---- rendering: tables ---------------------------------------------------
function readRow(tid, i, columns) {
  const o = {};
  for (const c of columns) o[c.name] = getF(`${tid}.${i}.${c.name}`);
  return o;
}
function shiftRowsUp(tid, columns, removeIdx, count) {
  // compact: move every row after removeIdx down by one, drop the last
  for (let i = removeIdx; i < count - 1; i++) {
    for (const c of columns) {
      const nextVal = getF(`${tid}.${i + 1}.${c.name}`);
      if (nextVal === "") delete state.fields[`${tid}.${i}.${c.name}`];
      else state.fields[`${tid}.${i}.${c.name}`] = nextVal;
    }
  }
  for (const c of columns) delete state.fields[`${tid}.${count - 1}.${c.name}`];
  state.rows[tid] = count - 1;
}

function renderTableBlock(block) {
  const tid = block.id;
  const columns = block.columns;
  const wrap = el("div", { class: "table-wrap" });
  const table = el("table", { class: "grid" });
  const headRow = el("tr", {}, columns.map((c) => el("th", {}, c.label)).concat([el("th", {}, "")]));
  table.appendChild(el("thead", {}, headRow));
  const tbody = el("tbody", {});
  const n = rowCount(tid);

  for (let i = 0; i < n; i++) {
    const tr = el("tr", {});
    const cellInputs = {};
    for (const c of columns) {
      const td = el("td", {});
      const key = `${tid}.${i}.${c.name}`;
      if (c.computed) {
        td.className = "computed";
        const span = el("span", {}, computeCell(c.computed, readRow(tid, i, columns)));
        td._compute = () => (span.textContent = computeCell(c.computed, readRow(tid, i, columns)));
        td.appendChild(span);
        td.dataset.compute = "1";
        tr._computedCells = (tr._computedCells || []).concat(td);
      } else if (c.type === "select") {
        const sel = el("select", {}, (c.options || []).map((o) => el("option", { value: o }, o)));
        sel.value = getF(key) || (c.options ? c.options[0] : "");
        const badge = el("span", { class: "badge " + (STATUS_COLOR[sel.value] || "grey") }, sel.value);
        const cell = el("div", { class: "status-cell" }, [sel, badge]);
        sel.addEventListener("change", () => {
          setF(key, sel.value);
          badge.className = "badge " + (STATUS_COLOR[sel.value] || "grey");
          badge.textContent = sel.value;
        });
        td.appendChild(cell);
      } else {
        const input = el("input", { type: c.type === "number" || c.type === "pct" ? "number" : "text", step: "any" });
        input.value = getF(key);
        input.addEventListener("input", () => {
          setF(key, input.value);
          if (tr._computedCells) tr._computedCells.forEach((cc) => cc._compute && cc._compute());
          if (c.name === "ipi" && tr._ipiBar) tr._ipiBar(parseFloat(input.value));
        });
        cellInputs[c.name] = input;
        // IPI mini-bar
        if (c.name === "ipi") {
          const track = el("div", { class: "bar-track" });
          const fill = el("div", { class: "bar-fill" });
          track.appendChild(fill);
          const paint = (v) => {
            v = isNaN(v) ? 0 : v;
            fill.style.width = Math.max(0, Math.min(100, (v / 5) * 100)) + "%";
            fill.style.background = v >= 3.0 ? "var(--green)" : v >= 2.6 ? "var(--amber)" : "var(--red)";
          };
          paint(parseFloat(input.value));
          tr._ipiBar = paint;
          td.appendChild(el("div", { class: "bar-wrap" }, [input, track]));
        } else {
          td.appendChild(input);
        }
      }
      tr.appendChild(td);
    }
    const delTd = el("td", {});
    delTd.appendChild(el("button", { class: "row-del", title: "Delete row",
      onclick: () => { shiftRowsUp(tid, columns, i, rowCount(tid)); scheduleSave(); renderTab(activeTab); } }, "✕"));
    tr.appendChild(delTd);
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  wrap.appendChild(table);

  const foot = el("div", { class: "table-foot" }, [
    el("button", { class: "add-row", onclick: () => { state.rows[tid] = rowCount(tid) + 1; scheduleSave(); renderTab(activeTab); } }, "+ Add row"),
    el("span", { class: "count-hint" }, `${n} row${n === 1 ? "" : "s"}`),
  ]);
  return el("div", { class: "block" }, [el("h3", {}, block.title), wrap, foot]);
}

// ---- rendering: tab + tabbar --------------------------------------------
function renderTab(secId) {
  activeTab = secId;
  const sec = SCHEMA.sections.find((s) => s.id === secId);
  const content = document.getElementById("content");
  content.innerHTML = "";
  content.appendChild(el("h1", {}, sec.label));
  for (const block of sec.blocks) {
    content.appendChild(block.type === "fields" ? renderFieldsBlock(block) : renderTableBlock(block));
  }
  document.querySelectorAll("#tabbar button").forEach((b) => b.classList.toggle("active", b.dataset.id === secId));
  window.scrollTo(0, 0);
}
function renderTabbar() {
  const bar = document.getElementById("tabbar");
  bar.innerHTML = "";
  SCHEMA.sections.forEach((s, i) => {
    const btn = el("button", { "data-id": s.id, onclick: () => renderTab(s.id) },
      [el("span", { class: "tab-num" }, String(i + 1).padStart(2, "0")), s.label]);
    bar.appendChild(btn);
  });
}

// ---- modal ---------------------------------------------------------------
function showModal(node) {
  const host = document.getElementById("modal");
  host.innerHTML = ""; host.appendChild(node);
  document.getElementById("modalOverlay").hidden = false;
}
function closeModal() { document.getElementById("modalOverlay").hidden = true; }
document.getElementById("modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") closeModal();
});

// ---- data menu actions ---------------------------------------------------
function toggleMenu(force) {
  const m = document.getElementById("dataMenu");
  m.hidden = force !== undefined ? !force : !m.hidden;
}
document.getElementById("btnDataMenu").addEventListener("click", (e) => { e.stopPropagation(); toggleMenu(); });
document.addEventListener("click", () => toggleMenu(false));
document.getElementById("dataMenu").addEventListener("click", (e) => e.stopPropagation());

document.getElementById("dataMenu").addEventListener("click", (e) => {
  const act = e.target.dataset.act;
  if (!act) return;
  toggleMenu(false);
  if (act === "template") downloadTemplate();
  else if (act === "importTemplate") document.getElementById("fileTemplate").click();
  else if (act === "importIpi") document.getElementById("fileIpi").click();
  else if (act === "importMilestones") askGraceThenImportMilestones();
  else if (act === "uploadPptx") document.getElementById("filePptx").click();
  else if (act === "reset") confirmReset();
});

document.getElementById("filePptx").addEventListener("change", async (e) => {
  const file = e.target.files[0]; e.target.value = "";
  if (!file) return;
  const form = new FormData(); form.append("file", file);
  try {
    const res = await fetch("/api/template/pptx", { method: "POST", body: form });
    if (!res.ok) throw new Error(await res.text());
    const info = await res.json();
    templateConfigured = true; updateExportButton();
    toast(`Template "${info.name}" set (${info.slide_count} slides). Export is ready.`, "ok");
  } catch (err) { toast("Template upload failed: " + tryJson(err.message), "error"); }
});

async function downloadTemplate() {
  try {
    const res = await fetch("/api/template");
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    const a = el("a", { href: URL.createObjectURL(blob), download: "strategic_health_check_template.xlsx" });
    document.body.appendChild(a); a.click(); a.remove();
    toast("Template downloaded — fill the Value column and re-import.", "ok");
  } catch (e) { toast("Download failed: " + e.message, "error"); }
}

document.getElementById("fileTemplate").addEventListener("change", async (e) => {
  const file = e.target.files[0]; e.target.value = "";
  if (!file) return;
  const form = new FormData(); form.append("file", file);
  try {
    const res = await fetch("/api/import/template", { method: "POST", body: form });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state = data.state; renderTab(activeTab);
    toast(`Imported ${data.applied} values from template.`, "ok");
  } catch (err) { toast("Template import failed: " + err.message, "error"); }
});

document.getElementById("fileIpi").addEventListener("change", async (e) => {
  const file = e.target.files[0]; e.target.value = "";
  if (!file) return;
  const form = new FormData(); form.append("file", file);
  try {
    const res = await fetch("/api/import/ipi", { method: "POST", body: form });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state = data.state; renderTab(activeTab);
    showIpiReference(data);
  } catch (err) { toast("IPI import failed: " + err.message, "error"); }
});

function showIpiReference(data) {
  const r = data.reference;
  const sb = r.status_breakdown || {};
  const refRows = [
    ["Enterprise IPI (written to Exec Summary)", r.enterprise_ipi],
    ["Strategic projects = SBP + SEP (written)", r.strategic],
    ["Total projects", r.total_projects],
    ["CP (change/BAU)", r.cp],
    ["SBP + SEP split", `${r.sbp} + ${r.sep}`],
    ...Object.entries(sb).map(([k, v]) => [k, v]),
  ];
  const table = el("table", { class: "ref-table" },
    refRows.map(([k, v]) => el("tr", {}, [el("td", {}, k), el("td", {}, String(v))])));
  showModal(el("div", {}, [
    el("h3", {}, "IPI Accountability imported"),
    el("p", {}, "Enterprise IPI and the strategic project count were written to the Executive Summary. The figures below are a read-only sanity check and were not written anywhere."),
    table,
    el("div", { class: "modal-actions" }, [el("button", { class: "btn primary", onclick: closeModal }, "Done")]),
  ]));
}

function askGraceThenImportMilestones() {
  const graceInput = el("input", { type: "number", min: "0", step: "1", value: String(state.config?.delayed_grace_days ?? 0) });
  showModal(el("div", {}, [
    el("h3", {}, "Import Strategic Milestones"),
    el("p", { html: "Classification: <b>complete</b> = progress ≥ 100%; <b>delayed</b> = past due date and not complete; <b>not-yet-due</b> = everything else." }),
    el("div", { class: "grace-row" }, [
      el("label", { html: "Grace period (days past due before counting as delayed):" }), graceInput,
    ]),
    el("div", { class: "modal-actions" }, [
      el("button", { class: "btn ghost", onclick: closeModal }, "Cancel"),
      el("button", { class: "btn primary", onclick: () => {
        closeModal();
        document.getElementById("fileMilestones").dataset.grace = String(parseInt(graceInput.value, 10) || 0);
        document.getElementById("fileMilestones").click();
      } }, "Choose file…"),
    ]),
  ]));
}

document.getElementById("fileMilestones").addEventListener("change", async (e) => {
  const file = e.target.files[0]; const grace = e.target.dataset.grace || "0"; e.target.value = "";
  if (!file) return;
  const form = new FormData(); form.append("file", file);
  try {
    const res = await fetch(`/api/import/milestones?grace_days=${encodeURIComponent(grace)}`, { method: "POST", body: form });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state = data.state; renderTab(activeTab);
    const r = data.reference;
    const table = el("table", { class: "ref-table" }, [
      ["Total milestones", r.total], ["Complete", r.complete], ["Delayed", r.delayed],
      ["Not-yet-due", r.not_yet_due], ["Grace period (days)", r.grace_days],
    ].map(([k, v]) => el("tr", {}, [el("td", {}, k), el("td", {}, String(v))])));
    showModal(el("div", {}, [
      el("h3", {}, "Strategic Milestones imported"),
      el("p", {}, "Milestone counts were written to the Executive Summary."),
      table,
      el("div", { class: "note" }, data.note),
      el("div", { class: "modal-actions" }, [el("button", { class: "btn primary", onclick: closeModal }, "Done")]),
    ]));
  } catch (err) { toast("Milestones import failed: " + err.message, "error"); }
});

function confirmReset() {
  showModal(el("div", {}, [
    el("h3", {}, "Reset to blank?"),
    el("p", {}, "This permanently clears all dashboard data on this machine and restores the default section labels. This cannot be undone."),
    el("div", { class: "modal-actions" }, [
      el("button", { class: "btn ghost", onclick: closeModal }, "Cancel"),
      el("button", { class: "btn primary", onclick: async () => {
        try {
          const res = await fetch("/api/reset", { method: "POST" });
          state = await res.json(); renderTab(activeTab); closeModal();
          toast("Dashboard reset to blank.", "ok");
        } catch (e) { toast("Reset failed: " + e.message, "error"); }
      } }, "Reset everything"),
    ]),
  ]));
}

document.getElementById("btnExport").addEventListener("click", async () => {
  try {
    const res = await fetch("/api/export/pptx", { method: "POST" });
    if (res.ok) {
      const blob = await res.blob();
      const a = el("a", { href: URL.createObjectURL(blob), download: "strategic_health_check.pptx" });
      document.body.appendChild(a); a.click(); a.remove();
      toast("PPTX exported.", "ok");
    } else {
      const msg = await res.text();
      toast(tryJson(msg), "");
    }
  } catch (e) { toast("Export failed: " + e.message, "error"); }
});
function tryJson(s) { try { return JSON.parse(s).detail || s; } catch { return s; } }

// ---- boot ----------------------------------------------------------------
async function boot() {
  const [schemaRes, stateRes] = await Promise.all([fetch("/api/schema"), fetch("/api/state")]);
  SCHEMA = await schemaRes.json();
  state = await stateRes.json();
  state.fields = state.fields || {}; state.rows = state.rows || {}; state.config = state.config || {};
  renderTabbar();
  renderTab(SCHEMA.sections[0].id);
  try {
    const info = await (await fetch("/api/template/pptx/info")).json();
    templateConfigured = !!info.configured;
  } catch (e) { /* ignore */ }
  updateExportButton();
}
boot();
