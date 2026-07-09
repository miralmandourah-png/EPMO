// Deck Builder frontend: holds the whole deck in memory (`deck`), renders
// editable tabs from a small config, and talks to /api/* for upload/export.
// Nothing here persists to localStorage/cookies -- state lives only in
// this tab's JS memory and is discarded on reload.

function emptyDeck() {
  return {
    cover: { title: "", subtitle: "", date_label: "", footer: "" },
    toc: [],
    exec_summary: {
      headline: "", subheadline: "", committed_label: "", committed_bri_m: 0, savings_total_m: 0,
      ipi_tiers: [], benefit_status_by_lob: [], savings_breakdown: [],
    },
    context_gwp: { headline: "", subheadline: "", years: [], by_lob: [], by_lob_year_label: "" },
    ipi_sector: { headline: "", subheadline: "", read_note: "", items: [] },
    lob_overview: { headline: "", narrative: "", rows: [] },
    initiatives: [],
    projects: [],
    nps_summary: { headline: "", subheadline: "", company_actual: 0, company_target: 0, by_lob: [] },
    nps_detail: { headline: "", narrative: "", rows: [] },
    recovery_tracker: { headline: "", subheadline: "", items: [] },
    scenarios: { headline: "", subheadline: "", committed_m: 0, items: [] },
    recommended_actions: { headline: "", subheadline: "", items: [] },
    closing: { message: "Thank you", footer: "" },
  };
}

let deck = emptyDeck();
let activeTab = null;

const STATUS_OPTS = ["on_track", "cautious", "at_risk"];
const TYPE_OPTS = ["growth", "savings"];

const GROUPS = [
  { id: "cover", label: "Cover", scalarPath: ["cover"], fields: [
      { name: "title", label: "Title" }, { name: "subtitle", label: "Subtitle" },
      { name: "date_label", label: "Date label" }, { name: "footer", label: "Footer (every slide)" },
  ]},
  { id: "toc", label: "Contents", listPath: ["toc"], columns: [
      { name: "number", label: "#" }, { name: "section", label: "Section" },
      { name: "description", label: "Description" }, { name: "page", label: "Page" },
  ]},
  { id: "exec_summary", label: "Executive Summary", scalarPath: ["exec_summary"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
      { name: "committed_label", label: "Committed label" }, { name: "committed_bri_m", label: "Committed BRI (SAR m)", type: "number" },
      { name: "savings_total_m", label: "Savings total (SAR m)", type: "number" },
    ],
    subTables: [
      { title: "By execution tier", listPath: ["exec_summary", "ipi_tiers"], columns: [
          { name: "label", label: "Label" }, { name: "bri_m", label: "BRI (m)", type: "number" },
          { name: "color_key", label: "Color", type: "select", options: STATUS_OPTS },
      ]},
      { title: "Benefit status by line of business", listPath: ["exec_summary", "benefit_status_by_lob"], columns: [
          { name: "lob", label: "LoB" }, { name: "percent_on_track", label: "% On track", type: "number" },
          { name: "bri_m", label: "BRI (m)", type: "number" },
      ]},
      { title: "Savings breakdown", listPath: ["exec_summary", "savings_breakdown"], columns: [
          { name: "category", label: "Category" }, { name: "amount_m", label: "Amount (m)", type: "number" },
          { name: "detail", label: "Detail" },
      ]},
    ],
  },
  { id: "context_gwp", label: "Context (GWP)", scalarPath: ["context_gwp"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
      { name: "by_lob_year_label", label: "By-LoB panel label" },
    ],
    subTables: [
      { title: "GWP by year", listPath: ["context_gwp", "years"], columns: [
          { name: "year", label: "Year" }, { name: "bau_m", label: "BAU (m)", type: "number" },
          { name: "initiative_m", label: "Initiative (m)", type: "number" },
      ]},
      { title: "GWP by line of business", listPath: ["context_gwp", "by_lob"], columns: [
          { name: "lob", label: "LoB" }, { name: "gwp_m", label: "GWP (m)", type: "number" },
      ]},
    ],
  },
  { id: "ipi_sector", label: "Execution Signal (IPI)", scalarPath: ["ipi_sector"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
      { name: "read_note", label: "Read note (italic caption)", wide: true },
    ],
    listPath: ["ipi_sector", "items"], columns: [
      { name: "lob", label: "LoB" }, { name: "ipi", label: "IPI (0-5)", type: "number" },
    ],
  },
  { id: "lob_overview", label: "LoB Overview", scalarPath: ["lob_overview"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "narrative", label: "Narrative", wide: true, type: "textarea" },
    ],
    listPath: ["lob_overview", "rows"], columns: [
      { name: "lob", label: "LoB" }, { name: "status_summary", label: "Status summary" },
      { name: "bri_m", label: "BRI (m)", type: "number" }, { name: "ipi", label: "IPI", type: "number" },
    ],
  },
  { id: "initiatives", label: "Initiatives", listPath: ["initiatives"],
    hint: "One row per initiative. Rows are grouped by LoB into one deep-dive slide per line of business.",
    columns: [
      { name: "lob", label: "LoB" }, { name: "initiative", label: "Initiative" }, { name: "subtitle", label: "Subtitle" },
      { name: "ipi", label: "IPI", type: "number" }, { name: "ti", label: "TI", type: "number" },
      { name: "bri_committed_m", label: "BRI committed (m)", type: "number" }, { name: "bri_actual_m", label: "BRI actual (m)", type: "number" },
      { name: "status", label: "Status", type: "select", options: STATUS_OPTS }, { name: "comments", label: "Comments" },
    ],
  },
  { id: "projects", label: "Projects", listPath: ["projects"],
    hint: "One row per delivery project. Rows are grouped by LoB into a projects slide per line of business.",
    columns: [
      { name: "lob", label: "LoB" }, { name: "initiative", label: "Initiative" }, { name: "project", label: "Project" },
      { name: "ipi", label: "IPI", type: "number" }, { name: "ti", label: "TI", type: "number" }, { name: "comments", label: "Comments" },
    ],
  },
  { id: "nps_summary", label: "NPS Summary", scalarPath: ["nps_summary"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
      { name: "company_actual", label: "Companywide actual", type: "number" }, { name: "company_target", label: "Companywide target", type: "number" },
    ],
    listPath: ["nps_summary", "by_lob"], columns: [
      { name: "lob", label: "LoB" }, { name: "actual", label: "Actual", type: "number" }, { name: "target", label: "Target", type: "number" },
    ],
  },
  { id: "nps_detail", label: "NPS by Segment", scalarPath: ["nps_detail"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "narrative", label: "Narrative", wide: true, type: "textarea" },
    ],
    listPath: ["nps_detail", "rows"], columns: [
      { name: "lob", label: "LoB" }, { name: "segment", label: "Segment" },
      { name: "actual", label: "Actual", type: "number" }, { name: "target", label: "Target", type: "number" },
    ],
  },
  { id: "recovery_tracker", label: "Recovery Tracker", scalarPath: ["recovery_tracker"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
    ],
    listPath: ["recovery_tracker", "items"], columns: [
      { name: "rank", label: "#", type: "number" }, { name: "initiative", label: "Initiative" }, { name: "lob", label: "LoB" },
      { name: "item_type", label: "Type", type: "select", options: TYPE_OPTS }, { name: "bri_m", label: "BRI (m)", type: "number" },
      { name: "status", label: "Status", type: "select", options: STATUS_OPTS },
      { name: "ipi", label: "IPI", type: "number" }, { name: "ti", label: "TI", type: "number" },
      { name: "root_cause", label: "Root cause" }, { name: "corrective_action", label: "Corrective action" }, { name: "owner", label: "Owner" },
    ],
  },
  { id: "scenarios", label: "Scenarios", scalarPath: ["scenarios"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
      { name: "committed_m", label: "Committed (SAR m)", type: "number" },
    ],
    listPath: ["scenarios", "items"], columns: [
      { name: "name", label: "Name" }, { name: "description", label: "Description" },
      { name: "value_m", label: "Value (m)", type: "number" }, { name: "percent", label: "Percent", type: "number" },
    ],
  },
  { id: "recommended_actions", label: "Recommended Actions", scalarPath: ["recommended_actions"], fields: [
      { name: "headline", label: "Headline", wide: true }, { name: "subheadline", label: "Subheadline", wide: true },
    ],
    listPath: ["recommended_actions", "items"], columns: [
      { name: "title", label: "Title" }, { name: "description", label: "Description" },
    ],
  },
  { id: "closing", label: "Closing", scalarPath: ["closing"], fields: [
      { name: "message", label: "Message" }, { name: "footer", label: "Footer (blank = use cover footer)" },
    ],
  },
];

function getAt(obj, path) {
  return path.reduce((o, k) => (o == null ? o : o[k]), obj);
}

function emptyRow(columns) {
  const row = {};
  for (const c of columns) row[c.name] = c.type === "number" ? 0 : "";
  return row;
}

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") node.className = v;
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  }
  for (const child of [].concat(children)) {
    node.appendChild(typeof child === "string" ? document.createTextNode(child) : child);
  }
  return node;
}

function renderScalarFields(container, basePath, fields) {
  const grid = el("div", { class: "field-grid" });
  for (const f of fields) {
    const value = getAt(deck, basePath)[f.name];
    const wrap = el("div", { class: "field" + (f.wide ? " wide" : "") });
    wrap.appendChild(el("label", {}, f.label));
    let input;
    if (f.type === "textarea") {
      input = el("textarea", {});
      input.value = value ?? "";
    } else {
      input = el("input", { type: f.type === "number" ? "number" : "text" });
      input.value = value ?? "";
    }
    input.addEventListener("input", () => {
      const target = getAt(deck, basePath);
      target[f.name] = f.type === "number" ? (parseFloat(input.value) || 0) : input.value;
    });
    wrap.appendChild(input);
    grid.appendChild(wrap);
  }
  container.appendChild(grid);
}

function renderTable(container, title, listPath, columns, hint) {
  if (title) container.appendChild(el("h3", {}, title));
  if (hint) container.appendChild(el("div", { class: "empty-hint" }, hint));
  const wrap = el("div", { class: "table-wrap" });
  const table = el("table", { class: "grid" });
  const thead = el("tr", {}, columns.map((c) => el("th", {}, c.label)).concat([el("th", {}, "")]));
  table.appendChild(el("thead", {}, thead));
  const tbody = el("tbody", {});

  const list = getAt(deck, listPath);
  list.forEach((row, idx) => {
    const tr = el("tr", {});
    for (const c of columns) {
      const td = el("td", {});
      let input;
      if (c.type === "select") {
        input = el("select", {}, c.options.map((o) => el("option", { value: o }, o)));
        input.value = row[c.name] || c.options[0];
      } else {
        input = el("input", { type: c.type === "number" ? "number" : "text" });
        input.value = row[c.name] ?? "";
      }
      input.addEventListener("input", () => {
        row[c.name] = c.type === "number" ? (parseFloat(input.value) || 0) : input.value;
      });
      td.appendChild(input);
      tr.appendChild(td);
    }
    const actionTd = el("td", { class: "row-actions" });
    const delBtn = el("button", { title: "Remove row", onclick: () => { list.splice(idx, 1); renderPanel(activeTab); } }, "✕");
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrap.appendChild(table);
  container.appendChild(wrap);

  const toolbar = el("div", { class: "table-toolbar" });
  const addBtn = el("button", {
    class: "add-row-btn",
    onclick: () => { list.push(emptyRow(columns)); renderPanel(activeTab); },
  }, "+ Add row");
  toolbar.appendChild(addBtn);
  toolbar.appendChild(el("span", { class: "empty-hint" }, `${list.length} row${list.length === 1 ? "" : "s"}`));
  container.appendChild(toolbar);
}

function renderPanel(tabId) {
  activeTab = tabId;
  const group = GROUPS.find((g) => g.id === tabId);
  const panel = document.getElementById("panel");
  panel.innerHTML = "";
  panel.appendChild(el("h2", {}, group.label));

  if (group.fields) renderScalarFields(panel, group.scalarPath, group.fields);
  if (group.listPath) renderTable(panel, group.fields ? "" : "", group.listPath, group.columns, group.hint);
  if (group.subTables) {
    for (const st of group.subTables) {
      renderTable(panel, st.title, st.listPath, st.columns);
    }
  }

  document.querySelectorAll("#tabNav button").forEach((b) => b.classList.toggle("active", b.dataset.id === tabId));
}

function renderNav() {
  const nav = document.getElementById("tabNav");
  nav.innerHTML = "";
  for (const g of GROUPS) {
    const btn = el("button", { "data-id": g.id, onclick: () => renderPanel(g.id) }, g.label);
    nav.appendChild(btn);
  }
}

function toast(msg, isError = false) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.toggle("error", isError);
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.hidden = true; }, 3500);
}

async function downloadBlob(url, filenameFallback, init) {
  const res = await fetch(url, init);
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(text || `Request failed (${res.status})`);
  }
  const blob = await res.blob();
  const disposition = res.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/);
  const filename = match ? match[1] : filenameFallback;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

document.getElementById("btnTemplate").addEventListener("click", () => {
  downloadBlob("/api/template", "blank_template.xlsx").catch((e) => toast(e.message, true));
});

document.getElementById("btnSample").addEventListener("click", () => {
  downloadBlob("/api/template?sample=true", "sample_template.xlsx").catch((e) => toast(e.message, true));
});

document.getElementById("btnUpload").addEventListener("click", () => document.getElementById("fileInput").click());

document.getElementById("fileInput").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const form = new FormData();
  form.append("file", file);
  try {
    const res = await fetch("/api/upload", { method: "POST", body: form });
    if (!res.ok) throw new Error(await res.text());
    deck = await res.json();
    renderPanel(activeTab || GROUPS[0].id);
    toast("Workbook loaded.");
  } catch (err) {
    toast("Upload failed: " + err.message, true);
  }
  e.target.value = "";
});

document.getElementById("btnExportPptx").addEventListener("click", async () => {
  try {
    await downloadBlob("/api/export/pptx", "deck.pptx", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(deck),
    });
    toast("PPTX exported.");
  } catch (err) {
    toast("Export failed: " + err.message, true);
  }
});

document.getElementById("btnExportXlsx").addEventListener("click", async () => {
  try {
    await downloadBlob("/api/export/xlsx", "deck.xlsx", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(deck),
    });
    toast("Excel exported.");
  } catch (err) {
    toast("Export failed: " + err.message, true);
  }
});

renderNav();
renderPanel(GROUPS[0].id);
