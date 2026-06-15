// =============================================================================
// 0. CONFIG — pick ONE of these URLs.
//    Option A (recommended): Published CSV  → no auth, fastest, works in browser
//    Option B: Direct XLSX export → requires sheet shared "Anyone with link"
// =============================================================================
const SHEET_ID = "1fx3FFlAPbF-_nbHEjUtEDrWW8LIwyygfYBZTwSVEVJw";
const GID      = "58409945";

// ✅ Your published CSV link (works perfectly in browser, no CORS issues)
const SHEET_URL = `https://docs.google.com/spreadsheets/d/e/2PACX-1vRe7_Vji3K-cSqRkLgK3SKI20HrHzdbrowl2PVcletw5iwZ03NdSF00C6cRqh2tr7EN72BvFWEg3rqi/pub?gid=${GID}&single=true&output=csv`;

// (Alt) XLSX export — uncomment if you prefer xlsx and the sheet is public:
// const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx&gid=${GID}`;

// =============================================================================
// 1. STATE
// =============================================================================
let rawRows = [];

const staticMetrics = {
  "Phase-01": { requiredFTE: 301.9, overloaded: 1, underutilised: 62,
    sections: { "Production SMS": 29.6, "Production Rolling": 58.1, "Inventory": 44.3, "Quality": 36.1, "Scrap Management": 40, "Distribution": 45 } },
  "Phase-02": { requiredFTE: 270.5, overloaded: 3, underutilised: 55,
    sections: { "Production SMS": 32, "Production Rolling": 60, "Inventory": 46, "Quality": 38, "Scrap Management": 42, "Distribution": 47 } },
  "Phase-03": { requiredFTE: 265.0, overloaded: 5, underutilised: 50,
    sections: { "Production SMS": 35, "Production Rolling": 63, "Inventory": 48, "Quality": 40, "Scrap Management": 44, "Distribution": 49 } },
  "Phase-04": { requiredFTE: 252.3, overloaded: 7, underutilised: 45,
    sections: { "Production SMS": 38, "Production Rolling": 66, "Inventory": 50, "Quality": 42, "Scrap Management": 46, "Distribution": 51 } },
};

let data = {};

// =============================================================================
// 2. HELPERS — robust column lookup
// =============================================================================
function pick(row, ...candidates) {
  for (const key of candidates) {
    if (row[key] !== undefined && String(row[key]).trim() !== "") return row[key];
  }
  const norm = s => String(s).toLowerCase().replace(/\s+/g, "");
  for (const k of Object.keys(row)) {
    if (candidates.some(c => norm(c) === norm(k))) {
      const v = row[k];
      if (v !== undefined && String(v).trim() !== "") return v;
    }
  }
  return "";
}

const getEnroll  = r => String(pick(r, "Employee Enroll", "Employee Enrol", "Emp Enroll", "EmployeeEnroll")).trim();
const getPhase   = r => String(pick(r, "Phase Remarks", "Phase", "PhaseRemarks")).trim();
const getRole    = r => String(pick(r, "Role")).trim();
const getSection = r => String(pick(r, "Section")).trim();

function uniqueEnrollCount(rows, predicate = () => true) {
  const set = new Set();
  for (const r of rows) {
    if (!predicate(r)) continue;
    const id = getEnroll(r);
    if (id) set.add(id);
  }
  return set.size;
}

const countEmployeesInPhase   = (rows, phase)          => uniqueEnrollCount(rows, r => getPhase(r) === phase);
const countEmployeesInSection = (rows, phase, section) => uniqueEnrollCount(rows, r => getPhase(r) === phase && getSection(r) === section);

function countUniqueRoles(rows, phase) {
  const set = new Set();
  for (const r of rows) {
    if (getPhase(r) !== phase) continue;
    const role = getRole(r);
    if (role) set.add(role);
  }
  return set.size;
}

// =============================================================================
// 3. BUILD `data`
// =============================================================================
function buildData() {
  data = {};
  for (const phase of Object.keys(staticMetrics)) {
    data[phase] = {
      employees:     countEmployeesInPhase(rawRows, phase),
      roles:         countUniqueRoles(rawRows, phase),
      requiredFTE:   staticMetrics[phase].requiredFTE,
      overloaded:    staticMetrics[phase].overloaded,
      underutilised: staticMetrics[phase].underutilised,
      sections:      staticMetrics[phase].sections,
    };
  }
}

// =============================================================================
// 4. LOAD SHEET — auto-detects CSV vs XLSX  ✅ THE FIX
// =============================================================================
async function loadSheet(url) {
  const res = await fetch(url, { redirect: "follow" });
  if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);

  const isCsv = /output=csv|\.csv($|\?)/i.test(url);

  let workbook;
  if (isCsv) {
    // 🟢 CSV path — read as text, parse as string
    const text = await res.text();
    if (!text || text.trim().startsWith("<")) {
      throw new Error("Got HTML instead of CSV — check the publish URL / sharing settings.");
    }
    workbook = XLSX.read(text, { type: "string" });
  } else {
    // 🟢 XLSX path — read as binary
    const buf = await res.arrayBuffer();
    workbook = XLSX.read(buf, { type: "array" });
  }

  const ws = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
}

// =============================================================================
// 5. RENDER
// =============================================================================
function render(data) {
  const phases = Object.keys(data);

  let html = `<thead><tr><th>Metric</th>${phases.map(p => `<th>${p}</th>`).join("")}</tr></thead><tbody>`;
  const rows = [
    ["Total Employees",     p => data[p].employees.toLocaleString()],
    ["Unique Roles",        p => data[p].roles],
    ["Required FTE",        p => data[p].requiredFTE],
    ["Overloaded Roles",    p => data[p].overloaded],
    ["Underutilised Roles", p => data[p].underutilised],
  ];
  rows.forEach(([metric, fn]) => {
    html += `<tr><td><strong>${metric}</strong></td>${phases.map(p => `<td>${fn(p)}</td>`).join("")}</tr>`;
  });
  html += "</tbody>";
  document.getElementById("compareTable").innerHTML = html;

  drawBar("empChart", phases, phases.map(p => data[p].employees), "Employees", "#3b82f6");
  drawBar("fteChart", phases, phases.map(p => data[p].requiredFTE), "Required FTE", "#10b981");

  drawGrouped("loadChart", phases, [
    { label: "Overloaded",    data: phases.map(p => data[p].overloaded),    color: "#ef4444" },
    { label: "Underutilised", data: phases.map(p => data[p].underutilised), color: "#3b82f6" }
  ]);

  const sections = Object.keys(data[phases[0]].sections);
  const colors = ["#2563eb", "#10b981", "#f59e0b", "#ef4444"];
  const datasets = phases.map((p, i) => ({
    label: p,
    data: sections.map(s => data[p].sections[s] || 0),
    color: colors[i % colors.length]
  }));
  drawGrouped("sectionChart", sections, datasets);
}

// =============================================================================
// 6. CHART HELPERS
// =============================================================================
const charts = {};
function drawBar(id, labels, values, label, color) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: { labels, datasets: [{ label, data: values, backgroundColor: color }] },
    options: { responsive: true, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });
}
function drawGrouped(id, labels, datasets) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: { labels, datasets: datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color })) },
    options: { responsive: true, plugins: { legend: { position: "top" } }, scales: { y: { beginAtZero: true } } }
  });
}

// =============================================================================
// 7. BOOT
// =============================================================================
(async () => {
  try {
    rawRows = await loadSheet(SHEET_URL);
    console.log("✅ Loaded rows:", rawRows.length);
    console.log("📋 Detected headers:", Object.keys(rawRows[0] || {}));
    console.log("🔎 Sample row:", rawRows[0]);

    // Show what phases & enroll values were detected (debugging)
    const phaseSeen  = new Set(rawRows.map(getPhase).filter(Boolean));
    const enrollSeen = new Set(rawRows.map(getEnroll).filter(Boolean));
    console.log("🏷️  Phases found:", [...phaseSeen]);
    console.log("👤 Unique Employee Enroll (all phases):", enrollSeen.size);

    buildData();
    console.table(
      Object.fromEntries(
        Object.entries(data).map(([k, v]) => [k, { employees: v.employees, roles: v.roles }])
      )
    );

    render(data);
  } catch (err) {
    console.error("❌ Could not load sheet:", err);
    alert("Could not load sheet.\n" + err.message);
  }
})();
