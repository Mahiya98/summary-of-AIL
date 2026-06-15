// Requires <script src="https://docs.google.com/spreadsheets/d/1fx3FFlAPbF-_nbHEjUtEDrWW8LIwyygfYBZTwSVEVJw/edit?gid=58409945#gid=58409945"></script>
async function loadSheet(url) {
  const buf  = await (await fetch(url)).arrayBuffer();
  const wb   = XLSX.read(buf, { type: "array" });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

loadSheet("your-file.xlsx").then(rows => {
  rawRows.length = 0;
  rawRows.push(...rows);

  // Recompute employees & roles for every phase
  for (const phase of Object.keys(staticMetrics)) {
    data[phase].employees = countEmployeesInPhase(rawRows, phase);
    data[phase].roles     = countUniqueRoles(rawRows, phase);
  }
  render(data);
});
// =============================================================================
// 1. RAW ROWS — replace this with your parsed sheet rows (e.g. via SheetJS/PapaParse)
//    Keys must match your column headers exactly.
// =============================================================================
const rawRows = [
  // Example — load your real data here
  // { "SBU": "AIL", "Section": "Production Rolling", "Shift": "A",
  //   "Employee Name": "Md. Jahangir Hossain", "Role": "Foreman",
  //   "Employee Enroll": "515194", "Work Centre": "Rolling Production Floor",
  //   "Task List": "Roll Change", "Time required/task": 140, "Frequency": 0.7,
  //   "Actual Time/Shift": 100, "Remarks": "", "Phase Remarks": "Phase-01" },
];

// =============================================================================
// 2. HELPERS — count unique Employee Enroll per phase / per section / overall
// =============================================================================
function uniqueEnrollCount(rows, predicate = () => true) {
  const set = new Set();
  for (const r of rows) {
    if (!predicate(r)) continue;
    const id = r["Employee Enroll"];
    if (id !== undefined && id !== null && String(id).trim() !== "") {
      set.add(String(id).trim());
    }
  }
  return set.size;
}

// Total unique employees in a phase
function countEmployeesInPhase(rows, phase) {
  return uniqueEnrollCount(rows, r => r["Phase Remarks"] === phase);
}

// Unique employees per section within a phase (useful for section breakdowns)
function countEmployeesInSection(rows, phase, section) {
  return uniqueEnrollCount(
    rows,
    r => r["Phase Remarks"] === phase && r["Section"] === section
  );
}

// Unique roles in a phase (column E)
function countUniqueRoles(rows, phase) {
  const set = new Set();
  rows
    .filter(r => r["Phase Remarks"] === phase)
    .forEach(r => r["Role"] && set.add(String(r["Role"]).trim()));
  return set.size;
}

// =============================================================================
// 3. PHASE DATA — employees & roles now computed dynamically from rawRows
//    (FTE, overloaded, underutilised kept as-is; swap to your formulas if needed)
// =============================================================================
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

const data = {};
for (const phase of Object.keys(staticMetrics)) {
  data[phase] = {
    employees:    countEmployeesInPhase(rawRows, phase),   // ✅ unique Employee Enroll (col F)
    roles:        countUniqueRoles(rawRows, phase),        // ✅ unique Role (col E)
    requiredFTE:  staticMetrics[phase].requiredFTE,
    overloaded:   staticMetrics[phase].overloaded,
    underutilised: staticMetrics[phase].underutilised,
    sections:     staticMetrics[phase].sections,
  };
}

// 🔁 FALLBACK: if rawRows is empty (e.g. before data load), keep original numbers
//    so the UI still renders. Remove this block once rawRows is wired up.
if (rawRows.length === 0) {
  const fallback = {
    "Phase-01": { employees: 1815, roles: 80 },
    "Phase-02": { employees: 1625, roles: 78 },
    "Phase-03": { employees: 1617, roles: 76 },
    "Phase-04": { employees: 1540, roles: 74 },
  };
  for (const p in fallback) Object.assign(data[p], fallback[p]);
}

// =============================================================================
// 4. RENDER (unchanged — already reads data[p].employees dynamically)
// =============================================================================
function render(data) {
  const phases = Object.keys(data);

  // ----- Comparison Table -----
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

  // ----- Charts -----
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
// 5. CHART HELPERS (unchanged)
// =============================================================================
const charts = {};

function drawBar(id, labels, values, label, color) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: { labels, datasets: [{ label, data: values, backgroundColor: color }] },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });
}

function drawGrouped(id, labels, datasets) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: {
      labels,
      datasets: datasets.map(d => ({
        label: d.label,
        data: d.data,
        backgroundColor: d.color
      }))
    },
    options: {
      responsive: true,
      plugins: { legend: { position: "top" } },
      scales: { y: { beginAtZero: true } }
    }
  });
}

// =============================================================================
// 6. INITIAL RENDER
// =============================================================================
render(data);
