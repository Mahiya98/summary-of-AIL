// Demo data based on your screenshots — replace by uploading CSV
const demoData = {
  "Phase-01": { employees: 1815, roles: 80, requiredFTE: 301.9, overloaded: 1, underutilised: 62,
    sections: { "Production SMS": 29.6, "Production Rolling": 58.1, "Inventory": 44.3, "Quality": 36.1, "Scrap Management": 40, "Distribution": 45 } },
  "Phase-02": { employees: 1625, roles: 78, requiredFTE: 270.5, overloaded: 3, underutilised: 55,
    sections: { "Production SMS": 32, "Production Rolling": 60, "Inventory": 46, "Quality": 38, "Scrap Management": 42, "Distribution": 47 } },
  "Phase-03": { employees: 1617, roles: 76, requiredFTE: 265.0, overloaded: 5, underutilised: 50,
    sections: { "Production SMS": 35, "Production Rolling": 63, "Inventory": 48, "Quality": 40, "Scrap Management": 44, "Distribution": 49 } },
  "Phase-04": { employees: 1540, roles: 74, requiredFTE: 252.3, overloaded: 7, underutilised: 45,
    sections: { "Production SMS": 38, "Production Rolling": 66, "Inventory": 50, "Quality": 42, "Scrap Management": 46, "Distribution": 51 } }
};

let currentData = demoData;

function render(data) {
  const phases = Object.keys(data);

  // KPIs (totals across phases)
  const totalEmp = phases.reduce((s, p) => s + data[p].employees, 0);
  const totalFTE = phases.reduce((s, p) => s + data[p].requiredFTE, 0).toFixed(1);
  const totalOver = phases.reduce((s, p) => s + data[p].overloaded, 0);
  const totalUnder = phases.reduce((s, p) => s + data[p].underutilised, 0);

  document.getElementById("kpiGrid").innerHTML = `
    <div class="kpi"><h3>Total Employees</h3><div class="val">${totalEmp.toLocaleString()}</div><div class="label">All phases</div></div>
    <div class="kpi"><h3>Total Required FTE</h3><div class="val">${totalFTE}</div><div class="label">Sum</div></div>
    <div class="kpi"><h3>Overloaded</h3><div class="val" style="color:#dc2626">${totalOver}</div><div class="label">Roles > 100%</div></div>
    <div class="kpi"><h3>Underutilised</h3><div class="val" style="color:#2563eb">${totalUnder}</div><div class="label">Roles < 60%</div></div>
    <div class="kpi"><h3>Phases</h3><div class="val">${phases.length}</div><div class="label">Compared</div></div>
  `;

  drawBar("empChart", phases, phases.map(p => data[p].employees), "Employees", "#3b82f6");
  drawBar("fteChart", phases, phases.map(p => data[p].requiredFTE), "Required FTE", "#10b981");

  // Overloaded vs Underutilised
  drawGrouped("loadChart", phases,
    [
      { label: "Overloaded", data: phases.map(p => data[p].overloaded), color: "#ef4444" },
      { label: "Underutilised", data: phases.map(p => data[p].underutilised), color: "#3b82f6" }
    ]);

  // Section workload comparison
  const sections = Object.keys(data[phases[0]].sections);
  const datasets = phases.map((p, i) => ({
    label: p,
    data: sections.map(s => data[p].sections[s] || 0),
    color: ["#2563eb", "#10b981", "#f59e0b", "#ef4444"][i]
  }));
  drawGrouped("sectionChart", sections, datasets);

  // Table
  let html = `<thead><tr><th>Metric</th>${phases.map(p => `<th>${p}</th>`).join("")}</tr></thead><tbody>`;
  const rows = [
    ["Total Employees", p => data[p].employees.toLocaleString()],
    ["Unique Roles", p => data[p].roles],
    ["Required FTE", p => data[p].requiredFTE],
    ["Overloaded Roles", p => data[p].overloaded],
    ["Underutilised Roles", p => data[p].underutilised],
  ];
  rows.forEach(([metric, fn]) => {
    html += `<tr><td><strong>${metric}</strong></td>${phases.map(p => `<td>${fn(p)}</td>`).join("")}</tr>`;
  });
  html += "</tbody>";
  document.getElementById("compareTable").innerHTML = html;
}

let charts = {};
function drawBar(id, labels, data, label, color) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: { labels, datasets: [{ label, data, backgroundColor: color }] },
    options: { responsive: true, plugins: { legend: { display: false } } }
  });
}
function drawGrouped(id, labels, datasets) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: {
      labels,
      datasets: datasets.map(d => ({ label: d.label, data: d.data, backgroundColor: d.color }))
    },
    options: { responsive: true }
  });
}

// CSV Upload — expects columns: Phase, Section, Role, HC, FTE, Workload, Status
document.getElementById("csvFile").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  Papa.parse(file, {
    header: true, dynamicTyping: true, skipEmptyLines: true,
    complete: (res) => {
      const grouped = {};
      res.data.forEach(r => {
        const phase = r.Phase || r.phase || r["Phase Remarks"];
        if (!phase) return;
        if (!grouped[phase]) grouped[phase] = { employees: 0, roles: new Set(), requiredFTE: 0, overloaded: 0, underutilised: 0, sections: {} };
        grouped[phase].employees += Number(r.HC || r.Headcount || 0);
        grouped[phase].roles.add(r.Role);
        grouped[phase].requiredFTE += Number(r.FTE || 0);
        const wl = Number(String(r.Workload || r["Load %"] || 0).replace("%",""));
        if (wl > 100) grouped[phase].overloaded++;
        if (wl < 60) grouped[phase].underutilised++;
        const sec = r.Section;
        if (sec) {
          if (!grouped[phase].sections[sec]) grouped[phase].sections[sec] = { sum: 0, n: 0 };
          grouped[phase].sections[sec].sum += wl;
          grouped[phase].sections[sec].n++;
        }
      });
      // finalize
      Object.keys(grouped).forEach(p => {
        grouped[p].roles = grouped[p].roles.size;
        grouped[p].requiredFTE = +grouped[p].requiredFTE.toFixed(1);
        const secAvg = {};
        Object.entries(grouped[p].sections).forEach(([k, v]) => secAvg[k] = +(v.sum / v.n).toFixed(1));
        grouped[p].sections = secAvg;
      });
      currentData = grouped;
      document.getElementById("status").textContent = `✅ Loaded ${res.data.length} rows from ${file.name}`;
      render(currentData);
    }
  });
});

render(currentData);
