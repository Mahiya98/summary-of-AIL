// Phase data based on your screenshots
const data = {
  "Phase-01": {
    employees: 1815, roles: 80, requiredFTE: 301.9, overloaded: 1, underutilised: 62,
    sections: { "Production SMS": 29.6, "Production Rolling": 58.1, "Inventory": 44.3, "Quality": 36.1, "Scrap Management": 40, "Distribution": 45 }
  },
  "Phase-02": {
    employees: 1625, roles: 78, requiredFTE: 270.5, overloaded: 3, underutilised: 55,
    sections: { "Production SMS": 32, "Production Rolling": 60, "Inventory": 46, "Quality": 38, "Scrap Management": 42, "Distribution": 47 }
  },
  "Phase-03": {
    employees: 1617, roles: 76, requiredFTE: 265.0, overloaded: 5, underutilised: 50,
    sections: { "Production SMS": 35, "Production Rolling": 63, "Inventory": 48, "Quality": 40, "Scrap Management": 44, "Distribution": 49 }
  },
  "Phase-04": {
    employees: 1540, roles: 74, requiredFTE: 252.3, overloaded: 7, underutilised: 45,
    sections: { "Production SMS": 38, "Production Rolling": 66, "Inventory": 50, "Quality": 42, "Scrap Management": 46, "Distribution": 51 }
  }
};

function render(data) {
  const phases = Object.keys(data);

  // ===== Comparison Table =====
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

  // ===== Charts =====
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

const charts = {};
function drawBar(id, labels, values, label, color) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), {
    type: "bar",
    data: { labels, datasets: [{ label, data: values, backgroundColor: color }] },
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

render(data);
