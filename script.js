"use strict";

/* ================= configuration ================= */
const DATA_URL = "CFS_Yard_Data.xlsx";
const SHEET_NAME = "Daily Data";
const CATS = [
  { key: "expLoad",  label: "Export load",      color: "#2e75b6" },
  { key: "imp",      label: "Import",           color: "#17a398" },
  { key: "empty",    label: "Empty",            color: "#9aa5b1" },
  { key: "openYard", label: "Export open yard", color: "#f2a541" },
  { key: "special",  label: "Special",          color: "#d1495b" }
];
const NAVY = "#1f3864", LIGHT = "#c8d9ee", AMBER = "#f2a541", RED = "#d1495b", TEAL = "#17a398";

const SCOPES = {
  total: { label: "Total yard", slots: null, keys: null },
  wh:    { label: "Warehouse container stacking", slots: 221, keys: ["openYard"] },
  yard:  { label: "Yard container stacking", slots: 449, keys: ["expLoad", "empty", "imp", "special"] }
};

let allRows = [];
let charts = {};
let sourceLabel = "";
let activeCats = CATS;
let scopeLabel = "Total yard";

/* ================= helpers ================= */
const $ = (id) => document.getElementById(id);
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const fmtInt = (n) => Math.round(n).toLocaleString("en-IN");
const fmtPct = (n) => n.toFixed(1) + "%";
const fmtDate = (d) => `${String(d.getDate()).padStart(2, "0")} ${MON[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`;
const fmtShort = (d) => `${String(d.getDate()).padStart(2, "0")} ${MON[d.getMonth()]}`;
const ymKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const ymLabel = (ym) => { const [y, m] = ym.split("-"); return `${MON[+m - 1]} ${y.slice(2)}`; };
const normKey = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
const toNum = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};

function parseDateCell(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === "number" && v > 20000) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  if (typeof v === "string") {
    const s = v.trim();
    let m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
    if (m) {
      let [, d, mo, y] = m.map(Number);
      if (y < 100) y += 2000;
      return new Date(y, mo - 1, d);
    }
    m = s.match(/^(\d{1,2})[ -]([A-Za-z]{3,})[ -](\d{2,4})$/);
    if (m) {
      const mi = MON.findIndex((x) => m[2].toLowerCase().startsWith(x.toLowerCase()));
      let y = +m[3]; if (y < 100) y += 2000;
      if (mi >= 0) return new Date(y, mi, +m[1]);
    }
  }
  return null;
}

/* ================= data loading ================= */
function parseWorkbook(wb, label) {
  const ws = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) { showBanner("The file has no readable sheet. Use the standard CFS_Yard_Data.xlsx template.", true); return false; }
  const raw = XLSX.utils.sheet_to_json(ws, { defval: null });
  const rows = [];
  raw.forEach((r) => {
    const o = {};
    Object.keys(r).forEach((k) => { o[normKey(k)] = r[k]; });
    const date = parseDateCell(o.date);
    const slots = toNum(o.groundslots);
    const tiers = toNum(o.tiers) || 3;
    if (!date || slots <= 0) return;
    const expLoad = toNum(o.exportload20) + 2 * toNum(o.exportload40);
    const empty = toNum(o.empty20) + 2 * toNum(o.empty40);
    const openYard = toNum(o.openyard20) + 2 * toNum(o.openyard40);
    const imp = toNum(o.import20) + 2 * toNum(o.import40);
    const special = toNum(o.specialteu);
    const whArea = toNum(o.warehouseareasqft);
    const whUsed = toNum(o.warehouseusedsqft);
    const cap = slots * tiers;
    const occ = expLoad + empty + openYard + imp + special;
    rows.push({
      date, ym: ymKey(date), slots, cap,
      expLoad, empty, openYard, imp, special,
      occ, avail: cap - occ, util: (occ / cap) * 100,
      whPct: whArea > 0 ? (whUsed / whArea) * 100 : null
    });
  });
  if (!rows.length) {
    showBanner("No daily rows were found. Check that the sheet is named \u201CDaily Data\u201D and the column headings in row 1 are unchanged.", true);
    return false;
  }
  rows.sort((a, b) => a.date - b.date);
  const byDate = new Map();
  rows.forEach((r) => byDate.set(r.date.getTime(), r));       // duplicates: later row wins
  allRows = [...byDate.values()].sort((a, b) => a.date - b.date);
  sourceLabel = label;
  buildMonthOptions();
  resetFilterInputs();
  applyFilters();
  return true;
}

async function loadRepoData() {
  try {
    const res = await fetch(DATA_URL, { cache: "no-store" });
    if (!res.ok) throw new Error(String(res.status));
    const buf = await res.arrayBuffer();
    parseWorkbook(XLSX.read(buf, { type: "array", cellDates: true }), "repository data");
  } catch (e) {
    loadEmbedded();
  }
}

$("uploadInput").addEventListener("change", (ev) => {
  const f = ev.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const ok = parseWorkbook(XLSX.read(reader.result, { type: "array", cellDates: true }), `uploaded file \u201C${f.name}\u201D`);
      if (ok) showBanner(`Previewing ${f.name} on this device only. To publish it for everyone, upload the same file to the GitHub repository (Add file \u203A Upload files) so it replaces CFS_Yard_Data.xlsx.`, false);
    } catch (e) {
      showBanner("That file could not be read as an Excel workbook. Save it as .xlsx using the standard template and try again.", true);
    }
  };
  reader.readAsArrayBuffer(f);
  ev.target.value = "";
});

function showBanner(msg, isError) {
  const b = $("banner");
  b.textContent = msg;
  b.classList.toggle("error", !!isError);
  b.hidden = false;
}

/* ================= filters ================= */
function buildMonthOptions() {
  const sel = $("monthSelect");
  const months = [...new Set(allRows.map((r) => r.ym))].sort();
  sel.innerHTML = "<option value=\"all\">All months</option>" +
    months.map((m) => `<option value="${m}">${MON[+m.split("-")[1] - 1]} ${m.split("-")[0]}</option>`).join("");
}

function resetFilterInputs() {
  $("monthSelect").value = "all";
  $("fromDate").value = "";
  $("toDate").value = "";
}

function filteredRows() {
  const m = $("monthSelect").value;
  const from = $("fromDate").value ? new Date($("fromDate").value + "T00:00:00") : null;
  const to = $("toDate").value ? new Date($("toDate").value + "T00:00:00") : null;
  return allRows.filter((r) =>
    (m === "all" || r.ym === m) && (!from || r.date >= from) && (!to || r.date <= to));
}

["scopeSelect", "monthSelect", "fromDate", "toDate", "metricSelect"].forEach((id) =>
  $(id).addEventListener("change", applyFilters));
$("resetFilters").addEventListener("click", () => { resetFilterInputs(); applyFilters(); });
$("toggleDaily").addEventListener("click", () => {
  const w = $("dailyWrap");
  w.hidden = !w.hidden;
  $("toggleDaily").textContent = w.hidden ? "Show daily records" : "Hide daily records";
});

/* ================= aggregation ================= */
function monthlyAgg(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!map.has(r.ym)) map.set(r.ym, []);
    map.get(r.ym).push(r);
  });
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([ym, rs]) => {
    const avg = (f) => rs.reduce((s, x) => s + f(x), 0) / rs.length;
    const wh = rs.filter((x) => x.whPct !== null);
    const slotsMin = Math.min(...rs.map((x) => x.slots));
    const slotsMax = Math.max(...rs.map((x) => x.slots));
    return {
      ym, label: ymLabel(ym), days: rs.length,
      slots: slotsMin === slotsMax ? String(slotsMin) : `${slotsMax}\u2192${slotsMin}`,
      avgCap: avg((x) => x.cap), avgOcc: avg((x) => x.occ), avgAvail: avg((x) => x.avail),
      avgUtil: avg((x) => x.util), avgAvailPct: 100 - avg((x) => x.util),
      peak: Math.max(...rs.map((x) => x.util)), low: Math.min(...rs.map((x) => x.util)),
      avgWh: wh.length ? wh.reduce((s, x) => s + x.whPct, 0) / wh.length : null
    };
  });
}

/* ================= rendering ================= */
function applyScope(rows) {
  const scope = $("scopeSelect").value;
  const def = SCOPES[scope] || SCOPES.total;
  if (!def.keys) return { rows, excluded: 0, cats: CATS, label: def.label };
  const valid = rows.filter((r) => r.cap === 2010);
  const cats = CATS.filter((c) => def.keys.includes(c.key));
  const out = valid.map((r) => {
    const tiers = r.cap / r.slots;
    const cap = def.slots * tiers;
    const occ = def.keys.reduce((s, k) => s + r[k], 0);
    return { ...r, slots: def.slots, cap, occ, avail: cap - occ, util: (occ / cap) * 100 };
  });
  return { rows: out, excluded: rows.length - valid.length, cats, label: def.label };
}

function applyFilters() {
  const scoped = applyScope(filteredRows());
  const rows = scoped.rows;
  activeCats = scoped.cats;
  scopeLabel = scoped.label;
  const sn = $("scopeNote");
  if (scoped.excluded > 0) {
    sn.textContent = `${scoped.excluded} report(s) before 06 Jun 2026 are not shown in this band view \u2014 the 221/449-slot split applies to the 670-slot layout.`;
    sn.hidden = false;
  } else {
    sn.hidden = true;
  }
  const notice = $("notice");
  if (!rows.length) {
    notice.textContent = "No records match the selected filters. Widen the date range or choose another month.";
    notice.hidden = false;
    return;
  }
  notice.hidden = true;
  const first = rows[0], last = rows[rows.length - 1];
  $("sourceMeta").textContent =
    `${sourceLabel} \u00B7 ${rows.length} of ${allRows.length} reporting days \u00B7 ${fmtDate(first.date)} \u2013 ${fmtDate(last.date)}`;
  renderStrip(last, scoped.label);
  renderKpis(rows);
  renderTrend(rows);
  renderMonthly(rows);
  renderComposition(rows);
  renderWarehouse(rows);
  renderMonthlyTable(rows);
  renderDailyTable(rows);
}

function renderStrip(r, label) {
  $("stripDate").textContent = fmtDate(r.date);
  const prefix = label === "Total yard" ? "" : `${label} \u00B7 `;
  $("stripFigures").textContent =
    `${prefix}${fmtInt(r.occ)} / ${fmtInt(r.cap)} TEU occupied (${fmtPct(r.util)}) \u00B7 ${fmtInt(r.avail)} TEU available`;
  const strip = $("yardStrip");
  strip.innerHTML = "";
  activeCats.forEach((c) => {
    const seg = document.createElement("div");
    seg.className = "seg";
    seg.style.background = c.color;
    seg.style.width = (r[c.key] / r.cap * 100) + "%";
    seg.title = `${c.label}: ${fmtInt(r[c.key])} TEU`;
    strip.appendChild(seg);
  });
  $("stripLegend").innerHTML = activeCats.map((c) =>
    `<span><span class="key" style="background:${c.color}"></span>${c.label} ${fmtInt(r[c.key])}</span>`
  ).join("") + `<span><span class="key" style="background:repeating-linear-gradient(135deg,#e8edf3,#e8edf3 3px,#f3f6fa 3px,#f3f6fa 6px);border:1px solid var(--line)"></span>Available ${fmtInt(r.avail)}</span>`;
}

function renderKpis(rows) {
  const avg = (f) => rows.reduce((s, x) => s + f(x), 0) / rows.length;
  const peak = rows.reduce((a, b) => (b.util > a.util ? b : a));
  const low = rows.reduce((a, b) => (b.util < a.util ? b : a));
  const wh = rows.filter((x) => x.whPct !== null);
  const kpis = [
    { label: "Avg yard utilisation", value: fmtPct(avg((x) => x.util)), sub: `${rows.length} reporting days`, cls: "" },
    { label: "Peak utilisation", value: fmtPct(peak.util), sub: fmtDate(peak.date), cls: "accent-red" },
    { label: "Lowest utilisation", value: fmtPct(low.util), sub: fmtDate(low.date), cls: "accent-teal" },
    { label: "Avg occupancy", value: `${fmtInt(avg((x) => x.occ))} TEU`, sub: `latest ${fmtInt(rows[rows.length - 1].occ)} TEU`, cls: "" },
    { label: "Avg available", value: `${fmtInt(avg((x) => x.avail))} TEU`, sub: `latest ${fmtInt(rows[rows.length - 1].avail)} TEU`, cls: "" },
    { label: "Avg covered WH util.", value: wh.length ? fmtPct(wh.reduce((s, x) => s + x.whPct, 0) / wh.length) : "\u2013", sub: "of total warehouse area", cls: "" }
  ];
  $("kpiGrid").innerHTML = kpis.map((k) =>
    `<div class="kpi ${k.cls}"><div class="label">${k.label}</div><div class="value">${k.value}</div><div class="sub">${k.sub}</div></div>`
  ).join("");
}

function killChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

const baseOpts = () => ({
  responsive: true,
  maintainAspectRatio: false,
  interaction: { mode: "index", intersect: false },
  plugins: {
    legend: { labels: { font: { family: "Inter", size: 11 }, boxWidth: 14 } },
    tooltip: { titleFont: { family: "Inter" }, bodyFont: { family: "Inter" } }
  }
});

const capChangePlugin = {
  id: "capChange",
  afterDatasetsDraw(chart) {
    const rows = chart.$rows;
    if (!rows) return;
    const { ctx, chartArea: { top, bottom }, scales: { x } } = chart;
    rows.forEach((r, i) => {
      if (i === 0 || r.cap === rows[i - 1].cap) return;
      const px = x.getPixelForValue(i);
      ctx.save();
      ctx.strokeStyle = RED;
      ctx.setLineDash([5, 4]);
      ctx.beginPath(); ctx.moveTo(px, top); ctx.lineTo(px, bottom); ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = RED;
      ctx.font = "600 11px Inter";
      const rightSide = i > rows.length / 2;
      ctx.textAlign = rightSide ? "right" : "left";
      ctx.fillText(`capacity ${fmtInt(rows[i - 1].cap)} \u2192 ${fmtInt(r.cap)} TEU`, px + (rightSide ? -6 : 6), top + 12);
      ctx.restore();
    });
  }
};

function renderTrend(rows) {
  const metric = $("metricSelect").value;
  const cfg = {
    util:  { title: "Daily yard utilisation", get: (r) => r.util,  fmt: fmtPct,             unit: "%" },
    occ:   { title: "Daily occupancy (TEU)",  get: (r) => r.occ,   fmt: (v) => fmtInt(v),   unit: "TEU" },
    avail: { title: "Daily available (TEU)",  get: (r) => r.avail, fmt: (v) => fmtInt(v),   unit: "TEU" }
  }[metric];
  $("trendTitle").textContent = cfg.title + (scopeLabel === "Total yard" ? "" : " \u2014 " + scopeLabel.toLowerCase());
  const mon = monthlyAgg(rows);
  const monMap = new Map(mon.map((m) => [m.ym, { util: m.avgUtil, occ: m.avgOcc, avail: m.avgAvail }[metric]]));
  killChart("trend");
  const c = new Chart($("trendChart"), {
    type: "line",
    data: {
      labels: rows.map((r) => fmtShort(r.date)),
      datasets: [
        {
          label: "Daily value", data: rows.map(cfg.get),
          borderColor: NAVY, backgroundColor: "rgba(46,117,182,0.12)",
          fill: metric === "util", pointRadius: 2, pointHoverRadius: 4, borderWidth: 2, tension: 0.15
        },
        {
          label: "Monthly average", data: rows.map((r) => monMap.get(r.ym)),
          borderColor: AMBER, borderDash: [7, 5], stepped: "middle",
          pointRadius: 0, borderWidth: 2.5, fill: false
        }
      ]
    },
    options: {
      ...baseOpts(),
      scales: {
        x: { ticks: { maxRotation: 45, autoSkip: true, maxTicksLimit: 20, font: { size: 10 } }, grid: { display: false } },
        y: { beginAtZero: true, title: { display: true, text: cfg.unit },
             ticks: { callback: (v) => metric === "util" ? v + "%" : fmtInt(v) } }
      },
      plugins: {
        ...baseOpts().plugins,
        tooltip: { callbacks: { label: (i) => `${i.dataset.label}: ${cfg.fmt(i.parsed.y)}` } }
      }
    },
    plugins: [capChangePlugin]
  });
  c.$rows = rows;
  charts.trend = c;
}

function renderMonthly(rows) {
  const mon = monthlyAgg(rows);
  killChart("monthly");
  charts.monthly = new Chart($("monthlyChart"), {
    type: "bar",
    data: {
      labels: mon.map((m) => m.label),
      datasets: [
        { label: "Avg occupancy", data: mon.map((m) => m.avgOcc), backgroundColor: NAVY, stack: "s" },
        { label: "Avg available", data: mon.map((m) => m.avgAvail), backgroundColor: LIGHT, stack: "s" }
      ]
    },
    options: {
      ...baseOpts(),
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, beginAtZero: true, title: { display: true, text: "TEU" }, ticks: { callback: (v) => fmtInt(v) } }
      },
      plugins: {
        ...baseOpts().plugins,
        tooltip: {
          callbacks: {
            label: (i) => `${i.dataset.label}: ${fmtInt(i.parsed.y)} TEU`,
            afterBody: (items) => {
              const m = mon[items[0].dataIndex];
              return [`Avg utilisation: ${fmtPct(m.avgUtil)}`, `Avg capacity: ${fmtInt(m.avgCap)} TEU`];
            }
          }
        }
      }
    }
  });
}

function renderComposition(rows) {
  killChart("comp");
  charts.comp = new Chart($("compChart"), {
    type: "line",
    data: {
      labels: rows.map((r) => fmtShort(r.date)),
      datasets: activeCats.map((c) => ({
        label: c.label, data: rows.map((r) => r[c.key]),
        borderColor: c.color, backgroundColor: c.color + "CC",
        fill: true, pointRadius: 0, borderWidth: 1, tension: 0.15
      }))
    },
    options: {
      ...baseOpts(),
      scales: {
        x: { ticks: { maxRotation: 45, autoSkip: true, maxTicksLimit: 14, font: { size: 10 } }, grid: { display: false } },
        y: { stacked: true, beginAtZero: true, title: { display: true, text: "TEU" }, ticks: { callback: (v) => fmtInt(v) } }
      },
      plugins: {
        ...baseOpts().plugins,
        tooltip: { callbacks: { label: (i) => `${i.dataset.label}: ${fmtInt(i.parsed.y)} TEU` } }
      }
    }
  });
}

function renderWarehouse(rows) {
  const mon = monthlyAgg(rows);
  const monMap = new Map(mon.map((m) => [m.ym, m.avgWh]));
  killChart("wh");
  charts.wh = new Chart($("whChart"), {
    type: "line",
    data: {
      labels: rows.map((r) => fmtShort(r.date)),
      datasets: [
        { label: "Daily warehouse %", data: rows.map((r) => r.whPct), borderColor: TEAL,
          backgroundColor: "rgba(23,163,152,0.12)", fill: true, pointRadius: 2, borderWidth: 2, tension: 0.15 },
        { label: "Monthly average", data: rows.map((r) => monMap.get(r.ym)), borderColor: AMBER,
          borderDash: [7, 5], stepped: "middle", pointRadius: 0, borderWidth: 2.5, fill: false }
      ]
    },
    options: {
      ...baseOpts(),
      scales: {
        x: { ticks: { maxRotation: 45, autoSkip: true, maxTicksLimit: 14, font: { size: 10 } }, grid: { display: false } },
        y: { title: { display: true, text: "%" }, suggestedMin: 40, suggestedMax: 90, ticks: { callback: (v) => v + "%" } }
      },
      plugins: {
        ...baseOpts().plugins,
        tooltip: { callbacks: { label: (i) => `${i.dataset.label}: ${i.parsed.y === null ? "\u2013" : fmtPct(i.parsed.y)}` } }
      }
    }
  });
}

function renderMonthlyTable(rows) {
  const mon = monthlyAgg(rows);
  const avg = (f) => rows.reduce((s, x) => s + f(x), 0) / rows.length;
  const wh = rows.filter((x) => x.whPct !== null);
  const body = mon.map((m) => `<tr>
    <td>${m.label}</td><td>${m.days}</td><td>${m.slots}</td>
    <td>${fmtInt(m.avgCap)}</td><td>${fmtInt(m.avgOcc)}</td><td>${fmtInt(m.avgAvail)}</td>
    <td>${fmtPct(m.avgUtil)}</td><td>${fmtPct(m.avgAvailPct)}</td>
    <td>${fmtPct(m.peak)}</td><td>${fmtPct(m.low)}</td>
    <td>${m.avgWh === null ? "\u2013" : fmtPct(m.avgWh)}</td></tr>`).join("");
  const total = `<tr class="total-row">
    <td>Selected period</td><td>${rows.length}</td><td>\u2013</td>
    <td>${fmtInt(avg((x) => x.cap))}</td><td>${fmtInt(avg((x) => x.occ))}</td><td>${fmtInt(avg((x) => x.avail))}</td>
    <td>${fmtPct(avg((x) => x.util))}</td><td>${fmtPct(100 - avg((x) => x.util))}</td>
    <td>${fmtPct(Math.max(...rows.map((x) => x.util)))}</td><td>${fmtPct(Math.min(...rows.map((x) => x.util)))}</td>
    <td>${wh.length ? fmtPct(wh.reduce((s, x) => s + x.whPct, 0) / wh.length) : "\u2013"}</td></tr>`;
  $("monthlyTableBody").innerHTML = body + total;
}

function renderDailyTable(rows) {
  $("dailyTableBody").innerHTML = rows.map((r) => `<tr>
    <td>${fmtDate(r.date)}</td><td>${r.slots}</td><td>${fmtInt(r.cap)}</td>
    <td>${fmtInt(r.expLoad)}</td><td>${fmtInt(r.empty)}</td><td>${fmtInt(r.openYard)}</td>
    <td>${fmtInt(r.imp)}</td><td>${fmtInt(r.special)}</td>
    <td>${fmtInt(r.occ)}</td><td>${fmtInt(r.avail)}</td>
    <td>${fmtPct(r.util)}</td><td>${r.whPct === null ? "\u2013" : fmtPct(r.whPct)}</td></tr>`).join("");
}


function loadEmbedded() {
  const tag = document.getElementById("embeddedData");
  if (!tag) {
    $("sourceMeta").textContent = "No data loaded yet";
    showBanner("Could not load CFS_Yard_Data.xlsx. Use Upload Excel to view your data.", true);
    return;
  }
  const data = JSON.parse(tag.textContent);
  allRows = data.map((r) => {
    const [y, m, d] = r.date.split("-").map(Number);
    return { ...r, date: new Date(y, m - 1, d) };
  });
  sourceLabel = "built-in data";
  buildMonthOptions();
  resetFilterInputs();
  applyFilters();
  showBanner("Showing the built-in data (up to 28 Jul 2026). The live GitHub Pages link always loads the latest CFS_Yard_Data.xlsx by itself \u2014 or use Upload Excel here to view a newer file.", false);
}

/* ================= start ================= */
document.addEventListener("DOMContentLoaded", loadRepoData);
