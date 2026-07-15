import React, { useState, useMemo, useEffect } from "react";
import Papa from "papaparse";
import { AreaChart, Area, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, Cell } from "recharts";
import * as XLSX from "xlsx";
import { ChevronRight, ChevronDown, TrendingUp, TrendingDown, Target, Percent, LayoutDashboard, LineChart as LineChartIcon, ArrowLeft, Wifi, WifiOff, Loader2, Download, Calendar, CheckCircle2, AlertTriangle, Gauge, Receipt, Package, Scale, Banknote } from "lucide-react";

// ---------- Google Sheet CSV sources ----------
const TARGET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRbKm9XW2L3mvPuUwTXCcLLt5nN3MFO0IciJ3ta5waPjerG0A459RtjwcDBinBgJeJxZpQsZBz9w8kZ/pub?output=csv";
const SKU_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSV9JuunrwMmHZ6rC7i2LROX_rvHI4pFFW9AnafZJsktehmLmfudPgNajTA03csnkTGIthOob8lgKGQ/pub?output=csv";

// ---------- Mock data (structure mirrors ABL: Section -> Unit -> Area) ----------
const BRANDS = ["AJI-Retail", "AJI-Bulk", "Hapima", "TasteMate"];
// SKU_wise data sheet's own Brand column uses the real product names (not the Retail/Bulk channel split
// used in the Target vs Progress sheet), and its SKU column holds just the pack size.
const SKU_BRANDS = ["AJI-NO-MOTO", "HAPIMA", "TasteMate"];

const SKU_LIST = {
  "AJI-NO-MOTO": ["450g", "200g", "100g", "20TK", "10TK"],
  "HAPIMA": ["15g"],
  "TasteMate": ["450g"],
};

function seedRand(seed) {
  let s = seed;
  return () => {
    s = (s * 9301 + 49297) % 233280;
    return s / 233280;
  };
}

function buildAreas(prefix, count, seedBase) {
  const rnd = seedRand(seedBase);
  return Array.from({ length: count }, (_, i) => {
    const perBrand = {};
    BRANDS.forEach((b) => {
      const target = Math.round(3000 + rnd() * 5000); // KG
      const achv = Math.round(target * (0.75 + rnd() * 0.5));
      const lastMonth = Math.round(achv * (0.85 + rnd() * 0.3));
      perBrand[b] = { target, achv, lastMonth };
    });
    return { name: `${prefix}-${i + 1}`, brands: perBrand };
  });
}

const SECTIONS = [
  {
    name: "South",
    units: [
      { name: "South Unit 1", areas: buildAreas("S1", 3, 11) },
      { name: "South Unit 2", areas: buildAreas("S2", 2, 23) },
    ],
  },
  {
    name: "North",
    units: [
      { name: "North Unit 1", areas: buildAreas("N1", 3, 37) },
      { name: "North Unit 2", areas: buildAreas("N2", 2, 41) },
    ],
  },
  {
    name: "CTG",
    units: [
      { name: "CTG Unit 1", areas: buildAreas("C1", 2, 53) },
      { name: "CTG Unit 2", areas: buildAreas("C2", 2, 61) },
    ],
  },
];

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

// Working-day-wise cumulative comparison: Running month vs Last month vs Last year (same month)
// Replace this generator with real rows pulled from your Google Sheet (day -> cumulative achv per SKU)
function workingDayTrend(seed, workingDays = 24) {
  const rnd = seedRand(seed);
  const build = (base, growthBias) => {
    let cum = 0;
    return Array.from({ length: workingDays }, (_, i) => {
      const daily = base * (0.85 + rnd() * 0.3) * growthBias;
      cum += daily;
      return Math.round(cum);
    });
  };
  const base = 90 + rnd() * 40; // KG per working day
  const runningMonth = build(base, 1.08); // slight growth bias for demo
  const lastMonth = build(base, 1.0);
  const lastYear = build(base * 0.85, 1.0);
  return Array.from({ length: workingDays }, (_, i) => ({
    day: `D${i + 1}`,
    "Running Month": runningMonth[i],
    "Last Month": lastMonth[i],
    "Last Year": lastYear[i],
  }));
}

// ---------- Live Google Sheets data layer ----------
function normalizeKey(k) {
  return String(k).toLowerCase().replace(/[^a-z0-9]/g, "");
}
function normalizeRow(row) {
  const out = {};
  Object.keys(row).forEach((k) => (out[normalizeKey(k)] = row[k]));
  return out;
}
function getField(nrow, candidates, fallback = "") {
  for (const c of candidates) {
    if (nrow[c] !== undefined && nrow[c] !== null && String(nrow[c]).trim() !== "") return nrow[c];
  }
  return fallback;
}

async function fetchCsv(url) {
  const bustUrl = url + (url.includes("?") ? "&" : "?") + "_t=" + Date.now();
  const res = await fetch(bustUrl, { cache: "no-store" });
  if (!res.ok) throw new Error("CSV fetch failed: " + res.status);
  const text = await res.text();
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}

async function fetchCsvRaw(url) {
  const bustUrl = url + (url.includes("?") ? "&" : "?") + "_t=" + Date.now();
  const res = await fetch(bustUrl, { cache: "no-store" });
  if (!res.ok) throw new Error("CSV fetch failed: " + res.status);
  const text = await res.text();
  const parsed = Papa.parse(text, { skipEmptyLines: true }); // array-of-arrays, no header merge
  return parsed.data;
}

function parseMonthKey(label) {
  const cleaned = String(label).replace(/[-_]/g, " ").trim();
  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? null : d;
}

// Target_vs_Progress sheet (wide format): row1 = brand group headers (merged), row2 = RF/Result/Prog.,
// data rows = Section, Unit, Area, then RF/Result/Prog. triplets per brand. Section/Unit cells are merged
// vertically in the sheet, so blanks are forward-filled from the row above. An optional "Month" column
// (e.g. "July-2026") can sit anywhere in the header row — commonly added at the far right — and turns
// each block of rows into its own selectable month.
function buildLiveSectionsWide(rawRows) {
  if (!rawRows || rawRows.length < 3) return null;
  const row1 = rawRows[0];
  const row2 = rawRows[1];

  let monthColIndex = row2.findIndex((h) => normalizeKey(h || "") === "month");
  if (monthColIndex === -1) monthColIndex = row1.findIndex((h) => normalizeKey(h || "") === "month");
  const hasMonthCol = monthColIndex !== -1;

  const findCol = (candidates) => {
    let idx = row2.findIndex((h) => candidates.includes(normalizeKey(h || "")));
    if (idx === -1) idx = row1.findIndex((h) => candidates.includes(normalizeKey(h || "")));
    return idx;
  };
  const totalWDColIndex = findCol(["totalworkingdays", "totalworkingday", "workingdaystotal"]);
  const passedWDColIndex = findCol(["workingdayspassed", "workingdaypassed", "workingdayspass", "dayspassed", "workingdaydone"]);
  const extraCols = [monthColIndex, totalWDColIndex, passedWDColIndex].filter((i) => i !== -1);

  const baseCols = 3; // Section, Unit, Area always occupy the first 3 columns

  const filledRow1 = [];
  let lastBrand = "";
  for (let i = 0; i < row1.length; i++) {
    const v = (row1[i] || "").toString().trim();
    if (i < baseCols || extraCols.includes(i)) {
      filledRow1.push("");
      continue;
    }
    if (v) lastBrand = v;
    filledRow1.push(lastBrand);
  }
  const colMeta = filledRow1.map((brandLabel, i) => {
    if (i < baseCols || extraCols.includes(i)) return null;
    const metric = (row2[i] || "").toString().trim().toLowerCase();
    return { brand: brandLabel, metric };
  });

  const byMonth = {}; // month -> sectionMap
  const paceByMonth = {}; // month -> { total, passed }
  let curMonth = "";
  let curSection = "";
  let curUnit = "";
  let curTotalWD = 0;
  let curPassedWD = 0;
  for (let r = 2; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.every((c) => !c || !String(c).trim())) continue;
    const monthRaw = hasMonthCol ? (row[monthColIndex] || "").toString().trim() : "";
    const secRaw = (row[0] || "").toString().trim();
    const unitRaw = (row[1] || "").toString().trim();
    const areaRaw = (row[2] || "").toString().trim();
    const totalWDRaw = totalWDColIndex !== -1 ? (row[totalWDColIndex] || "").toString().trim() : "";
    const passedWDRaw = passedWDColIndex !== -1 ? (row[passedWDColIndex] || "").toString().trim() : "";
    if (monthRaw) curMonth = monthRaw;
    if (secRaw) curSection = secRaw;
    if (unitRaw) curUnit = unitRaw;
    if (totalWDRaw) curTotalWD = parseInt(totalWDRaw, 10) || 0;
    if (passedWDRaw) curPassedWD = parseInt(passedWDRaw, 10) || 0;
    if (!areaRaw) continue;
    const month = hasMonthCol ? curMonth || "Unspecified" : "current";
    const section = curSection || "Unassigned";
    const unit = curUnit || "Unassigned";

    if (!paceByMonth[month] && (curTotalWD || curPassedWD)) {
      paceByMonth[month] = { total: curTotalWD, passed: curPassedWD };
    }

    byMonth[month] = byMonth[month] || {};
    byMonth[month][section] = byMonth[month][section] || {};
    byMonth[month][section][unit] = byMonth[month][section][unit] || {};
    const cells = {};
    BRANDS.forEach((b) => (cells[b] = { target: 0, achv: 0, lastMonth: 0 }));
    byMonth[month][section][unit][areaRaw] = cells;

    for (let i = baseCols; i < row.length; i++) {
      const meta = colMeta[i];
      if (!meta || !meta.brand) continue;
      const brandKey = BRANDS.find((b) => b.toLowerCase() === meta.brand.toLowerCase());
      if (!brandKey) continue; // e.g. the sheet's own "Total" columns
      const val = parseFloat(String(row[i]).replace(/,/g, "")) || 0;
      if (meta.metric === "rf") cells[brandKey].target = val;
      else if (meta.metric === "result") cells[brandKey].achv = val;
    }
  }

  const monthKeys = Object.keys(byMonth);
  if (!monthKeys.length) return null;

  const toSectionsArray = (sectionMap) =>
    Object.keys(sectionMap).map((secName) => ({
      name: secName,
      units: Object.keys(sectionMap[secName]).map((unitName) => ({
        name: unitName,
        areas: Object.keys(sectionMap[secName][unitName]).map((areaName) => ({
          name: areaName,
          brands: sectionMap[secName][unitName][areaName],
        })),
      })),
    }));

  const byMonthArrays = {};
  monthKeys.forEach((m) => (byMonthArrays[m] = toSectionsArray(byMonth[m])));

  if (!hasMonthCol) {
    return { hasMonths: false, months: [], byMonth: {}, sections: byMonthArrays["current"], pace: paceByMonth["current"] || null };
  }
  const months = monthKeys.sort((a, b) => {
    const da = parseMonthKey(a);
    const db = parseMonthKey(b);
    if (da && db) return db - da;
    return b.localeCompare(a);
  });
  return { hasMonths: true, months, byMonth: byMonthArrays, paceByMonth, sections: byMonthArrays[months[0]] };
}

// SKU_Daily_Data sheet -> raw rows used for working-day cumulative comparisons
function buildLiveSkuRows(rawRows) {
  const rows = rawRows
    .map(normalizeRow)
    .map((r) => {
      const dateStr = String(getField(r, ["day", "date"])).trim();
      const d = new Date(dateStr);
      return {
        date: d,
        workingDay: parseInt(getField(r, ["workingday"], 0), 10) || 0,
        brand: String(getField(r, ["brand"])).trim(),
        sku: String(getField(r, ["sku"])).trim(),
        district: String(getField(r, ["district"], "Unassigned")).trim() || "Unassigned",
        team: String(getField(r, ["team"], "Unassigned")).trim() || "Unassigned",
        amount: parseFloat(getField(r, ["qtykg", "qty", "achievementkg", "achievement", "amount"], 0)) || 0,
        invoice: parseFloat(getField(r, ["invoice"], 0)) || 0,
        pkt: parseFloat(getField(r, ["pktpc", "pkt"], 0)) || 0,
        qty: parseFloat(getField(r, ["qtykg", "qty"], 0)) || 0,
        amountTk: parseFloat(getField(r, ["amounttk", "amount"], 0)) || 0,
      };
    })
    .filter((r) => !isNaN(r.date.getTime()) && r.brand && r.sku && r.workingDay > 0);
  return rows.length ? rows : null;
}

function getAvailableMonths(skuRows) {
  const set = new Set();
  skuRows.forEach((r) => set.add(`${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`));
  return Array.from(set).sort().reverse();
}
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function monthLabel(ym) {
  const [y, m] = ym.split("-").map(Number);
  return `${MONTH_NAMES[m - 1]} ${y}`;
}

function normText(s) {
  return String(s || "").trim().toLowerCase();
}
function zeroTrend(maxDay = 26) {
  return Array.from({ length: maxDay }, (_, i) => ({ day: `D${i + 1}`, "Running Month": 0, "Last Month": 0, "Last Year": 0 }));
}
function computeWorkingDayTrend(skuRows, brand, sku, refMonth, maxDay = 26) {
  const filtered = skuRows.filter((r) => normText(r.brand) === normText(brand) && normText(r.sku) === normText(sku));
  if (!filtered.length) return null;
  let curY, curM;
  if (refMonth) {
    [curY, curM] = refMonth.split("-").map(Number);
    curM -= 1;
  } else {
    const latestDate = filtered.reduce((a, r) => (r.date > a ? r.date : a), filtered[0].date);
    curY = latestDate.getFullYear();
    curM = latestDate.getMonth();
  }
  let lmY = curY;
  let lmM = curM - 1;
  if (lmM < 0) {
    lmM = 11;
    lmY -= 1;
  }
  const lyY = curY - 1;
  const lyM = curM;

  const bucket = (matchFn) => {
    const daySum = {};
    filtered.forEach((r) => {
      if (matchFn(r.date)) daySum[r.workingDay] = (daySum[r.workingDay] || 0) + r.amount;
    });
    let cum = 0;
    const out = [];
    for (let d = 1; d <= maxDay; d++) {
      cum += daySum[d] || 0;
      out.push(Math.round(cum));
    }
    return out;
  };

  const running = bucket((d) => d.getFullYear() === curY && d.getMonth() === curM);
  const lastMonth = bucket((d) => d.getFullYear() === lmY && d.getMonth() === lmM);
  const lastYear = bucket((d) => d.getFullYear() === lyY && d.getMonth() === lyM);

  return Array.from({ length: maxDay }, (_, i) => ({
    day: `D${i + 1}`,
    "Running Month": running[i],
    "Last Month": lastMonth[i],
    "Last Year": lastYear[i],
  }));
}

// Same-working-day snapshot across 4 metrics (Invoice count, PKT, KG, Amount) for a specific SKU:
// Running Month (up to the latest working day reached) vs Last Month vs Last Year, at that same working day.
function computeMultiMetricSnapshot(skuRows, brand, sku, refMonth) {
  const filtered = skuRows.filter((r) => normText(r.brand) === normText(brand) && normText(r.sku) === normText(sku));
  if (!filtered.length) return null;
  let curY, curM;
  if (refMonth) {
    [curY, curM] = refMonth.split("-").map(Number);
    curM -= 1;
  } else {
    const latestDate = filtered.reduce((a, r) => (r.date > a ? r.date : a), filtered[0].date);
    curY = latestDate.getFullYear();
    curM = latestDate.getMonth();
  }
  let lmY = curY;
  let lmM = curM - 1;
  if (lmM < 0) {
    lmM = 11;
    lmY -= 1;
  }
  const lyY = curY - 1;
  const lyM = curM;

  const runningRows = filtered.filter((r) => r.date.getFullYear() === curY && r.date.getMonth() === curM);
  if (!runningRows.length) return null;
  const targetDay = runningRows.reduce((a, r) => Math.max(a, r.workingDay), 0);

  const sumUpTo = (matchFn, metric) =>
    filtered.filter((r) => matchFn(r.date) && r.workingDay <= targetDay).reduce((acc, r) => acc + (r[metric] || 0), 0);

  const metrics = ["invoice", "pkt", "qty", "amountTk"];
  const out = { day: targetDay };
  metrics.forEach((m) => {
    const running = sumUpTo((d) => d.getFullYear() === curY && d.getMonth() === curM, m);
    const lastMonth = sumUpTo((d) => d.getFullYear() === lmY && d.getMonth() === lmM, m);
    const lastYear = sumUpTo((d) => d.getFullYear() === lyY && d.getMonth() === lyM, m);
    out[m] = {
      running: Math.round(running * 100) / 100,
      lastMonth: Math.round(lastMonth * 100) / 100,
      lastYear: Math.round(lastYear * 100) / 100,
      vsMonth: lastMonth ? ((running - lastMonth) / lastMonth) * 100 : 0,
      vsYear: lastYear ? ((running - lastYear) / lastYear) * 100 : 0,
    };
  });
  return out;
}

// Same as computeWorkingDayTrend but returns cumulative day-wise series for all 4 metrics at once
// (Invoice, PKT, KG, Amount) instead of a single blended "amount" field.
function computeWorkingDayTrendMulti(skuRows, brand, sku, refMonth, maxDay = 26) {
  const filtered = skuRows.filter((r) => normText(r.brand) === normText(brand) && normText(r.sku) === normText(sku));
  if (!filtered.length) return null;
  let curY, curM;
  if (refMonth) {
    [curY, curM] = refMonth.split("-").map(Number);
    curM -= 1;
  } else {
    const latestDate = filtered.reduce((a, r) => (r.date > a ? r.date : a), filtered[0].date);
    curY = latestDate.getFullYear();
    curM = latestDate.getMonth();
  }
  let lmY = curY;
  let lmM = curM - 1;
  if (lmM < 0) {
    lmM = 11;
    lmY -= 1;
  }
  const lyY = curY - 1;
  const lyM = curM;

  const bucket = (matchFn, metric) => {
    const daySum = {};
    filtered.forEach((r) => {
      if (matchFn(r.date)) daySum[r.workingDay] = (daySum[r.workingDay] || 0) + (r[metric] || 0);
    });
    let cum = 0;
    const out = [];
    for (let d = 1; d <= maxDay; d++) {
      cum += daySum[d] || 0;
      out.push(Math.round(cum * 100) / 100);
    }
    return out;
  };

  const result = {};
  ["invoice", "pkt", "qty", "amountTk"].forEach((metric) => {
    const running = bucket((d) => d.getFullYear() === curY && d.getMonth() === curM, metric);
    const lastMonth = bucket((d) => d.getFullYear() === lmY && d.getMonth() === lmM, metric);
    const lastYear = bucket((d) => d.getFullYear() === lyY && d.getMonth() === lyM, metric);
    result[metric] = Array.from({ length: maxDay }, (_, i) => ({
      day: `D${i + 1}`,
      "Running Month": running[i],
      "Last Month": lastMonth[i],
      "Last Year": lastYear[i],
    }));
  });
  return result;
}
function zeroTrendMulti(maxDay = 26) {
  const out = {};
  ["invoice", "pkt", "qty", "amountTk"].forEach((m) => (out[m] = zeroTrend(maxDay)));
  return out;
}
// Demo-mode fallback: derive plausible Invoice/PKT/Amount series from the mock KG series.
function workingDayTrendMulti(seed, workingDays = 24) {
  const qty = workingDayTrend(seed, workingDays);
  const scale = (mult) =>
    qty.map((r) => ({
      day: r.day,
      "Running Month": Math.round(r["Running Month"] * mult * 100) / 100,
      "Last Month": Math.round(r["Last Month"] * mult * 100) / 100,
      "Last Year": Math.round(r["Last Year"] * mult * 100) / 100,
    }));
  return {
    invoice: scale(1 / 38),
    pkt: scale(1 / 0.45),
    qty,
    amountTk: scale(115),
  };
}

// Ranks a group field (District or Team) by current-month KG achievement, across ALL brands/SKUs,
// with MoM/YoY growth so it's easy to spot who's on top and who's lagging.
function computeGroupPerformance(skuRows, groupField, refMonth, brand = null, sku = null) {
  if (!skuRows || !skuRows.length) return [];
  const rows = brand && sku ? skuRows.filter((r) => normText(r.brand) === normText(brand) && normText(r.sku) === normText(sku)) : skuRows;
  if (!rows.length) return [];
  let curY, curM;
  if (refMonth) {
    [curY, curM] = refMonth.split("-").map(Number);
    curM -= 1;
  } else {
    const latestDate = rows.reduce((a, r) => (r.date > a ? r.date : a), rows[0].date);
    curY = latestDate.getFullYear();
    curM = latestDate.getMonth();
  }
  let lmY = curY;
  let lmM = curM - 1;
  if (lmM < 0) {
    lmM = 11;
    lmY -= 1;
  }
  const lyY = curY - 1;
  const lyM = curM;

  const groups = {};
  rows.forEach((r) => {
    const key = r[groupField] || "Unassigned";
    if (!groups[key]) groups[key] = { name: key, kg: 0, amountTk: 0, lmKg: 0, lyKg: 0 };
    if (r.date.getFullYear() === curY && r.date.getMonth() === curM) {
      groups[key].kg += r.qty;
      groups[key].amountTk += r.amountTk;
    } else if (r.date.getFullYear() === lmY && r.date.getMonth() === lmM) {
      groups[key].lmKg += r.qty;
    } else if (r.date.getFullYear() === lyY && r.date.getMonth() === lyM) {
      groups[key].lyKg += r.qty;
    }
  });

  return Object.values(groups)
    .map((g) => ({
      ...g,
      kg: Math.round(g.kg),
      amountTk: Math.round(g.amountTk),
      vsMonth: g.lmKg ? ((g.kg - g.lmKg) / g.lmKg) * 100 : 0,
      vsYear: g.lyKg ? ((g.kg - g.lyKg) / g.lyKg) * 100 : 0,
    }))
    .filter((g) => g.kg > 0 || g.lmKg > 0 || g.lyKg > 0)
    .sort((a, b) => b.vsMonth - a.vsMonth);
}

function computeSkuComparisonLive(skuRows, brand, refMonth) {
  const results = (SKU_LIST[brand] || []).map((s) => {
    const t = computeWorkingDayTrend(skuRows, brand, s, refMonth);
    if (t) {
      const last = t[t.length - 1];
      const vsMonth = last["Last Month"] ? ((last["Running Month"] - last["Last Month"]) / last["Last Month"]) * 100 : 0;
      const vsYear = last["Last Year"] ? ((last["Running Month"] - last["Last Year"]) / last["Last Year"]) * 100 : 0;
      return { sku: s, "Running Month": last["Running Month"], "Last Month": last["Last Month"], "Last Year": last["Last Year"], vsMonth, vsYear };
    }
    // No real rows for this SKU yet — show zeros rather than silently swapping in mock data
    return { sku: s, "Running Month": 0, "Last Month": 0, "Last Year": 0, vsMonth: 0, vsYear: 0 };
  });
  return results.length ? results : null;
}

function computeWorkingDayPace(skuRows) {
  if (!skuRows || !skuRows.length) return null;
  const latestDate = skuRows.reduce((a, r) => (r.date > a ? r.date : a), skuRows[0].date);
  const curY = latestDate.getFullYear();
  const curM = latestDate.getMonth();
  let passed = 0;
  skuRows.forEach((r) => {
    if (r.date.getFullYear() === curY && r.date.getMonth() === curM && r.workingDay > passed) passed = r.workingDay;
  });
  let lmY = curY;
  let lmM = curM - 1;
  if (lmM < 0) {
    lmM = 11;
    lmY -= 1;
  }
  let total = 0;
  skuRows.forEach((r) => {
    if (r.date.getFullYear() === lmY && r.date.getMonth() === lmM && r.workingDay > total) total = r.workingDay;
  });
  if (!total) total = 26;
  return { passed, total };
}

// ---------- Excel export ----------
function exportSkuComparisonToExcel(skuComparison, brand, monthLabelText) {
  const rows = skuComparison.map((d) => ({
    SKU: d.sku,
    "Running Month (KG)": d["Running Month"],
    "Last Month (KG)": d["Last Month"],
    "Last Year (KG)": d["Last Year"],
    "MoM %": Number(d.vsMonth.toFixed(1)),
    "YoY %": Number(d.vsYear.toFixed(1)),
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "SKU Comparison");
  XLSX.writeFile(wb, `${brand}-SKU-Analysis-${monthLabelText || "latest"}.xlsx`);
}

function exportDashboardToExcel(activeSections, brand) {
  const rows = [];
  activeSections.forEach((sec) =>
    sec.units.forEach((u) =>
      u.areas.forEach((a) => {
        const s = getCell(a.brands, brand);
        rows.push({
          Section: sec.name,
          Unit: u.name,
          Area: a.name,
          "Rolling Forecast (KG)": s.target,
          "Result (KG)": s.achv,
          "Progress %": Number(achPct(s).toFixed(1)),
        });
      })
    )
  );
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Target vs Progress");
  XLSX.writeFile(wb, `${brand}-Target-vs-Progress.xlsx`);
}




const fmt = (n) => `${new Intl.NumberFormat("en-BD").format(Math.round(n))} kg`;
const fmtNum = (n) => new Intl.NumberFormat("en-BD").format(Math.round(n));
const TOTAL_LABEL = "Total";
function getCell(areaBrands, brand) {
  if (brand === TOTAL_LABEL) {
    return BRANDS.reduce(
      (acc, b) => {
        const c = areaBrands[b];
        if (c) {
          acc.target += c.target;
          acc.achv += c.achv;
          acc.lastMonth += c.lastMonth;
        }
        return acc;
      },
      { target: 0, achv: 0, lastMonth: 0 }
    );
  }
  return areaBrands[brand];
}
function sumBrand(areas, brand) {
  return areas.reduce(
    (acc, a) => {
      const c = getCell(a.brands, brand);
      acc.target += c.target;
      acc.achv += c.achv;
      acc.lastMonth += c.lastMonth;
      return acc;
    },
    { target: 0, achv: 0, lastMonth: 0 }
  );
}
function sumUnits(units, brand) {
  return units.reduce(
    (acc, u) => {
      const s = sumBrand(u.areas, brand);
      acc.target += s.target;
      acc.achv += s.achv;
      acc.lastMonth += s.lastMonth;
      return acc;
    },
    { target: 0, achv: 0, lastMonth: 0 }
  );
}
function achPct(s) {
  return s.target ? (s.achv / s.target) * 100 : 0;
}
function growthPct(s) {
  return s.lastMonth ? ((s.achv - s.lastMonth) / s.lastMonth) * 100 : 0;
}

const NAVY = "#0A2647";
const NAVY_LIGHT = "#12395F";
const RED = "#C81D25";
const GREEN = "#1F9254";
const GOLD = "#D9A441";
const CREAM = "#F4F5F8";

function AchBadge({ pct }) {
  const good = pct >= 100;
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold"
      style={{
        background: good ? "rgba(31,146,84,0.12)" : "rgba(200,29,37,0.1)",
        color: good ? GREEN : RED,
      }}
    >
      {pct.toFixed(1)}%
    </span>
  );
}

function GrowthBadge({ pct }) {
  const up = pct >= 0;
  const Icon = up ? TrendingUp : TrendingDown;
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold"
      style={{ background: up ? "rgba(31,146,84,0.12)" : "rgba(200,29,37,0.1)", color: up ? GREEN : RED }}
    >
      <Icon size={12} />
      {Math.abs(pct).toFixed(1)}%
    </span>
  );
}

function DegrowthList({ title, data }) {
  return (
    <div className="bg-white rounded-2xl p-5" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
      <h2 className="text-sm font-bold mb-3" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>{title}</h2>
      {data.length === 0 ? (
        <p className="text-xs text-slate-400">No degrowth — everyone is at or above last month.</p>
      ) : (
        <div className="divide-y divide-slate-100">
          {data.map((g) => (
            <div key={g.name} className="flex items-center justify-between py-2 text-sm">
              <span className="font-medium text-slate-700">{g.name}</span>
              <div className="flex items-center gap-3">
                <span className="text-xs text-slate-400">{fmt(g.kg)} kg</span>
                <GrowthBadge pct={g.vsMonth} />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ProgressBar({ pct }) {
  const clamped = Math.min(pct, 130);
  return (
    <div className="w-28 h-1.5 rounded-full bg-slate-200 overflow-hidden">
      <div
        className="h-full rounded-full"
        style={{ width: `${Math.min(clamped, 100)}%`, background: pct >= 100 ? GREEN : NAVY_LIGHT }}
      />
    </div>
  );
}

// ---------- Row components ----------
function AreaRow({ area, brand }) {
  const s = getCell(area.brands, brand);
  const pct = achPct(s);
  return (
    <tr className="text-sm border-b border-slate-100 hover:bg-slate-50">
      <td className="py-2 pl-16 text-slate-600">{area.name}</td>
      <td className="py-2 text-right pr-4 text-slate-500">{fmt(s.target)}</td>
      <td className="py-2 text-right pr-4 font-medium text-slate-700">{fmt(s.achv)}</td>
      <td className="py-2 pr-6">
        <div className="flex items-center gap-2 justify-end">
          <ProgressBar pct={pct} />
          <AchBadge pct={pct} />
        </div>
      </td>
    </tr>
  );
}

function UnitBlock({ unit, brand }) {
  const [open, setOpen] = useState(false);
  const s = sumBrand(unit.areas, brand);
  const pct = achPct(s);
  return (
    <>
      <tr className="text-sm border-b border-slate-100 bg-slate-50/60 cursor-pointer" onClick={() => setOpen(!open)}>
        <td className="py-2.5 pl-8 font-medium text-slate-700 flex items-center gap-1">
          {open ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          {unit.name}
        </td>
        <td className="py-2.5 text-right pr-4 text-slate-500">{fmt(s.target)}</td>
        <td className="py-2.5 text-right pr-4 font-medium text-slate-700">{fmt(s.achv)}</td>
        <td className="py-2.5 pr-6">
          <div className="flex items-center gap-2 justify-end">
            <ProgressBar pct={pct} />
            <AchBadge pct={pct} />
          </div>
        </td>
      </tr>
      {open && unit.areas.map((a) => <AreaRow key={a.name} area={a} brand={brand} />)}
    </>
  );
}

function SectionBlock({ section, brand }) {
  const [open, setOpen] = useState(true);
  const s = sumUnits(section.units, brand);
  const pct = achPct(s);
  return (
    <>
      <tr className="text-sm cursor-pointer" onClick={() => setOpen(!open)} style={{ background: NAVY }}>
        <td className="py-3 pl-4 font-semibold text-white flex items-center gap-1.5">
          {open ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
          {section.name}
        </td>
        <td className="py-3 text-right pr-4 text-white/80">{fmt(s.target)}</td>
        <td className="py-3 text-right pr-4 font-semibold text-white">{fmt(s.achv)}</td>
        <td className="py-3 pr-6">
          <div className="flex items-center gap-2 justify-end">
            <AchBadge pct={pct} />
          </div>
        </td>
      </tr>
      {open && section.units.map((u) => <UnitBlock key={u.name} unit={u} brand={brand} />)}
    </>
  );
}

// All-SKU comparison for a brand: Running Month / Last Month / Last Year totals + growth
function buildSkuComparison(brandName) {
  return (SKU_LIST[brandName] || []).map((s) => {
    const t = workingDayTrend(s.length * 7 + brandName.length);
    const last = t[t.length - 1];
    const vsMonth = ((last["Running Month"] - last["Last Month"]) / last["Last Month"]) * 100;
    const vsYear = ((last["Running Month"] - last["Last Year"]) / last["Last Year"]) * 100;
    return {
      sku: s,
      "Running Month": last["Running Month"],
      "Last Month": last["Last Month"],
      "Last Year": last["Last Year"],
      vsMonth,
      vsYear,
    };
  });
}

// ---------- Product-wise circular progress ----------
function RadialProgress({ label, pct, target, achv, color }) {
  const r = 46;
  const c = 2 * Math.PI * r;
  const clamped = Math.max(0, Math.min(pct, 100));
  const offset = c - (clamped / 100) * c;
  return (
    <div
      className="flex-1 min-w-[180px] rounded-2xl p-5 flex flex-col items-center relative overflow-hidden"
      style={{ background: "white", boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -12px rgba(10,38,71,0.15)" }}
    >
      <div className="absolute top-0 left-0 right-0 h-1" style={{ background: `linear-gradient(90deg, ${color}, ${color}55)` }} />
      <div className="relative w-32 h-32 mt-1">
        <svg width="128" height="128" viewBox="0 0 128 128" className="-rotate-90">
          <circle cx="64" cy="64" r={r} fill="none" stroke="#EEF1F5" strokeWidth="11" />
          <circle
            cx="64"
            cy="64"
            r={r}
            fill="none"
            stroke={color}
            strokeWidth="11"
            strokeDasharray={c}
            strokeDashoffset={offset}
            strokeLinecap="round"
            style={{ transition: "stroke-dashoffset 0.8s ease" }}
          />
        </svg>
        <div className="absolute inset-0 flex flex-col items-center justify-center">
          <span className="text-2xl font-extrabold tracking-tight" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>{pct.toFixed(1)}%</span>
          <span className="text-[10px] uppercase tracking-wider text-slate-400 font-medium">progress</span>
        </div>
      </div>
      <div className="mt-3 text-sm font-bold" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>{label}</div>
      <div className="mt-2 flex gap-4 text-center">
        <div>
          <div className="text-[10px] uppercase tracking-wide text-slate-400">Rolling Forecast</div>
          <div className="text-xs font-semibold text-slate-600">{fmt(target)}</div>
        </div>
        <div className="w-px bg-slate-200" />
        <div>
          <div className="text-[10px] uppercase tracking-wide text-slate-400">Result</div>
          <div className="text-xs font-semibold" style={{ color }}>{fmt(achv)}</div>
        </div>
      </div>
    </div>
  );
}

const BRAND_COLORS = { "AJI-Retail": NAVY, "AJI-Bulk": GOLD, "Hapima": RED, "TasteMate": GREEN, "Total": "#475569", "AJI-NO-MOTO": NAVY, "HAPIMA": RED };
const TREND_METRIC_CONFIG = [
  { key: "invoice", label: "Invoice", color: NAVY, decimals: 0, tickDivide: false },
  { key: "pkt", label: "PKT (pc)", color: GOLD, decimals: 0, tickDivide: false },
  { key: "qty", label: "KG", color: GREEN, decimals: 1, tickDivide: true },
  { key: "amountTk", label: "Amount (Tk)", color: "#7C3AED", decimals: 0, tickDivide: true },
];

function ProductProgressRow({ allAreas }) {
  return (
    <div className="flex gap-3 mb-6 flex-wrap">
      {BRANDS.map((b) => {
        const s = sumBrand(allAreas, b);
        return (
          <RadialProgress
            key={b}
            label={b}
            pct={achPct(s)}
            target={s.target}
            achv={s.achv}
            color={BRAND_COLORS[b]}
          />
        );
      })}
    </div>
  );
}

function MetricSnapshotCard({ label, icon: Icon, color, data, unit, decimals = 0 }) {
  const fmtV = (n) => new Intl.NumberFormat("en-BD", { maximumFractionDigits: decimals }).format(n || 0);
  return (
    <div className="rounded-2xl p-4 bg-white" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
      <div className="flex items-center justify-between mb-2">
        <span className="text-[11px] font-semibold text-slate-400 uppercase tracking-wider">{label}</span>
        <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: `${color}18` }}>
          <Icon size={14} style={{ color }} />
        </div>
      </div>
      <div className="text-xl font-extrabold tracking-tight" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>
        {fmtV(data.running)}{unit}
      </div>
      <div className="flex items-center gap-3 mt-2 flex-wrap">
        <div className="flex items-center gap-1 text-[10px] text-slate-400 font-medium">MoM <GrowthBadge pct={data.vsMonth} /></div>
        <div className="flex items-center gap-1 text-[10px] text-slate-400 font-medium">YoY <GrowthBadge pct={data.vsYear} /></div>
      </div>
    </div>
  );
}

// ---------- KPI cards ----------
function KpiCard({ label, value, sub, icon: Icon, accent, badge }) {
  return (
    <div
      className="flex-1 min-w-[150px] rounded-2xl p-4 bg-white"
      style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}
    >
      <div className="flex items-center justify-between mb-3">
        <span className="text-[11px] font-semibold text-slate-400 uppercase tracking-wider">{label}</span>
        <div className="flex items-center gap-2">
          {badge}
          <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: `${accent}18` }}>
            <Icon size={14} style={{ color: accent }} />
          </div>
        </div>
      </div>
      <div className="text-2xl font-extrabold tracking-tight" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>{value}</div>
      {sub && <div className="text-xs text-slate-400 mt-1">{sub}</div>}
    </div>
  );
}

// ---------- Main dashboard ----------
export default function Dashboard() {
  const [page, setPage] = useState("dashboard"); // "dashboard" | "sku"
  const [brand, setBrand] = useState(BRANDS[0]);
  const [skuBrand, setSkuBrand] = useState(SKU_BRANDS[0]);
  const [sku, setSku] = useState(SKU_LIST[SKU_BRANDS[0]][0]);
  const [targetData, setTargetData] = useState(null);
  const [liveSkuRows, setLiveSkuRows] = useState(null);
  const [dataStatus, setDataStatus] = useState("loading"); // loading | live | mock
  const [selectedMonth, setSelectedMonth] = useState(null); // "YYYY-MM" or null = latest (SKU page)
  const [selectedTargetMonth, setSelectedTargetMonth] = useState(null); // Target sheet month (page 1)

  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    return () => document.head.removeChild(link);
  }, []);

  useEffect(() => {
    let cancelled = false;
    Promise.all([fetchCsvRaw(TARGET_CSV_URL), fetchCsv(SKU_CSV_URL)])
      .then(([targetRows, skuRows]) => {
        if (cancelled) return;
        const sections = buildLiveSectionsWide(targetRows);
        const skus = buildLiveSkuRows(skuRows);
        if (sections) setTargetData(sections);
        if (skus) setLiveSkuRows(skus);
        setDataStatus(sections || skus ? "live" : "mock");
      })
      .catch(() => {
        if (!cancelled) setDataStatus("mock");
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const activeSections = useMemo(() => {
    if (!targetData) return SECTIONS;
    if (targetData.hasMonths) {
      const m = selectedTargetMonth || targetData.months[0];
      return targetData.byMonth[m] || SECTIONS;
    }
    return targetData.sections || SECTIONS;
  }, [targetData, selectedTargetMonth]);
  const allAreas = useMemo(() => activeSections.flatMap((s) => s.units.flatMap((u) => u.areas)), [activeSections]);
  const totals = useMemo(() => sumBrand(allAreas, brand), [allAreas, brand]);
  const availableMonths = useMemo(() => (liveSkuRows ? getAvailableMonths(liveSkuRows) : []), [liveSkuRows]);
  const trendMulti = useMemo(() => {
    if (liveSkuRows) {
      const t = computeWorkingDayTrendMulti(liveSkuRows, skuBrand, sku, selectedMonth);
      return t || zeroTrendMulti();
    }
    return workingDayTrendMulti(sku.length * 7 + skuBrand.length);
  }, [sku, skuBrand, liveSkuRows, selectedMonth]);
  const trend = trendMulti.qty;
  const last = trend[trend.length - 1];
  const vsLastMonth = last["Last Month"] ? ((last["Running Month"] - last["Last Month"]) / last["Last Month"]) * 100 : 0;
  const vsLastYear = last["Last Year"] ? ((last["Running Month"] - last["Last Year"]) / last["Last Year"]) * 100 : 0;
  const skuComparison = useMemo(() => {
    if (liveSkuRows) {
      const c = computeSkuComparisonLive(liveSkuRows, skuBrand, selectedMonth);
      if (c) return c;
    }
    return buildSkuComparison(skuBrand);
  }, [skuBrand, liveSkuRows, selectedMonth]);
  const districtPerf = useMemo(() => (liveSkuRows ? computeGroupPerformance(liveSkuRows, "district", selectedMonth, skuBrand, sku).filter((g) => g.vsMonth < 0) : []), [liveSkuRows, selectedMonth, skuBrand, sku]);
  const teamPerf = useMemo(() => (liveSkuRows ? computeGroupPerformance(liveSkuRows, "team", selectedMonth, skuBrand, sku).filter((g) => g.vsMonth < 0) : []), [liveSkuRows, selectedMonth, skuBrand, sku]);
  const multiMetric = useMemo(() => {
    if (!liveSkuRows) return null;
    return computeMultiMetricSnapshot(liveSkuRows, skuBrand, sku, selectedMonth);
  }, [liveSkuRows, skuBrand, sku, selectedMonth]);
  const lastUpdatedLabel = useMemo(() => {
    if (!liveSkuRows || !liveSkuRows.length) return null;
    const maxDate = liveSkuRows.reduce((a, r) => (r.date > a ? r.date : a), liveSkuRows[0].date);
    return maxDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  }, [liveSkuRows]);
  const isCurrentMonthView = !targetData?.hasMonths || !selectedTargetMonth || selectedTargetMonth === targetData.months[0];
  const workingDayPace = useMemo(() => {
    let sheetPace = null;
    if (targetData) {
      if (targetData.hasMonths) {
        const m = selectedTargetMonth || targetData.months[0];
        sheetPace = targetData.paceByMonth?.[m] || null;
      } else {
        sheetPace = targetData.pace || null;
      }
    }
    if (sheetPace) return sheetPace;
    if (isCurrentMonthView && liveSkuRows) return computeWorkingDayPace(liveSkuRows);
    return null;
  }, [targetData, selectedTargetMonth, isCurrentMonthView, liveSkuRows]);
  const expectedPct = workingDayPace && workingDayPace.total ? (workingDayPace.passed / workingDayPace.total) * 100 : null;
  const actualPct = achPct(totals);
  const onTrack = expectedPct !== null ? actualPct >= expectedPct - 3 : null;

  const handleBrand = (b) => {
    setBrand(b);
  };
  const handleSkuBrand = (b) => {
    setSkuBrand(b);
    setSku(SKU_LIST[b][0]);
  };
  const monthSubLabel =
    targetData?.hasMonths && selectedTargetMonth && selectedTargetMonth !== targetData.months[0] ? `KG, ${selectedTargetMonth}` : "KG, current month";

  return (
    <div className="min-h-screen" style={{ background: CREAM, fontFamily: "'Inter', sans-serif" }}>
      {/* Gradient header banner */}
      <div style={{ background: `linear-gradient(120deg, ${NAVY} 0%, #16406B 60%, #0A2647 100%)` }}>
        <div className="max-w-6xl mx-auto px-4 pt-7 pb-8">
          <div className="flex items-center justify-between flex-wrap gap-4">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <div className="w-1.5 h-1.5 rounded-full" style={{ background: GOLD }} />
                <span
                  className="flex items-center gap-1 text-[10px] font-semibold px-2 py-0.5 rounded-full"
                  style={{
                    background: dataStatus === "live" ? "rgba(31,146,84,0.25)" : dataStatus === "loading" ? "rgba(255,255,255,0.12)" : "rgba(200,29,37,0.22)",
                    color: dataStatus === "live" ? "#5FE3A0" : dataStatus === "loading" ? "rgba(255,255,255,0.75)" : "#FF9B9B",
                  }}
                >
                  {dataStatus === "loading" && <Loader2 size={10} className="animate-spin" />}
                  {dataStatus === "live" && <Wifi size={10} />}
                  {dataStatus === "mock" && <WifiOff size={10} />}
                  {dataStatus === "loading" ? "Connecting…" : dataStatus === "live" ? "Live from Sheet" : "Demo data"}
                </span>
                {page === "dashboard" && targetData?.hasMonths && selectedTargetMonth && selectedTargetMonth !== targetData.months[0] ? (
                  <span className="text-[11px] text-white/50 font-medium">{selectedTargetMonth}</span>
                ) : (
                  lastUpdatedLabel && <span className="text-[11px] text-white/50 font-medium">Updated till {lastUpdatedLabel}</span>
                )}
              </div>
              <h1 className="text-2xl font-extrabold text-white" style={{ fontFamily: "'Sora', sans-serif" }}>
                {page === "dashboard" ? "ROLLING FORECAST DASHBOARD" : "SKU-wise Analysis"}
              </h1>
            </div>
            <div className="flex items-center gap-3 flex-wrap">
              {page === "dashboard" ? (
                <div className="flex gap-1 bg-white/10 backdrop-blur rounded-xl p-1 border border-white/10">
                  {[...BRANDS, TOTAL_LABEL].map((b) => (
                    <button
                      key={b}
                      onClick={() => handleBrand(b)}
                      className="px-3.5 py-1.5 rounded-lg text-sm font-semibold transition-all"
                      style={{
                        background: brand === b ? "white" : "transparent",
                        color: brand === b ? NAVY : "rgba(255,255,255,0.75)",
                      }}
                    >
                      {b}
                    </button>
                  ))}
                </div>
              ) : (
                <button
                  onClick={() => setPage("dashboard")}
                  className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold bg-white/10 border border-white/10 text-white hover:bg-white/20 transition-all"
                >
                  <ArrowLeft size={14} /> Back to Dashboard
                </button>
              )}
              {page === "dashboard" && targetData?.hasMonths && (
                <div className="flex items-center gap-1.5 bg-white/10 backdrop-blur rounded-xl px-3 py-2 border border-white/10">
                  <Calendar size={14} className="text-white/80" />
                  <select
                    value={selectedTargetMonth || targetData.months[0]}
                    onChange={(e) => setSelectedTargetMonth(e.target.value)}
                    className="text-sm font-semibold text-white bg-transparent outline-none"
                    style={{ colorScheme: "dark" }}
                  >
                    {targetData.months.map((m) => (
                      <option key={m} value={m} style={{ color: NAVY }}>{m}</option>
                    ))}
                  </select>
                </div>
              )}
              {page === "dashboard" && (
                <button
                  onClick={() => setPage("sku")}
                  className="flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-sm font-semibold transition-all ml-2"
                  style={{ background: GOLD, color: NAVY }}
                >
                  <LineChartIcon size={14} /> SKU-wise Analysis
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 -mt-5 pb-8">
      {page === "dashboard" ? (
      <>

        {/* Product-wise total progress (3 shapes) */}
        <ProductProgressRow allAreas={allAreas} />

        {isCurrentMonthView && workingDayPace && workingDayPace.total > 0 && (
          <p className="text-xs text-slate-400 mb-3 -mt-1">
            Day pass {expectedPct !== null ? expectedPct.toFixed(0) : "0"}% ({workingDayPace.passed}/{workingDayPace.total} working days)
          </p>
        )}

        {/* KPI row */}
        <div className="flex gap-3 mb-6 flex-wrap">
          <KpiCard label="Rolling Forecast" value={fmt(totals.target)} sub={monthSubLabel} icon={Target} accent={NAVY} />
          <KpiCard label="Result" value={fmt(totals.achv)} sub={monthSubLabel} icon={Target} accent={GREEN} />
          <KpiCard
            label="Progress"
            value={`${achPct(totals).toFixed(1)}%`}
            icon={Percent}
            accent={achPct(totals) >= 100 ? GREEN : RED}
            badge={
              isCurrentMonthView && workingDayPace && workingDayPace.total > 0 ? (
                <span className="flex items-center gap-1 text-[11px] font-semibold whitespace-nowrap" style={{ color: onTrack ? GREEN : RED }}>
                  {onTrack ? "🟢" : "🔴"} {onTrack ? "On Track" : "Off Track"}
                </span>
              ) : null
            }
          />
        </div>

        {/* Hierarchy table */}
        <div className="bg-white rounded-2xl overflow-hidden mb-3" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
          <table className="w-full">
            <thead>
              <tr className="text-[11px] text-slate-400 uppercase tracking-wider border-b border-slate-100">
                <th className="text-left py-3 pl-4 font-semibold">Section / Unit / Area</th>
                <th className="text-right py-3 pr-4 font-semibold">Rolling Forecast</th>
                <th className="text-right py-3 pr-4 font-semibold">Result</th>
                <th className="text-right py-3 pr-6 font-semibold">Progress</th>
              </tr>
            </thead>
            <tbody>
              {activeSections.map((s) => (
                <SectionBlock key={s.name} section={s} brand={brand} />
              ))}
            </tbody>
          </table>
        </div>
        <div className="flex justify-end mb-6">
          <button
            onClick={() => exportDashboardToExcel(activeSections, brand)}
            className="flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-medium text-slate-400 hover:text-slate-600 transition-colors"
          >
            <Download size={11} /> Export Excel
          </button>
        </div>
      </>
      ) : (
      <>
        {/* Brand + month selector for SKU page */}
        <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
          <div className="flex gap-1 bg-white rounded-xl p-1 w-fit" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
            {SKU_BRANDS.map((b) => (
              <button
                key={b}
                onClick={() => handleSkuBrand(b)}
                className="px-3.5 py-1.5 rounded-lg text-sm font-semibold transition-all"
                style={{
                  background: skuBrand === b ? NAVY : "transparent",
                  color: skuBrand === b ? "white" : "#475569",
                }}
              >
                {b}
              </button>
            ))}
          </div>
          <div className="flex items-center gap-2">
            {availableMonths.length > 0 && (
              <div className="flex items-center gap-1.5 bg-white rounded-xl px-3 py-2" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
                <Calendar size={14} style={{ color: NAVY }} />
                <select
                  value={selectedMonth || availableMonths[0]}
                  onChange={(e) => setSelectedMonth(e.target.value)}
                  className="text-sm font-semibold text-slate-700 bg-transparent outline-none"
                >
                  {availableMonths.map((m) => (
                    <option key={m} value={m}>{monthLabel(m)}</option>
                  ))}
                </select>
              </div>
            )}
            <button
              onClick={() => exportSkuComparisonToExcel(skuComparison, skuBrand, selectedMonth || (availableMonths[0] || ""))}
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-sm font-semibold transition-all"
              style={{ background: NAVY, color: "white" }}
            >
              <Download size={14} /> Export Excel
            </button>
          </div>
        </div>



        {/* SKU trend */}
        <div className="bg-white rounded-2xl p-5" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
          <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
            <h2 className="text-sm font-bold" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>SKU Deep-Dive (working-day cumulative) — {sku}</h2>
            <div className="flex gap-1 flex-wrap justify-end">
              {(SKU_LIST[skuBrand] || []).map((s) => (
                <button
                  key={s}
                  onClick={() => setSku(s)}
                  className="px-2.5 py-1 rounded-lg text-xs font-semibold transition-colors"
                  style={{
                    background: sku === s ? RED : "#F1F5F9",
                    color: sku === s ? "white" : "#475569",
                  }}
                >
                  {s}
                </button>
              ))}
            </div>
          </div>

          {multiMetric ? (
            <div className="mb-1">
              <p className="text-xs text-slate-400 mb-3">
                Same working day comparison — Day {multiMetric.day} (Running Month vs Last Month vs Last Year, up to this day)
              </p>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <MetricSnapshotCard label="Invoice" icon={Receipt} color={NAVY} data={multiMetric.invoice} unit="" />
                <MetricSnapshotCard label="PKT (pc)" icon={Package} color={GOLD} data={multiMetric.pkt} unit="" />
                <MetricSnapshotCard label="Qty (KG)" icon={Scale} color={GREEN} data={multiMetric.qty} unit=" kg" decimals={1} />
                <MetricSnapshotCard label="Amount (Tk)" icon={Banknote} color={RED} data={multiMetric.amountTk} unit=" tk" />
              </div>
            </div>
          ) : (
            <div className="flex gap-3 mb-1 flex-wrap">
              <div className="flex items-center gap-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-1.5 font-medium">
                vs Last Month <GrowthBadge pct={vsLastMonth} />
              </div>
              <div className="flex items-center gap-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-1.5 font-medium">
                vs Last Year <GrowthBadge pct={vsLastYear} />
              </div>
            </div>
          )}
        </div>

        {TREND_METRIC_CONFIG.map(({ key, label, color, decimals, tickDivide }) => {
          const data = trendMulti[key];
          // Compare at the same working day reached so far this month (matches the snapshot cards above),
          // not the last row of the array — otherwise a partial running month gets compared to a full last month.
          const dayIdx = multiMetric ? Math.min(multiMetric.day, data.length) - 1 : data.length - 1;
          const lastRow = data[dayIdx];
          const vsM = lastRow["Last Month"] ? ((lastRow["Running Month"] - lastRow["Last Month"]) / lastRow["Last Month"]) * 100 : 0;
          const vsY = lastRow["Last Year"] ? ((lastRow["Running Month"] - lastRow["Last Year"]) / lastRow["Last Year"]) * 100 : 0;
          const gradId = `gradTrend-${key}`;
          return (
            <div key={key} className="bg-white rounded-2xl p-5 mt-4" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
              <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                <h2 className="text-sm font-bold" style={{ color: NAVY, fontFamily: "'Sora', sans-serif" }}>{label} — {sku} (working-day cumulative)</h2>
                <div className="flex gap-3 flex-wrap">
                  <div className="flex items-center gap-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-1.5 font-medium">
                    vs Last Month <GrowthBadge pct={vsM} />
                  </div>
                  <div className="flex items-center gap-2 text-xs text-slate-500 bg-slate-50 rounded-lg px-3 py-1.5 font-medium">
                    vs Last Year <GrowthBadge pct={vsY} />
                  </div>
                </div>
              </div>

              <ResponsiveContainer width="100%" height={220}>
                <AreaChart data={data} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
                  <defs>
                    <linearGradient id={gradId} x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor={color} stopOpacity={0.35} />
                      <stop offset="100%" stopColor={color} stopOpacity={0} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#EEF1F5" vertical={false} />
                  <XAxis dataKey="day" tick={{ fontSize: 10, fill: "#94A3B8" }} axisLine={false} tickLine={false} />
                  <YAxis
                    tick={{ fontSize: 11, fill: "#94A3B8" }}
                    axisLine={false}
                    tickLine={false}
                    tickFormatter={(v) => (tickDivide ? `${Math.round(v / 1000)}k` : Math.round(v).toLocaleString())}
                  />
                  <Tooltip formatter={(v) => (decimals ? v.toFixed(decimals) : fmt(v))} contentStyle={{ borderRadius: 10, border: "1px solid #E2E8F0", fontSize: 12, boxShadow: "0 8px 24px rgba(10,38,71,0.12)" }} />
                  <Area type="monotone" dataKey="Running Month" stroke={color} strokeWidth={2.5} fill={`url(#${gradId})`} dot={false} />
                  <Area type="monotone" dataKey="Last Month" stroke={RED} strokeWidth={1.75} strokeDasharray="5 4" fill="transparent" dot={false} />
                  <Area type="monotone" dataKey="Last Year" stroke="#94A3B8" strokeWidth={1.75} strokeDasharray="2 3" fill="transparent" dot={false} />
                </AreaChart>
              </ResponsiveContainer>
              <div className="flex gap-4 mt-2 text-xs text-slate-500">
                <span className="flex items-center gap-1.5"><span className="w-3 h-0.5 inline-block rounded-full" style={{ background: color }} /> Running Month</span>
                <span className="flex items-center gap-1.5"><span className="w-3 h-0.5 inline-block rounded-full" style={{ background: RED }} /> Last Month</span>
                <span className="flex items-center gap-1.5"><span className="w-3 h-0.5 inline-block rounded-full" style={{ background: "#94A3B8" }} /> Last Year</span>
              </div>
            </div>
          );
        })}

        {/* District & Team degrowth for this exact SKU — bottom of page */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-4">
          <DegrowthList title={`District in Degrowth — ${sku}`} data={districtPerf} />
          <DegrowthList title={`Team in Degrowth — ${sku}`} data={teamPerf} />
        </div>
      </>
      )}

        <p className="text-xs text-slate-400 mt-4 text-center">
          {dataStatus === "live"
            ? "Connected to your Google Sheet — update the sheet daily and this refreshes on reload."
            : "Showing demo data — deploy this to Vercel (browser fetch is blocked in this preview) to pull live from your Google Sheet."}
        </p>
      </div>
    </div>
  );
}
