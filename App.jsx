import React, { useState, useMemo, useEffect } from "react";
import Papa from "papaparse";
import { ChevronRight, ChevronDown, Target, Percent, Wifi, WifiOff, Loader2, Download, Calendar, Package, CheckCircle2, AlertTriangle } from "lucide-react";

// ---------- Google Sheet CSV source ----------
const TARGET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRbKm9XW2L3mvPuUwTXCcLLt5nN3MFO0IciJ3ta5waPjerG0A459RtjwcDBinBgJeJxZpQsZBz9w8kZ/pub?output=csv";

// Fill this in with the direct download/view link of the Excel file you upload online
// (e.g. a Google Drive "anyone with link" share link, or a direct .xlsx URL). The
// "Export Excel" button just opens this link — nothing is generated in-app anymore.
const EXCEL_EXPORT_URL = "";

// ---------- Mock data (fallback shown before the live sheet connects) ----------
const BRANDS = ["AJI-Retail", "AJI-Bulk", "Hapima", "TasteMate"];

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
      perBrand[b] = { target, achv, lastMonth, ldSales: 0 };
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

// ---------- Live Google Sheet data layer ----------
function normalizeKey(k) {
  return String(k).toLowerCase().replace(/[^a-z0-9]/g, "");
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

// Target_vs_Progress sheet (wide format): row1 = brand group headers (merged), row2 = RF/Result/Prog./LD Sales,
// data rows = Section, Unit, Area, then RF/Result/Prog./LD Sales quadruplets per brand. Section/Unit cells are
// merged vertically in the sheet, so blanks are forward-filled from the row above. Optional standalone columns
// (found anywhere in the header row): "Month", "Total Working Days", "Working Days Passed", "Update Till Date".
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
  let updateTillColIndex = row2.findIndex((h) => {
    const n = normalizeKey(h || "");
    return n.includes("update") && n.includes("till");
  });
  if (updateTillColIndex === -1) {
    updateTillColIndex = row1.findIndex((h) => {
      const n = normalizeKey(h || "");
      return n.includes("update") && n.includes("till");
    });
  }
  const extraCols = [monthColIndex, totalWDColIndex, passedWDColIndex, updateTillColIndex].filter((i) => i !== -1);

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
    const metric = normalizeKey(row2[i] || "");
    return { brand: brandLabel, metric };
  });

  const byMonth = {}; // month -> sectionMap
  const paceByMonth = {}; // month -> { total, passed, updateTill }
  let curMonth = "";
  let curSection = "";
  let curUnit = "";
  let curTotalWD = 0;
  let curPassedWD = 0;
  let curUpdateTill = "";
  let hasLdSalesCol = false;
  for (let r = 2; r < rawRows.length; r++) {
    const row = rawRows[r];
    if (!row || row.every((c) => !c || !String(c).trim())) continue;
    const monthRaw = hasMonthCol ? (row[monthColIndex] || "").toString().trim() : "";
    const secRaw = (row[0] || "").toString().trim();
    const unitRaw = (row[1] || "").toString().trim();
    const areaRaw = (row[2] || "").toString().trim();
    const totalWDRaw = totalWDColIndex !== -1 ? (row[totalWDColIndex] || "").toString().trim() : "";
    const passedWDRaw = passedWDColIndex !== -1 ? (row[passedWDColIndex] || "").toString().trim() : "";
    const updateTillRaw = updateTillColIndex !== -1 ? (row[updateTillColIndex] || "").toString().trim() : "";
    if (monthRaw) curMonth = monthRaw;
    if (secRaw) curSection = secRaw;
    if (unitRaw) curUnit = unitRaw;
    if (totalWDRaw) curTotalWD = parseInt(totalWDRaw, 10) || 0;
    if (passedWDRaw) curPassedWD = parseInt(passedWDRaw, 10) || 0;
    if (updateTillRaw) curUpdateTill = updateTillRaw;
    if (!areaRaw) continue;
    const month = hasMonthCol ? curMonth || "Unspecified" : "current";
    const section = curSection || "Unassigned";
    const unit = curUnit || "Unassigned";

    if (!paceByMonth[month]) paceByMonth[month] = { total: 0, passed: 0, updateTill: "" };
    if (curTotalWD) paceByMonth[month].total = curTotalWD;
    if (curPassedWD) paceByMonth[month].passed = curPassedWD;
    if (curUpdateTill) paceByMonth[month].updateTill = curUpdateTill;

    byMonth[month] = byMonth[month] || {};
    byMonth[month][section] = byMonth[month][section] || {};
    byMonth[month][section][unit] = byMonth[month][section][unit] || {};
    const cells = {};
    BRANDS.forEach((b) => (cells[b] = { target: 0, achv: 0, lastMonth: 0, ldSales: 0 }));
    byMonth[month][section][unit][areaRaw] = { brands: cells };

    for (let i = baseCols; i < row.length; i++) {
      const meta = colMeta[i];
      if (!meta || !meta.brand) continue;
      const brandKey = BRANDS.find((b) => normalizeKey(b) === normalizeKey(meta.brand));
      if (!brandKey) continue; // e.g. the sheet's own "Total" columns
      const val = parseFloat(String(row[i]).replace(/,/g, "")) || 0;
      if (meta.metric === "rf") cells[brandKey].target = val;
      else if (meta.metric === "result") cells[brandKey].achv = val;
      else if (meta.metric.startsWith("ldsale")) {
        cells[brandKey].ldSales = val;
        hasLdSalesCol = true;
      }
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
          brands: sectionMap[secName][unitName][areaName].brands,
        })),
      })),
    }));

  const byMonthArrays = {};
  monthKeys.forEach((m) => (byMonthArrays[m] = toSectionsArray(byMonth[m])));

  if (!hasMonthCol) {
    return {
      hasMonths: false,
      months: [],
      byMonth: {},
      sections: byMonthArrays["current"],
      pace: paceByMonth["current"] || null,
      hasLdSales: hasLdSalesCol,
    };
  }
  const months = monthKeys.sort((a, b) => {
    const da = parseMonthKey(a);
    const db = parseMonthKey(b);
    if (da && db) return db - da;
    return b.localeCompare(a);
  });
  return { hasMonths: true, months, byMonth: byMonthArrays, paceByMonth, sections: byMonthArrays[months[0]], hasLdSales: hasLdSalesCol };
}

// ---------- Helpers ----------
const fmt = (n) => `${new Intl.NumberFormat("en-BD").format(Math.round(n))} kg`;
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
          acc.ldSales += c.ldSales || 0;
        }
        return acc;
      },
      { target: 0, achv: 0, lastMonth: 0, ldSales: 0 }
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
      acc.ldSales += c.ldSales || 0;
      return acc;
    },
    { target: 0, achv: 0, lastMonth: 0, ldSales: 0 }
  );
}
function sumUnits(units, brand) {
  return units.reduce(
    (acc, u) => {
      const s = sumBrand(u.areas, brand);
      acc.target += s.target;
      acc.achv += s.achv;
      acc.lastMonth += s.lastMonth;
      acc.ldSales += s.ldSales || 0;
      return acc;
    },
    { target: 0, achv: 0, lastMonth: 0, ldSales: 0 }
  );
}
function achPct(s) {
  return s.target ? (s.achv / s.target) * 100 : 0;
}

const NAVY = "#0A2647";
const NAVY_LIGHT = "#12395F";
const RED = "#C81D25";
const GREEN = "#1F9254";
const GOLD = "#D9A441";
const CREAM = "#F4F5F8";
const BRAND_COLORS = { "AJI-Retail": NAVY, "AJI-Bulk": GOLD, "Hapima": RED, "TasteMate": GREEN, "Total": "#475569" };

function AchBadge({ pct }) {
  const good = pct >= 100;
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold"
      style={{ background: good ? "rgba(31,146,84,0.12)" : "rgba(200,29,37,0.1)", color: good ? GREEN : RED }}
    >
      {pct.toFixed(1)}%
    </span>
  );
}

function ProgressBar({ pct }) {
  const clamped = Math.min(pct, 130);
  return (
    <div className="w-28 h-1.5 rounded-full bg-slate-200 overflow-hidden">
      <div className="h-full rounded-full" style={{ width: `${Math.min(clamped, 100)}%`, background: pct >= 100 ? GREEN : NAVY_LIGHT }} />
    </div>
  );
}

// ---------- Row components ----------
function AreaRow({ area, brand, showLd }) {
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
      {showLd && <td className="py-2 text-right pr-6 text-slate-500">{fmt(s.ldSales || 0)}</td>}
    </tr>
  );
}

function UnitBlock({ unit, brand, showLd }) {
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
        {showLd && <td className="py-2.5 text-right pr-6 font-medium text-slate-600">{fmt(s.ldSales || 0)}</td>}
      </tr>
      {open && unit.areas.map((a) => <AreaRow key={a.name} area={a} brand={brand} showLd={showLd} />)}
    </>
  );
}

function SectionBlock({ section, brand, showLd }) {
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
        {showLd && <td className="py-3 text-right pr-6 font-semibold text-white/90">{fmt(s.ldSales || 0)}</td>}
      </tr>
      {open && section.units.map((u) => <UnitBlock key={u.name} unit={u} brand={brand} showLd={showLd} />)}
    </>
  );
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

function ProductProgressRow({ allAreas }) {
  return (
    <div className="flex gap-3 mb-6 flex-wrap">
      {BRANDS.map((b) => {
        const s = sumBrand(allAreas, b);
        return <RadialProgress key={b} label={b} pct={achPct(s)} target={s.target} achv={s.achv} color={BRAND_COLORS[b]} />;
      })}
    </div>
  );
}

// ---------- KPI cards ----------
function KpiCard({ label, value, sub, icon: Icon, accent, badge }) {
  return (
    <div className="flex-1 min-w-[150px] rounded-2xl p-4 bg-white" style={{ boxShadow: "0 1px 2px rgba(10,38,71,0.04), 0 8px 24px -14px rgba(10,38,71,0.18)" }}>
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
  const [brand, setBrand] = useState(BRANDS[0]);
  const [targetData, setTargetData] = useState(null);
  const [dataStatus, setDataStatus] = useState("loading"); // loading | live | mock
  const [selectedTargetMonth, setSelectedTargetMonth] = useState(null);

  useEffect(() => {
    const link = document.createElement("link");
    link.href = "https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap";
    link.rel = "stylesheet";
    document.head.appendChild(link);
    return () => document.head.removeChild(link);
  }, []);

  useEffect(() => {
    let cancelled = false;
    fetchCsvRaw(TARGET_CSV_URL)
      .then((targetRows) => {
        if (cancelled) return;
        const sections = buildLiveSectionsWide(targetRows);
        if (sections) setTargetData(sections);
        setDataStatus(sections ? "live" : "mock");
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

  const isCurrentMonthView = !targetData?.hasMonths || !selectedTargetMonth || selectedTargetMonth === targetData.months[0];
  const showLd = false; // LD Sales removed from the dashboard view
  const ldGrandTotal = showLd ? totals.ldSales || 0 : 0;

  const currentPace = useMemo(() => {
    if (!targetData) return null;
    if (targetData.hasMonths) {
      const m = selectedTargetMonth || targetData.months[0];
      return targetData.paceByMonth?.[m] || null;
    }
    return targetData.pace || null;
  }, [targetData, selectedTargetMonth]);
  const workingDayPace = isCurrentMonthView ? currentPace : null;
  const expectedPct = workingDayPace && workingDayPace.total ? (workingDayPace.passed / workingDayPace.total) * 100 : null;
  const actualPct = achPct(totals);
  const onTrack = expectedPct !== null ? actualPct >= expectedPct - 3 : null;

  const lastUpdatedLabel = currentPace && currentPace.updateTill ? currentPace.updateTill : null;

  const handleBrand = (b) => setBrand(b);
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
                {targetData?.hasMonths && selectedTargetMonth && selectedTargetMonth !== targetData.months[0] ? (
                  <span className="text-[11px] text-white/50 font-medium">{selectedTargetMonth}</span>
                ) : (
                  lastUpdatedLabel && <span className="text-[11px] text-white/50 font-medium">Updated till {lastUpdatedLabel}</span>
                )}
              </div>
              <h1 className="text-2xl font-extrabold text-white" style={{ fontFamily: "'Sora', sans-serif" }}>ROLLING FORECAST DASHBOARD</h1>
            </div>
            <div className="flex items-center gap-3 flex-wrap">
              <div className="flex gap-1 bg-white/10 backdrop-blur rounded-xl p-1 border border-white/10">
                {[...BRANDS, TOTAL_LABEL].map((b) => (
                  <button
                    key={b}
                    onClick={() => handleBrand(b)}
                    className="px-3.5 py-1.5 rounded-lg text-sm font-semibold transition-all"
                    style={{ background: brand === b ? "white" : "transparent", color: brand === b ? NAVY : "rgba(255,255,255,0.75)" }}
                  >
                    {b}
                  </button>
                ))}
              </div>
              {targetData?.hasMonths && (
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
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 -mt-5 pb-8">
        {/* Product-wise total progress (4 shapes) */}
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
          {showLd && <KpiCard label="LD Sales" value={fmt(ldGrandTotal)} sub="from sheet, running month" icon={Package} accent={GOLD} />}
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
                {showLd && <th className="text-right py-3 pr-6 font-semibold">LD Sales</th>}
              </tr>
            </thead>
            <tbody>
              {activeSections.map((s) => (
                <SectionBlock key={s.name} section={s} brand={brand} showLd={showLd} />
              ))}
            </tbody>
          </table>
        </div>

        <div className="flex justify-end mb-6">
          {EXCEL_EXPORT_URL ? (
            <a
              href={EXCEL_EXPORT_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-medium text-slate-400 hover:text-slate-600 transition-colors"
            >
              <Download size={11} /> Export Excel
            </a>
          ) : (
            <span className="flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-medium text-slate-300 cursor-not-allowed" title="Set EXCEL_EXPORT_URL in the code first">
              <Download size={11} /> Export Excel
            </span>
          )}
        </div>

        <p className="text-xs text-slate-400 mt-4 text-center">
          {dataStatus === "live"
            ? "Connected to your Google Sheet — update the sheet daily and this refreshes on reload."
            : "Showing demo data — deploy this to Vercel (browser fetch is blocked in this preview) to pull live from your Google Sheet."}
        </p>
      </div>
    </div>
  );
}
