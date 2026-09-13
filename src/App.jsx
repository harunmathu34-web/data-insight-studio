import React, { useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, BarChart, Bar, LineChart, Line, PieChart, Pie,
  Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ScatterChart,
  Scatter, ZAxis,
} from "recharts";
import {
  Upload, Sparkles, Database, TrendingUp, Target, Filter, FileSpreadsheet,
  RefreshCcw, Download, AlertTriangle, CheckCircle2, Layers3, BarChart3,
  PieChart as PieIcon, LineChart as LineIcon, Activity, Search,
} from "lucide-react";

const COLORS = ["#2563EB", "#7C3AED", "#059669", "#EA580C", "#DB2777", "#0891B2", "#CA8A04", "#4F46E5"];
const NAVY = "#0B1324";

const emptyDataset = { name: "", rows: [], columns: [] };

function cleanRows(rows) {
  return (rows || []).filter((r) => Object.values(r || {}).some((v) => v !== null && v !== undefined && v !== ""));
}

function numeric(v) {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v !== "string") return null;
  const cleaned = v.replace(/,/g, "").replace(/%$/, "").trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function dateValue(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v !== "string" && typeof v !== "number") return null;
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? null : d;
}

function inferColumns(rows) {
  if (!rows.length) return [];
  return Object.keys(rows[0]).map((name) => {
    const values = rows.map((r) => r[name]).filter((v) => v !== null && v !== undefined && v !== "");
    const nums = values.filter((v) => numeric(v) !== null).length;
    const dates = values.filter((v) => dateValue(v) !== null && typeof v !== "number").length;
    const unique = new Set(values.map(String)).size;
    let type = "text";
    if (values.length && nums / values.length >= 0.85) type = "numeric";
    else if (values.length && dates / values.length >= 0.85) type = "date";
    else if (values.length && unique / values.length < 0.65) type = "category";
    return { name, type, unique, missing: rows.length - values.length };
  });
}

function formatNumber(n) {
  if (!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("en", { maximumFractionDigits: 1 }).format(n);
}

function compact(n) {
  if (!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat("en", { notation: "compact", maximumFractionDigits: 1 }).format(n);
}

function aggregate(rows, dimension, measure, limit = 10) {
  const map = new Map();
  rows.forEach((r) => {
    const key = r[dimension] === null || r[dimension] === undefined || r[dimension] === "" ? "(blank)" : String(r[dimension]);
    const value = measure ? numeric(r[measure]) : 1;
    if (value === null && measure) return;
    map.set(key, (map.get(key) || 0) + (value ?? 1));
  });
  let data = [...map.entries()].map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }));
  data.sort((a, b) => b.value - a.value);
  if (data.length > limit) {
    const top = data.slice(0, limit - 1);
    top.push({ name: "Other", value: data.slice(limit - 1).reduce((s, x) => s + x.value, 0) });
    data = top;
  }
  return data;
}

function monthlyTrend(rows, dateCol, measure) {
  const map = new Map();
  rows.forEach((r) => {
    const d = dateValue(r[dateCol]);
    const n = measure ? numeric(r[measure]) : 1;
    if (!d || (measure && n === null)) return;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    map.set(key, (map.get(key) || 0) + (n ?? 1));
  });
  return [...map.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }));
}

function findTarget(columns) {
  return columns.find((c) => /target|goal|planned|expected/i.test(c.name) && c.type === "numeric")?.name || null;
}

function ChartCard({ title, subtitle, icon: Icon, children, wide = false }) {
  return (
    <section className={`${wide ? "lg:col-span-2" : ""} rounded-2xl border border-slate-200 bg-white p-5 shadow-sm`}>
      <div className="mb-4 flex items-start justify-between gap-4">
        <div className="flex items-start gap-3">
          <div className="rounded-xl bg-slate-100 p-2 text-slate-700"><Icon size={17} /></div>
          <div><h3 className="font-semibold text-slate-900">{title}</h3>{subtitle && <p className="mt-0.5 text-xs text-slate-500">{subtitle}</p>}</div>
        </div>
      </div>
      <div className="h-[300px] w-full">{children}</div>
    </section>
  );
}

function Tip() {
  return <Tooltip contentStyle={{ borderRadius: 12, border: "1px solid #E2E8F0", boxShadow: "0 8px 30px rgba(15,23,42,.10)" }} />;
}

function Dashboard({ dataset, onReplace }) {
  const { rows, columns, name } = dataset;
  const numericCols = columns.filter((c) => c.type === "numeric");
  const dimensionCols = columns.filter((c) => c.type === "category" || c.type === "text");
  const dateCols = columns.filter((c) => c.type === "date");
  const [filters, setFilters] = useState({});

  const filterColumns = dimensionCols.slice(0, 4);
  const filteredRows = useMemo(() => rows.filter((r) => filterColumns.every((c) => !filters[c.name] || String(r[c.name]) === filters[c.name])), [rows, filters, filterColumns]);
  const measure = numericCols[0]?.name || null;
  const dimension = dimensionCols[0]?.name || columns[0]?.name;
  const secondDimension = dimensionCols[1]?.name || dimension;
  const date = dateCols[0]?.name || null;
  const target = findTarget(columns);

  const stats = useMemo(() => numericCols.slice(0, 4).map((c) => {
    const vals = filteredRows.map((r) => numeric(r[c.name])).filter((v) => v !== null);
    const sum = vals.reduce((a, b) => a + b, 0);
    return { name: c.name, sum, avg: vals.length ? sum / vals.length : 0, max: vals.length ? Math.max(...vals) : 0 };
  }), [filteredRows, numericCols]);

  const quality = useMemo(() => {
    const cells = rows.length * columns.length || 1;
    const missing = columns.reduce((s, c) => s + c.missing, 0);
    const seen = new Set(); let duplicates = 0;
    rows.forEach((r) => { const key = JSON.stringify(r); if (seen.has(key)) duplicates++; seen.add(key); });
    return { missing, missingPct: (missing / cells) * 100, duplicates };
  }, [rows, columns]);

  const charts = useMemo(() => {
    const result = [];
    if (dimension && (measure || rows.length)) result.push({ kind: "bar", title: measure ? `${measure} by ${dimension}` : `Records by ${dimension}`, data: aggregate(filteredRows, dimension, measure) });
    if (date && (measure || rows.length)) result.push({ kind: "line", title: date && measure ? `${measure} over time` : "Records over time", data: monthlyTrend(filteredRows, date, measure) });
    if (secondDimension) result.push({ kind: "pie", title: `${secondDimension} composition`, data: aggregate(filteredRows, secondDimension, null, 7) });
    if (dimension && measure && numericCols[1]) result.push({ kind: "scatter", title: `${measure} vs ${numericCols[1].name}`, x: measure, y: numericCols[1].name, data: filteredRows.map((r) => ({ x: numeric(r[measure]), y: numeric(r[numericCols[1].name]) })).filter((p) => p.x !== null && p.y !== null).slice(0, 120) });
    return result;
  }, [filteredRows, dimension, secondDimension, date, measure, numericCols, rows.length]);

  const best = charts[0]?.data?.[0];
  const total = measure ? (stats[0]?.sum || 0) : filteredRows.length;
  const targetActual = target && measure ? filteredRows.reduce((s, r) => s + (numeric(r[measure]) || 0), 0) : null;
  const targetTotal = target ? filteredRows.reduce((s, r) => s + (numeric(r[target]) || 0), 0) : null;
  const achievement = targetTotal ? (targetActual / targetTotal) * 100 : null;

  const insightText = useMemo(() => {
    const out = [];
    if (best) out.push(`${best.name} is the leading ${dimension} with ${formatNumber(best.value)} ${measure ? `in ${measure}` : "records"}.`);
    if (achievement !== null) out.push(`${measure} is at ${achievement.toFixed(1)}% of the ${target} target.`);
    if (quality.missingPct > 5) out.push(`${quality.missingPct.toFixed(1)}% of cells are missing and should be reviewed before reporting.`);
    if (quality.duplicates) out.push(`${formatNumber(quality.duplicates)} duplicate rows were detected.`);
    if (date && charts[1]?.data?.length >= 2) {
      const a = charts[1].data[0].value, b = charts[1].data.at(-1).value;
      if (a !== 0) out.push(`The latest period is ${Math.abs(((b - a) / a) * 100).toFixed(1)}% ${b >= a ? "above" : "below"} the first period.`);
    }
    return out.slice(0, 5);
  }, [best, dimension, measure, achievement, target, quality, date, charts]);

  function exportReport() {
    const report = { dataset: name, rows: filteredRows.length, columns: columns.map((c) => c.name), filters, kpis: stats, insights: insightText, generatedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(report, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = "data-insight-report.json"; a.click(); URL.revokeObjectURL(url);
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <header className="sticky top-0 z-20 border-b border-slate-800 bg-[#0B1324] text-white shadow-lg">
        <div className="mx-auto flex max-w-[1500px] items-center justify-between gap-4 px-6 py-4">
          <div className="flex items-center gap-3"><div className="rounded-xl bg-blue-600 p-2"><Sparkles size={19} /></div><div><div className="font-bold tracking-tight">Data Insight Studio</div><div className="text-xs text-slate-400">M&E Intelligence • Automated Analytics</div></div></div>
          <div className="flex items-center gap-2"><button onClick={exportReport} className="hidden items-center gap-2 rounded-lg border border-slate-600 px-3 py-2 text-sm hover:bg-white/10 sm:flex"><Download size={15}/> Export report</button><button onClick={onReplace} className="flex items-center gap-2 rounded-lg bg-white px-3 py-2 text-sm font-medium text-slate-900 hover:bg-slate-100"><RefreshCcw size={15}/> Replace data</button></div>
        </div>
      </header>

      <main className="mx-auto max-w-[1500px] space-y-6 px-6 py-7">
        <div className="flex flex-col justify-between gap-3 lg:flex-row lg:items-end"><div><div className="mb-1 flex items-center gap-2 text-xs font-medium uppercase tracking-wider text-blue-600"><Activity size={13}/> Executive dashboard</div><h1 className="text-2xl font-bold tracking-tight">{name}</h1><p className="mt-1 text-sm text-slate-500">{filteredRows.length.toLocaleString()} of {rows.length.toLocaleString()} records • {columns.length} fields analyzed automatically</p></div><div className="flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-600"><Database size={14}/> Live filtered analysis</div></div>

        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          <Kpi icon={Database} label="Records" value={filteredRows.length.toLocaleString()} sub={`${columns.length} fields`} />
          <Kpi icon={TrendingUp} label={measure || "Primary metric"} value={measure ? compact(total) : filteredRows.length.toLocaleString()} sub={measure ? `average ${compact(stats[0]?.avg || 0)}` : "row count"} />
          <Kpi icon={BarChart3} label="Top performer" value={best?.name || "—"} sub={best ? formatNumber(best.value) : "No categorical breakdown"} />
          <Kpi icon={Target} label="Target achievement" value={achievement === null ? "—" : `${achievement.toFixed(1)}%`} sub={target ? `${measure} vs ${target}` : "No target field detected"} />
        </div>

        <section className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm"><div className="mb-3 flex items-center gap-2 text-sm font-semibold"><Filter size={16} className="text-blue-600"/> Filters <span className="ml-auto text-xs font-normal text-slate-400">{Object.values(filters).filter(Boolean).length} active</span></div><div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">{filterColumns.map((c) => <label key={c.name} className="text-xs font-medium text-slate-600">{c.name}<select value={filters[c.name] || ""} onChange={(e) => setFilters((p) => ({ ...p, [c.name]: e.target.value }))} className="mt-1 w-full rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-sm font-normal text-slate-800 outline-none focus:border-blue-500"><option value="">All values</option>{[...new Set(filteredRows.map((r) => String(r[c.name] ?? "")).filter(Boolean))].slice(0, 100).map((v) => <option key={v} value={v}>{v}</option>)}</select></label>)}</div></section>

        <div className="grid gap-5 lg:grid-cols-2">{charts.map((chart, i) => <ChartCard key={i} title={chart.title} icon={chart.kind === "line" ? LineIcon : chart.kind === "pie" ? PieIcon : chart.kind === "scatter" ? Search : BarChart3} subtitle="Automatically selected from the structure and content of your data" wide={chart.kind === "line" && charts.length < 4}>{chart.kind === "bar" && <ResponsiveContainer><BarChart data={chart.data} margin={{ top: 8, right: 10, left: 0, bottom: 40 }}><CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" vertical={false}/><XAxis dataKey="name" angle={-25} textAnchor="end" height={60} tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tip/><Bar dataKey="value" radius={[6,6,0,0]}>{chart.data.map((_, j) => <Cell key={j} fill={COLORS[j % COLORS.length]}/>)}</Bar></BarChart></ResponsiveContainer>}{chart.kind === "line" && <ResponsiveContainer><LineChart data={chart.data} margin={{top:10,right:10,left:0,bottom:10}}><CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" vertical={false}/><XAxis dataKey="name" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tip/><Line type="monotone" dataKey="value" stroke="#2563EB" strokeWidth={3} dot={{r:3,fill:"#2563EB"}} activeDot={{r:6}}/></LineChart></ResponsiveContainer>}{chart.kind === "pie" && <ResponsiveContainer><PieChart><Pie data={chart.data} dataKey="value" nameKey="name" innerRadius={65} outerRadius={105} paddingAngle={3}>{chart.data.map((_, j) => <Cell key={j} fill={COLORS[j % COLORS.length]}/>)}</Pie><Tip/><Legend wrapperStyle={{fontSize:11}}/></PieChart></ResponsiveContainer>}{chart.kind === "scatter" && <ResponsiveContainer><ScatterChart margin={{top:10,right:10,bottom:10,left:0}}><CartesianGrid strokeDasharray="3 3"/><XAxis type="number" dataKey="x" name={chart.x} tick={{fontSize:11}}/><YAxis type="number" dataKey="y" name={chart.y} tick={{fontSize:11}}/><ZAxis range={[40,40]}/><Tip/><Scatter data={chart.data} fill="#7C3AED"/></ScatterChart></ResponsiveContainer>}</ChartCard>)}</div>

        <div className="grid gap-5 lg:grid-cols-3"><section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm lg:col-span-2"><div className="mb-4 flex items-center gap-2"><Sparkles size={17} className="text-blue-600"/><h2 className="font-semibold">Automated insights</h2></div>{insightText.length ? <div className="grid gap-3 sm:grid-cols-2">{insightText.map((x, i) => <div key={i} className="rounded-xl bg-slate-50 p-4 text-sm leading-6 text-slate-700"><span className="mr-2 inline-flex h-6 w-6 items-center justify-center rounded-full bg-blue-100 text-xs font-bold text-blue-700">{i+1}</span>{x}</div>)}</div> : <p className="text-sm text-slate-500">Upload richer data with measurable or categorical fields to generate more insights.</p>}</section><section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="mb-4 flex items-center gap-2"><AlertTriangle size={17} className="text-amber-500"/><h2 className="font-semibold">Data quality</h2></div><div className="space-y-3 text-sm"><Quality label="Missing cells" value={`${quality.missingPct.toFixed(1)}%`} bad={quality.missingPct > 5}/><Quality label="Duplicate rows" value={quality.duplicates.toLocaleString()} bad={quality.duplicates > 0}/><Quality label="Fields detected" value={columns.length.toString()} /><Quality label="Numeric fields" value={numericCols.length.toString()} /></div></section></div>

        <section className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="mb-4 flex items-center gap-2"><Layers3 size={17} className="text-slate-700"/><h2 className="font-semibold">Dataset structure</h2></div><div className="overflow-x-auto"><table className="w-full text-left text-sm"><thead><tr className="border-b border-slate-200 text-xs uppercase tracking-wide text-slate-400"><th className="px-3 py-3">Field</th><th className="px-3 py-3">Type</th><th className="px-3 py-3">Unique</th><th className="px-3 py-3">Missing</th></tr></thead><tbody>{columns.map((c) => <tr key={c.name} className="border-b border-slate-100 last:border-0"><td className="px-3 py-3 font-medium text-slate-800">{c.name}</td><td className="px-3 py-3"><span className="rounded-full bg-slate-100 px-2 py-1 text-xs">{c.type}</span></td><td className="px-3 py-3 text-slate-500">{c.unique.toLocaleString()}</td><td className="px-3 py-3 text-slate-500">{c.missing.toLocaleString()}</td></tr>)}</tbody></table></div></section>
      </main>
    </div>
  );
}

function Kpi({ icon: Icon, label, value, sub }) {
  return <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="mb-4 flex h-9 w-9 items-center justify-center rounded-xl bg-blue-50 text-blue-600"><Icon size={17}/></div><div className="text-xs font-medium uppercase tracking-wide text-slate-400">{label}</div><div className="mt-1 truncate text-2xl font-bold text-slate-900">{value}</div><div className="mt-1 truncate text-xs text-slate-500">{sub}</div></div>;
}

function Quality({ label, value, bad }) {
  return <div className="flex items-center justify-between rounded-xl bg-slate-50 px-3 py-2.5"><span className="text-slate-600">{label}</span><span className={`flex items-center gap-1 font-semibold ${bad ? "text-amber-600" : "text-emerald-600"}`}>{bad ? <AlertTriangle size={13}/> : <CheckCircle2 size={13}/>} {value}</span></div>;
}

function UploadScreen({ onFile, loading, error }) {
  const input = useRef(null);
  return <div className="min-h-screen bg-slate-50"><header className="bg-[#0B1324] px-6 py-5 text-white"><div className="mx-auto flex max-w-6xl items-center gap-3"><div className="rounded-xl bg-blue-600 p-2"><Sparkles size={19}/></div><div><div className="font-bold">Data Insight Studio</div><div className="text-xs text-slate-400">M&E Intelligence • Any data. Instant analysis.</div></div></div></header><main className="mx-auto flex min-h-[calc(100vh-76px)] max-w-5xl items-center justify-center px-6 py-12"><div className="w-full max-w-2xl text-center"><div className="mx-auto mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-600 text-white shadow-xl shadow-blue-200"><BarChart3 size={30}/></div><h1 className="text-4xl font-bold tracking-tight text-slate-900">Turn any dataset into a dashboard.</h1><p className="mx-auto mt-4 max-w-xl text-slate-500">Upload Excel or CSV data and Data Insight Studio will detect fields, build visualizations, calculate performance metrics, identify quality issues and surface the most important patterns.</p><button onClick={() => input.current?.click()} disabled={loading} className="mt-8 inline-flex items-center gap-2 rounded-xl bg-blue-600 px-6 py-3 font-semibold text-white shadow-lg shadow-blue-200 transition hover:bg-blue-700 disabled:opacity-60"><Upload size={18}/>{loading ? "Analyzing file…" : "Upload dataset"}</button><input ref={input} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(e) => onFile(e.target.files?.[0])}/><div className="mt-10 grid grid-cols-1 gap-3 text-left sm:grid-cols-3"><Feature icon={Database} title="Understands data" text="Detects numeric, date, categorical and text fields."/><Feature icon={BarChart3} title="Builds visuals" text="Selects useful charts from the dataset structure."/><Feature icon={Sparkles} title="Finds insights" text="Highlights leaders, trends, targets and quality issues."/></div>{error && <div className="mt-5 rounded-xl border border-red-200 bg-red-50 p-3 text-sm text-red-700">{error}</div>}</div></main></div>;
}

function Feature({ icon: Icon, title, text }) { return <div className="rounded-xl border border-slate-200 bg-white p-4"><Icon size={18} className="text-blue-600"/><div className="mt-3 font-semibold text-slate-800">{title}</div><div className="mt-1 text-xs leading-5 text-slate-500">{text}</div></div>; }

export default function App() {
  const [dataset, setDataset] = useState(emptyDataset);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleFile(file) {
    if (!file) return;
    setLoading(true); setError("");
    try {
      let sheets = {};
      if (file.name.toLowerCase().endsWith(".csv")) {
        const parsed = await new Promise((resolve, reject) => Papa.parse(file, { header: true, dynamicTyping: true, skipEmptyLines: true, complete: (r) => resolve(r.data), error: reject }));
        sheets["CSV"] = cleanRows(parsed);
      } else {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array", cellDates: true });
        wb.SheetNames.forEach((sheet) => { sheets[sheet] = cleanRows(XLSX.utils.sheet_to_json(wb.Sheets[sheet], { defval: null })); });
      }
      const usable = Object.entries(sheets).filter(([, r]) => r.length);
      if (!usable.length) throw new Error("No usable rows were found in this file.");
      const [sheetName, rows] = usable[0];
      const columns = inferColumns(rows);
      setDataset({ name: usable.length > 1 ? `${file.name} • ${sheetName}` : file.name, rows, columns });
    } catch (e) { setError(e.message || "Could not read this file."); }
    finally { setLoading(false); }
  }

  if (dataset.rows.length) return <Dashboard dataset={dataset} onReplace={() => setDataset(emptyDataset)} />;
  return <UploadScreen onFile={handleFile} loading={loading} error={error} />;
}
