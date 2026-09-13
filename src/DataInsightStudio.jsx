import React, { useState, useMemo, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  Upload, Home, MessageSquare, Send, Sparkles, FileSpreadsheet,
  AlertTriangle, Database, Loader2, Trash2, BarChart3, PieChartIcon,
  LineChartIcon, RefreshCcw, CheckCircle2, ChevronRight,
} from "lucide-react";

// ---------- palette ----------
const ACCENTS = ["#2563EB", "#0D9488", "#7C3AED", "#D97706", "#DB2777", "#059669"];
const NAVY = "#0B1324";

// ---------- helpers: parsing & profiling ----------
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const name = file.name.toLowerCase();
    if (name.endsWith(".csv")) {
      Papa.parse(file, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: (res) => resolve(res.data),
        error: reject,
      });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    }
  });
}

function isDateLike(v) {
  if (v instanceof Date) return true;
  if (typeof v !== "string") return false;
  return /^\d{4}-\d{2}(-\d{2})?$/.test(v) || /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(v);
}

function detectType(values) {
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && v !== "");
  if (nonEmpty.length === 0) return "text";
  const numeric = nonEmpty.filter((v) => typeof v === "number" || (!isNaN(parseFloat(v)) && isFinite(v))).length;
  if (numeric / nonEmpty.length > 0.85) return "numeric";
  const dates = nonEmpty.filter(isDateLike).length;
  if (dates / nonEmpty.length > 0.85) return "date";
  const uniq = new Set(nonEmpty.map(String)).size;
  if (uniq / nonEmpty.length < 0.6) return "categorical";
  return "text";
}

function profile(rows) {
  if (!rows || !rows.length) return { columns: [], numericCols: [], categoricalCols: [], dateCols: [] };
  const names = Object.keys(rows[0]);
  const columns = names.map((name) => {
    const values = rows.map((r) => r[name]);
    const type = detectType(values);
    const missing = values.filter((v) => v === null || v === undefined || v === "").length;
    return { name, type, missing };
  });
  return {
    columns,
    numericCols: columns.filter((c) => c.type === "numeric").map((c) => c.name),
    categoricalCols: columns.filter((c) => c.type === "categorical" || c.type === "text").map((c) => c.name),
    dateCols: columns.filter((c) => c.type === "date").map((c) => c.name),
  };
}

function num(v) {
  const n = typeof v === "number" ? v : parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function aggregate(rows, catCol, numCol, agg = "sum", limit = 8) {
  const buckets = new Map();
  rows.forEach((r) => {
    const key = r[catCol] === null || r[catCol] === undefined || r[catCol] === "" ? "(blank)" : String(r[catCol]);
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(numCol ? num(r[numCol]) : 1);
  });
  let entries = Array.from(buckets.entries()).map(([key, vals]) => {
    let value;
    if (agg === "count") value = vals.length;
    else if (agg === "avg") value = vals.reduce((a, b) => a + b, 0) / vals.length;
    else value = vals.reduce((a, b) => a + b, 0);
    return { name: key, value: Math.round(value * 100) / 100 };
  });
  entries.sort((a, b) => b.value - a.value);
  if (entries.length > limit) {
    const top = entries.slice(0, limit - 1);
    const restSum = entries.slice(limit - 1).reduce((a, e) => a + e.value, 0);
    top.push({ name: "Other", value: Math.round(restSum * 100) / 100 });
    return top;
  }
  return entries;
}

function timeSeries(rows, dateCol, numCol, agg = "sum") {
  const buckets = new Map();
  rows.forEach((r) => {
    let d = r[dateCol];
    if (!d) return;
    const key = String(d).slice(0, 7); // YYYY-MM
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(numCol ? num(r[numCol]) : 1);
  });
  let entries = Array.from(buckets.entries()).map(([key, vals]) => {
    let value = agg === "avg" ? vals.reduce((a, b) => a + b, 0) / vals.length : vals.reduce((a, b) => a + b, 0);
    return { name: key, value: Math.round(value * 100) / 100 };
  });
  entries.sort((a, b) => (a.name > b.name ? 1 : -1));
  return entries;
}

function formatCompact(n) {
  return new Intl.NumberFormat("en", { notation: "compact", maximumFractionDigits: 1 }).format(n);
}

// ---------- AI calls ----------
async function callClaude(messages, system) {
  // This hits OUR OWN /api/chat function (see api/chat.js), which holds the
  // real Anthropic API key on the server and forwards the request. The browser
  // never sees the key.
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ system, messages }),
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error?.message || data.error || "AI request failed");
  }
  const text = (data.content || [])
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n");
  return text;
}

function buildDataSystemPrompt(rows, prof) {
  const sample = rows.slice(0, 12);
  const stats = {};
  prof.numericCols.forEach((c) => {
    const vals = rows.map((r) => num(r[c]));
    stats[c] = {
      sum: Math.round(vals.reduce((a, b) => a + b, 0) * 100) / 100,
      avg: Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100) / 100,
      min: Math.min(...vals),
      max: Math.max(...vals),
    };
  });
  return `You are a data analyst assistant embedded in a dashboard tool. You have access to a dataset with ${rows.length} rows.
Columns: ${JSON.stringify(prof.columns.map((c) => ({ name: c.name, type: c.type })))}
Numeric column stats: ${JSON.stringify(stats)}
Sample rows (first 12): ${JSON.stringify(sample)}

Answer questions about this data clearly and concisely, in plain language, like you're explaining to someone with no data-analysis background. Use the sample rows and stats given, don't invent numbers you can't support.

If — and only if — the person asks you to build, make, show, plot or add a chart, end your reply with a fenced code block labeled "chart" containing ONLY valid JSON, no comments, in this exact shape:
\`\`\`chart
{"type": "bar" | "line" | "pie", "x": "<a column name from the list above>", "y": "<a numeric column name, or null for a count>", "agg": "sum" | "avg" | "count", "title": "<short chart title>"}
\`\`\`
Write a short one-sentence answer before the code block explaining what the chart shows. Never include a chart block unless asked for a chart.`;
}

function extractChartSpec(text) {
  const match = text.match(/```chart\s*([\s\S]*?)```/);
  if (!match) return { cleanText: text, spec: null };
  const cleanText = text.replace(match[0], "").trim();
  try {
    const spec = JSON.parse(match[1].trim());
    return { cleanText, spec };
  } catch {
    return { cleanText, spec: null };
  }
}

// ---------- small UI atoms ----------
function KpiCard({ icon: Icon, color, label, value, sub }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 flex flex-col gap-3">
      <div className="flex items-center justify-between">
        <div
          className="w-10 h-10 rounded-full flex items-center justify-center"
          style={{ backgroundColor: color + "1A" }}
        >
          <Icon size={18} style={{ color }} />
        </div>
      </div>
      <div>
        <div className="text-sm text-slate-500">{label}</div>
        <div className="text-2xl font-semibold text-slate-900 mt-0.5">{value}</div>
        {sub && <div className="text-xs text-slate-400 mt-1">{sub}</div>}
      </div>
    </div>
  );
}

function ChartCard({ title, icon: Icon, children, onRemove }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <Icon size={16} className="text-slate-400" />
          <h3 className="text-sm font-semibold text-slate-800">{title}</h3>
        </div>
        {onRemove && (
          <button onClick={onRemove} className="text-slate-300 hover:text-slate-500">
            <Trash2 size={14} />
          </button>
        )}
      </div>
      <div style={{ width: "100%", height: 240 }}>{children}</div>
    </div>
  );
}

function RenderChart({ type, data }) {
  if (type === "pie") {
    return (
      <ResponsiveContainer>
        <PieChart>
          <Pie data={data} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
            {data.map((_, i) => (
              <Cell key={i} fill={ACCENTS[i % ACCENTS.length]} />
            ))}
          </Pie>
          <Tooltip />
          <Legend wrapperStyle={{ fontSize: 11 }} />
        </PieChart>
      </ResponsiveContainer>
    );
  }
  if (type === "line") {
    return (
      <ResponsiveContainer>
        <LineChart data={data}>
          <CartesianGrid strokeDasharray="3 3" stroke="#EEF2F7" />
          <XAxis dataKey="name" tick={{ fontSize: 11 }} />
          <YAxis tick={{ fontSize: 11 }} />
          <Tooltip />
          <Line type="monotone" dataKey="value" stroke={ACCENTS[0]} strokeWidth={2} dot={{ r: 3 }} />
        </LineChart>
      </ResponsiveContainer>
    );
  }
  return (
    <ResponsiveContainer>
      <BarChart data={data}>
        <CartesianGrid strokeDasharray="3 3" stroke="#EEF2F7" />
        <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-20} textAnchor="end" height={50} />
        <YAxis tick={{ fontSize: 11 }} />
        <Tooltip />
        <Bar dataKey="value" radius={[4, 4, 0, 0]}>
          {data.map((_, i) => (
            <Cell key={i} fill={ACCENTS[i % ACCENTS.length]} />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

// ---------- main app ----------
export default function DataInsightStudio() {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [view, setView] = useState("upload"); // upload | dashboard | assistant
  const [error, setError] = useState("");
  const [loadingFile, setLoadingFile] = useState(false);
  const [insights, setInsights] = useState([]);
  const [insightsLoading, setInsightsLoading] = useState(false);
  const [customCharts, setCustomCharts] = useState([]);
  const [chatMessages, setChatMessages] = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const fileInputRef = useRef(null);
  const chatEndRef = useRef(null);

  const prof = useMemo(() => (rows ? profile(rows) : null), [rows]);

  const duplicateCount = useMemo(() => {
    if (!rows) return 0;
    const seen = new Set();
    let dupes = 0;
    rows.forEach((r) => {
      const key = JSON.stringify(r);
      if (seen.has(key)) dupes += 1;
      seen.add(key);
    });
    return dupes;
  }, [rows]);

  const dataQualityIssues = useMemo(() => {
    if (!prof) return [];
    const issues = [];
    prof.columns.forEach((c) => {
      if (c.missing > 0) {
        issues.push({ label: `Missing values in "${c.name}"`, count: c.missing, tone: "amber" });
      }
    });
    if (duplicateCount > 0) issues.push({ label: "Duplicate rows detected", count: duplicateCount, tone: "red" });
    return issues.slice(0, 5);
  }, [prof, duplicateCount]);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setError("");
    setLoadingFile(true);
    try {
      const data = await parseFile(file);
      const clean = data.filter((r) => Object.values(r).some((v) => v !== null && v !== undefined && v !== ""));
      if (!clean.length) throw new Error("No usable rows found in that file.");
      setRows(clean);
      setFileName(file.name);
      setCustomCharts([]);
      setChatMessages([]);
      setInsights([]);
      setView("dashboard");
    } catch (e) {
      setError(e.message || "Could not read that file. Try a .csv or .xlsx export.");
    } finally {
      setLoadingFile(false);
    }
  }, []);

  const generateInsights = useCallback(async () => {
    if (!rows || !prof) return;
    setInsightsLoading(true);
    try {
      const system = buildDataSystemPrompt(rows, prof);
      const text = await callClaude(
        [{ role: "user", content: "Give me exactly 4 short bullet-point insights about this dataset — trends, standouts, or anything worth a person's attention. One line each, no numbering, no markdown symbols. No chart block for this one." }],
        system
      );
      const bullets = text.split("\n").map((l) => l.replace(/^[-•*]\s*/, "").trim()).filter(Boolean).slice(0, 4);
      setInsights(bullets);
    } catch {
      setInsights(["Couldn't reach the AI assistant just now — you can still browse the charts below."]);
    } finally {
      setInsightsLoading(false);
    }
  }, [rows, prof]);

  const sendChat = useCallback(async () => {
    if (!chatInput.trim() || !rows || !prof) return;
    const userMsg = { role: "user", content: chatInput.trim() };
    const nextMessages = [...chatMessages, userMsg];
    setChatMessages(nextMessages);
    setChatInput("");
    setChatLoading(true);
    try {
      const system = buildDataSystemPrompt(rows, prof);
      const apiMessages = nextMessages.map((m) => ({ role: m.role, content: m.content }));
      const raw = await callClaude(apiMessages, system);
      const { cleanText, spec } = extractChartSpec(raw);
      let chartAdded = false;
      if (spec && spec.x && prof.columns.some((c) => c.name === spec.x)) {
        const data =
          spec.type === "line" && prof.dateCols.includes(spec.x)
            ? timeSeries(rows, spec.x, spec.y, spec.agg || "sum")
            : aggregate(rows, spec.x, spec.y, spec.agg || "sum");
        setCustomCharts((prev) => [...prev, { id: Date.now(), type: spec.type || "bar", title: spec.title || `${spec.x} breakdown`, data }]);
        chartAdded = true;
      }
      setChatMessages([...nextMessages, { role: "assistant", content: cleanText, chartAdded }]);
    } catch {
      setChatMessages([...nextMessages, { role: "assistant", content: "I couldn't reach the AI service just now — please try again in a moment." }]);
    } finally {
      setChatLoading(false);
      setTimeout(() => chatEndRef.current?.scrollIntoView({ behavior: "smooth" }), 50);
    }
  }, [chatInput, chatMessages, rows, prof]);

  // auto charts derived from the uploaded data
  const autoCharts = useMemo(() => {
    if (!rows || !prof) return [];
    const charts = [];
    const cat = prof.categoricalCols[0];
    const numCol = prof.numericCols[0];
    if (cat) {
      charts.push({ id: "auto-bar", type: "bar", title: numCol ? `Total ${numCol} by ${cat}` : `Count by ${cat}`, data: aggregate(rows, cat, numCol, numCol ? "sum" : "count") });
      charts.push({ id: "auto-pie", type: "pie", title: `Distribution of ${cat}`, data: aggregate(rows, cat, null, "count", 6) });
    }
    if (prof.dateCols[0] && numCol) {
      charts.push({ id: "auto-line", type: "line", title: `${numCol} over time`, data: timeSeries(rows, prof.dateCols[0], numCol, "sum") });
    }
    return charts;
  }, [rows, prof]);

  const kpis = useMemo(() => {
    if (!rows || !prof) return [];
    const cards = [{ icon: Database, color: ACCENTS[0], label: "Rows loaded", value: rows.length.toLocaleString(), sub: `${prof.columns.length} columns` }];
    prof.numericCols.slice(0, 3).forEach((c, i) => {
      const vals = rows.map((r) => num(r[c]));
      const sum = vals.reduce((a, b) => a + b, 0);
      cards.push({ icon: BarChart3, color: ACCENTS[(i + 1) % ACCENTS.length], label: `Total ${c}`, value: formatCompact(sum), sub: `avg ${formatCompact(sum / vals.length)}` });
    });
    return cards;
  }, [rows, prof]);

  const navItems = [
    { id: "upload", label: "Data Upload", icon: Upload, active: true },
    { id: "dashboard", label: "Dashboard", icon: Home, active: !!rows },
    { id: "assistant", label: "AI Assistant", icon: MessageSquare, active: !!rows },
  ];

  return (
    <div className="flex h-full min-h-[700px] bg-slate-50 font-sans text-slate-800">
      {/* Sidebar */}
      <aside className="w-56 shrink-0 text-white flex flex-col py-6 px-4 gap-1" style={{ backgroundColor: NAVY }}>
        <div className="flex items-center gap-2 px-2 mb-8">
          <div className="w-8 h-8 rounded-lg bg-blue-500 flex items-center justify-center">
            <Sparkles size={16} />
          </div>
          <div>
            <div className="text-sm font-semibold leading-tight">Insight Studio</div>
            <div className="text-[11px] text-slate-400 leading-tight">Any data. Instant analysis.</div>
          </div>
        </div>
        {navItems.map((item) => (
          <button
            key={item.id}
            disabled={!item.active}
            onClick={() => setView(item.id)}
            className={`flex items-center gap-3 px-3 py-2 rounded-lg text-sm text-left transition-colors ${
              view === item.id ? "bg-white/10 text-white" : item.active ? "text-slate-300 hover:bg-white/5" : "text-slate-600 cursor-not-allowed"
            }`}
          >
            <item.icon size={16} />
            {item.label}
          </button>
        ))}
        <div className="mt-auto px-2 pt-6 text-[11px] text-slate-500 leading-relaxed">
          {fileName ? (
            <div className="flex items-center gap-2 text-slate-300">
              <FileSpreadsheet size={13} />
              <span className="truncate">{fileName}</span>
            </div>
          ) : (
            "No dataset loaded yet"
          )}
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 overflow-y-auto p-8">
        {view === "upload" && (
          <UploadView
            onFile={handleFile}
            loading={loadingFile}
            error={error}
            fileInputRef={fileInputRef}
            hasData={!!rows}
            onGoDashboard={() => setView("dashboard")}
          />
        )}

        {view === "dashboard" && rows && prof && (
          <div className="max-w-6xl mx-auto flex flex-col gap-6">
            <div className="flex items-center justify-between">
              <div>
                <h1 className="text-xl font-semibold text-slate-900">Your data, analyzed</h1>
                <p className="text-sm text-slate-500 mt-0.5">{rows.length.toLocaleString()} rows from {fileName}</p>
              </div>
              <button
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2 text-sm border border-slate-300 rounded-lg px-3 py-1.5 text-slate-600 hover:bg-white"
              >
                <RefreshCcw size={14} /> Replace data
              </button>
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={(e) => handleFile(e.target.files?.[0])} />
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {kpis.map((k, i) => (
                <KpiCard key={i} {...k} />
              ))}
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="lg:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-6">
                {autoCharts.map((c) => (
                  <ChartCard key={c.id} title={c.title} icon={c.type === "pie" ? PieChartIcon : c.type === "line" ? LineChartIcon : BarChart3}>
                    <RenderChart type={c.type} data={c.data} />
                  </ChartCard>
                ))}
                {customCharts.map((c) => (
                  <ChartCard
                    key={c.id}
                    title={c.title}
                    icon={c.type === "pie" ? PieChartIcon : c.type === "line" ? LineChartIcon : BarChart3}
                    onRemove={() => setCustomCharts((prev) => prev.filter((x) => x.id !== c.id))}
                  >
                    <RenderChart type={c.type} data={c.data} />
                  </ChartCard>
                ))}
              </div>

              <div className="flex flex-col gap-6">
                <div className="bg-white rounded-xl border border-slate-200 p-5">
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-2">
                      <Sparkles size={16} className="text-blue-500" />
                      <h3 className="text-sm font-semibold text-slate-800">AI Insights</h3>
                    </div>
                    {!insightsLoading && (
                      <button onClick={generateInsights} className="text-xs text-blue-600 hover:underline">
                        {insights.length ? "Refresh" : "Generate"}
                      </button>
                    )}
                  </div>
                  {insightsLoading && (
                    <div className="flex items-center gap-2 text-sm text-slate-400 py-4">
                      <Loader2 size={14} className="animate-spin" /> Reading your data…
                    </div>
                  )}
                  {!insightsLoading && insights.length === 0 && (
                    <p className="text-sm text-slate-400">Click "Generate" and the assistant will read your data and summarize what stands out.</p>
                  )}
                  <ul className="flex flex-col gap-3">
                    {insights.map((line, i) => (
                      <li key={i} className="flex items-start gap-2 text-sm text-slate-600">
                        <CheckCircle2 size={14} className="text-emerald-500 mt-0.5 shrink-0" />
                        {line}
                      </li>
                    ))}
                  </ul>
                </div>

                <div className="bg-white rounded-xl border border-slate-200 p-5">
                  <div className="flex items-center gap-2 mb-3">
                    <AlertTriangle size={16} className="text-amber-500" />
                    <h3 className="text-sm font-semibold text-slate-800">Data Quality</h3>
                  </div>
                  {dataQualityIssues.length === 0 ? (
                    <p className="text-sm text-slate-400">No missing values or duplicate rows found.</p>
                  ) : (
                    <ul className="flex flex-col gap-3">
                      {dataQualityIssues.map((issue, i) => (
                        <li key={i} className="flex items-center justify-between text-sm">
                          <span className="text-slate-600">{issue.label}</span>
                          <span
                            className="text-xs font-medium px-2 py-0.5 rounded-full"
                            style={{
                              backgroundColor: issue.tone === "red" ? "#FEE2E2" : "#FEF3C7",
                              color: issue.tone === "red" ? "#B91C1C" : "#92400E",
                            }}
                          >
                            {issue.count}
                          </span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>

                <button
                  onClick={() => setView("assistant")}
                  className="flex items-center justify-between bg-blue-600 text-white rounded-xl p-4 text-sm font-medium hover:bg-blue-700"
                >
                  Ask the AI assistant about this data
                  <ChevronRight size={16} />
                </button>
              </div>
            </div>
          </div>
        )}

        {view === "assistant" && rows && prof && (
          <div className="max-w-3xl mx-auto flex flex-col h-full">
            <div className="mb-4">
              <h1 className="text-xl font-semibold text-slate-900">Ask about {fileName}</h1>
              <p className="text-sm text-slate-500 mt-0.5">Ask a question, or ask it to build you a chart — try "show me a pie chart of {prof.categoricalCols[0] || "a column"}".</p>
            </div>
            <div className="flex-1 bg-white border border-slate-200 rounded-xl p-4 overflow-y-auto flex flex-col gap-4 min-h-[380px] max-h-[480px]">
              {chatMessages.length === 0 && (
                <div className="text-sm text-slate-400 m-auto text-center max-w-xs">
                  No messages yet. Ask something like "which {prof.categoricalCols[0] || "category"} has the highest total?"
                </div>
              )}
              {chatMessages.map((m, i) => (
                <div key={i} className={`flex ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                  <div
                    className={`rounded-xl px-4 py-2.5 text-sm max-w-[85%] whitespace-pre-wrap ${
                      m.role === "user" ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-800"
                    }`}
                  >
                    {m.content}
                    {m.chartAdded && (
                      <div className="mt-2 text-xs opacity-75 flex items-center gap-1">
                        <BarChart3 size={12} /> Chart added to your dashboard
                      </div>
                    )}
                  </div>
                </div>
              ))}
              {chatLoading && (
                <div className="flex items-center gap-2 text-sm text-slate-400">
                  <Loader2 size={14} className="animate-spin" /> Thinking…
                </div>
              )}
              <div ref={chatEndRef} />
            </div>
            <div className="flex items-center gap-2 mt-3">
              <input
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && !chatLoading && sendChat()}
                placeholder="Ask a question about your data…"
                className="flex-1 border border-slate-300 rounded-lg px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
              <button
                onClick={sendChat}
                disabled={chatLoading || !chatInput.trim()}
                className="bg-blue-600 text-white rounded-lg px-4 py-2.5 disabled:opacity-40"
              >
                <Send size={16} />
              </button>
            </div>
            {customCharts.length > 0 && (
              <button onClick={() => setView("dashboard")} className="text-xs text-blue-600 hover:underline mt-3 self-start">
                View {customCharts.length} chart{customCharts.length > 1 ? "s" : ""} the assistant built →
              </button>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

function UploadView({ onFile, loading, error, fileInputRef, hasData, onGoDashboard }) {
  const [dragOver, setDragOver] = useState(false);
  return (
    <div className="max-w-xl mx-auto mt-20 flex flex-col items-center text-center">
      <div className="w-14 h-14 rounded-2xl bg-blue-600 flex items-center justify-center mb-5">
        <Upload size={24} className="text-white" />
      </div>
      <h1 className="text-2xl font-semibold text-slate-900">Bring in any dataset</h1>
      <p className="text-sm text-slate-500 mt-2 max-w-sm">
        Drop a CSV or Excel file below. Insight Studio will profile it, build charts automatically, and let you ask questions about it in plain language.
      </p>

      <label
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); onFile(e.dataTransfer.files?.[0]); }}
        className={`mt-8 w-full border-2 border-dashed rounded-xl py-14 flex flex-col items-center gap-2 cursor-pointer transition-colors ${
          dragOver ? "border-blue-500 bg-blue-50" : "border-slate-300 bg-white hover:border-slate-400"
        }`}
      >
        {loading ? (
          <>
            <Loader2 size={22} className="animate-spin text-blue-500" />
            <span className="text-sm text-slate-500">Reading your file…</span>
          </>
        ) : (
          <>
            <FileSpreadsheet size={22} className="text-slate-400" />
            <span className="text-sm text-slate-600">Click to choose a file, or drag one here</span>
            <span className="text-xs text-slate-400">.csv, .xlsx, or .xls</span>
          </>
        )}
        <input
          ref={fileInputRef}
          type="file"
          accept=".csv,.xlsx,.xls"
          className="hidden"
          onChange={(e) => onFile(e.target.files?.[0])}
        />
      </label>

      {error && <p className="text-sm text-red-600 mt-4">{error}</p>}

      {hasData && (
        <button onClick={onGoDashboard} className="text-sm text-blue-600 hover:underline mt-6">
          Go back to your current dashboard →
        </button>
      )}
    </div>
  );
}
