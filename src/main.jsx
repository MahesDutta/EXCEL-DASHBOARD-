import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Activity, BarChart3, CalendarDays, ChevronDown, Database,
  Download, FileSpreadsheet, Filter, Gauge, Layers3, Moon,
  RefreshCcw, Search, Sparkles, Sun, Table2, TrendingUp,
  Upload, X, ArrowUpRight, ArrowDownRight, CheckCircle2,
  AlertCircle, SlidersHorizontal
} from "lucide-react";
import {
  ResponsiveContainer, LineChart, Line, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  PieChart, Pie, Cell, AreaChart, Area
} from "recharts";
import "./styles.css";

const COLORS = ["#7c3aed","#06b6d4","#f59e0b","#10b981","#ef4444","#3b82f6","#ec4899","#84cc16"];
const DATE_WORDS = ["date","day","time","created","updated","invoice","order","transaction","period"];
const VALUE_WORDS = ["sales","sale","revenue","amount","value","price","cost","profit","income","total","turnover"];
const QTY_WORDS = ["quantity","qty","units","count","volume","pieces","pcs"];
const DIM_WORDS = ["product","customer","client","region","area","department","category","type","status","city","state","country","vendor","supplier","employee","name"];

function cleanHeader(v, i) {
  const s = String(v ?? "").trim();
  return s || `Column ${i + 1}`;
}
function isNumber(v) {
  return typeof v === "number" && Number.isFinite(v);
}
function toNumber(v) {
  if (isNumber(v)) return v;
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/[,₹$€£%\s]/g, ""));
  return Number.isFinite(n) ? n : null;
}
function toDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  const s = String(v ?? "").trim();
  if (!s) return null;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}
function fmtNumber(n) {
  if (!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(n);
}
function fmtCompact(n) {
  if (!Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  if (abs >= 1e9) return `${(n/1e9).toFixed(1)}B`;
  if (abs >= 1e6) return `${(n/1e6).toFixed(1)}M`;
  if (abs >= 1e3) return `${(n/1e3).toFixed(1)}K`;
  return fmtNumber(n);
}
function labelForDate(d) {
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}
function scoreColumn(name, words) {
  const n = name.toLowerCase().replace(/[^a-z0-9 ]/g," ");
  return words.reduce((s,w)=>s+(n.includes(w)?1:0),0);
}
function inferColumns(rows) {
  if (!rows.length) return {date:null,value:null,quantity:null,dimensions:[],numeric:[],all:[]};
  const keys = Object.keys(rows[0]);
  const info = keys.map(key => {
    const vals = rows.map(r=>r[key]).filter(v=>v!==null && v!==undefined && v!=="").slice(0,300);
    const nums = vals.map(toNumber).filter(v=>v!==null);
    const dates = vals.map(toDate).filter(Boolean);
    const numericRatio = vals.length ? nums.length/vals.length : 0;
    const dateRatio = vals.length ? dates.length/vals.length : 0;
    const unique = new Set(vals.map(v=>String(v))).size;
    return {key, numericRatio, dateRatio, unique, scoreValue:scoreColumn(key,VALUE_WORDS), scoreQty:scoreColumn(key,QTY_WORDS), scoreDate:scoreColumn(key,DATE_WORDS), scoreDim:scoreColumn(key,DIM_WORDS)};
  });
  const dateCandidates = info.filter(x=>x.dateRatio>.65).sort((a,b)=>b.scoreDate-a.scoreDate || b.dateRatio-a.dateRatio);
  const numeric = info.filter(x=>x.numericRatio>.7).sort((a,b)=>b.scoreValue-a.scoreValue);
  const valueCandidates = numeric.slice().sort((a,b)=>b.scoreValue-a.scoreValue);
  const qtyCandidates = numeric.slice().sort((a,b)=>b.scoreQty-a.scoreQty);
  const dimensions = info.filter(x=>!dateCandidates.some(d=>d.key===x.key) && !numeric.some(n=>n.key===x.key))
    .sort((a,b)=>b.scoreDim-a.scoreDim || a.unique-b.unique);
  return {
    date: dateCandidates[0]?.key || null,
    value: valueCandidates[0]?.key || numeric[0]?.key || null,
    quantity: qtyCandidates.find(x=>x.key!==valueCandidates[0]?.key)?.key || null,
    dimensions: dimensions.slice(0,5).map(x=>x.key),
    numeric: numeric.map(x=>x.key),
    all: keys
  };
}
function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, {type:"array", cellDates:true});
  const sheets = wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:true});
    if (!matrix.length) return {name, rows:[], headers:[]};
    const headerIndex = matrix.findIndex(row => row.filter(v=>v!=="" && v!=null).length >= 2);
    const h = matrix[Math.max(0,headerIndex)];
    const headers = h.map(cleanHeader);
    const rows = matrix.slice(Math.max(0,headerIndex)+1).filter(r => r.some(v=>v!=="" && v!=null)).map(r => {
      const obj={};
      headers.forEach((key,i)=>obj[key]=r[i] ?? "");
      return obj;
    });
    return {name, rows, headers};
  });
  return sheets;
}
function aggregateBy(rows, key, valueKey, limit=10) {
  if (!key) return [];
  const map = new Map();
  rows.forEach(r => {
    const k = String(r[key] ?? "Unknown").trim() || "Unknown";
    const v = valueKey ? toNumber(r[valueKey]) : 1;
    map.set(k, (map.get(k)||0)+(v ?? 0));
  });
  return [...map.entries()].sort((a,b)=>b[1]-a[1]).slice(0,limit).map(([name,value])=>({name,value}));
}
function trendByDate(rows,dateKey,valueKey) {
  if (!dateKey) return [];
  const map = new Map();
  rows.forEach(r => {
    const d=toDate(r[dateKey]); if(!d) return;
    const key=d.toISOString().slice(0,10);
    const v=valueKey ? (toNumber(r[valueKey]) ?? 0) : 1;
    map.set(key,(map.get(key)||0)+v);
  });
  return [...map.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(([date,value])=>({date:labelForDate(new Date(date+"T00:00:00")),value}));
}
function downloadCSV(rows) {
  if (!rows.length) return;
  const ws=XLSX.utils.json_to_sheet(rows);
  const csv=XLSX.utils.sheet_to_csv(ws);
  const blob=new Blob([csv],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="dashboard-filtered-data.csv"; a.click(); URL.revokeObjectURL(a.href);
}

function App() {
  const inputRef=useRef(null);
  const [sheets,setSheets]=useState([]);
  const [sheetIndex,setSheetIndex]=useState(0);
  const [fileName,setFileName]=useState("");
  const [error,setError]=useState("");
  const [dark,setDark]=useState(false);
  const [query,setQuery]=useState("");
  const [dateRange,setDateRange]=useState({from:"",to:""});
  const [filters,setFilters]=useState({});
  const [activeTab,setActiveTab]=useState("overview");

  const current=sheets[sheetIndex];
  const rows=current?.rows || [];
  const schema=useMemo(()=>inferColumns(rows),[rows]);

  const dimensionOptions=useMemo(()=>schema.dimensions.filter(k=>rows.some(r=>String(r[k]??"").trim())),[schema,rows]);
  const filterOptions=useMemo(()=>{
    const out={};
    dimensionOptions.forEach(k=>{
      const vals=[...new Set(rows.map(r=>String(r[k]??"").trim()).filter(Boolean))];
      out[k]=vals.slice(0,100);
    });
    return out;
  },[rows,dimensionOptions]);

  const bounds=useMemo(()=>{
    if(!schema.date) return {min:"",max:""};
    const ds=rows.map(r=>toDate(r[schema.date])).filter(Boolean).sort((a,b)=>a-b);
    return {min:ds[0]?ds[0].toISOString().slice(0,10):"",max:ds.at(-1)?ds.at(-1).toISOString().slice(0,10):""};
  },[rows,schema.date]);

  const filtered=useMemo(()=>{
    const q=query.trim().toLowerCase();
    return rows.filter(r=>{
      if(q && !Object.values(r).some(v=>String(v).toLowerCase().includes(q))) return false;
      if(schema.date && (dateRange.from || dateRange.to)) {
        const d=toDate(r[schema.date]); if(!d) return false;
        const day=d.toISOString().slice(0,10);
        if(dateRange.from && day<dateRange.from) return false;
        if(dateRange.to && day>dateRange.to) return false;
      }
      for(const k of Object.keys(filters)) if(filters[k] && String(r[k]??"")!==filters[k]) return false;
      return true;
    });
  },[rows,schema.date,dateRange,filters,query]);

  const stats=useMemo(()=>{
    const value=schema.value ? filtered.map(r=>toNumber(r[schema.value])).filter(v=>v!==null) : [];
    const qty=schema.quantity ? filtered.map(r=>toNumber(r[schema.quantity])).filter(v=>v!==null) : [];
    const total=value.reduce((a,b)=>a+b,0);
    const totalQty=qty.reduce((a,b)=>a+b,0);
    const avg=value.length?total/value.length:0;
    const max=value.length?Math.max(...value):0;
    return {total,totalQty,avg,max,count:filtered.length};
  },[filtered,schema]);

  const trend=useMemo(()=>trendByDate(filtered,schema.date,schema.value),[filtered,schema]);
  const dim=useMemo(()=>schema.dimensions.find(k=>k!==schema.date) || schema.all.find(k=>k!==schema.date && k!==schema.value && k!==schema.quantity),[schema]);
  const category=useMemo(()=>aggregateBy(filtered,dim,schema.value,8),[filtered,dim,schema.value]);
  const quantityCategory=useMemo(()=>aggregateBy(filtered,dim,schema.quantity,8),[filtered,dim,schema.quantity]);

  const previousStats=useMemo(()=>{
    if(!schema.date || !dateRange.from || !dateRange.to) return null;
    const from=new Date(dateRange.from), to=new Date(dateRange.to);
    const days=Math.max(1,Math.round((to-from)/86400000)+1);
    const pTo=new Date(from.getTime()-86400000), pFrom=new Date(pTo.getTime()-(days-1)*86400000);
    const pRows=rows.filter(r=>{
      const d=toDate(r[schema.date]); if(!d)return false;
      const x=d.toISOString().slice(0,10);
      return x>=pFrom.toISOString().slice(0,10)&&x<=pTo.toISOString().slice(0,10);
    });
    const vals=schema.value?pRows.map(r=>toNumber(r[schema.value])).filter(v=>v!==null):[];
    return {total:vals.reduce((a,b)=>a+b,0),count:pRows.length};
  },[rows,schema,dateRange]);

  const change=previousStats && previousStats.total!==0 ? ((stats.total-previousStats.total)/Math.abs(previousStats.total))*100 : null;

  async function handleFile(file){
    if(!file)return;
    setError("");
    try{
      const buffer=await file.arrayBuffer();
      const parsed=parseWorkbook(buffer);
      const usable=parsed.filter(s=>s.rows.length);
      if(!usable.length) throw new Error("No usable table was found in the workbook.");
      setSheets(usable); setSheetIndex(0); setFileName(file.name); setFilters({}); setDateRange({from:"",to:""}); setQuery(""); setActiveTab("overview");
    }catch(e){setError(e.message || "Could not read this file.");}
  }
  function reset(){
    setSheets([]);setFileName("");setError("");setFilters({});setDateRange({from:"",to:""});setQuery("");
  }
  function quickDate(type){
    if(!schema.date || !bounds.min)return;
    const max=new Date(bounds.max+"T00:00:00");
    let from=new Date(max);
    if(type==="today") from=max;
    if(type==="month") from=new Date(max.getFullYear(),max.getMonth(),1);
    if(type==="lastMonth") { from=new Date(max.getFullYear(),max.getMonth()-1,1); }
    if(type==="year") from=new Date(max.getFullYear(),0,1);
    setDateRange({from:from.toISOString().slice(0,10),to:bounds.max});
  }

  return (
    <div className={dark?"app dark":"app"}>
      <header className="topbar">
        <div className="brand"><div className="brandMark"><Sparkles size={19}/></div><div><strong>Excel Intelligence</strong><span>Dashboard Studio</span></div></div>
        <div className="topActions">
          {sheets.length>0 && <button className="ghost" onClick={reset}><RefreshCcw size={16}/> New file</button>}
          <button className="iconBtn" title="Toggle theme" onClick={()=>setDark(v=>!v)}>{dark?<Sun size={18}/>:<Moon size={18}/>}</button>
        </div>
      </header>

      {!sheets.length ? (
        <main className="landing">
          <section className="hero">
            <div className="eyebrow"><CheckCircle2 size={15}/> 100% browser-based · no API key</div>
            <h1>Turn your Excel data into a<br/><em>decision-ready dashboard.</em></h1>
            <p>Upload your workbook. The app detects dates, metrics and categories, then builds interactive KPIs, trends, rankings and filters automatically.</p>
            <div className="uploadCard" onClick={()=>inputRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files?.[0])}}>
              <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>handleFile(e.target.files?.[0])}/>
              <div className="uploadIcon"><Upload size={28}/></div>
              <h3>Drop your Excel file here</h3>
              <p>or <span>browse from your device</span></p>
              <small>Supports .xlsx, .xls and .csv · data stays in your browser</small>
            </div>
            {error && <div className="error"><AlertCircle size={18}/>{error}</div>}
            <div className="featureRow">
              <div><Database size={19}/><b>Smart detection</b><span>Finds dates, metrics & dimensions</span></div>
              <div><CalendarDays size={19}/><b>Date intelligence</b><span>Custom and quick date ranges</span></div>
              <div><BarChart3 size={19}/><b>Auto visuals</b><span>Charts adapt to your data</span></div>
            </div>
          </section>
        </main>
      ) : (
        <main className="dashboard">
          <section className="dashHead">
            <div>
              <div className="crumb"><FileSpreadsheet size={15}/>{fileName}<span>/</span>{current.name}</div>
              <h1>Performance overview</h1>
              <p>{fmtNumber(filtered.length)} records · {schema.all.length} columns · dashboard generated from your data</p>
            </div>
            <div className="headActions">
              <button className="outline" onClick={()=>downloadCSV(filtered)}><Download size={16}/> Export filtered</button>
            </div>
          </section>

          {sheets.length>1 && <div className="sheetStrip">{sheets.map((s,i)=><button className={i===sheetIndex?"sheet active":"sheet"} key={s.name} onClick={()=>{setSheetIndex(i);setFilters({});setDateRange({from:"",to:""})}}><Table2 size={15}/>{s.name}</button>)}</div>}

          <section className="controlPanel">
            <div className="controlTitle"><SlidersHorizontal size={17}/><b>Dashboard filters</b><span>{filtered.length.toLocaleString()} of {rows.length.toLocaleString()} rows</span></div>
            <div className="controls">
              {schema.date && <div className="dateControl">
                <CalendarDays size={16}/>
                <input type="date" value={dateRange.from} min={bounds.min} max={bounds.max} onChange={e=>setDateRange(v=>({...v,from:e.target.value}))}/>
                <span>→</span>
                <input type="date" value={dateRange.to} min={bounds.min} max={bounds.max} onChange={e=>setDateRange(v=>({...v,to:e.target.value}))}/>
              </div>}
              {schema.date && <div className="quick">
                <button onClick={()=>quickDate("today")}>Latest day</button><button onClick={()=>quickDate("month")}>This month</button><button onClick={()=>quickDate("lastMonth")}>Last month</button><button onClick={()=>quickDate("year")}>This year</button>
              </div>}
              {dimensionOptions.slice(0,3).map(k=><label className="selectWrap" key={k}><Filter size={14}/><select value={filters[k]||""} onChange={e=>setFilters(v=>({...v,[k]:e.target.value}))}><option value="">{k}</option>{filterOptions[k].map(v=><option key={v} value={v}>{v}</option>)}</select></label>)}
              <div className="search"><Search size={16}/><input placeholder="Search rows..." value={query} onChange={e=>setQuery(e.target.value)}/></div>
              {(query||Object.values(filters).some(Boolean)||dateRange.from||dateRange.to) && <button className="clear" onClick={()=>{setQuery("");setFilters({});setDateRange({from:"",to:""})}}><X size={15}/> Clear</button>}
            </div>
          </section>

          <nav className="tabs"><button className={activeTab==="overview"?"active":""} onClick={()=>setActiveTab("overview")}><Gauge size={16}/> Overview</button><button className={activeTab==="data"?"active":""} onClick={()=>setActiveTab("data")}><Table2 size={16}/> Data preview</button></nav>

          {activeTab==="overview" ? <section className="content">
            <div className="kpis">
              <KPI icon={<TrendingUp/>} label={schema.value||"Primary metric"} value={fmtCompact(stats.total)} sub={change!==null?`${change>=0?"+":""}${change.toFixed(1)}% vs previous period`:"Across filtered records"} positive={change===null?null:change>=0}/>
              <KPI icon={<Layers3/>} label="Records" value={fmtCompact(stats.count)} sub="Rows currently included"/>
              {schema.quantity && <KPI icon={<Activity/>} label={schema.quantity} value={fmtCompact(stats.totalQty)} sub="Total across filtered records"/>}
              <KPI icon={<BarChart3/>} label="Average" value={fmtCompact(stats.avg)} sub={`Average ${schema.value||"value"} per record`}/>
            </div>

            <div className="grid2">
              <div className="panel large">
                <PanelTitle title={`${schema.value||"Records"} over time`} subtitle={schema.date?`${schema.date} · ${trend.length} periods`:"No date column detected"} icon={<TrendingUp/>}/>
                {trend.length>1 ? <ResponsiveContainer width="100%" height={300}><AreaChart data={trend}><defs><linearGradient id="fill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopOpacity={0.25}/><stop offset="100%" stopOpacity={0}/></linearGradient></defs><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="date" tickLine={false} axisLine={false}/><YAxis tickLine={false} axisLine={false} tickFormatter={fmtCompact}/><Tooltip formatter={(v)=>fmtNumber(v)} /><Area type="monotone" dataKey="value" stroke="#7c3aed" strokeWidth={3} fill="url(#fill)" dot={false}/></AreaChart></ResponsiveContainer> : <Empty text={schema.date?"Not enough dated records for a trend.":"Add a recognizable date column to unlock trend analysis."}/>}
              </div>
              <div className="panel">
                <PanelTitle title={dim?`${dim} ranking`:"Category ranking"} subtitle="Top groups by primary metric" icon={<BarChart3/>}/>
                {category.length ? <ResponsiveContainer width="100%" height={300}><BarChart data={category} layout="vertical" margin={{left:10,right:18}}><CartesianGrid strokeDasharray="3 3" horizontal={false}/><XAxis type="number" tickFormatter={fmtCompact} tickLine={false} axisLine={false}/><YAxis type="category" dataKey="name" width={100} tickLine={false} axisLine={false}/><Tooltip formatter={(v)=>fmtNumber(v)}/><Bar dataKey="value" radius={[0,7,7,0]} fill="#06b6d4"/></BarChart></ResponsiveContainer> : <Empty text="No suitable categorical column detected."/>}
              </div>
            </div>

            <div className="grid2">
              <div className="panel">
                <PanelTitle title="Mix by category" subtitle="Share of the selected primary metric" icon={<Activity/>}/>
                {category.length ? <ResponsiveContainer width="100%" height={290}><PieChart><Pie data={category.slice(0,6)} dataKey="value" nameKey="name" innerRadius={72} outerRadius={104} paddingAngle={3}>{category.slice(0,6).map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip formatter={(v)=>fmtNumber(v)}/><Legend verticalAlign="bottom" height={36}/></PieChart></ResponsiveContainer> : <Empty text="No category data available."/ >}
              </div>
              <div className="panel">
                <PanelTitle title={schema.quantity?`${schema.quantity} by ${dim||"category"}`:"Data health"} subtitle={schema.quantity?"Top groups by quantity":"Detected structure from your workbook"} icon={<Database/>}/>
                {schema.quantity && quantityCategory.length ? <ResponsiveContainer width="100%" height={290}><BarChart data={quantityCategory}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="name" tickLine={false} axisLine={false}/><YAxis tickFormatter={fmtCompact} tickLine={false} axisLine={false}/><Tooltip formatter={(v)=>fmtNumber(v)}/><Bar dataKey="value" radius={[7,7,0,0]} fill="#f59e0b"/></BarChart></ResponsiveContainer> :
                <div className="health"><HealthRow label="Rows detected" value={rows.length.toLocaleString()} good/><HealthRow label="Columns detected" value={schema.all.length.toString()} good/><HealthRow label="Date field" value={schema.date||"Not detected"} good={!!schema.date}/><HealthRow label="Primary numeric field" value={schema.value||"Not detected"} good={!!schema.value}/></div>}
              </div>
            </div>
          </section> :
          <section className="dataPanel panel">
            <PanelTitle title="Filtered data preview" subtitle="Showing the first 100 matching records" icon={<Table2/>}/>
            <div className="tableScroll"><table><thead><tr>{schema.all.map(k=><th key={k}>{k}</th>)}</tr></thead><tbody>{filtered.slice(0,100).map((r,i)=><tr key={i}>{schema.all.map(k=><td key={k}>{String(r[k]??"")}</td>)}</tr>)}</tbody></table></div>
          </section>}
          <footer><span><CheckCircle2 size={14}/> Processed locally in your browser</span><span>Excel Intelligence Dashboard</span></footer>
        </main>
      )}
    </div>
  );
}
function KPI({icon,label,value,sub,positive}) {
  return <div className="kpi"><div className="kpiTop"><div className="kpiIcon">{icon}</div>{positive!==null && <span className={positive?"delta up":"delta down"}>{positive?<ArrowUpRight size={14}/>:<ArrowDownRight size={14}/>}</span>}</div><span className="kpiLabel">{label}</span><strong>{value}</strong><small>{sub}</small></div>;
}
function PanelTitle({title,subtitle,icon}) { return <div className="panelTitle"><div className="panelIcon">{icon}</div><div><h3>{title}</h3><p>{subtitle}</p></div></div>; }
function Empty({text}) { return <div className="empty"><BarChart3 size={27}/><span>{text}</span></div>; }
function HealthRow({label,value,good}) { return <div className="healthRow"><span>{label}</span><b>{value}</b><CheckCircle2 size={16} className={good?"ok":"muted"}/></div>; }

createRoot(document.getElementById("root")).render(<App />);
