import * as XLSX from "xlsx";

const HEADER_WORDS = ["name","date","age","sex","gender","status","outcome","county","ward","facility","site","quarter","month","year","target","result","unique","identifier","code","service","hiv","art","tb","prep","pep","sti","screen","received","provided","started","completed","client","kp","vp"];
const NULL_WORDS = new Set(["","n/a","na","n.a.","null","none","-","—"]);

function text(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    if (v.result !== undefined) return text(v.result);
    if (v.text !== undefined) return text(v.text);
  }
  return String(v).replace(/\s+/g, " ").trim();
}
function isBlank(v) { return v === null || v === undefined || NULL_WORDS.has(String(v).trim().toLowerCase()); }
function numeric(v) {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v !== "string") return null;
  const s = v.replace(/,/g, "").replace(/%/g, "").trim();
  if (!s || NULL_WORDS.has(s.toLowerCase())) return null;
  const n = Number(s); return Number.isFinite(n) ? n : null;
}
function dateValue(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 20000 && v < 60000) {
    const d = XLSX.SSF.parse_date_code(v); if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d));
  }
  if (typeof v !== "string") return null;
  const s = v.trim();
  if (!/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/.test(s) && !/^\d{4}[\/-]\d{1,2}[\/-]\d{1,2}$/.test(s)) return null;
  const p = s.split(/[\/-]/).map(Number);
  const d = p[0] > 31 ? new Date(Date.UTC(p[0], p[1] - 1, p[2])) : new Date(Date.UTC(p[2] < 100 ? 2000 + p[2] : p[2], p[1] - 1, p[0]));
  return Number.isNaN(d.getTime()) ? null : d;
}
function headerScore(row, rowIndex) {
  const values = row.map(text).filter(Boolean); if (!values.length) return -Infinity;
  const unique = new Set(values.map(v => v.toLowerCase())).size;
  const hits = values.reduce((n,v) => n + (HEADER_WORDS.some(k => v.toLowerCase().includes(k)) ? 1 : 0), 0);
  const strings = values.filter(v => Number.isNaN(Number(v))).length;
  return values.length * 2 + hits * 4 + unique / values.length * 8 + strings * .25 - rowIndex * .15;
}
function findHeaderRow(grid) {
  let best = 0, score = -Infinity;
  for (let i = 0; i < Math.min(20, grid.length); i++) { const s = headerScore(grid[i] || [], i); if (s > score) { score = s; best = i; } }
  return best;
}
function makeHeaders(grid, headerRow, width) {
  const used = new Map(); const headers = [];
  for (let c = 0; c < width; c++) {
    const current = text(grid[headerRow]?.[c]); const above = text(grid[headerRow - 1]?.[c]); const above2 = text(grid[headerRow - 2]?.[c]);
    let label = current || above || `Column ${c + 1}`;
    if (current && above && above !== current && !/^\d+$/.test(above)) label = `${above} :: ${current}`;
    else if (!current && above) label = above;
    else if (current && !above && above2 && !/^\d+$/.test(above2)) label = `${above2} :: ${current}`;
    const count = used.get(label) || 0; used.set(label, count + 1); headers.push(count ? `${label} (${count + 1})` : label);
  }
  return headers;
}
function topValues(map, limit = 12) { return [...map.entries()].sort((a,b) => b[1] - a[1]).slice(0,limit).map(([value,count]) => ({value,count})); }
function makeColumnStats(headers) {
  return headers.map((name,index) => ({ index,name,nonEmpty:0,missing:0,numericCount:0,dateCount:0,textCount:0,formulaCount:0,errorCount:0,min:null,max:null,sum:0,unique:new Set(),frequencies:new Map() }));
}
function updateColumn(stat,value,formula) {
  if (formula) stat.formulaCount++;
  if (isBlank(value)) { stat.missing++; return; }
  stat.nonEmpty++;
  const s = text(value);
  if (/#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(s)) stat.errorCount++;
  if (stat.unique.size < 50000) stat.unique.add(s);
  if (stat.frequencies.size < 5000) stat.frequencies.set(s,(stat.frequencies.get(s)||0)+1);
  const n = numeric(value), d = dateValue(value);
  if (n !== null && !(typeof value === "string" && /[\/-]/.test(value))) {
    stat.numericCount++; stat.sum += n; stat.min = stat.min === null ? n : Math.min(stat.min,n); stat.max = stat.max === null ? n : Math.max(stat.max,n);
  } else if (d) {
    stat.dateCount++; stat.min = stat.min === null ? d : stat.min < d ? stat.min : d; stat.max = stat.max === null ? d : stat.max > d ? stat.max : d;
  } else stat.textCount++;
}
function finalizeColumn(stat,rowCount) {
  const nonEmpty = stat.nonEmpty || 0; let type = "text";
  if (!nonEmpty) type = "empty"; else if (stat.dateCount / nonEmpty >= .8) type = "date"; else if (stat.numericCount / nonEmpty >= .8) type = "numeric"; else if (stat.unique.size / Math.max(nonEmpty,1) <= .65) type = "category";
  return { index:stat.index,name:stat.name,type,nonEmpty:stat.nonEmpty,missing:rowCount-stat.nonEmpty,missingPct:rowCount ? stat.missing/rowCount*100 : 0,unique:stat.unique.size,formulaCount:stat.formulaCount,errorCount:stat.errorCount,sum:stat.numericCount ? stat.sum : null,average:stat.numericCount ? stat.sum/stat.numericCount : null,min:stat.min instanceof Date ? stat.min.toISOString().slice(0,10) : stat.min,max:stat.max instanceof Date ? stat.max.toISOString().slice(0,10) : stat.max,topValues:topValues(stat.frequencies)};
}
function detectFormulaErrors(sheet) {
  let formulas=0,errors=0; const range=XLSX.utils.decode_range(sheet["!ref"]||"A1:A1");
  for(let r=range.s.r;r<=range.e.r;r++) for(let c=range.s.c;c<=range.e.c;c++) { const cell=sheet[XLSX.utils.encode_cell({r,c})]; if(!cell) continue; if(cell.f){formulas++;if(/#REF!|#DIV\/0!|#VALUE!/.test(String(cell.f)))errors++;} if(typeof cell.v === "string" && /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/.test(cell.v)) errors++; }
  return {formulas,errors};
}
function analyzeSheet(sheetName,sheet) {
  const ref=sheet["!ref"]; if(!ref) return {name:sheetName,rows:0,columns:0,headerRow:0,fields:[],rowCount:0,formulaCount:0,formulaErrors:0,insights:[],charts:[]};
  const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,blankrows:false});
  const width=XLSX.utils.decode_range(ref).e.c+1, headerRow=findHeaderRow(grid), headers=makeHeaders(grid,headerRow,width), stats=makeColumnStats(headers); let rowCount=0;
  for(let r=headerRow+1;r<grid.length;r++) { const row=grid[r]||[]; let has=false; for(let c=0;c<width;c++){const value=row[c];if(!isBlank(value))has=true;const cell=sheet[XLSX.utils.encode_cell({r,c})];updateColumn(stats[c],value,Boolean(cell?.f));} if(has)rowCount++; }
  const fields=stats.map(s=>finalizeColumn(s,rowCount)); const meta=detectFormulaErrors(sheet);
  const numericFields=fields.filter(f=>f.type==="numeric"), categoryFields=fields.filter(f=>f.type==="category"), dateFields=fields.filter(f=>f.type==="date");
  const charts=[]; if(categoryFields[0])charts.push({type:"bar",dimension:categoryFields[0].name,measure:numericFields[0]?.name||null,data:categoryFields[0].topValues}); if(categoryFields[1])charts.push({type:"pie",dimension:categoryFields[1].name,measure:null,data:categoryFields[1].topValues.slice(0,8)});
  const insights=[]; if(categoryFields[0]?.topValues[0])insights.push(`${categoryFields[0].topValues[0].value} is the largest ${categoryFields[0].name} group (${categoryFields[0].topValues[0].count}).`); if(numericFields[0])insights.push(`${numericFields[0].name} totals ${Math.round(numericFields[0].sum*100)/100} across ${numericFields[0].numericCount.toLocaleString()} populated records.`); const missing=fields.filter(f=>f.missing>0).sort((a,b)=>b.missingPct-a.missingPct)[0]; if(missing)insights.push(`${missing.name} has ${missing.missingPct.toFixed(1)}% missing values.`); if(meta.errors)insights.push(`${meta.errors.toLocaleString()} formula/error cells require attention.`);
  return {name:sheetName,rows:grid.length,columns:width,headerRow:headerRow+1,rowCount,fields,numericFields:numericFields.map(f=>f.name),categoryFields:categoryFields.map(f=>f.name),dateFields:dateFields.map(f=>f.name),formulaCount:meta.formulas,formulaErrors:meta.errors,insights,charts};
}
export async function analyzeWorkbook(file,onProgress=()=>{}) {
  const buffer=await file.arrayBuffer(); const workbook=XLSX.read(buffer,{type:"array",cellDates:true,cellFormula:true,cellNF:true,cellText:true}); const sheets=[];
  for(let i=0;i<workbook.SheetNames.length;i++){const name=workbook.SheetNames[i];onProgress({current:i+1,total:workbook.SheetNames.length,name});sheets.push(analyzeSheet(name,workbook.Sheets[name]));await new Promise(r=>setTimeout(r,0));}
  const allFields=sheets.flatMap(s=>s.fields.map(f=>({...f,sheet:s.name}))); const totalRows=sheets.reduce((a,s)=>a+s.rowCount,0); const totalColumns=sheets.reduce((a,s)=>a+s.columns,0); const totalMissing=allFields.reduce((a,f)=>a+f.missing,0); const totalCells=sheets.reduce((a,s)=>a+s.rowCount*s.columns,0)||1; const totalFormulaErrors=sheets.reduce((a,s)=>a+s.formulaErrors,0);
  return {fileName:file.name,fileSize:file.size,sheets,totals:{sheets:sheets.length,rows:totalRows,columns:totalColumns,fields:allFields.length,missing:totalMissing,missingPct:totalMissing/totalCells*100,formulaErrors:totalFormulaErrors},topFields:allFields.filter(f=>f.type!=="empty").sort((a,b)=>b.nonEmpty-a.nonEmpty).slice(0,25)};
}
export function workbookInsights(model){const out=[];const {sheets,totals}=model;const largest=[...sheets].sort((a,b)=>b.rowCount-a.rowCount)[0];if(largest)out.push(`${largest.name} is the largest data sheet with ${largest.rowCount.toLocaleString()} populated records and ${largest.columns.toLocaleString()} columns.`);out.push(`${totals.sheets} worksheets and ${totals.fields.toLocaleString()} detected fields were scanned.`);if(totals.formulaErrors)out.push(`${totals.formulaErrors.toLocaleString()} formula/error cells were detected, including broken references where present.`);out.push(`${totals.missingPct.toFixed(1)}% of analyzed cells are blank or marked as missing.`);const n=model.topFields.find(f=>f.type==="numeric");if(n)out.push(`${n.name} is the strongest numeric field by populated records (${n.nonEmpty.toLocaleString()}).`);return out.slice(0,6);}
export {numeric,dateValue};
