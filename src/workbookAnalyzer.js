import * as XLSX from "xlsx";

const NULLS = new Set(["", "n/a", "na", "n.a.", "null", "none", "-", "—", "not applicable"]);
const ERROR_RE = /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A|#NUM!|#NULL!/i;
const ROLE_RULES = {
  id:[/unique identifier|unique id|identifier code|unique.*code|\buic\b|nupi|ccc number|client id/],
  name:[/kp.*name|vp.*name|client.*name|participant.*name|beneficiary.*name/],
  typology:[/key.*population|vulnerable.*population|typology|\bkp\/vp\b/],
  sex:[/^sex\b|gender/], age:[/^age\b|age \(/], county:[/county/], subcounty:[/sub.?county/], ward:[/ward/], hotspot:[/hotspot/],
  sr:[/^sr name|service recipient|sub recipient/], pe:[/peer educator|\bpe name/],
  enrollmentDate:[/date of enrolment|date of enrollment|enrolment date|enrollment date|first contact/],
  hivEnrollment:[/hiv status.*enrol|hiv status at enrollment|status at enrollment/],
  appointment:[/next appointment|appointment date|next.*visit|return date|date.*return/],
  artRefill:[/art.*refill|arv.*refill|refill.*art|refill.*arv|next.*arv|arv.*due|art.*due/],
  artStatus:[/art status|on art|art outcome|current art|hiv care status/], hivTest:[/tested.*hiv|hiv.*test|hts.*result|hiv result|test result/],
  art:[/\bart\b|antiretroviral|arv/], tb:[/\btb\b|tuberculosis/], pep:[/\bpep\b/], tpt:[/\btpt\b|tb preventive/], condoms:[/condom/], lubricant:[/lubricant/], nsp:[/\bnsp\b|needle|syringe/],
  hepb:[/hepatitis ?b|\bhep b\b/], hepc:[/hepatitis ?c|\bhep c\b/], sti:[/\bsti\b|sexually transmitted/], prep:[/\bprep\b/], cancer:[/cervical cancer|anal cancer|cervical.*screen|anal.*screen/],
  fp:[/family planning|\bfp\b|contraceptive/], mental:[/mental health|psychosocial|\bmh\b/], substance:[/alcohol|drug.*screen|substance/], violence:[/violence|gbv|gender based/], pregnancy:[/pregnan/],
  status:[/programme status|program status|client status|subsequent visit status/], mmd:[/mmd|multi.?month dispensing/], viral:[/viral load|viral level/]
};
const SERVICES=["hivTest","art","tb","pep","tpt","condoms","lubricant","nsp","hepb","hepc","sti","prep","cancer","fp","mental","substance","violence"];
const DIMS=["typology","sex","county","subcounty","ward","hotspot","sr","pe","status","hivEnrollment","artStatus","pregnancy","mmd","viral"];
const DATE_HINT=/date|appointment|refill|due|visit|enrol|enroll|contact|start|end|issued|dispens|review|follow.?up/i;

const s=v=>v==null?"":v instanceof Date?v.toISOString().slice(0,10):typeof v==="object"?(v.result!==undefined?s(v.result):v.text!==undefined?s(v.text):String(v)):String(v).replace(/\s+/g," ").trim();
const norm=v=>s(v).toLowerCase().replace(/[’']/g,"'").replace(/\s+/g," ").trim();
const blank=v=>v==null||NULLS.has(norm(v));
const num=v=>{if(typeof v==="number"&&Number.isFinite(v))return v;if(typeof v!=="string")return null;const n=Number(v.replace(/,/g,"").replace(/%/g,"").trim());return Number.isFinite(n)?n:null};
function date(v){if(v instanceof Date&&!Number.isNaN(v.getTime()))return v;if(typeof v==="number"&&v>20000&&v<60000){const d=XLSX.SSF.parse_date_code(v);return d?new Date(Date.UTC(d.y,d.m-1,d.d)):null}if(typeof v!=="string")return null;const p=v.trim().split(/[\/-]/).map(Number);if(p.length!==3||p.some(Number.isNaN))return null;let y,day;if(p[0]>31){y=p[0];day=p[2]}else{day=p[0];y=p[2]<100?2000+p[2]:p[2]}const d=new Date(Date.UTC(y,p[1]-1,day));return Number.isNaN(d.getTime())?null:d}
const iso=d=>d instanceof Date&&!Number.isNaN(d.getTime())?d.toISOString().slice(0,10):null;
function headerRow(g){let best=0,score=-1e9;for(let r=0;r<Math.min(25,g.length);r++){const a=g[r]||[];const hits=a.reduce((n,x)=>n+(s(x)&&/name|date|age|sex|status|county|ward|unique|identifier|service|hiv|art|tb|prep|pep|sti|appointment|refill|result|outcome/i.test(s(x))?1:0),0);const sc=a.filter(x=>s(x)).length*2+hits*5-r*.1;if(sc>score){score=sc;best=r}}return best}
function headers(g,r,w){const used=new Map();return Array.from({length:w},(_,c)=>{const p=[s(g[r-2]?.[c]),s(g[r-1]?.[c]),s(g[r]?.[c])].filter((x,i,a)=>x&&!/^\d+$/.test(x)&&a.indexOf(x)===i);const b=p.join(" :: ")||`Column ${c+1}`;const n=used.get(b)||0;used.set(b,n+1);return n?`${b} (${n+1})`:b})}
function top(m,n=15){return [...m.entries()].sort((a,b)=>b[1]-a[1]).slice(0,n).map(([value,count])=>({value,count}))}
function roleFor(name){const n=norm(name);for(const [r,ps] of Object.entries(ROLE_RULES))if(ps.some(p=>p.test(n)))return r;return null}
function buildRoleMap(fields){const map={};for(const r of Object.keys(ROLE_RULES))map[r]=[];for(const f of fields){const r=roleFor(f.name);if(r)map[r].push(f)}return map}
function value(row,roleMap,r){const f=roleMap[r]?.[0];return f?s(row[f.index]):""}
function state(row,roleMap,r){let any=false;for(const f of roleMap[r]||[]){const x=norm(row[f.index]);if(!x)continue;any=true;if(/positive|reactive|detected|yes|^y$|started|provided|received|completed|active|on art|current|taking|dispensed|linked/.test(x))return"done";if(/negative|non.?reactive|^no$|^n$|not done|not received|not provided|not started/.test(x))return"not_done"}return any?"recorded":"not_recorded"}

function statsFor(g,hr,hs,w){
  const a=hs.map((name,index)=>({index,name,nonEmpty:0,missing:0,numeric:0,dates:0,errors:0,sum:0,min:null,max:null,unique:new Set(),freq:new Map()}));
  const rows=[];
  for(let r=hr+1;r<g.length;r++){
    const row=g[r]||[];let has=false;
    for(let c=0;c<w;c++){
      const v=row[c],st=a[c];
      if(blank(v)){st.missing++;continue}
      has=true;st.nonEmpty++;const x=s(v);if(ERROR_RE.test(x))st.errors++;
      const n=num(v);
      if(n!==null && !(typeof v==="string"&&/[\/-]/.test(x))){st.numeric++;st.sum+=n;st.min=st.min===null?n:Math.min(st.min,n);st.max=st.max===null?n:Math.max(st.max,n)}
      else if(v instanceof Date || (DATE_HINT.test(st.name)&&typeof v==="string"&&/\d[\/-]\d[\/-]\d/.test(x)) || (DATE_HINT.test(st.name)&&typeof v==="number")){const d=date(v);if(d){st.dates++;st.min=st.min===null?d:st.min<d?st.min:d;st.max=st.max===null?d:st.max>d?st.max:d}}
      if(st.unique.size<250)st.unique.add(x);
      if(st.freq.size<100)st.freq.set(x,(st.freq.get(x)||0)+1);
    }
    if(has)rows.push(row);
  }
  const fields=a.map(st=>{const type=!st.nonEmpty?"empty":st.dates/st.nonEmpty>.8?"date":st.numeric/st.nonEmpty>.8?"numeric":st.unique.size<=Math.min(80,Math.max(12,rows.length*.2))?"category":"text";return{index:st.index,name:st.name,type,nonEmpty:st.nonEmpty,missing:st.missing,missingPct:rows.length?st.missing/rows.length*100:0,unique:st.unique.size,errorCount:st.errors,sum:st.numeric?st.sum:null,average:st.numeric?st.sum/st.numeric:null,min:st.min instanceof Date?iso(st.min):st.min,max:st.max instanceof Date?iso(st.max):st.max,topValues:type==="category"||roleFor(st.name)?top(st.freq):[]}});
  return{rows,fields};
}

function cohort(rows,fields){
  const count=rows.length, roleMap=buildRoleMap(fields), dimensions={};
  for(const r of DIMS){const m=new Map();const fs=roleMap[r]||[];for(const row of rows){const v=fs[0]?s(row[fs[0].index]):"";if(v&&!NULLS.has(norm(v)))m.set(v,(m.get(v)||0)+1)}dimensions[r]=top(m,20)}
  const coverage={};
  for(const r of SERVICES){let done=0,notDone=0,recorded=0,notRecorded=0;const fs=roleMap[r]||[];for(const row of rows){let x="not_recorded",any=false;for(const f of fs){const v=norm(row[f.index]);if(!v)continue;any=true;if(/positive|reactive|detected|yes|^y$|started|provided|received|completed|active|on art|current|taking|dispensed|linked/.test(v)){x="done";break}if(/negative|non.?reactive|^no$|^n$|not done|not received|not provided|not started/.test(v))x="not_done"}if(x==="done")done++;else if(x==="not_done")notDone++;else if(any)recorded++;else notRecorded++}coverage[r]={denominator:count,done,recorded,notDone,notRecorded,pct:count?(done+recorded)/count*100:0}}
  let positive=0,negative=0;const testFields=roleMap.hivTest||[];for(const row of rows){for(const f of testFields){const x=norm(row[f.index]);if(/positive|reactive|detected/.test(x)){positive++;break}if(/negative|non.?reactive/.test(x)){negative++;break}}}
  const tested=positive+negative,ageGroups={"<18":0,"18–24":0,"25–34":0,"35–44":0,"45+":0},af=roleMap.age?.[0];if(af)for(const row of rows){const a=num(row[af.index]);if(a===null)continue;if(a<18)ageGroups["<18"]++;else if(a<25)ageGroups["18–24"]++;else if(a<35)ageGroups["25–34"]++;else if(a<45)ageGroups["35–44"]++;else ageGroups["45+"]++}
  const now=new Date();now.setHours(0,0,0,0);const due=[],aps=roleMap.appointment||[],refills=roleMap.artRefill||[];
  for(let i=0;i<rows.length;i++){const row=rows[i];let best=null,kind="";for(const f of [...refills,...aps]){const d=date(row[f.index]);if(d&&(!best||d>best)){best=d;kind=refills.includes(f)?"ARV/ART refill":"Next appointment"}}if(best){const days=Math.round((best-now)/86400000);if(days<=30)due.push({name:value(row,roleMap,"name")||`Record ${i+1}`,id:value(row,roleMap,"id"),date:iso(best),days,kind,status:days<0?"Overdue":days===0?"Due today":"Due soon",typology:value(row,roleMap,"typology"),pe:value(row,roleMap,"pe"),sr:value(row,roleMap,"sr"),artStatus:value(row,roleMap,"artStatus")})}}
  due.sort((a,b)=>a.days-b.days);const programmeStatuses={};const statusFields=roleMap.status||[];for(const row of rows){const v=statusFields[0]?s(row[statusFields[0].index]):"";if(v&&!NULLS.has(norm(v)))programmeStatuses[v]=(programmeStatuses[v]||0)+1}
  return{recordCount:count,identifierField:roleMap.id?.[0]?.name||null,nameField:roleMap.name?.[0]?.name||null,dimensions,ageGroups,coverage,hiv:{tested,testedPct:count?tested/count*100:0,positive,negative,positivityPct:tested?positive/tested*100:0},due,programmeStatuses,roles:Object.fromEntries(Object.keys(ROLE_RULES).map(k=>[k,(roleMap[k]||[]).map(f=>f.name)]))};
}
function quality(rows,fields){const id=fields.find(f=>roleFor(f.name)==="id"),seen=new Set(),dupes=new Set();if(id)for(const row of rows){const v=norm(row[id.index]);if(v){if(seen.has(v))dupes.add(v);else seen.add(v)}}return{duplicateRows:dupes.size,formulaErrors:fields.reduce((n,f)=>n+f.errorCount,0),outliers:[]}}
function cross(rows,fields){const cats=fields.filter(f=>f.type==="category"&&f.nonEmpty).slice(0,8),out=[];for(let i=0;i<cats.length;i++)for(let j=i+1;j<cats.length;j++){const a=cats[i],b=cats[j],m=new Map();for(const row of rows){const x=s(row[a.index]),y=s(row[b.index]);if(x&&y){const k=x+"|||"+y;m.set(k,(m.get(k)||0)+1)}}const pairs=top(m,10).map(z=>{const p=z.value.split("|||");return{x:p[0],y:p[1],count:z.count}});if(pairs.length)out.push({x:a.name,y:b.name,pairs})}return out.slice(0,8)}

export function analyzeSheet(name,sheet){
  const ref=sheet["!ref"];if(!ref)return{name,rows:0,columns:0,headerRow:0,fields:[],cohort:null,quality:null,crossRelations:[],numericRelations:[]};
  const range=XLSX.utils.decode_range(ref);
  const g=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:true});
  const w=range.e.c+1,hr=headerRow(g),hs=headers(g,hr,w),st=statsFor(g,hr,hs,w);
  return{name,rows:st.rows.length,columns:w,headerRow:hr+1,fields:st.fields,cohort:st.rows.length?cohort(st.rows,st.fields):null,quality:quality(st.rows,st.fields),crossRelations:st.rows.length<=30000?cross(st.rows,st.fields):[],numericRelations:[]};
}
export function analyzeWorkbook(workbook){const names=workbook.SheetNames||[],sheets=names.map(n=>analyzeSheet(n,workbook.Sheets[n])),register=sheets.find(x=>/register/i.test(x.name))||sheets.find(x=>x.cohort);return{sheetNames:names,sheetCount:sheets.length,totalRows:sheets.reduce((n,x)=>n+x.rows,0),totalColumns:sheets.reduce((n,x)=>n+x.columns,0),sheets,cohort:register?.cohort||null,workbookQuality:{formulaErrors:sheets.reduce((n,x)=>n+(x.quality?.formulaErrors||0),0),duplicateRows:sheets.reduce((n,x)=>n+(x.quality?.duplicateRows||0),0)}}}
export function workbookInsights(result){const c=result?.cohort,n=[];if(!c)return n;if(c.hiv?.tested)n.push(`HIV testing: ${c.hiv.tested} tested, ${c.hiv.positive} positive, ${c.hiv.negative} negative (${c.hiv.positivityPct.toFixed(1)}% positivity).`);for(const [r,x] of Object.entries(c.coverage||{}))if(x.denominator&&x.pct<80)n.push(`${r}: ${x.pct.toFixed(1)}% recorded/done coverage; ${x.notRecorded} records have no value.`);if(c.due?.length)n.push(`${c.due.length} records have an ARV refill or appointment due within 30 days or overdue.`);if(result.workbookQuality?.formulaErrors)n.push(`${result.workbookQuality.formulaErrors} formula/error cells were detected.`);return n.slice(0,12)}
export {ROLE_RULES};
