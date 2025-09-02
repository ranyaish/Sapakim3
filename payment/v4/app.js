console.log('payroll app.js v1.9.5');

//// ===================== Utils =====================
const $ = sel => document.querySelector(sel);
const logBox = $('#log');
function log(msg){ try{ console.log(msg); if(logBox){ logBox.textContent += `${msg}\n`; } }catch(_){} }

const fmt2 = n => (Math.round((+n||0)*100)/100).toFixed(2);
const toISODate = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()).toISOString().slice(0,10);
const dayStart = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const dayEnd = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
const hours = (ms) => ms/36e5;
const two = n => String(n).padStart(2,'0');
const hhmm = d => two(d.getHours())+':'+two(d.getMinutes());
const wdHe = i => ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'][i];
function overlap(a0,a1,b0,b1){ const s = a0>b0? a0:b0; const e = a1<b1? a1:b1; return Math.max(0, hours(e - s)); }
function normEmpName(name){ return String(name||'').replace(/\s+/g,' ').trim(); }

const STORAGE_KEY = 'payrollSessionV9_195';

//// ===================== Date parsing (strict-ish) =====================
function parseHebDateTime(s){
  if(!s) return null;
  if(s instanceof Date) return isNaN(s)? null : s;
  s = String(s).trim().replace(/\u200f|\u200e/g,''); // מנקה RTL/LRM אם יש

  // dd/mm/yyyy hh:mm[:ss] או dd.mm.yyyy hh:mm[:ss]
  let m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if(m){
    const d=+m[1], mo=+m[2], y=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0);
    const dt = new Date(y,mo-1,d,h,mi,se);
    return isNaN(dt)? null : dt;
  }
  // ISO: yyyy-mm-dd[ T]hh:mm[:ss]
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if(m){
    const y=+m[1], mo=+m[2], d=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0);
    const dt = new Date(y,mo-1,d,h,mi,se);
    return isNaN(dt)? null : dt;
  }
  return null;
}

//// ===================== Read workbook (all sheets + NBSP clean) =====================
async function readWorkbook(file){
  if (!window.XLSX) { alert('XLSX לא נטען'); throw new Error('XLSX missing'); }
  const ext = file.name.toLowerCase().split('.').pop();

  const reader = new FileReader();
  const load = new Promise(res=> reader.onload = () => res(reader.result));
  if(ext==='csv') reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
  const data = await load;

  const wb = (ext==='csv')
    ? XLSX.read(data, { type:'string', raw:false })
    : XLSX.read(data, { type:'array', raw:false, cellDates:true, dateNF:"dd/mm/yyyy hh:mm" });

  const allRows = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
    for (const row of rows) {
      allRows.push((row || []).map(x => String(x ?? '').replace(/\u00A0/g,' ').trim()));
    }
    allRows.push([]); // מפריד בין גליונות
  }
  log(`[payroll] rows loaded: ${allRows.length}`);
  return allRows;
}

//// ===================== Parse punches (robust + fallback) =====================
// ---------- מחליף את parsePunches הקיים ----------
function parsePunches(rows){
  // מנרמל מחרוזות (מסיר ‎NBSP / RTL, מקצר רווחים)
  const norm = s => String(s ?? '')
    .replace(/\u00A0/g,' ')
    .replace(/\u200f|\u200e/g,'')
    .replace(/\s+/g,' ')
    .trim();

  // "דו"ח שעות עבודה עבור <שם> בין ..." – מאפשר גם ל- / ל־ וניסוחים קרובים
  const titleRe = /דו.?ח.*שעות.*עבודה.*עבור\s+(.+?)\s+בין/i;

  // מזהה שורת כותרת ומחזיר את אינדקסי העמודות (כניסה/יציאה/יום)
  function findHeaderIndices(row){
    const cols = row.map(c => norm(c));
    const idxIn  = cols.findIndex(c => /(שעת\s*הגעה|כניסה)/.test(c));
    const idxOut = cols.findIndex(c => /(שעת\s*עזיבה|יציאה)/.test(c));
    const idxDay = cols.findIndex(c => /\bיום\b/.test(c));
    if (idxIn !== -1 && idxOut !== -1) return { idxIn, idxOut, idxDay };
    return null;
  }

  // תנאי עצירה לטבלה
  function isStopRow(row){
    const t = norm(row.join(' '));
    const c0 = norm(row[0] || '');
    return (
      t === '' ||
      /^סה"?כ\s*שעות/.test(c0) ||
      /^מספר\s*ימי\s*עבודה/.test(c0) ||
      /אחוזי\s*שכר/.test(t) ||
      titleRe.test(c0) // התחלה של דו"ח חדש
    );
  }

  // מחפש סביב שורה i את שם העובד האחרון שהופיע בכותרת
  function lookupEmployeeNear(idx){
    for(let i=Math.max(0, idx-10); i<idx; i++){
      for(const cell of (rows[i]||[])){
        const s = norm(cell);
        const m = s.match(titleRe);
        if (m) return normEmpName(m[1]);
      }
    }
    return null;
  }

  let currentEmp = null;
  let inTable = false;
  let hdr = null; // {idxIn, idxOut, idxDay}
  const punches = [];
  const seen = new Set(); // למניעת כפולים

  for (let rIdx = 0; rIdx < rows.length; rIdx++){
    const row = rows[rIdx] || [];
    const r0 = norm(row[0] || '');

    // כותרת דו"ח -> קובע עובד פעיל
    if (r0){
      const m = r0.match(titleRe);
      if (m){
        currentEmp = normEmpName(m[1]);
        inTable = false;
        hdr = null;
        continue;
      }
    }

    // טרם בתוך טבלה? נסה לזהות כותרת
    if (!inTable){
      const maybeHdr = findHeaderIndices(row);
      if (maybeHdr){
        inTable = true;
        hdr = maybeHdr;
        if (!currentEmp){
          const found = lookupEmployeeNear(rIdx);
          if (found) currentEmp = found;
        }
        continue;
      }
      // לא כותרת – ממשיכים
      continue;
    }

    // בתוך טבלה
    if (isStopRow(row)){
      inTable = false;
      hdr = null;
      currentEmp = null; // יזוהה שוב לפי הכותרת הבאה
      continue;
    }

    // קורא תאים לפי האינדקסים שמצאנו בכותרת
    const inVal  = row[hdr.idxIn];
    const outVal = row[hdr.idxOut];

    const din  = parseHebDateTime(inVal);
    const dout = parseHebDateTime(outVal);

    if (din && dout){
      const emp = currentEmp || lookupEmployeeNear(rIdx) || 'לא מזוהה';
      const key = `${normEmpName(emp)}|${din.toISOString()}|${dout.toISOString()}`;
      if (!seen.has(key)){
        punches.push({ employee: normEmpName(emp), dtIn: din, dtOut: dout });
        seen.add(key);
      }
    }
  }

  try { console.log('[payroll] punches parsed:', punches.length); } catch(_){}
  return punches;
}


function parsePunchesFallback(rows){
  const out = [];
  let currentEmp = null;
  const titleReLoose = /עבור\s+(.+?)\s+בין/;

  function guessEmp(around){
    for(let i=Math.max(0,around-12); i<around; i++){
      for(const cell of (rows[i]||[])){
        const m = String(cell||'').match(titleReLoose);
        if(m) return normEmpName(m[1]);
      }
    }
    return 'לא מזוהה';
  }

  for(let i=0;i<rows.length;i++){
    const r = rows[i] || [];
    for(const cell of r){
      const m = String(cell||'').match(titleReLoose);
      if(m){ currentEmp = normEmpName(m[1]); break; }
    }
    const d0 = parseHebDateTime(String(r[0]||'')); 
    const d1 = parseHebDateTime(String(r[1]||''));
    if(d0 && d1){
      const emp = currentEmp || guessEmp(i);
      out.push({ employee: emp, dtIn: d0, dtOut: d1 });
    }
  }
  log(`[payroll] fallback found: ${out.length}`);
  return out;
}

function safeParsePunches(rows){
  let p = parsePunches(rows);
  if (!p || p.length===0) {
    log('[payroll] primary parse=0; trying fallback');
    p = parsePunchesFallback(rows);
  }
  const s = new Set();
  p = p.filter(x=>{
    const k = `${normEmpName(x.employee)}|${new Date(x.dtIn).toISOString()}|${new Date(x.dtOut).toISOString()}`;
    if(s.has(k)) return false; s.add(k); return true;
  });
  log(`[payroll] punches total (dedup): ${p.length}`);
  return p;
}

//// ===================== Core calculations =====================
function iterateDailySegments(start, end){
  if(end<=start) end = new Date(end.getTime()+24*3600*1000);
  let cur = new Date(start);
  const segs = [];
  while(cur < end){
    const de = dayEnd(cur);
    const segEnd = (de<end ? de : end);
    segs.push({segStart: new Date(cur), segEnd});
    cur = segEnd;
  }
  return segs;
}
function shabbat150ForSegment(segStart, segEnd){
  const ds = dayStart(segStart);
  if(ds.getDay() !== 6) return 0;
  const w0 = new Date(ds.getFullYear(), ds.getMonth(), ds.getDate(), 8,0,0);
  const w1 = new Date(ds.getFullYear(), ds.getMonth(), ds.getDate(),17,0,0);
  return overlap(segStart, segEnd, w0, w1);
}
function buildDailyBase(_punches){
  const base = new Map();
  for(const p of _punches){
    const emp = normEmpName(p.employee);
    let s = p.dtIn, e = p.dtOut; if(e<=s) e = new Date(e.getTime()+24*3600*1000);
    const noShabbat = (s.getDay()===6 && s.getHours()>=16);
    const segs = iterateDailySegments(s,e);
    for(const seg of segs){
      const total = hours(seg.segEnd - seg.segStart);
      const sh150 = noShabbat ? 0 : shabbat150ForSegment(seg.segStart, seg.segEnd);
      const nonSh = Math.max(0, total - sh150);
      const iso = toISODate(seg.segStart);
      const key = emp + '|' + iso;
      const obj = base.get(key) || {employee: emp, date: iso, total:0, nonShabbat:0, shabbat150:0};
      obj.total += total; obj.nonShabbat += nonSh; obj.shabbat150 += sh150;
      base.set(key, obj);
    }
  }
  return Array.from(base.values()).sort((a,b)=> (a.employee.localeCompare(b.employee) || a.date.localeCompare(b.date)));
}
function applyOvertime(perDayBase, modeByEmp){
  return perDayBase.map(r=>{
    const mode = (modeByEmp[r.employee]?.mode || 'A').toUpperCase();
    let reg, ot125, ot150;
    if(mode==='A'){
      reg = Math.min(r.nonShabbat, 9);
      ot125 = Math.min(Math.max(r.nonShabbat-9,0), 2);
      ot150 = Math.max(r.nonShabbat-11, 0);
    } else {
      reg = r.nonShabbat; ot125 = 0; ot150 = 0;
    }
    const weighted = reg + 1.25*ot125 + 1.5*(ot150 + r.shabbat150);
    return {...r, reg, ot125, ot150, weighted};
  });
}

//// ===================== Extras per month =====================
function zeroExtras(){ return {travel:0,tips:0,bonus:0,advance:0}; }
let empConfig = {}; // { [name]: {mode:'A'|'B', rate:number, extras:{[YYYY-MM]:{...}} } }
function ensureEmpConfig(emp){
  emp = normEmpName(emp);
  if(!empConfig[emp]) empConfig[emp] = {mode:'A', rate:0, extras:{}};
  if(!empConfig[emp].extras) empConfig[emp].extras = {};
  return empConfig[emp];
}
function getExtras(emp, ym){
  const e = ensureEmpConfig(emp).extras;
  return {...(e[ym] || e.__default || zeroExtras())};
}

//// ===================== State / Session =====================
let punches = [];   // shifts
let perDayBase = [];
function serializeSession(){
  const selEmp = $('#employeeFilter')?.value || 'ALL';
  const punchesISO = punches.map(p=> ({
    employee: normEmpName(p.employee),
    dtInISO: new Date(p.dtIn).toISOString(),
    dtOutISO: new Date(p.dtOut).toISOString()
  }));
  return { __kind:"payroll-session", version:"1.9.5", savedAt:new Date().toISOString(), empConfig, punches:punchesISO, ui:{employeeFilter:selEmp} };
}
function reviveSession(obj){
  if(!obj || obj.__kind!=='payroll-session' || !Array.isArray(obj.punches)) throw new Error('קובץ JSON אינו בפורמט הנכון');
  punches = obj.punches.map(p=> ({ employee: normEmpName(p.employee), dtIn:new Date(p.dtInISO), dtOut:new Date(p.dtOutISO) }));
  empConfig = {};
  for(const [k,v] of Object.entries(obj.empConfig||{})){
    const emp = normEmpName(k);
    const cfg = {mode:(v.mode||'A'), rate:+(v.rate||0), extras:{}};
    if(v.extras && typeof v.extras === 'object'){
      for(const [kk,vv] of Object.entries(v.extras)){
        if(!vv) continue;
        cfg.extras[kk] = {
          travel:+(vv.travel||0),
          tips:+(vv.tips||0),
          bonus:+(vv.bonus||0),
          advance:+(vv.advance||0)
        };
      }
    }
    empConfig[emp] = cfg;
  }
  perDayBase = buildDailyBase(punches);
  renderEmployeeSelect();
  renderMonthlySummaryCard();
  saveLocal();
}
function saveLocal(){ try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(serializeSession())); }catch(e){} }
function loadLocalIfAny(){ try{ const raw = localStorage.getItem(STORAGE_KEY); if(raw) reviveSession(JSON.parse(raw)); }catch(e){} }
function clearLocal(){ try{ localStorage.removeItem(STORAGE_KEY); }catch(e){} }

//// ===================== UI helpers =====================
function renderEmployeeSelect(){
  const fsel = $('#employeeFilter');
  if(!fsel) return;
  const emps = [...new Set(buildDailyBase(punches).map(r=>normEmpName(r.employee)))].sort((a,b)=>a.localeCompare(b));
  fsel.innerHTML = '<option value="ALL">כל העובדים</option>' + emps.map(e=>`<option>${e}</option>`).join('');
  const disabled = emps.length===0;
  fsel.disabled = disabled;
  ['exportDaily','exportSummary','exportJSON','openCardBtn'].forEach(id => { const el = $('#'+id); if(el) el.disabled = disabled; });
}

function computeSummaryRows() {
  const perDay = applyOvertime(perDayBase, (()=>{ const m={}; for(const [k,v] of Object.entries(empConfig)) m[normEmpName(k)]={mode:v.mode}; return m; })())
    .map(r => ({...r, employee: normEmpName(r.employee)}));

  const byEmp = new Map();
  for (const r of perDay) {
    const emp = r.employee;
    const obj = byEmp.get(emp) || {employee:emp,total:0,reg:0,ot125:0,ot150:0,shabbat150:0,weighted:0,basePay:0,travel:0,tips:0,bonus:0,advance:0,finalPay:0};
    obj.total += r.total; obj.reg += r.reg; obj.ot125 += r.ot125; obj.ot150 += r.ot150; obj.shabbat150 += r.shabbat150; obj.weighted += r.weighted;
    obj.basePay += r.weighted * (+ensureEmpConfig(emp).rate || 0);
    byEmp.set(emp, obj);
  }
  for (const [emp, obj] of byEmp) {
    const exmap = ensureEmpConfig(emp).extras || {};
    let t=0, ti=0, b=0, a=0;
    for (const [k, v] of Object.entries(exmap)) {
      if (k === '__default') continue;
      t += +(v.travel||0); ti += +(v.tips||0); b += +(v.bonus||0); a += +(v.advance||0);
    }
    if (t===0 && ti===0 && b===0 && a===0 && exmap.__default) {
      t += +(exmap.__default.travel||0);
      ti += +(exmap.__default.tips||0);
      b += +(exmap.__default.bonus||0);
      a += +(exmap.__default.advance||0);
    }
    obj.travel=t; obj.tips=ti; obj.bonus=b; obj.advance=a;
    obj.finalPay = obj.basePay + t + ti + b - a;
  }
  return Array.from(byEmp.values()).sort((a,b)=> a.employee.localeCompare(b.employee));
}

function renderMonthlySummaryCard() {
  const tbody = document.querySelector('#summaryTable tbody');
  const tfootRow = document.querySelector('#summaryTotalsRow');
  if (!tbody || !tfootRow) return;

  const rows = computeSummaryRows();
  tbody.innerHTML = '';

  const totals = { employee:'סה״כ', total:0, reg:0, ot125:0, ot150:0, shabbat150:0, weighted:0, basePay:0, travel:0, tips:0, bonus:0, advance:0, finalPay:0 };

  for (const r of rows) {
    totals.total += r.total; totals.reg += r.reg; totals.ot125 += r.ot125; totals.ot150 += r.ot150;
    totals.shabbat150 += r.shabbat150; totals.weighted += r.weighted; totals.basePay += r.basePay;
    totals.travel += r.travel; totals.tips += r.tips; totals.bonus += r.bonus; totals.advance += r.advance; totals.finalPay += r.finalPay;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.employee}</td>
      <td>${fmt2(r.total)}</td>
      <td>${fmt2(r.reg)}</td>
      <td>${fmt2(r.ot125)}</td>
      <td>${fmt2(r.ot150)}</td>
      <td>${fmt2(r.shabbat150)}</td>
      <td>${fmt2(r.weighted)}</td>
      <td>${fmt2(r.basePay)}</td>
      <td>${fmt2(r.travel)}</td>
      <td>${fmt2(r.tips)}</td>
      <td>${fmt2(r.bonus)}</td>
      <td>${fmt2(r.advance)}</td>
      <td>${fmt2(r.finalPay)}</td>`;
    tbody.appendChild(tr);
  }

  tfootRow.innerHTML = `
    <th>${totals.employee}</th>
    <td>${fmt2(totals.total)}</td>
    <td>${fmt2(totals.reg)}</td>
    <td>${fmt2(totals.ot125)}</td>
    <td>${fmt2(totals.ot150)}</td>
    <td>${fmt2(totals.shabbat150)}</td>
    <td>${fmt2(totals.weighted)}</td>
    <td>${fmt2(totals.basePay)}</td>
    <td>${fmt2(totals.travel)}</td>
    <td>${fmt2(totals.tips)}</td>
    <td>${fmt2(totals.bonus)}</td>
    <td>${fmt2(totals.advance)}</td>
    <td>${fmt2(totals.finalPay)}</td>`;
}

//// ===================== Export XLSX (RTL + bold + zebra) =====================
async function exportXlsx(filename, headersHeb, rowsArray) {
  const X = window.XlsxPopulate;
  if (!X) { alert('XlsxPopulate לא נטען'); return; }
  const wb = await X.fromBlankAsync();
  const sheet = wb.sheet(0).name("Export");
  try {
    if (typeof sheet.rightToLeft === 'function') sheet.rightToLeft(true);
    else if (sheet._sheet) { sheet._sheet.sheetViews = sheet._sheet.sheetViews || [{ workbookViewId: 0 }]; sheet._sheet.sheetViews[0].rightToLeft = 1; }
  } catch (_) {}
  sheet.cell(1, 1).value([headersHeb]); sheet.row(1).style({ bold: true });
  if (rowsArray.length) { const dataMatrix = rowsArray.map(row => headersHeb.map(h => row[h])); sheet.cell(2, 1).value(dataMatrix); }
  const used = sheet.usedRange(); if (used) used.style({ horizontalAlignment: "right" });
  headersHeb.forEach((h, idx) => {
    let maxLen = (h || '').toString().length;
    for (const r of rowsArray) { const v = (r[h] == null ? '' : String(r[h])); if (v.length > maxLen) maxLen = v.length; }
    sheet.column(idx + 1).width(Math.min(Math.max(10, maxLen + 2), 40));
  });
  const lastRow = rowsArray.length + 1;
  const LIGHT = "EAF5FF", DARK = "D6ECFF";
  for (let r = 2; r <= lastRow; r++) sheet.row(r).style({ fill: (r % 2 === 0) ? LIGHT : DARK });
  const blob = await wb.outputAsync();
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename.endsWith('.xlsx') ? filename : (filename + '.xlsx'); a.click(); URL.revokeObjectURL(a.href);
}

function exportDaily(){
  const selEmp = ($('#employeeFilter').value || 'ALL');
  const modeMap = (()=>{ const m={}; for(const [k,v] of Object.entries(empConfig)) m[normEmpName(k)]={mode:v.mode}; return m; })();
  const perDayAll = applyOvertime(perDayBase, modeMap).map(r=>{
    const emp = r.employee;
    const rate = +ensureEmpConfig(emp).rate || 0;
    const pay = r.weighted * rate;
    const ym = r.date.slice(0,7);
    const ex = getExtras(emp, ym) || {};
    const travel  = +(ex.travel  || 0);
    const tips    = +(ex.tips    || 0);
    const bonus   = +(ex.bonus   || 0);
    const advance = +(ex.advance || 0);
    const extrasSum = travel + tips + bonus - advance;
    const finalPay = pay + extrasSum;
    return {
      'עובד': emp, 'תאריך': r.date, 'סה"כ שעות': fmt2(r.total), 'רגיל 100%': fmt2(r.reg),
      'נוספות 125%': fmt2(r.ot125), 'נוספות 150%': fmt2(r.ot150), 'שבת 150%': fmt2(r.shabbat150),
      'שעות משוקללות': fmt2(r.weighted), 'שכר לשעה': rate, 'שכר יום': fmt2(pay),
      'נסיעות': fmt2(travel), 'טיפים': fmt2(tips), 'תוספת שכר': fmt2(bonus),
      'החזר מקדמה': fmt2(advance), 'שכר ברוטו': fmt2(finalPay)
    };
  });
  const rows = (selEmp==='ALL') ? perDayAll : perDayAll.filter(r => r['עובד'] === selEmp);
  const headersHeb = ['עובד','תאריך','סה"כ שעות','רגיל 100%','נוספות 125%','נוספות 150%','שבת 150%','שעות משוקללות','שכר לשעה','שכר יום','נסיעות','טיפים','תוספת שכר','החזר מקדמה','שכר ברוטו'];
  exportXlsx(selEmp==='ALL' ? 'per_day.xlsx' : `per_day_${selEmp}.xlsx`, headersHeb, rows);
}

function exportSummary(){
  const rows = computeSummaryRows().map(r=> ({
    'עובד': r.employee, 'סה"כ שעות': fmt2(r.total), 'רגיל 100%': fmt2(r.reg), 'נוספות 125%': fmt2(r.ot125),
    'נוספות 150%': fmt2(r.ot150), 'שבת 150%': fmt2(r.shabbat150), 'שעות משוקללות': fmt2(r.weighted),
    'שכר בסיס': fmt2(r.basePay), 'נסיעות': fmt2(r.travel), 'טיפים': fmt2(r.tips),
    'תוספת שכר': fmt2(r.bonus), 'החזר מקדמה': fmt2(r.advance), 'שכר ברוטו': fmt2(r.finalPay)
  }));
  const headersHeb = ['עובד','סה"כ שעות','רגיל 100%','נוספות 125%','נוספות 150%','שבת 150%','שעות משוקללות','שכר בסיס','נסיעות','טיפים','תוספת שכר','החזר מקדמה','שכר ברוטו'];
  exportXlsx('summary.xlsx', headersHeb, rows);
}

//// ===================== File Handling & Bindings =====================
async function handleFile(ev){
  const file = ev.target.files[0]; if(!file) return;
  try{
    const rows = await readWorkbook(file);
    punches = safeParsePunches(rows).map(p=> ({...p, employee:normEmpName(p.employee)}));
    if(punches.length===0){ alert('לא נמצאו נתוני משמרות בקובץ.'); return; }
    perDayBase = buildDailyBase(punches);
    for(const e of new Set(perDayBase.map(r=>normEmpName(r.employee)))) ensureEmpConfig(e);
    renderEmployeeSelect();
    renderMonthlySummaryCard();
    saveLocal();
  }catch(err){ console.error(err); alert('שגיאה בקריאת הקובץ'); }
}

function openEmployeeCard(emp){
  alert(`כרטיס עובד (דמו) עבור: ${emp}\n(בגרסה נקייה זו הצגנו רק בדיקת טעינה/ייצוא)`);
}

function openCardFromSelect(){
  const emp = $('#employeeFilter')?.value;
  if(emp && emp!=='ALL') openEmployeeCard(emp);
  else alert('נא לבחור עובד מהרשימה.');
}

window.addEventListener('DOMContentLoaded', ()=>{
  const fileInp = $('#file'); if (fileInp) fileInp.addEventListener('change', handleFile);
  const impJson = $('#importJSONFile'); if (impJson) impJson.addEventListener('change', async (ev)=>{
    const f = ev.target.files[0]; if(!f) return;
    try{ const txt = await f.text(); reviveSession(JSON.parse(txt)); }
    catch(e){ alert('שגיאה בטעינת JSON'); }
    finally{ ev.target.value = ''; }
  });
  const btnDaily = $('#exportDaily'); if(btnDaily) btnDaily.addEventListener('click', exportDaily);
  const btnSum   = $('#exportSummary'); if(btnSum) btnSum.addEventListener('click', exportSummary);
  const btnJson  = $('#exportJSON'); if(btnJson) btnJson.addEventListener('click', ()=>{
    try{
      const data = serializeSession();
      const blob = new Blob([JSON.stringify(data, null, 2)], {type:'application/json'});
      const a = document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.download=`payroll_session_${(new Date).toISOString().slice(0,10)}.json`;
      a.click();
      URL.revokeObjectURL(a.href);
    }catch(e){ alert('שגיאה בייצוא הסשן'); }
  });
  const btnClr   = $('#clearSession'); if(btnClr) btnClr.addEventListener('click', ()=>{ clearLocal(); alert('נוקה הזיכרון המקומי.'); location.reload(); });
  const btnOpen  = $('#openCardBtn'); if(btnOpen) btnOpen.addEventListener('click', openCardFromSelect);

  loadLocalIfAny();
  const empSelTop = $('#employeeFilter');
  if (empSelTop) empSelTop.addEventListener('change', renderMonthlySummaryCard);
});
