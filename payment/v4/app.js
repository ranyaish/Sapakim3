// ===================== Utils =====================
const $ = sel => document.querySelector(sel);
const fmt2 = n => (Math.round((+n||0)*100)/100).toFixed(2);
const toISODate = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()).toISOString().slice(0,10);
const dayStart = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const dayEnd = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
const hours = (ms) => ms/36e5;
const two = n => String(n).padStart(2,'0');
const hhmm = d => two(d.getHours())+':'+two(d.getMinutes());
const wdHe = i => ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'][i];
const STORAGE_KEY = 'payrollSessionV7';
function normEmpName(name){ return String(name||'').replace(/\s+/g,' ').trim(); }

// ===================== Parsing workbook =====================
async function readWorkbook(file){
  const ext = file.name.toLowerCase().split('.').pop();
  const reader = new FileReader();
  const load = new Promise(res=> reader.onload = () => res(reader.result));
  if(ext==='csv') reader.readAsText(file); else reader.readAsArrayBuffer(file);
  const data = await load;
  const wb = (ext==='csv') ? XLSX.read(data, {type:'string', raw:false})
                           : XLSX.read(data, {type:'array', raw:false});
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
  return rows;
}

function parseHebDateTime(s){
  if(!s) return null;
  if(s instanceof Date) return s;
  s = String(s).trim();
  const m = s.match(/(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(m){ const d=+m[1], mo=+m[2], y=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0); return new Date(y,mo-1,d,h,mi,se); }
  const d = new Date(s); return isNaN(d) ? null : d;
}

function parsePunchesPrimary(rows){
  const titleRe = /דו"ח שעות עבודה עבור\s+(.+?)\s+בין/;
  let currentEmp = null, inTable = false; const punches = [];
  for(const row of rows){
    const r = (row||[]).map(x=> String(x??'').replace(/\u00A0/g,' ').trim());
    if(!r.length) continue;
    if(r[0]){
      const m = r[0].match(titleRe);
      if(m){ currentEmp = normEmpName(m[1]); inTable=false; continue; }
    }
    if(r.includes('שעת הגעה') && r.includes('שעת עזיבה') && r.includes('יום')){ inTable = true; continue; }
    if(inTable){
      if((r[0]||'').startsWith('סה"כ שעות עד') || r[0]==='סה"כ שעות'){ inTable=false; continue; }
      const din = parseHebDateTime(r[0]);
      const dout= parseHebDateTime(r[1]);
      if(din && dout && currentEmp){ punches.push({employee: currentEmp, dtIn: din, dtOut: dout}); }
    }
  }
  return punches;
}
function parsePunchesFallback(rows){
  const allRows = [];
  for(const row of rows){ allRows.push((row || []).map(x => String(x ?? '').replace(/\u00A0/g,' ').trim())); }
  const A = idx => (allRows[idx]?.[0]||'');
  const B = idx => (allRows[idx]?.[1]||'');
  const punches = [];
  let i=0;
  while(i<allRows.length){
    const m = A(i).match(/דו"ח שעות עבודה עבור\s+(.+?)\s+בין/);
    if(m){
      const emp = normEmpName(m[1]); i++;
      while(i<allRows.length && !A(i).includes('שעת הגעה') && !B(i).includes('שעת הגעה')) i++;
      i++;
      while(i<allRows.length){
        const a = A(i), b = B(i);
        if(/סה"כ שעות/.test(a) || /סה"כ שעות/.test(b)) break;
        const din = parseHebDateTime(A(i)); const dout = parseHebDateTime(B(i));
        if(din && dout) punches.push({employee: emp, dtIn: din, dtOut: dout});
        i++;
      }
    } else i++;
  }
  return punches;
}
function parsePunches(rows){
  const p1 = parsePunchesPrimary(rows);
  if(p1.length) return p1;
  return parsePunchesFallback(rows);
}

// ===================== Core calculations =====================
function overlap(a0,a1,b0,b1){ const s = a0>b0? a0:b0; const e = a1<b1? a1:b1; return Math.max(0, hours(e - s)); }
function iterateDailySegments(start, end){
  if(end<=start) end = new Date(end.getTime()+24*3600*1000);
  let cur = new Date(start); const segs = [];
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
  return Array.from(base.values()).sort((a,b)=> (a.employee.localeCompare(b,'he') || a.date.localeCompare(b.date)));
}
function applyOvertime(perDayBase, modeByEmp){
  return perDayBase.map(r=>{
    const mode = (modeByEmp[r.employee]?.mode || 'A').toUpperCase();
    let reg, ot125, ot150;
    if(mode==='A'){ reg = Math.min(r.nonShabbat, 9); ot125 = Math.min(Math.max(r.nonShabbat-9,0), 2); ot150 = Math.max(r.nonShabbat-11, 0); }
    else { reg = r.nonShabbat; ot125 = 0; ot150 = 0; }
    const weighted = reg + 1.25*ot125 + 1.5*(ot150 + r.shabbat150);
    return {...r, reg, ot125, ot150, weighted};
  });
}

// ===================== Extras per month =====================
function zeroExtras(){ return {travel:0,tips:0,bonus:0,advance:0}; }
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
function setExtras(emp, ym, vals){
  const e = ensureEmpConfig(emp).extras;
  e[ym] = { travel:+(vals.travel||0), tips:+(vals.tips||0), bonus:+(vals.bonus||0), advance:+(vals.advance||0) };
}

// ===================== State / Session =====================
let punches = [];
let perDayBase = [];
let empConfig = {};

function serializeSession(){
  const selEmp = $('#employeeFilter')?.value || 'ALL';
  const punchesISO = punches.map(p=> ({
    employee: normEmpName(p.employee),
    dtInISO: new Date(p.dtIn).toISOString(),
    dtOutISO: new Date(p.dtOut).toISOString()
  }));
  return { __kind:"payroll-session", version:7, savedAt:new Date().toISOString(), empConfig, punches:punchesISO, ui:{employeeFilter:selEmp} };
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
        cfg.extras[kk] = { travel:+(vv.travel||0), tips:+(vv.tips||0), bonus:+(vv.bonus||0), advance:+(vv.advance||0) };
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

// ===================== UI top selector =====================
function renderEmployeeSelect(){
  const fsel = $('#employeeFilter'); if(!fsel) return;
  const emps = [...new Set(buildDailyBase(punches).map(r=>normEmpName(r.employee)))].sort((a,b)=>a.localeCompare(b,'he'));
  fsel.innerHTML = '<option value="ALL">כל העובדים</option>' + emps.map(e=>`<option>${e}</option>`).join('');
  const dis = emps.length===0;
  fsel.disabled = dis;
  $('#exportDaily').disabled = dis; $('#exportSummary').disabled = dis; $('#exportJSON').disabled = dis;
}

// ===================== XLSX helpers =====================
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
  if (rowsArray.length) {
    const dataMatrix = rowsArray.map(row => headersHeb.map(h => row[h]));
    sheet.cell(2, 1).value(dataMatrix);
  }
  const used = sheet.usedRange(); if (used) used.style({ horizontalAlignment: "right" });
  headersHeb.forEach((h, idx) => {
    let maxLen = (h || '').toString().length;
    for (const r of rowsArray) {
      const v = (r[h] == null ? '' : String(r[h])); if (v.length > maxLen) maxLen = v.length;
    }
    sheet.column(idx + 1).width(Math.min(Math.max(10, maxLen + 2), 40));
  });
  const lastRow = rowsArray.length + 1, LIGHT = "EAF5FF", DARK  = "D6ECFF";
  for (let r = 2; r <= lastRow; r++) { sheet.row(r).style({ fill: (r % 2 === 0) ? LIGHT : DARK }); }
  const blob = await wb.outputAsync();
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename.endsWith('.xlsx') ? filename : (filename + '.xlsx'); a.click(); URL.revokeObjectURL(a.href);
}

// ===================== Export XLSX =====================
function mapConfigMode(cfg){ const m={}; for(const [k,v] of Object.entries(cfg)) m[normEmpName(k)]={mode:v.mode}; return m; }

function exportDaily(){
  const selEmp = ($('#employeeFilter').value || 'ALL');
  const perDayAll = applyOvertime(perDayBase, mapConfigMode(empConfig)).map(r=>{
    const emp = r.employee; const rate = +ensureEmpConfig(emp).rate || 0;
    const pay = r.weighted * rate; const ym = r.date.slice(0,7); const ex = getExtras(emp, ym) || {};
    const travel=+(ex.travel||0), tips=+(ex.tips||0), bonus=+(ex.bonus||0), advance=+(ex.advance||0);
    const extrasSum = travel + tips + bonus - advance; const finalPay = pay + extrasSum;
    return { employee:emp, date:r.date, total:fmt2(r.total), reg_100:fmt2(r.reg), ot_125:fmt2(r.ot125), ot_150:fmt2(r.ot150), shabbat_150:fmt2(r.shabbat150), weighted_hours:fmt2(r.weighted), hourly_rate:rate, pay:fmt2(pay), travel:fmt2(travel), tips:fmt2(tips), bonus:fmt2(bonus), advance:fmt2(advance), final_pay:fmt2(finalPay) };
  });
  const rows = (selEmp==='ALL') ? perDayAll : perDayAll.filter(r => r.employee === selEmp);
  const headersHeb = ['עובד','תאריך','סה"כ שעות','רגיל 100%','נוספות 125%','נוספות 150%','שבת 150%','שעות משוקללות','שכר לשעה','שכר יום','נסיעות','טיפים','תוספת שכר','החזר מקדמה','שכר ברוטו'];
  const rowsForXlsx = rows.map(r => ({'עובד':r.employee,'תאריך':r.date,'סה"כ שעות':r.total,'רגיל 100%':r.reg_100,'נוספות 125%':r.ot_125,'נוספות 150%':r.ot_150,'שבת 150%':r.shabbat_150,'שעות משוקללות':r.weighted_hours,'שכר לשעה':r.hourly_rate,'שכר יום':r.pay,'נסיעות':r.travel,'טיפים':r.tips,'תוספת שכר':r.bonus,'החזר מקדמה':r.advance,'שכר ברוטו':r.final_pay}));
  exportXlsx(selEmp==='ALL' ? 'per_day.xlsx' : `per_day_${selEmp}.xlsx`, headersHeb, rowsForXlsx);
}

function exportSummary(){
  const perDay = applyOvertime(perDayBase, mapConfigMode(empConfig)).map(r=>({...r, employee:normEmpName(r.employee)}));
  const byEmp = new Map();
  for(const r of perDay){
    const emp = r.employee;
    const obj = byEmp.get(emp) || {employee:emp,total:0,reg:0,ot125:0,ot150:0,shabbat150:0,weighted:0,basePay:0};
    obj.total += r.total; obj.reg += r.reg; obj.ot125 += r.ot125; obj.ot150 += r.ot150; obj.shabbat150 += r.shabbat150; obj.weighted += r.weighted;
    obj.basePay += r.weighted * (+ensureEmpConfig(emp).rate || 0); byEmp.set(emp, obj);
  }
  for(const [emp,obj] of byEmp){
    const exmap = ensureEmpConfig(emp).extras || {}; let t=0,ti=0,b=0,a=0;
    for(const [k,v] of Object.entries(exmap)){ if(k==='__default') continue; t+=+(v.travel||0); ti+=+(v.tips||0); b+=+(v.bonus||0); a+=+(v.advance||0); }
    if(t===0 && ti===0 && b===0 && a===0 && exmap.__default){ t+=+(exmap.__default.travel||0); ti+=+(exmap.__default.tips||0); b+=+(exmap.__default.bonus||0); a+=+(exmap.__default.advance||0); }
    obj.travel=t; obj.tips=ti; obj.bonus=b; obj.advance=a; obj.finalPay = obj.basePay + t + ti + b - a;
  }
  const rows = Array.from(byEmp.values()).sort((a,b)=> a.employee.localeCompare(b,'he')).map(r=> ({
    employee:r.employee,total:fmt2(r.total),reg_100:fmt2(r.reg),ot_125:fmt2(r.ot125),ot_150:fmt2(r.ot150),shabbat_150:fmt2(r.shabbat150),weighted_hours:fmt2(r.weighted),base_pay:fmt2(r.basePay),travel:fmt2(r.travel||0),tips:fmt2(r.tips||0),bonus:fmt2(r.bonus||0),advance:fmt2(r.advance||0),final_pay:fmt2(r.finalPay||0)
  }));
  const headersHeb = ['עובד','סה"כ שעות','רגיל 100%','נוספות 125%','נוספות 150%','שבת 150%','שעות משוקללות','שכר בסיס','נסיעות','טיפים','תוספת שכר','החזר מקדמה','שכר ברוטו'];
  const rowsForXlsx = rows.map(r => ({'עובד':r.employee,'סה"כ שעות':r.total,'רגיל 100%':r.reg_100,'נוספות 125%':r.ot_125,'נוספות 150%':r.ot_150,'שבת 150%':r.shabbat_150,'שעות משוקללות':r.weighted_hours,'שכר בסיס':r.base_pay,'נסיעות':r.travel,'טיפים':r.tips,'תוספת שכר':r.bonus,'החזר מקדמה':r.advance,'שכר ברוטו':r.final_pay}));
  exportXlsx('summary.xlsx', headersHeb, rowsForXlsx);
}

// ===================== Monthly summary card =====================
function computeSummaryRows() {
  const perDay = applyOvertime(perDayBase, mapConfigMode(empConfig)).map(r => ({...r, employee: normEmpName(r.employee)}));
  const byEmp = new Map();
  for (const r of perDay) {
    const emp = r.employee;
    const obj = byEmp.get(emp) || {employee: emp, total:0, reg:0, ot125:0, ot150:0, shabbat150:0, weighted:0, basePay:0, travel:0, tips:0, bonus:0, advance:0, finalPay:0};
    obj.total += r.total; obj.reg += r.reg; obj.ot125 += r.ot125; obj.ot150 += r.ot150; obj.shabbat150 += r.shabbat150; obj.weighted += r.weighted; obj.basePay += r.weighted * (+ensureEmpConfig(emp).rate || 0);
    byEmp.set(emp, obj);
  }
  for (const [emp, obj] of byEmp) {
    const exmap = ensureEmpConfig(emp).extras || {}; let t=0, ti=0, b=0, a=0;
    for (const [k, v] of Object.entries(exmap)) { if (k === '__default') continue; t += +(v.travel||0); ti += +(v.tips||0); b += +(v.bonus||0); a += +(v.advance||0); }
    if (t===0 && ti===0 && b===0 && a===0 && exmap.__default) { t += +(exmap.__default.travel||0); ti += +(exmap.__default.tips||0); b += +(exmap.__default.bonus||0); a += +(exmap.__default.advance||0); }
    obj.travel=t; obj.tips=ti; obj.bonus=b; obj.advance=a; obj.finalPay = obj.basePay + t + ti + b - a;
  }
  return Array.from(byEmp.values()).sort((a,b)=> a.employee.localeCompare(b,'he'));
}
function renderMonthlySummaryCard() {
  const tbody = document.querySelector('#summaryTable tbody');
  const tfootRow = document.querySelector('#summaryTotalsRow');
  if (!tbody || !tfootRow) return;

  const rows = computeSummaryRows();
  tbody.innerHTML = '';

  const totals = { employee:'סה״כ', total:0, reg:0, ot125:0, ot150:0, shabbat150:0, weighted:0, basePay:0, travel:0, tips:0, bonus:0, advance:0, finalPay:0 };

  for (const r of rows) {
    totals.total += r.total; totals.reg += r.reg; totals.ot125 += r.ot125; totals.ot150 += r.ot150; totals.shabbat150 += r.shabbat150; totals.weighted += r.weighted; totals.basePay += r.basePay; totals.travel += r.travel; totals.tips += r.tips; totals.bonus += r.bonus; totals.advance += r.advance; totals.finalPay += r.finalPay;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.employee}</td><td>${fmt2(r.total)}</td><td>${fmt2(r.reg)}</td><td>${fmt2(r.ot125)}</td><td>${fmt2(r.ot150)}</td><td>${fmt2(r.shabbat150)}</td><td>${fmt2(r.weighted)}</td><td>${fmt2(r.basePay)}</td><td>${fmt2(r.travel)}</td><td>${fmt2(r.tips)}</td><td>${fmt2(r.bonus)}</td><td>${fmt2(r.advance)}</td><td>${fmt2(r.finalPay)}</td>`;
    tbody.appendChild(tr);
  }
  tfootRow.innerHTML = `<th>${totals.employee}</th><td>${fmt2(totals.total)}</td><td>${fmt2(totals.reg)}</td><td>${fmt2(totals.ot125)}</td><td>${fmt2(totals.ot150)}</td><td>${fmt2(totals.shabbat150)}</td><td>${fmt2(totals.weighted)}</td><td>${fmt2(totals.basePay)}</td><td>${fmt2(totals.travel)}</td><td>${fmt2(totals.tips)}</td><td>${fmt2(totals.bonus)}</td><td>${fmt2(totals.advance)}</td><td>${fmt2(totals.finalPay)}</td>`;
}

// ===================== Employee Card (Modal) =====================
let currentEmpInModal = null;
let currentMonthKey = null; // "YYYY-MM"

function getEmployeesList(){
  const set = new Set();
  for(const p of punches){ set.add(normEmpName(p.employee)); }
  return Array.from(set).sort((a,b)=> a.localeCompare(b,'he'));
}

function monthsForEmployee(emp){
  const set = new Set();
  for(const p of punches.filter(x=> normEmpName(x.employee)===emp)){
    const d = new Date(p.dtIn);
    const key = d.getFullYear()+'-'+two(d.getMonth()+1);
    set.add(key);
  }
  return Array.from(set).sort();
}
function filterPunchesByMonth(emp, ym){
  return punches.filter(p=>{
    if(normEmpName(p.employee)!==emp) return false;
    const d = new Date(p.dtIn);
    const key = d.getFullYear()+'-'+two(d.getMonth()+1);
    return key===ym;
  });
}
function filterPerDayByMonth(emp, ym){
  const all = applyOvertime(perDayBase, mapConfigMode(empConfig));
  return all.filter(r=> normEmpName(r.employee)===emp && r.date.slice(0,7)===ym)
            .map(r=> ({...r, pay: r.weighted * (+ensureEmpConfig(emp).rate || 0)}));
}

function openEmployeeCard(emp){
  emp = normEmpName(emp);
  if (!emp) return;

  ensureEmpConfig(emp);
  currentEmpInModal = emp;

  $('#empModalTitle').textContent = `כרטיס עובד – ${emp}`;
  $('#empMode').value = ensureEmpConfig(emp).mode || 'A';
  $('#empRate').value = +ensureEmpConfig(emp).rate || 0;

  // חודשים
  const months = monthsForEmployee(emp);
  const monthSel  = $('#empMonthSel');
  const monthSel2 = $('#empMonthSel2');
  monthSel.innerHTML = '';
  monthSel2.innerHTML = '';

  if (months.length === 0) {
    monthSel.innerHTML  = '<option value="">—</option>';
    monthSel2.innerHTML = '<option value="">—</option>';
    currentMonthKey = null;
  } else {
    currentMonthKey = months[months.length - 1];
    const opts = months.map(m => `<option value="${m}">${m}</option>`).join('');
    monthSel.innerHTML  = opts;
    monthSel2.innerHTML = opts;
    monthSel.value  = currentMonthKey;
    monthSel2.value = currentMonthKey;
  }

  // טעינת נתונים לפאנלים
  loadExtrasIntoForm(emp, currentMonthKey);
  renderEmpDailyPanel(emp, currentMonthKey);
  renderEmpPunchesPanel(emp, currentMonthKey);
  updateModalTotals();
  updateModalFinal();

  // פתיחת מודאל + נעילת גלילה לרקע
  $('#empModal').classList.add('show');
  $('#empModal').setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');

  // ברירת מחדל: הצג את טאב "פרטי שכר" (או מה שתרצה)
  setActiveTab('config');

  // ליתר ביטחון: גלול את תוכן המודאל לראש
  const mb = document.querySelector('#empModal .modal-body');
  if (mb) mb.scrollTop = 0;
}

function closeEmpModal(){
  $('#empModal').classList.remove('show');
  $('#empModal').setAttribute('aria-hidden', 'true');   // ← תיקון ה-typo
  document.body.classList.remove('modal-open');
  currentEmpInModal = null;
}
window.closeEmpModal = closeEmpModal;

// חיבור הטאבים (פעם אחת ב-DOMContentLoaded)
document.addEventListener('DOMContentLoaded', () => {
  const bind = (id, key) => {
    const el = document.querySelector(id);
    if (el) el.addEventListener('click', () => setActiveTab(key));
  };
  bind('#tab-config',  'config');   // "פרטי שכר"
  bind('#tab-daily',   'daily');    // "פירוט יומי + סעיפים"
  bind('#tab-punches', 'punches');  // "משמרות"
});

function setActiveTab(id){
  ['config','daily','punches'].forEach(k=>{
    const tabBtn   = document.querySelector('#tab-' + k);
    const tabPanel = document.querySelector('#panel-' + k);
    if (tabBtn)   tabBtn.classList.toggle('active', k === id);
    if (tabPanel) tabPanel.style.display = (k === id ? '' : 'none');
  });
}

// ניווט בין עובדים
function openSiblingEmployee(step){
  if(!currentEmpInModal) return;
  const emps = getEmployeesList();
  const idx = emps.indexOf(normEmpName(currentEmpInModal));
  if(idx === -1 || emps.length === 0) return;
  const next = (idx + step + emps.length) % emps.length;
  openEmployeeCard(emps[next]);
}

// סעיפים + עידכוני סיכומים
function readExtrasFromForm(){
  return {
    travel: +($('#extra_travel').value||0),
    tips: +($('#extra_tips').value||0),
    bonus: +($('#extra_bonus').value||0),
    advance: +($('#extra_advance').value||0)
  };
}
function loadExtrasIntoForm(emp, ym){
  const ex = ym ? getExtras(emp, ym) : {travel:0,tips:0,bonus:0,advance:0};
  $('#extra_travel').value = +ex.travel || 0;
  $('#extra_tips').value = +ex.tips || 0;
  $('#extra_bonus').value = +ex.bonus || 0;
  $('#extra_advance').value = +ex.advance || 0;
}
function renderEmpDailyPanel(emp, ym){
  const rows = ym ? filterPerDayByMonth(emp, ym) : [];
  const tb = $('#empCardDaily tbody'); if(!tb) return;
  tb.innerHTML='';
  for(const r of rows){
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${r.date}</td><td>${fmt2(r.total)}</td><td>${fmt2(r.reg)}</td><td>${fmt2(r.ot125)}</td><td>${fmt2(r.ot150)}</td><td>${fmt2(r.shabbat150)}</td><td>${fmt2(r.weighted)}</td><td>${fmt2(r.pay)}</td>`;
    tb.appendChild(tr);
  }
}
function renderEmpPunchesPanel(emp, ym){
  const sh = ym ? filterPunchesByMonth(emp, ym) : [];
  const tb = $('#empPunches tbody'); if(!tb) return;
  tb.innerHTML='';
  for(const p of sh.sort((a,b)=> a.dtIn - b.dtIn)){
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${toISODate(p.dtIn)}</td><td>${wdHe(p.dtIn.getDay())}</td><td>${hhmm(p.dtIn)}</td><td>${hhmm(p.dtOut)}</td><td>${fmt2(hours(p.dtOut - p.dtIn))}</td>`;
    tb.appendChild(tr);
  }
}
function updateModalTotals(){
  if(!currentEmpInModal || !currentMonthKey){
    $('#empWgt').textContent='0'; $('#empBase').textContent='0'; $('#empExtrasSum').textContent='0';
    $('#empWgt_d').textContent='0'; $('#empBase_d').textContent='0'; $('#empExtrasSum_d').textContent='0';
    return;
  }
  const rows = filterPerDayByMonth(currentEmpInModal, currentMonthKey); let wsum = 0, base = 0;
  for(const r of rows){ wsum += r.weighted; base += r.pay; }
  const ex = getExtras(currentEmpInModal, currentMonthKey);
  const extrasSum = (+ex.travel||0) + (+ex.tips||0) + (+ex.bonus||0) - (+ex.advance||0);
  $('#empWgt').textContent = fmt2(wsum); $('#empBase').textContent = fmt2(base); $('#empExtrasSum').textContent = fmt2(extrasSum);
  $('#empWgt_d').textContent = fmt2(wsum); $('#empBase_d').textContent = fmt2(base); $('#empExtrasSum_d').textContent = fmt2(extrasSum);
}
function updateModalFinal(){
  const base = +($('#empBase').textContent||0), exSum = +($('#empExtrasSum').textContent||0);
  $('#empFinal').textContent = fmt2(base + exSum);
  $('#empFinal_d').textContent = fmt2(base + exSum);
}

// ===================== File Handling & Bindings =====================
async function handleFile(ev){
  const file = ev.target.files[0]; if(!file) return;
  try{
    const rows = await readWorkbook(file);
    punches = parsePunches(rows).map(p=> ({...p, employee:normEmpName(p.employee)}));
    if(punches.length===0){ alert('לא נמצאו נתוני משמרות בקובץ.'); return; }
    perDayBase = buildDailyBase(punches);
    for(const e of new Set(perDayBase.map(r=>normEmpName(r.employee)))) ensureEmpConfig(e);
    renderEmployeeSelect(); renderMonthlySummaryCard(); saveLocal();
  }catch(err){ console.error(err); alert('שגיאה בקריאת הקובץ'); }
}
function openCardFromSelect(){
  const emp = $('#employeeFilter').value;
  if(emp && emp!=='ALL') openEmployeeCard(emp);
  else alert('נא לבחור עובד מהרשימה.');
}

// ===================== Events =====================
window.addEventListener('DOMContentLoaded', ()=>{
  $('#file').addEventListener('change', handleFile);
  $('#exportDaily').addEventListener('click', exportDaily);
  $('#exportSummary').addEventListener('click', exportSummary);
  $('#exportJSON').addEventListener('click', ()=>{
    try{ const data = serializeSession(); const blob = new Blob([JSON.stringify(data, null, 2)], {type:'application/json'}); const a = document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=`payroll_session_${(new Date).toISOString().slice(0,10)}.json`; a.click(); URL.revokeObjectURL(a.href); }
    catch(e){ alert('שגיאה בייצוא הסשן'); }
  });
  $('#importJSON').addEventListener('click', ()=> $('#importJSONFile').click());
  $('#importJSONFile').addEventListener('change', async (ev)=>{
    const f = ev.target.files[0]; if(!f) return;
    try{ const txt = await f.text(); reviveSession(JSON.parse(txt)); } catch(e){ alert('שגיאה בטעינת JSON'); } finally { ev.target.value = ''; }
  });
  $('#clearSession').addEventListener('click', ()=>{ if(confirm('לנקות סשן מקומי?')){ clearLocal(); alert('נוקה הזיכרון המקומי.'); }});
  $('#openCardBtn').addEventListener('click', openCardFromSelect);

  // tabs
  $('#tab-config').onclick = ()=> setActiveTab('config');
  $('#tab-daily').onclick  = ()=> setActiveTab('daily');
  $('#tab-punches').onclick= ()=> setActiveTab('punches');

  // חודש בכרטיס
  $('#empMonthSel').addEventListener('change', ()=>{
    currentMonthKey = $('#empMonthSel').value || null; if(!currentEmpInModal) return;
    loadExtrasIntoForm(currentEmpInModal, currentMonthKey); renderEmpDailyPanel(currentEmpInModal, currentMonthKey); renderEmpPunchesPanel(currentEmpInModal, currentMonthKey); updateModalTotals(); updateModalFinal();
    $('#empMonthSel2').value = currentMonthKey || '';
  });
  $('#empMonthSel2').addEventListener('change', ()=>{ $('#empMonthSel').value = $('#empMonthSel2').value; $('#empMonthSel').dispatchEvent(new Event('change')); });

  // מצב/שכר לשעה
  $('#empMode').addEventListener('change', ()=>{ if(!currentEmpInModal) return; ensureEmpConfig(currentEmpInModal).mode = $('#empMode').value; saveLocal(); if(currentMonthKey){ renderEmpDailyPanel(currentEmpInModal, currentMonthKey); updateModalTotals(); updateModalFinal(); } renderMonthlySummaryCard(); });
  $('#empRate').addEventListener('input', ()=>{ if(!currentEmpInModal) return; ensureEmpConfig(currentEmpInModal).rate = +$('#empRate').value || 0; saveLocal(); if(currentMonthKey){ renderEmpDailyPanel(currentEmpInModal, currentMonthKey); updateModalTotals(); updateModalFinal(); } renderMonthlySummaryCard(); });

  // סעיפים
  $('#saveExtrasBtn').onclick = ()=>{ if(!currentEmpInModal) return; if(!currentMonthKey){ alert('אין חודש נבחר.'); return; } setExtras(currentEmpInModal, currentMonthKey, readExtrasFromForm()); updateModalTotals(); updateModalFinal(); saveLocal(); renderMonthlySummaryCard(); };

  // ניווט בין עובדים
  $('#empPrevBtn')?.addEventListener('click', ()=> openSiblingEmployee(-1));
  $('#empNextBtn')?.addEventListener('click', ()=> openSiblingEmployee(+1));
  document.addEventListener('keydown', (e)=>{
    const modalOpen = $('#empModal')?.classList.contains('show');
    if(!modalOpen) return;
    if(e.key === 'ArrowRight') openSiblingEmployee(-1); // RTL
    if(e.key === 'ArrowLeft')  openSiblingEmployee(+1);
  });

  // existing session
  loadLocalIfAny();
  $('#employeeFilter')?.addEventListener('change', renderMonthlySummaryCard);
});
