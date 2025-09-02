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
const STORAGE_KEY = 'payrollSessionV5';

function overlap(a0,a1,b0,b1){ const s = a0>b0? a0:b0; const e = a1<b1? a1:b1; return Math.max(0, hours(e - s)); }
function parseHebDateTime(s){
  if(!s) return null;
  if(s instanceof Date) return s;
  s = String(s).trim();
  const m = s.match(/(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(m){ const d=+m[1], mo=+m[2], y=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0); return new Date(y,mo-1,d,h,mi,se); }
  const d = new Date(s); return isNaN(d) ? null : d;
}
function normEmpName(name){ return String(name||'').replace(/\s+/g,' ').trim(); }

// ===================== Parsing workbook =====================
async function readWorkbook(file){
  const ext = file.name.toLowerCase().split('.').pop();
  const reader = new FileReader();
  const load = new Promise(res=> reader.onload = () => res(reader.result));
  if(ext==='csv') reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
  const data = await load;
  const wb = (ext==='csv') ? XLSX.read(data, {type:'string', raw:false})
                           : XLSX.read(data, {type:'array', raw:false});
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
  return rows;
}

function parsePunches(rows){
  const titleRe = /דו"ח שעות עבודה עבור\s+(.+?)\s+בין/;
  let currentEmp = null, inTable = false; const punches = [];
  for(const row of rows){
    const r = (row||[]).map(x=> x==null? "" : String(x).trim());
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

// ===================== Core calculations =====================
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
function summarizeByEmployee(perDayRows, rateByEmp, extrasByEmp){
  const byEmp = new Map();
  for(const r of perDayRows){
    const emp = normEmpName(r.employee);
    const obj = byEmp.get(emp) || {
      employee:emp,total:0,reg:0,ot125:0,ot150:0,shabbat150:0,weighted:0,basePay:0,
      travel:0,tips:0,bonus:0,advance:0,finalPay:0
    };
    obj.total += r.total; obj.reg += r.reg; obj.ot125 += r.ot125; obj.ot150 += r.ot150; obj.shabbat150 += r.shabbat150; obj.weighted += r.weighted;
    const rate = +(rateByEmp[emp]?.rate || 0);
    obj.basePay += r.weighted * rate;
    byEmp.set(emp, obj);
  }
  for(const [emp, row] of byEmp.entries()){
    const ex = extrasByEmp[emp] || {travel:0,tips:0,bonus:0,advance:0};
    row.travel = +ex.travel||0; row.tips = +ex.tips||0; row.bonus = +ex.bonus||0; row.advance = +ex.advance||0;
    row.finalPay = row.basePay + row.travel + row.tips + row.bonus - row.advance;
  }
  return Array.from(byEmp.values()).sort((a,b)=> a.employee.localeCompare(b.employee));
}

// ===================== UI State =====================
let punches = [];
let perDayBase = [];
// empConfig: { [emp]: { mode:'A'|'B', rate:number, extras:{travel,tips,bonus,advance} } }
let empConfig = {};
function ensureEmpConfig(emp){
  emp = normEmpName(emp);
  if(!empConfig[emp]) empConfig[emp] = {mode:'A', rate:0, extras:{travel:0,tips:0,bonus:0,advance:0}};
  if(!empConfig[emp].extras) empConfig[emp].extras = {travel:0,tips:0,bonus:0,advance:0};
  return empConfig[emp];
}

// ======= Session save/load =======
function serializeSession(){
  const selEmp = $('#employeeFilter')?.value || 'ALL';
  const punchesISO = punches.map(p=> ({
    employee: normEmpName(p.employee),
    dtInISO: new Date(p.dtIn).toISOString(),
    dtOutISO: new Date(p.dtOut).toISOString()
  }));
  return {
    __kind: "payroll-session",
    version: 5,
    savedAt: new Date().toISOString(),
    empConfig,
    punches: punchesISO,
    ui: { employeeFilter: selEmp }
  };
}
function reviveSession(obj){
  if(!obj || obj.__kind!=='payroll-session' || !Array.isArray(obj.punches)) throw new Error('קובץ JSON אינו בפורמט הנכון');
  punches = obj.punches.map(p=> ({ employee: normEmpName(p.employee), dtIn: new Date(p.dtInISO), dtOut: new Date(p.dtOutISO) }));
  empConfig = {};
  for(const [k,v] of Object.entries(obj.empConfig||{})){
    empConfig[normEmpName(k)] = {mode:(v.mode||'A'), rate:+(v.rate||0), extras:{travel:+(v.extras?.travel||0), tips:+(v.extras?.tips||0), bonus:+(v.extras?.bonus||0), advance:+(v.extras?.advance||0)}};
  }
  perDayBase = buildDailyBase(punches);
  renderEmployeeTable(); updateOutputs();
  const fsel = $('#employeeFilter');
  if(obj.ui?.employeeFilter && fsel.querySelector(`option[value="${CSS.escape(obj.ui.employeeFilter)}"]`)){ fsel.value = obj.ui.employeeFilter; updateOutputs(); }
  saveLocal();
}
function saveLocal(){ try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(serializeSession())); }catch(e){} }
function loadLocalIfAny(){ try{ const raw = localStorage.getItem(STORAGE_KEY); if(raw) reviveSession(JSON.parse(raw)); }catch(e){} }
function clearLocal(){ try{ localStorage.removeItem(STORAGE_KEY); }catch(e){} }

// ===================== Rendering =====================
function renderEmployeeTable(){
  const tbody = $('#empTable tbody'); tbody.innerHTML='';
  const emps = [...new Set(perDayBase.map(r=>normEmpName(r.employee)))].sort((a,b)=>a.localeCompare(b));
  for(const e of emps){ ensureEmpConfig(e); }

  for(const e of emps){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="link empLink" data-emp="${e}">${e}</span></td>
      <td><button class="btn alt openCardRow" data-emp="${e}">כרטיס</button></td>`;
    tbody.appendChild(tr);
  }
  for(const a of tbody.querySelectorAll('.empLink')){ a.addEventListener('click', ()=> openEmployeeCard(a.dataset.emp)); }
  for(const b of tbody.querySelectorAll('.openCardRow')){ b.addEventListener('click', ()=> openEmployeeCard(b.dataset.emp)); }

  const fsel = $('#employeeFilter'); 
  fsel.innerHTML = '<option value="ALL">כל העובדים</option>' + emps.map(e=>`<option>${e}</option>`).join(''); 
  fsel.disabled = false; 
  fsel.onchange = ()=>{ updateOutputs(); saveLocal(); };

  $('#exportDaily').disabled = false; $('#exportSummary').disabled = false; $('#exportJSON').disabled = false; $('#openCardBtn').disabled = false;
}

function updateOutputs(){
  const perDay = applyOvertime(perDayBase, mapConfigMode(empConfig));
  const perDayWithPay = perDay.map(r=> ({...r, employee:normEmpName(r.employee), pay: r.weighted * (+ensureEmpConfig(r.employee).rate || 0)}));

  renderSummary(summarizeByEmployee(perDayWithPay, mapConfigRate(empConfig), mapConfigExtras(empConfig)));
  renderDaily(perDayWithPay);
}

function mapConfigMode(cfg){ const m={}; for(const [k,v] of Object.entries(cfg)) m[normEmpName(k)]={mode:v.mode}; return m; }
function mapConfigRate(cfg){ const m={}; for(const [k,v]of Object.entries(cfg)) m[normEmpName(k)]={rate:+v.rate||0}; return m; }
function mapConfigExtras(cfg){ const m={}; for(const [k,v]of Object.entries(cfg)) m[normEmpName(k)]={...(v.extras||{})}; return m; }

function renderSummary(rows){
  const tbody = $('#summaryTable tbody'); tbody.innerHTML = '';
  for(const r of rows){
    const emp = normEmpName(r.employee);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="link" onclick="openEmployeeCard('${emp.replace(/'/g,"\\'")}')">${emp}</span></td>
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
}

function renderDaily(rows){
  const head = $('#dailyHead'); const body = $('#dailyTable tbody'); body.innerHTML='';
  const selEmp = $('#employeeFilter').value || 'ALL';
  if(selEmp==='ALL'){
    head.innerHTML = `<tr>
      <th>עובד</th><th>תאריך</th><th>סה"כ שעות</th><th>רגיל 100%</th><th>נוספות 125%</th>
      <th>נוספות 150%</th><th>שבת 150%</th><th>שעות משוקללות</th><th>שכר יום (₪)</th></tr>`;
    for(const r of rows){
      const emp = normEmpName(r.employee);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><span class="link" onclick="openEmployeeCard('${emp.replace(/'/g,"\\'")}')">${emp}</span></td>
        <td>${r.date}</td>
        <td>${fmt2(r.total)}</td>
        <td>${fmt2(r.reg)}</td>
        <td>${fmt2(r.ot125)}</td>
        <td>${fmt2(r.ot150)}</td>
        <td>${fmt2(r.shabbat150)}</td>
        <td>${fmt2(r.weighted)}</td>
        <td>${fmt2(r.pay)}</td>`;
      body.appendChild(tr);
    }
  } else {
    head.innerHTML = `<tr>
      <th>תאריך</th><th>יום</th><th>שעת כניסה</th><th>שעת יציאה</th><th>סה"כ שעות</th></tr>`;
    const rowsPunch = punches
      .filter(p=> normEmpName(p.employee)===selEmp)
      .sort((a,b)=> +a.dtIn - +b.dtIn);
    for(const p of rowsPunch){
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${toISODate(p.dtIn)}</td>
        <td>${wdHe(p.dtIn.getDay())}</td>
        <td>${hhmm(p.dtIn)}</td>
        <td>${hhmm(p.dtOut)}</td>
        <td>${fmt2(hours(p.dtOut - p.dtIn))}</td>`;
      body.appendChild(tr);
    }
  }
}

// ===================== Export CSV (BOM לעברית) =====================
function toCSV(rows, headers){
  const esc = s => '"'+String(s).replace(/"/g,'""')+'"';
  const lines = [headers.map(esc).join(',')];
  for(const r of rows){ lines.push(headers.map(h=> esc(r[h] ?? '')).join(',')); }
  return lines.join('\n');
}
function downloadCSV(name, csv){
  const BOM = '\uFEFF';
  const blob = new Blob([BOM + csv], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); URL.revokeObjectURL(a.href);
}

$('#exportDaily').onclick = ()=>{
  const selEmp = $('#employeeFilter').value || 'ALL';
  if(selEmp==='ALL'){
    const rows = applyOvertime(perDayBase, mapConfigMode(empConfig)).map(r=> ({
      employee:normEmpName(r.employee), date:r.date, total:fmt2(r.total), reg_100:fmt2(r.reg), ot_125:fmt2(r.ot125), ot_150:fmt2(r.ot150),
      shabbat_150:fmt2(r.shabbat150), weighted_hours:fmt2(r.weighted), hourly_rate:(+ensureEmpConfig(r.employee).rate||0),
      pay: fmt2(r.weighted * (+ensureEmpConfig(r.employee).rate||0))
    }));
    const csv = toCSV(rows, ['employee','date','total','reg_100','ot_125','ot_150','shabbat_150','weighted_hours','hourly_rate','pay']);
    downloadCSV('per_day.csv', csv);
  } else {
    const rows = punches
      .filter(p=> normEmpName(p.employee)===selEmp)
      .sort((a,b)=> +a.dtIn - +b.dtIn)
      .map(p=> ({
        date: toISODate(p.dtIn), weekday: wdHe(p.dtIn.getDay()),
        time_in: hhmm(p.dtIn), time_out: hhmm(p.dtOut),
        total_hours: fmt2(hours(p.dtOut - p.dtIn))
      }));
    const csv = toCSV(rows, ['date','weekday','time_in','time_out','total_hours']);
    downloadCSV(`per_day_${selEmp}.csv`, csv);
  }
};

$('#exportSummary').onclick = ()=>{
  const rows = summarizeByEmployee(
    applyOvertime(perDayBase, mapConfigMode(empConfig)).map(r=>({...r, employee:normEmpName(r.employee)})),
    mapConfigRate(empConfig),
    mapConfigExtras(empConfig)
  ).map(r=> ({
    employee:r.employee,total:fmt2(r.total),reg_100:fmt2(r.reg),ot_125:fmt2(r.ot125),ot_150:fmt2(r.ot150),
    shabbat_150:fmt2(r.shabbat150),weighted_hours:fmt2(r.weighted),base_pay:fmt2(r.basePay),
    travel:fmt2(r.travel),tips:fmt2(r.tips),bonus:fmt2(r.bonus),advance:fmt2(r.advance),final_pay:fmt2(r.finalPay)
  }));
  const csv = toCSV(rows, ['employee','total','reg_100','ot_125','ot_150','shabbat_150','weighted_hours','base_pay','travel','tips','bonus','advance','final_pay']);
  downloadCSV('summary.csv', csv);
};

// ===================== JSON Export/Import =====================
function downloadFile(name, content, mime='application/json'){
  const blob = new Blob([content], {type:mime});
  const a = document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; a.click(); URL.revokeObjectURL(a.href);
}
$('#exportJSON').onclick = ()=>{
  try{
    const data = serializeSession();
    downloadFile(`payroll_session_${(new Date).toISOString().slice(0,10)}.json`, JSON.stringify(data, null, 2));
  }catch(e){ alert('שגיאה בייצוא הסשן'); }
};
$('#importJSON').onclick = ()=> $('#importJSONFile').click();
$('#importJSONFile').addEventListener('change', async (ev)=>{
  const f = ev.target.files[0]; if(!f) return;
  try{ const txt = await f.text(); reviveSession(JSON.parse(txt)); }
  catch(e){ alert('שגיאה בטעינת JSON'); } 
  finally { ev.target.value = ''; }
});
$('#clearSession').onclick = ()=>{
  if(confirm('לנקות סשן מקומי?')){ clearLocal(); alert('נוקה הזיכרון המקומי.'); }
};

// ===================== File Handling =====================
$('#file').addEventListener('change', async (ev)=>{
  const file = ev.target.files[0]; if(!file) return;
  try{
    const rows = await readWorkbook(file);
    punches = parsePunches(rows).map(p=> ({...p, employee:normEmpName(p.employee)}));
    if(punches.length===0){ alert('לא נמצאו נתוני משמרות בקובץ.'); return; }
    perDayBase = buildDailyBase(punches);
    for(const e of new Set(perDayBase.map(r=>normEmpName(r.employee)))) ensureEmpConfig(e);
    renderEmployeeTable();
    updateOutputs();
    saveLocal();
  }catch(err){ console.error(err); alert('שגיאה בקריאת הקובץ'); }
});

// ===================== Employee Card (Modal) =====================
let currentEmpInModal = null;
let currentMonthKey = null; // "YYYY-MM"

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
  if(!emp) return;
  ensureEmpConfig(emp);
  currentEmpInModal = emp;
  $('#empModalTitle').textContent = `כרטיס עובד – ${emp}`;

  $('#empMode').value = ensureEmpConfig(emp).mode || 'A';
  $('#empRate').value = +ensureEmpConfig(emp).rate || 0;

  const months = monthsForEmployee(emp);
  const monthSel = $('#empMonthSel'); monthSel.innerHTML='';
  if(months.length===0){ monthSel.innerHTML = '<option value="">—</option>'; currentMonthKey = null; }
  else {
    currentMonthKey = months[months.length-1];
    monthSel.innerHTML = months.map(m=> `<option value="${m}">${m}</option>`).join('');
    monthSel.value = currentMonthKey;
  }

  const ex = ensureEmpConfig(emp).extras;
  $('#extra_travel').value = +ex.travel || 0;
  $('#extra_tips').value = +ex.tips || 0;
  $('#extra_bonus').value = +ex.bonus || 0;
  $('#extra_advance').value = +ex.advance || 0;

  renderEmpDailyPanel(emp, currentMonthKey);
  renderEmpPunchesPanel(emp, currentMonthKey);
  updateModalTotals();
  updateModalFinal();

  $('#empModal').classList.add('show');
  $('#empModal').setAttribute('aria-hidden','false');
}
function closeEmpModal(){
  $('#empModal').classList.remove('show');
  $('#empModal').setAttribute('aria-hidden','true');
  currentEmpInModal = null;
}
window.closeEmpModal = closeEmpModal; // נקרא מה-HTML

function setActiveTab(id){
  ['config','daily','punches'].forEach(k=>{
    $('#tab-'+k).classList.toggle('active', k===id);
    $('#panel-'+k).style.display = (k===id ? '' : 'none');
  });
}
$('#tab-config').onclick = ()=> setActiveTab('config');
$('#tab-daily').onclick = ()=> setActiveTab('daily');
$('#tab-punches').onclick = ()=> setActiveTab('punches');

$('#empMonthSel').addEventListener('change', ()=>{
  currentMonthKey = $('#empMonthSel').value || null;
  if(!currentEmpInModal || !currentMonthKey) return;
  renderEmpDailyPanel(currentEmpInModal, currentMonthKey);
  renderEmpPunchesPanel(currentEmpInModal, currentMonthKey);
  updateModalTotals();
});

$('#empMode').addEventListener('change', ()=>{
  if(!currentEmpInModal) return;
  ensureEmpConfig(currentEmpInModal).mode = $('#empMode').value;
  updateOutputs(); saveLocal();
  if(currentMonthKey){ renderEmpDailyPanel(currentEmpInModal, currentMonthKey); updateModalTotals(); }
});
$('#empRate').addEventListener('input', ()=>{
  if(!currentEmpInModal) return;
  ensureEmpConfig(currentEmpInModal).rate = +$('#empRate').value || 0;
  updateOutputs(); saveLocal();
  if(currentMonthKey){ renderEmpDailyPanel(currentEmpInModal, currentMonthKey); updateModalTotals(); }
});

function readExtrasFromModal(){
  return {
    travel: +($('#extra_travel').value||0),
    tips: +($('#extra_tips').value||0),
    bonus: +($('#extra_bonus').value||0),
    advance: +($('#extra_advance').value||0)
  };
}
$('#saveExtrasBtn').onclick = ()=>{
  if(!currentEmpInModal) return;
  ensureEmpConfig(currentEmpInModal).extras = readExtrasFromModal();
  updateOutputs(); saveLocal();
  updateModalFinal();
  alert('נשמר.');
};

function renderEmpDailyPanel(emp, ym){
  const rows = filterPerDayByMonth(emp, ym);
  const tb = $('#empCardDaily tbody'); tb.innerHTML='';
  for(const r of rows){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.date}</td>
      <td>${fmt2(r.total)}</td>
      <td>${fmt2(r.reg)}</td>
      <td>${fmt2(r.ot125)}</td>
      <td>${fmt2(r.ot150)}</td>
      <td>${fmt2(r.shabbat150)}</td>
      <td>${fmt2(r.weighted)}</td>
      <td>${fmt2(r.pay)}</td>`;
    tb.appendChild(tr);
  }
}
function renderEmpPunchesPanel(emp, ym){
  const sh = filterPunchesByMonth(emp, ym);
  const tb = $('#empPunches tbody'); tb.innerHTML='';
  for(const p of sh.sort((a,b)=> a.dtIn - b.dtIn)){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${toISODate(p.dtIn)}</td>
      <td>${wdHe(p.dtIn.getDay())}</td>
      <td>${hhmm(p.dtIn)}</td>
      <td>${hhmm(p.dtOut)}</td>
      <td>${fmt2(hours(p.dtOut - p.dtIn))}</td>`;
    tb.appendChild(tr);
  }
}
function updateModalTotals(){
  if(!currentEmpInModal || !currentMonthKey){ $('#empWgt').textContent='0'; $('#empBase').textContent='0'; return; }
  const rows = filterPerDayByMonth(currentEmpInModal, currentMonthKey);
  let wsum = 0, base = 0;
  for(const r of rows){ wsum += r.weighted; base += r.pay; }
  $('#empWgt').textContent = fmt2(wsum);
  $('#empBase').textContent = fmt2(base);
}
function updateModalFinal(){
  const base = +($('#empBase').textContent||0);
  const ex = readExtrasFromModal();
  const final = base + (ex.travel||0) + (ex.tips||0) + (ex.bonus||0) - (ex.advance||0);
  $('#empFinal').textContent = fmt2(final);
}
['extra_travel','extra_tips','extra_bonus','extra_advance'].forEach(id=>{
  document.addEventListener('input', (e)=>{ if(e.target && e.target.id===id) updateModalFinal(); });
});

$('#openCardBtn').onclick = ()=>{
  const emp = $('#employeeFilter').value;
  if(emp && emp!=='ALL') openEmployeeCard(emp);
  else alert('נא לבחור עובד מהרשימה.');
};

// ===================== Boot =====================
window.addEventListener('DOMContentLoaded', ()=>{ loadLocalIfAny(); });
