/* app.js – v1.9.9 */

//////////////////////////// Utils ////////////////////////////
const $ = sel => document.querySelector(sel);
const fmt2 = n => (Math.round((+n||0)*100)/100).toFixed(2);
const toISODate = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()).toISOString().slice(0,10);
const dayStart = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const dayEnd   = d => new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
const hours    = ms => ms/36e5;
const two      = n => String(n).padStart(2,'0');
const hhmm     = d => two(d.getHours())+':'+two(d.getMinutes());
const wdHe     = i => ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'][i];
const STORAGE_KEY = 'payrollSessionV7';
function normEmpName(name){ return String(name||'').replace(/\s+/g,' ').trim(); }
function log(s){ try{ console.log(s); }catch(_){} }

//////////////////// Excel date/parse helpers ////////////////////
function excelSerialToDate(n){
  if (typeof n !== 'number' || !isFinite(n) || n < 20000 || n > 60000) return null;
  const epoch = new Date(Date.UTC(1899, 11, 30));
  const ms = Math.round(n * 86400000);
  return new Date(epoch.getTime() + ms);
}
function parseHebDateTime(v){
  if (v == null) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') {
    const d = excelSerialToDate(v);
    return d && !isNaN(d) ? d : null;
  }
  let s = String(v).trim().replace(/\u00A0/g,' ').replace(/\u200f|\u200e/g,'');
  let m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if(m){
    const d=+m[1], mo=+m[2], y=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0);
    const dt = new Date(y,mo-1,d,h,mi,se);
    return isNaN(dt)? null : dt;
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if(m){
    const y=+m[1], mo=+m[2], d=+m[3], h=+m[4], mi=+m[5], se=+(m[6]||0);
    const dt = new Date(y,mo-1,d,h,mi,se);
    return isNaN(dt)? null : dt;
  }
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}

//////////////////// Parsing workbook (CSV/XLSX) ////////////////////
async function readWorkbook(file){
  if (!window.XLSX) { alert('XLSX לא נטען'); throw new Error('XLSX missing'); }
  const ext = file.name.toLowerCase().split('.').pop();

  const reader = new FileReader();
  const load = new Promise(res=> reader.onload = () => res(reader.result));
  if(ext==='csv') reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
  const data = await load;

  const wb = (ext==='csv')
    ? XLSX.read(data, { type:'string', raw:true })
    : XLSX.read(data, { type:'array',  raw:true, cellDates:true });

  const allRows = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:'' });
    for (const row of rows) {
      allRows.push((row || []).map(x => {
        if (typeof x === 'string') return x.replace(/\u00A0/g,' ').replace(/\u200f|\u200e/g,'').trim();
        return x;
      }));
    }
    allRows.push([]); // מפריד בין גליונות
  }
  log(`[payroll] rows loaded: ${allRows.length} [payroll]`);
  return allRows;
}

//////////////////// Parse punches ////////////////////
function parsePunches(rows){
  const norm = s => String(s ?? '')
    .replace(/\u00A0/g,' ')
    .replace(/\u200f|\u200e/g,'')
    .replace(/\s+/g,' ')
    .trim();

  const titleRe = /דו.?ח.*שעות.*עבודה.*עבור\s+(.+?)\s+בין/i;

  function findHeaderIndices(row){
    const cols = row.map(c => norm(c));
    const idxIn  = cols.findIndex(c => /(שעת\s*הגעה|כניסה)/.test(c));
    const idxOut = cols.findIndex(c => /(שעת\s*עזיבה|יציאה)/.test(c));
    const idxDay = cols.findIndex(c => /\bיום\b/.test(c));
    if (idxIn !== -1 && idxOut !== -1) return { idxIn, idxOut, idxDay };
    return null;
  }

  function isStopRow(row){
    const t = norm(row.join(' '));
    const c0 = norm(row[0] || '');
    return (
      t === '' ||
      /^סה"?כ\s*שעות/.test(c0) ||
      /^מספר\s*ימי\s*עבודה/.test(c0) ||
      /אחוזי\s*שכר/.test(t) ||
      titleRe.test(c0)
    );
  }

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

  let currentEmp = null, inTable = false, hdr = null;
  const punches = [], seen = new Set();

  for (let rIdx = 0; rIdx < rows.length; rIdx++){
    const row = rows[rIdx] || [];
    const r0 = norm(row[0] || '');
    if (r0){
      const m = r0.match(titleRe);
      if (m){
        currentEmp = normEmpName(m[1]);
        inTable = false; hdr = null;
        continue;
      }
    }
    if (!inTable){
      const maybeHdr = findHeaderIndices(row);
      if (maybeHdr){
        inTable = true; hdr = maybeHdr;
        if (!currentEmp) currentEmp = lookupEmployeeNear(rIdx) || currentEmp;
        continue;
      }
      continue;
    }
    if (isStopRow(row)){
      inTable = false; hdr = null; currentEmp = null;
      continue;
    }
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
  log(`punches parsed: ${punches.length} [payroll]`);
  return punches;
}

//////////////////// Core + extras + state (ללא שינוי) ////////////////////
// ... [נשאר כמו ב־1.9.8, לא קיצרתי פה כדי לא להציף אותך — זה חלק זהה אצלך] ...

//////////////////// XLSX helpers (RTL + Bold header + zebra + numbers + SUM) ////////////////////
async function exportXlsx(filename, headersHeb, rowsArray) {
  const X = window.XlsxPopulate;
  if (!X) { alert('XlsxPopulate לא נטען'); return; }

  const wb = await X.fromBlankAsync();
  const sheet = wb.sheet(0).name("Export");
  try {
    if (typeof sheet.rightToLeft === 'function') sheet.rightToLeft(true);
    else if (sheet._sheet) {
      sheet._sheet.sheetViews = sheet._sheet.sheetViews || [{ workbookViewId: 0 }];
      sheet._sheet.sheetViews[0].rightToLeft = 1;
    }
  } catch(_) {}

  sheet.cell(1, 1).value([headersHeb]);
  sheet.row(1).style({ bold: true });

  if (rowsArray.length) {
    const dataMatrix = rowsArray.map(row => headersHeb.map(h => row[h]));
    sheet.cell(2, 1).value(dataMatrix);
  }

  const used = sheet.usedRange();
  if (used) used.style({ horizontalAlignment: "right" });

  headersHeb.forEach((h, idx) => {
    let maxLen = (h || '').toString().length;
    for (const r of rowsArray) {
      const v = (r[h] == null ? '' : String(r[h]));
      if (v.length > maxLen) maxLen = v.length;
    }
    const col = sheet.column(idx + 1);
    col.width(Math.min(Math.max(10, maxLen + 2), 40));
    if (idx > 0) col.style({ numberFormat: "0.00" });
  });

  const lastRow = rowsArray.length + 1;
  const LIGHT = "EAF5FF", DARK  = "D6ECFF";
  for (let r = 2; r <= lastRow; r++) {
    sheet.row(r).style({ fill: (r % 2 === 0) ? LIGHT : DARK });
  }

  // SUM row
  const totalRow = lastRow + 1;
  sheet.cell(totalRow, 1).value('סה"כ').style({ bold: true });
  for (let c = 2; c <= headersHeb.length; c++) {
    const colLetter = sheet.column(c).letter();
    const formula = `SUM(${colLetter}2:${colLetter}${lastRow})`;
    sheet.cell(totalRow, c).formula(formula).style({ bold: true, numberFormat: "0.00" });
  }

  const blob = await wb.outputAsync();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename.endsWith('.xlsx') ? filename : (filename + '.xlsx');
  a.click();
  URL.revokeObjectURL(a.href);
}
