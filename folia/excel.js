// Folia — Excel editor. ExcelJS reads/writes .xlsx (with styles), HyperFormula calculates, a light grid edits.
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
let ctx, ready, hf = null, book = null, xid = null, xname = 'Book.xlsx', dirty = false, saveT = null;
let act = 0, sel = { r: 0, c: 0, r2: 0, c2: 0 }, editing = null, rangeMode = false, rowsShown = 0, colsShown = 0, undoS = [], redoS = [], clip = null;

// ---------- libs ----------
const load = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load ' + src)); document.head.appendChild(s); });
const libs = () => ready ||= Promise.all([load('vendor/exceljs.min.js'), load('vendor/hyperformula.full.min.js')]).catch(e => { ready = null; throw e; });
export const excelLibs = libs;

// ---------- helpers ----------
export const colName = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const colNum = s => [...s.toUpperCase()].reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const a1 = (r, c) => colName(c) + (r + 1);
const parseA1 = s => { const m = /^([A-Z]+)(\d+)$/i.exec(s.trim()); return m ? { r: +m[2] - 1, c: colNum(m[1]) } : null; };
const K = (r, c) => r + ':' + c;
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = d => (d.getTime() - EPOCH) / 864e5;
const fromSerial = n => new Date(EPOCH + Math.round(n * 864e5));
const argbToHex = a => typeof a === 'string' && a.length >= 6 ? '#' + a.slice(-6) : null;
const hexToArgb = h => 'FF' + h.replace('#', '').toUpperCase();
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const FMTS = { gen: 'General', num0: '#,##0', num2: '#,##0.00', ngn: '"₦"#,##0.00', usd: '"$"#,##0.00', pct: '0%', pct2: '0.00%', date: 'dd/mm/yyyy' };

// Display a value with an Excel number format (common cases)
export function fmtVal(v, nf) {
  if (v == null || v === '') return '';
  if (typeof v === 'object' && v.value) return String(v.value);           // HF error
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v !== 'number') return String(v);
  nf = nf || 'General';
  const body = nf.split(';')[0];
  if (/General/i.test(body) || !body) { if (Number.isInteger(v)) return String(v); return String(parseFloat(v.toPrecision(10))); }
  const isDate = /(^|[^"\\])[dy]|mmm|(^|[^"])m\/|\/m|h:mm/i.test(body.replace(/"[^"]*"/g, '')) && !/0\.0|#/.test(body);
  if (isDate) {
    const d = fromSerial(v), dd = String(d.getUTCDate()).padStart(2, '0'), mm = String(d.getUTCMonth() + 1).padStart(2, '0'), yy = d.getUTCFullYear();
    let out = /mmm/i.test(body) ? `${d.getUTCDate()} ${MON[d.getUTCMonth()]} ${yy}` : /^m/i.test(body) ? `${mm}/${dd}/${yy}` : /yyyy-mm-dd/i.test(body) ? `${yy}-${mm}-${dd}` : `${dd}/${mm}/${yy}`;
    if (/h/i.test(body)) out += ' ' + String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0');
    if (!/[dy]/i.test(body.replace(/mmm/i, ''))) out = String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0');
    return out;
  }
  const pct = /%/.test(body); let x = pct ? v * 100 : v;
  const dm = /\.([0#]+)/.exec(body); const dec = dm ? dm[1].length : 0;
  const sym = (/"([^"]+)"/.exec(body) || /\[\$([^\]-]+)/.exec(body) || /([₦$€£])/.exec(body) || [])[1] || '';
  const neg = x < 0; x = Math.abs(x);
  let s = x.toFixed(dec); if (/#,##|,0/.test(body)) { const [i, f] = s.split('.'); s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? '.' + f : ''); }
  const symFirst = sym && body.indexOf(sym) < body.search(/[0#]/);
  return (neg ? '-' : '') + (symFirst ? sym : '') + s + (pct ? '%' : '') + (sym && !symFirst ? ' ' + sym.trim() : '');
}
// Turn what the user typed into a raw value (+ a format it implies)
export function parseInput(t) {
  const s = String(t).trim();
  if (s === '') return { v: null };
  if (s[0] === '=') return { v: s };
  if (/^(true|false)$/i.test(s)) return { v: /^t/i.test(s) };
  let m = /^(-?)([₦$€£]?)\s*(-?[\d,]*\.?\d+)\s*(%?)$/.exec(s);
  if (m && (!/,/.test(m[3]) || /^-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(m[3]))) {
    let n = parseFloat(m[3].replace(/,/g, '')) * (m[1] ? -1 : 1); let nf = null;
    if (m[4]) { n /= 100; nf = /\./.test(m[3]) ? FMTS.pct2 : FMTS.pct; }
    else if (m[2]) nf = m[2] === '₦' ? FMTS.ngn : m[2] === '$' ? FMTS.usd : `"${m[2]}"#,##0.00`;
    else if (/,/.test(m[3])) nf = /\./.test(m[3]) ? FMTS.num2 : FMTS.num0;
    return { v: n, nf };
  }
  m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s);
  if (m && +m[2] <= 12 && +m[1] <= 31) return { v: toSerial(new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]))), nf: FMTS.date };
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) return { v: toSerial(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]))), nf: 'yyyy-mm-dd' };
  return { v: s };
}
function parseCSV(text) {
  const rows = []; let row = [], f = '', q = false;
  const sep = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
    else if (ch === '"') q = true; else if (ch === sep) { row.push(f); f = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
    else f += ch;
  }
  if (f !== '' || row.length) { row.push(f); rows.push(row); }
  return rows;
}

// ---------- book model: content lives in HyperFormula, styles here ----------
// book = { sheets: [{ name, st: {"r:c": {b,i,u,al,fill,color,nf}}, cw: {c: px}, merges: [] }] }
function cellRaw(c) {
  let v = c.value;
  if (v == null) return null;
  if (c.formula) return '=' + c.formula;
  if (v instanceof Date) return toSerial(v);
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.text != null) return String(v.text);
    if (v.error) return v.error;
    if (v.result != null) return v.result instanceof Date ? toSerial(v.result) : v.result;
    return null;
  }
  return v;
}
export async function importXlsx(bytes) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(bytes.slice().buffer);
  const sheets = [], data = {};
  for (const ws of wb.worksheets) {
    if (ws.state === 'veryHidden') continue;
    const name = ws.name, st = {}, cw = {}, grid = [];
    ws.eachRow({ includeEmpty: false }, (row, rn) => row.eachCell({ includeEmpty: false }, (c, cn) => {
      if (c.isMerged && c.master !== c) return;
      const r = rn - 1, cc = cn - 1; (grid[r] ||= [])[cc] = cellRaw(c);
      const s = {}; const f = c.font || {}, fl = c.fill, al = c.alignment || {};
      if (f.bold) s.b = 1; if (f.italic) s.i = 1; if (f.underline) s.u = 1;
      const col = argbToHex(f.color?.argb); if (col && col !== '#000000') s.color = col;
      if (fl?.type === 'pattern' && fl.pattern === 'solid') { const h = argbToHex(fl.fgColor?.argb); if (h) s.fill = h; }
      if (['left', 'center', 'right'].includes(al.horizontal)) s.al = al.horizontal;
      if (c.numFmt && c.numFmt !== 'General') s.nf = c.numFmt;
      if (c.value instanceof Date && !s.nf) s.nf = FMTS.date;
      if (Object.keys(s).length) st[K(r, cc)] = s;
    }));
    (ws.columns || []).forEach((col, i) => { if (col && col.width) cw[i] = Math.round(Math.min(60, col.width) * 7 + 5); });
    for (let r = 0; r < grid.length; r++) { grid[r] ||= []; for (let c = 0; c < grid[r].length; c++) if (grid[r][c] === undefined) grid[r][c] = null; }
    data[name] = grid.length ? grid : [[]];
    sheets.push({ name, st, cw, merges: (ws.model?.merges || []).slice() });
  }
  if (!sheets.length) { sheets.push({ name: 'Sheet1', st: {}, cw: {}, merges: [] }); data.Sheet1 = [[]]; }
  return { sheets, data };
}
export function importCSV(bytes, name) {
  const text = new TextDecoder().decode(bytes).replace(/^\uFEFF/, '');
  const st = {}; const grid = parseCSV(text).map((row, r) => row.map((t, c) => { const p = parseInput(t); if (p.nf) st[K(r, c)] = { nf: p.nf }; return p.v; }));
  const sn = (name.replace(/\.[^.]+$/, '') || 'Sheet1').slice(0, 31).replace(/[\[\]*?\/\\:]/g, '');
  return { sheets: [{ name: sn, st, cw: {}, merges: [] }], data: { [sn]: grid.length ? grid : [[]] } };
}
function makeHF(data) { return HyperFormula.buildFromSheets(data, { licenseKey: 'gpl-v3', maxRows: 100000, maxColumns: 2000 }); }
const sid = i => hf.getSheetId(book.sheets[i].name);
export async function exportXlsx(h = hf, b = book) {
  const wb = new ExcelJS.Workbook(); wb.creator = 'Folia';
  b.sheets.forEach(sh => {
    const id = h.getSheetId(sh.name), ws = wb.addWorksheet(sh.name);
    const raw = h.getSheetSerialized(id), val = h.getSheetValues(id);
    raw.forEach((row, r) => row.forEach((v, c) => {
      if (v == null || v === '') return; const cell = ws.getCell(r + 1, c + 1);
      if (typeof v === 'string' && v[0] === '=') { let res = val[r]?.[c]; if (res && typeof res === 'object') res = { error: res.value }; cell.value = { formula: v.slice(1), result: res ?? undefined }; }
      else cell.value = v;
    }));
    Object.entries(sh.st).forEach(([k, s]) => {
      const [r, c] = k.split(':').map(Number), cell = ws.getCell(r + 1, c + 1);
      if (s.b || s.i || s.u || s.color) cell.font = { name: 'Calibri', size: 11, bold: !!s.b, italic: !!s.i, underline: !!s.u, ...(s.color ? { color: { argb: hexToArgb(s.color) } } : {}) };
      if (s.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(s.fill) } };
      if (s.al) cell.alignment = { horizontal: s.al };
      if (s.nf) cell.numFmt = s.nf;
    });
    Object.entries(sh.cw).forEach(([c, px]) => { ws.getColumn(+c + 1).width = Math.max(2, (px - 5) / 7); });
    sh.merges.forEach(m => { try { ws.mergeCells(m); } catch {} });
  });
  return new Uint8Array(await wb.xlsx.writeBuffer());
}
function exportCSV() {
  const id = sid(act), val = hf.getSheetValues(id), st = book.sheets[act].st;
  return new TextEncoder().encode('\uFEFF' + val.map((row, r) => row.map((v, c) => {
    let t = v == null ? '' : typeof v === 'object' ? v.value : typeof v === 'number' && st[K(r, c)]?.nf && /[dy]/i.test(st[K(r, c)].nf) ? fmtVal(v, st[K(r, c)].nf) : String(v);
    return /[",\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; }).join(',')).join('\r\n'));
}
export function buildBook(imp) { book = { sheets: imp.sheets }; hf?.destroy?.(); hf = makeHF(imp.data); return { hf, book }; }

// ---------- undo ----------
function snap() { undoS.push(JSON.stringify({ data: hf.getAllSheetsSerialized(), sheets: book.sheets, act })); if (undoS.length > 40) undoS.shift(); redoS = []; }
function restoreSnap(s) { const o = JSON.parse(s); hf.destroy(); hf = makeHF(o.data); book.sheets = o.sheets; act = Math.min(o.act, book.sheets.length - 1); renderAll(); changed(); }
function undo() { if (!undoS.length) return ctx.toast('Nothing to undo'); redoS.push(JSON.stringify({ data: hf.getAllSheetsSerialized(), sheets: book.sheets, act })); restoreSnap(undoS.pop()); }
function redo() { if (!redoS.length) return; undoS.push(JSON.stringify({ data: hf.getAllSheetsSerialized(), sheets: book.sheets, act })); restoreSnap(redoS.pop()); }

// ---------- grid ----------
const grid = () => $('#xGrid');
const cellEl = (r, c) => grid().tBodies[0]?.rows[r]?.cells[c + 1];
const used = () => { const d = hf.getSheetDimensions(sid(act)); return { rows: d.height, cols: d.width }; };
function styleCell(td, s, v) {
  td.style.fontWeight = s?.b ? '700' : ''; td.style.fontStyle = s?.i ? 'italic' : ''; td.style.textDecoration = s?.u ? 'underline' : '';
  td.style.color = s?.color || ''; td.style.background = s?.fill || '';
  td.style.textAlign = s?.al || (typeof v === 'number' ? 'right' : typeof v === 'boolean' || (v && typeof v === 'object') ? 'center' : '');
}
function paintCell(r, c, vals) {
  const td = cellEl(r, c); if (!td) return; const v = vals ? vals[r]?.[c] : hf.getCellValue({ sheet: sid(act), row: r, col: c });
  const s = book.sheets[act].st[K(r, c)]; const t = fmtVal(v, s?.nf); if (td.textContent !== t) td.textContent = t; styleCell(td, s, v);
  td.classList.toggle('err', !!(v && typeof v === 'object'));
}
function paintAll() { const vals = hf.getSheetValues(sid(act)); for (let r = 0; r < rowsShown; r++) for (let c = 0; c < colsShown; c++) { const td = cellEl(r, c); if (!td) continue; if (vals[r]?.[c] != null || td.textContent || book.sheets[act].st[K(r, c)] || td.style.background) paintCell(r, c, vals); } }
function renderAll() {
  const u = used(); rowsShown = Math.max(u.rows + 40, 100); colsShown = Math.max(u.cols + 4, 26);
  const sh = book.sheets[act], g = grid();
  let h = '<colgroup><col style="width:44px">'; for (let c = 0; c < colsShown; c++) h += `<col style="width:${sh.cw[c] || 96}px">`; h += '</colgroup><thead><tr><th class="corner"></th>';
  for (let c = 0; c < colsShown; c++) h += `<th data-c="${c}">${colName(c)}<i class="rz" data-rz="${c}"></i></th>`; h += '</tr></thead><tbody>';
  const empty = '<td></td>'.repeat(colsShown); for (let r = 0; r < rowsShown; r++) h += `<tr><th data-r="${r}">${r + 1}</th>${empty}</tr>`;
  g.innerHTML = h + '</tbody>'; g.style.width = (44 + Array.from({ length: colsShown }, (_, c) => sh.cw[c] || 96).reduce((a, b) => a + b, 0)) + 'px';
  paintAll(); renderTabs(); showSel();
}
function growRows(n = 100) {
  const tb = grid().tBodies[0], empty = '<td></td>'.repeat(colsShown); let h = '';
  for (let r = rowsShown; r < rowsShown + n; r++) h += `<tr><th data-r="${r}">${r + 1}</th>${empty}</tr>`;
  tb.insertAdjacentHTML('beforeend', h); rowsShown += n;
}
function renderTabs() {
  const t = $('#xTabs'); t.innerHTML = '';
  book.sheets.forEach((s, i) => { const b = document.createElement('button'); b.className = 'xtab' + (i === act ? ' on' : ''); b.textContent = s.name; b.onclick = () => { if (i === act) sheetMenu(); else { commit(); act = i; sel = { r: 0, c: 0, r2: 0, c2: 0 }; renderAll(); } }; t.appendChild(b); });
  const add = document.createElement('button'); add.className = 'xtab add'; add.textContent = '+'; add.title = 'Add sheet';
  add.onclick = () => { commit(); snap(); let n = book.sheets.length + 1; while (hf.doesSheetExist('Sheet' + n)) n++; const nm = 'Sheet' + n; hf.addSheet(nm); book.sheets.push({ name: nm, st: {}, cw: {}, merges: [] }); act = book.sheets.length - 1; sel = { r: 0, c: 0, r2: 0, c2: 0 }; renderAll(); changed(); };
  t.appendChild(add);
}
function sheetMenu() { $('#xSheetName').value = book.sheets[act].name; $('#xSheetSheet').classList.remove('hidden'); }

// ---------- selection ----------
const norm = () => ({ r1: Math.min(sel.r, sel.r2), r2: Math.max(sel.r, sel.r2), c1: Math.min(sel.c, sel.c2), c2: Math.max(sel.c, sel.c2) });
function showSel() {
  $$('#xGrid td.in,#xGrid td.cur').forEach(td => td.classList.remove('in', 'cur')); $$('#xGrid th.hl').forEach(th => th.classList.remove('hl'));
  const n = norm(); if ((n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1) <= 20000) for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) cellEl(r, c)?.classList.add('in');
  cellEl(sel.r, sel.c)?.classList.add('cur');
  for (let c = n.c1; c <= n.c2; c++) grid().tHead.rows[0].cells[c + 1]?.classList.add('hl');
  for (let r = n.r1; r <= n.r2; r++) grid().tBodies[0].rows[r]?.cells[0].classList.add('hl');
  $('#xRef').value = n.r1 === n.r2 && n.c1 === n.c2 ? a1(sel.r, sel.c) : a1(n.r1, n.c1) + ':' + a1(n.r2, n.c2);
  if (!editing) { const raw = hf.getCellSerialized({ sheet: sid(act), row: sel.r, col: sel.c }); $('#xFx').value = rawText(raw, book.sheets[act].st[K(sel.r, sel.c)]?.nf); }
  const s = book.sheets[act].st[K(sel.r, sel.c)] || {};
  $('#xB').classList.toggle('on', !!s.b); $('#xI').classList.toggle('on', !!s.i); $('#xU').classList.toggle('on', !!s.u);
  $$('#xAlign button').forEach(b => b.classList.toggle('on', b.dataset.al === s.al));
  const fk = Object.keys(FMTS).find(k => FMTS[k] === s.nf); $('#xFmt').value = fk || (s.nf ? '' : 'gen');
  rangeSum(n);
}
function rangeSum(n) {
  const el = $('#xSum'); if (n.r1 === n.r2 && n.c1 === n.c2) { el.textContent = ''; return; }
  let sum = 0, cnt = 0, num = 0; const id = sid(act);
  if ((n.r2 - n.r1 + 1) * (n.c2 - n.c1 + 1) > 50000) { el.textContent = ''; return; }
  for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) { const v = hf.getCellValue({ sheet: id, row: r, col: c }); if (v != null && v !== '') cnt++; if (typeof v === 'number') { sum += v; num++; } }
  el.textContent = num ? `Sum ${fmtVal(sum, 'General')} · Avg ${fmtVal(sum / num, 'General')} · Count ${cnt}` : cnt ? `Count ${cnt}` : '';
}
function rawText(raw, nf) {
  if (raw == null) return ''; if (typeof raw === 'number') { if (nf && /%/.test(nf)) return parseFloat((raw * 100).toPrecision(12)) + '%'; if (nf && fmtVal(raw, nf).match(/^\d{2}\/\d{2}\/\d{4}|^\d{4}-/)) return fmtVal(raw, nf); return String(raw); }
  if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE'; return String(raw);
}
function moveTo(r, c, extend) {
  r = Math.max(0, r); c = Math.max(0, Math.min(colsShown - 1, c)); while (r >= rowsShown - 5) growRows();
  if (extend) { sel.r2 = r; sel.c2 = c; } else sel = { r, c, r2: r, c2: c };
  showSel(); const td = cellEl(extend ? sel.r2 : r, extend ? sel.c2 : c); td?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

// ---------- editing ----------
const gridFocus = () => { const w = $('#xWrap'); if (document.activeElement !== w) w.focus({ preventScroll: true }); };
function startEdit(initial) {
  if (editing) return; const td = cellEl(sel.r, sel.c); if (!td) return;
  const inp = $('#xIn'), wrap = $('#xWrap'); const tr = td.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
  inp.style.left = (tr.left - wr.left + wrap.scrollLeft) + 'px'; inp.style.top = (tr.top - wr.top + wrap.scrollTop) + 'px';
  inp.style.minWidth = tr.width + 'px'; inp.style.height = tr.height + 'px';
  const s = book.sheets[act].st[K(sel.r, sel.c)]; inp.style.fontWeight = s?.b ? '700' : ''; inp.style.textAlign = s?.al || 'left';
  inp.value = initial != null ? initial : $('#xFx').value; inp.classList.remove('hidden');
  editing = { r: sel.r, c: sel.c, src: 'cell' }; inp.focus({ preventScroll: true }); const L = inp.value.length; try { inp.setSelectionRange(L, L); } catch {}
}
function setCell(r, c, text, sheetIdx = act) {
  const p = parseInput(text), sh = book.sheets[sheetIdx];
  hf.setCellContents({ sheet: sid(sheetIdx), row: r, col: c }, [[p.v]]);
  if (p.nf && !(sh.st[K(r, c)]?.nf)) sh.st[K(r, c)] = { ...(sh.st[K(r, c)] || {}), nf: p.nf };
}
function commit(move) {
  if (!editing) return false; const e = editing; editing = null;
  const text = e.src === 'fx' ? $('#xFx').value : $('#xIn').value; if (e.src === 'cell') gridFocus(); $('#xIn').classList.add('hidden');
  const before = hf.getCellSerialized({ sheet: sid(act), row: e.r, col: e.c });
  if (rawText(before, book.sheets[act].st[K(e.r, e.c)]?.nf) !== text) {
    snap();
    try { setCell(e.r, e.c, text); } catch (er) { ctx.toast('That formula has a problem.'); }
    paintAll(); changed();
  }
  if (move === 'down') moveTo(e.r + 1, e.c); else if (move === 'right') moveTo(e.r, e.c + 1); else if (move === 'left') moveTo(e.r, e.c - 1); else showSel();
  return true;
}
function cancel() { if (!editing) return; if (editing.src === 'cell') gridFocus(); editing = null; $('#xIn').classList.add('hidden'); showSel(); }
function clearSel() { const n = norm(), id = sid(act); snap(); for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) hf.setCellContents({ sheet: id, row: r, col: c }, [[null]]); paintAll(); showSel(); changed(); }

// ---------- formatting ----------
function eachSel(fn) { const n = norm(), st = book.sheets[act].st; for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) { const k = K(r, c); const s = { ...(st[k] || {}) }; fn(s, r, c); Object.keys(s).forEach(x => (s[x] == null || s[x] === 0 || s[x] === '') && delete s[x]); if (Object.keys(s).length) st[k] = s; else delete st[k]; } }
function fmt(fn) { commit(); snap(); eachSel(fn); paintAll(); showSel(); changed(); }
function toggle(p) { const on = !(book.sheets[act].st[K(sel.r, sel.c)] || {})[p]; fmt(s => { s[p] = on ? 1 : 0; }); }
function autoSum() {
  commit(); const n = norm(), id = sid(act);
  if (n.r1 !== n.r2 || n.c1 !== n.c2) { // sum each column of the range into the row below
    snap(); for (let c = n.c1; c <= n.c2; c++) hf.setCellContents({ sheet: id, row: n.r2 + 1, col: c }, [[`=SUM(${a1(n.r1, c)}:${a1(n.r2, c)})`]]);
    paintAll(); moveTo(n.r2 + 1, n.c1); changed(); return;
  }
  let r = sel.r - 1; while (r >= 0 && typeof hf.getCellValue({ sheet: id, row: r, col: sel.c }) !== 'number') r--; const end = r;
  while (r >= 0 && typeof hf.getCellValue({ sheet: id, row: r, col: sel.c }) === 'number') r--;
  const f = end >= 0 ? `=SUM(${a1(r + 1, sel.c)}:${a1(end, sel.c)})` : '=SUM()';
  startEdit(f);
}
function rowsCols(op) {
  commit(); snap(); const n = norm(), id = sid(act), sh = book.sheets[act];
  const shift = (axis, at, d) => { const st = {}; Object.entries(sh.st).forEach(([k, s]) => { let [r, c] = k.split(':').map(Number); const v = axis === 'r' ? r : c;
      if (d < 0 && v >= at && v < at - d) return; const nv = v >= at ? v + d : v; if (axis === 'r') r = nv; else c = nv; st[K(r, c)] = s; }); sh.st = st;
    if (axis === 'c') { const cw = {}; Object.entries(sh.cw).forEach(([c, w]) => { c = +c; if (d < 0 && c >= at && c < at - d) return; cw[c >= at ? c + d : c] = w; }); sh.cw = cw; } };
  try {
    if (op === 'addRow') { hf.addRows(id, [n.r1, 1]); shift('r', n.r1, 1); }
    if (op === 'addCol') { hf.addColumns(id, [n.c1, 1]); shift('c', n.c1, 1); }
    if (op === 'delRow') { const k = n.r2 - n.r1 + 1; hf.removeRows(id, [n.r1, k]); shift('r', n.r1, -k); }
    if (op === 'delCol') { const k = n.c2 - n.c1 + 1; hf.removeColumns(id, [n.c1, k]); shift('c', n.c1, -k); }
  } catch (e) { console.error(e); ctx.toast("Couldn't do that here."); }
  renderAll(); changed();
}
function sortSel(desc) {
  commit(); const n = norm(), id = sid(act); if (n.r1 === n.r2) { // whole used block of rows, by this column
    const u = used(); n.r1 = 0; n.r2 = u.rows - 1; n.c1 = 0; n.c2 = u.cols - 1;
    const first = hf.getCellValue({ sheet: id, row: 0, col: sel.c }), second = hf.getCellValue({ sheet: id, row: 1, col: sel.c });
    if (typeof first === 'string' && typeof second !== 'string') n.r1 = 1; // keep a header row
  }
  if (n.r2 <= n.r1) return; snap();
  const rows = []; for (let r = n.r1; r <= n.r2; r++) { const raw = [], vals = [], st = []; for (let c = n.c1; c <= n.c2; c++) { raw.push(hf.getCellSerialized({ sheet: id, row: r, col: c })); vals.push(hf.getCellValue({ sheet: id, row: r, col: c })); st.push(book.sheets[act].st[K(r, c)]); } rows.push({ raw, vals, st }); }
  const kc = sel.c - n.c1; const key = x => x.vals[kc];
  rows.sort((a, b) => { const x = key(a), y = key(b); if (x == null || x === '') return 1; if (y == null || y === '') return -1; const o = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true }); return desc ? -o : o; });
  rows.forEach((row, i) => { const r = n.r1 + i; row.raw.forEach((v, j) => { const c = n.c1 + j; const isF = typeof v === 'string' && v[0] === '='; hf.setCellContents({ sheet: id, row: r, col: c }, [[isF ? row.vals[j] : v]]); if (row.st[j]) book.sheets[act].st[K(r, c)] = row.st[j]; else delete book.sheets[act].st[K(r, c)]; }); });
  paintAll(); showSel(); changed();
}

// ---------- copy / paste ----------
function selTSV() { const n = norm(), id = sid(act); const out = []; for (let r = n.r1; r <= n.r2; r++) { const row = []; for (let c = n.c1; c <= n.c2; c++) row.push(fmtVal(hf.getCellValue({ sheet: id, row: r, col: c }), book.sheets[act].st[K(r, c)]?.nf)); out.push(row.join('\t')); } return out.join('\n'); }
function onCopy(e, cut) {
  if (editing || $('#xlEd').classList.contains('hidden') || document.activeElement?.matches('input,textarea,select')) return;
  const n = norm(), t = selTSV(); e.clipboardData.setData('text/plain', t); e.preventDefault();
  hf.copy({ start: { sheet: sid(act), row: n.r1, col: n.c1 }, end: { sheet: sid(act), row: n.r2, col: n.c2 } });
  const st = []; for (let r = n.r1; r <= n.r2; r++) { const row = []; for (let c = n.c1; c <= n.c2; c++) row.push(book.sheets[act].st[K(r, c)]); st.push(row); } clip = { t, st };
  if (cut) clearSel();
}
function onPaste(e) {
  if (editing || $('#xlEd').classList.contains('hidden') || document.activeElement?.matches('input,textarea,select')) return;
  const t = e.clipboardData.getData('text/plain'); if (t == null) return; e.preventDefault(); snap(); const id = sid(act);
  if (clip && clip.t === t && hf.isClipboardEmpty() === false) { hf.paste({ sheet: id, row: sel.r, col: sel.c }); clip.st.forEach((row, i) => row.forEach((s, j) => { const k = K(sel.r + i, sel.c + j); if (s) book.sheets[act].st[k] = s; else delete book.sheets[act].st[k]; })); }
  else t.replace(/\r/g, '').replace(/\n$/, '').split('\n').forEach((line, i) => line.split('\t').forEach((cell, j) => setCell(sel.r + i, sel.c + j, cell)));
  renderAll(); changed();
}

// ---------- save ----------
async function autosave() {
  clearTimeout(saveT); saveT = null; if (!dirty || !xid) return;
  try { const bytes = await exportXlsx(); dirty = false; $('#xState').textContent = 'Saved'; await ctx.touchRecent(xid, xname, bytes, 'xlsx'); }
  catch (e) { console.error(e); $('#xState').textContent = 'Not saved'; }
}
function changed() { dirty = true; $('#xState').textContent = 'Editing…'; clearTimeout(saveT); saveT = setTimeout(autosave, 2500); }
function printSheet() {
  const id = sid(act), val = hf.getSheetValues(id), st = book.sheets[act].st, u = used();
  let h = '<table>'; for (let r = 0; r < u.rows; r++) { h += '<tr>'; for (let c = 0; c < u.cols; c++) { const s = st[K(r, c)], v = val[r]?.[c]; const css = [s?.b && 'font-weight:700', s?.i && 'font-style:italic', s?.fill && 'background:' + s.fill, s?.color && 'color:' + s.color, 'text-align:' + (s?.al || (typeof v === 'number' ? 'right' : 'left'))].filter(Boolean).join(';'); h += `<td style="${css}">${fmtVal(v, s?.nf).replace(/&/g, '&amp;').replace(/</g, '&lt;')}</td>`; } h += '</tr>'; }
  $('#xPrint').innerHTML = h + '</table>';
  const t = document.title; document.title = xname.replace(/\.xlsx$/i, ''); document.body.classList.add('printing-x');
  const done = () => { document.body.classList.remove('printing-x'); document.title = t; $('#xPrint').innerHTML = ''; }; window.addEventListener('afterprint', done, { once: true }); window.print(); setTimeout(done, 60000);
}
async function exportAs(kind) {
  commit();
  if (kind === 'pdf') return printSheet();
  ctx.busy(true, 'Saving…');
  try {
    if (kind === 'csv') { ctx.busy(false); await ctx.deliver(exportCSV(), xname.replace(/\.xlsx$/i, '') + (book.sheets.length > 1 ? ' - ' + book.sheets[act].name : '') + '.csv', 'text/csv'); return; }
    const bytes = await exportXlsx(); ctx.busy(false); const ok = await ctx.deliver(bytes, xname, XLSX_MIME);
    if (ok) { dirty = false; ctx.touchRecent(xid, xname, bytes, 'xlsx'); $('#xState').textContent = 'Saved'; }
  } catch (e) { console.error(e); ctx.busy(false); ctx.toast("Couldn't save: " + (e.message || e)); }
}

// ---------- open ----------
function show(name, id) {
  xname = name.replace(/\.(csv|xlsx|xlsm)$/i, '') + '.xlsx'; xid = id; dirty = false; undoS = []; redoS = []; act = 0; sel = { r: 0, c: 0, r2: 0, c2: 0 }; editing = null;
  $('#xName').value = xname.replace(/\.xlsx$/i, ''); $('#xState').textContent = 'Saved';
  $('#home').classList.add('hidden'); $('#editor').classList.add('hidden'); $('#wordEd').classList.add('hidden'); $('#xlEd').classList.remove('hidden');
  renderAll(); $('#xWrap').scrollTo(0, 0); if (matchMedia('(pointer:fine)').matches) gridFocus();
}
export async function openExcel(file, id) {
  ctx.busy(true, 'Opening…');
  try {
    await libs(); const csv = /\.csv$/i.test(file.name);
    buildBook(csv ? importCSV(file.bytes, file.name) : await importXlsx(file.bytes));
    const nid = id || Math.random().toString(36).slice(2, 10); show(file.name, nid);
    if (!id) { if (csv) { dirty = true; autosave(); } else ctx.touchRecent(nid, xname, file.bytes, 'xlsx'); }
  } catch (e) { console.error(e); ctx.toast("Couldn't open that spreadsheet."); }
  ctx.busy(false);
}
export async function newExcel() {
  ctx.busy(true, 'Starting…');
  try { await libs(); buildBook({ sheets: [{ name: 'Sheet1', st: {}, cw: {}, merges: [] }], data: { Sheet1: [[]] } }); show('Untitled.xlsx', Math.random().toString(36).slice(2, 10)); dirty = true; await autosave(); }
  catch (e) { console.error(e); ctx.toast("Couldn't start a spreadsheet."); }
  ctx.busy(false);
}

// ---------- wiring ----------
export function initExcel(c) {
  ctx = c; const wrap = $('#xWrap'), g = $('#xGrid');
  let drag = null;
  g.addEventListener('pointerdown', e => {
    const rz = e.target.closest('[data-rz]'); if (rz) { e.preventDefault(); const ci = +rz.dataset.rz, col = g.querySelectorAll('col')[ci + 1], x0 = e.clientX, w0 = col.getBoundingClientRect().width; snap();
      const mv = ev => { const w = Math.max(30, Math.round(w0 + ev.clientX - x0)); col.style.width = w + 'px'; book.sheets[act].cw[ci] = w; };
      const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); renderAll(); changed(); }; window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up); return; }
    const td = e.target.closest('td'), th = e.target.closest('th');
    if (td) { const r = td.parentNode.rowIndex - 1, cc = td.cellIndex - 1; const wasCur = !rangeMode && r === sel.r && cc === sel.c && sel.r === sel.r2 && sel.c === sel.c2;
      if (editing) commit();
      if (e.pointerType !== 'mouse') { drag = { touch: true, r, cc, wasCur }; return; }   // decide on pointerup so scrolling still works
      gridFocus(); if (e.shiftKey) moveTo(r, cc, true); else { moveTo(r, cc); drag = { r, cc }; } return; }
    if (th && th.dataset.c != null) { commit(); const cc = +th.dataset.c; sel = { r: 0, c: cc, r2: Math.max(used().rows - 1, rowsShown - 1), c2: cc }; showSel(); }
    else if (th && th.dataset.r != null) { commit(); const r = +th.dataset.r; sel = { r, c: 0, r2: r, c2: colsShown - 1 }; showSel(); }
  });
  g.addEventListener('pointermove', e => { if (!drag || drag.touch || e.buttons !== 1) return; const td = document.elementFromPoint(e.clientX, e.clientY)?.closest?.('#xGrid td'); if (td) { const r = td.parentNode.rowIndex - 1, cc = td.cellIndex - 1; if (r !== sel.r2 || cc !== sel.c2) { sel.r2 = r; sel.c2 = cc; showSel(); } } });
  g.addEventListener('pointerup', e => { const d = drag; drag = null; if (!d?.touch) return; if (Math.abs(wrap._sx - wrap.scrollLeft) > 4 || Math.abs(wrap._sy - wrap.scrollTop) > 4) return;
    if (rangeMode) moveTo(d.r, d.cc, true); else if (d.wasCur) startEdit(); else moveTo(d.r, d.cc); });
  g.addEventListener('pointerdown', () => { wrap._sx = wrap.scrollLeft; wrap._sy = wrap.scrollTop; }, true);
  g.addEventListener('dblclick', e => { if (e.target.closest('td')) startEdit(); });
  wrap.addEventListener('scroll', () => { if (wrap.scrollTop + wrap.clientHeight > wrap.scrollHeight - 300 && rowsShown < 100000) growRows(); });
  const inp = $('#xIn');
  inp.addEventListener('keydown', e => {
    if (['Enter', 'Tab', 'Escape'].includes(e.key)) e.stopPropagation();
    if (e.key === 'Enter') { e.preventDefault(); commit(e.shiftKey ? null : 'down'); }
    else if (e.key === 'Tab') { e.preventDefault(); commit(e.shiftKey ? 'left' : 'right'); }
    else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
  });
  inp.addEventListener('input', () => { $('#xFx').value = inp.value; });
  inp.addEventListener('blur', () => setTimeout(() => { if (editing?.src === 'cell' && document.activeElement !== inp) commit(); }, 0));
  const fx = $('#xFx');
  fx.addEventListener('focus', () => { if (editing?.src !== 'fx') { if (editing) commit(); editing = { r: sel.r, c: sel.c, src: 'fx' }; } });
  fx.addEventListener('keydown', e => { if (['Enter', 'Escape'].includes(e.key)) e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); commit('down'); fx.blur(); } else if (e.key === 'Escape') { cancel(); fx.blur(); } });
  fx.addEventListener('blur', () => { if (editing?.src === 'fx') commit(); });
  $('#xRef').addEventListener('change', e => { const v = e.target.value.trim().toUpperCase(), [p, q] = v.split(':').map(parseA1); if (p) { moveTo(p.r, p.c); if (q) moveTo(q.r, q.c, true); } else showSel(); });
  document.addEventListener('keydown', e => {
    if ($('#xlEd').classList.contains('hidden') || editing || (document.activeElement?.matches('input,textarea,select') && document.activeElement.id !== 'xIn') || !$('#xSheetSheet').classList.contains('hidden') || !$('#xSaveSheet').classList.contains('hidden')) return;
    const m = e.ctrlKey || e.metaKey, ex = e.shiftKey, cr = ex ? sel.r2 : sel.r, cc = ex ? sel.c2 : sel.c;
    if (m && /^[zZ]$/.test(e.key)) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (m && e.key === 'y') { e.preventDefault(); redo(); return; }
    if (m && e.key === 'b') { e.preventDefault(); toggle('b'); return; } if (m && e.key === 'i') { e.preventDefault(); toggle('i'); return; } if (m && e.key === 'u') { e.preventDefault(); toggle('u'); return; }
    if (m) return;
    const mv = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (mv) { e.preventDefault(); moveTo(cr + mv[0], cc + mv[1], ex); return; }
    if (e.key === 'Enter') { e.preventDefault(); moveTo(sel.r + (ex ? -1 : 1), sel.c); return; }
    if (e.key === 'Tab') { e.preventDefault(); moveTo(sel.r, sel.c + (ex ? -1 : 1)); return; }
    if (e.key === 'F2') { e.preventDefault(); startEdit(); return; }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); clearSel(); return; }
    if (e.key.length === 1) { e.preventDefault(); startEdit(e.key); }
  });
  document.addEventListener('copy', e => onCopy(e)); document.addEventListener('cut', e => onCopy(e, true)); document.addEventListener('paste', onPaste);
  // toolbar
  $$('#xBar button, #xTop .ib').forEach(b => b.addEventListener('pointerdown', e => { if (editing?.src === 'cell') e.preventDefault(); }));
  $('#xUndo').onclick = () => { commit(); undo(); }; $('#xRedo').onclick = () => { commit(); redo(); };
  $('#xB').onclick = () => toggle('b'); $('#xI').onclick = () => toggle('i'); $('#xU').onclick = () => toggle('u');
  $$('#xAlign button').forEach(b => b.onclick = () => fmt(s => { s.al = s.al === b.dataset.al ? '' : b.dataset.al; }));
  $$('#xColor button').forEach(b => b.onclick = () => fmt(s => { s.color = b.dataset.c === '#111111' ? '' : b.dataset.c; }));
  $$('#xFill button').forEach(b => b.onclick = () => fmt(s => { s.fill = b.dataset.c === 'none' ? '' : b.dataset.c; }));
  $('#xFmt').onchange = e => { const k = e.target.value; if (k) fmt(s => { s.nf = k === 'gen' ? '' : FMTS[k]; }); };
  $('#xSumBtn').onclick = autoSum;
  $$('#xBar [data-rc]').forEach(b => b.onclick = () => rowsCols(b.dataset.rc));
  $('#xSortA').onclick = () => sortSel(false); $('#xSortZ').onclick = () => sortSel(true);
  $('#xWrap').tabIndex = 0;
  $('#xRange').onclick = () => { rangeMode = !rangeMode; $('#xRange').classList.toggle('on', rangeMode); if (rangeMode) ctx.toast('Tap another cell to select a range'); };
  $('#xClear').onclick = () => { commit(); clearSel(); };
  $('#xName').addEventListener('change', e => { const v = e.target.value.trim().replace(/[\\/:*?"<>|]/g, '') || 'Untitled'; e.target.value = v; xname = v + '.xlsx'; dirty = true; autosave(); });
  $('#xSave').onclick = () => { commit(); $('#xSaveSheet').classList.remove('hidden'); };
  $$('#xSaveSheet [data-x]').forEach(b => b.onclick = () => { $('#xSaveSheet').classList.add('hidden'); exportAs(b.dataset.x); });
  $('#xSheetRename').onclick = () => { const nm = $('#xSheetName').value.trim().replace(/[\[\]*?\/\\:]/g, '').slice(0, 31); const old = book.sheets[act].name;
    if (!nm || nm === old) return $('#xSheetSheet').classList.add('hidden'); if (hf.doesSheetExist(nm)) return ctx.toast('A sheet has that name already');
    snap(); hf.renameSheet(sid(act), nm); book.sheets[act].name = nm; $('#xSheetSheet').classList.add('hidden'); renderTabs(); changed(); };
  $('#xSheetDel').onclick = () => { if (book.sheets.length < 2) return ctx.toast('A workbook needs at least one sheet'); if (!confirm(`Delete sheet "${book.sheets[act].name}"?`)) return;
    snap(); hf.removeSheet(sid(act)); book.sheets.splice(act, 1); act = Math.max(0, act - 1); $('#xSheetSheet').classList.add('hidden'); renderAll(); changed(); };
  $('#xBack').onclick = async () => { commit(); await autosave(); $('#xlEd').classList.add('hidden'); $('#home').classList.remove('hidden'); ctx.home(); };
  window.addEventListener('beforeunload', () => { if (dirty) autosave(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && dirty) autosave(); });
}
export const excelState = { get hf() { return hf; }, get book() { return book; }, get dirty() { return dirty; } };
