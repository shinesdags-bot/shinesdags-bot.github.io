// Folia — PDF editor (phase 1+2). pdf.js renders, pdf-lib writes.
import * as pdfjsLib from './vendor/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdf.worker.min.mjs';
import { Sync } from './sync.js';
import { initWord, openWord, newWord, wordLibs } from './word.js';
import { initExcel, openExcel, newExcel, excelLibs } from './excel.js';
import { initPpt, openPpt, newPpt, pptLibs } from './ppt.js';
const { PDFDocument, rgb, degrees, StandardFonts } = window.PDFLib;

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const uid = () => Math.random().toString(36).slice(2, 10);

// ---------- state ----------
let docs = [];          // {bytes, js}
let pages = [];         // {key, doc, idx, w0, h0, ox, oy, nat, urot, objs}
let imgs = {};          // id -> dataURL
let docName = 'Document.pdf', recentId = null;
let tool = 'select', sel = null, editing = null;
let props = { size: 14, color: '#111111', width: 2 };
let undoStack = [], dirty = false;
const canvasCache = new Map();

// ---------- helpers ----------
const toast = (t, ms = 2600) => { const e = $('#toast'); e.textContent = t; e.classList.remove('hidden'); clearTimeout(toast.t); toast.t = setTimeout(() => e.classList.add('hidden'), ms); };
const busy = (on, t = 'Working…') => { $('#busyT').textContent = t; $('#busy').classList.toggle('hidden', !on); };
const hex2rgb = h => { const n = parseInt(h.slice(1), 16); return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); };
const eff = p => (p.nat + p.urot) % 360;
const dims = p => (eff(p) % 180 ? { w: p.h0, h: p.w0 } : { w: p.w0, h: p.h0 });
const readFile = f => f.arrayBuffer().then(b => new Uint8Array(b));
function snapshot() { undoStack.push(JSON.stringify({ pages, imgs })); if (undoStack.length > 40) undoStack.shift(); dirty = true; }
function undo() { if (!undoStack.length) return toast('Nothing to undo'); const s = JSON.parse(undoStack.pop()); pages = s.pages; imgs = s.imgs; sel = null; renderAll(); }

// ---------- IndexedDB recents ----------
const idb = new Promise((res, rej) => { const r = indexedDB.open('folia', 1); r.onupgradeneeded = () => r.result.createObjectStore('files', { keyPath: 'id' }); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
const store = async mode => (await idb).transaction('files', mode).objectStore('files');
const dbAll = async () => new Promise(async r => { const q = (await store('readonly')).getAll(); q.onsuccess = () => r(q.result || []); q.onerror = () => r([]); });
const dbPut = async v => (await store('readwrite')).put(v);
const dbDel = async id => (await store('readwrite')).delete(id);
const dbGet = async id => new Promise(async r => { const q = (await store('readonly')).get(id); q.onsuccess = () => r(q.result); });

async function renderRecent() {
  const me = Sync.user?.id;
  const list = (await dbAll()).filter(f => !f.uid || f.uid === me).sort((a, b) => b.t - a.t).slice(0, 30);
  const box = $('#recent');
  if (!list.length) { box.innerHTML = `<p class="empty">${me ? 'No files yet. Open one and it shows up on all your devices.' : 'Files you open show up here, saved on this device.'}</p>`; return; }
  box.innerHTML = '';
  for (const f of list) {
    const b = document.createElement('div'); b.className = 'rc';
    const st = !me ? '' : f.synced && f.t <= f.synced ? '<span class="cl" title="On all your devices">☁︎</span>' : '<span class="cl wait" title="Waiting to sync">⟳</span>';
    b.innerHTML = `${isDocx(f) ? '<span class="tag doc">DOCX</span>' : isXlsx(f) ? '<span class="tag xls">XLSX</span>' : isPptx(f) ? '<span class="tag ppt">PPTX</span>' : '<span class="tag pdf">PDF</span>'}<b></b>${st}<small>${new Date(f.t).toLocaleDateString()}</small><button class="x" title="Remove">✕</button>`;
    b.querySelector('b').textContent = f.name;
    b.onclick = async e => {
      if (e.target.classList.contains('x')) { e.stopPropagation();
        if (me && f.synced && !confirm(`Remove "${f.name}" from all your devices?`)) return;
        await Sync.removeEverywhere(f); await dbDel(f.id); renderRecent(); return; }
      const r = await dbGet(f.id); if (isDocx(r)) openWord({ name: r.name, bytes: r.bytes }, r.id); else if (isXlsx(r)) openExcel({ name: r.name, bytes: r.bytes }, r.id); else if (isPptx(r)) openPpt({ name: r.name, bytes: r.bytes }, r.id); else openBuffers([{ name: r.name, bytes: r.bytes }], r.id); };
    box.appendChild(b);
  }
}
const isDocx = r => r.kind === 'docx' || /\.docx$/i.test(r.name || '');
const isPptx = r => r.kind === 'pptx' || /\.pptx$/i.test(r.name || '');
const isXlsx = r => r.kind === 'xlsx' || /\.xlsx$/i.test(r.name || '');
async function touchRecent(id, name, bytes, kind = 'pdf') { const old = await dbGet(id); await dbPut({ ...(old || {}), id, name, bytes, kind, t: Date.now(), uid: Sync.user?.id || old?.uid }); renderRecent(); Sync.syncNow(); }

// ---------- open ----------
async function loadDoc(bytes) {
  const js = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
  docs.push({ bytes, js }); const di = docs.length - 1, out = [];
  for (let i = 1; i <= js.numPages; i++) {
    const pg = await js.getPage(i);
    const v = pg.view; // [x0,y0,x1,y1]
    out.push({ key: uid(), doc: di, idx: i - 1, w0: v[2] - v[0], h0: v[3] - v[1], ox: v[0], oy: v[1], nat: (pg.rotate || 0) % 360, urot: 0, objs: [] });
  }
  return out;
}
async function openBuffers(files, rid = null) {
  busy(true, 'Opening…');
  try {
    docs = []; pages = []; imgs = {}; undoStack = []; sel = null; canvasCache.clear(); dirty = false;
    for (const f of files) pages.push(...await loadDoc(f.bytes));
    docName = files.length > 1 ? 'Merged.pdf' : files[0].name;
    if (files.length === 1) { recentId = rid || uid(); if (!rid) touchRecent(recentId, docName, files[0].bytes); } else recentId = null;
    $('#docName').textContent = docName;
    $('#home').classList.add('hidden'); $('#editor').classList.remove('hidden');
    setTool('select'); renderAll();
  } catch (e) { console.error(e); toast(e?.name === 'PasswordException' ? 'This PDF is password-protected.' : "Couldn't open that file."); }
  busy(false);
}
const openFiles = async list => {
  const all = [...list], dx = all.filter(f => /\.docx$/i.test(f.name) || /wordprocessingml/.test(f.type));
  if (dx.length) return openWord({ name: dx[0].name, bytes: await readFile(dx[0]) });
  const xl = all.filter(f => /\.(xlsx|xlsm|csv)$/i.test(f.name) || /spreadsheetml|text\/csv/.test(f.type));
  if (xl.length) return openExcel({ name: xl[0].name, bytes: await readFile(xl[0]) });
  const pp = all.filter(f => /\.pptx$/i.test(f.name) || /presentationml/.test(f.type));
  if (pp.length) return openPpt({ name: pp[0].name, bytes: await readFile(pp[0]) });
  if (all.some(f => /\.ppt$/i.test(f.name))) return toast('Old .ppt files aren\u2019t supported. Save it as .pptx first.');
  if (all.some(f => /\.xls$/i.test(f.name))) return toast('Old .xls files aren\u2019t supported yet. Save it as .xlsx first.');
  if (all.some(f => /\.doc$/i.test(f.name))) return toast('Old .doc files aren\u2019t supported yet. Save it as .docx first.');
  const fs = all.filter(f => /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name));
  if (!fs.length) return toast('Pick a PDF, Word, Excel or PowerPoint file.');
  openBuffers(await Promise.all(fs.map(async f => ({ name: f.name, bytes: await readFile(f) }))));
};

// ---------- render ----------
let io;
function pageWidthPx(p) { const vw = $('#viewer').clientWidth - 20; const d = dims(p); return Math.min(vw, d.w * 1.6, 1000); }
function renderAll() {
  const v = $('#viewer'); const keep = v.scrollTop; v.innerHTML = '';
  io?.disconnect(); io = new IntersectionObserver(es => es.forEach(e => e.isIntersecting && paintCanvas(e.target)), { root: v, rootMargin: '600px' });
  pages.forEach((p, i) => {
    const d = dims(p), W = pageWidthPx(p), s = W / d.w;
    const el = document.createElement('div'); el.className = 'page'; el.dataset.i = i; el.style.width = W + 'px'; el.style.height = d.h * s + 'px';
    el.innerHTML = `<div class="layer"></div><span class="num">${i + 1} / ${pages.length}</span>`;
    el._s = s; v.appendChild(el); io.observe(el); renderObjs(i); bindLayer(el, i);
  });
  v.scrollTop = keep;
}
async function paintCanvas(el) {
  const p = pages[+el.dataset.i]; if (!p || el.querySelector('canvas.pc')) return;
  const W = parseFloat(el.style.width), ck = p.key + ':' + eff(p) + ':' + Math.round(W);
  let c = canvasCache.get(ck);
  if (!c) {
    c = document.createElement('canvas'); c.className = 'pc';
    const dpr = Math.min(window.devicePixelRatio || 1, 2.5);
    if (p.doc == null) { c.width = 10; c.height = 10; c.getContext('2d').fillStyle = '#fff'; c.getContext('2d').fillRect(0, 0, 10, 10); }
    else {
      const pg = await docs[p.doc].js.getPage(p.idx + 1);
      const vp = pg.getViewport({ scale: (W / dims(p).w) * dpr, rotation: eff(p) });
      c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    }
    canvasCache.set(ck, c);
  }
  if (!el.isConnected || el.querySelector('canvas.pc')) return;
  el.prepend(c.cloneNode ? cloneCanvas(c) : c);
}
function cloneCanvas(c) { const n = document.createElement('canvas'); n.className = 'pc'; n.width = c.width; n.height = c.height; n.getContext('2d').drawImage(c, 0, 0); return n; }
const pageEl = i => $(`.page[data-i="${i}"]`);

function renderObjs(i) {
  const el = pageEl(i); if (!el) return; const L = el.querySelector('.layer'), s = el._s; L.innerHTML = '';
  pages[i].objs.forEach(o => {
    const d = document.createElement('div'); d.className = 'ob ' + o.t; d.dataset.id = o.id;
    d.style.left = o.x * s + 'px'; d.style.top = o.y * s + 'px';
    if (o.t === 'text') { d.textContent = o.text; d.style.fontSize = o.size * s + 'px'; d.style.color = o.color; }
    else { d.style.width = o.w * s + 'px'; d.style.height = o.h * s + 'px'; }
    if (o.t === 'rect') { d.classList.add(o.kind); if (o.kind === 'hl') { d.style.background = o.color; d.style.opacity = .4; } }
    if (o.t === 'img') d.innerHTML = `<img src="${imgs[o.src]}" alt="">`;
    if (o.t === 'ink') d.innerHTML = `<svg viewBox="0 0 ${o.bw} ${o.bh}" preserveAspectRatio="none"><path d="${inkPath(o.pts)}" fill="none" stroke="${o.color}" stroke-width="${o.lw}" stroke-linecap="round" stroke-linejoin="round" vector-effect="non-scaling-stroke" style="stroke-width:${o.lw * s}px"/></svg>`;
    if (sel && sel.id === o.id) { d.classList.add('sel'); if (o.t !== 'text') d.insertAdjacentHTML('beforeend', '<span class="h-rs"></span>'); d.insertAdjacentHTML('beforeend', '<button class="h-del">Delete</button>'); }
    bindObj(d, i, o); L.appendChild(d);
  });
}
const inkPath = pts => pts.map((p, k) => (k ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' ') + (pts.length === 1 ? ' l0.1 0' : '');

// ---------- interaction ----------
function ptIn(el, e) { const r = el.getBoundingClientRect(); return { x: (e.clientX - r.left) / el._s, y: (e.clientY - r.top) / el._s }; }
function select(o, i) { if (editing && (!o || editing.id !== o.id)) stopEdit(); sel = o ? { id: o.id, i } : null; pages.forEach((_, k) => renderObjs(k)); if (o && (o.size || o.color)) syncProps(o); }
function syncProps(o) { if (o.size) $('#pSize').value = o.size; if (o.color) $$('#pColor button').forEach(b => b.classList.toggle('on', b.dataset.c === o.color)); }
const findObj = id => { for (let i = 0; i < pages.length; i++) { const o = pages[i].objs.find(o => o.id === id); if (o) return { o, i }; } return null; };

function bindLayer(el, i) {
  const L = el.querySelector('.layer');
  L.addEventListener('pointerdown', e => {
    if (e.target !== L) return;
    const pt = ptIn(el, e);
    if (tool === 'select') { select(null); return; }
    if (tool === 'text') { e.preventDefault(); setTool('select'); snapshot(); const o = { id: uid(), t: 'text', x: pt.x, y: pt.y - props.size * .6, size: props.size, color: props.color, text: '' }; pages[i].objs.push(o); select(o, i); startEdit(o.id); return; }
    if (tool === 'whiteout' || tool === 'highlight') return dragRect(el, i, pt, e);
    if (tool === 'draw') return drawInk(el, i, pt, e);
  });
}
function dragRect(el, i, p0, e) {
  const L = el.querySelector('.layer'), s = el._s, g = document.createElement('div');
  const kind = tool === 'whiteout' ? 'wo' : 'hl', color = kind === 'hl' ? (props.color === '#111111' ? '#ffd23f' : props.color) : '#ffffff';
  g.className = 'ob rect ' + kind; if (kind === 'hl') { g.style.background = color; g.style.opacity = .4; } L.appendChild(g); L.setPointerCapture(e.pointerId);
  let r = { x: p0.x, y: p0.y, w: 0, h: 0 };
  const mv = ev => { const p = ptIn(el, ev); r = { x: Math.min(p.x, p0.x), y: Math.min(p.y, p0.y), w: Math.abs(p.x - p0.x), h: Math.abs(p.y - p0.y) }; Object.assign(g.style, { left: r.x * s + 'px', top: r.y * s + 'px', width: r.w * s + 'px', height: r.h * s + 'px' }); };
  const up = () => { L.removeEventListener('pointermove', mv); L.removeEventListener('pointerup', up); L.removeEventListener('pointercancel', up); g.remove();
    if (r.w > 3 && r.h > 3) { snapshot(); pages[i].objs.push({ id: uid(), t: 'rect', kind, color, ...r }); renderObjs(i); } };
  L.addEventListener('pointermove', mv); L.addEventListener('pointerup', up); L.addEventListener('pointercancel', up);
}
function drawInk(el, i, p0, e) {
  const L = el.querySelector('.layer'), s = el._s; L.setPointerCapture(e.pointerId);
  const ns = 'http://www.w3.org/2000/svg', svg = document.createElementNS(ns, 'svg'), path = document.createElementNS(ns, 'path');
  Object.assign(svg.style, { position: 'absolute', inset: 0, width: '100%', height: '100%', pointerEvents: 'none' });
  path.setAttribute('fill', 'none'); path.setAttribute('stroke', props.color); path.setAttribute('stroke-width', props.width * s); path.setAttribute('stroke-linecap', 'round'); path.setAttribute('stroke-linejoin', 'round');
  svg.appendChild(path); L.appendChild(svg); const pts = [[p0.x, p0.y]];
  const mv = ev => { (ev.getCoalescedEvents?.() || [ev]).forEach(c => { const p = ptIn(el, c); pts.push([p.x, p.y]); }); path.setAttribute('d', pts.map((p, k) => (k ? 'L' : 'M') + p[0] * s + ' ' + p[1] * s).join(' ')); };
  const up = () => { L.removeEventListener('pointermove', mv); L.removeEventListener('pointerup', up); L.removeEventListener('pointercancel', up); svg.remove();
    const pad = props.width; const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    const x = Math.min(...xs) - pad, y = Math.min(...ys) - pad, bw = Math.max(...xs) - x + pad, bh = Math.max(...ys) - y + pad;
    snapshot(); pages[i].objs.push({ id: uid(), t: 'ink', x, y, w: bw, h: bh, bw, bh, pts: pts.map(p => [p[0] - x, p[1] - y]), color: props.color, lw: props.width }); renderObjs(i); };
  L.addEventListener('pointermove', mv); L.addEventListener('pointerup', up); L.addEventListener('pointercancel', up);
}
function bindObj(d, i, o) {
  d.addEventListener('pointerdown', e => {
    if (tool === 'draw') return;
    if (e.target.classList.contains('h-del')) { e.stopPropagation(); snapshot(); pages[i].objs = pages[i].objs.filter(x => x.id !== o.id); sel = null; renderObjs(i); return; }
    if (editing && editing.id === o.id) return; // typing
    e.stopPropagation(); e.preventDefault();
    const wasSel = sel && sel.id === o.id; if (tool !== 'select') setTool('select');
    const el = pageEl(i), s = el._s, rs = e.target.classList.contains('h-rs'), p0 = ptIn(el, e), start = { ...o };
    if (!wasSel) select(o, i);
    let moved = false, snapped = false; const t = e.currentTarget.isConnected ? e.currentTarget : pageEl(i).querySelector(`[data-id="${o.id}"]`);
    const mv = ev => { const p = ptIn(el, ev), dx = p.x - p0.x, dy = p.y - p0.y; if (!moved && Math.hypot(dx, dy) < 2) return; moved = true; if (!snapped) { snapshot(); snapped = true; }
      const live = pages[i].objs.find(x => x.id === o.id); const node = pageEl(i).querySelector(`[data-id="${o.id}"]`);
      if (rs) { let w = Math.max(10, start.w + dx), h = Math.max(6, start.h + dy); if (o.t === 'img' || o.t === 'ink') { const r = start.w / start.h; h = w / r; } live.w = w; live.h = h; node.style.width = w * s + 'px'; node.style.height = h * s + 'px'; }
      else { live.x = start.x + dx; live.y = start.y + dy; node.style.left = live.x * s + 'px'; node.style.top = live.y * s + 'px'; } };
    const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); if (!moved && wasSel && o.t === 'text') startEdit(o.id); };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
  });
  if (o.t === 'text') d.addEventListener('dblclick', () => startEdit(o.id));
}
function startEdit(id) {
  const f = findObj(id); if (!f) return; editing = { id };
  const node = pageEl(f.i).querySelector(`[data-id="${id}"]`); node.querySelector('.h-del')?.remove();
  node.contentEditable = 'true'; node.focus();
  const r = document.createRange(); r.selectNodeContents(node); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  node.oninput = () => { f.o.text = node.innerText.replace(/\n$/, ''); };
  node.onblur = () => setTimeout(() => editing && editing.id === id && stopEdit(), 50);
}
function stopEdit() {
  if (!editing) return; const f = findObj(editing.id); editing = null;
  if (f && !f.o.text.trim()) { f.o && (pages[f.i].objs = pages[f.i].objs.filter(x => x.id !== f.o.id)); sel = null; }
  if (f) renderObjs(f.i);
}
function currentPage() { const v = $('#viewer'), mid = v.getBoundingClientRect().top + v.clientHeight / 2; let best = 0; $$('.page').forEach(el => { const r = el.getBoundingClientRect(); if (r.top < mid) best = +el.dataset.i; }); return best; }
function placeImage(dataURL, wpt) {
  const im = new Image(); im.onload = () => {
    const i = currentPage(), d = dims(pages[i]), w = Math.min(wpt, d.w * .8), h = w * im.height / im.width;
    const id = uid(); imgs[id] = dataURL; snapshot();
    const o = { id: uid(), t: 'img', src: id, x: (d.w - w) / 2, y: Math.max(10, d.h / 2 - h / 2), w, h }; pages[i].objs.push(o); setTool('select'); select(o, i);
    pageEl(i).scrollIntoView({ block: 'center', behavior: 'smooth' });
  }; im.src = dataURL;
}
const toPNG = (src, max = 1600) => new Promise(r => { const im = new Image(); im.onload = () => { const k = Math.min(1, max / Math.max(im.width, im.height)); const c = document.createElement('canvas'); c.width = im.width * k; c.height = im.height * k; c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); r(c.toDataURL('image/png')); }; im.src = src; });

function setTool(t) { if (t === 'edittext') return openEditText(); if (editing) stopEdit(); tool = t; $$('#toolbar button').forEach(b => b.classList.toggle('on', b.dataset.t === t)); $$('.layer').forEach(l => l.classList.toggle('drawing', ['draw', 'whiteout', 'highlight', 'text'].includes(t))); }

// ---------- props ----------
$('#pSize').onchange = e => { props.size = +e.target.value; applyToSel({ size: props.size }); };
$('#pWidth').onchange = e => { props.width = +e.target.value; };
$$('#pColor button').forEach(b => b.onclick = () => { $$('#pColor button').forEach(x => x.classList.remove('on')); b.classList.add('on'); props.color = b.dataset.c; applyToSel({ color: props.color }); });
function applyToSel(ch) { if (!sel) return; const f = findObj(sel.id); if (!f) return; const o = f.o;
  if (ch.size && o.t !== 'text') return; if (ch.color && !(o.t === 'text' || o.t === 'ink' || (o.t === 'rect' && o.kind === 'hl'))) return;
  snapshot(); Object.assign(o, ch); renderObjs(f.i); }

// ---------- toolbar ----------
$$('#toolbar button').forEach(b => b.onclick = () => {
  const t = b.dataset.t;
  if (t === 'image') return $('#imgInput').click();
  if (t === 'sign') return openSign();
  setTool(t);
  if (t === 'text') toast('Tap where you want to type');
  if (t === 'whiteout') toast('Drag over text to cover it, then add new text on top');
});
$('#imgInput').onchange = async e => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; const url = URL.createObjectURL(f); placeImage(await toPNG(url), 200); URL.revokeObjectURL(url); };
$('#undoBtn').onclick = undo;
$('#backBtn').onclick = () => { if (dirty && !confirm('Leave without saving your changes?')) return; $('#editor').classList.add('hidden'); $('#home').classList.remove('hidden'); renderRecent(); };
document.addEventListener('keydown', e => {
  if ($('#editor').classList.contains('hidden') || editing) return;
  if ((e.metaKey || e.ctrlKey) && e.key === 'z') { e.preventDefault(); undo(); }
  if ((e.key === 'Delete' || e.key === 'Backspace') && sel) { const f = findObj(sel.id); if (f) { snapshot(); pages[f.i].objs = pages[f.i].objs.filter(x => x.id !== sel.id); sel = null; renderObjs(f.i); } }
});
let rz; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => !$('#editor').classList.contains('hidden') && renderAll(), 200); });

// ---------- signature ----------
const sigKey = 'folia.sigs';
const savedSigs = () => JSON.parse(localStorage.getItem(sigKey) || '[]');
let sigMode = 'draw', sigPick = null, padDirty = false;
function openSign() { $('#signSheet').classList.remove('hidden'); setSigMode(savedSigs().length ? 'saved' : 'draw'); }
function setSigMode(m) { sigMode = m; $$('#signSheet .seg button').forEach(b => b.classList.toggle('on', b.dataset.m === m)); $('#smDraw').classList.toggle('hidden', m !== 'draw'); $('#smType').classList.toggle('hidden', m !== 'type'); $('#smSaved').classList.toggle('hidden', m !== 'saved'); $('#sigKeep').parentElement.classList.toggle('hidden', m === 'saved'); if (m === 'saved') renderSaved(); }
$$('#signSheet .seg button').forEach(b => b.onclick = () => setSigMode(b.dataset.m));
function renderSaved() { const box = $('#savedSigs'), list = savedSigs(); sigPick = list[0] || null; box.innerHTML = list.length ? '' : '<p class="empty">No saved signatures yet.</p>';
  list.forEach((s, k) => { const d = document.createElement('div'); d.className = 'sv' + (k === 0 ? ' on' : ''); d.innerHTML = `<img src="${s}"><button class="rm">✕</button>`;
    d.onclick = e => { if (e.target.classList.contains('rm')) { const l = savedSigs(); l.splice(k, 1); localStorage.setItem(sigKey, JSON.stringify(l)); return renderSaved(); } $$('.sv').forEach(x => x.classList.remove('on')); d.classList.add('on'); sigPick = s; };
    box.appendChild(d); }); }
(() => { const c = $('#sigPad'), x = c.getContext('2d'); let on = false;
  const P = e => { const r = c.getBoundingClientRect(); return [(e.clientX - r.left) * c.width / r.width, (e.clientY - r.top) * c.height / r.height]; };
  c.onpointerdown = e => { on = true; padDirty = true; c.setPointerCapture(e.pointerId); x.lineWidth = 6; x.lineCap = x.lineJoin = 'round'; x.strokeStyle = '#1a2a6c'; x.beginPath(); x.moveTo(...P(e)); x.lineTo(...P(e)); x.stroke(); };
  c.onpointermove = e => { if (!on) return; (e.getCoalescedEvents?.() || [e]).forEach(ev => x.lineTo(...P(ev))); x.stroke(); };
  c.onpointerup = c.onpointercancel = () => on = false;
  $('#sigClear').onclick = () => { x.clearRect(0, 0, c.width, c.height); padDirty = false; }; })();
$('#sigText').oninput = e => $('#sigPrev').textContent = e.target.value || 'Your name';
function trimCanvas(c) { const x = c.getContext('2d'), d = x.getImageData(0, 0, c.width, c.height).data; let a = c.width, b = c.height, r = 0, t = 0;
  for (let y = 0; y < c.height; y++) for (let X = 0; X < c.width; X++) if (d[(y * c.width + X) * 4 + 3] > 10) { a = Math.min(a, X); r = Math.max(r, X); b = Math.min(b, y); t = Math.max(t, y); }
  if (r <= a) return null; const o = document.createElement('canvas'); o.width = r - a + 12; o.height = t - b + 12; o.getContext('2d').drawImage(c, a - 6, b - 6, o.width, o.height, 0, 0, o.width, o.height); return o.toDataURL('image/png'); }
$('#sigUse').onclick = async () => {
  let url = null;
  if (sigMode === 'saved') url = sigPick;
  else if (sigMode === 'draw') url = padDirty ? trimCanvas($('#sigPad')) : null;
  else { const t = $('#sigText').value.trim(); if (t) { await document.fonts.load('600 120px "Dancing Script"'); const c = document.createElement('canvas'); c.width = 1400; c.height = 260; const x = c.getContext('2d'); x.font = '600 120px "Dancing Script", cursive'; x.fillStyle = '#1a2a6c'; x.textBaseline = 'middle'; x.fillText(t, 20, 130); url = trimCanvas(c); } }
  if (!url) return toast(sigMode === 'saved' ? 'Pick a saved signature' : sigMode === 'draw' ? 'Draw your signature first' : 'Type your name first');
  if (sigMode !== 'saved' && $('#sigKeep').checked) { const l = savedSigs(); if (!l.includes(url)) { l.unshift(url); localStorage.setItem(sigKey, JSON.stringify(l.slice(0, 6))); } }
  $('#signSheet').classList.add('hidden'); placeImage(url, 150); toast('Drag it into place. Use the blue dot to resize.');
};
$$('[data-close]').forEach(b => b.onclick = () => b.closest('.sheet').classList.add('hidden'));
$$('.sheet').forEach(s => s.addEventListener('pointerdown', e => { if (e.target === s) s.classList.add('hidden'); }));

// ---------- pages panel ----------
let picked = new Set();
$('#pagesBtn').onclick = () => { picked = new Set(); $('#pagesSheet').classList.remove('hidden'); renderThumbs(); };
async function renderThumbs() {
  const box = $('#thumbs'); box.innerHTML = '';
  pages.forEach((p, i) => {
    const d = document.createElement('div'); d.className = 'th' + (picked.has(p.key) ? ' on' : '');
    const c = document.createElement('canvas'); const dm = dims(p); c.width = 160; c.height = 160 * dm.h / dm.w;
    d.append(c); d.insertAdjacentHTML('beforeend', `<div>Page ${i + 1}</div><div class="acts"><button data-a="up" title="Move up">↑</button><button data-a="rot" title="Rotate">⟳</button><button data-a="del" title="Delete">🗑</button><button data-a="dn" title="Move down">↓</button></div>`);
    d.onclick = e => { const a = e.target.dataset.a;
      if (!a) { picked.has(p.key) ? picked.delete(p.key) : picked.add(p.key); d.classList.toggle('on'); return; }
      snapshot();
      if (a === 'up' && i > 0) [pages[i - 1], pages[i]] = [pages[i], pages[i - 1]];
      if (a === 'dn' && i < pages.length - 1) [pages[i + 1], pages[i]] = [pages[i], pages[i + 1]];
      if (a === 'del') { if (pages.length === 1) { undoStack.pop(); return toast("A PDF needs at least one page"); } pages.splice(i, 1); }
      if (a === 'rot') rotatePage(p);
      renderThumbs(); renderAll(); };
    box.appendChild(d);
    (async () => { const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); if (p.doc == null) return;
      const pg = await docs[p.doc].js.getPage(p.idx + 1); const vp = pg.getViewport({ scale: c.width / dm.w, rotation: eff(p) }); await pg.render({ canvasContext: x, viewport: vp }).promise; })();
  });
}
function rotImg(o) { const im = new Image(); im.onload = () => { const c = document.createElement('canvas'); c.width = im.height; c.height = im.width; const x = c.getContext('2d'); x.translate(c.width, 0); x.rotate(Math.PI / 2); x.drawImage(im, 0, 0); const id = uid(); imgs[id] = c.toDataURL('image/png'); o.src = id; const f = findObj(o.id); if (f) renderObjs(f.i); }; im.src = imgs[o.src]; }
function rotatePage(p) { // rotate 90° clockwise, carrying objects along
  const { h } = dims(p);
  p.objs.forEach(o => {
    const ow = o.t === 'text' ? o.size * Math.max(1, o.text.length) * .5 : o.w, oh = o.t === 'text' ? o.size * 1.15 : o.h;
    const nx = h - o.y - oh, ny = o.x; o.x = nx; o.y = ny;
    if (o.t !== 'text') { [o.w, o.h] = [o.h, o.w]; }
    if (o.t === 'ink') { o.pts = o.pts.map(([x, y]) => [o.bh - y, x]); [o.bw, o.bh] = [o.bh, o.bw]; }
    if (o.t === 'img') rotImg(o);
  });
  p.urot = (p.urot + 90) % 360;
}
$('#addPdfInput').onchange = async e => { const fs = [...e.target.files]; e.target.value = ''; if (!fs.length) return; busy(true, 'Adding pages…'); snapshot();
  for (const f of fs) pages.push(...await loadDoc(await readFile(f))); busy(false); renderThumbs(); renderAll(); toast('Pages added at the end'); };
$('#blankBtn').onclick = () => { snapshot(); pages.push({ key: uid(), doc: null, idx: 0, w0: 595.28, h0: 841.89, ox: 0, oy: 0, nat: 0, urot: 0, objs: [] }); renderThumbs(); renderAll(); };
$('#splitBtn').onclick = async () => { const list = pages.filter(p => picked.has(p.key)); if (!list.length) return toast('Tap pages to select them first');
  busy(true, 'Extracting…'); try { await deliver(await buildPdf(list), docName.replace(/\.pdf$/i, '') + ' (pages).pdf'); } catch (e) { console.error(e); toast('Extract failed'); } busy(false); };


// ---------- edit existing text (pdfedit: pdf.js 6 + pdf-lib, loaded on demand) ----------
let ptEditor = null, ptMod = null;
const pristine = () => docs.length === 1 && !pages.some(p => p.objs.length || p.urot || p.doc !== 0) && pages.length === (docs[0].js?.numPages || 0) && pages.every((p, k) => p.idx === k);
async function openEditText() {
  if (editing) stopEdit(); if (!pages.length) return;
  busy(true, 'Getting text ready…');
  try {
    const bytes = pristine() ? docs[0].bytes.slice() : new Uint8Array(await buildPdf(pages));
    ptMod ||= await import('./vendor/pdfedit.min.mjs');
    $('#editor').classList.add('hidden'); $('#ptEd').classList.remove('hidden');
    const host = $('#ptHost'); host.innerHTML = '';
    let warned = false;
    ptEditor = ptMod.createPdfEditor(host, bytes, {
      workerSrc: new URL('./vendor/pdfedit.worker.min.mjs', location.href).href,
      onError: m => { toast("This PDF's text can't be edited here: " + m, 5000); closeEditText(false); },
      onWarning: m => { if (!warned) { warned = true; toast(m, 5000); } },
    });
  } catch (e) { console.error(e); toast("Couldn't start text editing (" + (e?.message || e) + "). Use Whiteout + Add text instead.", 6000); $('#ptEd').classList.add('hidden'); $('#editor').classList.remove('hidden'); }
  busy(false);
}
async function closeEditText(keep) {
  const ed = ptEditor; ptEditor = null; if (!ed) { $('#ptEd').classList.add('hidden'); $('#editor').classList.remove('hidden'); return; }
  if (keep && ed.isDirty()) {
    busy(true, 'Saving text…');
    try { const out = await ed.getBytes(); ed.destroy(); $('#ptHost').innerHTML = ''; $('#ptEd').classList.add('hidden');
      const name = docName, id = recentId || uid(); await openBuffers([{ name, bytes: out }], id); recentId = id; await touchRecent(id, name, out); dirty = false; toast('Text updated'); }
    catch (e) { console.error(e); toast("Couldn't apply those edits.", 4000); ptEditor = ed; busy(false); return; }
    busy(false); return;
  }
  if (!keep && ed.isDirty() && !confirm('Discard your text edits?')) { ptEditor = ed; return; }
  ed.destroy(); $('#ptHost').innerHTML = ''; $('#ptEd').classList.add('hidden'); $('#editor').classList.remove('hidden');
}
// iPhone/iPad: WebKit won't place a caret in the invisible (opacity 0) text overlay pdfedit uses,
// so on a touch tap we reveal the tapped paragraph, focus it and put the caret where the finger was.
{ const host = $('#ptHost'); let p0 = null;
  host.addEventListener('touchstart', e => { const t = e.touches[0]; p0 = e.touches.length === 1 ? { x: t.clientX, y: t.clientY, t: Date.now() } : null; }, { capture: true, passive: true });
  host.addEventListener('touchend', ev => { const t = ev.changedTouches[0]; const e = { clientX: t.clientX, clientY: t.clientY };
    if (!p0 || Math.hypot(e.clientX - p0.x, e.clientY - p0.y) > 10 || Date.now() - p0.t > 600) return; p0 = null;
    const el = document.elementFromPoint(e.clientX, e.clientY)?.closest?.('.pdfedit-para'); if (!el || el.contentEditable !== 'true') return;
    $$('#ptHost .fo-tap').forEach(x => x !== el && x.classList.remove('fo-tap'));
    el.classList.add('fo-tap');
    if (document.activeElement !== el) el.focus({ preventScroll: true });
    const r = document.caretRangeFromPoint?.(e.clientX, e.clientY);
    if (r && el.contains(r.startContainer)) { const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
  }, true);
  host.addEventListener('focusout', e => { const el = e.target.closest?.('.pdfedit-para'); if (el) setTimeout(() => { if (document.activeElement !== el) el.classList.remove('fo-tap'); }, 0); }, true);
}
$('#ptDone').onclick = () => closeEditText(true);
$('#ptCancel').onclick = () => closeEditText(false);
$('#ptUndo').onclick = () => ptEditor?.canUndo() && ptEditor.undo();
$('#ptRedo').onclick = () => ptEditor?.canRedo() && ptEditor.redo();

// ---------- export ----------
async function buildPdf(list) {
  const out = await PDFDocument.create(), libs = {}, font = await out.embedFont(StandardFonts.Helvetica), imgCache = {};
  const lib = async di => libs[di] ||= await PDFDocument.load(docs[di].bytes, { ignoreEncryption: true });
  for (const p of list) {
    const d = dims(p), r = eff(p); let page, ox = 0, oy = 0;
    if (p.doc == null) page = out.addPage([d.w, d.h]);
    else if (r === 0) { const src = await lib(p.doc); [page] = await out.copyPages(src, [p.idx]); page.setRotation(degrees(0)); out.addPage(page); ox = p.ox; oy = p.oy; }
    else {
      const src = await lib(p.doc); const sp = src.getPage(p.idx);
      const [emb] = await out.embedPages([sp], [{ left: p.ox, bottom: p.oy, right: p.ox + p.w0, top: p.oy + p.h0 }]);
      page = out.addPage([d.w, d.h]); const W = p.w0, H = p.h0, X = p.ox, Y = p.oy;
      // embedded form is re-based at its box; place it upright
      const m = { 90: { x: 0, y: W, a: -90 }, 180: { x: W, y: H, a: 180 }, 270: { x: H, y: 0, a: 90 } }[r];
      page.drawPage(emb, { x: m.x, y: m.y, rotate: degrees(m.a), width: W, height: H });
      void X; void Y;
    }
    const top = d.h + oy, L = x => ox + x;
    for (const o of p.objs) {
      if (o.t === 'rect') page.drawRectangle({ x: L(o.x), y: top - o.y - o.h, width: o.w, height: o.h, color: hex2rgb(o.color), opacity: o.kind === 'hl' ? .4 : 1, blendMode: o.kind === 'hl' ? 'Multiply' : undefined });
      if (o.t === 'text') { const lines = clean(o.text).split('\n'); lines.forEach((ln, k) => page.drawText(ln, { x: L(o.x), y: top - o.y - (0.92 + 1.15 * k) * o.size, size: o.size, font, color: hex2rgb(o.color) })); }
      if (o.t === 'img') { const im = imgCache[o.src] ||= await out.embedPng(imgs[o.src]); const rr = o.rot || 0;
        if (!rr) page.drawImage(im, { x: L(o.x), y: top - o.y - o.h, width: o.w, height: o.h });
        else { const iw = rr % 180 ? o.h : o.w, ih = rr % 180 ? o.w : o.h; const cx = L(o.x) + o.w / 2, cy = top - o.y - o.h / 2; const a = -rr * Math.PI / 180;
          const px = cx - (iw / 2 * Math.cos(a) - ih / 2 * Math.sin(a)), py = cy - (iw / 2 * Math.sin(a) + ih / 2 * Math.cos(a));
          page.drawImage(im, { x: px, y: py, width: iw, height: ih, rotate: degrees(-rr) }); } }
      if (o.t === 'ink') { const sx = o.w / o.bw, sy = o.h / o.bh;
        const path = o.pts.map((q, k) => (k ? 'L' : 'M') + (o.x + q[0] * sx).toFixed(2) + ' ' + (o.y + q[1] * sy).toFixed(2)).join(' ') + (o.pts.length === 1 ? ' l0.1 0' : '');
        page.drawSvgPath(path, { x: ox, y: top, borderColor: hex2rgb(o.color), borderWidth: o.lw, borderLineCap: 1 }); }
    }
  }
  return out.save();
}
const clean = s => [...s].map(c => { const n = c.codePointAt(0); if (c === '\n' || (n >= 32 && n < 127) || (n >= 160 && n <= 255)) return c; return ({ '‘': "'", '’': "'", '“': '"', '”': '"', '–': '-', '—': '-', '…': '...', '•': '-', '€': 'EUR', '₦': 'N' })[c] || '?'; }).join('');
async function deliver(bytes, name, mime = 'application/pdf') {
  const file = new File([bytes], name, { type: mime });
  const ext = '.' + name.split('.').pop().toLowerCase();
  const mobile = /iPhone|iPad|Android/i.test(navigator.userAgent) || (navigator.maxTouchPoints > 1 && /Mac/.test(navigator.userAgent));
  if (mobile && navigator.canShare?.({ files: [file] })) { try { await navigator.share({ files: [file], title: name }); return true; } catch (e) { if (e.name === 'AbortError') return false; } }
  if (window.showSaveFilePicker) { try { const h = await showSaveFilePicker({ suggestedName: name, types: [{ description: ({ '.pdf': 'PDF', '.docx': 'Word document', '.xlsx': 'Excel workbook', '.csv': 'CSV', '.pptx': 'PowerPoint presentation' })[ext] || 'File', accept: { [mime]: [ext] } }] }); const w = await h.createWritable(); await w.write(file); await w.close(); toast('Saved'); return true; } catch (e) { if (e.name === 'AbortError') return false; } }
  const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); toast('Downloaded'); return true;
}
$('#saveBtn').onclick = async () => {
  if (editing) stopEdit(); busy(true, 'Saving…');
  try { const bytes = await buildPdf(pages); busy(false); const ok = await deliver(bytes, docName);
    if (ok) { dirty = false; if (recentId) touchRecent(recentId, docName, bytes); } }
  catch (e) { console.error(e); busy(false); toast("Couldn't save: " + (e.message || e)); }
};

// ---------- home wiring ----------
$('#openInput').onchange = e => { openFiles(e.target.files); e.target.value = ''; };
$('#mergeInput').onchange = e => { if (e.target.files.length < 2) toast('Pick two or more PDFs'); else openFiles(e.target.files); e.target.value = ''; };
$('#wordInput').onchange = e => { openFiles(e.target.files); e.target.value = ''; };
$('#xlInput').onchange = e => { openFiles(e.target.files); e.target.value = ''; };
$('#pptInput').onchange = e => { openFiles(e.target.files); e.target.value = ''; };
$$('.tile[data-ppt]').forEach(t => t.onclick = () => t.dataset.ppt === 'new' ? newPpt() : $('#pptInput').click());
$$('.tile[data-xl]').forEach(t => t.onclick = () => t.dataset.xl === 'new' ? newExcel() : $('#xlInput').click());
$$('.tile[data-word]').forEach(t => t.onclick = () => t.dataset.word === 'new' ? newWord() : $('#wordInput').click());
$$('.tile[data-open]').forEach(t => t.onclick = () => (t.dataset.open === 'merge' ? $('#mergeInput') : $('#openInput')).click());
$$('.tile.soon').forEach(t => t.onclick = () => toast('That editor is on the way in the next update.'));
const drop = $('#drop');
['dragenter', 'dragover'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
document.addEventListener('drop', e => e.dataTransfer?.files?.length && openFiles(e.dataTransfer.files));
if ('launchQueue' in window) launchQueue.setConsumer(async p => { if (p.files?.length) openFiles(await Promise.all(p.files.map(h => h.getFile()))); });
window.addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });
if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js', { scope: './' }).catch(() => {});
// ---------- account + sync ----------
let syncState = 'off', syncMsg = '';
function paintAcct() {
  const u = Sync.user, chip = $('#acctChip');
  chip.className = 'acct ' + (u ? syncState : 'off');
  chip.textContent = !u ? 'Sync off' : syncState === 'syncing' ? 'Syncing…' : syncState === 'error' ? 'Sync issue' : 'Synced';
  $('#acctIn').classList.toggle('hidden', !!u); $('#acctOut').classList.toggle('hidden', !u);
  if (u) { $('#acctEmail').textContent = u.email || ''; $('#acctState').textContent = syncState === 'error' ? "Couldn't sync: " + syncMsg : syncState === 'syncing' ? 'Syncing now…' : 'Your files are on all your devices.'; }
}
Sync.init({ all: dbAll, put: dbPut, del: dbDel, get: dbGet,
  status: (s, m) => { syncState = s; syncMsg = m || ''; paintAcct(); if (s === 'error' && m && $('#acctSheet').classList.contains('hidden')) toast("Couldn't sync: " + m); },
  changed: () => { if (!$('#home').classList.contains('hidden')) renderRecent(); } });
$('#acctChip').onclick = () => { paintAcct(); $('#acctSheet').classList.remove('hidden'); };
async function doAuth(create) {
  const em = $('#aEmail').value.trim(), pw = $('#aPass').value;
  if (!em || !pw) return toast('Enter your email and password');
  if (create && pw.length < 6) return toast('Use at least 6 characters');
  busy(true, create ? 'Creating account…' : 'Signing in…');
  try { await Sync.signIn(em, pw, create); $('#aPass').value = ''; renderRecent(); toast('Signed in. Syncing your files.'); }
  catch (e) { toast(e.message); }
  busy(false); paintAcct();
}
$('#aIn').onclick = () => doAuth(false);
$('#aNew').onclick = () => doAuth(true);
$('#aPass').addEventListener('keydown', e => e.key === 'Enter' && doAuth(false));
$('#aShow').onclick = () => { const p = $('#aPass'); p.type = p.type === 'password' ? 'text' : 'password'; $('#aShow').textContent = p.type === 'password' ? 'Show' : 'Hide'; };
$('#aSync').onclick = () => Sync.syncNow();
$('#aOut').onclick = () => { Sync.signOut(); renderRecent(); toast('Signed out. Files stay on this device.'); };
initWord({ toast, busy, deliver, toPNG, touchRecent, home: renderRecent });
initExcel({ toast, busy, deliver, touchRecent, home: renderRecent });
initPpt({ toast, busy, deliver, toPNG, touchRecent, home: renderRecent });
if ('requestIdleCallback' in window) requestIdleCallback(() => wordLibs().then(excelLibs).catch(() => {}), { timeout: 8000 });
renderRecent();
window.__folia = { get pages() { return pages; }, buildPdf, openBuffers, setTool };
// ?open=test/x.pdf opens a file shipped with Folia (same-origin paths only; used for testing)
{ const o = new URLSearchParams(location.search).get('open'); if (o && /^[\w\-\/]+\.(pdf|docx|xlsx|pptx|csv)$/i.test(o) && !o.startsWith('/') && !o.includes('..')) fetch(o).then(r => r.ok ? r.blob() : Promise.reject()).then(b => openFiles([new File([b], o.split('/').pop(), { type: b.type })])).catch(() => toast("Couldn't open that file.")); }
