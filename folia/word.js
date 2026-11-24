// Folia — Word editor. mammoth reads .docx, a contenteditable page edits it, docx/html writes .docx.
// Folia keeps its own HTML in a custom property so its formatting survives a round trip; if the text no longer
// matches (edited in Word since), the file is read fresh with mammoth.
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
let ctx, ready, wid = null, wname = 'Document.docx', dirty = false, saveT = null, lastRange = null;
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const EXPORT_CSS = 'table{border-collapse:collapse}td,th{border:1px solid #999999;padding:4px}';

function load(src) { return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load ' + src)); document.head.appendChild(s); }); }
function libs() {
  return ready ||= (async () => {
    await Promise.all([load('vendor/mammoth.browser.min.js'), load('vendor/fflate.min.js'), load('vendor/docx.iife.js')]);
    await load('vendor/docx-math.iife.js'); await load('vendor/docx-html.iife.js');
  })().catch(e => { ready = null; throw e; });
}
const ed = () => $('#wPage');
const norm = s => (s || '').replace(/\s+/g, ' ').trim();
const textOf = html => { const d = document.createElement('div'); d.innerHTML = html; return norm(d.textContent); };
const b64e = s => btoa(unescape(encodeURIComponent(s)));
const b64d = s => decodeURIComponent(escape(atob(s)));

// ---------- read ----------
function foliaHtml(bytes) {
  try {
    const z = fflate.unzipSync(bytes, { filter: f => f.name === 'docProps/custom.xml' });
    const xml = z['docProps/custom.xml'] && fflate.strFromU8(z['docProps/custom.xml']); if (!xml) return null;
    const m = xml.match(/name="FoliaHtml"[^>]*>\s*<vt:lpwstr>([^<]*)<\/vt:lpwstr>/); return m ? b64d(m[1].trim()) : null;
  } catch { return null; }
}
async function readDocx(bytes) {
  const r = await mammoth.convertToHtml({ arrayBuffer: bytes.slice().buffer }, {
    styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh"],
    convertImage: mammoth.images.imgElement(img => img.read('base64').then(b => ({ src: 'data:' + img.contentType + ';base64,' + b })))
  });
  let html = r.value.replace(/<a id="[^"]*"><\/a>/g, '');
  const mine = foliaHtml(bytes);
  if (mine && textOf(mine) === textOf(html)) { // restore Folia's formatting; images come back from the file itself
    const srcs = [...html.matchAll(/<img [^>]*src="([^"]+)"/g)].map(m => m[1]); let k = 0;
    html = mine.replace(/<img([^>]*?)data-fimg="\d+"([^>]*)>/g, (all, a, b) => `<img${a}src="${srcs[k++] || ''}"${b}>`);
  }
  return html || '<p><br></p>';
}

// ---------- write ----------
function cleanHtml() {
  const d = ed().cloneNode(true);
  d.querySelectorAll('font').forEach(f => { const s = document.createElement('span'); if (f.color) s.style.color = f.color; s.innerHTML = f.innerHTML; f.replaceWith(s); });
  d.querySelectorAll('[contenteditable],[data-sel]').forEach(e => { e.removeAttribute('contenteditable'); e.removeAttribute('data-sel'); });
  d.querySelectorAll('div').forEach(x => { if (!x.querySelector('p,h1,h2,h3,ul,ol,table,div')) { const p = document.createElement('p'); p.innerHTML = x.innerHTML; for (const a of x.attributes) p.setAttribute(a.name, a.value); x.replaceWith(p); } });
  return d.innerHTML;
}
async function buildDocx() {
  await libs();
  const html = cleanHtml();
  let k = 0; const lite = html.replace(/<img([^>]*?)src="[^"]*"([^>]*)>/g, (a, x, y) => `<img${x}data-fimg="${k++}"${y}>`);
  const kids = docxHtml.htmlToContent(html, { stylesheet: EXPORT_CSS, maxImageWidth: 600 });
  const doc = new docx.Document({
    creator: 'Folia', title: wname.replace(/\.docx$/i, ''),
    customProperties: [{ name: 'FoliaHtml', value: b64e(lite) }],
    styles: { default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 120, line: 276 } } } } },
    sections: [{ children: kids.length ? kids : [new docx.Paragraph('')] }]
  });
  return new Uint8Array(await (await docx.Packer.toBlob(doc)).arrayBuffer());
}
async function autosave() {
  clearTimeout(saveT); saveT = null; if (!dirty || !wid) return;
  try { const bytes = await buildDocx(); dirty = false; $('#wState').textContent = 'Saved'; await ctx.touchRecent(wid, wname, bytes, 'docx'); }
  catch (e) { console.error(e); $('#wState').textContent = 'Not saved'; }
}
const changed = () => { dirty = true; $('#wState').textContent = 'Editing…'; clearTimeout(saveT); saveT = setTimeout(autosave, 2500); };

// ---------- open ----------
async function show(html, name, id) {
  wname = /\.docx$/i.test(name) ? name : name.replace(/\.[^.]*$/, '') + '.docx'; wid = id; dirty = false;
  $('#wName').value = wname.replace(/\.docx$/i, '');
  ed().innerHTML = html; $('#wState').textContent = 'Saved';
  $('#home').classList.add('hidden'); $('#editor').classList.add('hidden'); $('#wordEd').classList.remove('hidden');
  try { document.execCommand('styleWithCSS', false, true); document.execCommand('defaultParagraphSeparator', false, 'p'); } catch {}
  window.scrollTo(0, 0); syncBar();
}
export async function openWord(file, id) {
  ctx.busy(true, 'Opening…');
  try { await libs(); const html = await readDocx(file.bytes); const nid = id || Math.random().toString(36).slice(2, 10);
    await show(html, file.name, nid); if (!id) ctx.touchRecent(nid, wname, file.bytes, 'docx'); }
  catch (e) { console.error(e); ctx.toast("Couldn't open that Word file."); }
  ctx.busy(false);
}
export async function newWord() {
  ctx.busy(true, 'Starting…');
  try { await libs(); await show('<h1>Untitled</h1><p><br></p>', 'Untitled.docx', Math.random().toString(36).slice(2, 10)); dirty = true; await autosave();
    const h = ed().querySelector('h1'); const r = document.createRange(); r.selectNodeContents(h); const s = getSelection(); s.removeAllRanges(); s.addRange(r); ed().focus(); }
  catch (e) { console.error(e); ctx.toast("Couldn't start a document."); }
  ctx.busy(false);
}
export const wordLibs = libs;

// ---------- toolbar ----------
function keepRange() { const s = getSelection(); if (s.rangeCount && ed().contains(s.anchorNode)) lastRange = s.getRangeAt(0).cloneRange(); }
function restore() { ed().focus({ preventScroll: true }); if (lastRange) { const s = getSelection(); s.removeAllRanges(); s.addRange(lastRange); } }
function cmd(c, v = null) { restore(); document.execCommand(c, false, v); keepRange(); changed(); syncBar(); }
function blockOf(n) { while (n && n !== ed()) { if (n.nodeType === 1 && /^(P|H1|H2|H3|LI|DIV|TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; } return null; }
function cellOf(n) { while (n && n !== ed()) { if (n.nodeType === 1 && /^(TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; } return null; }
function syncBar() {
  const s = getSelection(); const n = s.anchorNode; if (!n || !ed().contains(n)) { $('#wTbl').classList.add('hidden'); return; }
  ['bold', 'italic', 'underline', 'strikeThrough', 'insertUnorderedList', 'insertOrderedList', 'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'].forEach(c => { const b = $(`#wordEd [data-c="${c}"]`); if (b) { let on = false; try { on = document.queryCommandState(c); } catch {} b.classList.toggle('on', on); } });
  let b = n.nodeType === 1 ? n : n.parentNode; while (b && b !== ed() && !/^(P|H1|H2|H3|LI)$/.test(b.tagName)) b = b.parentNode;
  $('#wStyle').value = b && /^H[123]$/.test(b.tagName) ? b.tagName.toLowerCase() : 'p';
  $('#wTbl').classList.toggle('hidden', !cellOf(n));
}
function setSize(pt) {
  restore(); const s = getSelection(); if (!s.rangeCount) return;
  if (s.isCollapsed) { const b = blockOf(s.anchorNode); if (b) b.style.fontSize = pt + 'pt'; changed(); return; }
  document.execCommand('styleWithCSS', false, false); document.execCommand('fontSize', false, '7'); document.execCommand('styleWithCSS', false, true);
  ed().querySelectorAll('font[size="7"]').forEach(f => { const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.innerHTML = f.innerHTML; f.replaceWith(sp); });
  ed().querySelectorAll('span[style*="xxx-large"]').forEach(sp => sp.style.fontSize = pt + 'pt');
  keepRange(); changed();
}
function insertTable() {
  const row = c => '<tr>' + Array(3).fill(`<${c}><br></${c}>`).join('') + '</tr>';
  const html = `<table><tbody>${row('td')}${row('td')}${row('td')}</tbody></table><p><br></p>`;
  restore(); let n = getSelection().anchorNode, list = null; while (n && n !== ed()) { if (n.nodeType === 1 && /^(UL|OL)$/.test(n.tagName)) list = n; n = n.parentNode; }
  if (list) { // never put a table inside a bullet; drop it after the list
    const li = getSelection().anchorNode?.nodeType === 1 ? getSelection().anchorNode.closest?.('li') : getSelection().anchorNode?.parentNode?.closest?.('li');
    if (li && !norm(li.textContent)) li.remove();
    const t = document.createElement('template'); t.innerHTML = html; const first = t.content.firstChild, last = t.content.lastChild; list.after(t.content);
    const r = document.createRange(); r.selectNodeContents(first.querySelector('td')); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); keepRange(); changed(); syncBar(); return;
  }
  cmd('insertHTML', html);
}
function tableOp(op) {
  restore(); const c = cellOf(getSelection().anchorNode); if (!c) return; const tr = c.parentNode, tb = tr.closest('table'), ci = [...tr.children].indexOf(c);
  if (op === 'row') { const n = tr.cloneNode(true); n.querySelectorAll('td,th').forEach(x => x.innerHTML = '<br>'); tr.after(n); }
  if (op === 'col') tb.querySelectorAll('tr').forEach(r => { const x = document.createElement(r.children[ci]?.tagName || 'td'); x.innerHTML = '<br>'; r.children[ci].after(x); });
  if (op === 'delrow') { if (tb.rows.length > 1) tr.remove(); else tb.remove(); }
  if (op === 'delcol') { if (tr.children.length > 1) tb.querySelectorAll('tr').forEach(r => r.children[ci]?.remove()); else tb.remove(); }
  changed(); syncBar();
}
async function insertImage(f) {
  const src = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(f); });
  const png = await ctx.toPNG(src, 1400); cmd('insertHTML', `<img src="${png}" style="max-width:100%">`);
}
async function exportAs(kind) {
  if (kind === 'docx') {
    ctx.busy(true, 'Saving…');
    try { const bytes = await buildDocx(); ctx.busy(false); const ok = await ctx.deliver(bytes, wname, DOCX); if (ok) { dirty = false; ctx.touchRecent(wid, wname, bytes, 'docx'); $('#wState').textContent = 'Saved'; } }
    catch (e) { console.error(e); ctx.busy(false); ctx.toast("Couldn't save: " + (e.message || e)); }
  } else { await autosave(); const t = document.title; document.title = wname.replace(/\.docx$/i, ''); document.body.classList.add('printing'); const done = () => { document.body.classList.remove('printing'); document.title = t; }; window.addEventListener('afterprint', done, { once: true }); window.print(); setTimeout(done, 60000); }
}

export function initWord(c) {
  ctx = c;
  const page = ed();
  page.addEventListener('input', changed);
  page.addEventListener('paste', e => { // plain-ish paste keeps documents clean
    const html = e.clipboardData?.getData('text/html'); if (html) return;
    const t = e.clipboardData?.getData('text/plain'); if (t != null) { e.preventDefault(); document.execCommand('insertText', false, t); }
  });
  document.addEventListener('selectionchange', () => { if ($('#wordEd').classList.contains('hidden')) return; keepRange(); syncBar(); });
  $$('#wordEd [data-c]').forEach(b => { b.addEventListener('pointerdown', e => e.preventDefault()); b.onclick = () => cmd(b.dataset.c, b.dataset.v || null); });
  $('#wStyle').onchange = e => cmd('formatBlock', '<' + e.target.value + '>');
  $('#wSize').onchange = e => { if (e.target.value) setSize(+e.target.value); e.target.value = ''; };
  $$('#wColor button').forEach(b => { b.addEventListener('pointerdown', e => e.preventDefault()); b.onclick = () => cmd('foreColor', b.dataset.c); });
  $$('#wMark button').forEach(b => { b.addEventListener('pointerdown', e => e.preventDefault()); b.onclick = () => cmd('hiliteColor', b.dataset.c); });
  $('#wLink').onclick = () => { keepRange(); const u = prompt('Link address'); if (u) cmd('createLink', /^[a-z]+:/i.test(u) ? u : 'https://' + u); };
  $('#wImg').onchange = e => { const f = e.target.files[0]; if (f) insertImage(f); e.target.value = ''; };
  $('#wImgBtn').onclick = () => { keepRange(); $('#wImg').click(); };
  $('#wTable').onclick = insertTable;
  $$('#wTbl [data-t]').forEach(b => { b.addEventListener('pointerdown', e => e.preventDefault()); b.onclick = () => tableOp(b.dataset.t); });
  $('#wName').addEventListener('change', e => { const v = e.target.value.trim().replace(/[\\/:*?"<>|]/g, '') || 'Untitled'; e.target.value = v; wname = v + '.docx'; dirty = true; autosave(); });
  $('#wSave').onclick = () => $('#wSaveSheet').classList.remove('hidden');
  $$('#wSaveSheet [data-x]').forEach(b => b.onclick = () => { $('#wSaveSheet').classList.add('hidden'); exportAs(b.dataset.x); });
  $('#wBack').onclick = async () => { await autosave(); $('#wordEd').classList.add('hidden'); $('#home').classList.remove('hidden'); ctx.home(); };
  window.addEventListener('beforeunload', () => { if (dirty) autosave(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && dirty) autosave(); });
}
export const wordState = { get dirty() { return dirty; }, get id() { return wid; }, buildDocx, readDocx };
