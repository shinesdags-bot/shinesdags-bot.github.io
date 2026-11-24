// Folia — PowerPoint editor. Reads the .pptx package directly (fflate + DOMParser), draws slides as HTML,
// and writes edits back into the original XML so themes, layouts, notes and anything we don't draw survive.
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const NS = { a: 'http://schemas.openxmlformats.org/drawingml/2006/main', p: 'http://schemas.openxmlformats.org/presentationml/2006/main', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', rel: 'http://schemas.openxmlformats.org/package/2006/relationships', ct: 'http://schemas.openxmlformats.org/package/2006/content-types' };
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const RT = { slide: NS.r + '/slide', layout: NS.r + '/slideLayout', master: NS.r + '/slideMaster', theme: NS.r + '/theme', image: NS.r + '/image', notes: NS.r + '/notesSlide' };
const EMU_PT = 12700;
let ctx, ready, pk = null, pid = null, pname = 'Presentation.pptx', dirty = false, saveT = null;
let cur = 0, selShape = null, editingShape = null, undoS = [], redoS = [], stageScale = 1, shapes = [], thumbT = null;

const load = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load ' + src)); document.head.appendChild(s); });
const libs = () => ready ||= (window.fflate ? Promise.resolve() : load('vendor/fflate.min.js')).catch(e => { ready = null; throw e; });
export const pptLibs = libs;

// ---------- package ----------
// pk = { files: {path: Uint8Array}, docs: {path: Document}, dirtyDocs: Set, slides: [path], sz: {cx, cy} }
const dec = new TextDecoder(), enc = new TextEncoder();
const kids = (el, ns, name) => el ? [...el.children].filter(c => c.namespaceURI === ns && c.localName === name) : [];
const kid = (el, ns, name) => kids(el, ns, name)[0] || null;
const A = (el, name) => kid(el, NS.a, name), P = (el, name) => kid(el, NS.p, name);
const path1 = (el, ...steps) => steps.reduce((e, s) => { if (!e) return null; const [ns, n] = s.split(':'); return kid(e, NS[ns], n); }, el);
const desc = (el, ns, name) => el ? [...el.getElementsByTagNameNS(ns, name)] : [];
function doc(path) {
  if (pk.docs[path]) return pk.docs[path]; const b = pk.files[path]; if (!b) return null;
  const d = new DOMParser().parseFromString(dec.decode(b), 'application/xml'); pk.docs[path] = d; return d;
}
const touch = path => pk.dirtyDocs.add(path);
function flush() { const ser = new XMLSerializer(); for (const p of pk.dirtyDocs) { const d = pk.docs[p]; if (!d) continue; let x = ser.serializeToString(d); if (!x.startsWith('<?xml')) x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + x; pk.files[p] = enc.encode(x); } pk.dirtyDocs.clear(); }
const dirOf = p => p.slice(0, p.lastIndexOf('/') + 1);
const relsPath = p => dirOf(p) + '_rels/' + p.slice(p.lastIndexOf('/') + 1) + '.rels';
function resolve(base, t) { if (t.startsWith('/')) return t.slice(1); const parts = (dirOf(base) + t).split('/'), out = []; for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.') out.push(x); } return out.join('/'); }
function rels(p) { const d = doc(relsPath(p)); const m = {}; if (!d) return m; for (const r of desc(d.documentElement, NS.rel, 'Relationship')) m[r.getAttribute('Id')] = { type: r.getAttribute('Type'), target: r.getAttribute('TargetMode') === 'External' ? null : resolve(p, r.getAttribute('Target')), el: r }; return m; }
const relOfType = (p, type) => Object.values(rels(p)).find(r => r.type === type)?.target || null;
function ensureRels(p) { const rp = relsPath(p); if (!pk.files[rp] && !pk.docs[rp]) { pk.docs[rp] = new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"/>`, 'application/xml'); touch(rp); } return doc(rp); }
function addRel(p, type, target) {
  const d = ensureRels(p), root = d.documentElement; const ids = desc(root, NS.rel, 'Relationship').map(r => +(r.getAttribute('Id').replace(/\D/g, '')) || 0);
  const id = 'rId' + (Math.max(0, ...ids) + 1); const r = d.createElementNS(NS.rel, 'Relationship');
  r.setAttribute('Id', id); r.setAttribute('Type', type); r.setAttribute('Target', relTarget(p, target)); root.appendChild(r); touch(relsPath(p)); return id;
}
function relTarget(from, to) { const a = dirOf(from).split('/').filter(Boolean), b = to.split('/'); let i = 0; while (i < a.length && a[i] === b[i]) i++; return '../'.repeat(a.length - i) + b.slice(i).join('/'); }
function contentType(partPath, type) { const d = doc('[Content_Types].xml'), root = d.documentElement; const pn = '/' + partPath; if (desc(root, NS.ct, 'Override').some(o => o.getAttribute('PartName') === pn)) return; const o = d.createElementNS(NS.ct, 'Override'); o.setAttribute('PartName', pn); o.setAttribute('ContentType', type); root.appendChild(o); touch('[Content_Types].xml'); }
function defaultType(ext, type) { const d = doc('[Content_Types].xml'), root = d.documentElement; if (desc(root, NS.ct, 'Default').some(o => o.getAttribute('Extension').toLowerCase() === ext)) return; const o = d.createElementNS(NS.ct, 'Default'); o.setAttribute('Extension', ext); o.setAttribute('ContentType', type); root.insertBefore(o, root.firstChild); touch('[Content_Types].xml'); }
function dropContentType(partPath) { const d = doc('[Content_Types].xml'); desc(d.documentElement, NS.ct, 'Override').filter(o => o.getAttribute('PartName') === '/' + partPath).forEach(o => o.remove()); touch('[Content_Types].xml'); }
const PRES = 'ppt/presentation.xml';
function readSlides() { const d = doc(PRES), rs = rels(PRES); pk.slides = desc(d.documentElement, NS.p, 'sldId').map(s => rs[s.getAttributeNS(NS.r, 'id')]?.target).filter(Boolean); const sz = desc(d.documentElement, NS.p, 'sldSz')[0]; pk.sz = { cx: +(sz?.getAttribute('cx') || 12192000), cy: +(sz?.getAttribute('cy') || 6858000) }; }
export function openPackage(bytes) { const files = fflate.unzipSync(bytes); pk = { files, docs: {}, dirtyDocs: new Set(), slides: [], sz: null, media: {} }; if (!files[PRES]) throw new Error('Not a PowerPoint file'); readSlides(); return pk; }
export function packBytes() { flush(); return fflate.zipSync(pk.files, { level: 6 }); }

// ---------- theme & colours ----------
function themeOf(master) { const t = relOfType(master, RT.theme); const d = t && doc(t); const out = { clr: {}, major: 'Calibri Light', minor: 'Calibri' }; if (!d) return out;
  const cs = desc(d.documentElement, NS.a, 'clrScheme')[0]; if (cs) for (const c of cs.children) { const v = c.firstElementChild; out.clr[c.localName] = v?.getAttribute('lastClr') || v?.getAttribute('val') || '000000'; }
  const mj = desc(d.documentElement, NS.a, 'majorFont')[0], mn = desc(d.documentElement, NS.a, 'minorFont')[0];
  if (mj) out.major = A(mj, 'latin')?.getAttribute('typeface') || out.major; if (mn) out.minor = A(mn, 'latin')?.getAttribute('typeface') || out.minor; return out; }
function clrMapOf(master) { const m = P(doc(master).documentElement, 'clrMap'); const o = {}; if (m) for (const at of m.attributes) o[at.name] = at.value; return o; }
const hex2rgb = h => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
function rgb2hsl([r, g, b]) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2; if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
function hsl2rgb([h, s, l]) { if (!s) return [l, l, l].map(v => Math.round(v * 255)); const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; const f = t => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; }; return [f(h + 1 / 3), f(h), f(h - 1 / 3)].map(v => Math.round(v * 255)); }
// colour element (srgbClr/schemeClr/sysClr/prstClr) -> css
function colorOf(el, C, phClr) {
  if (!el) return null; let hex;
  if (el.localName === 'srgbClr') hex = el.getAttribute('val');
  else if (el.localName === 'sysClr') hex = el.getAttribute('lastClr') || (el.getAttribute('val') === 'window' ? 'FFFFFF' : '000000');
  else if (el.localName === 'schemeClr') { let v = el.getAttribute('val'); if (v === 'phClr') return phClr || null; v = C.map[v] || v; hex = C.theme.clr[v] || '000000'; }
  else if (el.localName === 'prstClr') hex = ({ black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '008000', yellow: 'FFFF00', gray: '808080' })[el.getAttribute('val')] || '000000';
  else return null;
  let rgb = hex2rgb(hex || '000000'), alpha = 1;
  for (const m of el.children) { const v = +m.getAttribute('val') / 100000;
    if (m.localName === 'lumMod' || m.localName === 'lumOff') { const h = rgb2hsl(rgb); h[2] = m.localName === 'lumMod' ? h[2] * v : Math.min(1, h[2] + v); rgb = hsl2rgb(h); }
    else if (m.localName === 'tint') rgb = rgb.map(c => Math.round(c + (255 - c) * (1 - v)));
    else if (m.localName === 'shade') rgb = rgb.map(c => Math.round(c * v));
    else if (m.localName === 'alpha') alpha = v; }
  return alpha < 1 ? `rgba(${rgb.join(',')},${alpha})` : '#' + rgb.map(c => c.toString(16).padStart(2, '0')).join('');
}
const firstColor = el => el ? [...el.children].find(c => /Clr$/.test(c.localName)) : null;
function fillOf(spPr, C, styleRef, media) { // -> css background or null; 'none' when explicitly none
  if (spPr) for (const f of spPr.children) {
    if (f.localName === 'noFill') return 'none';
    if (f.localName === 'solidFill') return colorOf(firstColor(f), C);
    if (f.localName === 'gradFill') { const gs = desc(f, NS.a, 'gs').map(g => colorOf(firstColor(g), C) + ' ' + (+g.getAttribute('pos') / 1000) + '%'); const ang = +(A(f, 'lin')?.getAttribute('ang') || 5400000) / 60000; return gs.length ? `linear-gradient(${ang + 90}deg,${gs.join(',')})` : null; }
    if (f.localName === 'blipFill' && media) { const id = A(f, 'blip')?.getAttributeNS(NS.r, 'embed'); const u = id && media(id); return u ? `center/cover no-repeat url("${u}")` : null; }
    if (f.localName === 'grpFill') return null;
  }
  if (styleRef && +styleRef.getAttribute('idx') > 0) return colorOf(firstColor(styleRef), C);
  return null;
}

// ---------- placeholder inheritance ----------
const phOf = sp => desc(path1(sp, 'p:nvSpPr', 'p:nvPr') || path1(sp, 'p:nvPicPr', 'p:nvPr') || path1(sp, 'p:nvGraphicFramePr', 'p:nvPr'), NS.p, 'ph')[0] || null;
const phType = ph => ph ? (ph.getAttribute('type') || 'body') : null;
const normT = t => t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t;
function findPh(partPath, ph, toMaster) {
  const d = doc(partPath); if (!d || !ph) return null; const t = phType(ph), idx = ph.getAttribute('idx');
  const all = desc(desc(d.documentElement, NS.p, 'spTree')[0], NS.p, 'sp').map(sp => ({ sp, ph: phOf(sp) })).filter(x => x.ph);
  if (!toMaster && idx != null) { const m = all.find(x => x.ph.getAttribute('idx') === idx); if (m) return m.sp; }
  return (all.find(x => phType(x.ph) === t) || all.find(x => normT(phType(x.ph)) === normT(t)))?.sp || null;
}
function chainOf(slidePath) { const layout = relOfType(slidePath, RT.layout); const master = layout && relOfType(layout, RT.master); return { layout, master }; }

// ---------- rendering ----------
function makeCtx(slidePath) {
  const { layout, master } = chainOf(slidePath); const theme = themeOf(master); const map = Object.assign({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' }, clrMapOf(master));
  const mdoc = doc(master); const txStyles = P(mdoc.documentElement, 'txStyles');
  return { slidePath, layout, master, theme, map, txStyles, defStyle: P(doc(PRES).documentElement, 'defaultTextStyle') };
}
function mediaUrl(partPath, rid) { const t = rels(partPath)[rid]?.target; if (!t || !pk.files[t]) return null; if (!pk.media[t]) { const ext = t.split('.').pop().toLowerCase(); const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp', webp: 'image/webp' }[ext]; if (!type) return null; pk.media[t] = URL.createObjectURL(new Blob([pk.files[t]], { type })); } return pk.media[t]; }
function bgOf(C) {
  for (const part of [C.slidePath, C.layout, C.master]) { const cSld = P(doc(part).documentElement, 'cSld'), bg = P(cSld, 'bg'); if (!bg) continue;
    const bgPr = P(bg, 'bgPr'); if (bgPr) { const f = fillOf(bgPr, C, null, id => mediaUrl(part, id)); if (f && f !== 'none') return f; }
    const ref = P(bg, 'bgRef'); if (ref) return colorOf(firstColor(ref), C) || '#fff'; }
  return '#fff';
}
function xfrmOf(sp) { const x = A(path1(sp, 'p:spPr') || path1(sp, 'p:grpSpPr'), 'xfrm') || P(sp, 'xfrm'); if (!x) return null; const off = A(x, 'off'), ext = A(x, 'ext'); if (!off || !ext) return null;
  return { x: +off.getAttribute('x'), y: +off.getAttribute('y'), cx: +ext.getAttribute('cx'), cy: +ext.getAttribute('cy'), rot: +(x.getAttribute('rot') || 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1', el: x }; }
// level props for paragraph lvl from the cascade of lstStyle nodes
function lvlNodes(C, sp, chain, lvl) {
  const out = []; const nm = `lvl${lvl + 1}pPr`;
  const push = ls => { const n = ls && A(ls, nm); if (n) out.push(n); };
  push(path1(sp, 'p:txBody', 'a:lstStyle'));
  for (const s of chain) push(path1(s, 'p:txBody', 'a:lstStyle'));
  const ph = phOf(sp); const t = ph ? normT(phType(ph)) : null;
  if (t === 'title') push(P(C.txStyles, 'titleStyle')); else if (t === 'body') push(P(C.txStyles, 'bodyStyle')); else if (ph) push(P(C.txStyles, 'otherStyle'));
  else { push(C.defStyle); push(P(C.txStyles, 'otherStyle')); }
  return out;
}
const pick = (nodes, fn) => { for (const n of nodes) { const v = fn(n); if (v != null) return v; } return null; };
function runStyle(C, rPr, pNodes, defColor, strong) { // pNodes: [pPr, ...lvl nodes]; strong: shape style colour beats inherited text styles
  const rs = [rPr, ...pNodes.map(n => A(n, 'defRPr'))].filter(Boolean);
  const sz = pick(rs, r => r.getAttribute('sz')); const b = pick(rs, r => r.getAttribute('b')), i = pick(rs, r => r.getAttribute('i')), u = pick(rs, r => r.getAttribute('u'));
  const col = pick(strong ? [rPr, ...pNodes.filter(n => n.localName === 'pPr').map(n => A(n, 'defRPr'))].filter(Boolean) : rs, r => { const f = A(r, 'solidFill'); return f ? colorOf(firstColor(f), C) : null; });
  let face = pick(rs, r => A(r, 'latin')?.getAttribute('typeface')); if (face === '+mj-lt') face = C.theme.major; else if (!face || face === '+mn-lt') face = C.theme.minor;
  return { sz: sz ? +sz / 100 : 18, b: b === '1' || b === 'true', i: i === '1' || i === 'true', u: !!u && u !== 'none', color: col || defColor, face, base: rPr?.getAttribute('baseline') };
}
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const fontStack = f => `"${f}",${/light/i.test(f) ? '"Segoe UI Light",' : ''}Carlito,Calibri,"Segoe UI",Helvetica,Arial,sans-serif`;
function textHTML(C, sp, chain, k, fontScale, defColor, hint, strong) {
  const tx = P(sp, 'txBody'); if (!tx) return ''; const ps = kids(tx, NS.a, 'p'); let html = '', num = {};
  const allEmpty = ps.every(p => !desc(p, NS.a, 't').some(t => t.textContent));
  ps.forEach((p, pi) => {
    const pPr = A(p, 'pPr'); const lvl = +(pPr?.getAttribute('lvl') || 0); const nodes = [pPr, ...lvlNodes(C, sp, chain, lvl)].filter(Boolean);
    const algn = pick(nodes, n => n.getAttribute('algn')) || 'l'; const marL = +(pick(nodes, n => n.getAttribute('marL')) || 0), ind = +(pick(nodes, n => n.getAttribute('indent')) || 0);
    let bu = null; const buN = pick(nodes, n => A(n, 'buNone') ? 'none' : A(n, 'buChar') ? 'c:' + A(n, 'buChar').getAttribute('char') : A(n, 'buAutoNum') ? 'n:' + A(n, 'buAutoNum').getAttribute('type') : null);
    const hasText = desc(p, NS.a, 't').some(t => t.textContent);
    if (buN && buN !== 'none' && hasText) { if (buN[0] === 'c') bu = buN.slice(2); else { num[lvl] = (num[lvl] || 0) + 1; bu = /alpha/i.test(buN) ? String.fromCharCode((/UC/.test(buN) ? 64 : 96) + num[lvl]) + (/Paren/.test(buN) ? ')' : '.') : num[lvl] + (/Paren/.test(buN) ? ')' : '.'); } }
    if (!buN || buN === 'none') num = {};
    const spcB = pick(nodes, n => { const s = A(n, 'spcBef'); return s ? (A(s, 'spcPts') ? +A(s, 'spcPts').getAttribute('val') / 100 + 'pt' : A(s, 'spcPct') ? +A(s, 'spcPct').getAttribute('val') / 100000 + 'em' : null) : null; });
    const lnPct = pick(nodes, n => { const s = A(n, 'lnSpc'); return s && A(s, 'spcPct') ? +A(s, 'spcPct').getAttribute('val') / 100000 : null; });
    const first = runStyle(C, A(p, 'r') ? A(A(p, 'r'), 'rPr') : A(p, 'endParaRPr'), nodes, defColor, strong);
    const px = pt => (pt * EMU_PT * k * fontScale).toFixed(2) + 'px';
    let inner = '', ri = 0;
    for (const r of p.children) {
      if (r.localName === 'r' || r.localName === 'fld') { const rix = r.localName === 'r' ? ri++ : 0; const st = runStyle(C, A(r, 'rPr'), nodes, defColor, strong); const t = A(r, 't')?.textContent || ''; if (!t) continue;
        inner += `<span data-ri="${pi}.${rix}" style="font-size:${px(st.sz)};${st.b ? 'font-weight:700;' : ''}${st.i ? 'font-style:italic;' : ''}${st.u ? 'text-decoration:underline;' : ''}color:${st.color};font-family:${fontStack(st.face).replace(/"/g, "'")}${st.base ? ';vertical-align:' + (+st.base > 0 ? 'super' : 'sub') : ''}">${esc(t)}</span>`; }
      else if (r.localName === 'br') inner += '<br>';
    }
    if (!inner) inner = '<br>';
    const ph = hint && allEmpty && pi === 0 ? ` data-hint="${hint}"` : '';
    html += `<p data-pi="${pi}" data-lvl="${lvl}"${bu ? ` data-bu="${esc(bu)}"` : ''}${ph} style="text-align:${{ l: 'left', ctr: 'center', r: 'right', just: 'justify' }[algn] || 'left'};margin:${pi ? spcB || 0 : 0} 0 0 ${marL * k}px;text-indent:${ind * k}px;font-size:${px(first.sz)};color:${first.color};font-family:${fontStack(first.face).replace(/"/g, "'")};line-height:${lnPct ? lnPct * 1.2 : 1.2}${bu ? `;--bw:${-ind * k}px` : ''}">${inner}</p>`;
  });
  return html;
}
function bodyPrOf(sp, chain) { const out = { anchor: 't', l: 91440, t: 45720, r: 91440, b: 45720, fontScale: 1, wrap: true, vert: null };
  for (const s of [...chain].reverse().concat(sp)) { const b = path1(s, 'p:txBody', 'a:bodyPr'); if (!b) continue;
    if (b.getAttribute('anchor')) out.anchor = b.getAttribute('anchor'); for (const [k2, at] of [['l', 'lIns'], ['t', 'tIns'], ['r', 'rIns'], ['b', 'bIns']]) if (b.getAttribute(at) != null) out[k2] = +b.getAttribute(at);
    if (b.getAttribute('wrap')) out.wrap = b.getAttribute('wrap') !== 'none'; const na = A(b, 'normAutofit'); out.fontScale = na?.getAttribute('fontScale') ? +na.getAttribute('fontScale') / 100000 : s === sp && na ? 1 : out.fontScale; }
  return out; }
function renderShape(C, sp, part, k, host, opts) {
  const nm = sp.localName; const interactive = opts.interactive && part === C.slidePath && !opts.inGroup;
  if (nm === 'grpSp') { const gx = xfrmOf(sp); const x = A(path1(sp, 'p:grpSpPr'), 'xfrm'); const ch = x && { x: +A(x, 'chOff')?.getAttribute('x') || 0, y: +A(x, 'chOff')?.getAttribute('y') || 0, cx: +A(x, 'chExt')?.getAttribute('cx') || 1, cy: +A(x, 'chExt')?.getAttribute('cy') || 1 };
    const map = gx && ch ? (b => ({ ...b, x: gx.x + (b.x - ch.x) * gx.cx / ch.cx, y: gx.y + (b.y - ch.y) * gx.cy / ch.cy, cx: b.cx * gx.cx / ch.cx, cy: b.cy * gx.cy / ch.cy })) : (b => b);
    const pm = opts.map || (b => b); for (const c of sp.children) if (['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame'].includes(c.localName)) renderShape(C, c, part, k, host, { ...opts, inGroup: true, map: b => pm(map(b)) }); return; }
  if (!['sp', 'pic', 'cxnSp', 'graphicFrame'].includes(nm)) return;
  const ph = phOf(sp); if (part !== C.slidePath && ph) return; // layout/master placeholders are templates, not content
  if (ph && opts.interactive === false && ['dt', 'ftr', 'sldNum'].includes(phType(ph)) && !desc(sp, NS.a, 't').some(t => t.textContent)) return;
  const chain = ph && part === C.slidePath ? [findPh(C.layout, ph), findPh(C.master, ph, true)].filter(Boolean) : [];
  let box = xfrmOf(sp); for (const s of chain) if (!box) box = xfrmOf(s); if (!box) return; if (opts.map) box = { ...box, ...opts.map(box) };
  const el = document.createElement('div'); el.className = 'pshape';
  el.style.cssText = `left:${box.x * k}px;top:${box.y * k}px;width:${box.cx * k}px;height:${box.cy * k}px;${box.rot || box.flipH || box.flipV ? `transform:rotate(${box.rot}deg)${box.flipH ? ' scaleX(-1)' : ''}${box.flipV ? ' scaleY(-1)' : ''};` : ''}`;
  const spPr = P(sp, 'spPr'), style = P(sp, 'style'); const media = id => mediaUrl(part, id);
  const geom = A(spPr, 'prstGeom')?.getAttribute('prst') || (chain.map(s => A(P(s, 'spPr'), 'prstGeom')?.getAttribute('prst')).find(Boolean)) || 'rect';
  if (nm === 'pic') { const id = path1(sp, 'p:blipFill', 'a:blip')?.getAttributeNS(NS.r, 'embed'); const u = id && media(id); if (u) { const im = document.createElement('img'); im.src = u; im.draggable = false; el.appendChild(im); } else { el.classList.add('pmissing'); el.textContent = 'Picture'; } }
  else if (nm === 'graphicFrame') { const tbl = desc(sp, NS.a, 'tbl')[0];
    if (tbl) { const cols = desc(A(tbl, 'tblGrid'), NS.a, 'gridCol').map(g => +g.getAttribute('w') * k); let h = '<table class="ptbl"><colgroup>' + cols.map(w => `<col style="width:${w}px">`).join('') + '</colgroup>';
      for (const tr of kids(tbl, NS.a, 'tr')) { h += `<tr style="height:${+tr.getAttribute('h') * k}px">`; for (const tc of kids(tr, NS.a, 'tc')) { if (tc.getAttribute('hMerge') || tc.getAttribute('vMerge')) continue; const tcPr = A(tc, 'tcPr'); const f = fillOf(tcPr, C); const txt = kids(A(tc, 'txBody'), NS.a, 'p').map(p => esc(desc(p, NS.a, 't').map(t => t.textContent).join(''))).join('<br>');
        const sz = +(desc(tc, NS.a, 'rPr')[0]?.getAttribute('sz') || 1800) / 100; h += `<td${tc.getAttribute('gridSpan') ? ` colspan="${tc.getAttribute('gridSpan')}"` : ''}${tc.getAttribute('rowSpan') ? ` rowspan="${tc.getAttribute('rowSpan')}"` : ''} style="font-size:${sz * EMU_PT * k}px;${f && f !== 'none' ? 'background:' + f : ''}">${txt}</td>`; } h += '</tr>'; }
      el.innerHTML = h + '</table>'; }
    else { el.classList.add('pmissing'); el.textContent = /chart/.test(path1(sp, 'a:graphic', 'a:graphicData')?.getAttribute('uri') || '') ? 'Chart' : 'Object'; } }
  else {
    let fill = fillOf(spPr, C, A(style, 'fillRef'), media); if (!fill) for (const s of chain) { const f = fillOf(P(s, 'spPr'), C, null); if (f) { fill = f; break; } }
    const ln = A(spPr, 'ln'); let border = null;
    if (ln && !A(ln, 'noFill')) { const c = A(ln, 'solidFill') ? colorOf(firstColor(A(ln, 'solidFill')), C) : (A(style, 'lnRef') ? colorOf(firstColor(A(style, 'lnRef')), C) : null); if (c) border = `${Math.max(1, (+(ln.getAttribute('w') || 12700)) * k)}px solid ${c}`; }
    else if (!ln && style && +(A(style, 'lnRef')?.getAttribute('idx') || 0) > 0) border = `1px solid ${colorOf(firstColor(A(style, 'lnRef')), C)}`;
    if (nm === 'cxnSp') { el.style.height = Math.max(1, box.cy * k) + 'px'; el.style.borderTop = border || '1px solid #444'; }
    else { if (fill && fill !== 'none') el.style.background = fill; if (border) el.style.border = border; el.style.boxSizing = 'border-box';
      if (geom === 'ellipse') el.style.borderRadius = '50%'; else if (/roundRect/.test(geom)) el.style.borderRadius = Math.min(box.cx, box.cy) * k * .1667 + 'px'; }
    if (P(sp, 'txBody')) {
      const bp = bodyPrOf(sp, chain); const fontRef = A(style, 'fontRef'); const defColor = (fontRef && colorOf(firstColor(fontRef), C)) || ('#' + (C.theme.clr[C.map.tx1 || 'dk1'] || '000000'));
      const t = ph ? normT(phType(ph)) : null; const hint = opts.interactive && ph ? (t === 'title' ? 'Click to add title' : phType(ph) === 'subTitle' ? 'Click to add subtitle' : 'Click to add text') : null;
      const tb = document.createElement('div'); tb.className = 'ptb'; tb.style.cssText = `padding:${bp.t * k}px ${bp.r * k}px ${bp.b * k}px ${bp.l * k}px;justify-content:${{ t: 'flex-start', ctr: 'center', b: 'flex-end' }[bp.anchor] || 'flex-start'};${bp.wrap ? '' : 'white-space:nowrap;'}`;
      const tx = document.createElement('div'); tx.className = 'ptx'; tx.innerHTML = textHTML(C, sp, chain, k, bp.fontScale, defColor, hint, !ph && !!(fontRef && firstColor(fontRef))); tx.dataset.fs = bp.fontScale; tb.appendChild(tx); el.appendChild(tb);
      if (geom === 'ellipse' || nm === 'cxnSp') el.style.overflow = 'visible';
    }
  }
  if (interactive) { el.classList.add('pedit'); const rec = { el, node: sp, k, C, chain, text: !!P(sp, 'txBody') }; el._rec = rec; shapes.push(rec);
    const h = document.createElement('i'); h.className = 'prz'; el.appendChild(h); }
  host.appendChild(el);
}
export function renderSlide(path, host, width, interactive) {
  const C = makeCtx(path); const k = width / pk.sz.cx; host.innerHTML = ''; host.style.width = width + 'px'; host.style.height = pk.sz.cy * k + 'px'; host.style.background = bgOf(C);
  const showM = !(P(doc(C.layout).documentElement, 'cSld') && doc(C.layout).documentElement.getAttribute('showMasterSp') === '0') && doc(path).documentElement.getAttribute('showMasterSp') !== '0';
  const tree = part => desc(P(doc(part).documentElement, 'cSld'), NS.p, 'spTree')[0];
  const opts = { interactive };
  if (showM) for (const c of tree(C.master)?.children || []) renderShape(C, c, C.master, k, host, opts);
  if (doc(path).documentElement.getAttribute('showMasterSp') !== '0') for (const c of tree(C.layout)?.children || []) renderShape(C, c, C.layout, k, host, opts);
  for (const c of tree(path)?.children || []) renderShape(C, c, path, k, host, opts);
  return k;
}

// ---------- editor ----------
const curPath = () => pk.slides[cur];
function snap() { flush(); undoS.push({ files: { ...pk.files }, cur }); if (undoS.length > 30) undoS.shift(); redoS = []; }
function restore(s) { pk.files = { ...s.files }; pk.docs = {}; pk.dirtyDocs.clear(); readSlides(); cur = Math.min(s.cur, pk.slides.length - 1); selShape = null; editingShape = null; renderAllP(); changed(); }
function undo() { endEdit(); if (!undoS.length) return ctx.toast('Nothing to undo'); flush(); redoS.push({ files: { ...pk.files }, cur }); restore(undoS.pop()); }
function redo() { endEdit(); if (!redoS.length) return; flush(); undoS.push({ files: { ...pk.files }, cur }); restore(redoS.pop()); }
function fitWidth() { const st = $('#pStageWrap'); const w = st.clientWidth - 24, h = st.clientHeight - 24; return Math.max(200, Math.min(w, h * pk.sz.cx / pk.sz.cy)); }
function renderStage() { shapes = []; const host = $('#pStage'); stageScale = renderSlide(curPath(), host, fitWidth(), true); if (selShape) { const r = shapes.find(s => s.node === selShape.node); selShape = r || null; r?.el.classList.add('sel'); } $('#pCount').textContent = `${cur + 1} / ${pk.slides.length}`; }
function renderThumbs() {
  const t = $('#pThumbs'); t.innerHTML = '';
  pk.slides.forEach((p, i) => { const b = document.createElement('button'); b.className = 'pthumb' + (i === cur ? ' on' : ''); const n = document.createElement('span'); n.textContent = i + 1; const h = document.createElement('div'); h.className = 'pslide'; b.append(n, h);
    try { renderSlide(p, h, 150, false); } catch (e) { console.error(e); } b.onclick = () => { endEdit(); cur = i; selShape = null; renderStage(); $$('.pthumb').forEach((x, j) => x.classList.toggle('on', j === cur)); }; t.appendChild(b); });
  t.children[cur]?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
}
function refreshThumb() { clearTimeout(thumbT); thumbT = setTimeout(() => { const b = $('#pThumbs').children[cur]; const h = b?.querySelector('.pslide'); if (h) try { renderSlide(curPath(), h, 150, false); } catch {} }, 300); }
function renderAllP() { renderStage(); renderThumbs(); }
function select(rec) { $$('.pshape.sel').forEach(e => e.classList.remove('sel')); selShape = rec; rec?.el.classList.add('sel'); updateBar(); }
function updateBar() { const on = !!selShape; $$('#pBar [data-need]').forEach(b => b.disabled = !on); }

// position -> XML
function ensureXfrm(rec) { const sp = rec.node; let x = xfrmOf(sp); if (x && (A(P(sp, 'spPr'), 'xfrm') || P(sp, 'xfrm'))) return x.el;
  let box = null; for (const s of rec.chain) if (!box) box = xfrmOf(s); const d = sp.ownerDocument;
  const spPr = P(sp, 'spPr') || sp.insertBefore(d.createElementNS(NS.p, 'p:spPr'), P(sp, 'style') || P(sp, 'txBody')); const xf = d.createElementNS(NS.a, 'a:xfrm');
  const off = d.createElementNS(NS.a, 'a:off'), ext = d.createElementNS(NS.a, 'a:ext'); off.setAttribute('x', box?.x || 0); off.setAttribute('y', box?.y || 0); ext.setAttribute('cx', box?.cx || 914400); ext.setAttribute('cy', box?.cy || 914400);
  xf.append(off, ext); spPr.insertBefore(xf, spPr.firstChild); return xf; }
function writeBox(rec, b) { const xf = ensureXfrm(rec); A(xf, 'off').setAttribute('x', Math.round(b.x)); A(xf, 'off').setAttribute('y', Math.round(b.y)); A(xf, 'ext').setAttribute('cx', Math.round(b.cx)); A(xf, 'ext').setAttribute('cy', Math.round(b.cy)); touch(curPath()); }

// text DOM -> XML
function rgbToHex(c) { const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); return m ? [m[1], m[2], m[3]].map(v => (+v).toString(16).padStart(2, '0')).join('').toUpperCase() : null; }
function writeText(rec) {
  const sp = rec.node, d = sp.ownerDocument, tx = P(sp, 'txBody'), box = rec.el.querySelector('.ptx'); if (!tx || !box) return;
  const oldPs = kids(tx, NS.a, 'p'); const fs = +box.dataset.fs || 1; const k = rec.k;
  const blocks = [...box.childNodes].filter(n => n.nodeType === 1 && /^(P|DIV)$/.test(n.nodeName)); if (!blocks.length) blocks.push(box);
  const newPs = []; let lastPi = 0;
  for (const bl of blocks) {
    const pi = bl.dataset?.pi != null ? +bl.dataset.pi : lastPi; lastPi = pi; const src = oldPs[pi] || oldPs[oldPs.length - 1];
    const p = d.createElementNS(NS.a, 'a:p'); const sPPr = src && A(src, 'pPr'); if (sPPr) p.appendChild(sPPr.cloneNode(true));
    const lvl = +(sPPr?.getAttribute('lvl') || 0); const nodes = [sPPr, ...lvlNodes(rec.C, sp, rec.chain, lvl)].filter(Boolean);
    const srcRuns = src ? kids(src, NS.a, 'r') : []; const baseRPr = srcRuns[0] ? A(srcRuns[0], 'rPr') : (src && A(src, 'endParaRPr'));
    const def = runStyle(rec.C, null, nodes, '#000');
    const walk = n => {
      if (n.nodeType === 3) { if (!n.textContent) return; const pe = n.parentElement, cs = getComputedStyle(pe); const r = d.createElementNS(NS.a, 'a:r');
        const span = pe.closest('[data-ri]'); const [spi, sri] = span ? span.dataset.ri.split('.').map(Number) : [null, null];
        const origR = span ? kids(oldPs[spi], NS.a, 'r')[sri] || kids(oldPs[spi], NS.a, 'fld')[0] : null; const origRPr = origR ? A(origR, 'rPr') : baseRPr;
        const rp = d.createElementNS(NS.a, 'a:rPr'); if (origRPr) { for (const at of origRPr.attributes) rp.setAttribute(at.name, at.value); [...origRPr.childNodes].forEach(c => rp.appendChild(c.cloneNode(true))); }
        if (!rp.getAttribute('lang')) rp.setAttribute('lang', 'en-US');
        const ref = span ? { b: span.style.fontWeight === '700', i: span.style.fontStyle === 'italic', u: /underline/.test(span.style.textDecoration), sz: parseFloat(span.style.fontSize) / (EMU_PT * k * fs), color: rgbToHex(span.style.color) } : { ...def, color: null };
        const b = +cs.fontWeight >= 600, it = cs.fontStyle === 'italic'; let u = false; for (let e = pe; e && e !== box; e = e.parentElement) if (/underline/.test(getComputedStyle(e).textDecorationLine) || e.nodeName === 'U') u = true;
        if (b !== ref.b) rp.setAttribute('b', b ? '1' : '0'); if (it !== ref.i) rp.setAttribute('i', it ? '1' : '0'); if (u !== ref.u) rp.setAttribute('u', u ? 'sng' : 'none');
        const pt = parseFloat(cs.fontSize) / (EMU_PT * k * fs); if (Math.abs(pt - ref.sz) > .4) rp.setAttribute('sz', Math.round(pt * 2) * 50);
        const hex = rgbToHex(cs.color); if (hex && hex !== ref.color && (span || hex !== rgbToHex(getComputedStyle(bl).color))) { kids(rp, NS.a, 'solidFill').forEach(f => f.remove()); const f = d.createElementNS(NS.a, 'a:solidFill'), c = d.createElementNS(NS.a, 'a:srgbClr'); c.setAttribute('val', hex); f.appendChild(c); const ln = A(rp, 'ln'); rp.insertBefore(f, ln ? ln.nextSibling : rp.firstChild); }
        const t = d.createElementNS(NS.a, 'a:t'); t.textContent = n.textContent.replace(/\u00a0/g, ' '); r.append(rp, t); p.appendChild(r); return; }
      if (n.nodeName === 'BR') { if (n.nextSibling || n.previousSibling) p.appendChild(d.createElementNS(NS.a, 'a:br')); return; }
      n.childNodes.forEach(walk);
    };
    bl.childNodes.forEach(walk);
    while (p.lastChild?.localName === 'br') p.lastChild.remove();
    const end = d.createElementNS(NS.a, 'a:endParaRPr'); end.setAttribute('lang', 'en-US'); const bsz = baseRPr?.getAttribute('sz'); if (bsz) end.setAttribute('sz', bsz); p.appendChild(end);
    newPs.push(p);
  }
  oldPs.forEach(p => p.remove()); newPs.forEach(p => tx.appendChild(p)); touch(curPath());
}
function startEdit(rec, x, y) {
  if (!rec.text) return; select(rec); editingShape = rec; const tx = rec.el.querySelector('.ptx'); rec.el.classList.add('editing');
  tx.querySelectorAll('[data-hint]').forEach(p => p.removeAttribute('data-hint'));
  rec.before = tx.innerHTML; tx.contentEditable = 'true'; try { document.execCommand('defaultParagraphSeparator', false, 'p'); document.execCommand('styleWithCSS', false, true); } catch {}
  tx.focus({ preventScroll: true });
  let range = x != null && document.caretRangeFromPoint ? document.caretRangeFromPoint(x, y) : null;
  if (!range || !tx.contains(range.startContainer)) { range = document.createRange(); range.selectNodeContents(tx); range.collapse(false); }
  const s = getSelection(); s.removeAllRanges(); s.addRange(range);
}
function endEdit() {
  const rec = editingShape; if (!rec) return; editingShape = null; const tx = rec.el.querySelector('.ptx'); tx.contentEditable = 'false'; rec.el.classList.remove('editing');
  if (tx.innerHTML !== rec.before) { snapBefore(rec); writeText(rec); changed(); renderStage(); refreshThumb(); }
}
function snapBefore(rec) { // snapshot without the in-progress DOM change: XML hasn't changed yet, so a normal snap is right
  snap(); }
function eachRun(sp, fn) { for (const p of desc(P(sp, 'txBody'), NS.a, 'p')) { for (const r of kids(p, NS.a, 'r')) { let rp = A(r, 'rPr'); if (!rp) { rp = sp.ownerDocument.createElementNS(NS.a, 'a:rPr'); rp.setAttribute('lang', 'en-US'); r.insertBefore(rp, r.firstChild); } fn(rp, p); } const e = A(p, 'endParaRPr'); if (e) fn(e, p); } }
function boxFormat(kind, val) {
  if (editingShape && ['b', 'i', 'u', 'color'].includes(kind)) { const cmd = { b: 'bold', i: 'italic', u: 'underline', color: 'foreColor' }[kind]; const s = getSelection(); if (s && !s.isCollapsed) { document.execCommand(cmd, false, val); return; } }
  const rec = editingShape || selShape; if (!rec || !rec.text) return; endEdit(); snap(); const sp = rec.node; const C = rec.C;
  if (['b', 'i', 'u'].includes(kind)) { const first = desc(P(sp, 'txBody'), NS.a, 'r')[0]; const p0 = first?.parentNode; const st = first ? runStyle(C, A(first, 'rPr'), [A(p0, 'pPr'), ...lvlNodes(C, sp, rec.chain, +(A(p0, 'pPr')?.getAttribute('lvl') || 0))].filter(Boolean), '#000') : {}; const on = !st[kind];
    eachRun(sp, rp => rp.setAttribute(kind, kind === 'u' ? (on ? 'sng' : 'none') : on ? '1' : '0')); }
  if (kind === 'color') eachRun(sp, rp => { kids(rp, NS.a, 'solidFill').forEach(f => f.remove()); const f = sp.ownerDocument.createElementNS(NS.a, 'a:solidFill'), c = sp.ownerDocument.createElementNS(NS.a, 'a:srgbClr'); c.setAttribute('val', val.replace('#', '').toUpperCase()); f.appendChild(c); rp.insertBefore(f, rp.firstChild); });
  if (kind === 'size') eachRun(sp, (rp, p) => { const st = runStyle(C, rp, [A(p, 'pPr'), ...lvlNodes(C, sp, rec.chain, +(A(p, 'pPr')?.getAttribute('lvl') || 0))].filter(Boolean), '#000'); const n = Math.max(6, Math.min(200, Math.round(st.sz * val))); rp.setAttribute('sz', n * 100); });
  if (kind === 'size') { const na = desc(path1(sp, 'p:txBody', 'a:bodyPr'), NS.a, 'normAutofit')[0]; if (na) { na.removeAttribute('fontScale'); na.removeAttribute('lnSpcReduction'); } }
  if (kind === 'align') for (const p of kids(P(sp, 'txBody'), NS.a, 'p')) { let pp = A(p, 'pPr'); if (!pp) { pp = sp.ownerDocument.createElementNS(NS.a, 'a:pPr'); p.insertBefore(pp, p.firstChild); } pp.setAttribute('algn', val); }
  touch(curPath()); renderStage(); refreshThumb(); changed();
}
function shapeFill(hex) { const rec = selShape; if (!rec || rec.node.localName !== 'sp') return; endEdit(); snap(); const d = rec.node.ownerDocument; ensureXfrm(rec); const spPr = P(rec.node, 'spPr');
  [...spPr.children].filter(c => /^(noFill|solidFill|gradFill|blipFill|pattFill)$/.test(c.localName)).forEach(c => c.remove());
  let f; if (hex === 'none') f = d.createElementNS(NS.a, 'a:noFill'); else { f = d.createElementNS(NS.a, 'a:solidFill'); const c = d.createElementNS(NS.a, 'a:srgbClr'); c.setAttribute('val', hex.replace('#', '').toUpperCase()); f.appendChild(c); }
  const after = A(spPr, 'prstGeom') || A(spPr, 'custGeom') || A(spPr, 'xfrm'); spPr.insertBefore(f, after ? after.nextSibling : spPr.firstChild); touch(curPath()); renderStage(); refreshThumb(); changed(); }
const tree = path => desc(P(doc(path).documentElement, 'cSld'), NS.p, 'spTree')[0];
function nextId(path) { return Math.max(1, ...desc(doc(path).documentElement, NS.p, 'cNvPr').map(c => +c.getAttribute('id') || 0)) + 1; }
function frag(path, xml) { const d = doc(path); const w = new DOMParser().parseFromString(`<w xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}">${xml}</w>`, 'application/xml'); return d.importNode(w.documentElement.firstElementChild, true); }
function addTextBox() { endEdit(); snap(); const path = curPath(), id = nextId(path), cx = pk.sz.cx * .5, cy = pk.sz.cy * .12, x = (pk.sz.cx - cx) / 2, y = (pk.sz.cy - cy) / 2;
  const el = frag(path, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400" dirty="0"/><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`);
  tree(path).appendChild(el); touch(path); renderStage(); refreshThumb(); changed(); const rec = shapes.find(s => s.node === el); if (rec) { startEdit(rec); const s = getSelection(), r = document.createRange(); r.selectNodeContents(rec.el.querySelector('.ptx')); s.removeAllRanges(); s.addRange(r); } }
async function addImage(file) {
  endEdit(); const url = URL.createObjectURL(file); const data = await ctx.toPNG(url, 1600); URL.revokeObjectURL(url);
  const bin = Uint8Array.from(atob(data.split(',')[1]), c => c.charCodeAt(0)); const dims = await new Promise(r => { const im = new Image(); im.onload = () => r([im.width, im.height]); im.src = data; });
  snap(); const path = curPath(), media = `ppt/media/folia_${Date.now().toString(36)}.png`; pk.files[media] = bin; defaultType('png', 'image/png'); const rid = addRel(path, RT.image, media);
  let cx = pk.sz.cx * .4, cy = cx * dims[1] / dims[0]; if (cy > pk.sz.cy * .7) { cy = pk.sz.cy * .7; cx = cy * dims[0] / dims[1]; } const id = nextId(path);
  const el = frag(path, `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((pk.sz.cx - cx) / 2)}" y="${Math.round((pk.sz.cy - cy) / 2)}"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
  tree(path).appendChild(el); touch(path); renderStage(); refreshThumb(); changed(); select(shapes.find(s => s.node === el)); }
function delShape() { const rec = selShape; if (!rec) return; endEdit(); snap(); rec.node.remove(); touch(curPath()); selShape = null; renderStage(); refreshThumb(); changed(); updateBar(); }

// ---------- slides ----------
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
function newSlidePath() { let n = 1; while (pk.files['ppt/slides/slide' + n + '.xml'] || pk.docs['ppt/slides/slide' + n + '.xml']) n++; return 'ppt/slides/slide' + n + '.xml'; }
function insertSldId(rid, afterIdx) { const d = doc(PRES), lst = desc(d.documentElement, NS.p, 'sldIdLst')[0]; const ids = desc(lst, NS.p, 'sldId'); const id = Math.max(255, ...ids.map(s => +s.getAttribute('id'))) + 1;
  const el = d.createElementNS(NS.p, 'p:sldId'); el.setAttribute('id', id); el.setAttributeNS(NS.r, 'r:id', rid); const ref = ids[afterIdx]; lst.insertBefore(el, ref ? ref.nextSibling : null); touch(PRES); }
function createSlide(xml, relsXml, afterIdx) {
  const path = newSlidePath(); pk.docs[path] = new DOMParser().parseFromString(xml, 'application/xml'); touch(path);
  pk.docs[relsPath(path)] = new DOMParser().parseFromString(relsXml, 'application/xml'); touch(relsPath(path));
  contentType(path, SLIDE_CT); const rid = addRel(PRES, RT.slide, path); insertSldId(rid, afterIdx); readSlides(); return path;
}
function layouts() { const m = chainOf(curPath()).master; const d = doc(m); const ids = desc(P(d.documentElement, 'sldLayoutIdLst'), NS.p, 'sldLayoutId'); const rs = rels(m); return ids.map(i => rs[i.getAttributeNS(NS.r, 'id')]?.target).filter(Boolean).map(p => ({ path: p, type: doc(p).documentElement.getAttribute('type'), name: P(doc(p).documentElement, 'cSld')?.getAttribute('name') || '' })); }
function addSlide() {
  endEdit(); snap(); const ls = layouts(); const curL = chainOf(curPath()).layout; const curT = doc(curL).documentElement.getAttribute('type');
  const L = (curT && curT !== 'title' ? ls.find(l => l.path === curL) : null) || ls.find(l => l.type === 'obj') || ls.find(l => l.type !== 'title') || ls[0];
  let id = 2, sps = '';
  for (const sp of desc(tree(L.path), NS.p, 'sp')) { const ph = phOf(sp); if (!ph || ['dt', 'ftr', 'sldNum'].includes(phType(ph))) continue;
    const at = (ph.getAttribute('type') ? ` type="${ph.getAttribute('type')}"` : '') + (ph.getAttribute('idx') ? ` idx="${ph.getAttribute('idx')}"` : '');
    sps += `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${at}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`; id++; }
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const path = newSlidePath(); const rx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${RT.layout}" Target="${relTarget(path, L.path)}"/></Relationships>`;
  createSlide(xml, rx, cur); cur++; selShape = null; renderAllP(); changed();
}
function dupSlide() { endEdit(); snap(); flush(); const src = curPath(); const xml = dec.decode(pk.files[src]); const rd = doc(relsPath(src)); const rx = rd ? new XMLSerializer().serializeToString(rd) : `<Relationships xmlns="${NS.rel}"/>`;
  const path = createSlide(xml, rx, cur); desc(doc(relsPath(path)).documentElement, NS.rel, 'Relationship').filter(r => r.getAttribute('Type') === RT.notes).forEach(r => r.remove()); touch(relsPath(path)); cur++; renderAllP(); changed(); }
function delSlide() {
  if (pk.slides.length < 2) return ctx.toast('A presentation needs at least one slide'); endEdit(); snap(); const path = curPath();
  const d = doc(PRES), rs = rels(PRES); const rid = Object.keys(rs).find(k => rs[k].target === path);
  desc(d.documentElement, NS.p, 'sldId').filter(s => s.getAttributeNS(NS.r, 'id') === rid).forEach(s => s.remove()); rs[rid]?.el.remove(); touch(PRES); touch(relsPath(PRES));
  const notes = relOfType(path, RT.notes); for (const p of [path, notes].filter(Boolean)) { delete pk.files[p]; delete pk.docs[p]; delete pk.files[relsPath(p)]; delete pk.docs[relsPath(p)]; pk.dirtyDocs.delete(p); pk.dirtyDocs.delete(relsPath(p)); dropContentType(p); }
  readSlides(); cur = Math.min(cur, pk.slides.length - 1); selShape = null; renderAllP(); changed();
}
function moveSlide(dir) { const j = cur + dir; if (j < 0 || j >= pk.slides.length) return; endEdit(); snap(); const lst = desc(doc(PRES).documentElement, NS.p, 'sldIdLst')[0]; const ids = desc(lst, NS.p, 'sldId');
  if (dir < 0) lst.insertBefore(ids[cur], ids[j]); else lst.insertBefore(ids[j], ids[cur]); touch(PRES); readSlides(); cur = j; renderAllP(); changed(); }

// ---------- save ----------
async function autosave() { clearTimeout(saveT); saveT = null; if (!dirty || !pid) return; try { const b = packBytes(); dirty = false; $('#pState').textContent = 'Saved'; await ctx.touchRecent(pid, pname, b, 'pptx'); } catch (e) { console.error(e); $('#pState').textContent = 'Not saved'; } }
function changed() { dirty = true; $('#pState').textContent = 'Editing…'; clearTimeout(saveT); saveT = setTimeout(autosave, 2500); }
function printDeck() {
  endEdit(); const host = $('#pPrint'); host.innerHTML = ''; const W = 1000; pk.slides.forEach(p => { const s = document.createElement('div'); s.className = 'pslide pprintpage'; host.appendChild(s); renderSlide(p, s, W, false); });
  const st = document.createElement('style'); st.id = 'pPageSize'; st.textContent = `@page{size:${pk.sz.cx / 914400}in ${pk.sz.cy / 914400}in;margin:0}`; document.head.appendChild(st);
  $$('#pPrint .pslide').forEach(s => { s.style.zoom = (pk.sz.cx / 914400 * 96) / W; });
  const t = document.title; document.title = pname.replace(/\.pptx$/i, ''); document.body.classList.add('printing-p');
  const done = () => { document.body.classList.remove('printing-p'); document.title = t; host.innerHTML = ''; st.remove(); }; window.addEventListener('afterprint', done, { once: true }); setTimeout(() => window.print(), 300); setTimeout(done, 60000);
}
async function exportAs(kind) { endEdit(); if (kind === 'pdf') return printDeck(); ctx.busy(true, 'Saving…');
  try { const b = packBytes(); ctx.busy(false); const ok = await ctx.deliver(b, pname, PPTX_MIME); if (ok) { dirty = false; ctx.touchRecent(pid, pname, b, 'pptx'); $('#pState').textContent = 'Saved'; } }
  catch (e) { console.error(e); ctx.busy(false); ctx.toast("Couldn't save: " + (e.message || e)); } }

// ---------- open ----------
function show(name, id) {
  pname = name.replace(/\.pptx$/i, '') + '.pptx'; pid = id; dirty = false; undoS = []; redoS = []; cur = 0; selShape = null; editingShape = null;
  $('#pName').value = pname.replace(/\.pptx$/i, ''); $('#pState').textContent = 'Saved';
  ['#home', '#editor', '#wordEd', '#xlEd'].forEach(s => $(s)?.classList.add('hidden')); $('#pptEd').classList.remove('hidden');
  requestAnimationFrame(() => { renderAllP(); updateBar(); });
}
export async function openPpt(file, id) {
  ctx.busy(true, 'Opening…');
  try { await libs(); Object.values(pk?.media || {}).forEach(u => URL.revokeObjectURL(u)); openPackage(file.bytes); if (!pk.slides.length) throw new Error('No slides');
    const nid = id || Math.random().toString(36).slice(2, 10); show(file.name, nid); if (!id) ctx.touchRecent(nid, pname, file.bytes, 'pptx'); }
  catch (e) { console.error(e); ctx.toast("Couldn't open that presentation."); }
  ctx.busy(false);
}
export async function newPpt() {
  ctx.busy(true, 'Starting…');
  try { await libs(); const r = await fetch('vendor/blank.pptx'); openPackage(new Uint8Array(await r.arrayBuffer()));
    // empty the placeholders so they show "Click to add…" hints
    for (const r of desc(doc(pk.slides[0]).documentElement, NS.a, 'r')) r.remove(); touch(pk.slides[0]);
    show('Untitled.pptx', Math.random().toString(36).slice(2, 10)); dirty = true; await autosave(); }
  catch (e) { console.error(e); ctx.toast("Couldn't start a presentation."); }
  ctx.busy(false);
}

// ---------- wiring ----------
export function initPpt(c) {
  ctx = c; const stage = $('#pStage');
  let drag = null;
  stage.addEventListener('pointerdown', e => {
    const el = e.target.closest('.pshape.pedit'); const rec = el?._rec;
    if (editingShape && rec === editingShape && !e.target.classList.contains('prz')) return; // let the caret move
    if (!rec) { endEdit(); select(null); return; }
    if (editingShape) endEdit();
    const r2 = shapes.find(s => s.node === rec.node) || rec; const wasSel = selShape?.node === r2.node; select(r2);
    const box = xfrmOf(r2.node) || r2.chain.map(xfrmOf).find(Boolean); if (!box) return;
    drag = { rec: r2, x0: e.clientX, y0: e.clientY, box: { ...box }, resize: e.target.classList.contains('prz'), moved: false, wasSel, id: e.pointerId };
    stage.setPointerCapture(e.pointerId); e.preventDefault();
  });
  stage.addEventListener('pointermove', e => { if (!drag) return; const dx = (e.clientX - drag.x0) / stageScale, dy = (e.clientY - drag.y0) / stageScale;
    if (!drag.moved && Math.hypot(e.clientX - drag.x0, e.clientY - drag.y0) < 4) return; if (!drag.moved) { drag.moved = true; snap(); }
    const el = drag.rec.el, b = drag.box, k = stageScale;
    if (drag.resize) { const lock = drag.rec.node.localName === 'pic'; let cx = Math.max(91440, b.cx + dx), cy = Math.max(91440, lock ? cx * b.cy / b.cx : b.cy + dy); drag.nb = { ...b, cx, cy }; el.style.width = cx * k + 'px'; el.style.height = cy * k + 'px'; }
    else { drag.nb = { ...b, x: b.x + dx, y: b.y + dy }; el.style.left = drag.nb.x * k + 'px'; el.style.top = drag.nb.y * k + 'px'; } });
  stage.addEventListener('pointerup', e => { const d = drag; drag = null; if (!d) return;
    if (d.moved && d.nb) { writeBox(d.rec, d.nb); changed(); renderStage(); refreshThumb(); }
    else if (d.wasSel && d.rec.text) { const r = shapes.find(s => s.node === d.rec.node); if (r) startEdit(r, e.clientX, e.clientY); } });
  stage.addEventListener('dblclick', e => { const rec = e.target.closest('.pshape.pedit')?._rec; if (rec && !editingShape) startEdit(rec, e.clientX, e.clientY); });
  $('#pStageWrap').addEventListener('pointerdown', e => { if (e.target.id === 'pStageWrap') { endEdit(); select(null); } });
  stage.addEventListener('keydown', e => { if (editingShape && e.key === 'Escape') { e.preventDefault(); endEdit(); } });
  document.addEventListener('keydown', e => {
    if ($('#pptEd').classList.contains('hidden') || editingShape || document.activeElement?.matches('input,textarea,select,[contenteditable="true"]')) return;
    const m = e.ctrlKey || e.metaKey;
    if (m && /^[zZ]$/.test(e.key)) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; } if (m && e.key === 'y') { e.preventDefault(); redo(); return; }
    if (!selShape) { if (e.key === 'PageDown' || e.key === 'ArrowDown' || e.key === 'ArrowRight') { if (cur < pk.slides.length - 1) { cur++; renderStage(); renderThumbs(); } } else if (e.key === 'PageUp' || e.key === 'ArrowUp' || e.key === 'ArrowLeft') { if (cur > 0) { cur--; renderStage(); renderThumbs(); } } return; }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); delShape(); return; }
    if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startEdit(selShape); return; }
    const mv = { ArrowDown: [0, 1], ArrowUp: [0, -1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[e.key];
    if (mv) { e.preventDefault(); const b = xfrmOf(selShape.node) || selShape.chain.map(xfrmOf).find(Boolean); if (!b) return; snap(); const st = (e.shiftKey ? 10 : 1) * 91440 / 2; writeBox(selShape, { ...b, x: b.x + mv[0] * st, y: b.y + mv[1] * st }); renderStage(); refreshThumb(); changed(); }
  });
  $$('#pBar button, #pTop .ib').forEach(b => b.addEventListener('pointerdown', e => { if (editingShape) e.preventDefault(); }));
  $('#pUndo').onclick = undo; $('#pRedo').onclick = redo;
  $('#pText').onclick = addTextBox; $('#pImg').onclick = () => { endEdit(); $('#pImgInput').click(); };
  $('#pImgInput').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) addImage(f); };
  $('#pB').onclick = () => boxFormat('b'); $('#pI').onclick = () => boxFormat('i'); $('#pU').onclick = () => boxFormat('u');
  $('#pBig').onclick = () => boxFormat('size', 1.15); $('#pSmall').onclick = () => boxFormat('size', 1 / 1.15);
  $$('#pAlign button').forEach(b => b.onclick = () => boxFormat('align', b.dataset.al));
  $$('#pColor button').forEach(b => b.onclick = () => boxFormat('color', b.dataset.c));
  $$('#pFill button').forEach(b => b.onclick = () => shapeFill(b.dataset.c));
  $('#pDel').onclick = delShape;
  $('#pAddSlide').onclick = addSlide; $('#pSlideMenu').onclick = () => { endEdit(); $('#pSlideSheet').classList.remove('hidden'); };
  $('#pDup').onclick = () => { $('#pSlideSheet').classList.add('hidden'); dupSlide(); }; $('#pDelSlide').onclick = () => { $('#pSlideSheet').classList.add('hidden'); if (confirm('Delete this slide?')) delSlide(); };
  $('#pUp').onclick = () => { $('#pSlideSheet').classList.add('hidden'); moveSlide(-1); }; $('#pDown').onclick = () => { $('#pSlideSheet').classList.add('hidden'); moveSlide(1); };
  $('#pPrev').onclick = () => { endEdit(); if (cur > 0) { cur--; selShape = null; renderStage(); renderThumbs(); } }; $('#pNext').onclick = () => { endEdit(); if (cur < pk.slides.length - 1) { cur++; selShape = null; renderStage(); renderThumbs(); } };
  $('#pName').addEventListener('change', e => { const v = e.target.value.trim().replace(/[\\/:*?"<>|]/g, '') || 'Untitled'; e.target.value = v; pname = v + '.pptx'; dirty = true; autosave(); });
  $('#pSave').onclick = () => { endEdit(); $('#pSaveSheet').classList.remove('hidden'); };
  $$('#pSaveSheet [data-x]').forEach(b => b.onclick = () => { $('#pSaveSheet').classList.add('hidden'); exportAs(b.dataset.x); });
  $('#pBack').onclick = async () => { endEdit(); await autosave(); $('#pptEd').classList.add('hidden'); $('#home').classList.remove('hidden'); ctx.home(); };
  let rz; window.addEventListener('resize', () => { if ($('#pptEd').classList.contains('hidden')) return; clearTimeout(rz); rz = setTimeout(() => { endEdit(); renderStage(); }, 150); });
  window.addEventListener('beforeunload', () => { if (dirty) autosave(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && dirty) { endEdit(); autosave(); } });
}
export const pptState = { get pk() { return pk; }, get cur() { return cur; } };
