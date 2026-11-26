// Folia inline text editor: tap a line of text on the PDF and edit it right where it sits.
// Digital PDFs: text comes from the PDF itself. Scanned pages: text is read automatically with OCR.
// On save, each changed line is covered in its own background colour and redrawn with the PDF's own
// font when it can be reused, else the closest bundled font (Lato, Noto Sans, Noto Serif), per character.
const { PDFDocument, PDFName, PDFDict, rgb, decodePDFRawStream } = window.PDFLib;
const V = new URL('vendor/', location.href).href;

const loadScript = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not download ' + src.split('/').pop())); document.head.appendChild(s); });
let fkP = null; const getFontkit = () => fkP ||= (window.fontkit ? Promise.resolve(window.fontkit) : loadScript(V + 'fontkit.umd.min.js').then(() => window.fontkit)).catch(e => { fkP = null; throw e; });
const fileCache = {}; const fetchFont = f => fileCache[f] ||= fetch(V + 'fonts/' + f).then(r => { if (!r.ok) throw new Error('font ' + f); return r.arrayBuffer(); }).then(b => new Uint8Array(b));

const strip = n => String(n || '').replace(/^\//, '').replace(/^[A-Z]{6}\+/, '');
const isSerif = n => /serif|times|roman|georgia|garamond|cmr|lmr|minion|cambria|book|palatino|baskerville/i.test(n) && !/sans/i.test(n);
const isBold = n => /bold|black|heavy|semibold|demi/i.test(n);
const isItalic = n => /italic|oblique/i.test(n);

// ---------- text blocks ----------
function groupLines(items) {
  // items: {s, x, x2, y (baseline), size, font, base}
  items = items.filter(it => it.size > 1).sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  for (const it of items) {
    let L = null;
    for (let k = lines.length - 1; k >= 0 && k >= lines.length - 4; k--) { const l = lines[k]; if (Math.abs(l.y - it.y) < Math.min(l.size, it.size) * 0.4 && it.size / l.size > 0.6 && it.size / l.size < 1.6) { L = l; break; } }
    if (!L) lines.push(L = { y: it.y, size: it.size, items: [] });
    L.items.push(it);
  }
  const blocks = [];
  for (const l of lines) {
    l.items.sort((a, b) => a.x - b.x);
    let cur = null, space = false;
    for (const it of l.items) {
      if (!it.s.trim()) { if (cur) space = true; continue; }
      const gap = cur ? it.x - cur.x2 : 0;
      if (cur && gap < cur.size * 1.1 && gap > -cur.size * 0.6) {
        if ((space || gap > cur.size * 0.12) && !cur.text.endsWith(' ') && !it.s.startsWith(' ')) cur.text += ' ';
        cur.text += it.s; cur.x2 = Math.max(cur.x2, it.x2); cur.size = Math.max(cur.size, it.size);
      } else {
        cur = { text: it.s, x: it.x, x2: it.x2, y: l.y, size: it.size, font: it.font, base: it.base, ocr: it.ocr };
        blocks.push(cur);
      }
      space = /\s$/.test(it.s);
    }
  }
  for (const b of blocks) { b.text = b.text.replace(/\s+$/, '').replace(/^\s+/, ''); b.orig = b.text; }
  return blocks.filter(b => b.text);
}

// colours from the rendered canvas inside a block's box (canvas px)
function sampleColours(ctx, x, y, w, h) {
  x = Math.max(0, Math.floor(x)); y = Math.max(0, Math.floor(y)); w = Math.max(2, Math.ceil(w)); h = Math.max(2, Math.ceil(h));
  let d; try { d = ctx.getImageData(x, y, w, h).data; } catch { return { bg: [1, 1, 1], fg: [0, 0, 0] }; }
  const px = i => [d[i], d[i + 1], d[i + 2]];
  const border = []; for (let i = 0; i < w; i++) { border.push(px(i * 4), px(((h - 1) * w + i) * 4)); } for (let j = 0; j < h; j++) { border.push(px(j * w * 4), px((j * w + w - 1) * 4)); }
  const med = arr => { const s = [...arr].sort((a, b) => a - b); return s[s.length >> 1]; };
  const bg = [0, 1, 2].map(c => med(border.map(p => p[c])));
  let maxD = 0; const dist = []; for (let i = 0; i < d.length; i += 4) { const dd = Math.abs(d[i] - bg[0]) + Math.abs(d[i + 1] - bg[1]) + Math.abs(d[i + 2] - bg[2]); dist.push(dd); if (dd > maxD) maxD = dd; }
  let fg = [0, 0, 0];
  if (maxD > 60) { let n = 0, s = [0, 0, 0]; for (let k = 0; k < dist.length; k++) if (dist[k] > maxD * 0.9) { s[0] += d[k * 4]; s[1] += d[k * 4 + 1]; s[2] += d[k * 4 + 2]; n++; } fg = s.map(v => v / n); }
  return { bg: bg.map(v => v / 255), fg: fg.map(v => v / 255) };
}
const css = c => `rgb(${c.map(v => Math.round(v * 255)).join(',')})`;

// ---------- the editor ----------
export async function openInline({ host, bytes, pdfjsLib, getOcr, status }) {
  const js = await pdfjsLib.getDocument({ data: bytes.slice(), fontExtraProperties: true }).promise;
  const pages = [];
  for (let i = 0; i < js.numPages; i++) { const pg = await js.getPage(i + 1); const vp = pg.getViewport({ scale: 1 }); pages.push({ pg, i, w: vp.width, h: vp.height, blocks: null, el: null }); }
  let zoom = 1, active = null, dirty = false, destroyed = false;
  const undo = [];
  host.innerHTML = ''; host.classList.add('ilhost');
  const wrap = document.createElement('div'); wrap.className = 'ilwrap'; host.appendChild(wrap);
  const fit = () => Math.max(0.3, (Math.min(host.clientWidth || 360, 1100) - 16) / pages[0].w);

  async function buildBlocks(P, canvas, scale) {
    const tc = await P.pg.getTextContent();
    const chars = tc.items.map(t => t.str || '').join('').replace(/\s/g, '').length;
    let items = [];
    if (chars >= 15) {
      for (const it of tc.items) {
        if (typeof it.str !== 'string' || !it.str) continue;
        const [a, b, c, d, e, f] = it.transform;
        if (Math.abs(b) > 0.01 || Math.abs(c) > 0.01 || a <= 0) continue; // rotated text: not editable here
        const size = Math.abs(d) || Math.hypot(a, b);
        let base = ''; try { const fo = P.pg.commonObjs.get(it.fontName); base = strip(fo?.name || ''); } catch {}
        items.push({ s: it.str, x: e, x2: e + it.width, y: f, size, font: it.fontName, base, family: tc.styles[it.fontName]?.fontFamily || 'sans-serif' });
      }
    } else {
      status('Reading the scanned page ' + (P.i + 1) + '…');
      const O = await getOcr();
      const s = Math.min(2.5, 1700 / Math.max(P.w, P.h)), vp = P.pg.getViewport({ scale: s });
      const c = document.createElement('canvas'); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
      await P.pg.render({ canvasContext: x, viewport: vp }).promise;
      const r = await O.recognize(c, m => m.status === 'recognizing text' && status(`Reading the scanned page ${P.i + 1}… ${Math.round(m.progress * 100)}%`));
      for (const l of r.lines) {
        // a line's box runs from the tallest letter to the lowest descender, so the font size depends on what letters it has
        const lh = (l.y1 - l.y0) / s, desc = /[gjpqy,;()\[\]]/.test(l.text), size = desc ? lh / 1.0 : lh / 0.82;
        for (const w of l.words) { const [x0, y0] = vp.convertToPdfPoint(w.x0, l.y1); const [x1] = vp.convertToPdfPoint(w.x1, l.y1); items.push({ s: w.text, x: x0, x2: x1, y: y0 + (desc ? size * 0.22 : 0), size, ocr: true, base: '', family: 'sans-serif' }); }
      }
      status('');
    }
    const blocks = groupLines(items);
    const ctx = canvas.getContext('2d'), k = canvas.width / (P.w * scale);
    for (const b of blocks) {
      const [vx, vy] = P.vp0.convertToViewportPoint(b.x, b.y);
      const top = vy - b.size * 0.85, wpx = (b.x2 - b.x), hpx = b.size * 1.12;
      Object.assign(b, sampleColours(ctx, vx * scale * k, top * scale * k, wpx * scale * k, hpx * scale * k));
      const it = items.find(i => i.font === b.font); b.family = it?.family || 'sans-serif';
      b.bold = isBold(b.base); b.italic = isItalic(b.base); b.serif = isSerif(b.base) || /serif/.test(b.family) && !/sans/.test(b.family);
    }
    return blocks;
  }

  async function renderPage(P) {
    const scale = zoom * fit(), dpr = Math.min(window.devicePixelRatio || 1, 2.5);
    P.vp0 = P.pg.getViewport({ scale: 1 });
    const vp = P.pg.getViewport({ scale: scale * dpr });
    const el = P.el || document.createElement('div'); el.className = 'ilpage'; el.innerHTML = '';
    el.style.width = P.w * scale + 'px'; el.style.height = P.h * scale + 'px';
    const c = document.createElement('canvas'); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height); el.appendChild(c);
    if (!P.el) { P.el = el; wrap.appendChild(el); }
    await P.pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    if (!P.blocks) P.blocks = await buildBlocks(P, c, scale);
    const ov = document.createElement('div'); ov.className = 'ilov'; el.appendChild(ov);
    P.blocks.forEach((b, n) => {
      const [vx, vy] = P.vp0.convertToViewportPoint(b.x, b.y);
      const d = document.createElement('div'); d.className = 'ilb' + (b.text !== b.orig ? ' changed' : ''); d.dataset.p = P.i; d.dataset.n = n;
      const fs = b.size * scale;
      Object.assign(d.style, { left: vx * scale - 2 + 'px', top: (vy - b.size * 0.9) * scale + 'px', minWidth: (b.x2 - b.x) * scale + 4 + 'px', height: b.size * 1.15 * scale + 'px', fontSize: fs + 'px', lineHeight: b.size * 1.15 * scale + 'px',
        fontFamily: `"${b.font || 'x'}", ${b.serif ? '"Noto Serif", Georgia, "Times New Roman", serif' : 'Lato, "Helvetica Neue", Helvetica, Arial, sans-serif'}`,
        fontWeight: b.bold ? '700' : '400', fontStyle: b.italic ? 'italic' : 'normal' });
      d.style.setProperty('--fg', css(b.fg)); d.style.setProperty('--bg', css(b.bg));
      d.textContent = b.text;
      ov.appendChild(d);
    });
  }
  async function renderAll() { for (const P of pages) { if (destroyed) return; await renderPage(P); } }

  // tap a line: it becomes an editable field in place, original covered, caret where you tapped
  function startEdit(d, x, y) {
    if (active === d) return;
    if (active) stopEdit();
    active = d; d.classList.add('editing'); d.contentEditable = 'true'; d.spellcheck = false; d.setAttribute('autocapitalize', 'off'); d.setAttribute('autocorrect', 'off');
    d.focus({ preventScroll: true });
    let r = null; try { r = document.caretRangeFromPoint ? document.caretRangeFromPoint(x, y) : null; } catch {}
    if (!r || !d.contains(r.startContainer)) { r = document.createRange(); r.selectNodeContents(d); r.collapse(false); }
    const s = getSelection(); s.removeAllRanges(); s.addRange(r);
  }
  function stopEdit() {
    const d = active; if (!d) return; active = null;
    d.contentEditable = 'false'; d.classList.remove('editing');
    const b = pages[+d.dataset.p].blocks[+d.dataset.n]; const t = d.textContent.replace(/\s+/g, ' ').trim();
    if (t !== b.text) { undo.push([b, b.text]); b.text = t; dirty = true; }
    d.textContent = b.text; d.classList.toggle('changed', b.text !== b.orig);
  }
  let t0 = null;
  const onDown = e => { t0 = { x: e.clientX, y: e.clientY, t: Date.now() }; };
  const onUp = e => {
    const d = e.target.closest?.('.ilb'); if (!t0) return; const moved = Math.hypot(e.clientX - t0.x, e.clientY - t0.y) > 10; t0 = null;
    if (!d) { if (!moved && active && !e.target.closest('.ilb')) { stopEdit(); document.activeElement?.blur?.(); } return; }
    if (moved || d === active) return;
    e.preventDefault(); startEdit(d, e.clientX, e.clientY);
  };
  wrap.addEventListener('pointerdown', onDown); wrap.addEventListener('pointerup', onUp);
  wrap.addEventListener('click', e => { const d = e.target.closest?.('.ilb'); if (!d) return; if (d !== active) startEdit(d, e.clientX, e.clientY); else if (document.activeElement !== d) d.focus({ preventScroll: true }); });
  // iOS raises the keyboard most reliably from touchend, so focus there too
  wrap.addEventListener('touchend', e => { const d = e.target.closest?.('.ilb'); if (d && d === active && document.activeElement !== d) d.focus({ preventScroll: true }); }, { passive: true });
  wrap.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === 'Escape') { e.preventDefault(); stopEdit(); e.target.blur(); } });
  wrap.addEventListener('focusout', e => { if (active && e.target === active) setTimeout(() => { if (active && document.activeElement !== active) stopEdit(); }, 0); });
  wrap.addEventListener('paste', e => { if (!active) return; e.preventDefault(); const t = (e.clipboardData?.getData('text/plain') || '').replace(/\s+/g, ' '); document.execCommand('insertText', false, t); });

  await renderAll();

  async function getBytes() {
    stopEdit();
    const fontkit = await getFontkit();
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true }); doc.registerFontkit(fontkit);
    // the PDF's own font programs, by base name
    const own = {};
    for (const [, obj] of doc.context.enumerateIndirectObjects()) {
      if (!(obj instanceof PDFDict) || obj.get(PDFName.of('Type'))?.toString() !== '/FontDescriptor') continue;
      const f = obj.lookup(PDFName.of('FontFile2')) || obj.lookup(PDFName.of('FontFile3')); if (!f) continue;
      try { const raw = decodePDFRawStream(f).decode(); const fk = fontkit.create(raw); if (fk && fk.characterSet?.length) own[strip(obj.get(PDFName.of('FontName')))] = { raw, fk }; } catch {}
    }
    const emb = new Map();
    const embed = async (key, raw) => { if (!emb.has(key)) emb.set(key, doc.embedFont(raw, { subset: !/^Noto/.test(key) })); return emb.get(key); };
    const fkCache = {}; const fkOf = async f => fkCache[f] ||= fontkit.create(await fetchFont(f));
    const has = (fk, ch) => { try { return fk.glyphForCodePoint(ch.codePointAt(0)).id > 0; } catch { return false; } };
    for (const P of pages) for (const b of P.blocks || []) {
      if (b.text === b.orig) continue;
      const page = doc.getPage(P.i);
      page.drawRectangle({ x: b.x - 1, y: b.y - b.size * 0.28, width: (b.x2 - b.x) + 2, height: b.size * 1.22, color: rgb(...b.bg) });
      if (!b.text) continue;
      // font chain: the PDF's own font, then a bundled one in the same style, then Noto
      const chain = [];
      if (own[b.base]) chain.push({ key: 'own:' + b.base, fk: own[b.base].fk, raw: own[b.base].raw });
      const style = b.bold ? 'Bold' : 'Regular';
      const firstF = /lato/i.test(b.base) || (!b.serif && !b.base) ? `Lato-${style}.ttf` : b.serif ? `NotoSerif-${style}.ttf` : `NotoSans-${style}.ttf`;
      for (const f of [firstF, `NotoSans-${style}.ttf`]) if (!chain.some(c => c.key === f)) chain.push({ key: f, fk: await fkOf(f), raw: await fetchFont(f) });
      // split the text into runs by the first font in the chain that has each character
      const runs = []; for (const ch of b.text) { const c = chain.find(c => ch === ' ' || has(c.fk, ch)) || chain[chain.length - 1]; const last = runs[runs.length - 1]; if (last && last.c === c) last.t += ch; else runs.push({ c, t: ch }); }
      let x = b.x;
      for (const r of runs) { const font = await embed(r.c.key, r.c.raw); page.drawText(r.t, { x, y: b.y, size: b.size, font, color: rgb(...b.fg) }); x += font.widthOfTextAtSize(r.t, b.size); }
    }
    return doc.save();
  }

  return {
    isDirty: () => { if (active) stopEdit(); return dirty || pages.some(P => (P.blocks || []).some(b => b.text !== b.orig)); },
    getBytes,
    undo: () => { stopEdit(); const u = undo.pop(); if (!u) return false; u[0].text = u[1]; renderAll(); return true; },
    zoom: async dz => { stopEdit(); zoom = Math.max(0.6, Math.min(4, zoom * dz)); await renderAll(); return zoom; },
    count: () => pages.reduce((n, P) => n + (P.blocks?.length || 0), 0),
    destroy: () => { destroyed = true; host.innerHTML = ''; host.classList.remove('ilhost'); js.destroy(); },
  };
}
