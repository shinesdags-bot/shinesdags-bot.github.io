// Folia OCR: text recognition with Tesseract.js 5 (Apache-2.0), all on the device, no server.
// Adds an invisible, selectable text layer to scanned PDF pages (so Edit text, search and copy work),
// and turns photos into searchable PDFs.
const { PDFDocument, StandardFonts, rgb } = window.PDFLib;
const base = new URL('vendor/ocr/', location.href).href;
let workerP = null;

function loadScript(src) {
  return new Promise((res, rej) => { if (window.Tesseract) return res(); const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not download the text recognizer')); document.head.appendChild(s); });
}

export async function getWorker(onProgress) {
  if (!workerP) workerP = (async () => {
    let err;
    for (let k = 0; k < 3; k++) { try { await loadScript(base + 'tesseract.min.js'); break; } catch (e) { err = e; await new Promise(r => setTimeout(r, 1200)); } }
    if (!window.Tesseract) throw err;
    return window.Tesseract.createWorker('eng', 1, {
      workerPath: base + 'worker.min.js', corePath: base, langPath: base.replace(/\/$/, ''), gzip: true, workerBlobURL: false,
      logger: m => getWorker.cb && getWorker.cb(m),
    });
  })().catch(e => { workerP = null; throw e; });
  getWorker.cb = onProgress;
  return workerP;
}

// canvas/img/blob -> [{text, x0, y0, x1, y1}] lines in image pixels, plus the full text
export async function recognize(image, onProgress) {
  const w = await getWorker(onProgress);
  const { data } = await w.recognize(image, {}, { text: true, blocks: true });
  const lines = [];
  for (const b of data.blocks || []) for (const p of b.paragraphs || []) for (const l of p.lines || []) {
    const words = (l.words || []).filter(x => x.text.trim() && x.confidence > 20);
    if (words.length) lines.push({ words: words.map(x => ({ text: x.text, ...x.bbox })), text: words.map(x => x.text).join(' '), ...l.bbox });
  }
  return { lines, text: data.text || '' };
}

const clean = s => [...s].map(c => { const n = c.codePointAt(0); if ((n >= 32 && n < 127) || (n >= 160 && n <= 255)) return c; return ({ '‘': "'", '’': "'", '“': '"', '”': '"', '–': '-', '—': '-', '…': '...', '•': '-', '€': 'EUR', '₦': 'N', '|': 'I' })[c] || ''; }).join('');

// Draw recognized lines as invisible text on a pdf-lib page. k = image px per PDF point.
function layText(page, font, lines, k, ox = 0, oy = 0, H) {
  for (const l of lines) {
    const t = clean(l.text); if (!t.trim()) continue;
    const hPt = (l.y1 - l.y0) / k, wPt = (l.x1 - l.x0) / k;
    let size = Math.max(4, hPt * 0.78); const tw = font.widthOfTextAtSize(t, size);
    if (tw > 0) size = Math.max(4, Math.min(size * 1.25, size * wPt / tw));
    const x = ox + l.x0 / k, y = oy + H - l.y1 / k + hPt * 0.2;
    page.drawText(t, { x, y, size, font, color: rgb(0, 0, 0), opacity: 0 });
  }
}

const renderPage = async (js, i, maxPx = 2200) => {
  const pg = await js.getPage(i + 1);
  const v1 = pg.getViewport({ scale: 1, rotation: 0 });
  const s = Math.min(3, maxPx / Math.max(v1.width, v1.height));
  const vp = pg.getViewport({ scale: s, rotation: 0 });
  const c = document.createElement('canvas'); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
  await pg.render({ canvasContext: x, viewport: vp }).promise;
  return { c, s, view: pg.view };
};

// Add a text layer to every page that has (almost) no real text. Returns { bytes, text, done }.
export async function ocrPdf(bytes, pdfjsLib, status) {
  const js = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const out = []; let done = 0;
  for (let i = 0; i < js.numPages; i++) {
    const tc = await (await js.getPage(i + 1)).getTextContent();
    const have = tc.items.map(t => t.str).join('').replace(/\s/g, '').length;
    if (have > 40) { out.push(tc.items.map(t => t.str).join(' ')); continue; }
    status(`Reading page ${i + 1} of ${js.numPages}…`);
    const { c, s, view } = await renderPage(js, i);
    const r = await recognize(c, m => m.status === 'recognizing text' && status(`Reading page ${i + 1} of ${js.numPages}… ${Math.round(m.progress * 100)}%`));
    layText(doc.getPage(i), font, r.lines, s, view[0], view[1], view[3] - view[1]);
    out.push(r.text); done++;
  }
  return { bytes: done ? await doc.save() : bytes, text: out.join('\n\n'), done, pages: js.numPages };
}

// Photos -> searchable PDF (one page per photo, A4-ish width, photo plus invisible text)
export async function photosToPdf(files, status) {
  const doc = await PDFDocument.create(); const font = await doc.embedFont(StandardFonts.Helvetica); const texts = [];
  for (let n = 0; n < files.length; n++) {
    status(`Preparing photo ${n + 1} of ${files.length}…`);
    const bmp = await createImageBitmap(files[n], { imageOrientation: 'from-image' }).catch(async () => { const im = new Image(); im.src = URL.createObjectURL(files[n]); await im.decode(); return im; });
    const k0 = Math.min(1, 2400 / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k0); c.height = Math.round(bmp.height * k0);
    const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(bmp, 0, 0, c.width, c.height);
    const jpg = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
    const img = await doc.embedJpg(new Uint8Array(await jpg.arrayBuffer()));
    const W = 595, k = c.width / W, H = c.height / k;
    const page = doc.addPage([W, H]); page.drawImage(img, { x: 0, y: 0, width: W, height: H });
    const r = await recognize(c, m => m.status === 'recognizing text' && status(`Reading photo ${n + 1} of ${files.length}… ${Math.round(m.progress * 100)}%`));
    layText(page, font, r.lines, k, 0, 0, H); texts.push(r.text);
  }
  return { bytes: await doc.save(), text: texts.join('\n\n') };
}
