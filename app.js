/* Đối chiếu Sticker — chạy hoàn toàn trên trình duyệt.
   Excel: SheetJS + JSZip (ảnh). PDF: pdf.js (text + render). OCR: Tesseract.js (khi chữ trên proof đã bị outline). */
'use strict';
const $ = (s, r = document) => r.querySelector(s);
pdfjsLib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';

const DEFAULT_SIZE = `English,French,Italian,Spanish,China Tops,China Bottoms
XS,XP,XS,ECH,165/84A,165/64A
S,P,S,CH,170/88A,170/68A
M,M,M,M,170/96A,170/72A
L,G,L,G,170/104A,170/80A
XL,XG,XL,EG,170/112A,170/90A
XXL,XXG,XXL,EEG,170/116A,170/96A`;
const ST = { ok: 'Đạt', err: 'Lỗi', warn: 'Cảnh báo', man: 'Kiểm tra mắt' };
const RANK = { ok: 0, man: 1, warn: 2, err: 3 };
const OCR_SCALE = 6;          // ~430 dpi
const TEXT_SCALE = 3;

const state = { sheets: [], xName: '', pdfFiles: [], cache: {}, results: [] };

/* ---------------- tiện ích ---------------- */
const nfc = s => (s == null ? '' : String(s)).normalize('NFC').trim();
const deacc = s => nfc(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D');
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const uniq = a => [...new Set(a)];
const worst = arr => arr.reduce((w, s) => (RANK[s] > RANK[w] ? s : w), 'ok');
const pill = s => `<span class="pill s-${s}">${ST[s]}</span>`;
const fz = s => String(s || '').toUpperCase().replace(/[OQD]/g, '0').replace(/[IL|]/g, '1').replace(/S/g, '5').replace(/[BE]/g, '8').replace(/Z/g, '2').replace(/G/g, '6');
const itemNum = s => (String(s).match(/LB\s*0*(\d{3,6})/gi) || []).map(x => x.replace(/LB\s*0*/i, ''));
function lev(a, b) {
  const m = a.length, n = b.length, d = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
  for (let j = 1; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++)
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[m][n];
}
function store(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
function load(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }

/* ---------------- thiết lập ---------------- */
const SET_IDS = ['sRN', 'sCA', 'sMonth', 'sType', 'sComp', 'sSize'];
function initSettings() {
  $('#sSize').value = DEFAULT_SIZE;
  SET_IDS.forEach(id => { const v = load('stk_' + id); if (v != null) $('#' + id).value = v; });
  SET_IDS.forEach(id => $('#' + id).addEventListener('change', () => { store('stk_' + id, $('#' + id).value); reevaluateAll(); }));
  $('#btnResetSet').onclick = () => {
    SET_IDS.forEach(id => { try { localStorage.removeItem('stk_' + id); } catch (e) {} });
    $('#sRN').value = '54867'; $('#sCA').value = '67359'; $('#sMonth').value = ''; $('#sType').value = 'auto'; $('#sComp').value = ''; $('#sSize').value = DEFAULT_SIZE;
    reevaluateAll();
  };
}
function getSettings() {
  const rows = $('#sSize').value.split(/\r?\n/).map(l => l.split(',').map(s => s.trim())).filter(r => r.length >= 4 && r[0]);
  const size = {};
  rows.slice(1).forEach(r => { size[r[0].toUpperCase()] = { en: r[0], fr: r[1], it: r[2], es: r[3], cnTop: r[4] || '', cnBot: r[5] || '' }; });
  const tokens = new Set();
  Object.values(size).forEach(s => [s.en, s.fr, s.it, s.es].forEach(t => t && tokens.add(t.toUpperCase())));
  return { rn: $('#sRN').value.trim(), ca: $('#sCA').value.trim(), month: $('#sMonth').value.trim(), type: $('#sType').value, comp: $('#sComp').value.trim(), size, tokens };
}

/* ---------------- đọc Excel tác nghiệp ---------------- */
async function loadExcel(file) {
  $('#xInfo').textContent = 'Đang đọc ' + file.name + '…';
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  let images = {};
  try { images = await extractImages(buf); } catch (e) { console.warn('Không đọc được ảnh', e); }
  state.sheets = wb.SheetNames.map(n => parseSheet(n, wb.Sheets[n], images[n] || []));
  state.xName = file.name;
  const nOrd = state.sheets.reduce((a, s) => a + s.orders.length, 0);
  $('#xInfo').innerHTML = `<b>${esc(file.name)}</b><br>${state.sheets.length} sheet · ${nOrd} dòng đơn hàng`;
  renderSheets();
  updateButtons();
  reevaluateAll();
}

function parseSheet(name, ws, imgs) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true }).map(r => r.map(c => (c == null ? '' : nfc(c))));
  const sh = { name, aoa, meta: {}, title: '', labelSize: null, labelSizeText: '', items: [], orders: [], colorGroups: [], styles: new Set(), images: imgs, notes: [] };
  const metaKeys = { 'mua': 'season', 'style': 'style', 'item': 'item', 'thi truong': 'market' };
  aoa.forEach((row, ri) => {
    row.forEach((c, ci) => {
      if (!c) return;
      if (!sh.title) sh.title = c;
      const k = deacc(c).toLowerCase().replace(/:$/, '').trim();
      if (metaKeys[k] && !sh.meta[metaKeys[k]]) {
        const v = row.slice(ci + 1).find(x => x);
        if (v) sh.meta[metaKeys[k]] = v;
      }
      const m = deacc(c).match(/kich thuoc nhan[^:]*:\s*([\d.,]+)\s*[*x×]\s*([\d.,]+)\s*mm/i);
      if (m && !sh.labelSize) { sh.labelSize = [parseFloat(m[1].replace(',', '.')), parseFloat(m[2].replace(',', '.'))]; sh.labelSizeText = c; }
    });
  });
  sh.items = uniq([...itemNum(sh.meta.item || ''), ...itemNum(name)]);

  // Bảng đơn hàng (header có VPONO) và bảng style-color (header có TECHPACKNO + Color code)
  for (let ri = 0; ri < aoa.length; ri++) {
    const row = aoa[ri];
    const up = row.map(c => c.toUpperCase().replace(/\s+/g, ''));
    if (up.includes('VPONO')) {
      const col = k => up.findIndex(c => k.includes(c));
      const ix = {
        ship: col(['SHIPTONAME']), label: col(['LABEL']), vpo: col(['VPONO']), tp: col(['TECHPACKNO']),
        style: col(['CUSTSTYLE', 'MASTSTYLE']), color: col(['CUSTCOLORCODE', 'MASTCOLORCODE', 'COLORCODE']),
        size: col(['CUSTSIZEDESC', 'MASTSIZEDESC', 'SIZE']), fac: col(['FACTORYID']), care: col(['CARECONTENT']), upc: col(['CUSTOMERARTICLE', 'UPC'])
      };
      for (let rj = ri + 1; rj < aoa.length; rj++) {
        const r = aoa[rj];
        const u2 = r.map(c => c.toUpperCase().replace(/\s+/g, ''));
        if (u2.includes('VPONO')) break;
        const vpo = (r[ix.vpo] || '').replace(/\s/g, '');
        if (!/^\d{6,}$/.test(vpo)) continue;
        const g = i => (i >= 0 ? r[i] || '' : '');
        sh.orders.push({ row: rj + 1, ship: g(ix.ship), market: g(ix.label), vpo, tp: g(ix.tp), style: g(ix.style).toUpperCase(), color: g(ix.color).toUpperCase(), size: g(ix.size).toUpperCase(), fac: g(ix.fac), care: g(ix.care), upc: g(ix.upc) });
      }
      ri += 1;
    } else if (up.includes('TECHPACKNO') && up.some(c => c === 'COLORCODE')) {
      const cc = up.map((c, i) => (/COLORCODE$/.test(c) ? i : -1)).filter(i => i >= 0);
      const sc = ['TECHPACKNO', 'MASTSTYLE', 'CUSTSTYLE'].map(k => up.indexOf(k)).filter(i => i >= 0);
      for (let rj = ri + 1; rj < aoa.length; rj++) {
        const r = aoa[rj];
        if (!r.some(x => x)) break;
        sc.forEach(i => r[i] && sh.styles.add(String(r[i]).toUpperCase()));
        const g = cc.map(i => (r[i] || '').toUpperCase()).filter(Boolean);
        if (g.length) {
          const set = new Set();
          g.forEach(c => { set.add(c); if (/^Z/.test(c)) set.add(c.slice(1)); });
          sh.colorGroups.push(set);
        }
      }
    }
  }
  sh.orders.forEach(o => sh.styles.add(o.style));
  sh.pos = uniq(sh.orders.map(o => o.vpo));
  // gom theo style-color
  sh.sc = {};
  sh.orders.forEach(o => {
    const k = o.style + '|' + o.color;
    const e = sh.sc[k] || (sh.sc[k] = { style: o.style, color: o.color, sizes: new Set(), pos: new Set(), fac: new Set(), markets: new Set(), care: new Set() });
    e.sizes.add(o.size); e.pos.add(o.vpo); o.fac && e.fac.add(o.fac); o.market && e.markets.add(o.market); o.care && e.care.add(o.care);
  });
  return sh;
}

async function extractImages(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = async p => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
  const rels = async p => { const d = await xml(p); const m = {}; if (d) [...d.getElementsByTagName('Relationship')].forEach(r => { m[r.getAttribute('Id')] = r.getAttribute('Target'); }); return m; };
  const resolve = (base, t) => { if (t.startsWith('/')) return t.slice(1); const parts = base.split('/'); parts.pop(); t.split('/').forEach(s => { if (s === '..') parts.pop(); else if (s !== '.') parts.push(s); }); return parts.join('/'); };
  const wbx = await xml('xl/workbook.xml'); const wbr = await rels('xl/_rels/workbook.xml.rels');
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const out = {};
  for (const s of [...wbx.getElementsByTagNameNS('*', 'sheet')]) {
    const name = s.getAttribute('name');
    const rid = s.getAttributeNS(RNS, 'id') || s.getAttribute('r:id');
    const sp = resolve('xl/workbook.xml', wbr[rid] || '');
    const sr = await rels(sp.replace(/([^/]+)$/, '_rels/$1.rels'));
    const list = [];
    for (const t of Object.values(sr)) {
      if (!/drawings\/drawing/.test(t)) continue;
      const dp = resolve(sp, t);
      const dx = await xml(dp); if (!dx) continue;
      const dr = await rels(dp.replace(/([^/]+)$/, '_rels/$1.rels'));
      const anchors = [...dx.getElementsByTagNameNS('*', 'twoCellAnchor'), ...dx.getElementsByTagNameNS('*', 'oneCellAnchor')];
      for (const a of anchors) {
        const blip = a.getElementsByTagNameNS('*', 'blip')[0]; if (!blip) continue;
        const emb = blip.getAttributeNS(RNS, 'embed') || blip.getAttribute('r:embed');
        const mp = resolve(dp, dr[emb] || ''); const f = zip.file(mp); if (!f) continue;
        const from = a.getElementsByTagNameNS('*', 'from')[0];
        const row = from ? +from.getElementsByTagNameNS('*', 'row')[0].textContent : 0;
        const ext = mp.split('.').pop().toLowerCase();
        const data = await f.async('uint8array');
        if (data.length < 400) continue; // bỏ ảnh rỗng / icon
        const url = URL.createObjectURL(new Blob([data], { type: ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : 'image/jpeg' }));
        list.push({ row, url });
      }
    }
    list.sort((a, b) => a.row - b.row);
    out[name] = list;
  }
  return out;
}

function imgCaption(sh, row) {
  for (let r = row; r >= Math.max(0, row - 4); r--) {
    const t = (sh.aoa[r] || []).find(c => c && c.length < 80 && !/^\d+$/.test(c));
    if (t) return t;
  }
  return '';
}

function renderSheets() {
  const html = state.sheets.map(sh => {
    const scs = Object.values(sh.sc).map(e => `${e.style}-${e.color}`);
    return `<div class="card sheet">
      <h4>${esc(sh.name)}</h4>
      <div class="kv">
        <b>Mùa</b><span>${esc(sh.meta.season || '—')}</span>
        <b>Style</b><span>${esc(sh.meta.style || '—')}</span>
        <b>Item</b><span>${esc(sh.meta.item || '—')}</span>
        <b>Thị trường</b><span>${esc(sh.meta.market || '—')}</span>
        <b>Kích thước</b><span>${sh.labelSize ? sh.labelSize.join(' × ') + ' mm' : '—'}</span>
        <b>Đơn hàng</b><span>${sh.orders.length} dòng · ${sh.pos.length} PO</span>
      </div>
      ${scs.length ? `<div class="chips">${scs.slice(0, 24).map(s => `<span class="chip">${esc(s)}</span>`).join('')}${scs.length > 24 ? `<span class="chip">+${scs.length - 24}</span>` : ''}</div>` : ''}
      ${sh.images.length ? `<details class="sub"><summary>${sh.images.length} hình trong sheet</summary><div class="gallery" style="margin-top:8px">${sh.images.map(im => `<figure><img src="${im.url}" data-zoom alt=""><figcaption>Dòng ${im.row + 1} · ${esc(imgCaption(sh, im.row))}</figcaption></figure>`).join('')}</div></details>` : ''}
    </div>`;
  }).join('');
  $('#xSummary').innerHTML = `<h3>Tác nghiệp: ${esc(state.xName)}</h3><div class="sheets">${html}</div>`;
}

/* ---------------- đọc proof PDF ---------------- */
let worker = null, ocrLogger = null;
async function getWorker() {
  if (!worker) {
    const base = new URL('lib/tesseract/', location.href).href;
    worker = await Tesseract.createWorker('eng', 1, { workerPath: base + 'worker.min.js', corePath: base, langPath: base, gzip: true, logger: m => ocrLogger && ocrLogger(m) });
    await worker.setParameters({ tessedit_pageseg_mode: '11' });
  }
  return worker;
}

function parseHeader(text) {
  const t = text.replace(/\s+/g, ' ');
  const h = {};
  h.ticket = (t.match(/\b([A-Z]{2}\d{6,}[A-Z]?)\b/) || [])[1] || '';
  const it = t.match(/\bLB\s*0*(\d{3,6})(?:\s*(C\s*\/\s*\d+))?/i);
  h.item = it ? it[1] : ''; h.itemText = it ? it[0] : '';
  const sz = t.match(/(\d+(?:[.,]\d+)?)\s*mm\s*[x×*]\s*(\d+(?:[.,]\d+)?)\s*mm/i);
  h.size = sz ? [parseFloat(sz[1].replace(',', '.')), parseFloat(sz[2].replace(',', '.'))] : null;
  h.supplierPO = (t.match(/PO\s*#\s*([A-Z]{2,4}\d{5,})/i) || [])[1] || '';
  let seg = '';
  const i = t.search(/VS\s*PO\s*#/i);
  if (i >= 0) {
    seg = t.slice(i + 6);
    const end = seg.search(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/);
    if (end >= 0) seg = seg.slice(0, end);
    seg = seg.replace(/\s+/g, '');
    h.pos = uniq(seg.match(/\d{10}/g) || []);
  } else h.pos = uniq(t.match(/\b\d{10}\b/g) || []);
  h.date = (t.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s*\d{4}[^P]*/) || [''])[0].trim();
  return h;
}

function mergeLines(lines) {
  lines.sort((a, b) => a.y0 - b.y0 || a.x0 - b.x0);
  const out = [];
  lines.forEach(l => {
    const p = out[out.length - 1];
    const h = Math.min(l.y1 - l.y0, p ? p.y1 - p.y0 : 1e9);
    if (p && Math.abs((l.y0 + l.y1) / 2 - (p.y0 + p.y1) / 2) < h * 0.45) {
      p.text = (p.x0 <= l.x0 ? p.text + ' ' + l.text : l.text + ' ' + p.text).trim();
      p.x0 = Math.min(p.x0, l.x0); p.x1 = Math.max(p.x1, l.x1); p.y0 = Math.min(p.y0, l.y0); p.y1 = Math.max(p.y1, l.y1);
    } else out.push({ ...l });
  });
  return out;
}

function clusterColumns(lines, W) {
  const L = lines.filter(l => l.text.trim() && (l.x1 - l.x0) < W * 0.45).map(l => ({ ...l, cx: (l.x0 + l.x1) / 2 }));
  L.sort((a, b) => a.cx - b.cx);
  const cols = []; let last = -1e9;
  L.forEach(l => { if (l.cx - last > W * 0.03) cols.push([]); cols[cols.length - 1].push(l); last = l.cx; });
  return cols.map(c => {
    const ml = mergeLines(c);
    return { lines: ml, text: ml.map(l => l.text).join('\n'), x0: Math.min(...ml.map(l => l.x0)), x1: Math.max(...ml.map(l => l.x1)), y0: Math.min(...ml.map(l => l.y0)), y1: Math.max(...ml.map(l => l.y1)) };
  }).filter(c => c.text.replace(/[^A-Za-z0-9%]/g, '').length >= 12);
}

const keyScore = t => (t.match(/RN\s*[#H*]?\s*\d{4,}|CA\s*[#H*]?\s*\d{4,}|ID\s*[#H*]?\s*\d{6,}|\d{5,8}\s*-\s*[A-Z0-9]{3}|made in/gi) || []).length;

function rotateCanvas(src, deg) {
  if (!deg) return src;
  const c = document.createElement('canvas');
  c.width = src.height; c.height = src.width;
  const g = c.getContext('2d');
  g.translate(c.width / 2, c.height / 2); g.rotate(deg * Math.PI / 180); g.drawImage(src, -src.width / 2, -src.height / 2);
  return c;
}

// Chỉ giữ nét tối (chữ đen): bỏ watermark xanh nhạt "VISUAL ONLY", khung đỏ, nền màu
function binarize(src) {
  const c = document.createElement('canvas'); c.width = src.width; c.height = src.height;
  const g = c.getContext('2d'); g.drawImage(src, 0, 0);
  const d = g.getImageData(0, 0, c.width, c.height), a = d.data;
  for (let i = 0; i < a.length; i += 4) { const v = Math.max(a[i], a[i + 1], a[i + 2]) < 150 ? 0 : 255; a[i] = a[i + 1] = a[i + 2] = v; a[i + 3] = 255; }
  g.putImageData(d, 0, 0);
  return c;
}
function cropCanvas(src, x, y, w, h) {
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(h));
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  g.drawImage(src, -x, -y);
  return c;
}
async function ocrLines(canvas, psm = '11') {
  const w = await getWorker();
  await w.setParameters({ tessedit_pageseg_mode: psm });
  const { data } = await w.recognize(canvas);
  let lines = data.lines;
  if (!lines || !lines.length) {
    lines = [];
    (data.blocks || []).forEach(b => (b.paragraphs || []).forEach(p => (p.lines || []).forEach(l => lines.push(l))));
  }
  const L = lines.map(l => ({ text: (l.text || '').replace(/\n/g, ' ').trim(), x0: l.bbox.x0, y0: l.bbox.y0, x1: l.bbox.x1, y1: l.bbox.y1, conf: l.confidence }));
  L.words = (data.words || []).filter(w => w.text && w.text.trim()).map(w => ({ text: w.text.trim(), x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1 }));
  return L;
}
// Gom từ thành cột nhãn theo khoảng trống ngang giữa các nhãn
function wordColumns(words, W, H) {
  const ws = words.filter(w => (w.x1 - w.x0) < W * 0.3).sort((a, b) => a.x0 - b.x0);
  if (!ws.length) return [];
  const hs = ws.map(w => w.y1 - w.y0).sort((a, b) => a - b);
  const mh = hs[Math.floor(hs.length / 2)] || 20;
  // mật độ từ theo trục X — khe giữa các nhãn có mật độ thấp (chỉ footer/header chạy ngang)
  const B = 4, nb = Math.ceil(W / B) + 1, cover = new Array(nb).fill(0);
  ws.forEach(w => { for (let b = Math.floor(w.x0 / B); b <= Math.floor(w.x1 / B); b++) cover[b]++; });
  const T = 3, segs = []; let s = -1;
  for (let b = 0; b < nb; b++) {
    if (cover[b] >= T && s < 0) s = b;
    if ((cover[b] < T || b === nb - 1) && s >= 0) { segs.push([s * B, b * B]); s = -1; }
  }
  const merged = [];
  segs.forEach(g => { const p = merged[merged.length - 1]; if (p && g[0] - p[1] < mh * 0.6) p[1] = g[1]; else merged.push([...g]); });
  const groups = merged.map(([x0, x1]) => ({ x0, x1, words: [] }));
  ws.forEach(w => { const cx = (w.x0 + w.x1) / 2; const g = groups.find(g => cx >= g.x0 && cx <= g.x1); if (g) g.words.push(w); });
  // bỏ cụm chữ cách xa theo chiều dọc (footer, header)
  groups.forEach(g => {
    const sorted = g.words.sort((a, b) => a.y0 - b.y0); const parts = [];
    sorted.forEach(w => { const p = parts[parts.length - 1]; if (p && w.y0 - p.y1 < mh * 25) { p.ws.push(w); p.y1 = Math.max(p.y1, w.y1); } else parts.push({ ws: [w], y1: w.y1 }); });
    g.words = parts.sort((a, b) => b.ws.length - a.ws.length)[0]?.ws || [];
    if (g.words.length) { g.x0 = Math.min(...g.words.map(w => w.x0)); g.x1 = Math.max(...g.words.map(w => w.x1)); }
  });
  return groups.filter(g => g.words.length).map(g => {
    const ml = mergeLines(g.words.map(w => ({ ...w })));
    return { lines: ml, text: ml.map(l => l.text).join('\n'), x0: g.x0, x1: g.x1, y0: Math.min(...ml.map(l => l.y0)), y1: Math.max(...ml.map(l => l.y1)) };
  }).filter(c => c.text.replace(/[^A-Za-z0-9%]/g, '').length >= 12);
}

async function processPdf(file, prog) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const out = { name: file.name, pages: [], headerText: '' };
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp1 = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent();
    const items = tc.items.filter(i => i.str && i.str.trim()).map(i => {
      const fh = Math.hypot(i.transform[2], i.transform[3]) || i.height || 8;
      return { str: i.str, x: i.transform[4], y: vp1.height - i.transform[5], w: i.width, h: fh };
    });
    const fullText = tc.items.map(i => i.str + (i.hasEOL ? '\n' : ' ')).join('');
    out.headerText += fullText + '\n';
    const textMode = keyScore(fullText) >= 3;
    const S = textMode ? TEXT_SCALE : OCR_SCALE;
    const vp = page.getViewport({ scale: S });
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    prog(`Trang ${p}/${pdf.numPages}: đang dựng hình…`);
    await page.render({ canvasContext: ctx, viewport: vp }).promise;

    let lines = [], rot = 0, img = canvas, mode = textMode ? 'text' : 'ocr';
    if (textMode) {
      lines = items.map(i => ({ text: i.str, x0: i.x * S, x1: (i.x + i.w) * S, y0: (i.y - i.h) * S, y1: i.y * S }));
    } else {
      // cắt bỏ phần chữ thật (header PO#/LB/mm, footer) — phần nhãn đã bị outline nên không có text
      // (watermark kiểu "VISUAL ONLY" cũng là text thật nên chỉ dựa vào dòng header PO#/LB/mm)
      const hm = items.filter(i => i.y < vp1.height * 0.5 && /PO\s*#|\bLB\s*\d|\d\s*mm\b/i.test(i.str));
      let top = 0;
      if (hm.length) {
        const base = Math.max(...hm.map(i => i.y)), lh = Math.max(...hm.map(i => i.h));
        const cont = items.filter(i => i.y > base && i.y <= base + lh * 2.6 && /\d/.test(i.str));
        top = Math.max(base, ...cont.map(i => i.y)) + lh * 0.6;
      }
      const bottom = vp1.height;
      const crop = cropCanvas(canvas, 0, Math.floor(top * S), canvas.width, Math.max(50, Math.floor((bottom - top) * S)));
      let best = null;
      for (const deg of [0, 90, 270]) {
        const c = rotateCanvas(crop, deg), bin = binarize(c);
        prog(`Trang ${p}/${pdf.numPages}: OCR bố cục${deg ? ' (xoay ' + deg + '°)' : ''}…`);
        const ls = await ocrLines(bin, '11');
        const sc = keyScore(ls.map(l => l.text).join('\n'));
        if (!best || sc > best.sc) best = { sc, ls, c, bin, deg };
        if (sc >= 4) break;
      }
      lines = best.ls; img = best.c; rot = best.deg;
      // lượt 2: đọc lại từng cột nhãn ở chế độ khối chữ (chính xác hơn cho mã/size/tháng)
      const cols1 = lines.words && lines.words.length ? wordColumns(lines.words, img.width, img.height) : clusterColumns(lines, img.width);
      const pad = img.width * 0.006;
      for (let k = 0; k < cols1.length; k++) {
        const c = cols1[k];
        prog(`Trang ${p}/${pdf.numPages}: đọc nhãn ${k + 1}/${cols1.length}…`);
        const x = Math.max(0, c.x0 - pad), y = Math.max(0, c.y0 - img.height * 0.03);
        const h = Math.min(img.height - y, c.y1 - c.y0 + img.height * 0.06);
        const ls = await ocrLines(cropCanvas(best.bin, x, y, Math.min(img.width - x, c.x1 - c.x0 + 2 * pad), h), '6');
        if (ls.length) c.text2 = ls.map(l => l.text).filter(Boolean).join('\n');
      }
      cols1.forEach(c => { if (c.text2 && keyScore(c.text2) >= keyScore(c.text) - 1) { c.text1 = c.text; c.text = c.text2; } });
      lines.cols = cols1;
    }
    const cols = lines.cols || clusterColumns(lines, img.width);
    // ảnh hiển thị thu nhỏ + ảnh cắt từng cột
    const dispScale = Math.min(1, 1400 / img.width);
    cols.forEach(c => {
      const pad = 12 * (img.width / 2500);
      const x = Math.max(0, c.x0 - pad), y = Math.max(0, c.y0 - pad);
      const w = Math.min(img.width - x, c.x1 - c.x0 + 2 * pad), h = Math.min(img.height - y, c.y1 - c.y0 + 2 * pad);
      const k = Math.min(1, 600 / Math.max(w, h));
      const cc = document.createElement('canvas'); cc.width = Math.max(1, w * k); cc.height = Math.max(1, h * k);
      cc.getContext('2d').drawImage(img, x, y, w, h, 0, 0, cc.width, cc.height);
      c.crop = cc.toDataURL('image/jpeg', 0.85);
    });
    const disp = document.createElement('canvas');
    disp.width = img.width * dispScale; disp.height = img.height * dispScale;
    disp.getContext('2d').drawImage(img, 0, 0, disp.width, disp.height);
    out.pages.push({ n: p, mode, rot, cols, dispUrl: disp.toDataURL('image/jpeg', 0.85), dispScale, dispW: disp.width, dispH: disp.height, rawText: lines.map(l => l.text).join('\n') });
  }
  out.header = parseHeader(out.headerText);
  return out;
}

/* ---------------- phân tích 1 nhãn ---------------- */
function parseLabel(col, set) {
  const t = col.text, U = deacc(t).toUpperCase();
  const lines = t.split('\n').map(s => s.trim()).filter(Boolean);
  const r = { text: t, crop: col.crop, box: [col.x0, col.y0, col.x1, col.y1] };
  const sc = U.match(/\b(\d{5,8})\s*[-–]\s*([A-Z0-9]{2,6})\b/);
  if (sc) { r.style = sc[1]; r.colorRaw = sc[2]; }
  r.rn = (U.match(/RN\s*[#H*]?\s*(\d{5})/) || [])[1] || '';
  r.ca = (U.match(/CA\s*[#H*]?\s*(\d{5})/) || [])[1] || '';
  r.id = (U.match(/(?:ID|[1L|\[(]D)\s*[#H*]?\s*(\d{6,10})/) || [])[1] || '';
  const idL = lines.findIndex(l => /(?:ID|[1L|\[(]D)\s*[#H*]?\s*\d/i.test(l));
  if (idL >= 0) for (const l of lines.slice(idL + 1, idL + 3)) {
    const m = l.match(/^\s*(\d{2})\s*[\/ .-]?\s*(\d{2})\s*$/); if (m) { r.month = m[1] + ' ' + m[2]; break; }
  }
  const madeKeys = ['MADE IN', 'FABRIQU', 'HECHO EN', 'BBRICATO'];
  r.madeIn = madeKeys.filter(k => U.includes(k)).length;
  const cn = U.match(/(\d{3})\s*\/\s*(\d{2,3})\s*([A-Z])\b/);
  if (cn) r.china = `${cn[1]}/${cn[2]}${cn[3]}`;
  for (let i = lines.length - 1; i >= 0; i--) {
    const toks = deacc(lines[i]).toUpperCase().split('/').map(s => s.replace(/[^A-Z0-9]/g, '')).filter(Boolean);
    if (toks.length && toks.every(x => set.tokens.has(x)) && set.size[toks[0]]) { r.sizeTokens = toks; r.size = toks[0]; break; }
  }
  if (col.text1) { // bổ sung trường còn thiếu từ lượt OCR thứ nhất
    const r1 = parseLabel({ ...col, text: col.text1, text1: null }, set);
    ['style', 'colorRaw', 'rn', 'ca', 'id', 'month', 'china', 'sizeTokens', 'size'].forEach(k => { if (!r[k] && r1[k]) r[k] = r1[k]; });
    if (r1.style) r.alt = { style: r1.style, colorRaw: r1.colorRaw };
    r.madeIn = Math.max(r.madeIn, r1.madeIn);
  }
  r.pcts = (t.match(/(\d{1,3})\s*%/g) || []).map(x => parseInt(x, 10));
  r.kind = (r.rn || r.ca || r.id || r.style || r.size) ? 'size' : (r.pcts.length ? 'comp' : 'other');
  return r;
}

/* ---------------- đối chiếu ---------------- */
function colorEq(sh, a, b) {
  if (a === b || 'Z' + a === b || a === 'Z' + b) return true;
  return sh.colorGroups.some(g => g.has(a) && g.has(b));
}
function matchSC(sh, style, color) {
  const es = Object.values(sh.sc);
  let e = es.find(x => x.style === style && colorEq(sh, color, x.color));
  if (e) return { e, fuzzy: false };
  const hits = es.filter(x => fz(x.style) === fz(style) && (fz(x.color) === fz(color) || sh.colorGroups.some(g => g.has(x.color) && [...g].some(c => fz(c) === fz(color)))));
  if (hits.length === 1) return { e: hits[0], fuzzy: true };
  if (hits.length > 1) return { e: null, ambiguous: hits };
  const key = `${style}-${color}`;
  const near = es.map(x => ({ x, d: lev(key, `${x.style}-${x.color}`) })).sort((a, b) => a.d - b.d)[0];
  return { e: null, near: near && near.d <= 3 ? near.x : null };
}
function expectedSizeTokens(s) { return uniq([s.en, s.fr, s.es].map(x => (x || '').toUpperCase()).filter(Boolean)); }
function parseComp(s) {
  const pcts = (s.match(/(\d{1,3})\s*%/g) || []).map(x => parseInt(x, 10));
  const fibers = uniq((deacc(s).toLowerCase().match(/[a-z]{4,}/g) || []).filter(w => FIBERS.has(w)));
  return { pcts, fibers };
}
const FIBERS = new Set(['polyester', 'polyamide', 'elastane', 'spandex', 'nylon', 'cotton', 'viscose', 'modal', 'lyocell', 'silk', 'wool', 'recycled', 'acrylic', 'linen', 'cupro', 'acetate', 'rayon', 'lace', 'body', 'lining', 'mesh', 'trim', 'decoration', 'exclusive']);

function pickSheet(header) {
  const sh = state.sheets;
  if (!sh.length) return -1;
  let c = sh.map((s, i) => ({ i, s })).filter(x => header.item && x.s.items.includes(header.item));
  if (c.length > 1 && header.size) { const c2 = c.filter(x => x.s.labelSize && x.s.labelSize[0] === header.size[0]); if (c2.length) c = c2; }
  if (c.length) return c[0].i;
  const ov = sh.map((s, i) => ({ i, n: header.pos.filter(p => s.pos.includes(p)).length })).sort((a, b) => b.n - a.n);
  return ov[0].n ? ov[0].i : 0;
}

function evaluate(pd, si) {
  const set = getSettings();
  const sh = state.sheets[si];
  const h = pd.header;
  const checks = [];
  const add = (group, item, expected, actual, status, note = '') => checks.push({ group, item, expected, actual, status, note });
  const labels = [], comps = [];
  pd.pages.forEach(pg => pg.cols.forEach(c => {
    const r = parseLabel(c, set); r.page = pg.n;
    if (r.kind === 'size') labels.push(r); else if (r.kind === 'comp') comps.push(r);
  }));

  /* --- thông tin chung --- */
  if (!sh) { add('Thông tin chung', 'Sheet tác nghiệp', '', '', 'err', 'Chưa có file tác nghiệp'); return { checks, labels, comps, si }; }
  add('Thông tin chung', 'Mã nhãn (Item)', sh.meta.item || sh.items.join(', '), h.itemText || '(không đọc được)',
    h.item ? (sh.items.includes(h.item) ? 'ok' : 'err') : 'warn', h.item && !sh.items.includes(h.item) ? 'Mã nhãn trên proof không khớp sheet đang chọn' : '');
  if (sh.labelSize || h.size) {
    let st = 'man', note = '';
    if (sh.labelSize && h.size) {
      const wOk = sh.labelSize[0] === h.size[0], lOk = sh.labelSize[1] === h.size[1];
      const flex = /tuy thuoc|trung binh/i.test(deacc(sh.labelSizeText));
      st = wOk && lOk ? 'ok' : (wOk && flex ? 'warn' : 'err');
      if (st === 'warn') note = 'Chiều dài khác — tác nghiệp ghi độ dài tùy thành phần vải, xác nhận lại';
    }
    add('Thông tin chung', 'Kích thước nhãn', sh.labelSize ? sh.labelSize.join(' × ') + ' mm' : '—', h.size ? h.size.join(' × ') + ' mm' : '(không đọc được)', st, note);
  }
  // PO
  const proofSC = new Set();
  labels.forEach(l => {
    if (!l.style) return;
    let m = matchSC(sh, l.style, l.colorRaw);
    if ((!m.e || m.fuzzy) && l.alt) { const m2 = matchSC(sh, l.alt.style, l.alt.colorRaw); if (m2.e && !m2.fuzzy) { m = m2; l.style = l.alt.style; l.colorRaw = l.alt.colorRaw; } }
    l.match = m; if (m.e) proofSC.add(m.e);
  });
  if (h.pos.length) {
    const extra = h.pos.filter(p => !sh.pos.includes(p));
    const missing = sh.pos.filter(p => !h.pos.includes(p));
    const relMissing = missing.filter(p => [...proofSC].some(e => e.pos.has(p)));
    add('Thông tin chung', 'PO trên proof có trong tác nghiệp', `${sh.pos.length} PO`, `${h.pos.length} PO`, extra.length ? 'err' : 'ok', extra.length ? 'PO lạ: ' + extra.join(', ') : '');
    if (relMissing.length) add('Thông tin chung', 'PO của style-color này bị thiếu trên proof', relMissing.join(', '), 'Không có', 'err', 'Các PO này có đặt style-color in trên proof');
    const other = missing.filter(p => !relMissing.includes(p));
    if (other.length) add('Thông tin chung', 'PO trong tác nghiệp không có trên proof', other.join(', '), 'Không có', 'warn', 'Có thể thuộc proof/ticket khác — xác nhận lại');
  } else add('Thông tin chung', 'Danh sách PO', sh.pos.join(', '), '(không tìm thấy trên proof)', 'man');

  /* --- từng nhãn --- */
  const pickChina = (s, l) => {
    if (set.type === 'top') return s.cnTop; if (set.type === 'bottom') return s.cnBot;
    return l.china && l.china === s.cnBot ? s.cnBot : s.cnTop;
  };
  const months = uniq(labels.map(l => l.month).filter(Boolean));
  labels.forEach(l => {
    const iss = [];
    // style-color
    if (!l.style) iss.push(['warn', 'Không đọc được style-color']);
    else if (l.match.e) { l.scText = `${l.match.e.style}-${l.match.e.color}`; if (l.match.fuzzy) iss.push(['ok', `OCR đọc "${l.style}-${l.colorRaw}", khớp ${l.scText}`]); }
    else if (l.match.ambiguous) iss.push(['warn', `"${l.style}-${l.colorRaw}" gần nhiều mã: ` + l.match.ambiguous.map(x => x.style + '-' + x.color).join(', ')]);
    else iss.push(['err', `Style-color ${l.style}-${l.colorRaw} không có trong tác nghiệp` + (l.match.near ? ` (gần giống ${l.match.near.style}-${l.match.near.color})` : '')]);
    const e = l.match && l.match.e;
    // RN / CA / ID
    if (set.rn) iss.push(!l.rn ? ['warn', 'Không đọc được RN#'] : l.rn === set.rn ? ['ok', ''] : ['err', `RN# ${l.rn} ≠ ${set.rn}`]);
    if (set.ca) iss.push(!l.ca ? ['warn', 'Không đọc được CA#'] : l.ca === set.ca ? ['ok', ''] : ['err', `CA# ${l.ca} ≠ ${set.ca}`]);
    const facs = e ? [...e.fac] : uniq(sh.orders.map(o => o.fac).filter(Boolean));
    if (!l.id) { const f = facs.find(x => l.text.includes(x)); if (f) l.id = f; }
    if (facs.length) iss.push(!l.id ? ['warn', 'Không đọc được ID#'] : facs.includes(l.id) ? ['ok', ''] : ['err', `ID# ${l.id} ≠ Factory ID ${facs.join('/')}`]);
    // tháng
    if (set.month) iss.push(!l.month ? ['warn', 'Không đọc được tháng SX'] : l.month.replace(/\D/g, '') === set.month.replace(/\D/g, '') ? ['ok', ''] : ['err', `Tháng ${l.month} ≠ ${set.month}`]);
    // size
    if (!l.size) iss.push(['warn', 'Không đọc được size']);
    else {
      const s = set.size[l.size];
      const exp = expectedSizeTokens(s);
      const got = l.sizeTokens;
      if (exp.join('/') === got.join('/')) {} else if (exp.every(x => got.includes(x)) && got.every(x => exp.includes(x))) iss.push(['warn', `Thứ tự size: ${got.join('/')} (chuẩn ${exp.join('/')})`]);
      else iss.push(['err', `Dịch size sai: ${got.join('/')} — chuẩn ${exp.join('/')}`]);
      const cn = pickChina(s, l);
      if (cn) iss.push(!l.china ? ['warn', 'Không đọc được size China'] : l.china === cn ? ['ok', ''] : ['err', `Size China ${l.china} ≠ ${cn}`]);
      if (e && !e.sizes.has(l.size)) iss.push(['warn', `Size ${l.size} không có trong đơn của ${l.scText}`]);
      l.sizeText = got.join('/') + (l.china ? ' · ' + l.china : '');
    }
    if (l.madeIn < 3) iss.push(['warn', `Chỉ đọc được ${l.madeIn}/4 dòng "made in Vietnam" — kiểm tra mắt`]);
    l.issues = iss.filter(x => x[0] !== 'ok' || x[1]);
    l.status = worst(iss.map(x => x[0]));
  });
  if (months.length > 1) add('Từng nhãn', 'Tháng SX đồng nhất giữa các nhãn', '1 giá trị', months.join(' | '), 'warn');
  else if (months.length && !set.month) add('Từng nhãn', 'Tháng sản xuất in trên nhãn', '(nhập ở Thiết lập để kiểm)', months[0], 'man');

  // độ phủ size theo style-color
  if (!labels.length) add('Từng nhãn', 'Nhãn size trên proof', '', 'Không nhận diện được nhãn nào', 'err', 'Thử file PDF rõ hơn hoặc kiểm tra mắt');
  proofSC.forEach(e => {
    const got = new Set(labels.filter(l => l.match && l.match.e === e && l.size).map(l => l.size));
    const order = ['XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL'];
    const srt = a => a.sort((x, y) => (order.indexOf(x) + 1 || 99) - (order.indexOf(y) + 1 || 99));
    const miss = srt([...e.sizes].filter(s => !got.has(s)));
    const extra = srt([...got].filter(s => !e.sizes.has(s)));
    add('Từng nhãn', `Đủ size cho ${e.style}-${e.color}`, srt([...e.sizes]).join(', '), srt([...got]).join(', '),
      miss.length ? 'err' : extra.length ? 'warn' : 'ok', (miss.length ? 'Thiếu: ' + miss.join(', ') + '. ' : '') + (extra.length ? 'Thừa (không đặt hàng): ' + extra.join(', ') : ''));
  });
  const dup = labels.filter(l => l.size && l.match && l.match.e).map(l => l.scText + ' ' + l.size).filter((k, i, a) => a.indexOf(k) !== i);
  if (dup.length) add('Từng nhãn', 'Nhãn trùng', '', uniq(dup).join(', '), 'warn');
  labels.forEach((l, i) => l.issues.forEach(([s, msg]) => s !== 'ok' && add('Từng nhãn', `Nhãn #${i + 1} ${l.size || ''}`, '', '', s, msg)));

  /* --- thành phần --- */
  const compText = comps.map(c => c.text).join('\n');
  const careSrc = set.comp || uniq([...proofSC].flatMap(e => [...e.care])).join('; ');
  if (comps.length) {
    if (careSrc) {
      const ex = parseComp(careSrc), got = parseComp(compText);
      const missF = ex.fibers.filter(f => !got.fibers.includes(f));
      const pe = [...ex.pcts].sort((a, b) => a - b).join(','), pg = [...got.pcts].sort((a, b) => a - b).join(',');
      add('Thành phần', 'Tên sợi / thành phần', ex.fibers.join(', '), got.fibers.join(', '), missF.length ? 'err' : 'ok', missF.length ? 'Thiếu: ' + missF.join(', ') : '');
      add('Thành phần', 'Tỷ lệ %', ex.pcts.join('% ') + '%', got.pcts.join('% ') + '%', pe === pg ? 'ok' : 'warn', pe === pg ? '' : 'Khác — OCR có thể đọc sai số nhỏ, kiểm tra mắt');
    } else {
      const sum = comps.flatMap(c => c.pcts).reduce((a, b) => a + b, 0);
      add('Thành phần', 'Thành phần vải', '(nhập thành phần ở Thiết lập để so tự động)', comps.flatMap(c => c.pcts).join('% ') + '%', 'man', sum % 100 ? `Tổng % đọc được = ${sum}, không chia hết cho 100 — kiểm tra mắt` : '');
    }
  } else add('Thành phần', 'Thành phần vải', careSrc || '', 'Không nhận diện được cột thành phần', 'man');
  add('Kiểm tra mắt', 'Biểu tượng giặt (care symbols)', 'Theo care code trong tác nghiệp', 'Xem ảnh nhãn bên dưới', 'man');
  add('Kiểm tra mắt', 'Chữ Trung Quốc, bố cục, vị trí gấp, màu in', 'Theo hình mẫu tác nghiệp', 'Xem ảnh nhãn bên dưới', 'man');
  return { checks, labels, comps, si, compText };
}

/* ---------------- hiển thị ---------------- */
function renderResult(idx) {
  const R = state.results[idx];
  const { pd, ev } = R;
  const cnt = { ok: 0, err: 0, warn: 0, man: 0 };
  ev.checks.forEach(c => cnt[c.status]++);
  const verdict = cnt.err ? ['err', `CÓ ${cnt.err} LỖI — chưa duyệt sản xuất, phản hồi lại nhà cung cấp nhãn`] :
    cnt.warn ? ['warn', `Không thấy lỗi chắc chắn, nhưng còn ${cnt.warn} mục cần xác nhận`] : ['ok', 'Khớp tác nghiệp — chỉ còn các mục kiểm tra mắt'];
  const h = pd.header;
  const sel = `<select class="sheetSel" data-i="${idx}">${state.sheets.map((s, i) => `<option value="${i}" ${i === ev.si ? 'selected' : ''}>${esc(s.name)}</option>`).join('')}</select>`;
  const rows = ev.checks.filter(c => !(c.group === 'Từng nhãn' && /^Nhãn #/.test(c.item)));
  const pg = pd.pages;
  const pagesHtml = pg.map(p => {
    const boxes = ev.labels.concat(ev.comps).filter(l => l.page === p.n).map(l => {
      const [x0, y0, x1, y1] = l.box.map(v => v * p.dispScale);
      const col = l.status ? { ok: '#1e7a4c', err: '#d92d20', warn: '#dc9b00', man: '#3b5aa8' }[l.status] : '#3b5aa8';
      return `<rect x="${x0 - 4}" y="${y0 - 4}" width="${x1 - x0 + 8}" height="${y1 - y0 + 8}" fill="none" stroke="${col}" stroke-width="3" rx="4"/>`;
    }).join('');
    return `<div style="position:relative"><img class="pagecv" src="${p.dispUrl}" data-zoom alt="Proof trang ${p.n}">
      <svg viewBox="0 0 ${p.dispW} ${p.dispH}" style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none">${boxes}</svg></div>
      <div class="muted small">Trang ${p.n} · ${p.mode === 'ocr' ? 'đọc bằng OCR' + (p.rot ? ' (xoay ' + p.rot + '°)' : '') : 'đọc chữ trực tiếp từ PDF'}</div>`;
  }).join('');
  const sh = state.sheets[ev.si];
  const html = `
  <div class="card result" id="res${idx}">
    <div class="rhead">
      <div>
        <h2>${esc(pd.name)}</h2>
        <div class="muted small">Ticket <b>${esc(h.ticket || '—')}</b> · ${esc(h.itemText || '—')} · ${h.size ? h.size.join('×') + 'mm' : ''} ${h.supplierPO ? '· PO NCC ' + esc(h.supplierPO) : ''} ${h.date ? '· ' + esc(h.date) : ''}</div>
        <div class="small" style="margin-top:4px">So với sheet: ${sel}</div>
      </div>
      <div class="badges">
        <span class="badge s-err">${cnt.err} lỗi</span><span class="badge s-warn">${cnt.warn} cảnh báo</span>
        <span class="badge s-ok">${cnt.ok} đạt</span><span class="badge s-man">${cnt.man} kiểm tra mắt</span>
      </div>
    </div>
    <div class="verdict s-${verdict[0]}">${verdict[1]}</div>
    <h3>Kết quả kiểm tra</h3>
    <div class="tbl-wrap"><table><thead><tr><th>Nhóm</th><th>Hạng mục</th><th>Tác nghiệp</th><th>Proof</th><th>Kết quả</th><th>Ghi chú</th></tr></thead><tbody>
      ${rows.map(c => `<tr><td>${esc(c.group)}</td><td>${esc(c.item)}</td><td class="mono">${esc(c.expected)}</td><td class="mono">${esc(c.actual)}</td><td class="st">${pill(c.status)}</td><td>${esc(c.note)}</td></tr>`).join('')}
    </tbody></table></div>
    <h3>Chi tiết từng nhãn (${ev.labels.length})</h3>
    <div class="tbl-wrap"><table><thead><tr><th>#</th><th>Ảnh</th><th>Size</th><th>Style-color</th><th>RN#</th><th>CA#</th><th>ID#</th><th>Tháng</th><th>Size in trên nhãn</th><th>Kết quả</th><th>Vấn đề</th></tr></thead><tbody>
      ${ev.labels.map((l, i) => `<tr><td>${i + 1}</td><td><img class="thumb" src="${l.crop}" data-zoom alt=""></td><td><b>${esc(l.size || '?')}</b></td>
        <td class="mono">${esc(l.scText || (l.style ? l.style + '-' + l.colorRaw : '?'))}</td><td class="mono">${esc(l.rn)}</td><td class="mono">${esc(l.ca)}</td><td class="mono">${esc(l.id)}</td><td class="mono">${esc(l.month || '')}</td>
        <td class="mono">${esc(l.sizeText || '')}</td><td class="st">${pill(l.status)}</td><td class="small">${l.issues.filter(x => x[0] !== 'ok').map(x => esc(x[1])).join('<br>') || ''}</td></tr>`).join('')}
    </tbody></table></div>
    <div class="cols" style="margin-top:14px">
      <div><h3>Proof (khung màu = kết quả từng nhãn)</h3>${pagesHtml}</div>
      <div>
        <h3>Thành phần vải đọc được</h3>
        ${ev.comps.map(c => `<img class="thumb" style="max-width:160px;max-height:260px;display:inline-block;margin:0 6px 6px 0" src="${c.crop}" data-zoom alt="">`).join('')}
        <div class="ocrtext">${esc(ev.compText || '(không có)')}</div>
        ${sh && sh.images.length ? `<h3>Hình mẫu trong tác nghiệp (${esc(sh.name)})</h3><div class="gallery">${sh.images.map(im => `<figure><img src="${im.url}" data-zoom alt=""><figcaption>Dòng ${im.row + 1} · ${esc(imgCaption(sh, im.row))}</figcaption></figure>`).join('')}</div>` : ''}
        <details class="sub"><summary>Toàn bộ chữ đọc được (OCR)</summary><div class="ocrtext">${esc(pd.pages.map(p => p.rawText).join('\n\n'))}</div></details>
      </div>
    </div>
  </div>`;
  const old = document.getElementById('res' + idx);
  if (old) old.outerHTML = html; else $('#results').insertAdjacentHTML('beforeend', html);
}

function reevaluateAll() {
  state.results.forEach((R, i) => { R.ev = evaluate(R.pd, R.ev ? R.ev.si : pickSheet(R.pd.header)); renderResult(i); });
}

function updateButtons() {
  $('#btnRun').disabled = !(state.sheets.length && state.pdfFiles.length);
  $('#btnExport').disabled = !state.results.length;
}

async function run() {
  $('#btnRun').disabled = true;
  const pr = $('#progress'); pr.classList.remove('hidden');
  const bar = $('.bar i', pr), txt = $('.ptext', pr);
  state.results = []; $('#results').innerHTML = '';
  try {
    for (let f = 0; f < state.pdfFiles.length; f++) {
      const file = state.pdfFiles[f];
      const key = file.name + '|' + file.size + '|' + file.lastModified;
      const base = f / state.pdfFiles.length, span = 1 / state.pdfFiles.length;
      let stage = '';
      ocrLogger = m => { if (m.status === 'recognizing text') { bar.style.width = ((base + span * m.progress) * 100).toFixed(1) + '%'; txt.textContent = `${file.name} — ${stage} ${(m.progress * 100).toFixed(0)}%`; } else if (/loading|initializ/.test(m.status)) txt.textContent = 'Đang tải bộ OCR…'; };
      const prog = s => { stage = s; txt.textContent = `${file.name} — ${s}`; bar.style.width = (base * 100).toFixed(1) + '%'; };
      const pd = state.cache[key] || (state.cache[key] = await processPdf(file, prog));
      const R = { pd, ev: null };
      state.results.push(R);
      R.ev = evaluate(pd, pickSheet(pd.header));
      renderResult(state.results.length - 1);
    }
    bar.style.width = '100%'; txt.textContent = 'Xong.';
  } catch (e) {
    console.error(e); txt.textContent = 'Lỗi: ' + e.message;
  }
  $('#btnRun').disabled = false; updateButtons();
}

function exportExcel() {
  const wb = XLSX.utils.book_new();
  const sum = [['File proof', 'Ticket', 'Item', 'Sheet tác nghiệp', 'Lỗi', 'Cảnh báo', 'Đạt', 'Kiểm tra mắt', 'Kết luận']];
  const chk = [['File proof', 'Nhóm', 'Hạng mục', 'Tác nghiệp', 'Proof', 'Kết quả', 'Ghi chú']];
  const lab = [['File proof', '#', 'Size', 'Style-color', 'RN#', 'CA#', 'ID#', 'Tháng', 'Size in trên nhãn', 'Kết quả', 'Vấn đề']];
  state.results.forEach(({ pd, ev }) => {
    const c = { ok: 0, err: 0, warn: 0, man: 0 }; ev.checks.forEach(x => c[x.status]++);
    sum.push([pd.name, pd.header.ticket, pd.header.itemText, state.sheets[ev.si] ? state.sheets[ev.si].name : '', c.err, c.warn, c.ok, c.man, c.err ? 'CÓ LỖI' : c.warn ? 'CẦN XÁC NHẬN' : 'KHỚP']);
    ev.checks.forEach(x => chk.push([pd.name, x.group, x.item, x.expected, x.actual, ST[x.status], x.note]));
    ev.labels.forEach((l, i) => lab.push([pd.name, i + 1, l.size || '', l.scText || (l.style ? l.style + '-' + l.colorRaw : ''), l.rn, l.ca, l.id, l.month || '', l.sizeText || '', ST[l.status], l.issues.filter(x => x[0] !== 'ok').map(x => x[1]).join('; ')]));
  });
  const add = (aoa, name, w) => { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = w.map(wch => ({ wch })); XLSX.utils.book_append_sheet(wb, ws, name); };
  add(sum, 'Tổng hợp', [28, 14, 16, 30, 6, 9, 6, 12, 14]);
  add(chk, 'Kiểm tra', [28, 14, 34, 34, 34, 12, 50]);
  add(lab, 'Chi tiết nhãn', [28, 4, 6, 16, 8, 8, 10, 8, 26, 12, 60]);
  const d = new Date(); const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `DoiChieuSticker_${stamp}.xlsx`);
}

/* ---------------- sự kiện ---------------- */
function bindDrop(labelSel, inputSel, handler) {
  const d = $(labelSel), inp = $(inputSel);
  inp.addEventListener('change', () => inp.files.length && handler([...inp.files]));
  ['dragenter', 'dragover'].forEach(ev => d.addEventListener(ev, e => { e.preventDefault(); d.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => d.addEventListener(ev, e => { e.preventDefault(); d.classList.remove('over'); }));
  d.addEventListener('drop', e => e.dataTransfer.files.length && handler([...e.dataTransfer.files]));
}
bindDrop('#dropX', '#fileX', fs => loadExcel(fs[0]).catch(e => { console.error(e); $('#xInfo').textContent = 'Không đọc được file: ' + e.message; }));
bindDrop('#dropP', '#fileP', fs => {
  state.pdfFiles = fs.filter(f => /\.pdf$/i.test(f.name));
  $('#pInfo').innerHTML = state.pdfFiles.map(f => esc(f.name)).join('<br>');
  updateButtons();
});
$('#btnRun').onclick = run;
$('#btnExport').onclick = exportExcel;
document.addEventListener('change', e => {
  if (e.target.matches('select.sheetSel')) { const i = +e.target.dataset.i; const R = state.results[i]; R.ev = evaluate(R.pd, +e.target.value); renderResult(i); }
});
document.addEventListener('click', e => {
  const z = e.target.closest('[data-zoom]');
  const lb = $('#lightbox');
  if (z) { $('img', lb).src = z.src; lb.classList.remove('hidden'); }
  else if (e.target.closest('#lightbox')) lb.classList.add('hidden');
});
initSettings();
