const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const FREE_CONTACTS = 10;
const STORE_BILLING = 'https://store.microsoft.com/billing';
const PRO_SKU = 'addresspress_pro';
const PRO_IDS = [PRO_SKU];
let STORE_URL = 'https://apps.microsoft.com/detail/9PJZFJ139JC2';
let proAvailable = false; // Pro limits apply only when the Store actually sells the add-on
let proSku = PRO_SKU;
let isPro = false;
try { isPro = localStorage.getItem('addresspress-pro') === '1'; } catch (_) {}
const unlocked = () => isPro || !proAvailable;

async function billing() {
  if (!('getDigitalGoodsService' in window)) return null;
  try { return await window.getDigitalGoodsService(STORE_BILLING); } catch (_) { return null; }
}
function setPro(v) {
  isPro = v;
  try { localStorage.setItem('addresspress-pro', v ? '1' : '0'); } catch (_) {}
  const b = $('proBtn');
  b.textContent = v ? '★ Pro' : '★ Get Pro';
  b.classList.toggle('is-pro', v);
  b.hidden = !proAvailable && !v;
  document.body.classList.toggle('unlocked', unlocked());
  if (typeof renderAll === 'function') renderAll();
}
async function checkPro() {
  const svc = await billing();
  if (!svc) return false;
  try {
    const details = await svc.getDetails(PRO_IDS).catch(() => []);
    proAvailable = details.length > 0;
    if (details[0]) { proSku = details[0].itemId; if (!PRO_IDS.includes(proSku)) PRO_IDS.push(proSku); }
    const list = await svc.listPurchases();
    const owned = list.some((p) => PRO_IDS.includes(p.itemId));
    setPro(owned);
    return owned;
  } catch (_) { return isPro; }
}
function proMessage(text) { const m = $('proMsg'); m.hidden = !text; m.textContent = text || ''; }
async function openPro(reason) {
  proMessage(reason || '');
  const svc = await billing();
  if (!svc) {
    $('buyBtn').textContent = 'Get AddressPress on Microsoft Store';
    $('restoreBtn').hidden = true;
  } else {
    $('restoreBtn').hidden = false;
    try {
      const [d] = await svc.getDetails([proSku]);
      const price = d && d.price ? new Intl.NumberFormat(undefined, { style: 'currency', currency: d.price.currency }).format(Number(d.price.value)) : '';
      $('buyBtn').textContent = price ? `Unlock Pro – ${price}` : 'Unlock Pro';
    } catch (_) { $('buyBtn').textContent = 'Unlock Pro'; }
  }
  if (!$('proDialog').open) $('proDialog').showModal();
}
async function buyPro() {
  const svc = await billing();
  if (!svc) { window.open(STORE_URL, '_blank'); return; }
  const methods = [{ supportedMethods: STORE_BILLING, data: { sku: proSku } }];
  try {
    let req;
    try { req = new PaymentRequest(methods); }
    catch (_) { req = new PaymentRequest(methods, { total: { label: 'Total', amount: { currency: 'USD', value: '0' } } }); }
    const res = await req.show();
    await res.complete('success');
  } catch (e) {
    if (e && e.name === 'AbortError') return;
    proMessage('Purchase could not be completed: ' + (e.message || e));
    return;
  }
  if (await checkPro()) { proMessage('Thank you! Pro is unlocked. ★'); setTimeout(() => $('proDialog').close(), 1500); }
}
function needsPro(reason) {
  if (unlocked()) return false;
  openPro(reason);
  return true;
}


// ---------- sizes (mm) ----------
const ENVS = [
  { id: 'n10', name: 'No. 10 – US business (4⅛ × 9½ in)', w: 241.3, h: 104.8 },
  { id: 'a7', name: 'A7 – greeting card (5¼ × 7¼ in)', w: 184.2, h: 133.4 },
  { id: 'a2', name: 'A2 – note card (4⅜ × 5¾ in)', w: 146.1, h: 111.1 },
  { id: 'n634', name: 'No. 6¾ – US personal (3⅝ × 6½ in)', w: 165.1, h: 92.1 },
  { id: 'dl', name: 'DL (110 × 220 mm)', w: 220, h: 110 },
  { id: 'c5', name: 'C5 (162 × 229 mm)', w: 229, h: 162 },
  { id: 'c6', name: 'C6 (114 × 162 mm)', w: 162, h: 114 },
  { id: 'custom', name: 'My own size… (Pro)', custom: true },
];
// sheet w/h, cols/rows, label w/h, left/top margin, pitch x/y
const LABS = [
  { id: 'us30', name: 'US Letter – 30 labels, 1 × 2⅝ in (Avery 5160 size)', pw: 215.9, ph: 279.4, cols: 3, rows: 10, w: 66.68, h: 25.4, left: 4.76, top: 12.7, px: 69.85, py: 25.4 },
  { id: 'us20', name: 'US Letter – 20 labels, 1 × 4 in (Avery 5161 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 10, w: 101.6, h: 25.4, left: 3.97, top: 12.7, px: 106.36, py: 25.4 },
  { id: 'us10', name: 'US Letter – 10 labels, 2 × 4 in (Avery 5163 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 5, w: 101.6, h: 50.8, left: 3.97, top: 12.7, px: 106.36, py: 50.8 },
  { id: 'us14', name: 'US Letter – 14 labels, 1⅓ × 4 in (Avery 5162 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 7, w: 101.6, h: 33.87, left: 4.06, top: 21.08, px: 106.43, py: 33.87 },
  { id: 'us6', name: 'US Letter – 6 labels, 3⅓ × 4 in (Avery 5164 size)', pw: 215.9, ph: 279.4, cols: 2, rows: 3, w: 101.6, h: 84.67, left: 4.06, top: 12.7, px: 106.43, py: 84.67 },
  { id: 'us80', name: 'US Letter – 80 return labels, ½ × 1¾ in (Avery 5167 size)', pw: 215.9, ph: 279.4, cols: 4, rows: 20, w: 44.45, h: 12.7, left: 7.62, top: 12.7, px: 52.07, py: 12.7 },
  { id: 'a421', name: 'A4 – 21 labels, 63.5 × 38.1 mm (L7160 size)', pw: 210, ph: 297, cols: 3, rows: 7, w: 63.5, h: 38.1, left: 7.21, top: 15.15, px: 66.04, py: 38.1 },
  { id: 'a414', name: 'A4 – 14 labels, 99.1 × 38.1 mm (L7163 size)', pw: 210, ph: 297, cols: 2, rows: 7, w: 99.1, h: 38.1, left: 4.65, top: 15.15, px: 101.6, py: 38.1 },
  { id: 'a424', name: 'A4 – 24 labels, 70 × 37 mm', pw: 210, ph: 297, cols: 3, rows: 8, w: 70, h: 37, left: 0, top: 0.5, px: 70, py: 37 },
];
const FONTS = { sans: 'Segoe UI, Arial, sans-serif', serif: 'Georgia, Times New Roman, serif', script: 'Segoe Script, Lucida Handwriting, cursive' };
const num = (v, lo, hi, d) => { const n = parseFloat(String(v).replace(',', '.')); return isFinite(n) && n >= lo && n <= hi ? n : d; };

// ---------- state ----------
let st = { mode: 'env', env: 'n10', ew: 190, eh: 130, lab: 'us30', start: 1, labWhat: 'to', font: 'sans', size: 1, showRet: true, ret: '', dx: 0, dy: 0, contacts: [] };
try { Object.assign(st, JSON.parse(localStorage.getItem('addresspress') || '{}')); } catch (_) {}
let saveTimer;
function save() { clearTimeout(saveTimer); saveTimer = setTimeout(() => { try { localStorage.setItem('addresspress', JSON.stringify(st)); } catch (_) { toast('Could not save – storage is full'); } }, 250); }
const uid = () => Math.random().toString(36).slice(2, 10);
const linesOf = (c) => [c.name].concat(String(c.addr || '').split(/\r?\n/)).map((s) => s.trim()).filter(Boolean);
const retLines = () => st.ret.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
const chosen = () => st.contacts.filter((c) => c.on !== false);
function envSize() { const e = ENVS.find((x) => x.id === st.env) || ENVS[0]; return e.custom ? { custom: true, w: num(st.ew, 90, 400, 190), h: num(st.eh, 60, 300, 130) } : e; }
const labSheet = () => LABS.find((x) => x.id === st.lab) || LABS[0];

// ---------- pages (SVG in millimetres) ----------
const widthOf = (s, fs, script) => s.length * fs * (script ? 0.62 : 0.52); // rough but safe estimate
function block(lines, x, y, maxW, maxH, fs, anchor) { // text block that shrinks to fit; y = top
  const script = st.font === 'script';
  let f = fs * st.size;
  const fit = []; // wrap long lines at a space before shrinking the letters
  for (const l of lines) { let rest = l; while (widthOf(rest, f, script) > maxW && rest.includes(' ')) { let cut = rest.length; do { cut = rest.lastIndexOf(' ', cut - 1); } while (cut > 0 && widthOf(rest.slice(0, cut), f, script) > maxW); if (cut <= 0) break; fit.push(rest.slice(0, cut)); rest = rest.slice(cut + 1); } fit.push(rest); }
  const first = lines[0]; lines = fit.length * f * 1.32 <= maxH * 1.25 ? fit : lines;
  const longest = Math.max(1, ...lines.map((l) => widthOf(l, 1, script)));
  f = Math.min(f, maxW / longest, maxH / (lines.length * 1.32));
  const lh = f * 1.32, x0 = anchor === 'middle' ? x + maxW / 2 : x;
  return { h: lines.length * lh, svg: lines.map((l, i) => `<text x="${x0.toFixed(2)}" y="${(y + f + i * lh).toFixed(2)}" font-size="${f.toFixed(2)}" text-anchor="${anchor || 'start'}"${i === 0 && l && first.startsWith(l) ? ' font-weight="600"' : ''}>${esc(l)}</text>`).join('') };
}
const wrap = (w, h, body, guides) => ({ w, h, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}mm" height="${h}mm" font-family="${FONTS[st.font]}" fill="#111"><rect width="${w}" height="${h}" fill="#fff"/>${guides || ''}<g transform="translate(${num(st.dx, -30, 30, 0)} ${num(st.dy, -30, 30, 0)})">${body}</g></svg>` });
function envPage(lines, screen) {
  const e = envSize(), ret = st.showRet ? retLines() : [];
  let s = '';
  if (ret.length) s += block(ret, 9, 8, e.w * 0.42, e.h * 0.28, 3.4).svg;
  const bw = e.w * 0.5, bh = e.h * 0.42, b = block(lines, 0, 0, bw, bh, 5);
  s += block(lines, e.w * 0.4, e.h * 0.56 - b.h / 2, bw, bh, 5).svg;
  const stamp = screen ? `<rect x="${e.w - 30}" y="8" width="22" height="26" fill="none" stroke="#bbb" stroke-dasharray="1.5 1.5" stroke-width=".4"/><text x="${e.w - 19}" y="23" font-size="3" fill="#aaa" text-anchor="middle" font-family="Segoe UI, Arial">stamp</text>` : '';
  return wrap(e.w, e.h, s, stamp);
}
function labelPages(items, screen, test) {
  const L = labSheet(), per = L.cols * L.rows, pages = [];
  const first = Math.max(0, Math.min(per - 1, Math.round(num(st.start, 1, per, 1)) - 1));
  const slots = Array(first).fill(null).concat(items);
  for (let p = 0; p < Math.max(1, Math.ceil(slots.length / per)); p++) {
    let s = '', g = '';
    for (let i = 0; i < per; i++) {
      const x = L.left + (i % L.cols) * L.px, y = L.top + Math.floor(i / L.cols) * L.py, it = slots[p * per + i];
      if (screen || test) g += `<rect x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${L.w}" height="${L.h}" rx="1.5" fill="none" stroke="${test ? '#111' : '#c7ccd1'}" stroke-width="${test ? 0.25 : 0.3}"/>`;
      if (screen && p === 0 && i < first) g += `<text x="${(x + L.w / 2).toFixed(2)}" y="${(y + L.h / 2 + 1).toFixed(2)}" font-size="3" fill="#b0b5ba" text-anchor="middle" font-family="Segoe UI, Arial">used</text>`;
      if (!it) continue;
      const pad = 3.5, b = block(it, 0, 0, L.w - 2 * pad, L.h - 4, 4);
      s += block(it, x + pad, y + (L.h - b.h) / 2 - 0.4, L.w - 2 * pad, L.h - 4, 4).svg;
    }
    pages.push(wrap(L.pw, L.ph, s, g));
  }
  return pages;
}
let pages = [], pageNo = 0;
function build(screen) {
  const list = chosen().map(linesOf);
  if (st.mode === 'env') return list.length ? list.map((l) => envPage(l, screen)) : [envPage(['Add an address', 'on the left to begin'], screen)];
  if (st.labWhat === 'ret') { const L = labSheet(), r = retLines(), n = L.cols * L.rows - Math.max(0, Math.round(num(st.start, 1, 99, 1)) - 1); return labelPages(r.length ? Array(Math.max(1, n)).fill(r) : [], screen); }
  return labelPages(list, screen);
}

// ---------- screen ----------
let toastTimer;
function toast(msg) { const t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 5000); }
function renderPage() {
  const p = pages[pageNo];
  $('page').innerHTML = p.svg.replace(/ width="[\d.]+mm" height="[\d.]+mm"/, ` style="aspect-ratio:${p.w}/${p.h};width:min(100%, calc((100vh - 170px) * ${p.w / p.h}))"`);
  const unit = st.mode === 'env' ? 'Envelope' : 'Sheet';
  $('pager').textContent = pages.length > 1 ? `${unit} ${pageNo + 1} of ${pages.length}` : 'Preview';
  $('prev').disabled = pageNo === 0; $('next').disabled = pageNo >= pages.length - 1;
}
function renderContacts() {
  const ul = $('contacts');
  ul.innerHTML = st.contacts.length ? st.contacts.map((c) => `<li><input type="checkbox" data-id="${c.id}" ${c.on !== false ? 'checked' : ''} aria-label="Print"><span class="nm" data-id="${c.id}" title="Click to change">${esc(c.name || '(no name)')} <small>${esc(String(c.addr || '').split(/\r?\n/)[0] || '')}</small></span><button data-del="${c.id}" title="Delete">✕</button></li>`).join('') : '<li class="none">No addresses yet. Click “+ Add an address”.</li>';
  $('count').textContent = st.contacts.length ? `(${chosen().length} of ${st.contacts.length} ticked)` : '';
}
function renderAll() {
  document.querySelectorAll('#modes button').forEach((b) => b.classList.toggle('on', b.dataset.m === st.mode));
  const env = st.mode === 'env';
  $('envF').hidden = !env; $('envCustomF').hidden = !env || st.env !== 'custom'; $('retChkF').hidden = !env;
  $('labF').hidden = env; $('labOpts').hidden = env; $('testBtn').hidden = env;
  ['env', 'lab', 'labWhat', 'font'].forEach((k) => { $(k).value = st[k]; }); $('size').value = String(st.size);
  $('showRet').checked = !!st.showRet;
  [['ew', 'ew'], ['eh', 'eh'], ['start', 'start'], ['dx', 'dx'], ['dy', 'dy']].forEach(([id, k]) => { if (document.activeElement !== $(id)) $(id).value = st[k]; });
  if (document.activeElement !== $('ret')) $('ret').value = st.ret;
  renderContacts();
  pages = build(true); pageNo = Math.min(pageNo, pages.length - 1);
  const n = chosen().length;
  $('hint').textContent = env ? `${n} envelope${n === 1 ? '' : 's'} will be printed – one for each ticked address.` : st.labWhat === 'ret' ? 'The sheet is filled with your return address.' : `${n} label${n === 1 ? '' : 's'} on ${pages.length} sheet${pages.length === 1 ? '' : 's'}.`;
  renderPage();
}

// ---------- actions ----------
function printPages(list) {
  const p = list[0];
  $('pageRule').textContent = `@page { size: ${p.w}mm ${p.h}mm; margin: 0; } @media print { #printArea svg { width: ${p.w}mm !important; height: ${p.h}mm !important; } }`;
  $('printArea').innerHTML = list.map((x) => x.svg).join('');
  window.print();
}
function doPrint() {
  if (st.mode === 'env' && !chosen().length) { toast('Tick at least one address first.'); return; }
  if (st.mode === 'lab' && st.labWhat === 'to' && !chosen().length) { toast('Tick at least one address first.'); return; }
  if (st.mode === 'lab' && st.labWhat === 'ret' && !retLines().length) { toast('Type your return address first.'); return; }
  if (st.mode === 'env' && envSize().custom && needsPro('Your own envelope size is part of Pro. All standard sizes are free.')) return;
  printPages(build(false));
}
function testPage() { printPages(labelPages([], false, true)); toast('Hold the test page behind a label sheet against a window to check that the boxes match.'); }
let editing = null;
function openContact(c) { editing = c ? c.id : null; $('cTitle').textContent = c ? 'Change this address' : 'Add an address'; $('cName').value = c ? c.name : ''; $('cAddr').value = c ? c.addr : ''; $('cNext').hidden = !!c; $('cDialog').showModal(); $('cName').focus(); }
function canAdd(n) {
  if (unlocked() || st.contacts.length + (n || 1) <= FREE_CONTACTS) return true;
  openPro(`The free version keeps up to ${FREE_CONTACTS} addresses. Printing is never limited.`); return false;
}
function saveContact(again) {
  const name = $('cName').value.trim(), addr = $('cAddr').value.replace(/\s+$/, '');
  if (!name && !addr.trim()) { toast('Type a name or an address.'); return; }
  if (editing) { const c = st.contacts.find((x) => x.id === editing); if (c) { c.name = name; c.addr = addr; } }
  else { if (!canAdd(1)) return; st.contacts.push({ id: uid(), name, addr, on: true }); }
  save(); renderAll();
  if (again && !editing) { $('cName').value = ''; $('cAddr').value = ''; $('cName').focus(); } else $('cDialog').close();
}
function parseCsv(text) {
  const rows = []; let row = [], cell = '', q = false;
  const delim = (text.split(/\r?\n/)[0].match(/;/g) || []).length > (text.split(/\r?\n/)[0].match(/,/g) || []).length ? ';' : ',';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
    else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cell); cell = ''; }
    else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); cell = ''; if (row.some((c) => c.trim())) rows.push(row); row = []; }
    else cell += ch;
  }
  row.push(cell); if (row.some((c) => c.trim())) rows.push(row);
  return rows;
}
function csvToContacts(text) {
  const rows = parseCsv(text.replace(/^﻿/, '')); if (!rows.length) return [];
  const head = rows[0].map((h) => h.trim().toLowerCase());
  const find = (...keys) => head.findIndex((h) => keys.some((k) => h === k || h.includes(k)));
  const iName = find('name'), hasHead = iName >= 0 || find('address', 'street', 'city') >= 0;
  if (!hasHead) return rows.map((r) => ({ id: uid(), name: (r[0] || '').trim(), addr: r.slice(1).map((c) => c.trim()).filter(Boolean).join('\n'), on: true })).filter((c) => c.name || c.addr);
  const iFirst = find('first name', 'given name', 'first'), iLast = find('last name', 'family name', 'surname', 'last');
  const iFull = head.findIndex((h) => h === 'name' || h === 'full name' || h === 'display name');
  const iA1 = find('address 1', 'address1', 'street', 'address line 1', 'address'), iA2 = find('address 2', 'address2', 'address line 2');
  const iCity = find('city', 'town'), iState = find('state', 'region', 'province'), iZip = find('zip', 'postal', 'postcode'), iCountry = find('country');
  const g = (r, i) => (i >= 0 && r[i] ? r[i].trim() : '');
  return rows.slice(1).map((r) => {
    const name = g(r, iFull) || [g(r, iFirst), g(r, iLast)].filter(Boolean).join(' ') || g(r, iName);
    const a1 = g(r, iA1), cityLine = [[g(r, iCity), g(r, iState)].filter(Boolean).join(', '), g(r, iZip)].filter(Boolean).join(' ');
    const addr = [a1, iA2 !== iA1 ? g(r, iA2) : '', cityLine, g(r, iCountry)].filter(Boolean).join('\n');
    return { id: uid(), name, addr, on: true };
  }).filter((c) => c.name || c.addr);
}
function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 60000); toast(`Saved ${name} to your Downloads folder`); }
function exportCsv() {
  if (!st.contacts.length) { toast('There are no addresses to save yet.'); return; }
  const q = (s) => '"' + String(s || '').replace(/"/g, '""') + '"';
  const rows = [['Name', 'Address']].concat(st.contacts.map((c) => [c.name, c.addr]));
  download(new Blob(['﻿' + rows.map((r) => r.map(q).join(',')).join('\r\n')], { type: 'text/csv' }), 'addresses.csv');
}
const SAMPLE = [['Mr and Mrs Robert Miller', '128 Maple Street\nSpringfield, IL 62704'], ['Grandma Rose Carter', '47 Orchard Lane\nPortland, OR 97205'], ['Dr Helen Brooks', '9 Harbour View\nApartment 3B\nBoston, MA 02110'], ['The Nguyen Family', '2210 Lakeside Drive\nAustin, TX 78701'], ['Emma and Liam Reed', '15 Church Road\nBristol BS1 4QA\nUnited Kingdom']];

// ---------- events ----------
ENVS.forEach((e) => $('env').add(new Option(e.name, e.id)));
LABS.forEach((l) => $('lab').add(new Option(l.name, l.id)));
const change = () => { save(); renderAll(); };
document.querySelectorAll('#modes button').forEach((b) => { b.onclick = () => { st.mode = b.dataset.m; pageNo = 0; change(); }; });
['env', 'lab', 'labWhat', 'font'].forEach((k) => { $(k).onchange = () => { st[k] = $(k).value; pageNo = 0; change(); }; });
$('size').onchange = () => { st.size = Number($('size').value); change(); };
$('showRet').onchange = () => { st.showRet = $('showRet').checked; change(); };
[['ew', 'ew'], ['eh', 'eh'], ['start', 'start'], ['dx', 'dx'], ['dy', 'dy'], ['ret', 'ret']].forEach(([id, k]) => { $(id).oninput = () => { st[k] = $(id).value; change(); }; });
$('contacts').addEventListener('change', (e) => { const id = e.target.dataset.id; const c = st.contacts.find((x) => x.id === id); if (c) { c.on = e.target.checked; pageNo = 0; change(); } });
$('contacts').addEventListener('click', (e) => {
  const del = e.target.dataset.del;
  if (del) { const c = st.contacts.find((x) => x.id === del); if (c && confirm(`Delete ${c.name || 'this address'}?`)) { st.contacts = st.contacts.filter((x) => x.id !== del); change(); } return; }
  if (e.target.closest('.nm')) { const c = st.contacts.find((x) => x.id === e.target.closest('.nm').dataset.id); if (c) openContact(c); }
});
$('addBtn').onclick = () => { if (canAdd(1)) openContact(null); };
$('allBtn').onclick = () => { st.contacts.forEach((c) => { c.on = true; }); change(); };
$('noneBtn').onclick = () => { st.contacts.forEach((c) => { c.on = false; }); change(); };
$('cSave').onclick = () => saveContact(false); $('cNext').onclick = () => saveContact(true); $('cCancel').onclick = () => $('cDialog').close();
$('prev').onclick = () => { if (pageNo > 0) { pageNo--; renderPage(); } };
$('next').onclick = () => { if (pageNo < pages.length - 1) { pageNo++; renderPage(); } };
$('printBtn').onclick = doPrint; $('testBtn').onclick = testPage;
$('moreBtn').onclick = (e) => { e.stopPropagation(); $('menu').hidden = !$('menu').hidden; };
document.addEventListener('click', () => { $('menu').hidden = true; });
$('importBtn').onclick = () => {
  if (needsPro('Importing a CSV file is part of Pro. You can always add addresses by hand.')) return;
  DataIn.importList({
    title: 'Import addresses from Excel or CSV',
    fields: [
      { key: 'name', label: 'Name', required: true, guess: /^(full ?)?name|recipient|tên|contact/i },
      { key: 'company', label: 'Company', guess: /company|organi[sz]ation|business|firm/i },
      { key: 'a1', label: 'Street / address line 1', guess: /address ?(line)? ?1|street|^address|địa chỉ/i },
      { key: 'a2', label: 'Address line 2', guess: /address ?(line)? ?2|apt|suite|unit|flat/i },
      { key: 'city', label: 'City', guess: /city|town|suburb|thành phố/i },
      { key: 'state', label: 'State / county', guess: /state|province|region|county/i },
      { key: 'zip', label: 'ZIP / postcode', guess: /zip|postal|post ?code/i },
      { key: 'country', label: 'Country', guess: /country|nation|quốc gia/i },
    ],
    onDone: (rows) => {
      const have = new Set(st.contacts.map((c) => (c.name + '|' + c.addr).toLowerCase()));
      let dup = 0;
      const list = [];
      rows.forEach((o) => {
        const cityLine = [[o.city, o.state].filter(Boolean).join(', '), o.zip].filter(Boolean).join(' ');
        const addr = [o.company, o.a1, o.a2, cityLine, o.country].filter(Boolean).join('\n');
        const key = (o.name + '|' + addr).toLowerCase();
        if (have.has(key)) { dup++; return; }
        have.add(key); list.push({ id: uid(), name: o.name, addr, on: true });
      });
      if (!list.length) { toast(dup ? 'All of these addresses are already in your list.' : 'No addresses were found.'); return; }
      if (!canAdd(list.length)) return;
      st.contacts = st.contacts.concat(list); change();
      toast(`${list.length} addresses added${dup ? ` – ${dup} duplicate${dup > 1 ? 's' : ''} skipped` : ''}.`);
    },
  });
};
$('exportBtn').onclick = exportCsv;
$('sampleBtn').onclick = () => { if (!canAdd(SAMPLE.length)) return; SAMPLE.forEach(([name, addr]) => st.contacts.push({ id: uid(), name, addr, on: true })); if (!st.ret) st.ret = 'John and Mary Carter\n31 Willow Way\nMadison, WI 53703'; change(); };
$('clearBtn').onclick = () => { if (st.contacts.length && confirm('Delete all addresses? This cannot be undone.')) { st.contacts = []; change(); } };
$('proBtn').onclick = () => { if (!isPro) openPro(''); };
$('buyBtn').onclick = buyPro;
$('restoreBtn').onclick = async () => { proMessage((await checkPro()) ? 'Pro is unlocked. ★' : 'No purchase was found on this Microsoft account.'); };
$('closePro').onclick = () => $('proDialog').close();
window.addEventListener('resize', () => renderPage());

setPro(isPro);
checkPro();
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
