// ---------- Import a list from Excel, CSV or the clipboard (shared by the Press apps) ----------
// importList({ title, fields: [{ key, label, required, guess: /regex/ }], onDone(rows) })
// rows = [{ key: value, ... }] – only non-empty rows.
const DataIn = (() => {
  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  let xlsxReady = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxReady) return xlsxReady;
    xlsxReady = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL; s.onload = () => res(window.XLSX); s.onerror = () => { xlsxReady = null; rej(new Error('Excel reader could not be loaded – connect to the internet once, or save the sheet as CSV.')); };
      document.head.appendChild(s);
    });
    return xlsxReady;
  }
  function parseDelimited(text) {
    text = String(text || '').replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    const first = text.split('\n').slice(0, 5).join('\n');
    const count = (ch) => (first.match(new RegExp(ch === '\t' ? '\t' : '\\' + ch, 'g')) || []).length;
    const delim = ['\t', ';', ','].map((d) => [d, count(d)]).sort((a, b) => b[1] - a[1])[0];
    const d = delim[1] ? delim[0] : '\t';
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c;
      } else if (c === '"' && cell === '') q = true;
      else if (c === d) { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += c;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.map((r) => r.map((x) => x.trim())).filter((r) => r.some((x) => x));
  }
  async function readFile(file) {
    if (/\.(xlsx|xlsm|xls|ods)$/i.test(file.name)) {
      const X = await loadXlsx();
      const wb = X.read(await file.arrayBuffer(), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }).map((r) => r.map((x) => String(x == null ? '' : x).trim())).filter((r) => r.some((x) => x));
    }
    const buf = await file.arrayBuffer();
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) text = new TextDecoder('windows-1252').decode(buf); // old Excel CSV
    return parseDelimited(text);
  }
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function css() {
    if (document.getElementById('dataInCss')) return;
    const st = document.createElement('style'); st.id = 'dataInCss';
    st.textContent = `
.di-dlg{max-width:760px;width:calc(100vw - 32px)}
.di-dlg h2{margin:0 0 4px}
.di-src{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:10px 0}
.di-dlg textarea{width:100%;min-height:90px;font:13px/1.4 Consolas,monospace;padding:8px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit;box-sizing:border-box}
.di-map{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin:10px 0}
.di-map label{display:flex;flex-direction:column;gap:3px;font-size:13px;font-weight:600}
.di-map select{padding:6px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit}
.di-map select option{color:#111}
.di-prev{max-height:220px;overflow:auto;border:1px solid #8885;border-radius:8px}
.di-prev table{border-collapse:collapse;width:100%;font-size:13px}
.di-prev th,.di-prev td{border-bottom:1px solid #8884;padding:4px 8px;text-align:left;white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis}
.di-prev th{position:sticky;top:0;background:Canvas}
.di-msg{font-size:13px;opacity:.8;margin:6px 0}
.di-err{color:#dc2626}
.di-btns{display:flex;gap:8px;justify-content:flex-end;margin-top:12px}`;
    document.head.appendChild(st);
  }
  function importList(opts) {
    css();
    let dlg = document.getElementById('dataInDlg');
    if (dlg) dlg.remove();
    dlg = document.createElement('dialog'); dlg.id = 'dataInDlg'; dlg.className = 'dlg di-dlg';
    dlg.innerHTML = `<h2>${esc(opts.title || 'Import a list')}</h2>
      <p class="di-msg">Open an Excel or CSV file, or copy the cells in Excel / Google Sheets and paste them below.</p>
      <div class="di-src"><button type="button" class="primary" data-a="file">📂 Open Excel or CSV file…</button><span class="di-msg" data-r="fname"></span></div>
      <textarea data-r="paste" placeholder="…or paste here (Ctrl+V)"></textarea>
      <div data-r="mapBox" hidden>
        <label class="check"><input type="checkbox" data-r="head" checked> The first row has the column names</label>
        <div class="di-map" data-r="map"></div>
        <div class="di-prev" data-r="prev"></div>
        <p class="di-msg" data-r="count"></p>
      </div>
      <p class="di-msg di-err" data-r="err"></p>
      <div class="di-btns"><button type="button" class="ghost" data-a="cancel">Cancel</button><button type="button" class="primary" data-a="ok" disabled>Import</button></div>
      <input type="file" accept=".xlsx,.xls,.xlsm,.ods,.csv,.txt,.tsv" hidden data-r="file">`;
    document.body.appendChild(dlg);
    const R = (k) => dlg.querySelector(`[data-r="${k}"]`), A = (k) => dlg.querySelector(`[data-a="${k}"]`);
    let table = [];
    const cols = () => { const w = Math.max(0, ...table.map((r) => r.length)); return Array.from({ length: w }, (_, i) => (R('head').checked && table[0] && table[0][i]) || `Column ${String.fromCharCode(65 + (i % 26))}`); };
    function guessMap() {
      const names = cols(), used = new Set();
      R('map').innerHTML = opts.fields.map((f) => {
        let pick = -1;
        if (R('head').checked && f.guess) pick = names.findIndex((n, i) => !used.has(i) && f.guess.test(n));
        if (pick < 0 && f.required) pick = names.findIndex((_, i) => !used.has(i));
        if (pick >= 0) used.add(pick);
        return `<label>${esc(f.label)}${f.required ? ' *' : ''}<select data-k="${f.key}"><option value="-1">${f.required ? '– choose –' : '(not used)'}</option>${names.map((n, i) => `<option value="${i}"${i === pick ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select></label>`;
      }).join('');
      R('map').querySelectorAll('select').forEach((s) => (s.onchange = preview));
      preview();
    }
    function mapped() {
      const m = {}; R('map').querySelectorAll('select').forEach((s) => (m[s.dataset.k] = Number(s.value)));
      const body = R('head').checked ? table.slice(1) : table;
      return body.map((r) => { const o = {}; opts.fields.forEach((f) => { o[f.key] = m[f.key] >= 0 ? (r[m[f.key]] || '').trim() : ''; }); return o; })
        .filter((o) => opts.fields.some((f) => f.required && o[f.key]));
    }
    function preview() {
      const rows = mapped(), shown = opts.fields.filter((f) => Number(R('map').querySelector(`[data-k="${f.key}"]`).value) >= 0);
      R('prev').innerHTML = `<table><tr>${shown.map((f) => `<th>${esc(f.label)}</th>`).join('')}</tr>${rows.slice(0, 8).map((o) => `<tr>${shown.map((f) => `<td>${esc(o[f.key])}</td>`).join('')}</tr>`).join('')}</table>`;
      R('count').textContent = rows.length ? `${rows.length} row${rows.length > 1 ? 's' : ''} will be imported${rows.length > 8 ? ' (the first 8 are shown)' : ''}.` : 'No rows yet – choose the right columns.';
      A('ok').disabled = !rows.length;
    }
    function setTable(t, label) {
      table = t; R('err').textContent = '';
      if (!table.length) { R('err').textContent = 'No rows were found.'; R('mapBox').hidden = true; A('ok').disabled = true; return; }
      R('fname').textContent = label || '';
      // a first row that looks like headings
      const h = table[0].join(' ');
      R('head').checked = opts.fields.some((f) => f.guess && f.guess.test(h)) || /name|address|street|city|zip|email|title|date|company/i.test(h);
      R('mapBox').hidden = false; guessMap();
    }
    A('file').onclick = () => R('file').click();
    R('file').onchange = async () => {
      const f = R('file').files[0]; if (!f) return;
      R('fname').textContent = 'Reading…';
      try { setTable(await readFile(f), f.name); } catch (e) { R('fname').textContent = ''; R('err').textContent = e.message || 'That file could not be read.'; }
      R('file').value = '';
    };
    R('paste').oninput = () => { if (R('paste').value.trim()) setTable(parseDelimited(R('paste').value), ''); };
    R('head').onchange = guessMap;
    A('cancel').onclick = () => dlg.close();
    A('ok').onclick = () => { const rows = mapped(); dlg.close(); opts.onDone(rows); };
    dlg.addEventListener('close', () => setTimeout(() => dlg.remove(), 50));
    dlg.showModal();
  }
  return { importList, parseDelimited, readFile };
})();
