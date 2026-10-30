// Excello dosya işlemleri: xlsx/xlsm/xltx (ExcelJS), xls/ods/csv/numbers (SheetJS), PDF (yazdır)
(function (G) {
'use strict';
const X = G.XF, M = G.XM;
const loaded = {};
function loadScript(src) {
  if (loaded[src]) return loaded[src];
  return loaded[src] = new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error(src + ' yüklenemedi')); }; document.head.append(s); });
}
const needExcelJS = () => G.ExcelJS ? Promise.resolve() : loadScript('vendor/exceljs.min.js');
const needSheetJS = () => G.XLSX ? Promise.resolve() : loadScript('vendor/xlsx.full.min.js');

const FORMATS = {
  xlsx: { ext: 'xlsx', name: 'Excel Çalışma Kitabı (.xlsx)', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  xltx: { ext: 'xltx', name: 'Excel Şablonu (.xltx)', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template' },
  xlsm: { ext: 'xlsm', name: 'Excel Makrolu (.xlsm)', mime: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
  xls: { ext: 'xls', name: 'Excel 97-2003 (.xls)', mime: 'application/vnd.ms-excel' },
  xlt: { ext: 'xlt', name: 'Excel 97-2003 Şablonu (.xlt)', mime: 'application/vnd.ms-excel' },
  ods: { ext: 'ods', name: 'OpenDocument (.ods)', mime: 'application/vnd.oasis.opendocument.spreadsheet' },
  csv: { ext: 'csv', name: 'CSV (noktalı virgül)', mime: 'text/csv' },
};
const OPEN_ACCEPT = '.xlsx,.xlsm,.xltx,.xltm,.xls,.xlt,.xlsb,.ods,.fods,.csv,.tsv,.txt,.numbers,.excello,.json';

// ---------- Tema renkleri ----------
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
function tint(hex, t) {
  let [r, g, b] = [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16));
  const f = v => t < 0 ? v * (1 + t) : v + (255 - v) * t;
  return '#' + [r, g, b].map(v => Math.round(f(v)).toString(16).padStart(2, '0')).join('');
}
function colorOf(c) {
  if (!c) return null;
  if (c.argb) { const a = c.argb.length === 8 ? c.argb.slice(2) : c.argb; return '#' + a.toLowerCase(); }
  if (c.theme != null) { const b = THEME[c.theme]; return b ? tint(b, c.tint || 0) : null; }
  if (c.indexed != null) { if (c.indexed === 64 || c.indexed === 65) return null; const b = INDEXED[c.indexed]; return b ? '#' + b.toLowerCase() : null; }
  return null;
}
const argb = c => c ? { argb: 'FF' + c.replace('#', '').toUpperCase().padStart(6, '0').slice(-6) } : undefined;
const BSTY = { thin: 'thin', hair: 'thin', dotted: 'dashed', dashDot: 'dashed', dashDotDot: 'dashed', dashed: 'dashed', medium: 'medium', mediumDashed: 'medium', mediumDashDot: 'medium', mediumDashDotDot: 'medium', slantDashDot: 'medium', thick: 'thick', double: 'thick' };

function styleFromExcel(st) {
  if (!st) return null; const s = {};
  const f = st.font;
  if (f) { if (f.bold) s.b = 1; if (f.italic) s.i = 1; if (f.underline) s.u = 1; if (f.strike) s.st = 1; if (f.size && f.size !== 11) s.fs = f.size; const fc = colorOf(f.color); if (fc && fc !== '#000000') s.fc = fc; }
  const fl = st.fill;
  if (fl && fl.type === 'pattern' && fl.pattern && fl.pattern !== 'none') { const bg = colorOf(fl.fgColor) || colorOf(fl.bgColor); if (bg) s.bg = bg; }
  else if (fl && fl.type === 'gradient' && fl.stops && fl.stops[0]) { const bg = colorOf(fl.stops[0].color); if (bg) s.bg = bg; }
  const al = st.alignment;
  if (al) { if (al.horizontal && /left|center|right/.test(al.horizontal)) s.ha = al.horizontal.replace('centerContinuous', 'center'); if (al.vertical === 'top' || al.vertical === 'middle') s.va = al.vertical; if (al.wrapText) s.wrap = 1; }
  if (st.numFmt && !/^General$/i.test(st.numFmt)) s.nf = st.numFmt;
  const bd = st.border;
  if (bd) { const o = {}; for (const [k, n] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']]) { const e = bd[k]; if (e && e.style) o[n] = { s: BSTY[e.style] || 'thin', c: colorOf(e.color) || '#000000' }; } if (Object.keys(o).length) s.bd = o; }
  return Object.keys(s).length ? s : null;
}
function styleToExcel(s) {
  const o = {};
  const font = {}; if (s.b) font.bold = true; if (s.i) font.italic = true; if (s.u) font.underline = true; if (s.st) font.strike = true; if (s.fs) font.size = s.fs; if (s.fc) font.color = argb(s.fc);
  if (Object.keys(font).length) o.font = Object.assign({ name: 'Calibri', size: 11 }, font);
  if (s.bg) o.fill = { type: 'pattern', pattern: 'solid', fgColor: argb(s.bg), bgColor: argb(s.bg) };
  if (s.ha || s.va || s.wrap) o.alignment = { horizontal: s.ha, vertical: s.va || (s.wrap ? 'top' : undefined), wrapText: !!s.wrap };
  if (s.nf) o.numFmt = s.nf;
  if (s.bd) { o.border = {}; for (const [n, k] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (s.bd[n]) o.border[k] = { style: s.bd[n].s === 'dashed' ? 'dashed' : s.bd[n].s, color: argb(s.bd[n].c || '#000000') }; }
  return o;
}
const isDateNF = nf => X.isDateFormat(nf);

// ---------- Okuma ----------
async function open(file) {
  const name = file.name || 'Dosya';
  const ext = (name.split('.').pop() || '').toLowerCase();
  const buf = await file.arrayBuffer();
  let wb;
  if (ext === 'excello' || ext === 'json') { wb = M.Workbook.fromJSON(JSON.parse(new TextDecoder().decode(buf))); }
  else if (['xlsx', 'xlsm', 'xltx', 'xltm'].includes(ext)) {
    try { wb = await readExcelJS(buf); } catch (e) { console.warn('ExcelJS okuyamadı, SheetJS deneniyor', e); wb = await readSheetJS(buf, ext); }
  } else wb = await readSheetJS(buf, ext);
  wb.name = name.replace(/\.[^.]+$/, '');
  wb.format = ['xlsx', 'xlsm', 'xltx', 'xls', 'xlt', 'ods', 'csv'].includes(ext) ? ext : (ext === 'xltm' ? 'xlsm' : 'xlsx');
  wb.recalc(); wb.undoS = []; wb.redoS = []; wb.dirty = false;
  return wb;
}

function refToRange(ref) {
  const [a, b] = ref.replace(/\$/g, '').split(':'); const p = X.parseAddr(a), q = X.parseAddr(b || a);
  if (!p || !q) return null; return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
}
const rangeRef = g => X.addr(g.r1, g.c1) + (g.r2 !== g.r1 || g.c2 !== g.c1 ? ':' + X.addr(g.r2, g.c2) : '');

async function readExcelJS(buf) {
  await needExcelJS();
  const ewb = new G.ExcelJS.Workbook(); await ewb.xlsx.load(buf);
  const wb = new M.Workbook(); wb.sheets = [];
  let meta = null;
  ewb.eachSheet(ws => {
    if (ws.name === '_excello') { try { meta = JSON.parse(String(ws.getCell('A1').value || '') + String(ws.getCell('A2').value || '') + String(ws.getCell('A3').value || '')); } catch (e) { } return; }
    const s = new M.Sheet(ws.name); wb.sheets.push(s);
    // sütunlar
    const cc = ws.columnCount || 0;
    for (let c = 1; c <= Math.min(cc, 2000); c++) { const col = ws.getColumn(c); if (col.width) s.colW[c - 1] = Math.round(col.width * 7 + 5); if (col.hidden) s.hiddenC.add(c - 1); }
    const shared = {}, arrays = [];
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const r = rn - 1;
      if (row.height) s.rowH[r] = Math.round(row.height * 4 / 3); if (row.hidden) s.hiddenR.add(r);
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        const c = cn - 1;
        if (cell.type === G.ExcelJS.ValueType.Merge) { const st = styleFromExcel(cell.style); if (st) s.cells.set(M.key(r, c), { v: null, f: null, s: st }); return; }
        let v = cell.value, f = null;
        if (v && typeof v === 'object' && !(v instanceof Date)) {
          if (v.formula != null || v.sharedFormula != null) {
            f = v.formula;
            if (f != null) shared[cell.address] = { f, r, c };
            if (v.shareType === 'array' && v.ref) arrays.push({ r, c, ref: v.ref });
            if (f == null && v.sharedFormula) { const m = shared[v.sharedFormula]; try { f = cell.formula || (m ? X.shiftFormula(m.f, r - m.r, c - m.c) : null); } catch (e) { f = m ? X.shiftFormula(m.f, r - m.r, c - m.c) : null; } }
            v = v.result; if (v && typeof v === 'object' && !(v instanceof Date)) v = v.error ? X.E(v.error) : null;
          }
          else if (v.richText) v = v.richText.map(t => t.text).join('');
          else if (v.text != null) v = typeof v.text === 'object' && v.text.richText ? v.text.richText.map(t => t.text).join('') : v.text;
          else if (v.error) v = X.E(v.error);
          else v = null;
        }
        if (v instanceof Date) v = X.dateToSerial(v);
        if (X.isErr(v)) v = f ? null : v.code;
        const st = styleFromExcel(cell.style);
        if (v instanceof Date) v = X.dateToSerial(v);
        if (typeof v === 'number' && cell.type === G.ExcelJS.ValueType.Date && st && !st.nf) st.nf = 'dd.mm.yyyy';
        if (f) { try { X.parse(f); } catch (e) { f = null; } }
        if (v == null && f == null && !st) return;
        s.cells.set(M.key(r, c), { v: v == null ? null : v, f: f || null, s: st });
      });
    });
    for (const a of arrays) { const g = refToRange(a.ref); if (!g) continue; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== a.r || c !== a.c) { const k = M.key(r, c); const x = s.cells.get(k); if (x && x.f == null) { if (x.s) delete x.v; else s.cells.delete(k); } } }
    // birleştirmeler
    const merges = (ws.model && ws.model.merges) || [];
    for (const m of merges) { const g = refToRange(m); if (g) s.merges.push(g); }
    // dondurma
    const vw = ws.views && ws.views[0];
    if (vw && vw.state === 'frozen') { s.freezeR = vw.ySplit || 0; s.freezeC = vw.xSplit || 0; }
    if (vw && vw.showGridLines === false) s.grid = false;
    // filtre
    if (ws.autoFilter) { const ref = typeof ws.autoFilter === 'string' ? ws.autoFilter : (ws.autoFilter.from && ws.autoFilter.to ? null : null); const g = ref && refToRange(ref); if (g) s.filter = Object.assign(g, { crit: {} }); }
    // veri doğrulama
    try {
      const dvm = ws.dataValidations && ws.dataValidations.model || {};
      let n = 0;
      for (const a in dvm) {
        if (n++ > 3000) break; const d = dvm[a]; const g = refToRange(a.split(' ')[0]); if (!g || !d) continue;
        const o = { range: g, msg: d.error || '' };
        if (d.type === 'list') { o.type = 'list'; const fm = (d.formulae && d.formulae[0]) || ''; if (/^".*"$/.test(fm)) o.items = fm.slice(1, -1).split(','); else o.src = fm; }
        else if (d.type === 'whole' || d.type === 'decimal' || d.type === 'date' || d.type === 'textLength') { o.type = d.type === 'textLength' ? 'textlen' : d.type; const a1 = parseFloat(d.formulae && d.formulae[0]), a2 = parseFloat(d.formulae && d.formulae[1]); if (d.operator === 'between' || !d.operator) { o.min = isNaN(a1) ? null : a1; o.max = isNaN(a2) ? null : a2; } else if (/greaterThan/.test(d.operator)) o.min = a1; else if (/lessThan/.test(d.operator)) o.max = a1; }
        else continue;
        // aynı kuralları birleştir
        const prev = s.dv.find(x => JSON.stringify({ ...x, range: 0 }) === JSON.stringify({ ...o, range: 0 }) && x.range.c1 === g.c1 && x.range.c2 === g.c2 && x.range.r2 + 1 === g.r1);
        if (prev) prev.range.r2 = g.r2; else s.dv.push(o);
      }
    } catch (e) { }
    // koşullu biçim
    try {
      for (const cfm of ws.conditionalFormattings || []) {
        const g = refToRange(String(cfm.ref).split(' ')[0]); if (!g) continue;
        for (const ru of cfm.rules || []) {
          const st = {}; if (ru.style) { const s2 = styleFromExcel(ru.style); if (s2) { if (s2.bg) st.bg = s2.bg; if (s2.fc) st.fc = s2.fc; if (s2.b) st.b = 1; } }
          const num = x => { const v = parseFloat(x); return isNaN(v) ? String(x || '').replace(/^"|"$/g, '') : v; };
          let o = null;
          if (ru.type === 'cellIs') { const map = { greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le', between: 'between', equal: 'eq', notEqual: 'ne' }; if (map[ru.operator]) o = { type: map[ru.operator], a: num(ru.formulae && ru.formulae[0]), b: num(ru.formulae && ru.formulae[1]) }; }
          else if (ru.type === 'containsText') o = { type: ru.operator === 'notContains' ? 'notcontains' : ru.operator === 'beginsWith' ? 'begins' : 'contains', a: ru.text };
          else if (ru.type === 'colorScale') { const cs = (ru.color || []).map(colorOf); o = { type: 'scale', c1: cs[0], c2: cs.length > 2 ? cs[1] : null, c3: cs[cs.length - 1] }; }
          else if (ru.type === 'dataBar') o = { type: 'bar', color: colorOf(ru.color) || '#4f8ef7' };
          else if (ru.type === 'duplicateValues') o = { type: 'dup' };
          else if (ru.type === 'uniqueValues') o = { type: 'unique' };
          else if (ru.type === 'top10') o = { type: ru.bottom ? 'bottom' : 'top', a: ru.rank || 10 };
          else if (ru.type === 'aboveAverage') o = { type: ru.aboveAverage === false ? 'below' : 'above' };
          else if (ru.type === 'iconSet') o = { type: 'icons' };
          if (o) { o.range = { ...g }; o.style = st; s.cf.push(o); }
        }
      }
    } catch (e) { }
  });
  if (meta && meta.sheets) meta.sheets.forEach(ms => { const s = wb.sheets.find(x => x.name === ms.name); if (s) { s.charts = ms.charts || []; if (ms.cf && !s.cf.length) s.cf = ms.cf; } });
  if (!wb.sheets.length) wb.sheets.push(new M.Sheet('Sayfa1'));
  return wb;
}

function decodeText(buf) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, ''); }
  catch (e) { try { return new TextDecoder('windows-1254').decode(buf); } catch (e2) { return new TextDecoder('latin1').decode(buf); } }
}
async function readSheetJS(buf, ext) {
  await needSheetJS();
  const XL = G.XLSX;
  let xwb;
  const isText = ['csv', 'tsv', 'txt'].includes(ext);
  if (isText) {
    const txt = decodeText(buf);
    const first = txt.split(/\r?\n/)[0] || '';
    const FS = ext === 'tsv' || first.split('\t').length > first.split(';').length && first.split('\t').length > first.split(',').length ? '\t' : (first.split(';').length >= first.split(',').length ? ';' : ',');
    xwb = XL.read(txt, { type: 'string', FS, raw: true, cellDates: false });
  } else xwb = XL.read(new Uint8Array(buf), { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false });
  const wb = new M.Workbook(); wb.sheets = [];
  for (const sn of xwb.SheetNames) {
    const ws = xwb.Sheets[sn]; const s = new M.Sheet(sn); wb.sheets.push(s);
    if (!ws || !ws['!ref']) continue;
    for (const a in ws) {
      if (a[0] === '!') continue;
      const cell = ws[a]; const p = XL.utils.decode_cell(a);
      let v = cell.v, f = cell.f || null, st = null;
      if (cell.t === 'e') v = f ? null : (cell.w || '#N/A');
      if (cell.t === 'd' && v instanceof Date) v = X.dateToSerial(v);
      if (isText && typeof v === 'string') { const pi = parseInput(v); v = pi.v; if (pi.nf) st = { nf: pi.nf }; }
      if (cell.z && cell.z !== 'General' && typeof v === 'number') {
        let z = cell.z;
        // ODS tarih stilleri ayraçsız gelebiliyor; okunur bir biçime çevir
        if (ext === 'ods' && X.isDateFormat(z) && !/[./\-\s]/.test(z.replace(/"[^"]*"/g, ''))) z = v % 1 ? 'dd.mm.yyyy hh:mm' : 'dd.mm.yyyy';
        st = Object.assign(st || {}, { nf: z });
      }
      if (f) { try { X.parse(f); } catch (e) { f = null; } }
      if (v == null && !f) continue;
      s.cells.set(M.key(p.r, p.c), { v: v == null ? null : v, f, s: st });
    }
    (ws['!merges'] || []).forEach(m => s.merges.push({ r1: m.s.r, c1: m.s.c, r2: m.e.r, c2: m.e.c }));
    (ws['!cols'] || []).forEach((c, i) => { if (!c) return; if (c.wpx) s.colW[i] = Math.round(c.wpx); else if (c.wch) s.colW[i] = Math.round(c.wch * 7 + 5); if (c.hidden) s.hiddenC.add(i); });
    (ws['!rows'] || []).forEach((r, i) => { if (!r) return; if (r.hpx) s.rowH[i] = Math.round(r.hpx); if (r.hidden) s.hiddenR.add(i); });
  }
  if (!wb.sheets.length) wb.sheets.push(new M.Sheet('Sayfa1'));
  return wb;
}

// Kullanıcı girdisini değere çevir (hücreye yazma ve CSV için)
function parseInput(text) {
  const t = String(text);
  if (t === '') return { v: null };
  if (t[0] === "'") return { v: t.slice(1) };
  const tt = t.trim();
  const u = tt.toLocaleUpperCase('tr-TR');
  if (u === 'DOĞRU' || u === 'TRUE') return { v: true };
  if (u === 'YANLIŞ' || u === 'FALSE') return { v: false };
  const n = X.parseNumLoose(tt);
  if (n != null) {
    let nf = null;
    if (/%$/.test(tt)) nf = /[.,]\d/.test(tt) ? '0.00%' : '0%';
    else if (/₺|TL$/i.test(tt)) nf = '#,##0.00 ₺';
    else if (/^\$/.test(tt)) nf = '$#,##0.00';
    else if (/€/.test(tt)) nf = '#,##0.00 €';
    else if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(tt)) nf = /,/.test(tt) ? '#,##0.00' : '#,##0';
    if (/^0\d+$/.test(tt)) return { v: tt }; // baştaki sıfırlı numaralar (telefon, kod) metin kalsın
    return { v: n, nf };
  }
  const d = X.parseDateLoose(tt); if (d) return { v: d.v, nf: d.time ? 'dd.mm.yyyy hh:mm' : 'dd.mm.yyyy' };
  const tm = X.parseTimeLoose(tt); if (tm != null) return { v: tm, nf: 'hh:mm' };
  return { v: t };
}

// ---------- Yazma ----------
async function writeExcelJS(wb, fmt) {
  await needExcelJS();
  const ewb = new G.ExcelJS.Workbook(); ewb.creator = 'Excello'; ewb.created = new Date();
  const meta = { v: 1, sheets: [] };
  for (const s of wb.sheets) {
    const si = wb.sheets.indexOf(s);
    const views = [{ state: s.freezeR || s.freezeC ? 'frozen' : 'normal', xSplit: s.freezeC || 0, ySplit: s.freezeR || 0, showGridLines: s.grid !== false }];
    const ws = ewb.addWorksheet(s.name, { views });
    for (const c in s.colW) ws.getColumn(+c + 1).width = Math.max(1, (s.colW[c] - 5) / 7);
    for (const c of s.hiddenC) ws.getColumn(c + 1).hidden = true;
    for (const [k, cell] of s.cells) {
      const r = M.kr(k), c = M.kc(k); const ec = ws.getCell(r + 1, c + 1);
      if (cell.f != null) {
        const res = wb.value(si, r, c); const result = X.isErr(res) ? { error: /^#(DÖNGÜ|TAŞMA|HESAP)!$/.test(res.code) ? '#REF!' : res.code } : res;
        if (cell._sp) { ec.value = { formula: cell.f, result, shareType: 'array', ref: X.addr(r, c) + ':' + X.addr(r + cell._sp.h - 1, c + cell._sp.w - 1) }; }
        else ec.value = { formula: cell.f, result };
      }
      else if (cell.v != null) ec.value = cell.v;
      if (cell.s) ec.style = styleToExcel(cell.s);
    }
    // taşan dizi değerleri
    if (s._spill && s._spillVer === wb.ver) for (const [k, sp] of s._spill) { const v = sp.v; if (v == null || X.isErr(v)) continue; const ec = ws.getCell(M.kr(k) + 1, M.kc(k) + 1); ec.value = v; const nf = wb.numFmt(si, M.kr(k), M.kc(k)); if (nf && !s.cells.has(k)) ec.numFmt = nf; }
    for (const r in s.rowH) ws.getRow(+r + 1).height = s.rowH[r] * 0.75;
    for (const r of s.hiddenR) ws.getRow(r + 1).hidden = true;
    for (const r of s.filterR) ws.getRow(r + 1).hidden = true;
    for (const m of s.merges) { try { ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1); } catch (e) { } }
    if (s.filter) ws.autoFilter = rangeRef({ r1: s.filter.r1, c1: s.filter.c1, r2: s.filter.r1, c2: s.filter.c2 });
    // doğrulama
    for (const d of s.dv) {
      const g = d.range; let n = 0; const o = { allowBlank: true, showErrorMessage: true, error: d.msg || 'Geçersiz değer', errorTitle: 'Excello' };
      if (d.type === 'list') { o.type = 'list'; o.formulae = [d.items && d.items.length ? '"' + d.items.join(',') + '"' : (d.src || '').replace(/^=/, '')]; }
      else { o.type = d.type === 'textlen' ? 'textLength' : d.type; if (d.min != null && d.max != null) { o.operator = 'between'; o.formulae = [d.min, d.max]; } else if (d.min != null) { o.operator = 'greaterThanOrEqual'; o.formulae = [d.min]; } else if (d.max != null) { o.operator = 'lessThanOrEqual'; o.formulae = [d.max]; } else continue; }
      const r2 = Math.min(g.r2, g.r1 + 2000);
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= g.c2; c++) { if (n++ > 5000) break; ws.getCell(r + 1, c + 1).dataValidation = o; }
    }
    // koşullu biçim
    for (const ru of s.cf) {
      const ref = rangeRef(ru.range); const st = {}; if (ru.style) { if (ru.style.bg) st.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(ru.style.bg), fgColor: argb(ru.style.bg) }; if (ru.style.fc || ru.style.b) st.font = { color: argb(ru.style.fc), bold: !!ru.style.b }; }
      const q = v => typeof v === 'number' ? String(v) : '"' + String(v || '').replace(/"/g, '""') + '"';
      let rule = null; const map = { gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual', between: 'between', eq: 'equal', ne: 'notEqual' };
      const tl = X.addr(ru.range.r1, ru.range.c1);
      if (map[ru.type]) rule = { type: 'cellIs', operator: map[ru.type], formulae: ru.type === 'between' ? [q(ru.a), q(ru.b)] : [q(ru.a)], style: st };
      else if (ru.type === 'contains') rule = { type: 'containsText', operator: 'containsText', text: String(ru.a), style: st };
      else if (ru.type === 'notcontains') rule = { type: 'containsText', operator: 'notContains', text: String(ru.a), style: st };
      else if (ru.type === 'begins') rule = { type: 'containsText', operator: 'beginsWith', text: String(ru.a), style: st };
      else if (ru.type === 'scale') rule = { type: 'colorScale', cfvo: ru.c2 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }], color: (ru.c2 ? [ru.c1 || '#f8696b', ru.c2, ru.c3 || '#63be7b'] : [ru.c1 || '#f8696b', ru.c3 || '#63be7b']).map(argb) };
      else if (ru.type === 'bar') rule = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: argb(ru.color || '#4f8ef7') };
      else if (ru.type === 'top' || ru.type === 'bottom') rule = { type: 'top10', rank: ru.a || 10, bottom: ru.type === 'bottom', style: st };
      else if (ru.type === 'above' || ru.type === 'below') rule = { type: 'aboveAverage', aboveAverage: ru.type === 'above', style: st };
      else if (ru.type === 'dup') rule = { type: 'expression', formulae: [`COUNTIF(${rangeRef({ ...ru.range, r1: ru.range.r1, c1: ru.range.c1 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')},${tl})>1`], style: st };
      else if (ru.type === 'unique') rule = { type: 'expression', formulae: [`COUNTIF(${rangeRef(ru.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')},${tl})=1`], style: st };
      else if (ru.type === 'blank') rule = { type: 'expression', formulae: [`LEN(TRIM(${tl}))=0`], style: st };
      else if (ru.type === 'notblank') rule = { type: 'expression', formulae: [`LEN(TRIM(${tl}))>0`], style: st };
      else if (ru.type === 'neg') rule = { type: 'cellIs', operator: 'lessThan', formulae: ['0'], style: st };
      else if (ru.type === 'icons') rule = { type: 'iconSet', iconSet: '3Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] };
      if (rule) { rule.priority = 1; try { ws.addConditionalFormatting({ ref, rules: [rule] }); } catch (e) { } }
    }
    // grafikleri resim olarak göm
    for (const ch of s.charts) {
      try {
        const data = G.XC.extract(wb, si, ch); const png = G.XC.toPNG(ch, data, ch.w || 480, ch.h || 300);
        const id = ewb.addImage({ base64: png, extension: 'png' });
        const pos = chartAnchor(s, ch);
        ws.addImage(id, { tl: { col: pos.c, row: pos.r }, ext: { width: ch.w || 480, height: ch.h || 300 } });
      } catch (e) { console.warn(e); }
    }
    meta.sheets.push({ name: s.name, charts: s.charts });
  }
  if (meta.sheets.some(m => m.charts && m.charts.length)) {
    const ms = ewb.addWorksheet('_excello', { state: 'veryHidden' });
    const j = JSON.stringify(meta); ms.getCell('A1').value = j.slice(0, 32000); ms.getCell('A2').value = j.slice(32000, 64000) || null; ms.getCell('A3').value = j.slice(64000, 96000) || null;
  }
  ewb.views = [{ activeTab: wb.active }];
  let out = new Uint8Array(await ewb.xlsx.writeBuffer());
  if (fmt === 'xltx') out = await patchContentType(out, 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml');
  return out;
}
function chartAnchor(s, ch) {
  // grafik konumunu (px) hücre konumuna çevir
  let x = 0, c = 0; while (c < 500) { const w = s.hiddenC.has(c) ? 0 : (s.colW[c] || 88); if (x + w > ch.x) break; x += w; c++; }
  let y = 0, r = 0; while (r < 100000) { const hh = s.hiddenR.has(r) ? 0 : (s.rowH[r] || 26); if (y + hh > ch.y) break; y += hh; r++; }
  return { c: c + (ch.x - x) / (s.colW[c] || 88), r: r + (ch.y - y) / (s.rowH[r] || 26) };
}
async function patchContentType(u8, ct) {
  await needSheetJS();
  const CFB = G.XLSX.CFB;
  const z = CFB.read(u8, { type: 'array' });
  const ent = CFB.find(z, '[Content_Types].xml');
  let xml = new TextDecoder().decode(ent.content instanceof Uint8Array ? ent.content : new Uint8Array(ent.content));
  xml = xml.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', ct);
  ent.content = new TextEncoder().encode(xml); ent.size = ent.content.length;
  const outp = CFB.write(z, { type: 'array', fileType: 'zip', compression: true });
  return outp instanceof Uint8Array ? outp : new Uint8Array(outp);
}

async function writeSheetJS(wb, fmt) {
  await needSheetJS();
  const XL = G.XLSX; const xwb = XL.utils.book_new();
  wb.sheets.forEach((s, si) => {
    const ws = {}; let mr = 0, mc = 0;
    for (const [k, cell] of s.cells) {
      const r = M.kr(k), c = M.kc(k); const v = wb.value(si, r, c);
      if (v == null && cell.f == null) continue;
      const o = {};
      if (X.isErr(v)) { o.t = 'e'; o.v = 0x0F; o.w = v.code; }
      else if (typeof v === 'number') { o.t = 'n'; o.v = v; }
      else if (typeof v === 'boolean') { o.t = 'b'; o.v = v; }
      else { o.t = 's'; o.v = String(v == null ? '' : v); }
      if (cell.f != null && !cell._sp) o.f = cell.f;
      if (cell.s && cell.s.nf) o.z = cell.s.nf;
      ws[XL.utils.encode_cell({ r, c })] = o; mr = Math.max(mr, r); mc = Math.max(mc, c);
    }
    if (s._spill && s._spillVer === wb.ver) for (const [k, sp] of s._spill) { const r = M.kr(k), c = M.kc(k), v = sp.v; if (v == null || X.isErr(v)) continue; ws[XL.utils.encode_cell({ r, c })] = typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : { t: 's', v: String(v) }; mr = Math.max(mr, r); mc = Math.max(mc, c); }
    ws['!ref'] = XL.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: mr, c: mc } });
    if (s.merges.length) ws['!merges'] = s.merges.map(m => ({ s: { r: m.r1, c: m.c1 }, e: { r: m.r2, c: m.c2 } }));
    const cols = []; for (const c in s.colW) cols[+c] = { wpx: s.colW[c] }; for (const c of s.hiddenC) cols[c] = Object.assign(cols[c] || {}, { hidden: true }); if (cols.length) ws['!cols'] = cols;
    XL.utils.book_append_sheet(xwb, ws, s.name.slice(0, 31));
  });
  const out = XL.write(xwb, { bookType: fmt === 'xls' || fmt === 'xlt' ? 'biff8' : fmt, type: 'array' });
  return new Uint8Array(out);
}
function writeCSV(wb) {
  const si = wb.active; const d = wb.dims(si); const lines = [];
  for (let r = 0; r < d.rows; r++) {
    const row = [];
    for (let c = 0; c < d.cols; c++) {
      const v = wb.value(si, r, c); const nf = wb.numFmt(si, r, c);
      let t = v == null ? '' : typeof v === 'number' ? (X.isDateFormat(nf) ? X.formatValue(v, nf) : X.fmtGeneral(v)) : X.formatValue(v);
      if (/[;"\n\r]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
      row.push(t);
    }
    lines.push(row.join(';'));
  }
  return new TextEncoder().encode('﻿' + lines.join('\r\n'));
}

async function save(wb, fmt) {
  fmt = fmt || 'xlsx';
  let data;
  if (fmt === 'xlsx' || fmt === 'xltx' || fmt === 'xlsm') data = await writeExcelJS(wb, fmt === 'xlsm' ? 'xlsx' : fmt);
  else if (fmt === 'csv') data = writeCSV(wb);
  else if (fmt === 'excello') data = new TextEncoder().encode(JSON.stringify(wb.toJSON()));
  else data = await writeSheetJS(wb, fmt);
  const f = FORMATS[fmt] || { ext: fmt, mime: 'application/octet-stream' };
  return { blob: new Blob([data], { type: f.mime }), name: (wb.name || 'Excello') + '.' + f.ext, mime: f.mime, ext: f.ext };
}

// ---------- Yazdır / PDF ----------
function printHTML(wb, onlySel) {
  const si = wb.active, s = wb.sheet; const g = onlySel || wb.usedRange(si);
  let html = '<table>';
  const skip = new Set(); const mm = new Map();
  for (const m of s.merges) { mm.set(m.r1 * 1e5 + m.c1, m); for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) skip.add(r * 1e5 + c); }
  html += '<colgroup>'; for (let c = g.c1; c <= g.c2; c++) if (!s.hiddenC.has(c)) html += `<col style="width:${s.colW[c] || 88}px">`; html += '</colgroup>';
  for (let r = g.r1; r <= g.r2; r++) {
    if (wb.isRowHidden(si, r)) continue;
    html += `<tr style="height:${s.rowH[r] || 22}px">`;
    for (let c = g.c1; c <= g.c2; c++) {
      if (s.hiddenC.has(c) || skip.has(r * 1e5 + c)) continue;
      const cell = wb.cell(si, r, c); const st = Object.assign({}, cell && cell.s, wb.cfStyle(si, r, c)); const v = wb.value(si, r, c);
      let css = '';
      if (st.b) css += 'font-weight:700;'; if (st.i) css += 'font-style:italic;'; if (st.u) css += 'text-decoration:underline;'; if (st.fc) css += `color:${st.fc};`; if (st.bg) css += `background:${st.bg};`;
      if (st.fs) css += `font-size:${st.fs}pt;`; css += `text-align:${st.ha || (typeof v === 'number' ? 'right' : 'left')};`; if (st.wrap) css += 'white-space:pre-wrap;';
      if (st.bd) for (const [n, k] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (st.bd[n]) css += `border-${k}:${st.bd[n].s === 'thick' ? 2.5 : st.bd[n].s === 'medium' ? 1.8 : 1}px ${st.bd[n].s === 'dashed' ? 'dashed' : 'solid'} ${st.bd[n].c || '#000'};`;
      const m = mm.get(r * 1e5 + c); const span = m ? ` colspan="${m.c2 - m.c1 + 1}" rowspan="${m.r2 - m.r1 + 1}"` : '';
      html += `<td${span} style="${css}">${G.UI.esc(wb.display(si, r, c))}</td>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  const charts = s.charts.map(ch => `<div class="ch"><img src="${G.XC.toPNG(ch, G.XC.extract(wb, si, ch), ch.w || 480, ch.h || 300)}" style="width:${ch.w || 480}px"></div>`).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><title>${G.UI.esc(wb.name)}</title><style>
  body{font:11pt -apple-system,"Segoe UI",system-ui,sans-serif;margin:16px;color:#111} table{border-collapse:collapse;table-layout:fixed}
  td{border:1px solid ${s.grid !== false ? '#d4d4d8' : 'transparent'};padding:2px 5px;overflow:hidden;white-space:nowrap;vertical-align:bottom}
  .ch{margin-top:16px;page-break-inside:avoid} @page{margin:12mm}</style></head><body>${html}${charts}</body></html>`;
}
function print(wb, sel) {
  const html = printHTML(wb, sel);
  const fr = document.createElement('iframe'); fr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0'; document.body.append(fr);
  fr.contentDocument.open(); fr.contentDocument.write(html); fr.contentDocument.close();
  setTimeout(() => { try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch (e) { const w = open(); } setTimeout(() => fr.remove(), 60000); }, 300);
}

G.XIO = { open, save, print, parseInput, FORMATS, OPEN_ACCEPT, needExcelJS, needSheetJS, rangeRef, refToRange };
})(window);
