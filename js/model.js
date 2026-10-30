// Excello veri modeli: çalışma kitabı, sayfalar, hücreler, geri al / yinele
(function (G) {
'use strict';
const X = G.XF;
const KC = X.MAXC;
const key = (r, c) => r * KC + c;
const kr = k => Math.floor(k / KC), kc = k => k % KC;

let uid = 1;
class Sheet {
  constructor(name) {
    this.name = name;
    this.cells = new Map();
    this.colW = {}; this.rowH = {};
    this.hiddenR = new Set(); this.hiddenC = new Set(); this.filterR = new Set();
    this.merges = [];      // {r1,c1,r2,c2}
    this.cf = [];          // koşullu biçim kuralları
    this.dv = [];          // veri doğrulama
    this.charts = [];      // grafikler
    this.filter = null;    // {r1,c1,r2,c2, crit:{col:[değerler]}}
    this.freezeR = 0; this.freezeC = 0;
    this.grid = true;
  }
}
function cloneSheet(s) {
  const n = new Sheet(s.name);
  for (const [k, c] of s.cells) n.cells.set(k, { v: c.v, f: c.f, s: c.s });
  n.colW = { ...s.colW }; n.rowH = { ...s.rowH };
  n.hiddenR = new Set(s.hiddenR); n.hiddenC = new Set(s.hiddenC); n.filterR = new Set(s.filterR);
  n.merges = s.merges.map(m => ({ ...m }));
  n.cf = JSON.parse(JSON.stringify(s.cf)); n.dv = JSON.parse(JSON.stringify(s.dv)); n.charts = JSON.parse(JSON.stringify(s.charts));
  n.filter = s.filter ? JSON.parse(JSON.stringify(s.filter)) : null;
  n.freezeR = s.freezeR; n.freezeC = s.freezeC; n.grid = s.grid;
  return n;
}

class Workbook {
  constructor() {
    this.sheets = [new Sheet('Sayfa1')];
    this.active = 0; this.ver = 1; this.name = 'Yeni Çalışma';
    this.undoS = []; this.redoS = []; this.tx = null; this.dirty = false;
    this.listeners = [];
  }
  get sheet() { return this.sheets[this.active]; }
  sheetIndex(name) { const u = name.toLocaleUpperCase('tr-TR'); return this.sheets.findIndex(s => s.name.toLocaleUpperCase('tr-TR') === u); }
  dims(si) {
    const s = this.sheets[si]; let rows = 0, cols = 0;
    if (s._dimVer === this.ver) return this._dimSpill(s, s._dim);
    for (const k of s.cells.keys()) { const r = kr(k) + 1, c = kc(k) + 1; if (r > rows) rows = r; if (c > cols) cols = c; }
    s._dimVer = this.ver; s._dim = { rows, cols }; return this._dimSpill(s, s._dim);
  }
  _dimSpill(s, d) {
    if (s._spillVer !== this.ver || !s._spill || !s._spill.size) return d;
    let rows = d.rows, cols = d.cols; for (const k of s._spill.keys()) { rows = Math.max(rows, kr(k) + 1); cols = Math.max(cols, kc(k) + 1); }
    return { rows, cols };
  }
  isRowHidden(si, r) { const s = this.sheets[si]; return s.hiddenR.has(r) || s.filterR.has(r); }
  isColHidden(si, c) { return this.sheets[si].hiddenC.has(c); }

  cell(si, r, c) { return this.sheets[si].cells.get(key(r, c)); }
  // Hesaplanmış değer
  value(si, r, c) {
    const cell = this.sheets[si] && this.sheets[si].cells.get(key(r, c));
    if (!cell || (cell.f == null && (cell.v == null || cell.v === '') && this.sheets[si]._spill)) { const sh = this.sheets[si]; if (sh && sh._spillVer === this.ver) { const sp = sh._spill.get(key(r, c)); if (sp) return sp.v; } return null; }
    if (cell.f == null) return cell.v === undefined ? null : cell.v;
    if (cell._ver === this.ver) return cell._v;
    if (cell._busy) return X.E('#DÖNGÜ!');
    cell._busy = true;
    let v;
    try {
      if (!cell._ast || cell._astF !== cell.f) { cell._ast = X.parse(cell.f); cell._astF = cell.f; }
      v = X.evaluate(cell._ast, { model: this, si, r, c });
      cell._sp = null;
      if (X.isRng(v)) v = this._spillOut(si, r, c, v, cell);
      if (v == null) v = 0;
    } catch (e) { v = X.isErr(e) ? e : X.E('#NAME?'); }
    cell._busy = false; cell._v = v; cell._ver = this.ver;
    return v;
  }
  // Dizi sonucu yan hücrelere "taşır" (FİLTRE, SIRALA, BENZERSİZ…)
  _spillOut(si, r, c, v, cell) {
    const H = v.h, W = v.w; if (!H || !W) return X.E('#VALUE!');
    if (H === 1 && W === 1) return v.get(0, 0);
    if (H * W > 200000) return X.E('#TAŞMA!');
    const sh = this.sheets[si]; if (sh._spillVer !== this.ver) { sh._spill = new Map(); sh._spillVer = this.ver; }
    const me = key(r, c); cell._blk = null;
    for (let i = 0; i < H; i++) for (let j = 0; j < W; j++) {
      if (!i && !j) continue; const k = key(r + i, c + j); const o = sh.cells.get(k);
      // engel: sonucun yazılacağı alan (kullanıcıya göstermek için saklanır)
      if (o && ((o.v != null && o.v !== '') || o.f != null)) { cell._blk = { r1: r, c1: c, r2: r + H - 1, c2: c + W - 1 }; return X.E('#TAŞMA!'); }
      const sp = sh._spill.get(k); if (sp && sp.src !== me) { cell._blk = { r1: r, c1: c, r2: r + H - 1, c2: c + W - 1 }; return X.E('#TAŞMA!'); }
    }
    const g = v.grid();
    for (let i = 0; i < H; i++) for (let j = 0; j < W; j++) { if (!i && !j) continue; let x = g[i][j]; sh._spill.set(key(r + i, c + j), { v: x == null ? null : x, src: me }); }
    cell._sp = { h: H, w: W };
    const f0 = g[0][0]; return f0 == null ? 0 : f0;
  }
  spillAt(si, r, c) { const sh = this.sheets[si]; return sh._spillVer === this.ver && sh._spill ? sh._spill.get(key(r, c)) || null : null; }
  // Tüm formülleri satır sırasıyla hesapla (derin özyinelemeyi önler)
  recalc() {
    this.ver++;
    const DYN = /(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|TRANSPOSE|TEXTSPLIT|RANDARRAY|TAKE|CHOOSECOLS)\(/;
    for (let si = 0; si < this.sheets.length; si++) {
      const sh = this.sheets[si]; sh._spill = new Map(); sh._spillVer = this.ver;
      const ks = [], first = []; for (const [k, c] of sh.cells) if (c.f != null) { (c._sp || DYN.test(c.f) || /^[^(]*:/.test(c.f) ? first : ks).push(k); }
      first.sort((a, b) => a - b); ks.sort((a, b) => a - b);
      for (const k of first) this.value(si, kr(k), kc(k));
      for (const k of ks) this.value(si, kr(k), kc(k));
    }
    this.sheets.forEach(s => s._cfCache = null);
  }
  numFmt(si, r, c) { const cell = this.cell(si, r, c); if (cell && cell.s && cell.s.nf) return cell.s.nf; const sp = this.spillAt(si, r, c); if (sp) { const sh = this.sheets[si], top = sh.cells.get(key(kr(sp.src), c)) || sh.cells.get(sp.src); return top && top.s && top.s.nf; } return undefined; }
  display(si, r, c) { const v = this.value(si, r, c); return X.formatValue(v, this.numFmt(si, r, c)); }
  // Düzenleme kutusunda gösterilecek metin
  editText(si, r, c) {
    const cell = this.cell(si, r, c); if (!cell) return '';
    if (cell.f != null) return '=' + X.toDisplay(cell.f);
    const v = cell.v; if (v == null) return '';
    if (typeof v === 'number') {
      const nf = cell.s && cell.s.nf;
      if (X.isDateFormat(nf)) return X.formatValue(v, v % 1 ? (Math.floor(v) ? 'dd.mm.yyyy hh:mm' : 'hh:mm') : 'dd.mm.yyyy');
      if (nf && nf.includes('%')) return X.fmtGeneral(+(v * 100).toPrecision(12)) + '%';
      return X.fmtGeneral(v);
    }
    if (typeof v === 'boolean') return v ? 'DOĞRU' : 'YANLIŞ';
    if (typeof v === 'string' && (v.startsWith('=') || X.parseNumLoose(v) != null)) return "'" + v;
    return String(v);
  }

  // ---- düzenleme (geri alınabilir) ----
  begin() { if (!this.tx) this.tx = { cells: new Map(), sheets: new Map(), list: null, depth: 0 }; this.tx.depth++; }
  end() {
    const t = this.tx; if (!t) return; if (--t.depth > 0) return; this.tx = null;
    if (t.cells.size || t.sheets.size || t.list) { this.undoS.push(t); if (this.undoS.length > 100) this.undoS.shift(); this.redoS = []; }
    this.dirty = true; this.recalc(); this.emit();
  }
  _rec(si, k) {
    const t = this.tx; if (!t || t.sheets.has(this.sheets[si])) return;
    const mk = si + ':' + k; if (t.cells.has(mk)) return;
    const c = this.sheets[si].cells.get(k);
    t.cells.set(mk, { sheet: this.sheets[si], k, before: c ? { v: c.v, f: c.f, s: c.s } : null });
  }
  snapSheet(si) { const t = this.tx; if (!t) return; const s = this.sheets[si]; if (!t.sheets.has(s)) t.sheets.set(s, cloneSheet(s)); }
  snapList() { const t = this.tx; if (t && !t.list) t.list = { sheets: this.sheets.slice(), names: this.sheets.map(s => s.name), active: this.active }; }
  setCell(si, r, c, data) {
    const k = key(r, c); this._rec(si, k);
    const cells = this.sheets[si].cells;
    if (!data || (data.v == null && data.f == null && !data.s)) { cells.delete(k); return; }
    cells.set(k, { v: data.v == null ? null : data.v, f: data.f == null ? null : data.f, s: data.s || null });
  }
  patchCell(si, r, c, patch) {
    const old = this.cell(si, r, c) || { v: null, f: null, s: null };
    this.setCell(si, r, c, Object.assign({ v: old.v, f: old.f, s: old.s }, patch));
  }
  setStyle(si, r, c, patch) {
    const old = this.cell(si, r, c); const s = Object.assign({}, old && old.s, patch);
    for (const k in s) if (s[k] == null || s[k] === false || s[k] === '') delete s[k];
    this.patchCell(si, r, c, { s: Object.keys(s).length ? s : null });
  }
  _swap(t) {
    // t içindeki durumu uygula, tersini döndür
    const inv = { cells: new Map(), sheets: new Map(), list: null, depth: 0 };
    if (t.list) { inv.list = { sheets: this.sheets.slice(), names: this.sheets.map(s => s.name), active: this.active }; this.sheets = t.list.sheets.slice(); t.list.names.forEach((n, i) => this.sheets[i].name = n); this.active = Math.min(t.list.active, this.sheets.length - 1); }
    for (const [sh, snap] of t.sheets) {
      inv.sheets.set(sh, cloneSheet(sh));
      Object.assign(sh, cloneSheet(snap), { name: sh.name });
    }
    for (const [mk, ch] of t.cells) {
      const c = ch.sheet.cells.get(ch.k);
      inv.cells.set(mk, { sheet: ch.sheet, k: ch.k, before: c ? { v: c.v, f: c.f, s: c.s } : null });
      if (ch.before) ch.sheet.cells.set(ch.k, { ...ch.before }); else ch.sheet.cells.delete(ch.k);
    }
    return inv;
  }
  undo() { const t = this.undoS.pop(); if (!t) return false; this.redoS.push(this._swap(t)); this.dirty = true; this.recalc(); this.emit(); return true; }
  redo() { const t = this.redoS.pop(); if (!t) return false; this.undoS.push(this._swap(t)); this.dirty = true; this.recalc(); this.emit(); return true; }
  on(fn) { this.listeners.push(fn); }
  emit() { this.listeners.forEach(f => f()); }

  // ---- sayfa işlemleri ----
  addSheet(name, at) {
    this.begin(); this.snapList();
    if (!name) { let i = this.sheets.length + 1; while (this.sheetIndex('Sayfa' + i) >= 0) i++; name = 'Sayfa' + i; }
    const s = new Sheet(name); if (at == null) this.sheets.push(s); else this.sheets.splice(at, 0, s);
    this.active = this.sheets.indexOf(s); this.end(); return s;
  }
  removeSheet(si) { if (this.sheets.length < 2) return; this.begin(); this.snapList(); this.sheets.splice(si, 1); this.active = Math.max(0, Math.min(this.active, this.sheets.length - 1)); this.end(); }
  renameSheet(si, name) {
    name = name.trim(); if (!name || /[\\/?*[\]:]/.test(name)) return false;
    const ex = this.sheetIndex(name); if (ex >= 0 && ex !== si) return false;
    const old = this.sheets[si].name; this.begin(); this.snapList();
    this.forEachFormula((ss, r, c, cell) => { const nf = X.renameSheetInFormula(cell.f, old, name); if (nf !== cell.f) this.patchCell(ss, r, c, { f: nf }); });
    this.sheets[si].name = name; this.end(); return true;
  }
  duplicateSheet(si) { this.begin(); this.snapList(); const n = cloneSheet(this.sheets[si]); let b = n.name + ' (2)', i = 2; while (this.sheetIndex(b) >= 0) b = n.name + ' (' + (++i) + ')'; n.name = b; this.sheets.splice(si + 1, 0, n); this.active = si + 1; this.end(); }
  moveSheet(si, d) { const j = si + d; if (j < 0 || j >= this.sheets.length) return; this.begin(); this.snapList(); const [s] = this.sheets.splice(si, 1); this.sheets.splice(j, 0, s); this.active = j; this.end(); }

  forEachFormula(fn) { this.sheets.forEach((s, si) => { for (const [k, c] of Array.from(s.cells)) if (c.f != null) fn(si, kr(k), kc(k), c); }); }

  // Satır/sütun ekle (n>0) / sil (n<0)
  shift(si, axis, at, n) {
    this.begin();
    const sh = this.sheets[si];
    this.sheets.forEach((_, i) => this.snapSheet(i));
    // formülleri düzelt
    this.sheets.forEach((s, ssi) => { for (const c of s.cells.values()) if (c.f != null) c.f = X.adjustFormula(c.f, s.name, sh.name, axis, at, n); });
    const moved = new Map();
    for (const [k, c] of sh.cells) {
      let r = kr(k), cc = kc(k);
      let p = axis === 'r' ? r : cc;
      if (n < 0 && p >= at && p < at - n) continue;
      if (p >= at) p += n;
      if (axis === 'r') r = p; else cc = p;
      moved.set(key(r, cc), c);
    }
    sh.cells = moved;
    const mvObj = o => { const out = {}; for (const k in o) { let p = +k; if (n < 0 && p >= at && p < at - n) continue; if (p >= at) p += n; out[p] = o[k]; } return out; };
    const mvSet = s => new Set([...s].filter(p => !(n < 0 && p >= at && p < at - n)).map(p => p >= at ? p + n : p));
    if (axis === 'r') { sh.rowH = mvObj(sh.rowH); sh.hiddenR = mvSet(sh.hiddenR); sh.filterR = mvSet(sh.filterR); }
    else { sh.colW = mvObj(sh.colW); sh.hiddenC = mvSet(sh.hiddenC); }
    const adjR = (o, a1, a2) => {
      let lo = o[a1], hi = o[a2];
      if (n > 0) { if (lo >= at) lo += n; if (hi >= at) hi += n; }
      else { const d = -n; if (lo >= at + d) lo -= d; else if (lo >= at) lo = at; if (hi >= at + d) hi -= d; else if (hi >= at) hi = at - 1; }
      if (hi < lo) return false; o[a1] = lo; o[a2] = hi; return true;
    };
    const [a1, a2] = axis === 'r' ? ['r1', 'r2'] : ['c1', 'c2'];
    sh.merges = sh.merges.filter(m => adjR(m, a1, a2) && (m.r2 > m.r1 || m.c2 > m.c1));
    sh.cf = sh.cf.filter(x => adjR(x.range, a1, a2));
    sh.dv = sh.dv.filter(x => adjR(x.range, a1, a2));
    sh.charts = sh.charts.filter(x => adjR(x.range, a1, a2));
    if (sh.filter && !adjR(sh.filter, a1, a2)) sh.filter = null;
    this.end();
  }

  // Alan değerlerini 2B dizi olarak al
  values(si, r1, c1, r2, c2) { const g = []; for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(this.value(si, r, c)); g.push(row); } return g; }
  usedRange(si) { const d = this.dims(si); return { r1: 0, c1: 0, r2: Math.max(0, d.rows - 1), c2: Math.max(0, d.cols - 1) }; }
  mergeAt(si, r, c) { for (const m of this.sheets[si].merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }

  // ---- koşullu biçim ----
  cfStyle(si, r, c) {
    const s = this.sheets[si]; if (!s.cf.length) return null;
    let out = null;
    for (const rule of s.cf) {
      const g = rule.range; if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) continue;
      const v = this.value(si, r, c);
      const st = this.cfEval(si, rule, v);
      if (st) out = Object.assign(out || {}, st);
    }
    return out;
  }
  cfStats(si, rule) {
    const s = this.sheets[si]; s._cfCache = s._cfCache || new Map();
    if (s._cfCache.has(rule)) return s._cfCache.get(rule);
    const g = rule.range; const vals = []; const counts = new Map();
    const d = this.dims(si); const r2 = Math.min(g.r2, d.rows - 1), c2 = Math.min(g.c2, d.cols - 1);
    for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { const v = this.value(si, r, c); if (typeof v === 'number') vals.push(v); if (v != null && v !== '') { const k = typeof v === 'string' ? v.toLocaleLowerCase('tr-TR') : String(v); counts.set(k, (counts.get(k) || 0) + 1); } }
    const sorted = vals.slice().sort((a, b) => a - b);
    const st = { min: sorted[0], max: sorted[sorted.length - 1], avg: vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0, sorted, counts };
    s._cfCache.set(rule, st); return st;
  }
  cfEval(si, rule, v) {
    const num = typeof v === 'number' ? v : null;
    const a = rule.a, b = rule.b;
    const sty = rule.style || {};
    const txt = v == null ? '' : String(typeof v === 'number' ? X.fmtGeneral(v) : v).toLocaleLowerCase('tr-TR');
    const at = String(a == null ? '' : a).toLocaleLowerCase('tr-TR');
    const st = () => this.cfStats(si, rule);
    switch (rule.type) {
      case 'gt': return num != null && num > a ? sty : null;
      case 'lt': return num != null && num < a ? sty : null;
      case 'ge': return num != null && num >= a ? sty : null;
      case 'le': return num != null && num <= a ? sty : null;
      case 'between': return num != null && num >= Math.min(a, b) && num <= Math.max(a, b) ? sty : null;
      case 'eq': return (num != null && typeof a === 'number' ? num === a : txt === at) ? sty : null;
      case 'ne': return (num != null && typeof a === 'number' ? num !== a : txt !== at) && v != null ? sty : null;
      case 'contains': return txt.includes(at) && v != null ? sty : null;
      case 'notcontains': return !txt.includes(at) && v != null ? sty : null;
      case 'begins': return txt.startsWith(at) && v != null ? sty : null;
      case 'blank': return v == null || v === '' ? sty : null;
      case 'notblank': return v != null && v !== '' ? sty : null;
      case 'dup': case 'unique': { if (v == null || v === '') return null; const k = typeof v === 'string' ? v.toLocaleLowerCase('tr-TR') : String(v); const n = st().counts.get(k) || 0; return (rule.type === 'dup' ? n > 1 : n === 1) ? sty : null; }
      case 'top': { if (num == null) return null; const s = st().sorted; const k = Math.max(1, Math.round(a || 10)); return num >= s[Math.max(0, s.length - k)] ? sty : null; }
      case 'bottom': { if (num == null) return null; const s = st().sorted; const k = Math.max(1, Math.round(a || 10)); return num <= s[Math.min(s.length - 1, k - 1)] ? sty : null; }
      case 'above': return num != null && num > st().avg ? sty : null;
      case 'below': return num != null && num < st().avg ? sty : null;
      case 'neg': return num != null && num < 0 ? sty : null;
      case 'scale': { if (num == null) return null; const s = st(); const t = s.max === s.min ? 0.5 : (num - s.min) / (s.max - s.min); return { bg: mix(rule.c1 || '#f8696b', rule.c3 || '#63be7b', t, rule.c2 || '#ffeb84') }; }
      case 'bar': { if (num == null) return null; const s = st(); const lo = Math.min(0, s.min), hi = Math.max(0, s.max); return { bar: hi === lo ? 0 : (num - lo) / (hi - lo), barColor: rule.color || '#4f8ef7' }; }
      case 'icons': { if (num == null) return null; const s = st(); const t = s.max === s.min ? 1 : (num - s.min) / (s.max - s.min); return { icon: t >= 0.67 ? 'up' : t >= 0.33 ? 'mid' : 'down' }; }
    }
    return null;
  }

  // ---- veri doğrulama ----
  dvAt(si, r, c) { for (const d of this.sheets[si].dv) { const g = d.range; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return d; } return null; }
  dvItems(si, d) {
    if (d.items && d.items.length) return d.items;
    if (d.src) { try { const ast = X.parse(d.src.replace(/^=/, '')); const v = X.evaluate(ast, { model: this, si, r: 0, c: 0 }); if (X.isRng(v)) return v.list().filter(x => x != null && x !== '').map(x => typeof x === 'number' ? X.fmtGeneral(x) : String(x)); } catch (e) { } }
    return [];
  }
  dvCheck(si, r, c, v) {
    const d = this.dvAt(si, r, c); if (!d || v == null || v === '') return null;
    const fail = d.msg || 'Bu hücreye girilen değer kurala uymuyor.';
    if (d.type === 'list') { const items = this.dvItems(si, d).map(x => x.toLocaleLowerCase('tr-TR')); const s = (typeof v === 'number' ? X.fmtGeneral(v) : String(v)).toLocaleLowerCase('tr-TR'); return items.includes(s) ? null : fail; }
    if (d.type === 'whole' || d.type === 'decimal' || d.type === 'date') {
      if (typeof v !== 'number') return fail; if (d.type === 'whole' && !Number.isInteger(v)) return fail;
      if (d.min != null && v < d.min) return fail; if (d.max != null && v > d.max) return fail; return null;
    }
    if (d.type === 'textlen') { const l = String(v).length; if (d.min != null && l < d.min) return fail; if (d.max != null && l > d.max) return fail; }
    return null;
  }

  // ---- kalıcılık ----
  toJSON() {
    return {
      name: this.name, active: this.active,
      sheets: this.sheets.map(s => ({
        name: s.name, cells: Array.from(s.cells, ([k, c]) => [k, c.v, c.f, c.s]),
        colW: s.colW, rowH: s.rowH, hiddenR: [...s.hiddenR], hiddenC: [...s.hiddenC], filterR: [...s.filterR],
        merges: s.merges, cf: s.cf, dv: s.dv, charts: s.charts, filter: s.filter, freezeR: s.freezeR, freezeC: s.freezeC, grid: s.grid
      }))
    };
  }
  static fromJSON(o) {
    const wb = new Workbook(); wb.name = o.name || 'Çalışma'; wb.sheets = [];
    for (const so of o.sheets) {
      const s = new Sheet(so.name);
      for (const [k, v, f, st] of so.cells) s.cells.set(k, { v: v == null ? null : v, f: f == null ? null : f, s: st || null });
      Object.assign(s, { colW: so.colW || {}, rowH: so.rowH || {}, merges: so.merges || [], cf: so.cf || [], dv: so.dv || [], charts: so.charts || [], filter: so.filter || null, freezeR: so.freezeR || 0, freezeC: so.freezeC || 0, grid: so.grid !== false });
      s.hiddenR = new Set(so.hiddenR || []); s.hiddenC = new Set(so.hiddenC || []); s.filterR = new Set(so.filterR || []);
      wb.sheets.push(s);
    }
    if (!wb.sheets.length) wb.sheets.push(new Sheet('Sayfa1'));
    wb.active = Math.min(o.active || 0, wb.sheets.length - 1);
    wb.recalc(); return wb;
  }
}

function hex(c) { c = c.replace('#', ''); if (c.length === 3) c = c.split('').map(x => x + x).join(''); return [0, 2, 4].map(i => parseInt(c.substr(i, 2), 16)); }
function mix(a, b, t, mid) {
  if (mid) { if (t < 0.5) return mix(a, mid, t * 2); return mix(mid, b, (t - 0.5) * 2); }
  const x = hex(a), y = hex(b); return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
}

G.XM = { Workbook, Sheet, cloneSheet, key, kr, kc, mix, hex };
})(typeof window !== 'undefined' ? window : globalThis);
