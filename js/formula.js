// Excello hesap motoru: ayrıştırıcı, değerlendirici, biçimlendirici ve fonksiyon kataloğu
(function (G) {
'use strict';

// ---------- Hatalar ----------
class FErr { constructor(code) { this.code = code; } toString() { return this.code; } }
const ERRS = {};
['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#N/A', '#NUM!', '#NULL!', '#DÖNGÜ!', '#TAŞMA!', '#HESAP!'].forEach(c => ERRS[c] = new FErr(c));
const E = c => ERRS[c] || new FErr(c);
const isErr = v => v instanceof FErr;
const DIV0 = E('#DIV/0!'), VAL = E('#VALUE!'), REF = E('#REF!'), NAME = E('#NAME?'), NA = E('#N/A'), NUM = E('#NUM!');

const MAXR = 1048576, MAXC = 16384;

// ---------- Hücre adresleri ----------
function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
function colIndex(s) { let n = 0; s = s.toUpperCase(); for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; }
function addr(r, c) { return colName(c) + (r + 1); }
function parseAddr(a) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim()); return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null; }
function quoteSheet(n) { return /^[A-Za-z_\u00C0-\u024F][A-Za-z0-9_.\u00C0-\u024F]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }

// ---------- Tarih ----------
const DAYMS = 86400000;
function serialToDate(s) { return new Date(Math.round((s - 25569) * DAYMS)); } // UTC alanları kullanılır
function dateToSerial(d) { return d.getTime() / DAYMS + 25569; }
function ymdToSerial(y, m, d) { return Date.UTC(y, m - 1, d) / DAYMS + 25569; }
function nowSerial() { const d = new Date(); return (d.getTime() - d.getTimezoneOffset() * 60000) / DAYMS + 25569; }
const AYLAR = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];
const GUNLER = ['Pazar', 'Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi'];

// ---------- Fonksiyon kataloğu ----------
const FN = {};      // İngilizce ad -> tanım
const TR2EN = {};   // Türkçe ad -> İngilizce ad
const trUp = s => s.toLocaleUpperCase('tr-TR');
function def(en, tr, cat, desc, args, fn, opt) {
  const d = Object.assign({ en, tr, cat, desc, args: args || [], fn }, opt || {});
  FN[en] = d; TR2EN[trUp(tr)] = en; return d;
}
function lookupFn(name) {
  const u = name.toUpperCase();
  if (FN[u]) return u;
  const t = trUp(name);
  if (TR2EN[t]) return TR2EN[t];
  if (TR2EN[u]) return TR2EN[u];
  if (FN[t]) return t;
  return null;
}

// ---------- Belirteçlere ayırma ----------
const RE_ERR = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!|DÖNGÜ!)/i;
const RE_QSHEET = /^'((?:[^']|'')+)'!/;
const RE_SHEET = /^([^\s!'"(),;:+\-*\/^&=<>%{}#$]+)!/;
const RE_CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(.])/;
const RE_COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.])/;
const RE_ROWS = /^(\$?)(\d+):(\$?)(\d+)(?![A-Za-z0-9_(.])/;
const RE_ID = /^[A-Za-zÇĞİÖŞÜçğıöşü_\\][A-Za-z0-9ÇĞİÖŞÜçğıöşü_.]*/;

function hasSemi(src) { let q = false; for (const ch of src) { if (ch === '"') q = !q; else if (!q && ch === ';') return true; } return false; }

function tokenize(src) {
  const toks = []; let i = 0; const n = src.length; const semi = hasSemi(src); const stack = [];
  while (i < n) {
    const ch = src[i]; const rest = src.slice(i); let m;
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) { if (j >= n) throw new Error('Kapanmamış tırnak'); if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; }
      toks.push({ t: 'str', v: s }); i = j + 1; continue;
    }
    if ((m = RE_ERR.exec(rest))) { toks.push({ t: 'err', v: m[0].toUpperCase() }); i += m[0].length; continue; }
    let sheet = null, sl = 0;
    if ((m = RE_QSHEET.exec(rest))) { sheet = m[1].replace(/''/g, "'"); sl = m[0].length; }
    else if ((m = RE_SHEET.exec(rest))) { sheet = m[1]; sl = m[0].length; }
    const r2 = rest.slice(sl);
    if ((m = RE_COLS.exec(r2))) { toks.push({ t: 'cols', sheet, ac1: !!m[1], c1: colIndex(m[2]), ac2: !!m[3], c2: colIndex(m[4]) }); i += sl + m[0].length; continue; }
    if ((m = RE_CELL.exec(r2))) {
      const c = colIndex(m[2]), r = +m[4] - 1;
      if (c < MAXC && r >= 0 && r < MAXR) { toks.push({ t: 'ref', sheet, ac: !!m[1], c, ar: !!m[3], r }); i += sl + m[0].length; continue; }
    }
    if ((m = RE_ROWS.exec(r2))) { toks.push({ t: 'rows', sheet, ar1: !!m[1], r1: +m[2] - 1, ar2: !!m[3], r2: +m[4] - 1 }); i += sl + m[0].length; continue; }
    if (sheet) throw new Error('Geçersiz başvuru');
    if (/[0-9]/.test(ch) || ((ch === '.' || ch === ',') && /[0-9]/.test(src[i + 1] || ''))) {
      m = semi ? /^(\d+(,\d*)?|,\d+)(e[+-]?\d+)?/i.exec(rest) : /^(\d+(\.\d*)?|\.\d+)(e[+-]?\d+)?/i.exec(rest);
      if (m) {
        let txt = m[0].replace(',', '.'); let len = m[0].length;
        if (!semi && !m[2] && !m[3] && !stack.some(x => x) && src[i + len] === ',' && /\d/.test(src[i + len + 1] || '')) {
          const m2 = /^,(\d+)/.exec(src.slice(i + len)); txt += '.' + m2[1]; len += m2[0].length;
        }
        toks.push({ t: 'num', v: txt }); i += len; continue;
      }
    }
    if ((m = RE_ID.exec(rest))) {
      const name = m[0]; i += name.length;
      let k = i; while (k < n && src[k] === ' ') k++;
      if (src[k] === '(') { const en = lookupFn(name); toks.push({ t: 'func', v: en || name.toUpperCase(), known: !!en }); continue; }
      const u = trUp(name);
      if (u === 'TRUE' || u === 'DOĞRU' || name.toUpperCase() === 'TRUE') { toks.push({ t: 'bool', v: true }); continue; }
      if (u === 'FALSE' || u === 'YANLIŞ' || name.toUpperCase() === 'FALSE') { toks.push({ t: 'bool', v: false }); continue; }
      toks.push({ t: 'name', v: name }); continue;
    }
    const two = src.substr(i, 2);
    if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%'.includes(ch)) { toks.push({ t: 'op', v: ch }); i++; continue; }
    if (ch === '(') { const last = toks[toks.length - 1]; stack.push(!!(last && last.t === 'func')); toks.push({ t: '(' }); i++; continue; }
    if (ch === ')') { stack.pop(); toks.push({ t: ')' }); i++; continue; }
    if (ch === ';' || ch === ',') { toks.push({ t: 'sep' }); i++; continue; }
    if (ch === ':') { toks.push({ t: ':' }); i++; continue; }
    throw new Error('Beklenmeyen karakter: ' + ch);
  }
  return toks;
}

function refText(t) {
  const sh = t.sheet != null ? quoteSheet(t.sheet) + '!' : '';
  if (t.t === 'ref') return sh + (t.ac ? '$' : '') + colName(t.c) + (t.ar ? '$' : '') + (t.r + 1);
  if (t.t === 'cols') return sh + (t.ac1 ? '$' : '') + colName(t.c1) + ':' + (t.ac2 ? '$' : '') + colName(t.c2);
  if (t.t === 'rows') return sh + (t.ar1 ? '$' : '') + (t.r1 + 1) + ':' + (t.ar2 ? '$' : '') + (t.r2 + 1);
  return '#REF!';
}

function serialize(toks, lang) {
  const tr = lang === 'tr'; let out = '';
  for (const t of toks) {
    switch (t.t) {
      case 'num': out += tr ? t.v.replace('.', ',') : t.v; break;
      case 'str': out += '"' + t.v.replace(/"/g, '""') + '"'; break;
      case 'bool': out += tr ? (t.v ? 'DOĞRU' : 'YANLIŞ') : (t.v ? 'TRUE' : 'FALSE'); break;
      case 'err': out += t.v; break;
      case 'ref': case 'cols': case 'rows': out += refText(t); break;
      case 'func': out += (tr && FN[t.v] ? FN[t.v].tr : t.v); break;
      case 'name': out += t.v; break;
      case 'op': out += t.v; break;
      case '(': out += '('; break;
      case ')': out += ')'; break;
      case 'sep': out += tr ? ';' : ','; break;
      case ':': out += ':'; break;
      case 'refErr': out += '#REF!'; break;
    }
  }
  return out;
}

// Görüntüleme: İngilizce saklanan formül -> Türkçe
function toDisplay(f) { try { return serialize(tokenize(f), 'tr'); } catch (e) { return f; } }
// Kayıt: kullanıcı girdisi (TR/EN) -> İngilizce
function toStore(f) { return serialize(tokenize(f), 'en'); }

// Göreli başvuruları kaydır (kopyala/doldur)
function shiftFormula(f, dr, dc) {
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    if (t.t === 'ref') {
      const r = t.ar ? t.r : t.r + dr, c = t.ac ? t.c : t.c + dc;
      if (r < 0 || c < 0 || r >= MAXR || c >= MAXC) toks[k] = { t: 'refErr' }; else { t.r = r; t.c = c; }
    } else if (t.t === 'cols') { if (!t.ac1) t.c1 += dc; if (!t.ac2) t.c2 += dc; if (t.c1 < 0 || t.c2 < 0) toks[k] = { t: 'refErr' }; }
    else if (t.t === 'rows') { if (!t.ar1) t.r1 += dr; if (!t.ar2) t.r2 += dr; if (t.r1 < 0 || t.r2 < 0) toks[k] = { t: 'refErr' }; }
  }
  return serialize(toks, 'en');
}

// Satır/sütun ekleme-silme sonrası başvuruları düzelt.
// axis: 'r' | 'c', at: indeks, n: >0 ekleme, <0 silme; sheetName: değişen sayfa; homeSheet: formülün sayfası
function adjustFormula(f, homeSheet, sheetName, axis, at, n) {
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  const hits = t => (t.sheet == null ? homeSheet : t.sheet) === sheetName;
  const mv = (v, isEnd) => {
    if (n > 0) return v >= at ? v + n : v;
    const d = -n; if (v < at) return v; if (v >= at + d) return v - d;
    return isEnd ? (at - 1) : (at); // silinen alana düşen uç
  };
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    const nx = toks[k + 1], nx2 = toks[k + 2];
    if (t.t === 'ref' && nx && nx.t === ':' && nx2 && nx2.t === 'ref' && hits(t)) {
      // aralık
      const key = axis === 'r' ? 'r' : 'c';
      let a = t[key], b = nx2[key]; const lo = Math.min(a, b), hi = Math.max(a, b);
      if (n < 0 && lo >= at && hi < at - n) { toks.splice(k, 3, { t: 'refErr' }); continue; }
      let nlo = mv(lo, false), nhi = mv(hi, true);
      if (n < 0 && lo >= at && lo < at - n) nlo = at;
      if (n < 0 && nhi < nlo) { toks.splice(k, 3, { t: 'refErr' }); continue; }
      t[key] = nlo; nx2[key] = nhi; k += 2; continue;
    }
    if (t.t === 'ref' && hits(t)) {
      const key = axis === 'r' ? 'r' : 'c';
      if (n < 0 && t[key] >= at && t[key] < at - n) { toks[k] = { t: 'refErr' }; continue; }
      t[key] = mv(t[key], false);
    } else if (t.t === 'cols' && axis === 'c' && hits(t)) { t.c1 = mv(t.c1, false); t.c2 = mv(t.c2, true); if (t.c2 < t.c1) toks[k] = { t: 'refErr' }; }
    else if (t.t === 'rows' && axis === 'r' && hits(t)) { t.r1 = mv(t.r1, false); t.r2 = mv(t.r2, true); if (t.r2 < t.r1) toks[k] = { t: 'refErr' }; }
  }
  return serialize(toks, 'en');
}
function renameSheetInFormula(f, oldN, newN) {
  let toks; try { toks = tokenize(f); } catch (e) { return f; }
  let ch = false; for (const t of toks) if ((t.t === 'ref' || t.t === 'cols' || t.t === 'rows') && t.sheet === oldN) { t.sheet = newN; ch = true; }
  return ch ? serialize(toks, 'en') : f;
}

// ---------- Ayrıştırıcı ----------
function parse(src) {
  const toks = tokenize(src); let p = 0;
  const peek = () => toks[p], next = () => toks[p++];
  const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  function expr(minP) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.t !== 'op' || !(t.v in PREC) || PREC[t.v] < minP) break;
      next(); const pr = PREC[t.v];
      const right = expr(t.v === '^' ? pr + 0.5 : pr + 1);
      left = { k: 'bin', op: t.v, a: left, b: right };
    }
    return left;
  }
  function unary() {
    const t = peek();
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) { next(); const a = unary(); return t.v === '-' ? { k: 'neg', a } : a; }
    let a = postfix(primary());
    return a;
  }
  function postfix(a) { while (peek() && peek().t === 'op' && peek().v === '%') { next(); a = { k: 'pct', a }; } return a; }
  function primary() {
    const t = next();
    if (!t) throw new Error('Eksik ifade');
    switch (t.t) {
      case 'num': return { k: 'v', v: parseFloat(t.v) };
      case 'str': return { k: 'v', v: t.v };
      case 'bool': return { k: 'v', v: t.v };
      case 'err': return { k: 'v', v: E(t.v) };
      case 'refErr': return { k: 'v', v: REF };
      case 'ref': {
        if (peek() && peek().t === ':' && toks[p + 1] && toks[p + 1].t === 'ref') {
          next(); const b = next();
          return { k: 'rng', sheet: t.sheet, r1: Math.min(t.r, b.r), c1: Math.min(t.c, b.c), r2: Math.max(t.r, b.r), c2: Math.max(t.c, b.c) };
        }
        return { k: 'ref', sheet: t.sheet, r: t.r, c: t.c };
      }
      case 'cols': return { k: 'rng', sheet: t.sheet, r1: 0, c1: Math.min(t.c1, t.c2), r2: MAXR - 1, c2: Math.max(t.c1, t.c2), whole: 1 };
      case 'rows': return { k: 'rng', sheet: t.sheet, r1: Math.min(t.r1, t.r2), c1: 0, r2: Math.max(t.r1, t.r2), c2: MAXC - 1, whole: 1 };
      case 'name': return { k: 'v', v: NAME };
      case '(': { const e = expr(0); if (!peek() || peek().t !== ')') throw new Error('Parantez eksik'); next(); return e; }
      case 'func': {
        if (!peek() || peek().t !== '(') throw new Error('Parantez eksik'); next();
        const args = [];
        if (peek() && peek().t === ')') { next(); return { k: 'call', name: t.v, args, known: t.known }; }
        for (;;) {
          if (peek() && (peek().t === 'sep' || peek().t === ')')) args.push({ k: 'missing' });
          else args.push(expr(0));
          const s = next();
          if (!s) throw new Error('Parantez eksik');
          if (s.t === ')') break;
          if (s.t !== 'sep') throw new Error('Ayraç bekleniyor');
        }
        return { k: 'call', name: t.v, args, known: t.known };
      }
    }
    throw new Error('Beklenmeyen öğe');
  }
  const ast = expr(0);
  if (p < toks.length) throw new Error('Fazla öğe');
  return ast;
}

// ---------- Değerlendirme ----------
class Range {
  constructor(model, si, r1, c1, r2, c2) {
    this.model = model; this.si = si; this.r1 = r1; this.c1 = c1;
    if (r2 >= MAXR - 1 || c2 >= MAXC - 1) { const d = model.dims(si); if (r2 >= MAXR - 1) r2 = Math.max(r1, d.rows - 1); if (c2 >= MAXC - 1) c2 = Math.max(c1, d.cols - 1); }
    this.r2 = r2; this.c2 = c2;
  }
  get h() { return this.r2 - this.r1 + 1; }
  get w() { return this.c2 - this.c1 + 1; }
  get(i, j) { return this.model.value(this.si, this.r1 + i, this.c1 + j); }
  *each() { for (let r = this.r1; r <= this.r2; r++) for (let c = this.c1; c <= this.c2; c++) yield this.model.value(this.si, r, c); }
  list() { return Array.from(this.each()); }
  grid() { const g = []; for (let i = 0; i < this.h; i++) { const row = []; for (let j = 0; j < this.w; j++) row.push(this.get(i, j)); g.push(row); } return g; }
  rowHidden(i) { return this.model.isRowHidden ? this.model.isRowHidden(this.si, this.r1 + i) : false; }
}
// 2B dizi (fonksiyon sonucu) için Range benzeri sarmalayıcı
class Arr {
  constructor(g) { this.g = g; }
  get h() { return this.g.length; } get w() { return this.g[0] ? this.g[0].length : 0; }
  get(i, j) { return this.g[i][j]; }
  *each() { for (const row of this.g) for (const v of row) yield v; }
  list() { return Array.from(this.each()); }
  grid() { return this.g; }
  rowHidden() { return false; }
}
const isRng = v => v instanceof Range || v instanceof Arr;
const LIFT_T = new Set(['num', 'text', 'date', 'bool']);

function scalar(v) {
  if (isRng(v)) { if (v.h === 1 && v.w === 1) return v.get(0, 0); return VAL; }
  return v;
}

function parseNumLoose(s) {
  if (typeof s !== 'string') return null;
  let t = s.trim(); if (!t) return null;
  let pct = false; if (t.endsWith('%')) { pct = true; t = t.slice(0, -1).trim(); }
  t = t.replace(/^[₺$€]\s*|\s*(₺|TL|\$|€)$/gi, '');
  let n = null;
  if (/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) n = parseFloat(t.replace(/\./g, '').replace(',', '.'));
  else if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t) && /\./.test(t)) n = parseFloat(t.replace(/,/g, ''));
  else if (/^[-+]?(\d+(,\d*)?|,\d+)(e[-+]?\d+)?$/i.test(t)) n = parseFloat(t.replace(',', '.'));
  else if (/^[-+]?(\d+(\.\d*)?|\.\d+)(e[-+]?\d+)?$/i.test(t)) n = parseFloat(t);
  if (n == null || isNaN(n)) return null;
  return pct ? n / 100 : n;
}
function parseDateLoose(s) {
  if (typeof s !== 'string') return null; const t = s.trim(); let m;
  if ((m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    let y = +m[3]; if (y < 100) y += y < 50 ? 2000 : 1900;
    const d = +m[1], mo = +m[2]; if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    let v = ymdToSerial(y, mo, d); if (m[4]) v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400;
    return { v, time: !!m[4] };
  }
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    let v = ymdToSerial(+m[1], +m[2], +m[3]); if (m[4]) v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400;
    return { v, time: !!m[4] };
  }
  return null;
}
function parseTimeLoose(s) {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(String(s).trim());
  return m ? (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400 : null;
}

function toNum(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = parseNumLoose(v); if (n != null) return n;
  const d = parseDateLoose(v); if (d) return d.v;
  const t = parseTimeLoose(v); if (t != null) return t;
  return VAL;
}
function toStr(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'DOĞRU' : 'YANLIŞ';
  if (typeof v === 'number') return fmtGeneral(v);
  return String(v);
}
function toBool(v) {
  v = scalar(v);
  if (isErr(v)) return v;
  if (v == null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const u = trUp(String(v));
  if (u === 'TRUE' || u === 'DOĞRU') return true;
  if (u === 'FALSE' || u === 'YANLIŞ') return false;
  return VAL;
}

function typeRank(v) { return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0; }
function cmp(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeRank(a), tb = typeRank(b);
  if (ta !== tb) return ta - tb;
  if (ta === 2) { const x = a.toLocaleLowerCase('tr-TR'), y = b.toLocaleLowerCase('tr-TR'); return x < y ? -1 : x > y ? 1 : 0; }
  if (ta === 3) return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

function evaluate(node, ctx) {
  switch (node.k) {
    case 'v': return node.v;
    case 'missing': return null;
    case 'ref': {
      const si = node.sheet == null ? ctx.si : ctx.model.sheetIndex(node.sheet);
      if (si < 0) return REF;
      return ctx.model.value(si, node.r, node.c);
    }
    case 'rng': {
      const si = node.sheet == null ? ctx.si : ctx.model.sheetIndex(node.sheet);
      if (si < 0) return REF;
      return new Range(ctx.model, si, node.r1, node.c1, node.r2, node.c2);
    }
    case 'neg': case 'pct': {
      const A0 = evaluate(node.a, ctx); const f = x => { const a = toNum(x); return isErr(a) ? a : node.k === 'neg' ? -a : a / 100; };
      if (isRng(A0) && A0.h * A0.w > 1) return new Arr(A0.grid().map(row => row.map(f)));
      return f(A0);
    }
    case 'bin': {
      const A0 = evaluate(node.a, ctx), B0 = evaluate(node.b, ctx);
      if ((isRng(A0) && A0.h * A0.w > 1) || (isRng(B0) && B0.h * B0.w > 1)) return broadcast(A0, B0, (x, y) => binOp(node.op, x, y));
      return binOp(node.op, scalar(A0), scalar(B0));
    }
    case 'call': {
      const d = FN[node.name];
      if (!d) return NAME;
      const args = node.args.map(a => d.rawRef && a.k === 'ref' ? refRange(a, ctx) : evaluate(a, ctx));
      // tek değer bekleyen yere liste verilirse her öğe için ayrı hesaplanır (Excel 365 gibi): MBUL("a";A1:A9)
      const lift = []; args.forEach((a, i) => { const ad = d.args[Math.min(i, d.args.length - 1)]; if (ad && !ad.many && isRng(a) && a.h * a.w > 1 && (LIFT_T.has(ad.t) || (ad.t === 'any' && d.lift))) lift.push(i); });
      if (lift.length) {
        let H = 1, W = 1; for (const i of lift) { H = Math.max(H, args[i].h); W = Math.max(W, args[i].w); }
        if (H * W <= 200000) {
          const out = []; for (let y = 0; y < H; y++) { const row = []; for (let x = 0; x < W; x++) {
            const a2 = args.slice(); for (const i of lift) { const g = args[i]; a2[i] = (y >= g.h && g.h > 1) || (x >= g.w && g.w > 1) ? NA : g.get(g.h === 1 ? 0 : y, g.w === 1 ? 0 : x); }
            let v; if (!d.errOk && a2.some(isErr)) v = a2.find(isErr); else { try { v = d.fn.apply(ctx, a2); if (typeof v === 'number' && !isFinite(v)) v = NUM; } catch (e) { v = isErr(e) ? e : VAL; } }
            row.push(isRng(v) ? scalar(v) : v === undefined ? null : v); } out.push(row); }
          return new Arr(out);
        }
      }
      if (!d.errOk) for (const a of args) if (isErr(a)) return a;
      try {
        const r = d.fn.apply(ctx, args);
        if (typeof r === 'number' && !isFinite(r)) return NUM;
        return r === undefined ? null : r;
      } catch (e) { return isErr(e) ? e : VAL; }
    }
  }
  return VAL;
}
function refRange(a, ctx) { const si = a.sheet == null ? ctx.si : ctx.model.sheetIndex(a.sheet); if (si < 0) return REF; return new Range(ctx.model, si, a.r, a.c, a.r, a.c); }
function broadcast(A, B, f) {
  const ga = isRng(A) ? A.grid() : [[A]], gb = isRng(B) ? B.grid() : [[B]];
  const H = Math.max(ga.length, gb.length), W = Math.max(ga[0].length, gb[0].length);
  const pick = (g, i, j) => { const r = g.length === 1 ? 0 : i, c = g[0].length === 1 ? 0 : j; return g[r] && g[r][c] !== undefined ? g[r][c] : NA; };
  const out = []; for (let i = 0; i < H; i++) { const row = []; for (let j = 0; j < W; j++) row.push(f(pick(ga, i, j), pick(gb, i, j))); out.push(row); }
  return new Arr(out);
}
function binOp(op, a, b) {
  {
      if (isErr(a)) return a; if (isErr(b)) return b;
      if (op === '&') return toStr(a) + toStr(b);
      if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
        const c = cmp(a, b);
        return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
      }
      const x = toNum(a), y = toNum(b);
      if (isErr(x)) return x; if (isErr(y)) return y;
      switch (op) {
        case '+': return x + y; case '-': return x - y; case '*': return x * y;
        case '/': return y === 0 ? DIV0 : x / y;
        case '^': { const r = Math.pow(x, y); return isFinite(r) ? r : NUM; }
      }
      return VAL;
  }
}

// ---------- Yardımcılar (fonksiyonlar için) ----------
function nums(args, opt) {
  // opt.a: metin->0, mantıksal sayılır (…A fonksiyonları)
  const out = [];
  for (const a of args) {
    if (isRng(a)) {
      for (const v of a.each()) {
        if (isErr(v)) throw v;
        if (typeof v === 'number') out.push(v);
        else if (opt && opt.a && v != null) out.push(typeof v === 'boolean' ? (v ? 1 : 0) : 0);
      }
    } else {
      if (a == null) continue;
      if (isErr(a)) throw a;
      const n = toNum(a); if (isErr(n)) throw n; out.push(n);
    }
  }
  return out;
}
function flat(args) { const out = []; for (const a of args) { if (isRng(a)) for (const v of a.each()) out.push(v); else out.push(a); } return out; }
function N(v) { const n = toNum(v); if (isErr(n)) throw n; return n; }
function S(v) { const s = toStr(v); if (isErr(s)) throw s; return s; }
function B(v) { const b = toBool(v); if (isErr(b)) throw b; return b; }
function opt(v, d) { return v == null ? d : v; }
function asRange(v) { if (isRng(v)) return v; return new Arr([[v]]); }
function wild(p) { return new RegExp('^' + p.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'is'); }

function crit(c) {
  c = scalar(c);
  if (typeof c === 'number') return v => (typeof v === 'number' && v === c) || (typeof v === 'string' && parseNumLoose(v) === c);
  if (typeof c === 'boolean') return v => v === c;
  if (c == null) c = '';
  const m = /^(<>|>=|<=|=|>|<)?([\s\S]*)$/.exec(String(c)); const op = m[1] || '='; const val = m[2];
  const num = parseNumLoose(val) ?? (parseDateLoose(val) ? parseDateLoose(val).v : null);
  if (op === '=' || op === '<>') {
    let f;
    if (val === '') f = v => v == null || v === '';
    else if (num != null) f = v => typeof v === 'number' ? Math.abs(v - num) < 1e-12 : (typeof v === 'string' && parseNumLoose(v) === num);
    else { const u = trUp(val); if (u === 'DOĞRU' || u === 'TRUE') f = v => v === true; else if (u === 'YANLIŞ' || u === 'FALSE') f = v => v === false; else { const re = wild(val); f = v => v != null && typeof v !== 'number' && re.test(String(v)); } }
    return op === '=' ? f : v => !f(v);
  }
  if (num != null) return v => typeof v === 'number' && (op === '>' ? v > num : op === '<' ? v < num : op === '>=' ? v >= num : v <= num);
  return v => typeof v === 'string' && (c => op === '>' ? c > 0 : op === '<' ? c < 0 : op === '>=' ? c >= 0 : c <= 0)(cmp(v, val));
}
// Çoklu ölçüt eşleşmesi: [rng1, crit1, rng2, crit2...] -> boolean dizi
function multiMask(pairs) {
  if (pairs.length % 2) throw VAL;
  let mask = null, len = 0;
  for (let k = 0; k < pairs.length; k += 2) {
    const r = asRange(pairs[k]), f = crit(pairs[k + 1]); const l = r.list();
    if (mask == null) { len = l.length; mask = l.map(v => f(v)); }
    else { if (l.length !== len) throw VAL; for (let i = 0; i < len; i++) mask[i] = mask[i] && f(l[i]); }
  }
  return mask || [];
}

function mean(a) { return a.reduce((s, x) => s + x, 0) / a.length; }
function variance(a, pop) { if (a.length < (pop ? 1 : 2)) throw DIV0; const m = mean(a); return a.reduce((s, x) => s + (x - m) * (x - m), 0) / (a.length - (pop ? 0 : 1)); }
function round(x, d) { const f = Math.pow(10, d); const y = +(Math.abs(x) * f).toPrecision(15); return Math.sign(x) * Math.round(y) / f; }
function sortedNums(args) { return nums(args).sort((a, b) => a - b); }
function percentile(a, k) { if (!a.length || k < 0 || k > 1) throw NUM; const s = a.slice().sort((x, y) => x - y); const h = (s.length - 1) * k; const lo = Math.floor(h); return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]); }
function ymd(s) { const d = serialToDate(Math.floor(N(s))); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay() }; }
function eq(a, b) { return cmp(a, b) === 0; }
function lookIndex(look, list, mode) {
  // mode 0 tam, 1 küçük-eşit(sıralı), -1 büyük-eşit (sıralı azalan), 2 joker
  if (mode === 0 || mode === 2) {
    if (typeof look === 'string' && (mode === 2 || /[*?]/.test(look))) { const re = wild(look); for (let i = 0; i < list.length; i++) if (list[i] != null && re.test(String(list[i]))) return i; return -1; }
    for (let i = 0; i < list.length; i++) if (list[i] != null && eq(list[i], look)) return i;
    return -1;
  }
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i]; if (v == null || typeRank(v) !== typeRank(look)) continue;
    const c = cmp(v, look);
    if (mode === 1) { if (c <= 0) best = i; else break; } else { if (c >= 0) best = i; else break; }
  }
  return best;
}

// ---------- Fonksiyonlar ----------
const A = (n, d, t, o) => Object.assign({ n, d, t: t || 'any' }, o || {});
const many = { many: true };
const M = 'Matematik', ST = 'İstatistik', L = 'Mantıksal', T = 'Metin', D = 'Tarih ve Saat', LK = 'Arama', F = 'Finans', I = 'Bilgi';

def('SUM', 'TOPLA', M, 'Seçilen sayıları toplar.', [A('Sayılar', 'Toplanacak hücreler', 'range', many)], (...a) => nums(a).reduce((s, x) => s + x, 0), { pop: 1 });
def('SUMIF', 'ETOPLA', M, 'Koşula uyan hücreleri toplar.', [A('Aranan alan', 'Koşulun bakılacağı hücreler', 'range'), A('Koşul', 'Ör: >100, Ankara, <>0', 'crit'), A('Toplam alanı', 'Toplanacak hücreler (boşsa ilk alan)', 'range', { o: 1 })],
  (r, c, s) => { const f = crit(c); const l = asRange(r).list(); const sl = s ? asRange(s).list() : l; let t = 0; l.forEach((v, i) => { if (f(v) && typeof sl[i] === 'number') t += sl[i]; }); return t; }, { pop: 1 });
def('SUMIFS', 'ÇOKETOPLA', M, 'Birden çok koşula uyanları toplar.', [A('Toplam alanı', 'Toplanacak hücreler', 'range'), A('Ölçüt alanı', '', 'range'), A('Ölçüt', '', 'crit'), A('Ek ölçüt…', '', 'any', { many: true, pair: true, o: 1, pairT: ['range', 'crit'], pairL: ['Ölçüt alanı', 'Ölçüt'] })],
  (s, ...p) => { const m = multiMask(p); const l = asRange(s).list(); let t = 0; l.forEach((v, i) => { if (m[i] && typeof v === 'number') t += v; }); return t; });
def('SUMPRODUCT', 'TOPLA.ÇARPIM', M, 'Aralıkları çarpıp toplar.', [A('Dizi', 'Aynı boyutlu alanlar', 'range', many)],
  (...a) => { const ls = a.map(x => asRange(x).list()); const n = ls[0].length; if (ls.some(l => l.length !== n)) throw VAL; let t = 0; for (let i = 0; i < n; i++) { let p = 1; for (const l of ls) { const v = l[i]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : 0; } t += p; } return t; });
def('PRODUCT', 'ÇARPIM', M, 'Sayıları çarpar.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); return n.length ? n.reduce((p, x) => p * x, 1) : 0; });
def('ROUND', 'YUVARLA', M, 'Sayıyı belirtilen basamağa yuvarlar.', [A('Sayı', '', 'num'), A('Basamak', 'Virgülden sonra kaç hane', 'num', { def: '2' })], (x, d) => round(N(x), N(opt(d, 0))), { pop: 1 });
def('ROUNDUP', 'YUKARIYUVARLA', M, 'Yukarı yuvarlar.', [A('Sayı', '', 'num'), A('Basamak', '', 'num', { def: '0' })], (x, d) => { x = N(x); const f = Math.pow(10, N(opt(d, 0))); return Math.sign(x) * Math.ceil(+(Math.abs(x) * f).toPrecision(15)) / f; });
def('ROUNDDOWN', 'AŞAĞIYUVARLA', M, 'Aşağı yuvarlar.', [A('Sayı', '', 'num'), A('Basamak', '', 'num', { def: '0' })], (x, d) => { x = N(x); const f = Math.pow(10, N(opt(d, 0))); return Math.sign(x) * Math.floor(+(Math.abs(x) * f).toPrecision(15)) / f; });
def('INT', 'TAMSAYI', M, 'Aşağı doğru en yakın tam sayı.', [A('Sayı', '', 'num')], x => Math.floor(N(x)));
def('TRUNC', 'NSAT', M, 'Ondalık kısmı atar.', [A('Sayı', '', 'num'), A('Basamak', '', 'num', { o: 1 })], (x, d) => { const f = Math.pow(10, N(opt(d, 0))); return Math.trunc(N(x) * f) / f; });
def('ABS', 'MUTLAK', M, 'Mutlak değer.', [A('Sayı', '', 'num')], x => Math.abs(N(x)));
def('MOD', 'MOD', M, 'Bölümden kalan.', [A('Sayı', '', 'num'), A('Bölen', '', 'num')], (x, y) => { x = N(x); y = N(y); if (y === 0) throw DIV0; return x - y * Math.floor(x / y); });
def('QUOTIENT', 'BÖLÜM', M, 'Bölümün tam kısmı.', [A('Pay', '', 'num'), A('Payda', '', 'num')], (x, y) => { y = N(y); if (!y) throw DIV0; return Math.trunc(N(x) / y); });
def('POWER', 'KUVVET', M, 'Üs alır.', [A('Sayı', '', 'num'), A('Üs', '', 'num')], (x, y) => Math.pow(N(x), N(y)));
def('SQRT', 'KAREKÖK', M, 'Karekök.', [A('Sayı', '', 'num')], x => { x = N(x); if (x < 0) throw NUM; return Math.sqrt(x); });
def('CEILING', 'TAVANAYUVARLA', M, 'Katına yukarı yuvarlar.', [A('Sayı', '', 'num'), A('Kat', '', 'num', { def: '1' })], (x, s) => { s = N(opt(s, 1)); if (!s) return 0; return Math.ceil(+(N(x) / s).toPrecision(15)) * s; });
def('FLOOR', 'TABANAYUVARLA', M, 'Katına aşağı yuvarlar.', [A('Sayı', '', 'num'), A('Kat', '', 'num', { def: '1' })], (x, s) => { s = N(opt(s, 1)); if (!s) throw DIV0; return Math.floor(+(N(x) / s).toPrecision(15)) * s; });
def('MROUND', 'KYUVARLA', M, 'En yakın katına yuvarlar.', [A('Sayı', '', 'num'), A('Kat', '', 'num')], (x, s) => { s = N(s); if (!s) return 0; return Math.round(N(x) / s) * s; });
def('PI', 'Pİ', M, 'π sayısı.', [], () => Math.PI);
def('RAND', 'S_SAYI_ÜRET', M, '0-1 arası rastgele sayı.', [], () => Math.random(), { vol: 1 });
def('RANDBETWEEN', 'RASTGELEARADA', M, 'İki sayı arasında rastgele tam sayı.', [A('Alt', '', 'num'), A('Üst', '', 'num')], (a, b) => { a = Math.ceil(N(a)); b = Math.floor(N(b)); return a + Math.floor(Math.random() * (b - a + 1)); }, { vol: 1 });
def('EXP', 'ÜS', M, 'e üzeri.', [A('Sayı', '', 'num')], x => Math.exp(N(x)));
def('LN', 'LN', M, 'Doğal logaritma.', [A('Sayı', '', 'num')], x => { x = N(x); if (x <= 0) throw NUM; return Math.log(x); });
def('LOG', 'LOG', M, 'Logaritma.', [A('Sayı', '', 'num'), A('Taban', '', 'num', { def: '10' })], (x, b) => { x = N(x); b = N(opt(b, 10)); if (x <= 0 || b <= 0) throw NUM; return Math.log(x) / Math.log(b); });
def('LOG10', 'LOG10', M, '10 tabanında logaritma.', [A('Sayı', '', 'num')], x => { x = N(x); if (x <= 0) throw NUM; return Math.log10(x); });
def('SIGN', 'İŞARET', M, 'Sayının işareti (1, 0, -1).', [A('Sayı', '', 'num')], x => Math.sign(N(x)));
def('EVEN', 'ÇİFT', M, 'En yakın çift sayıya yuvarlar.', [A('Sayı', '', 'num')], x => { x = N(x); const s = Math.sign(x) || 1; let v = Math.ceil(Math.abs(x)); if (v % 2) v++; return s * v; });
def('ODD', 'TEK', M, 'En yakın tek sayıya yuvarlar.', [A('Sayı', '', 'num')], x => { x = N(x); const s = Math.sign(x) || 1; let v = Math.ceil(Math.abs(x)); if (v % 2 === 0) v++; return s * v; });
def('FACT', 'ÇARPINIM', M, 'Faktöriyel.', [A('Sayı', '', 'num')], x => { x = Math.floor(N(x)); if (x < 0) throw NUM; let r = 1; for (let i = 2; i <= x; i++) r *= i; return r; });
def('GCD', 'OBEB', M, 'En büyük ortak bölen.', [A('Sayılar', '', 'range', many)], (...a) => nums(a).map(Math.floor).reduce((x, y) => { while (y) [x, y] = [y, x % y]; return x; }));
def('LCM', 'OKEK', M, 'En küçük ortak kat.', [A('Sayılar', '', 'range', many)], (...a) => nums(a).map(Math.floor).reduce((x, y) => { let g = x, h = y; while (h) [g, h] = [h, g % h]; return x * y / g; }));
def('SIN', 'SİN', M, 'Sinüs (radyan).', [A('Açı', '', 'num')], x => Math.sin(N(x)));
def('COS', 'COS', M, 'Kosinüs (radyan).', [A('Açı', '', 'num')], x => Math.cos(N(x)));
def('TAN', 'TAN', M, 'Tanjant (radyan).', [A('Açı', '', 'num')], x => Math.tan(N(x)));
def('DEGREES', 'DERECE', M, 'Radyanı dereceye çevirir.', [A('Açı', '', 'num')], x => N(x) * 180 / Math.PI);
def('RADIANS', 'RADYAN', M, 'Dereceyi radyana çevirir.', [A('Açı', '', 'num')], x => N(x) * Math.PI / 180);
def('SUBTOTAL', 'ALTTOPLAM', M, 'Gizli satırları dikkate alabilen ara toplam.', [A('İşlev no', '9=Toplam 1=Ortalama 2=Say 4=Mak 5=Min', 'num'), A('Alanlar', '', 'range', many)],
  function (k, ...a) { k = N(k); const skip = k > 100; k %= 100; const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
    if (!map[k]) throw VAL;
    const b = skip ? a.map(r => isRng(r) ? new Arr(r.grid().filter((_, i) => !r.rowHidden(i))) : r) : a;
    return FN[map[k]].fn.apply(this, b.map(r => (r instanceof Arr && !r.g.length) ? null : r)); });

def('AVERAGE', 'ORTALAMA', ST, 'Sayıların ortalaması.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); if (!n.length) throw DIV0; return mean(n); }, { pop: 1 });
def('AVERAGEIF', 'EĞERORTALAMA', ST, 'Koşula uyanların ortalaması.', [A('Aranan alan', '', 'range'), A('Koşul', '', 'crit'), A('Ortalama alanı', '', 'range', { o: 1 })],
  (r, c, s) => { const f = crit(c); const l = asRange(r).list(); const sl = s ? asRange(s).list() : l; let t = 0, k = 0; l.forEach((v, i) => { if (f(v) && typeof sl[i] === 'number') { t += sl[i]; k++; } }); if (!k) throw DIV0; return t / k; });
def('AVERAGEIFS', 'ÇOKEĞERORTALAMA', ST, 'Çok koşullu ortalama.', [A('Ortalama alanı', '', 'range'), A('Ölçüt alanı', '', 'range'), A('Ölçüt', '', 'crit'), A('Ek ölçüt…', '', 'any', { many: true, pair: true, o: 1, pairT: ['range', 'crit'], pairL: ['Ölçüt alanı', 'Ölçüt'] })],
  (s, ...p) => { const m = multiMask(p); const l = asRange(s).list(); let t = 0, k = 0; l.forEach((v, i) => { if (m[i] && typeof v === 'number') { t += v; k++; } }); if (!k) throw DIV0; return t / k; });
def('COUNT', 'BAĞ_DEĞ_SAY', ST, 'Sayı içeren hücreleri sayar.', [A('Alanlar', '', 'range', many)], (...a) => { let k = 0; for (const x of a) { if (isRng(x)) { for (const v of x.each()) if (typeof v === 'number') k++; } else if (x != null && !isErr(toNum(x))) k++; } return k; }, { errOk: 1, pop: 1 });
def('COUNTA', 'BAĞ_DEĞ_DOLU_SAY', ST, 'Dolu hücreleri sayar.', [A('Alanlar', '', 'range', many)], (...a) => flat(a).filter(v => v != null && v !== '').length, { errOk: 1, pop: 1 });
def('COUNTBLANK', 'BOŞLUKSAY', ST, 'Boş hücreleri sayar.', [A('Alan', '', 'range')], r => asRange(r).list().filter(v => v == null || v === '').length);
def('COUNTIF', 'EĞERSAY', ST, 'Koşula uyan hücreleri sayar.', [A('Alan', '', 'range'), A('Koşul', 'Ör: >50, Evet, <>""', 'crit')], (r, c) => { const f = crit(c); return asRange(r).list().filter(v => f(v)).length; }, { pop: 1 });
def('COUNTIFS', 'ÇOKEĞERSAY', ST, 'Birden çok koşula uyanları sayar.', [A('Ölçüt alanı', '', 'range'), A('Ölçüt', '', 'crit'), A('Ek ölçüt…', '', 'any', { many: true, pair: true, o: 1, pairT: ['range', 'crit'], pairL: ['Ölçüt alanı', 'Ölçüt'] })], (...p) => multiMask(p).filter(Boolean).length);
def('MAX', 'MAK', ST, 'En büyük değer.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; }, { pop: 1 });
def('MIN', 'MİN', ST, 'En küçük değer.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; }, { pop: 1 });
def('MAXIFS', 'ÇOKEĞERMAK', ST, 'Koşula uyanların en büyüğü.', [A('Mak alanı', '', 'range'), A('Ölçüt alanı', '', 'range'), A('Ölçüt', '', 'crit'), A('Ek ölçüt…', '', 'any', { many: true, pair: true, o: 1, pairT: ['range', 'crit'], pairL: ['Ölçüt alanı', 'Ölçüt'] })],
  (s, ...p) => { const m = multiMask(p); const v = asRange(s).list().filter((x, i) => m[i] && typeof x === 'number'); return v.length ? Math.max(...v) : 0; });
def('MINIFS', 'ÇOKEĞERMİN', ST, 'Koşula uyanların en küçüğü.', [A('Min alanı', '', 'range'), A('Ölçüt alanı', '', 'range'), A('Ölçüt', '', 'crit'), A('Ek ölçüt…', '', 'any', { many: true, pair: true, o: 1, pairT: ['range', 'crit'], pairL: ['Ölçüt alanı', 'Ölçüt'] })],
  (s, ...p) => { const m = multiMask(p); const v = asRange(s).list().filter((x, i) => m[i] && typeof x === 'number'); return v.length ? Math.min(...v) : 0; });
def('MEDIAN', 'ORTANCA', ST, 'Ortanca değer.', [A('Sayılar', '', 'range', many)], (...a) => { const s = sortedNums(a); if (!s.length) throw NUM; const h = s.length >> 1; return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2; });
def('MODE', 'ENÇOK_OLAN', ST, 'En sık tekrar eden değer.', [A('Sayılar', '', 'range', many)], (...a) => { const c = new Map(); let best = null, bc = 1; for (const x of nums(a)) { const k = (c.get(x) || 0) + 1; c.set(x, k); if (k > bc) { bc = k; best = x; } } if (best == null) throw NA; return best; });
def('STDEV', 'STDSAPMA', ST, 'Standart sapma (örneklem).', [A('Sayılar', '', 'range', many)], (...a) => Math.sqrt(variance(nums(a), false)));
def('STDEVP', 'STDSAPMAS', ST, 'Standart sapma (popülasyon).', [A('Sayılar', '', 'range', many)], (...a) => Math.sqrt(variance(nums(a), true)));
def('VAR', 'VAR', ST, 'Varyans (örneklem).', [A('Sayılar', '', 'range', many)], (...a) => variance(nums(a), false));
def('VARP', 'VARS', ST, 'Varyans (popülasyon).', [A('Sayılar', '', 'range', many)], (...a) => variance(nums(a), true));
def('LARGE', 'BÜYÜK', ST, 'k. en büyük değer.', [A('Alan', '', 'range'), A('Sıra (k)', '1 = en büyük', 'num', { def: '1' })], (r, k) => { const s = sortedNums([r]); k = Math.floor(N(k)); if (k < 1 || k > s.length) throw NUM; return s[s.length - k]; });
def('SMALL', 'KÜÇÜK', ST, 'k. en küçük değer.', [A('Alan', '', 'range'), A('Sıra (k)', '1 = en küçük', 'num', { def: '1' })], (r, k) => { const s = sortedNums([r]); k = Math.floor(N(k)); if (k < 1 || k > s.length) throw NUM; return s[k - 1]; });
def('RANK', 'RANK', ST, 'Sayının listedeki sırası.', [A('Sayı', '', 'num'), A('Liste', '', 'range'), A('Sıralama', '0 = büyükten küçüğe, 1 = küçükten büyüğe', 'enum', { o: 1, opts: [['0', 'Büyükten küçüğe'], ['1', 'Küçükten büyüğe']] })],
  (x, r, o) => { x = N(x); const s = nums([r]); if (!s.includes(x)) throw NA; return 1 + s.filter(v => N(opt(o, 0)) ? v < x : v > x).length; });
def('PERCENTILE', 'YÜZDEBİRLİK', ST, 'Yüzdelik değer.', [A('Alan', '', 'range'), A('Yüzde (0-1)', 'Ör: 0,9', 'num')], (r, k) => percentile(nums([r]), N(k)));
def('QUARTILE', 'DÖRTTEBİRLİK', ST, 'Çeyrek değer.', [A('Alan', '', 'range'), A('Çeyrek', '0-4', 'num')], (r, q) => percentile(nums([r]), N(q) / 4));
def('CORREL', 'KORELASYON', ST, 'İki seri arasındaki korelasyon.', [A('Seri 1', '', 'range'), A('Seri 2', '', 'range')], (a, b) => { const x = asRange(a).list(), y = asRange(b).list(); const p = []; x.forEach((v, i) => { if (typeof v === 'number' && typeof y[i] === 'number') p.push([v, y[i]]); }); if (p.length < 2) throw DIV0; const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); let sxy = 0, sxx = 0, syy = 0; p.forEach(([u, v]) => { sxy += (u - mx) * (v - my); sxx += (u - mx) ** 2; syy += (v - my) ** 2; }); return sxy / Math.sqrt(sxx * syy); });

def('IF', 'EĞER', L, 'Koşul doğruysa bir değer, yanlışsa başka değer.', [A('Koşul', 'Ör: Puan 50 den büyük', 'cond'), A('Doğruysa', 'Koşul sağlanırsa', 'any'), A('Yanlışsa', 'Sağlanmazsa', 'any', { o: 1 })],
  function IF(c, a, b) {
    if (isRng(c) && c.h * c.w > 1) { const el = (v, i, j) => isRng(v) ? (v.h * v.w === 1 ? v.get(0, 0) : v.get(Math.min(i, v.h - 1), Math.min(j, v.w - 1))) : v; return new Arr(c.grid().map((row, i) => row.map((x, j) => IF(x, el(a, i, j), el(b, i, j))))); }
    const x = toBool(c); if (isErr(x)) return x; return x ? (a === undefined ? true : a == null ? 0 : scalar(a)) : (b === undefined ? false : b == null ? 0 : scalar(b)); }, { errOk: 1 });
def('IFS', 'ÇOKEĞER', L, 'Sırayla koşulları dener.', [A('Koşul / sonuç…', 'Koşul-sonuç çiftleri', 'any', { many: true, pair: true, pairT: ['cond', 'any'], pairL: ['Koşul', 'Sonuç'] })], (...p) => { for (let i = 0; i + 1 < p.length; i += 2) { const b = toBool(p[i]); if (isErr(b)) return b; if (b) return scalar(p[i + 1]); } return NA; }, { errOk: 1 });
def('AND', 'VE', L, 'Tüm koşullar doğruysa DOĞRU.', [A('Koşullar', '', 'cond', many)], (...a) => { let any = false; for (const v of flat(a)) { if (isErr(v)) throw v; if (v == null || typeof v === 'string') continue; any = true; if (!B(v)) return false; } if (!any) throw VAL; return true; });
def('OR', 'YADA', L, 'Herhangi biri doğruysa DOĞRU.', [A('Koşullar', '', 'cond', many)], (...a) => { let any = false; for (const v of flat(a)) { if (isErr(v)) throw v; if (v == null || typeof v === 'string') continue; any = true; if (B(v)) return true; } if (!any) throw VAL; return false; });
def('XOR', 'ÖZELVEYA', L, 'Tek sayıda koşul doğruysa DOĞRU.', [A('Koşullar', '', 'cond', many)], (...a) => flat(a).filter(v => v != null && typeof v !== 'string' && B(v)).length % 2 === 1);
def('NOT', 'DEĞİL', L, 'Mantıksal değeri tersine çevirir.', [A('Değer', '', 'any')], v => !B(v));
def('IFERROR', 'EĞERHATA', L, 'Hata varsa yerine başka değer gösterir.', [A('Değer', '', 'any'), A('Hata olursa', 'Ör: 0 veya ""', 'any')], (v, a) => { v = scalar(v); return isErr(v) ? scalar(a) : v; }, { errOk: 1 });
def('IFNA', 'EĞERYOKSA', L, '#YOK hatası yerine değer gösterir.', [A('Değer', '', 'any'), A('Yoksa', '', 'any')], (v, a) => { v = scalar(v); return v === NA ? scalar(a) : v; }, { errOk: 1 });
def('SWITCH', 'İLKEŞLEŞEN', L, 'Değeri seçeneklerle karşılaştırır.', [A('İfade', '', 'any'), A('Değer / sonuç…', '', 'any', { many: true, pair: true, pairT: ['any', 'any'], pairL: ['Eşleşecek değer', 'Sonuç'] })], (x, ...p) => { x = scalar(x); for (let i = 0; i + 1 < p.length; i += 2) if (eq(x, scalar(p[i]))) return scalar(p[i + 1]); return p.length % 2 ? scalar(p[p.length - 1]) : NA; });
def('TRUE', 'DOĞRU', L, 'DOĞRU değeri.', [], () => true);
def('FALSE', 'YANLIŞ', L, 'YANLIŞ değeri.', [], () => false);

def('CONCATENATE', 'BİRLEŞTİR', T, 'Metinleri birleştirir.', [A('Metinler', '', 'any', many)], (...a) => a.map(x => S(x)).join(''));
def('CONCAT', 'ARALIKBİRLEŞTİR', T, 'Alanlardaki metinleri birleştirir.', [A('Metinler', '', 'range', many)], (...a) => flat(a).map(x => S(x)).join(''));
def('TEXTJOIN', 'METİNBİRLEŞTİR', T, 'Ayraçla birleştirir.', [A('Ayraç', 'Ör: ", "', 'text'), A('Boşları atla', '', 'bool', { def: 'TRUE' }), A('Metinler', '', 'range', many)], (d, ig, ...a) => flat(a).map(x => S(x)).filter(x => !B(ig) || x !== '').join(S(d)));
def('LEFT', 'SOLDAN', T, 'Soldan karakter alır.', [A('Metin', '', 'text'), A('Karakter sayısı', '', 'num', { def: '1' })], (s, n) => S(s).slice(0, N(opt(n, 1))));
def('RIGHT', 'SAĞDAN', T, 'Sağdan karakter alır.', [A('Metin', '', 'text'), A('Karakter sayısı', '', 'num', { def: '1' })], (s, n) => { n = N(opt(n, 1)); s = S(s); return n ? s.slice(-n) : ''; });
def('MID', 'PARÇAAL', T, 'Ortadan parça alır.', [A('Metin', '', 'text'), A('Başlangıç', '', 'num'), A('Karakter sayısı', '', 'num')], (s, a, n) => S(s).substr(N(a) - 1, N(n)));
def('LEN', 'UZUNLUK', T, 'Karakter sayısı.', [A('Metin', '', 'text')], s => S(s).length);
def('UPPER', 'BÜYÜKHARF', T, 'BÜYÜK HARFE çevirir.', [A('Metin', '', 'text')], s => trUp(S(s)));
def('LOWER', 'KÜÇÜKHARF', T, 'küçük harfe çevirir.', [A('Metin', '', 'text')], s => S(s).toLocaleLowerCase('tr-TR'));
def('PROPER', 'YAZIM.DÜZENİ', T, 'Her Kelimenin İlk Harfini Büyütür.', [A('Metin', '', 'text')], s => S(s).toLocaleLowerCase('tr-TR').replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + trUp(b)));
def('TRIM', 'KIRP', T, 'Fazla boşlukları siler.', [A('Metin', '', 'text')], s => S(s).replace(/ +/g, ' ').trim());
def('CLEAN', 'TEMİZ', T, 'Yazdırılamayan karakterleri siler.', [A('Metin', '', 'text')], s => S(s).replace(/[\x00-\x1f]/g, ''));
def('SUBSTITUTE', 'YERİNEKOY', T, 'Metindeki bir parçayı değiştirir.', [A('Metin', '', 'text'), A('Eski', '', 'text'), A('Yeni', '', 'text'), A('Kaçıncı', '', 'num', { o: 1 })],
  (s, o, n, k) => { s = S(s); o = S(o); n = S(n); if (!o) return s; if (k == null) return s.split(o).join(n); k = N(k); let i = -1; for (let j = 0; j < k; j++) { i = s.indexOf(o, i + 1); if (i < 0) return s; } return s.slice(0, i) + n + s.slice(i + o.length); });
def('REPLACE', 'DEĞİŞTİR', T, 'Konuma göre değiştirir.', [A('Metin', '', 'text'), A('Başlangıç', '', 'num'), A('Karakter sayısı', '', 'num'), A('Yeni metin', '', 'text')], (s, a, n, t) => { s = S(s); a = N(a) - 1; return s.slice(0, a) + S(t) + s.slice(a + N(n)); });
def('FIND', 'BUL', T, 'Metin içinde arar (büyük/küçük harf duyarlı).', [A('Aranan', '', 'text'), A('Metin', '', 'text'), A('Başlangıç', '', 'num', { o: 1 })], (f, s, a) => { const i = S(s).indexOf(S(f), N(opt(a, 1)) - 1); if (i < 0) throw VAL; return i + 1; });
def('SEARCH', 'MBUL', T, 'Metin içinde arar.', [A('Aranan', '', 'text'), A('Metin', '', 'text'), A('Başlangıç', '', 'num', { o: 1 })], (f, s, a) => { const i = S(s).toLocaleLowerCase('tr-TR').indexOf(S(f).toLocaleLowerCase('tr-TR'), N(opt(a, 1)) - 1); if (i < 0) throw VAL; return i + 1; });
def('TEXT', 'METNEÇEVİR', T, 'Sayıyı biçimli metne çevirir.', [A('Değer', '', 'num'), A('Biçim', 'Ör: #.##0,00 veya gg.aa.yyyy', 'text')], (v, f) => { const n = toNum(v); return isErr(n) ? S(v) : formatValue(n, trFormatCode(S(f))); });
def('VALUE', 'SAYIYAÇEVİR', T, 'Metni sayıya çevirir.', [A('Metin', '', 'text')], s => N(S(s)));
def('REPT', 'YİNELE', T, 'Metni tekrarlar.', [A('Metin', '', 'text'), A('Tekrar', '', 'num')], (s, n) => S(s).repeat(Math.max(0, N(n))));
def('EXACT', 'ÖZDEŞ', T, 'İki metin tamamen aynı mı?', [A('Metin 1', '', 'text'), A('Metin 2', '', 'text')], (a, b) => S(a) === S(b));
def('CHAR', 'DAMGA', T, 'Kodun karakteri.', [A('Kod', '', 'num')], n => String.fromCharCode(N(n)));
def('CODE', 'KOD', T, 'Karakterin kodu.', [A('Metin', '', 'text')], s => { s = S(s); if (!s) throw VAL; return s.charCodeAt(0); });

def('TODAY', 'BUGÜN', D, 'Bugünün tarihi.', [], () => Math.floor(nowSerial()), { vol: 1, fmt: 'dd.mm.yyyy' });
def('NOW', 'ŞİMDİ', D, 'Şu anki tarih ve saat.', [], () => nowSerial(), { vol: 1, fmt: 'dd.mm.yyyy hh:mm' });
def('DATE', 'TARİH', D, 'Yıl, ay, günden tarih oluşturur.', [A('Yıl', '', 'num'), A('Ay', '', 'num'), A('Gün', '', 'num')], (y, m, d) => ymdToSerial(N(y), N(m), N(d)), { fmt: 'dd.mm.yyyy' });
def('TIME', 'ZAMAN', D, 'Saat, dakika, saniyeden zaman.', [A('Saat', '', 'num'), A('Dakika', '', 'num'), A('Saniye', '', 'num')], (h, m, s) => ((N(h) * 3600 + N(m) * 60 + N(s)) / 86400) % 1, { fmt: 'hh:mm' });
def('YEAR', 'YIL', D, 'Tarihin yılı.', [A('Tarih', '', 'date')], s => ymd(s).y);
def('MONTH', 'AY', D, 'Tarihin ayı.', [A('Tarih', '', 'date')], s => ymd(s).m);
def('DAY', 'GÜN', D, 'Tarihin günü.', [A('Tarih', '', 'date')], s => ymd(s).d);
def('HOUR', 'SAAT', D, 'Saat.', [A('Zaman', '', 'date')], s => Math.floor(((N(s) % 1) * 86400 + 0.5) / 3600) % 24);
def('MINUTE', 'DAKİKA', D, 'Dakika.', [A('Zaman', '', 'date')], s => Math.floor(((N(s) % 1) * 86400 + 0.5) / 60) % 60);
def('SECOND', 'SANİYE', D, 'Saniye.', [A('Zaman', '', 'date')], s => Math.floor((N(s) % 1) * 86400 + 0.5) % 60);
def('WEEKDAY', 'HAFTANINGÜNÜ', D, 'Haftanın kaçıncı günü.', [A('Tarih', '', 'date'), A('Tür', '', 'enum', { o: 1, opts: [['2', 'Pazartesi=1 … Pazar=7'], ['1', 'Pazar=1 … Cumartesi=7'], ['3', 'Pazartesi=0 … Pazar=6']] })],
  (s, t) => { const w = ymd(s).wd; t = N(opt(t, 1)); return t === 2 ? (w + 6) % 7 + 1 : t === 3 ? (w + 6) % 7 : w + 1; });
def('WEEKNUM', 'HAFTASAY', D, 'Yılın kaçıncı haftası.', [A('Tarih', '', 'date'), A('Tür', '1=Pazar 2=Pazartesi başlar', 'num', { o: 1 })],
  (s, t) => { s = Math.floor(N(s)); const y = ymd(s).y; const j1 = ymdToSerial(y, 1, 1); const st = N(opt(t, 1)) === 2 ? 1 : 0; const wd = (ymd(j1).wd - st + 7) % 7; return Math.floor((s - j1 + wd) / 7) + 1; });
def('EDATE', 'SERİTARİH', D, 'Ay ekler/çıkarır.', [A('Başlangıç', '', 'date'), A('Ay sayısı', '', 'num')], (s, m) => { const d = ymd(s); const nm = d.m - 1 + N(m); const y = d.y + Math.floor(nm / 12), mo = ((nm % 12) + 12) % 12; const last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate(); return ymdToSerial(y, mo + 1, Math.min(d.d, last)); }, { fmt: 'dd.mm.yyyy' });
def('EOMONTH', 'SERİAY', D, 'Ayın son günü.', [A('Başlangıç', '', 'date'), A('Ay sayısı', '0 = bu ay', 'num', { def: '0' })], (s, m) => { const d = ymd(s); return Date.UTC(d.y, d.m + N(m), 0) / DAYMS + 25569; }, { fmt: 'dd.mm.yyyy' });
def('DATEDIF', 'ETARİHLİ', D, 'İki tarih arası fark.', [A('Başlangıç', '', 'date'), A('Bitiş', '', 'date'), A('Birim', '', 'enum', { opts: [['"Y"', 'Yıl'], ['"M"', 'Ay'], ['"D"', 'Gün'], ['"YM"', 'Yıl hariç ay'], ['"MD"', 'Ay hariç gün']] })],
  (a, b, u) => { a = Math.floor(N(a)); b = Math.floor(N(b)); if (a > b) throw NUM; const x = ymd(a), y = ymd(b); u = trUp(S(u));
    let months = (y.y - x.y) * 12 + y.m - x.m; if (y.d < x.d) months--;
    switch (u) { case 'Y': return Math.floor(months / 12); case 'M': return months; case 'D': return b - a; case 'YM': return months % 12;
      case 'MD': { let d = y.d - x.d; if (d < 0) d += new Date(Date.UTC(y.y, y.m - 1, 0)).getUTCDate(); return d; }
      case 'YD': { let s2 = ymdToSerial(y.y, x.m, x.d); if (s2 > b) s2 = ymdToSerial(y.y - 1, x.m, x.d); return b - s2; } }
    throw NUM; });
def('DAYS', 'GÜNSAY', D, 'İki tarih arası gün.', [A('Bitiş', '', 'date'), A('Başlangıç', '', 'date')], (b, a) => Math.floor(N(b)) - Math.floor(N(a)));
function isWork(s, hol) { const w = ymd(s).wd; return w !== 0 && w !== 6 && !hol.has(s); }
function holSet(h) { return new Set(h == null ? [] : nums([h]).map(Math.floor)); }
def('NETWORKDAYS', 'TAMİŞGÜNÜ', D, 'İki tarih arası iş günü.', [A('Başlangıç', '', 'date'), A('Bitiş', '', 'date'), A('Tatiller', '', 'range', { o: 1 })],
  (a, b, h) => { a = Math.floor(N(a)); b = Math.floor(N(b)); const hs = holSet(h); const sg = a <= b ? 1 : -1; let k = 0; for (let s = Math.min(a, b); s <= Math.max(a, b); s++) if (isWork(s, hs)) k++; return k * sg; });
def('WORKDAY', 'İŞGÜNÜ', D, 'İş günü ekler.', [A('Başlangıç', '', 'date'), A('Gün', '', 'num'), A('Tatiller', '', 'range', { o: 1 })],
  (a, n, h) => { let s = Math.floor(N(a)); n = Math.trunc(N(n)); const hs = holSet(h); const st = n < 0 ? -1 : 1; while (n) { s += st; if (isWork(s, hs)) n -= st; } return s; }, { fmt: 'dd.mm.yyyy' });

def('VLOOKUP', 'DÜŞEYARA', LK, 'Tablonun ilk sütununda arar, aynı satırdan değer getirir.', [A('Aranan değer', 'Ne arıyorsunuz?', 'any'), A('Tablo', 'Arama yapılacak tablo', 'range'), A('Sütun no', 'Tablodan kaçıncı sütun gelsin', 'num'), A('Eşleşme', '', 'enum', { def: 'FALSE', opts: [['FALSE', 'Tam eşleşme'], ['TRUE', 'Yaklaşık (sıralı liste)']] })],
  (v, t, c, ap) => { v = scalar(v); if (isErr(v)) return v; t = asRange(t); c = Math.floor(N(c)); if (c < 1 || c > t.w) throw REF; const col = []; for (let i = 0; i < t.h; i++) col.push(t.get(i, 0)); const i = lookIndex(v, col, B(opt(ap, true)) ? 1 : 0); if (i < 0) throw NA; return t.get(i, c - 1); }, { errOk: 1 });
def('HLOOKUP', 'YATAYARA', LK, 'Tablonun ilk satırında arar.', [A('Aranan değer', '', 'any'), A('Tablo', '', 'range'), A('Satır no', '', 'num'), A('Eşleşme', '', 'enum', { def: 'FALSE', opts: [['FALSE', 'Tam eşleşme'], ['TRUE', 'Yaklaşık']] })],
  (v, t, r, ap) => { v = scalar(v); if (isErr(v)) return v; t = asRange(t); r = Math.floor(N(r)); if (r < 1 || r > t.h) throw REF; const row = []; for (let j = 0; j < t.w; j++) row.push(t.get(0, j)); const j = lookIndex(v, row, B(opt(ap, true)) ? 1 : 0); if (j < 0) throw NA; return t.get(r - 1, j); }, { errOk: 1 });
def('XLOOKUP', 'ÇAPRAZARA', LK, 'Bir listede arar, diğer listeden karşılığını getirir.', [A('Aranan değer', '', 'any'), A('Arama listesi', '', 'range'), A('Sonuç listesi', '', 'range'), A('Bulunamazsa', '', 'any', { o: 1 }), A('Eşleşme', '', 'enum', { o: 1, opts: [['0', 'Tam'], ['-1', 'Tam veya küçük'], ['1', 'Tam veya büyük'], ['2', 'Joker (*?)']] })],
  (v, la, ra, nf, mm) => { v = scalar(v); if (isErr(v)) return v; la = asRange(la); ra = asRange(ra); const l = la.list(); mm = N(opt(mm, 0)); let i;
    if (mm === 0 || mm === 2) i = lookIndex(v, l, mm); else { i = -1; let best = null; l.forEach((x, k) => { if (x == null || typeRank(x) !== typeRank(v)) return; const c = cmp(x, v); if (c === 0) { if (i < 0 || cmp(best, v) !== 0) { i = k; best = x; } } else if (mm === -1 ? c < 0 : c > 0) { if (i < 0 || (cmp(best, v) !== 0 && (mm === -1 ? cmp(x, best) > 0 : cmp(x, best) < 0))) { i = k; best = x; } } }); }
    if (i < 0) { if (nf !== undefined && nf !== null) return scalar(nf); throw NA; }
    const vertical = la.w === 1; return vertical ? ra.get(i, 0) : ra.get(0, i); }, { errOk: 1 });
def('INDEX', 'İNDİS', LK, 'Tablodaki satır/sütun kesişimini getirir.', [A('Tablo', '', 'range'), A('Satır no', '', 'num'), A('Sütun no', '', 'num', { o: 1 })],
  (t, r, c) => { t = asRange(t); r = Math.floor(N(opt(r, 0))); c = Math.floor(N(opt(c, 0))); if (t.h === 1 && c === 0) { c = r; r = 1; } if (c === 0) c = 1; if (r === 0) r = 1; if (r < 1 || r > t.h || c < 1 || c > t.w) throw REF; return t.get(r - 1, c - 1); });
def('MATCH', 'KAÇINCI', LK, 'Değerin listedeki sırası.', [A('Aranan değer', '', 'any'), A('Liste', '', 'range'), A('Eşleşme', '', 'enum', { def: '0', opts: [['0', 'Tam eşleşme'], ['1', 'Küçük veya eşit (artan sıralı)'], ['-1', 'Büyük veya eşit (azalan sıralı)']] })],
  (v, l, m) => { v = scalar(v); if (isErr(v)) return v; const i = lookIndex(v, asRange(l).list(), N(opt(m, 1))); if (i < 0) throw NA; return i + 1; }, { errOk: 1 });
def('CHOOSE', 'ELEMAN', LK, 'Sıra numarasına göre değer seçer.', [A('Sıra', '', 'num'), A('Değerler', '', 'any', many)], (k, ...v) => { k = Math.floor(N(k)); if (k < 1 || k > v.length) throw VAL; return scalar(v[k - 1]); });
def('ROW', 'SATIR', LK, 'Satır numarası.', [A('Hücre', '', 'range', { o: 1 })], function (r) { return r instanceof Range ? r.r1 + 1 : this.r + 1; }, { rawRef: 1 });
def('COLUMN', 'SÜTUN', LK, 'Sütun numarası.', [A('Hücre', '', 'range', { o: 1 })], function (r) { return r instanceof Range ? r.c1 + 1 : this.c + 1; }, { rawRef: 1 });
def('ROWS', 'SATIRSAY', LK, 'Alandaki satır sayısı.', [A('Alan', '', 'range')], r => asRange(r).h);
def('COLUMNS', 'SÜTUNSAY', LK, 'Alandaki sütun sayısı.', [A('Alan', '', 'range')], r => asRange(r).w);

def('ISBLANK', 'EBOŞSA', I, 'Hücre boş mu?', [A('Değer', '', 'any')], v => { v = scalar(v); return v == null; }, { errOk: 1, lift: 1 });
def('ISNUMBER', 'ESAYIYSA', I, 'Sayı mı?', [A('Değer', '', 'any')], v => typeof scalar(v) === 'number', { errOk: 1, lift: 1 });
def('ISTEXT', 'EMETİNSE', I, 'Metin mi?', [A('Değer', '', 'any')], v => typeof scalar(v) === 'string', { errOk: 1, lift: 1 });
def('ISERROR', 'EHATALIYSA', I, 'Hata mı?', [A('Değer', '', 'any')], v => isErr(scalar(v)), { errOk: 1, lift: 1 });
def('ISNA', 'EYOKSA', I, '#YOK hatası mı?', [A('Değer', '', 'any')], v => scalar(v) === NA, { errOk: 1, lift: 1 });
def('ISEVEN', 'ÇİFTMİ', I, 'Çift mi?', [A('Sayı', '', 'num')], v => Math.floor(Math.abs(N(v))) % 2 === 0);
def('ISODD', 'TEKMİ', I, 'Tek mi?', [A('Sayı', '', 'num')], v => Math.floor(Math.abs(N(v))) % 2 === 1);
def('NA', 'YOKSAY', I, '#N/A hatası.', [], () => NA);

def('PMT', 'DEVRESEL_ÖDEME', F, 'Kredi taksiti.', [A('Faiz (dönemlik)', 'Ör: yıllık %24 aylık için 0,02', 'num'), A('Taksit sayısı', '', 'num'), A('Kredi tutarı', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Ödeme zamanı', '0 = dönem sonu', 'num', { o: 1 })],
  (r, n, pv, fv, t) => { r = N(r); n = N(n); pv = N(pv); fv = N(opt(fv, 0)); t = N(opt(t, 0)); if (!r) return -(pv + fv) / n; const q = Math.pow(1 + r, n); return -(r * (fv + pv * q)) / ((1 + r * t) * (q - 1)); }, { fmt: '#,##0.00' });
def('FV', 'GD', F, 'Gelecek değer.', [A('Faiz', '', 'num'), A('Dönem', '', 'num'), A('Ödeme', '', 'num'), A('Bugünkü değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 })],
  (r, n, p, pv, t) => { r = N(r); n = N(n); p = N(p); pv = N(opt(pv, 0)); t = N(opt(t, 0)); if (!r) return -(pv + p * n); const q = Math.pow(1 + r, n); return -(pv * q + p * (1 + r * t) * (q - 1) / r); });
def('PV', 'BD', F, 'Bugünkü değer.', [A('Faiz', '', 'num'), A('Dönem', '', 'num'), A('Ödeme', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 })],
  (r, n, p, fv, t) => { r = N(r); n = N(n); p = N(p); fv = N(opt(fv, 0)); t = N(opt(t, 0)); if (!r) return -(fv + p * n); const q = Math.pow(1 + r, n); return -(fv + p * (1 + r * t) * (q - 1) / r) / q; });
def('NPV', 'NBD', F, 'Net bugünkü değer.', [A('İskonto oranı', '', 'num'), A('Nakit akışları', '', 'range', many)], (r, ...a) => { r = N(r); return nums(a).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); });
def('IRR', 'İÇ_VERİM_ORANI', F, 'İç verim oranı.', [A('Nakit akışları', '', 'range'), A('Tahmin', '', 'num', { o: 1 })],
  (a, g) => { const v = nums([a]); let r = N(opt(g, 0.1)); for (let k = 0; k < 100; k++) { let f = 0, d = 0; v.forEach((x, i) => { f += x / Math.pow(1 + r, i); d -= i * x / Math.pow(1 + r, i + 1); }); if (!d) break; const nr = r - f / d; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw NUM; }, { fmt: '0.00%' });

// ---------- Ek işlevler ----------
const alias = (en, tr, base, desc) => { const b = FN[base]; def(en, tr, b.cat, desc || b.desc, b.args, b.fn, { errOk: b.errOk, fmt: b.fmt, rawRef: b.rawRef }); };
const grid2 = v => isRng(v) ? v.grid() : [[v]];
const toArr = g => g.length ? new Arr(g) : NA;

// Matematik
def('SUMSQ', 'TOPKARE', M, 'Sayıların karelerinin toplamı.', [A('Sayılar', '', 'range', many)], (...a) => nums(a).reduce((s, x) => s + x * x, 0));
def('COMBIN', 'KOMBİNASYON', M, 'Kaç farklı grup seçilebilir (sırasız).', [A('Toplam', '', 'num'), A('Seçilen', '', 'num')], (n, k) => { n = Math.floor(N(n)); k = Math.floor(N(k)); if (k < 0 || n < k) throw NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); });
def('PERMUT', 'PERMÜTASYON', M, 'Kaç farklı sıralama yapılabilir.', [A('Toplam', '', 'num'), A('Seçilen', '', 'num')], (n, k) => { n = Math.floor(N(n)); k = Math.floor(N(k)); if (k < 0 || n < k) throw NUM; let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; });
def('ROMAN', 'ROMEN', M, 'Sayıyı Romen rakamına çevirir.', [A('Sayı', '', 'num')], x => { x = Math.floor(N(x)); if (x < 0 || x > 3999) throw VAL; const m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; for (const [v, t] of m) while (x >= v) { s += t; x -= v; } return s; });
def('ARABIC', 'ARAP', M, 'Romen rakamını sayıya çevirir.', [A('Romen rakamı', '', 'text')], s => { s = S(s).toUpperCase().trim(); const v = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }; let t = 0; for (let i = 0; i < s.length; i++) { const a = v[s[i]], b = v[s[i + 1]] || 0; if (!a) throw VAL; t += a < b ? -a : a; } return t; });
def('DEC2BIN', 'ONDALIKTANİKİLİYE', M, 'Ondalık sayıyı ikiliye çevirir.', [A('Sayı', '', 'num')], x => Math.floor(N(x)).toString(2));
def('BIN2DEC', 'İKİLİKTENONDALIĞA', M, 'İkili sayıyı ondalığa çevirir.', [A('İkili sayı', '', 'text')], s => { const n = parseInt(S(s), 2); if (isNaN(n)) throw NUM; return n; });
def('DEC2HEX', 'ONDALIKTANONALTILIYA', M, 'Ondalık sayıyı onaltılığa çevirir.', [A('Sayı', '', 'num')], x => Math.floor(N(x)).toString(16).toUpperCase());
def('HEX2DEC', 'ONALTILIKTANONDALIĞA', M, 'Onaltılık sayıyı ondalığa çevirir.', [A('Onaltılık sayı', '', 'text')], s => { const n = parseInt(S(s), 16); if (isNaN(n)) throw NUM; return n; });
def('ASIN', 'ASİN', M, 'Ark sinüs.', [A('Sayı', '', 'num')], x => Math.asin(N(x)));
def('ACOS', 'ACOS', M, 'Ark kosinüs.', [A('Sayı', '', 'num')], x => Math.acos(N(x)));
def('ATAN', 'ATAN', M, 'Ark tanjant.', [A('Sayı', '', 'num')], x => Math.atan(N(x)));
def('ATAN2', 'ATAN2', M, 'x, y koordinatının açısı.', [A('x', '', 'num'), A('y', '', 'num')], (x, y) => Math.atan2(N(y), N(x)));
alias('CEILING.MATH', 'TAVANAYUVARLA.MATEMATİK', 'CEILING');
alias('FLOOR.MATH', 'TABANAYUVARLA.MATEMATİK', 'FLOOR');
def('RANDARRAY', 'RASTGELEDİZİ', M, 'Rastgele sayılardan oluşan bir tablo üretir.', [A('Satır', '', 'num', { def: '5' }), A('Sütun', '', 'num', { o: 1 }), A('En az', '', 'num', { o: 1 }), A('En çok', '', 'num', { o: 1 }), A('Tam sayı', '', 'bool', { o: 1 })],
  (r, c, lo, hi, w) => { r = N(opt(r, 1)); c = N(opt(c, 1)); lo = N(opt(lo, 0)); hi = N(opt(hi, 1)); const whole = w != null && B(w); const g = []; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)); g.push(row); } return toArr(g); }, { vol: 1 });

// İstatistik
def('AVERAGEA', 'ORTALAMAA', ST, 'Metin ve mantıksal değerleri de sayarak ortalama.', [A('Değerler', '', 'range', many)], (...a) => { const n = nums(a, { a: 1 }); if (!n.length) throw DIV0; return mean(n); });
def('MAXA', 'MAKA', ST, 'Metinleri 0 sayarak en büyük.', [A('Değerler', '', 'range', many)], (...a) => { const n = nums(a, { a: 1 }); return n.length ? Math.max(...n) : 0; });
def('MINA', 'MİNA', ST, 'Metinleri 0 sayarak en küçük.', [A('Değerler', '', 'range', many)], (...a) => { const n = nums(a, { a: 1 }); return n.length ? Math.min(...n) : 0; });
alias('STDEV.S', 'STDSAPMA.S', 'STDEV'); alias('STDEV.P', 'STDSAPMA.P', 'STDEVP'); alias('VAR.S', 'VAR.S', 'VAR'); alias('VAR.P', 'VAR.P', 'VARP');
alias('MODE.SNGL', 'ENÇOK_OLAN.TEK', 'MODE'); alias('RANK.EQ', 'RANK.EŞİT', 'RANK');
alias('PERCENTILE.INC', 'YÜZDEBİRLİK.DHL', 'PERCENTILE'); alias('QUARTILE.INC', 'DÖRTTEBİRLİK.DHL', 'QUARTILE');
def('RANK.AVG', 'RANK.ORT', ST, 'Sıra (eşitlerde ortalama sıra).', [A('Sayı', '', 'num'), A('Liste', '', 'range'), A('Sıralama', '', 'enum', { o: 1, opts: [['0', 'Büyükten küçüğe'], ['1', 'Küçükten büyüğe']] })],
  (x, r, o) => { x = N(x); const s = nums([r]); const eqn = s.filter(v => v === x).length; if (!eqn) throw NA; const before = s.filter(v => N(opt(o, 0)) ? v < x : v > x).length; return before + (eqn + 1) / 2; });
def('PERCENTRANK', 'YÜZDERANK', ST, 'Değerin yüzde kaçlık dilimde olduğu.', [A('Liste', '', 'range'), A('Değer', '', 'num')], (r, x) => { const s = sortedNums([r]); x = N(x); if (!s.length || x < s[0] || x > s[s.length - 1]) throw NA; const below = s.filter(v => v < x).length; return Math.floor(below / (s.length - 1) * 1000) / 1000; }, { fmt: '0.0%' });
def('TRIMMEAN', 'KIRPORTALAMA', ST, 'En uç değerleri atarak ortalama.', [A('Liste', '', 'range'), A('Atılacak oran', 'Ör: 0,2 = %20', 'num')], (r, p) => { const s = sortedNums([r]); const k = Math.floor(s.length * N(p) / 2); const t = s.slice(k, s.length - k); if (!t.length) throw NUM; return mean(t); });
def('GEOMEAN', 'GEOORT', ST, 'Geometrik ortalama (büyüme oranları için).', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); if (!n.length || n.some(x => x <= 0)) throw NUM; return Math.exp(n.reduce((s, x) => s + Math.log(x), 0) / n.length); });
def('HARMEAN', 'HARORT', ST, 'Harmonik ortalama.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); if (!n.length || n.some(x => x <= 0)) throw NUM; return n.length / n.reduce((s, x) => s + 1 / x, 0); });
def('AVEDEV', 'ORTSAP', ST, 'Ortalamadan ortalama sapma.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); const m = mean(n); return n.reduce((s, x) => s + Math.abs(x - m), 0) / n.length; });
def('DEVSQ', 'SAPKARE', ST, 'Sapmaların kareleri toplamı.', [A('Sayılar', '', 'range', many)], (...a) => { const n = nums(a); const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0); });
function pairs(ya, xa) { const y = asRange(ya).list(), x = asRange(xa).list(); if (y.length !== x.length) throw NA; const p = []; y.forEach((v, i) => { if (typeof v === 'number' && typeof x[i] === 'number') p.push([x[i], v]); }); if (p.length < 2) throw DIV0; const mx = mean(p.map(q => q[0])), my = mean(p.map(q => q[1])); let sxy = 0, sxx = 0, syy = 0; p.forEach(([u, v]) => { sxy += (u - mx) * (v - my); sxx += (u - mx) ** 2; syy += (v - my) ** 2; }); if (!sxx) throw DIV0; return { mx, my, sxy, sxx, syy, b: sxy / sxx }; }
def('SLOPE', 'EĞİM', ST, 'Doğrusal eğilimin eğimi.', [A('Y değerleri', '', 'range'), A('X değerleri', '', 'range')], (y, x) => pairs(y, x).b);
def('INTERCEPT', 'KESMENOKTASI', ST, 'Doğrunun Y eksenini kestiği nokta.', [A('Y değerleri', '', 'range'), A('X değerleri', '', 'range')], (y, x) => { const p = pairs(y, x); return p.my - p.b * p.mx; });
def('RSQ', 'RKARE', ST, 'R kare (uyum gücü).', [A('Y değerleri', '', 'range'), A('X değerleri', '', 'range')], (y, x) => { const p = pairs(y, x); return p.sxy * p.sxy / (p.sxx * p.syy); });
def('PEARSON', 'PEARSON', ST, 'Pearson korelasyon katsayısı.', [A('Seri 1', '', 'range'), A('Seri 2', '', 'range')], (a, b) => FN.CORREL.fn(a, b));
def('FORECAST.LINEAR', 'TAHMİN.DOĞRUSAL', ST, 'Geçmiş verilerden gelecek değeri tahmin eder.', [A('Tahmin edilecek X', 'Ör: gelecek ayın numarası', 'num'), A('Bilinen Y değerleri', 'Ör: satışlar', 'range'), A('Bilinen X değerleri', 'Ör: ay numaraları', 'range')], (x, y, xs) => { const p = pairs(y, xs); return p.my - p.b * p.mx + p.b * N(x); });
alias('FORECAST', 'TAHMİN', 'FORECAST.LINEAR');
// EĞİLİM / BÜYÜME: bilinen değerlerden doğrusal / üstel eğilimle yeni değerler (liste döndürebilir)
const fitXY = (ya, xa, log) => { const y = asRange(ya).list(); const x = xa == null ? y.map((_, i) => i + 1) : asRange(xa).list(); if (y.length !== x.length) throw REF; const p = []; y.forEach((v, i) => { if (typeof v === 'number' && typeof x[i] === 'number' && (!log || v > 0)) p.push([x[i], log ? Math.log(v) : v]); }); if (p.length < 2) throw DIV0; const n = p.length, mx = p.reduce((a, q) => a + q[0], 0) / n, my = p.reduce((a, q) => a + q[1], 0) / n; let sxy = 0, sxx = 0; p.forEach(([a, b]) => { sxy += (a - mx) * (b - my); sxx += (a - mx) * (a - mx); }); const b = sxx ? sxy / sxx : 0; return { at: v => log ? Math.exp(my - b * mx + b * v) : my - b * mx + b * v, n: y.length }; };
const fitOut = (f, nx) => { if (nx == null) return new Arr(Array.from({ length: f.n }, (_, i) => [f.at(i + 1)])); if (isRng(nx)) return new Arr(nx.grid().map(row => row.map(v => typeof v === 'number' ? f.at(v) : VAL))); return f.at(N(nx)); };
def('TREND', 'EĞİLİM', ST, 'Doğrusal eğilimle yeni değerleri tahmin eder (bir ya da birden çok).', [A('Bilinen Y değerleri', 'Ör: satışlar', 'range'), A('Bilinen X değerleri', 'Boş bırakılırsa 1, 2, 3…', 'range', { o: 1 }), A('Yeni X değerleri', 'Ör: gelecek dönem numaraları', 'range', { o: 1 })], (y, x, nx) => fitOut(fitXY(y, x == null || x === '' ? null : x), nx == null || nx === '' ? null : nx));
def('GROWTH', 'BÜYÜME', ST, 'Üstel büyümeyle (yüzdesel artış) yeni değerleri tahmin eder.', [A('Bilinen Y değerleri', 'Pozitif sayılar', 'range'), A('Bilinen X değerleri', 'Boş bırakılırsa 1, 2, 3…', 'range', { o: 1 }), A('Yeni X değerleri', '', 'range', { o: 1 })], (y, x, nx) => fitOut(fitXY(y, x == null || x === '' ? null : x, true), nx == null || nx === '' ? null : nx));
const erf = x => { const s = Math.sign(x); x = Math.abs(x); const t = 1 / (1 + 0.3275911 * x); const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x); return s * y; };
const ncdf = z => 0.5 * (1 + erf(z / Math.SQRT2));
function ninv(p) { if (p <= 0 || p >= 1) throw NUM; const a = [-39.6968302866538, 220.946098424521, -275.928510446969, 138.357751867269, -30.6647980661472, 2.50662827745924], b = [-54.4760987982241, 161.585836858041, -155.698979859887, 66.8013118877197, -13.2806815528857], c = [-7.78489400243029e-3, -0.322396458041136, -2.40075827716184, -2.54973253934373, 4.37466414146497, 2.93816398269878], d = [7.78469570904146e-3, 0.32246712907004, 2.445134137143, 3.75440866190742]; const pl = 0.02425; let q, r; if (p < pl) { q = Math.sqrt(-2 * Math.log(p)); return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); } if (p > 1 - pl) { q = Math.sqrt(-2 * Math.log(1 - p)); return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); } q = p - 0.5; r = q * q; return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
def('NORM.DIST', 'NORM.DAĞ', ST, 'Normal dağılım.', [A('x', '', 'num'), A('Ortalama', '', 'num'), A('Standart sapma', '', 'num'), A('Birikimli', '', 'bool', { def: 'TRUE' })], (x, m, s, c) => { x = N(x); m = N(m); s = N(s); if (s <= 0) throw NUM; const z = (x - m) / s; return B(c) ? ncdf(z) : Math.exp(-z * z / 2) / (s * Math.sqrt(2 * Math.PI)); });
def('NORM.INV', 'NORM.TERS', ST, 'Normal dağılımın tersi.', [A('Olasılık', '', 'num'), A('Ortalama', '', 'num'), A('Standart sapma', '', 'num')], (p, m, s) => N(m) + N(s) * ninv(N(p)));
def('NORM.S.DIST', 'NORM.S.DAĞ', ST, 'Standart normal dağılım.', [A('z', '', 'num'), A('Birikimli', '', 'bool', { def: 'TRUE' })], (z, c) => { z = N(z); return B(opt(c, true)) ? ncdf(z) : Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI); });
def('NORM.S.INV', 'NORM.S.TERS', ST, 'Standart normal dağılımın tersi.', [A('Olasılık', '', 'num')], p => ninv(N(p)));
def('STANDARDIZE', 'STANDARTLAŞTIRMA', ST, 'Z puanı.', [A('x', '', 'num'), A('Ortalama', '', 'num'), A('Standart sapma', '', 'num')], (x, m, s) => { s = N(s); if (s <= 0) throw NUM; return (N(x) - N(m)) / s; });
def('BINOM.DIST', 'BİNOM.DAĞ', ST, 'Binom dağılımı.', [A('Başarı sayısı', '', 'num'), A('Deneme', '', 'num'), A('Olasılık', '', 'num'), A('Birikimli', '', 'bool')], (k, n, p, c) => { k = Math.floor(N(k)); n = Math.floor(N(n)); p = N(p); const pmf = i => FN.COMBIN.fn(n, i) * p ** i * (1 - p) ** (n - i); if (!B(c)) return pmf(k); let s = 0; for (let i = 0; i <= k; i++) s += pmf(i); return s; });
def('POISSON.DIST', 'POISSON.DAĞ', ST, 'Poisson dağılımı.', [A('x', '', 'num'), A('Ortalama', '', 'num'), A('Birikimli', '', 'bool')], (x, m, c) => { x = Math.floor(N(x)); m = N(m); const pmf = i => { let f = 1; for (let j = 2; j <= i; j++) f *= j; return Math.exp(-m) * m ** i / f; }; if (!B(c)) return pmf(x); let s = 0; for (let i = 0; i <= x; i++) s += pmf(i); return s; });

// Mantıksal / bilgi
def('ISLOGICAL', 'EMANTIKSALSA', I, 'Mantıksal değer mi?', [A('Değer', '', 'any')], v => typeof scalar(v) === 'boolean', { errOk: 1 });
def('ISNONTEXT', 'EMETİNDEĞİLSE', I, 'Metin değil mi?', [A('Değer', '', 'any')], v => typeof scalar(v) !== 'string', { errOk: 1 });
def('ISERR', 'EHATA', I, '#YOK dışında bir hata mı?', [A('Değer', '', 'any')], v => { v = scalar(v); return isErr(v) && v !== NA; }, { errOk: 1 });
def('ISFORMULA', 'EFORMÜLSE', I, 'Hücrede formül var mı?', [A('Hücre', '', 'range')], function (r) { if (!(r instanceof Range)) return false; const c = this.model.cell(r.si, r.r1, r.c1); return !!(c && c.f); }, { rawRef: 1 });
def('N', 'S', I, 'Değeri sayıya çevirir (metin 0).', [A('Değer', '', 'any')], v => { v = scalar(v); return typeof v === 'number' ? v : v === true ? 1 : 0; });
def('T', 'M', T, 'Değer metinse metni, değilse boş döndürür.', [A('Değer', '', 'any')], v => { v = scalar(v); return typeof v === 'string' ? v : ''; });
def('TYPE', 'TÜR', I, 'Değerin türü (1 sayı, 2 metin, 4 mantıksal, 16 hata, 64 dizi).', [A('Değer', '', 'any')], v => isRng(v) ? 64 : (v = scalar(v), typeof v === 'number' || v == null ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 16), { errOk: 1 });
def('ERROR.TYPE', 'HATA.TİPİ', I, 'Hatanın numarası.', [A('Hata', '', 'any')], v => { v = scalar(v); if (!isErr(v)) return NA; return { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7 }[v.code] || 8; }, { errOk: 1 });

// Metin
function nth(s, d, k, fromEnd) { let i = -1; if (!fromEnd) { for (let j = 0; j < k; j++) { i = s.indexOf(d, i + 1); if (i < 0) return -1; } return i; } i = s.length; for (let j = 0; j < k; j++) { i = s.lastIndexOf(d, i - 1); if (i < 0) return -1; } return i; }
def('TEXTBEFORE', 'ÖNCEKİMETİN', T, 'Bir işaretten önceki metni alır (ör: e-postada @ öncesi).', [A('Metin', '', 'text'), A('İşaret', 'Ör: @ veya -', 'text'), A('Kaçıncı', '-1 = sondan', 'num', { o: 1 })], (s, d, k) => { s = S(s); d = S(d); k = Math.trunc(N(opt(k, 1))); const i = nth(s, d, Math.abs(k), k < 0); if (i < 0) throw NA; return s.slice(0, i); });
def('TEXTAFTER', 'SONRAKİMETİN', T, 'Bir işaretten sonraki metni alır.', [A('Metin', '', 'text'), A('İşaret', 'Ör: @ veya -', 'text'), A('Kaçıncı', '-1 = sondan', 'num', { o: 1 })], (s, d, k) => { s = S(s); d = S(d); k = Math.trunc(N(opt(k, 1))); const i = nth(s, d, Math.abs(k), k < 0); if (i < 0) throw NA; return s.slice(i + d.length); });
def('TEXTSPLIT', 'METİNBÖL', T, 'Metni işarete göre hücrelere böler (ör: "Ad Soyad" → 2 hücre).', [A('Metin', '', 'text'), A('Sütun işareti', 'Ör: boşluk için " "', 'text'), A('Satır işareti', '', 'text', { o: 1 })],
  (s, cd, rd) => { s = S(s); cd = S(cd); const rows = rd != null && S(rd) ? s.split(S(rd)) : [s]; const g = rows.map(r => cd ? r.split(cd) : [r]); const w = Math.max(...g.map(r => r.length)); return toArr(g.map(r => { while (r.length < w) r.push(NA); return r; })); });
def('FIXED', 'SAYIDÜZENLE', T, 'Sayıyı belirli ondalıkla metne çevirir.', [A('Sayı', '', 'num'), A('Ondalık', '', 'num', { def: '2' }), A('Binlik ayraç olmasın', '', 'bool', { o: 1 })], (x, d, nc) => { d = Math.max(0, N(opt(d, 2))); const code = (nc != null && B(nc) ? '0' : '#,##0') + (d ? '.' + '0'.repeat(d) : ''); return formatValue(round(N(x), d), code); });
def('DOLLAR', 'LİRA', T, 'Sayıyı para biçimli metne çevirir.', [A('Sayı', '', 'num'), A('Ondalık', '', 'num', { def: '2' })], (x, d) => { d = Math.max(0, N(opt(d, 2))); return formatValue(round(N(x), d), '#,##0' + (d ? '.' + '0'.repeat(d) : '') + ' ₺'); });
def('NUMBERVALUE', 'SAYIDEĞERİ', T, 'Metni sayıya çevirir (ayraçları belirterek).', [A('Metin', '', 'text'), A('Ondalık ayracı', '', 'text', { o: 1 }), A('Binlik ayracı', '', 'text', { o: 1 })], (s, ds, gs) => { s = S(s).trim(); if (ds != null || gs != null) { const g = gs != null ? S(gs) : (S(ds) === ',' ? '.' : ','); s = s.split(g).join('').replace(S(opt(ds, '.')), '.'); const n = parseFloat(s); if (isNaN(n)) throw VAL; return n; } return N(s); });
def('UNICHAR', 'UNICODEKARAKTERİ', T, 'Kod numarasının karakteri.', [A('Kod', '', 'num')], n => String.fromCodePoint(N(n)));
def('UNICODE', 'UNICODE', T, 'Karakterin kod numarası.', [A('Metin', '', 'text')], s => { s = S(s); if (!s) throw VAL; return s.codePointAt(0); });

// Tarih
def('DATEVALUE', 'TARİHSAYISI', D, 'Metin halindeki tarihi gerçek tarihe çevirir.', [A('Tarih metni', 'Ör: 30.10.2026', 'text')], s => { const d = parseDateLoose(S(s)); if (!d) throw VAL; return Math.floor(d.v); }, { fmt: 'dd.mm.yyyy' });
def('TIMEVALUE', 'ZAMANSAYISI', D, 'Metin halindeki saati gerçek saate çevirir.', [A('Saat metni', 'Ör: 14:30', 'text')], s => { const t = parseTimeLoose(S(s)); if (t == null) { const d = parseDateLoose(S(s)); if (d) return d.v % 1; throw VAL; } return t; }, { fmt: 'hh:mm' });
def('DAYS360', 'GÜN360', D, '360 günlük yıla göre gün farkı (muhasebe).', [A('Başlangıç', '', 'date'), A('Bitiş', '', 'date'), A('Avrupa yöntemi', '', 'bool', { o: 1 })], (a, b, eu) => { const x = ymd(a), y = ymd(b); let d1 = x.d, d2 = y.d; if (eu != null && B(eu)) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; } else { if (d1 === 31) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; } return (y.y - x.y) * 360 + (y.m - x.m) * 30 + d2 - d1; });
def('YEARFRAC', 'YILORAN', D, 'İki tarih arası yılın kesri.', [A('Başlangıç', '', 'date'), A('Bitiş', '', 'date'), A('Esas', '0=30/360 1=gerçek 2=gerçek/360 3=gerçek/365', 'num', { o: 1 })], (a, b, bs) => { a = Math.floor(N(a)); b = Math.floor(N(b)); if (a > b) [a, b] = [b, a]; bs = N(opt(bs, 0)); if (bs === 0 || bs === 4) return FN.DAYS360.fn(a, b, bs === 4) / 360; if (bs === 2) return (b - a) / 360; if (bs === 3) return (b - a) / 365; const y1 = ymd(a).y, y2 = ymd(b).y; let days = 0; for (let y = y1; y <= y2; y++) days += (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 366 : 365; return (b - a) / (days / (y2 - y1 + 1)); });
def('ISOWEEKNUM', 'ISOHAFTASAY', D, 'ISO hafta numarası (Pazartesi başlar).', [A('Tarih', '', 'date')], s => { const d = serialToDate(Math.floor(N(s))); const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())); const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day); const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - y0) / DAYMS + 1) / 7); });
function weekendFn(w) { w = w == null ? 1 : w; if (typeof w === 'string') { const m = w.split('').map(x => x === '1'); return d => m[(d + 6) % 7]; } w = N(w); const map = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] }; if (w >= 11 && w <= 17) { const one = (w - 11 + 0) % 7; const day = [0, 1, 2, 3, 4, 5, 6][(one + 0) % 7]; return d => d === day; } const pr = map[w] || map[1]; return d => pr.includes(d); }
def('NETWORKDAYS.INTL', 'TAMİŞGÜNÜ.ULUSL', D, 'Hafta sonu günlerini seçerek iş günü sayısı.', [A('Başlangıç', '', 'date'), A('Bitiş', '', 'date'), A('Hafta sonu', '', 'enum', { o: 1, opts: [['1', 'Cumartesi-Pazar'], ['11', 'Yalnız Pazar'], ['17', 'Yalnız Cumartesi'], ['7', 'Cuma-Cumartesi']] }), A('Tatiller', '', 'range', { o: 1 })],
  (a, b, w, h) => { a = Math.floor(N(a)); b = Math.floor(N(b)); const we = weekendFn(w); const hs = holSet(h); let k = 0; for (let s = Math.min(a, b); s <= Math.max(a, b); s++) if (!we(ymd(s).wd) && !hs.has(s)) k++; return a <= b ? k : -k; });
def('WORKDAY.INTL', 'İŞGÜNÜ.ULUSL', D, 'Hafta sonu günlerini seçerek iş günü ekler.', [A('Başlangıç', '', 'date'), A('Gün', '', 'num'), A('Hafta sonu', '', 'enum', { o: 1, opts: [['1', 'Cumartesi-Pazar'], ['11', 'Yalnız Pazar'], ['17', 'Yalnız Cumartesi'], ['7', 'Cuma-Cumartesi']] }), A('Tatiller', '', 'range', { o: 1 })],
  (a, n, w, h) => { let s = Math.floor(N(a)); n = Math.trunc(N(n)); const we = weekendFn(w); const hs = holSet(h); const st = n < 0 ? -1 : 1; while (n) { s += st; if (!we(ymd(s).wd) && !hs.has(s)) n -= st; } return s; }, { fmt: 'dd.mm.yyyy' });

// Arama ve başvuru
def('LOOKUP', 'ARA', LK, 'Sıralı listede arar, karşılığını getirir.', [A('Aranan', '', 'any'), A('Arama listesi', '', 'range'), A('Sonuç listesi', '', 'range', { o: 1 })], (v, l, r) => { v = scalar(v); const ll = asRange(l).list(); const i = lookIndex(v, ll, 1); if (i < 0) throw NA; return (r ? asRange(r).list() : ll)[i]; }, { errOk: 1 });
def('XMATCH', 'ÇAPRAZEŞLEŞTİR', LK, 'Değerin listedeki sırası (gelişmiş).', [A('Aranan', '', 'any'), A('Liste', '', 'range'), A('Eşleşme', '', 'enum', { o: 1, opts: [['0', 'Tam'], ['-1', 'Tam veya küçük'], ['1', 'Tam veya büyük'], ['2', 'Joker']] }), A('Arama yönü', '', 'enum', { o: 1, opts: [['1', 'Baştan'], ['-1', 'Sondan']] })],
  (v, l, m, sm) => { v = scalar(v); let ls = asRange(l).list(); m = N(opt(m, 0)); const rev = N(opt(sm, 1)) === -1; if (rev) ls = ls.slice().reverse(); let i; if (m === 0 || m === 2) i = lookIndex(v, ls, m); else { i = -1; let best = null; ls.forEach((x, k) => { if (x == null || typeRank(x) !== typeRank(v)) return; const c = cmp(x, v); if (c === 0 && (i < 0 || cmp(best, v) !== 0)) { i = k; best = x; } else if (c !== 0 && (m === -1 ? c < 0 : c > 0) && (i < 0 || (cmp(best, v) !== 0 && (m === -1 ? cmp(x, best) > 0 : cmp(x, best) < 0)))) { i = k; best = x; } }); } if (i < 0) throw NA; return rev ? ls.length - i : i + 1; }, { errOk: 1 });
def('OFFSET', 'KAYDIR', LK, 'Bir hücreden belli sayıda satır/sütun ötedeki alanı verir.', [A('Başlangıç hücresi', '', 'range'), A('Satır kaydır', '', 'num'), A('Sütun kaydır', '', 'num'), A('Yükseklik', '', 'num', { o: 1 }), A('Genişlik', '', 'num', { o: 1 })],
  function (r, dr, dc, hh, ww) { if (!(r instanceof Range)) throw VAL; const r1 = r.r1 + Math.trunc(N(dr)), c1 = r.c1 + Math.trunc(N(dc)); const H = Math.trunc(N(opt(hh, r.h))), W = Math.trunc(N(opt(ww, r.w))); if (r1 < 0 || c1 < 0 || H < 1 || W < 1) throw REF; return new Range(this.model, r.si, r1, c1, r1 + H - 1, c1 + W - 1); }, { rawRef: 1, vol: 1 });
def('INDIRECT', 'DOLAYLI', LK, 'Metin olarak yazılmış adresi gerçek başvuruya çevirir.', [A('Adres metni', 'Ör: "B"&A1 veya Sayfa2!A1', 'text')],
  function (t) { t = S(t).trim(); try { const ast = parse(t); if (ast.k === 'ref') { const si = ast.sheet == null ? this.si : this.model.sheetIndex(ast.sheet); if (si < 0) throw REF; return new Range(this.model, si, ast.r, ast.c, ast.r, ast.c); } if (ast.k === 'rng') { const si = ast.sheet == null ? this.si : this.model.sheetIndex(ast.sheet); if (si < 0) throw REF; return new Range(this.model, si, ast.r1, ast.c1, ast.r2, ast.c2); } } catch (e) { } throw REF; }, { vol: 1 });
def('ADDRESS', 'ADRES', LK, 'Satır ve sütun numarasından adres metni üretir.', [A('Satır no', '', 'num'), A('Sütun no', '', 'num'), A('Mutlak', '1=$A$1 4=A1', 'num', { o: 1 }), A('A1 stili', '', 'bool', { o: 1 }), A('Sayfa adı', '', 'text', { o: 1 })],
  (r, c, ab, a1, sh) => { r = Math.floor(N(r)); c = Math.floor(N(c)); ab = N(opt(ab, 1)); const cr = ab === 1 || ab === 3 ? '$' : '', rr = ab === 1 || ab === 2 ? '$' : ''; return (sh != null ? quoteSheet(S(sh)) + '!' : '') + cr + colName(c - 1) + rr + r; });
def('HYPERLINK', 'KÖPRÜ', LK, 'Bağlantı metni.', [A('Adres', '', 'text'), A('Görünen ad', '', 'text', { o: 1 })], (u, n) => n != null ? S(n) : S(u));
def('TRANSPOSE', 'DEVRİK_DÖNÜŞÜM', LK, 'Satırları sütun, sütunları satır yapar (sonuç yana taşar).', [A('Alan', '', 'range')], r => { const g = grid2(r); return toArr(g[0].map((_, j) => g.map(row => row[j]))); });
def('FILTER', 'FİLTRE', LK, 'Koşula uyan satırları listeler (sonuç aşağı taşar).', [A('Listelenecek alan', 'Ör: tüm tablo', 'range'), A('Koşul', 'Ör: Şehir sütunu = Ankara', 'cond'), A('Hiç yoksa', '', 'any', { o: 1 })],
  (a, inc, ie) => { const g = grid2(a), m = grid2(inc); const byRow = m.length === g.length; const out = []; if (byRow) g.forEach((row, i) => { const v = m[i][0]; if (isErr(v)) throw v; if (v === true || (typeof v === 'number' && v !== 0)) out.push(row); }); else { const cols = []; m[0].forEach((v, j) => { if (v === true || (typeof v === 'number' && v)) cols.push(j); }); g.forEach(row => out.push(cols.map(j => row[j]))); if (!cols.length) out.length = 0; } if (!out.length) { if (ie !== undefined) return scalar(ie); throw E('#HESAP!'); } return toArr(out); }, { errOk: 1 });
def('SORT', 'SIRALA', LK, 'Alanı sıralı olarak listeler (sonuç taşar).', [A('Alan', '', 'range'), A('Hangi sütuna göre', '1 = ilk sütun', 'num', { o: 1 }), A('Yön', '', 'enum', { o: 1, opts: [['1', 'Artan (A→Z)'], ['-1', 'Azalan (Z→A)']] })],
  (a, k, o) => { const g = grid2(a).map(r => r.slice()); k = Math.floor(N(opt(k, 1))) - 1; o = N(opt(o, 1)); g.sort((x, y) => cmp(x[k], y[k]) * (o < 0 ? -1 : 1)); return toArr(g); });
def('SORTBY', 'SIRALAÖLÇÜT', LK, 'Alanı başka bir sütuna göre sıralar.', [A('Alan', '', 'range'), A('Ölçüt sütunu', '', 'range'), A('Yön', '', 'enum', { o: 1, opts: [['1', 'Artan'], ['-1', 'Azalan']] })],
  (a, by, o) => { const g = grid2(a); const b = asRange(by).list(); o = N(opt(o, 1)); const idx = g.map((_, i) => i).sort((i, j) => cmp(b[i], b[j]) * (o < 0 ? -1 : 1) || i - j); return toArr(idx.map(i => g[i])); });
def('UNIQUE', 'BENZERSİZ', LK, 'Tekrarsız (benzersiz) değerleri listeler (sonuç taşar).', [A('Alan', '', 'range'), A('Sütunlara göre', '', 'bool', { o: 1 }), A('Yalnız bir kez geçenler', '', 'bool', { o: 1 })],
  (a, bc, once) => { const g = grid2(a); const key = r => r.map(v => v == null ? '' : typeof v === 'string' ? 'S' + v.toLocaleLowerCase('tr-TR') : String(v)).join('\u0001'); const cnt = new Map(); g.forEach(r => { const k = key(r); cnt.set(k, (cnt.get(k) || 0) + 1); }); const seen = new Set(), out = []; g.forEach(r => { const k = key(r); if (seen.has(k)) return; seen.add(k); if (once != null && B(once) && cnt.get(k) > 1) return; out.push(r); }); return toArr(out); });
def('SEQUENCE', 'SIRALI', LK, '1, 2, 3… gibi sıralı sayılar üretir (sonuç taşar).', [A('Satır sayısı', '', 'num', { def: '10' }), A('Sütun sayısı', '', 'num', { o: 1 }), A('Başlangıç', '', 'num', { o: 1 }), A('Artış', '', 'num', { o: 1 })],
  (r, c, s, st) => { r = Math.floor(N(r)); c = Math.floor(N(opt(c, 1))); s = N(opt(s, 1)); st = N(opt(st, 1)); if (r < 1 || c < 1 || r * c > 100000) throw VAL; const g = []; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(s + (i * c + j) * st); g.push(row); } return toArr(g); });
def('TAKE', 'AL', LK, 'Alanın baştan/sondan belirli satırlarını alır.', [A('Alan', '', 'range'), A('Satır', 'Negatif = sondan', 'num'), A('Sütun', '', 'num', { o: 1 })], (a, r, c) => { let g = grid2(a); r = Math.trunc(N(r)); g = r >= 0 ? g.slice(0, r) : g.slice(r); if (c != null) { c = Math.trunc(N(c)); g = g.map(row => c >= 0 ? row.slice(0, c) : row.slice(c)); } return toArr(g); });
def('CHOOSECOLS', 'SÜTUNSEÇ', LK, 'Alandan istenen sütunları alır.', [A('Alan', '', 'range'), A('Sütun numaraları', '', 'num', many)], (a, ...ks) => { const g = grid2(a); const idx = ks.map(k => { k = Math.trunc(N(k)); return k < 0 ? g[0].length + k : k - 1; }); return toArr(g.map(r => idx.map(i => { if (i < 0 || i >= r.length) throw VAL; return r[i]; }))); });

// Finans
const pmtF = (r, n, pv, fv, t) => { if (!r) return -(pv + fv) / n; const q = Math.pow(1 + r, n); return -(r * (fv + pv * q)) / ((1 + r * t) * (q - 1)); };
const fvF = (r, n, p, pv, t) => { if (!r) return -(pv + p * n); const q = Math.pow(1 + r, n); return -(pv * q + p * (1 + r * t) * (q - 1) / r); };
const ipmtF = (r, per, n, pv, fv, t) => { const p = pmtF(r, n, pv, fv, t); if (per === 1 && t === 1) return 0; let i = fvF(r, per - 1, p, pv, t) * r; if (t === 1) i /= 1 + r; return i; };
def('RATE', 'FAİZ_ORANI', F, 'Kredinin dönemlik faiz oranını bulur.', [A('Taksit sayısı', '', 'num'), A('Taksit (eksi)', 'Ör: -1000', 'num'), A('Kredi tutarı', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 }), A('Tahmin', '', 'num', { o: 1 })],
  (n, p, pv, fv, t, g) => { n = N(n); p = N(p); pv = N(pv); fv = N(opt(fv, 0)); t = N(opt(t, 0)); let r = N(opt(g, 0.1)); const f = r => !r ? pv + p * n + fv : pv * (1 + r) ** n + p * (1 + r * t) * ((1 + r) ** n - 1) / r + fv; for (let k = 0; k < 200; k++) { const y = f(r), d = (f(r + 1e-7) - y) / 1e-7; if (!d) break; const nr = r - y / d; if (Math.abs(nr - r) < 1e-12) return nr; r = nr; } if (Math.abs(f(r)) < 1e-6) return r; throw NUM; }, { fmt: '0.00%' });
def('NPER', 'TAKSİT_SAYISI', F, 'Borcun kaç taksitte biteceği.', [A('Faiz (dönemlik)', '', 'num'), A('Taksit (eksi)', '', 'num'), A('Kredi tutarı', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 })],
  (r, p, pv, fv, t) => { r = N(r); p = N(p); pv = N(pv); fv = N(opt(fv, 0)); t = N(opt(t, 0)); if (!r) return -(pv + fv) / p; const z = p * (1 + r * t) / r; return Math.log((z - fv) / (z + pv)) / Math.log(1 + r); });
def('IPMT', 'FAİZTUTARI', F, 'Belirli bir taksitteki faiz payı.', [A('Faiz (dönemlik)', '', 'num'), A('Kaçıncı taksit', '', 'num'), A('Taksit sayısı', '', 'num'), A('Kredi tutarı', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 })],
  (r, per, n, pv, fv, t) => ipmtF(N(r), N(per), N(n), N(pv), N(opt(fv, 0)), N(opt(t, 0))), { fmt: '#,##0.00' });
def('PPMT', 'ANA_PARA_ÖDEMESİ', F, 'Belirli bir taksitteki anapara payı.', [A('Faiz (dönemlik)', '', 'num'), A('Kaçıncı taksit', '', 'num'), A('Taksit sayısı', '', 'num'), A('Kredi tutarı', '', 'num'), A('Gelecek değer', '', 'num', { o: 1 }), A('Tür', '', 'num', { o: 1 })],
  (r, per, n, pv, fv, t) => { r = N(r); n = N(n); pv = N(pv); fv = N(opt(fv, 0)); t = N(opt(t, 0)); return pmtF(r, n, pv, fv, t) - ipmtF(r, N(per), n, pv, fv, t); }, { fmt: '#,##0.00' });
def('CUMIPMT', 'TOPÖDENENFAİZ', F, 'İki taksit arası ödenen toplam faiz.', [A('Faiz', '', 'num'), A('Taksit sayısı', '', 'num'), A('Kredi tutarı', '', 'num'), A('İlk taksit', '', 'num'), A('Son taksit', '', 'num'), A('Tür', '', 'num')],
  (r, n, pv, a, b, t) => { let s = 0; for (let k = N(a); k <= N(b); k++) s += ipmtF(N(r), k, N(n), N(pv), 0, N(t)); return s; }, { fmt: '#,##0.00' });
def('CUMPRINC', 'TOPANAPARA', F, 'İki taksit arası ödenen toplam anapara.', [A('Faiz', '', 'num'), A('Taksit sayısı', '', 'num'), A('Kredi tutarı', '', 'num'), A('İlk taksit', '', 'num'), A('Son taksit', '', 'num'), A('Tür', '', 'num')],
  (r, n, pv, a, b, t) => { r = N(r); n = N(n); pv = N(pv); t = N(t); let s = 0; for (let k = N(a); k <= N(b); k++) s += pmtF(r, n, pv, 0, t) - ipmtF(r, k, n, pv, 0, t); return s; }, { fmt: '#,##0.00' });
def('SLN', 'DA', F, 'Doğrusal amortisman (yıllık).', [A('Maliyet', '', 'num'), A('Hurda değeri', '', 'num'), A('Ömür (yıl)', '', 'num')], (c, s, l) => { l = N(l); if (!l) throw DIV0; return (N(c) - N(s)) / l; }, { fmt: '#,##0.00' });
def('DDB', 'ÇAFA', F, 'Azalan bakiyeler amortismanı.', [A('Maliyet', '', 'num'), A('Hurda değeri', '', 'num'), A('Ömür', '', 'num'), A('Dönem', '', 'num'), A('Oran', '', 'num', { o: 1 })], (c, s, l, p, f) => { c = N(c); s = N(s); l = N(l); p = N(p); f = N(opt(f, 2)); let acc = 0, d = 0; for (let i = 1; i <= p; i++) { d = Math.min((c - acc) * f / l, Math.max(0, c - s - acc)); acc += d; } return d; }, { fmt: '#,##0.00' });
def('EFFECT', 'ETKİN', F, 'Yıllık efektif faiz.', [A('Nominal faiz', '', 'num'), A('Yılda kaç dönem', '', 'num')], (n, p) => { p = Math.floor(N(p)); return Math.pow(1 + N(n) / p, p) - 1; }, { fmt: '0.00%' });
def('NOMINAL', 'NOMİNAL', F, 'Yıllık nominal faiz.', [A('Efektif faiz', '', 'num'), A('Yılda kaç dönem', '', 'num')], (e, p) => { p = Math.floor(N(p)); return p * (Math.pow(1 + N(e), 1 / p) - 1); }, { fmt: '0.00%' });
def('XNPV', 'ANBD', F, 'Tarihli nakit akışlarının net bugünkü değeri.', [A('Oran', '', 'num'), A('Nakit akışları', '', 'range'), A('Tarihler', '', 'range')], (r, v, d) => { r = N(r); const vs = nums([v]), ds = nums([d]); if (vs.length !== ds.length) throw NUM; return vs.reduce((s, x, i) => s + x / Math.pow(1 + r, (ds[i] - ds[0]) / 365), 0); });
def('XIRR', 'AİÇVERİMORANI', F, 'Tarihli nakit akışlarının iç verim oranı.', [A('Nakit akışları', '', 'range'), A('Tarihler', '', 'range'), A('Tahmin', '', 'num', { o: 1 })],
  (v, d, g) => { const vs = nums([v]), ds = nums([d]); let r = N(opt(g, 0.1)); const f = r => vs.reduce((s, x, i) => s + x / Math.pow(1 + r, (ds[i] - ds[0]) / 365), 0); for (let k = 0; k < 200; k++) { const y = f(r), dd = (f(r + 1e-7) - y) / 1e-7; if (!dd) break; const nr = r - y / dd; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw NUM; }, { fmt: '0.00%' });
def('MIRR', 'D_İÇ_VERİM_ORANI', F, 'Değiştirilmiş iç verim oranı.', [A('Nakit akışları', '', 'range'), A('Finansman oranı', '', 'num'), A('Yeniden yatırım oranı', '', 'num')],
  (v, fr, rr) => { const vs = nums([v]); fr = N(fr); rr = N(rr); const n = vs.length; let pos = 0, neg = 0; vs.forEach((x, i) => { if (x > 0) pos += x * Math.pow(1 + rr, n - 1 - i); else neg += x / Math.pow(1 + fr, i); }); if (!neg || !pos) throw DIV0; return Math.pow(-pos / neg, 1 / (n - 1)) - 1; }, { fmt: '0.00%' });

// ölçütü liste olan EĞERSAY/ETOPLA/EĞERORTALAMA her öğe için ayrı hesaplar (özet tablolar için)
['SUMIF', 'COUNTIF', 'AVERAGEIF'].forEach(n => { const f = FN[n].fn; FN[n].fn = (r, c, s2) => isRng(c) && c.h * c.w > 1 ? new Arr(c.grid().map(row => row.map(x => { try { return f(r, x, s2); } catch (e) { return isErr(e) ? e : VAL; } }))) : f(r, c, s2); });
['MAXIFS', 'MINIFS'].forEach(n => { const f = FN[n].fn; FN[n].fn = (v, g, c, ...rest) => isRng(c) && c.h * c.w > 1 ? new Arr(c.grid().map(row => row.map(x => { try { return f(v, g, x, ...rest); } catch (e) { return isErr(e) ? e : VAL; } }))) : f(v, g, c, ...rest); });

// ---------- Sayı biçimleri ----------
function fmtGeneral(v) {
  if (!isFinite(v)) return '#NUM!';
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  const a = Math.abs(v);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'e').replace('.', ',').toUpperCase();
  let s = String(parseFloat(v.toPrecision(10)));
  if (/e/i.test(s)) s = v.toExponential(5).toUpperCase();
  return s.replace('.', ',');
}
function splitSections(code) { const out = []; let cur = '', q = false; for (let i = 0; i < code.length; i++) { const ch = code[i]; if (ch === '"') q = !q; if (ch === '\\' && !q) { cur += ch + (code[i + 1] || ''); i++; continue; } if (ch === ';' && !q) { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out; }
function stripCode(s) { return s.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, ''); }
function isDateCode(code) { const s = stripCode(splitSections(code)[0]); return /[dmyhs]/i.test(s) && !/[0#?]/.test(s.replace(/s\.0+/i, '')) || /[dmyh].*s\.0/i.test(s); }
function pad(n, w) { return String(n).padStart(w, '0'); }

function formatDate(v, sec) {
  const totalSec = Math.round(v * 86400); const day = Math.floor(totalSec / 86400); const d = serialToDate(day);
  const Y = d.getUTCFullYear(), Mo = d.getUTCMonth(), D = d.getUTCDate(), W = d.getUTCDay();
  let secs = totalSec - day * 86400; let H = Math.floor(secs / 3600), Mi = Math.floor((secs % 3600) / 60), Se = secs % 60;
  const ampm = /AM\/PM|A\/P/i.test(sec); let hh = H; if (ampm) { hh = H % 12 || 12; }
  const elapsedH = /\[h+\]/i.test(sec);
  let out = ''; let i = 0; const s = sec; let lastWasH = false;
  const toks = [];
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"') { const j = s.indexOf('"', i + 1); toks.push({ lit: s.slice(i + 1, j < 0 ? s.length : j) }); i = j < 0 ? s.length : j + 1; continue; }
    if (ch === '\\') { toks.push({ lit: s[i + 1] || '' }); i += 2; continue; }
    if (ch === '[') { const j = s.indexOf(']', i); const inner = s.slice(i + 1, j); if (/^h+$/i.test(inner)) toks.push({ k: 'H', n: inner.length, el: 1 }); else if (/^m+$/i.test(inner)) toks.push({ k: 'EM', n: inner.length }); i = j + 1; continue; }
    if (/^AM\/PM/i.test(s.slice(i))) { toks.push({ lit: H < 12 ? 'ÖÖ' : 'ÖS' }); i += 5; continue; }
    const m = /^(y+|m+|d+|h+|s+)/i.exec(s.slice(i));
    if (m) { toks.push({ k: m[0][0].toLowerCase(), n: m[0].length }); i += m[0].length; continue; }
    if (ch === '0' && toks.length && toks[toks.length - 1].k === 's') { const m2 = /^0+/.exec(s.slice(i)); toks.push({ k: 'frac', n: m2[0].length }); i += m2[0].length; continue; }
    if (ch === '.' && /^\.0/.test(s.slice(i)) && toks.length && toks[toks.length - 1].k === 's') { toks.push({ lit: ',' }); i++; continue; }
    toks.push({ lit: ch }); i++;
  }
  // m -> dakika mı?
  for (let k = 0; k < toks.length; k++) {
    if (toks[k].k !== 'm' || toks[k].n > 2) continue;
    let p = k - 1; while (p >= 0 && toks[p].lit !== undefined) p--;
    let q = k + 1; while (q < toks.length && toks[q].lit !== undefined) q++;
    if ((p >= 0 && (toks[p].k === 'h' || toks[p].k === 'H')) || (q < toks.length && toks[q].k === 's')) toks[k].k = 'min';
  }
  for (const t of toks) {
    if (t.lit !== undefined) { out += t.lit; continue; }
    switch (t.k) {
      case 'y': out += t.n <= 2 ? pad(Y % 100, 2) : Y; break;
      case 'm': out += t.n === 1 ? (Mo + 1) : t.n === 2 ? pad(Mo + 1, 2) : t.n === 3 ? AYLAR[Mo].slice(0, 3) : t.n === 5 ? AYLAR[Mo][0] : AYLAR[Mo]; break;
      case 'd': out += t.n === 1 ? D : t.n === 2 ? pad(D, 2) : t.n === 3 ? GUNLER[W].slice(0, 3) : GUNLER[W]; break;
      case 'h': out += t.n === 1 ? hh : pad(hh, 2); break;
      case 'H': out += elapsedH ? pad(Math.floor(totalSec / 3600), t.n) : pad(H, t.n); break;
      case 'EM': out += pad(Math.floor(totalSec / 60), t.n); break;
      case 'min': out += t.n === 1 ? Mi : pad(Mi, 2); break;
      case 's': out += t.n === 1 ? Se : pad(Se, 2); break;
      case 'frac': out += pad(Math.round((v * 86400 % 1) * Math.pow(10, t.n)), t.n); break;
    }
  }
  return out;
}

function groupInt(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function formatNumber(v, sec, hadNeg) {
  // metin parçaları ve sayı deseni
  let pre = '', post = '', pat = '', i = 0, inNum = false, done = false;
  let pct = 0;
  const s = sec.replace(/\[[^\]]*\]/g, m => { const c = /^\[\$([^-\]]*)/.exec(m); return c ? '"' + c[1] + '"' : ''; });
  while (i < s.length) {
    const ch = s[i];
    let lit = null;
    if (ch === '"') { const j = s.indexOf('"', i + 1); lit = s.slice(i + 1, j < 0 ? s.length : j); i = j < 0 ? s.length : j + 1; }
    else if (ch === '\\') { lit = s[i + 1] || ''; i += 2; }
    else if (ch === '_') { lit = ' '; i += 2; }
    else if (ch === '*') { i += 2; continue; }
    else if (!done && /[0#?.,]/.test(ch) || (inNum && !done && /[eE+\-]/.test(ch) && /E[+-]/i.test(s.slice(i, i + 2)) ) || (inNum && /[0#?+\-]/.test(ch) && /E/i.test(pat))) { pat += ch; inNum = true; i++; continue; }
    else { if (ch === '%') pct++; lit = ch; i++; }
    if (inNum) done = true;
    if (done || inNum) post += lit; else pre += lit;
  }
  let x = Math.abs(v) * Math.pow(100, pct);
  if (!pat) return (v < 0 && !hadNeg ? '-' : '') + pre + post;
  const sci = /E[+-]/i.exec(pat);
  let numStr;
  if (sci) {
    const mant = pat.slice(0, sci.index); const dec = (mant.split('.')[1] || '').length;
    const e = x === 0 ? 0 : Math.floor(Math.log10(x)); const m = x / Math.pow(10, e);
    numStr = m.toFixed(dec).replace('.', ',') + 'E' + (e < 0 ? '-' : '+') + pad(Math.abs(e), (pat.slice(sci.index + 2).length) || 2);
  } else {
    let p = pat; let scale = 0; while (p.endsWith(',')) { scale++; p = p.slice(0, -1); } x /= Math.pow(1000, scale);
    const [ip, fp = ''] = p.split('.');
    const thousands = ip.includes(',');
    const minInt = (ip.replace(/,/g, '').match(/0/g) || []).length;
    const minDec = (fp.match(/0/g) || []).length, maxDec = (fp.match(/[0#?]/g) || []).length;
    let fixed = round(x, maxDec).toFixed(maxDec);
    let [a, b = ''] = fixed.split('.');
    while (b.length > minDec && b.endsWith('0')) b = b.slice(0, -1);
    if (a === '0' && minInt === 0) a = '';
    a = a.padStart(minInt, '0');
    if (thousands) a = groupInt(a);
    numStr = a + (b.length || p.includes('.') && minDec ? ',' + b : (p.endsWith('.') ? ',' : ''));
  }
  const neg = v < 0 && !hadNeg && Math.abs(v) * Math.pow(100, pct) >= Math.pow(10, -((pat.split('.')[1] || '').length)) / 2;
  return (neg ? '-' : '') + pre + numStr + post;
}

function formatValue(v, code) {
  if (v == null) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'DOĞRU' : 'YANLIŞ';
  if (typeof v === 'string') { if (code && code.includes('@')) { const secs = splitSections(code); const sec = secs[3] ?? (secs.length === 1 ? secs[0] : null); if (sec) return sec.replace(/"([^"]*)"/g, '$1').replace('@', v); } return v; }
  if (!code || /^(General|Genel)$/i.test(code)) return fmtGeneral(v);
  if (code === '@') return fmtGeneral(v);
  const secs = splitSections(code);
  let sec = secs[0], hadNeg = false;
  if (v < 0 && secs[1] != null && secs[1] !== '') { sec = secs[1]; hadNeg = true; }
  else if (v === 0 && secs[2] != null && secs[2] !== '') sec = secs[2];
  if (/^(General|Genel)$/i.test(sec.replace(/\[[^\]]*\]/g, ''))) return fmtGeneral(hadNeg ? -v : v);
  try {
    if (isDateCode(sec)) return formatDate(hadNeg ? -v : v, sec);
    return formatNumber(v, sec, hadNeg);
  } catch (e) { return fmtGeneral(v); }
}
// Biçimdeki renk etiketi ([Red], [Kırmızı])
const CODE_COLORS = { red: '#d32f2f', kırmızı: '#d32f2f', blue: '#1565c0', mavi: '#1565c0', green: '#2e7d32', yeşil: '#2e7d32', black: '#000', siyah: '#000', magenta: '#c2185b', cyan: '#00838f', yellow: '#f9a825', white: '#fff' };
function formatColor(v, code) {
  if (typeof v !== 'number' || !code) return null;
  const secs = splitSections(code); let sec = secs[0];
  if (v < 0 && secs[1]) sec = secs[1]; else if (v === 0 && secs[2]) sec = secs[2];
  const m = /\[([^\]$]+)\]/.exec(sec); return m ? CODE_COLORS[m[1].toLocaleLowerCase('tr-TR')] || null : null;
}
// METNEÇEVİR içindeki Türkçe biçim kodlarını İngilizceye çevir
function trFormatCode(f) {
  let s = f;
  if (/[gGaA]/.test(stripCode(s)) && !/AM\/PM/i.test(s)) s = s.replace(/("[^"]*")|g/gi, (m, q) => q || 'd').replace(/("[^"]*")|a/gi, (m, q) => q || 'm');
  const pat = (/[0#?][0#?.,]*/.exec(stripCode(s)) || [''])[0];
  const lc = pat.lastIndexOf(','), ld = pat.lastIndexOf('.');
  let isTr = false;
  if (lc > ld) isTr = ld >= 0 || pat.length - lc - 1 !== 3;
  else if (ld > lc) isTr = lc < 0 && pat.length - ld - 1 === 3;
  if (isTr) s = s.replace(/("[^"]*")|[.,]/g, (m, q) => q || (m === '.' ? ',' : '.'));
  return s;
}
function isDateFormat(code) { return !!code && !/^(General|Genel|@)$/i.test(code) && isDateCode(code); }

G.XF = {
  FErr, E, isErr, ERRS, colName, colIndex, addr, parseAddr, quoteSheet, MAXR, MAXC,
  tokenize, serialize, parse, evaluate, toDisplay, toStore, shiftFormula, adjustFormula, renameSheetInFormula,
  FN, lookupFn, Range, Arr, isRng,
  formatValue, formatColor, fmtGeneral, isDateFormat, trFormatCode,
  parseNumLoose, parseDateLoose, parseTimeLoose, toNum, toStr, cmp,
  serialToDate, dateToSerial, ymdToSerial, nowSerial, AYLAR, GUNLER, crit
};
})(typeof window !== 'undefined' ? window : globalThis);
