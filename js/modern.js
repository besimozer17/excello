// Excello — modern görünüm: aynı hesap motoru ve dosya biçimleri, Excel'e benzemeyen arayüz
(function (G) {
'use strict';
const X = G.XF, M = G.XM, U = G.UI, IO = G.XIO, C = G.XC;
const { h, icon, $, $$ } = U;
const LS = { get(k, d) { try { const v = localStorage.getItem('excello.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } }, set(k, v) { try { localStorage.setItem('excello.' + k, JSON.stringify(v)); } catch (e) { } } };
const isMobile = () => matchMedia('(pointer: coarse)').matches;
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

// klasik görünümle aynı yerel kayıt: iki görünüm aynı çalışmayı açar
const DB = {
  db: null,
  open() { if (this.db) return Promise.resolve(this.db); return new Promise((res, rej) => { let r; try { r = indexedDB.open('excello', 1); } catch (e) { return rej(e); } r.onupgradeneeded = () => r.result.createObjectStore('docs', { keyPath: 'id' }); r.onsuccess = () => res(this.db = r.result); r.onerror = () => rej(r.error); }); },
  async tx(mode, fn) { const work = (async () => { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('docs', mode); const st = t.objectStore('docs'); const rq = fn(st); t.oncomplete = () => res(rq && rq.result); t.onerror = () => rej(t.error); }); })(); return Promise.race([work, new Promise((_, rej) => setTimeout(() => rej(new Error('zaman aşımı')), 3000))]); },
  put(d) { return this.tx('readwrite', s => s.put(d)).catch(e => { try { localStorage.setItem('excello.doc.' + d.id, JSON.stringify(d)); } catch (e2) { throw e; } }); },
  get(id) { return this.tx('readonly', s => s.get(id)).then(r => r || JSON.parse(localStorage.getItem('excello.doc.' + id) || 'null')).catch(() => JSON.parse(localStorage.getItem('excello.doc.' + id) || 'null')); },
  all() { return this.tx('readonly', s => s.getAll()); },
  del(id) { try { localStorage.removeItem('excello.doc.' + id); } catch (e) { } return this.tx('readwrite', s => s.delete(id)); }
};

// sütun türleri
const KINDS = [['text', 'Metin', 'Aa'], ['num', 'Sayı', '#'], ['money', 'Para', '₺'], ['pct', 'Yüzde', '%'], ['date', 'Tarih', '◷']];
const KIND_NF = { text: null, num: '#,##0.##', money: '#,##0.00 ₺', pct: '0%', date: 'dd.mm.yyyy' };
const AGGS = [['sum', 'Toplam'], ['avg', 'Ortalama'], ['max', 'En büyük'], ['min', 'En küçük'], ['count', 'Sayı'], ['distinct', 'Farklı']];
const OPS = [['=', 'eşittir'], ['<>', 'eşit değil'], ['>', 'büyüktür'], ['<', 'küçüktür'], ['>=', 'en az'], ['<=', 'en çok'], ['contains', 'içerir'], ['empty', 'boş'], ['filled', 'dolu']];
const TONES = [['Yeşil', '#dcfce7', '#15803d'], ['Kırmızı', '#fee2e2', '#b91c1c'], ['Turuncu', '#ffedd5', '#c2410c'], ['Sarı', '#fef9c3', '#a16207'], ['Mavi', '#dbeafe', '#1d4ed8'], ['Mor', '#ede9fe', '#6d28d9'], ['Pembe', '#fce7f3', '#be185d'], ['Turkuaz', '#ccfbf1', '#0f766e'], ['Gri', '#f4f4f5', '#3f3f46']];
const CF_OPS = [['gt', 'büyükse'], ['lt', 'küçükse'], ['ge', 'en az'], ['le', 'en çok'], ['between', 'arasındaysa'], ['eq', 'eşitse'], ['contains', 'içeriyorsa'], ['above', 'ortalamanın üstündeyse'], ['below', 'ortalamanın altındaysa'], ['top', 'en büyük N içindeyse'], ['bottom', 'en küçük N içindeyse'], ['dup', 'tekrar ediyorsa'], ['blank', 'boşsa']];
const isTotalText = v => typeof v === 'string' && /^(genel\s+)?(toplam|total|ara toplam)/i.test(v.trim());
const lc = s => String(s == null ? '' : s).toLocaleLowerCase('tr-TR');

class Modern {
  constructor() {
    this.wb = new M.Workbook();
    this.sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    this.view = LS.get('m.view', 'list');
    this.vs = {}; // sayfa başına görünüm durumu (arama, filtre, sıralama, toplam türleri)
    this.applyTheme(LS.get('theme', 'auto'));
    this.wiz = new G.XW.Wizards(this);
    this.build();
    this.attach(this.wb);
    this.boot();
  }
  async boot() {
    try {
      const id = LS.get('cur', null); const doc = id && await DB.get(id);
      if (doc) { const wb = M.Workbook.fromJSON(doc.data); wb.format = doc.format || 'xlsx'; this.docId = id; this.attach(wb); }
      else { this.docId = this.newId(); this.sample(); }
    } catch (e) { this.docId = this.newId(); }
    const openUrl = new URLSearchParams(location.search).get('open'); if (openUrl) this.openFromUrl(openUrl);
    if ('serviceWorker' in navigator && /^https?:/.test(location.protocol) && location.hostname !== 'localhost') { const had = !!navigator.serviceWorker.controller; navigator.serviceWorker.register('sw.js').catch(() => { }); navigator.serviceWorker.addEventListener('controllerchange', () => { if (had && !this._reloaded) { this._reloaded = true; this.saveLocal().then(() => location.reload()); } }); }
    addEventListener('beforeunload', () => this.saveLocal());
    document.addEventListener('visibilitychange', () => { if (document.hidden) this.saveLocal(); });
  }
  newId() { return 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
  applyTheme(t) {
    this.theme = t; LS.set('theme', t);
    const dark = t === 'dark' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.classList.toggle('dark', dark);
    if (this.T) this.render();
  }
  attach(wb) {
    this.wb = wb; wb.listeners = []; wb.on(() => this.onChange());
    if (wb.active >= wb.sheets.length) wb.active = 0;
    this.render();
  }
  onChange() { this.render(); this.scheduleSave(); }
  get vsCur() { const k = this.wb.active; return this.vs[k] || (this.vs[k] = { q: '', filters: [], sort: null, agg: {} }); }

  // ---------- iskelet ----------
  build() {
    const root = $('#m');
    this.side = h('aside', { class: 'm-side' });
    this.head = h('header', { class: 'm-head' });
    this.tools = h('div', { class: 'm-tools' });
    this.body = h('div', { class: 'm-body' });
    this.main = h('main', { class: 'm-main' }, this.head, this.tools, this.body);
    this.scrim = h('div', { class: 'm-scrim', onclick: () => root.classList.remove('side-open') });
    root.append(this.side, this.scrim, this.main);
    addEventListener('keydown', e => {
      if (e.target.closest && e.target.closest('input,textarea,select')) return;
      const g = this.view === 'list' && !U.topDialog() ? this.selRect() : null;
      if (g) {
        const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
        if (mv) { e.preventDefault(); const k = e.shiftKey ? 'b' : 'a'; const p = this.msel[k]; const n = { ri: Math.max(0, Math.min(this._vr.length - 1, p.ri + mv[0])), ci: Math.max(0, Math.min(this._vc.length - 1, p.ci + mv[1])) }; if (e.shiftKey) this.msel.b = n; else this.msel = { si: this.msel.si, a: n, b: { ...n } }; this.paintSel(); const td = this.body.querySelector(`.m-t td[data-ri="${n.ri}"][data-ci="${n.ci}"]`); td && td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); return; }
        if (e.key === 'Escape') { this.msel = null; this.paintSel(); return; }
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); this.selClear(); return; }
        if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); this.editAt(this.msel.a); return; }
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); this.editAt(this.msel.a, e.key); return; }
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); this.wb.undo(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); this.wb.redo(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); this.saveAs('xlsx'); }
    });
    document.addEventListener('copy', e => { if (e.target.closest && e.target.closest('input,textarea')) return; const t = this.view === 'list' && this.selRect() ? this.selTSV() : ''; if (t) { e.preventDefault(); e.clipboardData.setData('text/plain', t); U.toast('✓ Kopyalandı'); } });
    document.addEventListener('paste', e => { if (e.target.closest && e.target.closest('input,textarea')) return; if (this.view === 'list' && this.selRect() && !U.topDialog()) { e.preventDefault(); this.selPaste(e.clipboardData.getData('text/plain')); } });
  }

  // ---------- tablo bilgisi ----------
  table(si) {
    const wb = this.wb; si = si == null ? wb.active : si; const d = wb.dims(si);
    const T = { si, hr: -1, r0: 0, r2: -1, cols: [], totals: [] };
    if (!d.rows) return T;
    let head = true, any = false;
    for (let c = 0; c < d.cols; c++) { const v = wb.value(si, 0, c); if (v == null || v === '') continue; any = true; if (typeof v !== 'string') head = false; }
    head = head && any;
    T.hr = head ? 0 : -1; T.r0 = head ? 1 : 0;
    // en alttaki "Toplam" satırları veri değil
    const rowTotal = r => { for (let c = 0; c < d.cols; c++) { const v = wb.value(si, r, c), cl = wb.cell(si, r, c); if (isTotalText(v)) return true; if (cl && cl.f && /^(SUM|SUBTOTAL)\(/.test(cl.f)) { const m = /\(\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)\)/.exec(cl.f); if (m && +m[2] - 1 < r) return true; } } return false; };
    const rowEmpty = r => { for (let c = 0; c < d.cols; c++) { const v = wb.value(si, r, c); if (v != null && v !== '') return false; } return true; };
    let r2 = d.rows - 1;
    while (r2 >= T.r0 && (rowEmpty(r2) || rowTotal(r2))) { if (!rowEmpty(r2)) T.totals.unshift(r2); r2--; }
    T.r2 = r2;
    for (let c = 0; c < d.cols; c++) {
      let filled = 0, f = 0, num = 0, txt = 0;
      for (let r = T.r0; r <= T.r2; r++) { const v = wb.value(si, r, c); if (v == null || v === '') continue; filled++; if (typeof v === 'number') num++; else txt++; const cl = wb.cell(si, r, c); if (cl && cl.f != null) f++; else if (!cl && wb.spillAt(si, r, c)) f++; }
      const name = head ? wb.display(si, 0, c) : '';
      if (!filled && !name) continue;
      const nf = wb.numFmt(si, T.r0, c) || '';
      const kind = X.isDateFormat(nf) ? 'date' : /%/.test(nf) ? 'pct' : /[₺$€]/.test(nf) ? 'money' : num > txt ? 'num' : 'text';
      T.cols.push({ c, name: name || 'Sütun ' + X.colName(c), kind, nf, computed: filled > 0 && f >= filled / 2, filled });
    }
    return T;
  }
  rowsOf(T) {
    // arama + filtre + sıralama sonrası görünen satırlar
    const wb = this.wb, vs = this.vs[T.si] || { q: '', filters: [], sort: null };
    let rows = []; for (let r = T.r0; r <= T.r2; r++) rows.push(r);
    const q = lc(vs.q).trim();
    if (q) rows = rows.filter(r => T.cols.some(c => lc(wb.display(T.si, r, c.c)).includes(q)));
    for (const f of vs.filters) rows = rows.filter(r => this.test(T.si, r, f));
    if (vs.sort) { const { c, dir } = vs.sort; rows.sort((a, b) => { const x = wb.value(T.si, a, c), y = wb.value(T.si, b, c); if (x == null || x === '') return 1; if (y == null || y === '') return -1; return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'tr')) * dir; }); }
    return rows;
  }
  test(si, r, f) {
    const v = this.wb.value(si, r, f.c), d = this.wb.display(si, r, f.c);
    if (f.op === 'empty') return v == null || v === ''; if (f.op === 'filled') return !(v == null || v === '');
    const n = X.parseNumLoose(String(f.val)); const dt = X.parseDateLoose(String(f.val));
    const cmp = typeof v === 'number' ? (n != null ? n : dt ? dt.v : null) : null;
    if (f.op === 'contains') return lc(d).includes(lc(f.val));
    if (cmp != null) { if (f.op === '=') return v === cmp; if (f.op === '<>') return v !== cmp; if (f.op === '>') return v > cmp; if (f.op === '<') return v < cmp; if (f.op === '>=') return v >= cmp; if (f.op === '<=') return v <= cmp; }
    const a = lc(d), b = lc(f.val);
    if (f.op === '=') return a === b; if (f.op === '<>') return a !== b;
    const k = a.localeCompare(b, 'tr'); return f.op === '>' ? k > 0 : f.op === '<' ? k < 0 : f.op === '>=' ? k >= 0 : k <= 0;
  }
  agg(T, col, rows, kind) {
    const wb = this.wb; const vals = rows.map(r => wb.value(T.si, r, col.c)).filter(v => v != null && v !== '');
    const nums = vals.filter(v => typeof v === 'number');
    switch (kind) {
      case 'sum': return nums.reduce((a, b) => a + b, 0);
      case 'avg': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
      case 'max': return nums.length ? Math.max(...nums) : null;
      case 'min': return nums.length ? Math.min(...nums) : null;
      case 'count': return vals.length;
      case 'distinct': return new Set(vals.map(v => lc(v))).size;
    }
  }

  // ---------- çizim ----------
  render() {
    if (this._rq) return; this._rq = requestAnimationFrame(() => { this._rq = 0; this.renderNow(); });
  }
  renderNow() {
    const T = this.T = this.table();
    if (this.msel && (this.msel.si !== this.wb.active || this.view !== 'list')) this.msel = null;
    this.renderSide(); this.renderHead(T); this.renderTools(T);
    const sc = this.body.scrollTop;
    this.body.innerHTML = '';
    if (!T.cols.length) this.body.append(this.emptyState());
    else if (this.view === 'cards') this.body.append(this.viewCards(T));
    else if (this.view === 'summary') this.body.append(this.viewSummary(T));
    else if (this.view === 'chart') this.body.append(this.viewChart(T));
    else this.body.append(this.viewList(T));
    this.body.scrollTop = sc;
  }
  renderSide() {
    const wb = this.wb, s = this.side; s.innerHTML = '';
    s.append(h('div', { class: 'm-brand' }, h('span', { class: 'logo' , text: 'E' }), h('div', null, h('b', { text: 'Excello' }), h('small', { text: 'Tablolarım' }))));
    s.append(h('input', { class: 'm-docname', value: wb.name || 'Yeni Çalışma', title: 'Çalışmanın adı', onchange: e => { wb.name = e.target.value.trim() || 'Çalışma'; this.scheduleSave(); } }));
    const tl = h('nav', { class: 'm-tables' }, h('div', { class: 'm-sec', text: 'Tablolar' }));
    wb.sheets.forEach((sh, i) => {
      const T = i === wb.active ? this.T : null; const n = T ? Math.max(0, T.r2 - T.r0 + 1) : null;
      tl.append(h('button', { class: 'm-tab' + (i === wb.active ? ' on' : ''), onclick: () => { wb.active = i; $('#m').classList.remove('side-open'); this.render(); }, oncontextmenu: e => { e.preventDefault(); this.sheetMenu(e.currentTarget, i); } },
        h('span', { class: 'm-tab-ic', html: icon('table') }), h('span', { class: 'm-tab-t', text: sh.name }), n != null ? h('small', { text: String(n) }) : null,
        h('span', { class: 'm-tab-more', html: icon('dots'), onclick: e => { e.stopPropagation(); this.sheetMenu(e.currentTarget, i); } })));
    });
    tl.append(h('button', { class: 'm-tab add', onclick: () => this.newTable() }, h('span', { class: 'm-tab-ic', html: icon('plus') }), h('span', { class: 'm-tab-t', text: 'Yeni tablo' })));
    s.append(tl);
    const act = (ic, t, fn) => h('button', { class: 'm-act', onclick: () => { $('#m').classList.remove('side-open'); fn(); } }, h('span', { html: icon(ic) }), h('span', { text: t }));
    s.append(h('div', { class: 'm-sec', text: 'Dosya' }),
      act('file', 'Yeni çalışma', () => this.newDoc()),
      act('folder', 'Dosya aç (Excel, CSV…)', () => this.openFile()),
      act('save', 'Excel olarak kaydet', () => this.saveAs('xlsx')),
      act('share', 'Başka biçimde kaydet', () => this.saveMenu()),
      act('layers', 'Son çalışmalar', () => this.recent()),
      act('table', 'Örnek tablo', () => this.sample()));
    s.append(h('div', { class: 'spacer' }),
      h('div', { class: 'm-foot' },
        h('button', { class: 'm-act', onclick: () => this.applyTheme(document.documentElement.classList.contains('dark') ? 'light' : 'dark') }, h('span', { html: icon('moon') }), h('span', { text: 'Koyu / açık' }))));
  }
  renderHead(T) {
    const wb = this.wb, sh = wb.sheet; const hd = this.head; hd.innerHTML = '';
    const n = Math.max(0, T.r2 - T.r0 + 1);
    hd.append(h('button', { class: 'ic m-burger', html: icon('menu'), 'aria-label': 'Menü', onclick: () => $('#m').classList.toggle('side-open') }),
      h('div', { class: 'm-titlebox' },
        h('input', { class: 'm-title', value: sh.name, 'aria-label': 'Tablo adı', onchange: e => { const v = e.target.value.trim(); if (v && v !== sh.name) { if (wb.sheetIndex(v) >= 0) { U.toast('Bu adda bir tablo var'); e.target.value = sh.name; return; } wb.renameSheet(wb.active, v); } } }),
        h('small', { class: 'm-sub', text: n + ' kayıt · ' + T.cols.length + ' sütun' + (this._saved ? ' · ' + this._saved : '') })),
      h('button', { class: 'ic m-tmore', html: icon('dots'), title: 'Tablo işlemleri', 'aria-label': 'Tablo işlemleri', onclick: e => this.sheetMenu(e.currentTarget, wb.active) }),
      h('div', { class: 'spacer' }),
      h('button', { class: 'ic', html: icon('undo'), title: 'Geri al', disabled: !wb.undoS.length, onclick: () => wb.undo() }),
      h('button', { class: 'ic', html: icon('redo'), title: 'Yinele', disabled: !wb.redoS.length, onclick: () => wb.redo() }),
      h('button', { class: 'btn primary m-save', html: icon('save') + '<span>Excel olarak kaydet</span>', onclick: () => this.saveAs('xlsx') }));
  }
  renderTools(T) {
    const t = this.tools; t.innerHTML = ''; const vs = this.vsCur;
    const views = [['list', 'Liste', 'list'], ['cards', 'Kartlar', 'layers'], ['summary', 'Özet', 'sigma'], ['chart', 'Grafik', 'chart']];
    t.append(h('div', { class: 'm-views' }, ...views.map(([k, tx, ic]) => h('button', { class: 'm-view' + (this.view === k ? ' on' : ''), onclick: () => { this.view = k; LS.set('m.view', k); this.render(); } }, h('span', { html: icon(ic) }), h('span', { text: tx })))));
    const q = h('input', { class: 'm-search', placeholder: 'Ara…', value: vs.q, oninput: () => { vs.q = q.value; clearTimeout(this._qt); this._qt = setTimeout(() => { this.render(); setTimeout(() => { const el = $('.m-search'); if (el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); } }, 20); }, 250); } });
    const chips = h('div', { class: 'm-chips' });
    vs.filters.forEach((f, i) => { const col = T.cols.find(c => c.c === f.c); chips.append(h('span', { class: 'm-chip' }, h('span', { text: (col ? col.name : '?') + ' ' + OPS.find(o => o[0] === f.op)[1] + (f.op === 'empty' || f.op === 'filled' ? '' : ' ' + f.val) }), h('button', { html: icon('x'), 'aria-label': 'Kaldır', onclick: () => { vs.filters.splice(i, 1); this.render(); } }))); });
    if (vs.sort) { const col = T.cols.find(c => c.c === vs.sort.c); chips.append(h('span', { class: 'm-chip sort' }, h('span', { text: (col ? col.name : '?') + (vs.sort.dir > 0 ? ' ↑' : ' ↓') }), h('button', { html: icon('x'), onclick: () => { vs.sort = null; this.render(); } }))); }
    t.append(h('div', { class: 'm-searchbox' }, h('span', { html: icon('search') }), q),
      h('div', { class: 'm-actions' },
        h('button', { class: 'btn m-calc', html: icon('wand') + '<span>Hesapla</span>', onclick: () => { this.focusCol(T, T.cols[0] && T.cols[0].c); this.wiz.openRecipes(); } }),
        h('button', { class: 'btn', html: icon('pointer') + '<span>Bul ve getir</span>', onclick: () => this.lookupTool() }),
        h('button', { class: 'btn', html: icon('plus') + '<span>Sütun</span>', onclick: () => this.addColumn(T) }),
        h('button', { class: 'btn', html: icon('filter') + '<span>Filtre</span>', onclick: e => this.filterDialog(T) }),
        h('button', { class: 'btn', html: icon('fill') + '<span>Renk</span>', onclick: () => this.colorDialog(T) })),
      chips);
  }
  emptyState() {
    return h('div', { class: 'm-empty' }, h('div', { class: 'm-empty-ic', html: icon('table') }), h('h2', { text: 'Bu tablo boş' }), h('p', { text: 'Sütunları ekleyerek başlayın ya da bir Excel dosyası açın.' }),
      h('div', { class: 'btnrow' }, h('button', { class: 'btn primary', html: icon('plus') + '<span>Sütun ekle</span>', onclick: () => this.addColumn(this.T) }), h('button', { class: 'btn', html: icon('folder') + '<span>Dosya aç</span>', onclick: () => this.openFile() })));
  }
  badge(col) { const k = KINDS.find(x => x[0] === col.kind); return h('span', { class: 'm-badge k-' + col.kind + (col.computed ? ' m-fx' : ''), text: col.computed ? 'ƒ' : k[2], title: col.computed ? 'Hesaplanan sütun' : k[1] }); }

  // ---------- Liste ----------
  viewList(T) {
    const wb = this.wb, si = T.si, vs = this.vsCur; const rows = this.rowsOf(T);
    const tbl = h('table', { class: 'm-t' });
    const hr = h('tr', null, h('th', { class: 'm-rn' }));
    for (const col of T.cols) hr.append(h('th', { class: 'k-' + col.kind + (col.computed ? ' m-fx' : ''), onclick: e => this.colMenu(e.currentTarget, T, col), oncontextmenu: e => { e.preventDefault(); this.colMenu(e.currentTarget, T, col); } }, h('div', { class: 'm-th' }, this.badge(col), h('span', { class: 'm-th-t', text: col.name }), vs.sort && vs.sort.c === col.c ? h('span', { class: 'm-sortm', text: vs.sort.dir > 0 ? '↑' : '↓' }) : null, h('span', { class: 'm-th-dd', html: icon('down') }))));
    const tb = h('tbody');
    this._vr = rows; this._vc = T.cols; this._vT = T;
    rows.forEach((r, i) => {
      const tr = h('tr', null, h('td', { class: 'm-rn', text: String(i + 1), title: 'Kaydı aç', onclick: () => this.rowForm(T, r), oncontextmenu: e => { e.preventDefault(); this.msel = { si, a: { ri: i, ci: 0 }, b: { ri: i, ci: T.cols.length - 1 } }; this.paintSel(); this.selMenu(e); } }));
      T.cols.forEach((col, ci) => {
        const v = wb.value(si, r, col.c); const st = wb.cfStyle ? wb.cfStyle(si, r, col.c) : null;
        const td = h('td', { class: 'k-' + col.kind + (col.computed ? ' m-fx' : '') + (X.isErr(v) ? ' err' : '') + (typeof v === 'number' ? ' num' : ''), text: wb.display(si, r, col.c), 'data-r': r, 'data-c': col.c, 'data-ri': i, 'data-ci': ci });
        this.paintCell(td, si, r, col.c, st);
        tr.append(td);
      });
      tb.append(tr);
    });
    this.bindSel(tb, T);
    const tf = h('tr', null, h('td', { class: 'm-rn' }));
    for (const col of T.cols) {
      const kind = vs.agg[col.c] || (col.kind === 'text' ? 'count' : col.kind === 'date' ? 'count' : col.kind === 'pct' ? 'avg' : 'sum');
      const v = this.agg(T, col, rows, kind); const nm = AGGS.find(a => a[0] === kind)[1];
      const txt = v == null ? '—' : kind === 'count' || kind === 'distinct' ? String(v) : X.formatValue(v, col.nf || null);
      tf.append(h('td', { class: 'm-agg', title: 'Dokunarak hesap türünü değiştirin', onclick: e => U.menu(e.currentTarget, AGGS.map(([k, t]) => ({ text: t, on: k === kind, onClick: () => { vs.agg[col.c] = k; this.render(); } }))) }, h('small', { text: nm }), h('b', { text: txt })));
    }
    tbl.append(h('thead', null, hr), tb, h('tfoot', null, tf));
    const wrap = h('div', { class: 'm-card m-listwrap' }, h('div', { class: 'm-scrollx' }, tbl));
    if (!rows.length) wrap.append(h('div', { class: 'm-none', text: vs.q || vs.filters.length ? 'Aramaya / filtreye uyan kayıt yok' : 'Henüz kayıt yok' }));
    wrap.append(h('button', { class: 'm-addrow', html: icon('plus') + '<span>Yeni kayıt</span>', onclick: () => this.rowForm(T, null) }));
    requestAnimationFrame(() => this.paintSel());
    return wrap;
  }
  // hücre rengi: önce hücrenin kendi biçimi, sonra renk kuralları
  paintCell(el, si, r, c, st) {
    const own = (this.wb.cell(si, r, c) || {}).s || {};
    const bg = (st && st.bg) || own.bg, fc = (st && st.fc) || own.fc;
    if (bg) { el.style.background = bg; el.style.color = fc || '#18181b'; } else if (fc) el.style.color = fc;
    if (own.b || (st && st.b)) el.style.fontWeight = '700';
    if (st && st.bar != null) el.style.backgroundImage = `linear-gradient(90deg, ${st.barColor}55 ${(st.bar * 100).toFixed(1)}%, transparent ${(st.bar * 100).toFixed(1)}%)`;
    if (st && st.icon) el.prepend(h('span', { class: 'm-ico ' + st.icon, text: st.icon === 'up' ? '▲ ' : st.icon === 'mid' ? '● ' : '▼ ' }));
  }
  // ---------- seçim: fareyle sürükle, dokunmatikte nokta-kutu ----------
  selRect() { const m = this.msel; if (!m || m.si !== this.wb.active || !this._vr) return null; const n = this._vr.length - 1, k = this._vc.length - 1; if (n < 0) return null; const cl = (v, hi) => Math.max(0, Math.min(hi, v)); return { r1: cl(Math.min(m.a.ri, m.b.ri), n), r2: cl(Math.max(m.a.ri, m.b.ri), n), c1: cl(Math.min(m.a.ci, m.b.ci), k), c2: cl(Math.max(m.a.ci, m.b.ci), k) }; }
  selCells() { const g = this.selRect(); if (!g) return []; const out = []; for (let i = g.r1; i <= g.r2; i++) for (let j = g.c1; j <= g.c2; j++) out.push({ ri: i, ci: j, r: this._vr[i], col: this._vc[j] }); return out; }
  paintSel() {
    const g = this.selRect(); const tb = this.body.querySelector('.m-t tbody'); if (!tb) { this.selBar(); return; }
    U.$$('td.sel', tb).forEach(td => { td.classList.remove('sel', 'sel-a'); td.style.removeProperty('--st'); td.style.removeProperty('--sb'); td.style.removeProperty('--sl'); td.style.removeProperty('--sr'); });
    if (!this._hdrag) U.$$('.m-handle', tb).forEach(x => x.remove());
    if (g) {
      for (let i = g.r1; i <= g.r2; i++) for (let j = g.c1; j <= g.c2; j++) {
        const td = tb.querySelector(`td[data-ri="${i}"][data-ci="${j}"]`); if (!td) continue;
        td.classList.add('sel'); if (i === this.msel.a.ri && j === this.msel.a.ci) td.classList.add('sel-a');
        td.style.setProperty('--st', i === g.r1 ? 1 : 0); td.style.setProperty('--sb', i === g.r2 ? 1 : 0); td.style.setProperty('--sl', j === g.c1 ? 1 : 0); td.style.setProperty('--sr', j === g.c2 ? 1 : 0);
        if (i === g.r2 && j === g.c2 && !this._hdrag) { const hd = h('span', { class: 'm-handle', title: 'Sürükleyerek seçimi genişletin' }); td.append(hd); this.bindHandle(hd); }
      }
    }
    this.selBar();
  }
  cellAt(x, y) { const el = document.elementFromPoint(x, y); const td = el && el.closest && el.closest('td[data-ri]'); return td ? { ri: +td.dataset.ri, ci: +td.dataset.ci, td } : null; }
  autoScroll(y, x) { const sc = this.body, rb = sc.getBoundingClientRect(); if (y > rb.bottom - 40) sc.scrollTop += 18; else if (y < rb.top + 60) sc.scrollTop -= 18; const sx = this.body.querySelector('.m-scrollx'); if (sx) { const rx = sx.getBoundingClientRect(); if (x > rx.right - 30) sx.scrollLeft += 18; else if (x < rx.left + 50) sx.scrollLeft -= 18; } }
  bindHandle(hd) {
    // nokta: parmakla / fareyle sürükleyince seçim büyür
    hd.addEventListener('pointerdown', e => {
      e.preventDefault(); e.stopPropagation(); this._hdrag = true; hd.style.visibility = 'hidden';
      const move = ev => { ev.preventDefault(); const c = this.cellAt(ev.clientX, ev.clientY); if (c && (c.ri !== this.msel.b.ri || c.ci !== this.msel.b.ci)) { this.msel.b = { ri: c.ri, ci: c.ci }; this.paintSel(); } this.autoScroll(ev.clientY, ev.clientX); };
      const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); removeEventListener('pointercancel', up); this._hdrag = false; this.paintSel(); };
      addEventListener('pointermove', move, { passive: false }); addEventListener('pointerup', up); addEventListener('pointercancel', up);
    });
  }
  bindSel(tb, T) {
    const si = T.si;
    tb.addEventListener('pointerdown', e => {
      const td = e.target.closest('td[data-ri]'); if (!td || td.classList.contains('editing') || e.target.closest('.m-handle')) return;
      const at = { ri: +td.dataset.ri, ci: +td.dataset.ci };
      if (e.pointerType === 'mouse') {
        if (e.button !== 0) return;
        const g = this.selRect(); const wasOnly = g && g.r1 === g.r2 && g.c1 === g.c2 && g.r1 === at.ri && g.c1 === at.ci;
        if (e.shiftKey && this.msel) { this.msel.b = at; this.paintSel(); return; }
        this.msel = { si, a: at, b: { ...at } }; this.paintSel(); let moved = false;
        const move = ev => { const c = this.cellAt(ev.clientX, ev.clientY); if (c && (c.ri !== this.msel.b.ri || c.ci !== this.msel.b.ci)) { moved = true; this.msel.b = { ri: c.ri, ci: c.ci }; this.paintSel(); } if (ev.buttons) this.autoScroll(ev.clientY, ev.clientX); };
        const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); if (!moved && wasOnly) this.editAt(at); };
        addEventListener('pointermove', move); addEventListener('pointerup', up);
        return;
      }
      // dokunmatik: dokun = seç (seçiliyse düzenle); basılı tut + kaydır = alan seç; basılı tutup bırak = menü; hızlı kaydırma = sayfa kayar
      this._touchAt = Date.now();
      const x0 = e.clientX, y0 = e.clientY; let gone = false, held = false, moved = false;
      const lp = setTimeout(() => { held = true; this._tdrag = true; if (!this.inSel(at)) { this.msel = { si, a: at, b: { ...at } }; this.paintSel(); } navigator.vibrate && navigator.vibrate(15); }, 380);
      const move = ev => {
        if (held) { const c = this.cellAt(ev.clientX, ev.clientY); if (c && (!moved || c.ri !== this.msel.b.ri || c.ci !== this.msel.b.ci)) { if (!moved) this.msel = { si, a: at, b: { ...at } }; moved = true; this.msel.b = { ri: c.ri, ci: c.ci }; this.paintSel(); } this.autoScroll(ev.clientY, ev.clientX); return; }
        if (Math.abs(ev.clientX - x0) > 8 || Math.abs(ev.clientY - y0) > 8) { gone = true; clearTimeout(lp); end(); }
      };
      const up = () => {
        clearTimeout(lp); end();
        if (held) { if (!moved) this.selMenu({ clientX: x0, clientY: y0 }); return; }
        if (gone) return; const g = this.selRect(); const wasOnly = g && g.r1 === g.r2 && g.c1 === g.c2 && g.r1 === at.ri && g.c1 === at.ci; if (wasOnly) this.editAt(at); else { this.msel = { si, a: at, b: { ...at } }; this.paintSel(); }
      };
      const end = () => { this._tdrag = false; removeEventListener('pointermove', move); removeEventListener('pointerup', up); removeEventListener('pointercancel', end); };
      addEventListener('pointermove', move); addEventListener('pointerup', up); addEventListener('pointercancel', end);
    });
    // basılı tutarak seçerken sayfa kaymasın
    tb.addEventListener('touchmove', e => { if (this._tdrag) e.preventDefault(); }, { passive: false });
    tb.addEventListener('dblclick', e => { const td = e.target.closest('td[data-ri]'); if (td) this.editAt({ ri: +td.dataset.ri, ci: +td.dataset.ci }); });
    tb.addEventListener('contextmenu', e => { const td = e.target.closest('td[data-ri]'); if (!td) return; e.preventDefault(); if (this._touchAt && Date.now() - this._touchAt < 1500) return; const at = { ri: +td.dataset.ri, ci: +td.dataset.ci }; if (!this.inSel(at)) { this.msel = { si, a: at, b: { ...at } }; this.paintSel(); } this.selMenu(e); });
  }
  inSel(at) { const g = this.selRect(); return !!g && at.ri >= g.r1 && at.ri <= g.r2 && at.ci >= g.c1 && at.ci <= g.c2; }
  editAt(at, text) { const td = this.body.querySelector(`.m-t td[data-ri="${at.ri}"][data-ci="${at.ci}"]`); if (td) this.editCell(td, this._vT, this._vr[at.ri], this._vc[at.ci], text); }
  // alt çubuk: seçimin toplamı, ortalaması… (Excel'in durum çubuğu gibi)
  selBar() {
    let bar = $('#m-selbar'); const cells = this.view === 'list' ? this.selCells() : [];
    if (!cells.length) { if (bar) bar.remove(); return; }
    if (!bar) { bar = h('div', { id: 'm-selbar' }); document.body.append(bar); }
    const wb = this.wb, si = this.wb.active; const nums = cells.map(c => wb.value(si, c.r, c.col.c)).filter(v => typeof v === 'number');
    const nf = (cells.find(c => typeof wb.value(si, c.r, c.col.c) === 'number') || {}).col; const f = v => X.formatValue(v, (nf && nf.nf) || '#,##0.##');
    const sum = nums.reduce((a, b) => a + b, 0);
    const parts = [h('span', { class: 'm-sb-n', text: cells.length > 1 ? cells.length + ' hücre' : (cells[0].col.name) })];
    if (nums.length) { parts.push(h('span', null, h('small', { text: 'Toplam' }), h('b', { text: f(sum) }))); if (nums.length > 1) parts.push(h('span', null, h('small', { text: 'Ortalama' }), h('b', { text: f(sum / nums.length) })), h('span', { class: 'm-sb-x' }, h('small', { text: 'En küçük' }), h('b', { text: f(Math.min(...nums)) })), h('span', { class: 'm-sb-x' }, h('small', { text: 'En büyük' }), h('b', { text: f(Math.max(...nums)) }))); }
    else parts.push(h('span', null, h('small', { text: 'Dolu' }), h('b', { text: String(cells.filter(c => { const v = wb.value(si, c.r, c.col.c); return v != null && v !== ''; }).length) })));
    bar.innerHTML = ''; bar.append(...parts, h('button', { class: 'm-sb-btn', html: icon('dots'), title: 'Seçenekler', 'aria-label': 'Seçenekler', onclick: e => this.selMenu(e.currentTarget) }), h('button', { class: 'm-sb-btn', html: icon('x'), title: 'Seçimi bırak', 'aria-label': 'Seçimi bırak', onclick: () => { this.msel = null; this.paintSel(); } }));
  }
  selTSV() { const g = this.selRect(); if (!g) return ''; const wb = this.wb, si = wb.active; const lines = []; for (let i = g.r1; i <= g.r2; i++) { const row = []; for (let j = g.c1; j <= g.c2; j++) row.push(wb.display(si, this._vr[i], this._vc[j].c)); lines.push(row.join('\t')); } return lines.join('\n'); }
  async selCopy() { const t = this.selTSV(); try { await navigator.clipboard.writeText(t); U.toast('✓ Kopyalandı'); } catch (e) { U.toast('Kopyalamak için Ctrl+C kullanın'); } }
  selPaste(text) {
    const g = this.selRect(); if (!g || !text) return; const wb = this.wb, si = wb.active; const T = this._vT;
    const grid = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
    let skipped = 0; wb.begin();
    grid.forEach((row, i) => row.forEach((v, j) => { const ri = g.r1 + i, ci = g.c1 + j; if (ri >= this._vr.length || ci >= this._vc.length) return; const col = this._vc[ci]; if (col.computed) { skipped++; return; } this.commitInput(si, this._vr[ri], col.c, v); }));
    wb.end(); U.toast('✓ Yapıştırıldı' + (skipped ? ' (' + skipped + ' hesaplanan hücre atlandı)' : ''));
  }
  async selClear() {
    const cells = this.selCells().filter(c => !c.col.computed); if (!cells.length) { U.toast('Hesaplanan hücreler silinmez'); return; }
    if (cells.length > 20 && !await U.confirm(cells.length + ' hücrenin içi temizlensin mi?', 'Temizle', 'Temizle')) return;
    const wb = this.wb, si = wb.active; wb.begin(); cells.forEach(c => wb.patchCell(si, c.r, c.col.c, { v: null, f: null })); wb.end();
  }
  selMenu(at) {
    const cells = this.selCells(); if (!cells.length) return; const wb = this.wb, si = wb.active, T = this._vT, vs = this.vsCur;
    const anchor = at && at.getBoundingClientRect ? at : { left: at.clientX, right: at.clientX, top: at.clientY, bottom: at.clientY };
    const g = this.selRect(); const one = cells.length === 1; const c0 = cells[0]; const rowsSel = [...new Set(cells.map(c => c.r))]; const colsSel = [...new Set(cells.map(c => c.col))];
    const nums = cells.map(c => wb.value(si, c.r, c.col.c)).filter(v => typeof v === 'number'); const sum = nums.reduce((a, b) => a + b, 0);
    const setAll = patch => { wb.begin(); cells.forEach(c => wb.setStyle(si, c.r, c.col.c, patch)); wb.end(); };
    const sub = items => setTimeout(() => U.menu(anchor, items), 10);
    const disp = wb.display(si, c0.r, c0.col.c);
    U.menu(anchor, [
      { head: one ? c0.col.name + ': ' + (disp || 'boş') : cells.length + ' hücre seçili' + (nums.length ? ' · Toplam ' + X.formatValue(sum, (colsSel.find(c => c.kind !== 'text') || {}).nf || '#,##0.##') : '') },
      ...(one && !c0.col.computed ? [{ text: 'Düzenle', icon: 'rules', onClick: () => this.editAt({ ri: g.r1, ci: g.c1 }) }] : []),
      { text: 'Kopyala', icon: 'copy', kbd: 'Ctrl+C', onClick: () => this.selCopy() },
      { text: 'Yapıştır', icon: 'paste', kbd: 'Ctrl+V', onClick: async () => { try { this.selPaste(await navigator.clipboard.readText()); } catch (e) { U.toast('Yapıştırmak için Ctrl+V kullanın'); } } },
      ...(nums.length > 1 ? [{ text: 'Toplamı kopyala (' + X.formatValue(sum, '#,##0.##') + ')', icon: 'sigma', onClick: async () => { try { await navigator.clipboard.writeText(X.formatValue(sum, '0.##')); U.toast('✓ Toplam kopyalandı'); } catch (e) { } } }] : []),
      { text: 'İçini temizle', icon: 'eraser', kbd: 'Delete', onClick: () => this.selClear() },
      { sep: 1 },
      { text: 'Biçim ▸', icon: 'hash', onClick: () => sub(KINDS.map(([k, t, b]) => ({ text: b + '  ' + t, onClick: () => setAll({ nf: KIND_NF[k] }) }))) },
      { text: 'Renk ▸', icon: 'fill', onClick: () => sub([...TONES.map(t => ({ text: t[0], swatch: t[1], onClick: () => setAll({ bg: t[1], fc: t[2] }) })), { sep: 1 }, { text: 'Rengi kaldır', icon: 'eraser', onClick: () => setAll({ bg: null, fc: null }) }]) },
      { text: 'Kalın', icon: 'bold', onClick: () => { const b = !((wb.cell(si, c0.r, c0.col.c) || {}).s || {}).b; setAll({ b: b ? 1 : null }); } },
      { sep: 1 },
      ...(one && disp ? [{ text: 'Sadece “' + disp + '” olanları göster', icon: 'filter', onClick: () => { vs.filters.push({ c: c0.col.c, op: '=', val: disp }); this.msel = null; this.render(); } }, { text: '“' + disp + '” olmayanları göster', icon: 'eyeOff', onClick: () => { vs.filters.push({ c: c0.col.c, op: '<>', val: disp }); this.msel = null; this.render(); } }] : []),
      { text: 'Küçükten büyüğe sırala', icon: 'sortA', onClick: () => { vs.sort = { c: c0.col.c, dir: 1 }; this.msel = null; this.render(); } },
      { text: 'Büyükten küçüğe sırala', icon: 'sortD', onClick: () => { vs.sort = { c: c0.col.c, dir: -1 }; this.msel = null; this.render(); } },
      { sep: 1 },
      { text: 'Seçimle formül kur ▸', icon: 'calc', onClick: () => sub(this.selFormulaItems()) },
      { text: 'Seçimden grafik yap', icon: 'chart', onClick: () => { const lab = colsSel.find(c => c.kind === 'text' || c.kind === 'date') || T.cols.find(c => c.kind === 'text') || T.cols[0]; const vals = colsSel.filter(c => c.kind !== 'text' && c.kind !== 'date').map(c => ({ c: c.c, op: 'sum' })); this.chartWizard(T, null, { lab: lab.c, labMode: lab.kind === 'date' ? 'month' : 'value', vals: vals.length ? vals : [{ c: -1, op: 'count' }], conds: rowsSel.length < this._vr.length && rowsSel.length > 1 ? [] : [] }); } },
      { text: 'Bu sütunla hesapla…', icon: 'wand', onClick: () => { this.focusCol(T, c0.col.c); this.wiz.openRecipes(); } },
      { text: 'Renk kuralı…', icon: 'rules', onClick: () => this.colorDialog(T, c0.col) },
      { sep: 1 },
      ...(rowsSel.length === 1 ? [{ text: 'Kaydı aç', icon: 'file', onClick: () => this.rowForm(T, c0.r) }] : []),
      { text: rowsSel.length > 1 ? rowsSel.length + ' kaydı sil' : 'Kaydı sil', icon: 'trash', danger: 1, onClick: async () => { if (!await U.confirm(rowsSel.length + ' kayıt (satır) silinsin mi? Geri al ile geri getirebilirsiniz.', 'Kayıt sil', 'Sil')) return; wb.begin(); rowsSel.slice().sort((a, b) => b - a).forEach(r => wb.shift(si, 'r', r, -1)); wb.end(); this.msel = null; this.render(); } },
      { text: 'Yeni kayıt ekle', icon: 'plus', onClick: () => this.rowForm(T, null) }
    ]);
  }
  // ---------- seçimle formül: sonuçlar "Hesaplar" tablosuna, formülüyle ----------
  selRefs() {
    const g = this.selRect(); if (!g) return null; const T = this._vT; const sh = X.quoteSheet(this.wb.sheets[T.si].name) + '!';
    const rows = this._vr.slice(g.r1, g.r2 + 1).slice().sort((a, b) => a - b); const cols = this._vc.slice(g.c1, g.c2 + 1);
    const contig = rows.every((r, i) => !i || r === rows[i - 1] + 1) && cols.every((c, i) => !i || c.c === cols[i - 1].c + 1);
    const abs = (r, c) => '$' + X.colName(c) + '$' + (r + 1);
    const block = contig ? sh + abs(rows[0], cols[0].c) + (rows.length > 1 || cols.length > 1 ? ':' + abs(rows[rows.length - 1], cols[cols.length - 1].c) : '') : null;
    const list = []; rows.forEach(r => cols.forEach(c => list.push(sh + abs(r, c.c))));
    return { rows, cols, contig, block, any: block || list.join(','), sheetRows: { r1: rows[0], r2: rows[rows.length - 1] } };
  }
  calcSheet() {
    const wb = this.wb; let i = wb.sheetIndex('Hesaplar');
    if (i < 0) { const cur = wb.active; wb.begin(); const s = wb.addSheet('Hesaplar'); i = wb.sheets.indexOf(s); const hs = { b: 1, bg: '#16a34a', fc: '#ffffff', ha: 'center' }; wb.patchCell(i, 0, 0, { v: 'Hesap', f: null }); wb.setStyle(i, 0, 0, hs); wb.patchCell(i, 0, 1, { v: 'Sonuç', f: null }); wb.setStyle(i, 0, 1, hs); s.colW = { 0: 220, 1: 160 }; wb.active = cur; wb.end(); }
    const d = wb.dims(i); let r = 1; while (r < d.rows && (wb.value(i, r, 0) != null || wb.value(i, r, 1) != null)) r++;
    return { si: i, r };
  }
  addCalc(name, f, nf) {
    const wb = this.wb; const t = this.calcSheet();
    wb.begin(); wb.patchCell(t.si, t.r, 0, { v: name, f: null }); wb.patchCell(t.si, t.r, 1, { f: X.toStore(f), v: null }); if (nf) wb.setStyle(t.si, t.r, 1, { nf }); wb.end();
    const v = wb.display(t.si, t.r, 1);
    U.toast('✓ ' + name + ' = ' + v + ' (Hesaplar tablosuna eklendi)');
    return t;
  }
  selFormulaItems() {
    const S = this.selRefs(); if (!S) return []; const wb = this.wb, T = this._vT;
    const numCols = S.cols.filter(c => c.kind !== 'text' && c.kind !== 'date'); const nf = (numCols[0] || {}).nf || null;
    const what = (S.cols.length === 1 ? S.cols[0].name : S.cols.map(c => c.name).join(' + ')) + (S.rows.length < this._vr.length ? ' (' + S.rows.length + ' kayıt)' : '');
    const agg = (fn, tr, nm, keepNf) => ({ text: nm + ' (' + tr + ')', onClick: () => this.addCalc(nm + ' ' + what, fn + '(' + S.any + ')', keepNf ? nf : null) });
    const items = [{ head: 'Seçim: ' + what }];
    if (numCols.length) items.push(agg('SUM', 'TOPLA', 'Toplam', 1), agg('AVERAGE', 'ORTALAMA', 'Ortalama', 1), agg('MAX', 'MAK', 'En büyük', 1), agg('MIN', 'MİN', 'En küçük', 1));
    items.push(agg('COUNTA', 'BAĞ_DEĞ_DOLU_SAY', 'Dolu hücre sayısı', 0));
    if (S.cols.length === 1 && S.cols[0].kind === 'text' && S.contig) items.push({ text: 'Farklı değer sayısı', onClick: () => this.addCalc('Farklı ' + what, 'COUNTA(UNIQUE(' + S.block + '))') });
    items.push({ sep: 1 });
    items.push({ text: 'DÜŞEYARA ile bul (Bul ve getir)…', icon: 'pointer', onClick: () => { const key = S.cols[0]; const vals = []; S.rows.forEach(r => { const d = wb.display(T.si, r, key.c); if (d && !vals.includes(d)) vals.push(d); }); this.lookupTool({ si: T.si, key: key.c, rets: S.cols.slice(1).map(c => c.c), picks: vals }); } });
    if (numCols.length === 1 && S.contig && S.rows.length > 1) items.push({ text: 'Sonraki değeri tahmin et (TAHMİN)', icon: 'line', onClick: () => { const y = S.block; this.addCalc('Tahmin: sonraki ' + numCols[0].name, 'FORECAST.LINEAR(ROWS(' + y + ')+1,' + y + ',SEQUENCE(ROWS(' + y + ')))', nf); } });
    items.push(S.contig ? { text: 'Başka işlev (234 işlev)…', icon: 'calc', onClick: () => this.selFunction(S) } : { text: 'Başka işlev: sıralama/filtre varken seçim parçalı', icon: 'calc', disabled: true });
    return items;
  }
  selFunction(S) {
    // seçilen alan işlevin ilk alan kutusuna yazılır, sonuç Hesaplar tablosuna
    const wb = this.wb, T = this._vT; const t = this.calcSheet();
    const g = { r1: S.rows[0], r2: S.rows[S.rows.length - 1], c1: S.cols[0].c, c2: S.cols[S.cols.length - 1].c };
    this.sel = { r: g.r1, c: g.c1, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 };
    this.fnTarget = () => ({ si: t.si, r: t.r, c: 1 });
    this._calcName = (S.cols.length === 1 ? S.cols[0].name : 'Hesap');
    this.wiz.openFunctions();
    // pencere kapanınca hedefi bırak
    const off = () => { if (!U.topDialog()) { this.fnTarget = null; this._calcName = null; } else setTimeout(off, 400); }; setTimeout(off, 400);
  }
  editCell(td, T, r, col, startText) {
    const wb = this.wb, si = T.si; const cell = wb.cell(si, r, col.c);
    if (X.isErr(wb.value(si, r, col.c)) && this.errHint(si, r, col.c)) return;
    if (col.computed || (cell && cell.f != null) || (!cell && wb.spillAt(si, r, col.c))) {
      const f = cell && cell.f != null ? cell.f : null;
      U.toast(f ? 'Bu değer hesaplanıyor: =' + X.toDisplay(f) : 'Bu değer bir hesabın sonucu'); return;
    }
    if (td.querySelector('input')) return;
    const old = wb.editText(si, r, col.c);
    const inp = h('input', { class: 'm-cellinp', value: startText != null ? startText : old });
    const list = this.choices(T, col);
    td.textContent = ''; td.classList.add('editing'); td.append(inp); inp.focus(); if (startText == null) inp.select();
    if (list.length && list.length <= 12 && col.kind === 'text') {
      const dl = h('datalist', { id: 'm-dl' }, ...list.map(v => h('option', { value: v }))); td.append(dl); inp.setAttribute('list', 'm-dl');
    }
    let done = false;
    const finish = ok => { if (done) return; done = true; if (ok && inp.value !== old) { if (!this.commitInput(si, r, col.c, inp.value)) this.render(); } else this.render(); };
    inp.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === 'Tab') { e.preventDefault(); finish(true); if (this.msel) { const a = this.msel.a; const n = e.key === 'Tab' ? { ri: a.ri, ci: Math.min(this._vc.length - 1, a.ci + 1) } : { ri: Math.min(this._vr.length - 1, a.ri + 1), ci: a.ci }; this.msel = { si: this.msel.si, a: n, b: { ...n } }; } } else if (e.key === 'Escape') finish(false); });
    inp.addEventListener('blur', () => finish(true));
  }
  errHint(si, r, c) {
    const wb = this.wb, v = wb.value(si, r, c), cell = wb.cell(si, r, c);
    if (v.code === '#TAŞMA!' && cell && cell._blk) {
      const b = cell._blk, blockers = [];
      for (let i = b.r1; i <= b.r2; i++) for (let j = b.c1; j <= b.c2; j++) { if (i === r && j === c) continue; const o = wb.cell(si, i, j); if (o && ((o.v != null && o.v !== '') || o.f != null)) blockers.push(X.addr(i, j)); }
      U.confirm('Bu hesabın sonucu birden çok satıra yazılan bir liste, ama yazılacağı yerde ' + (blockers.length ? blockers.length + ' dolu hücre var.' : 'başka bir liste var.') + (blockers.length ? ' Engel olan hücreleri temizleyeyim mi?' : ''), 'Sonuç için yer yok', blockers.length ? 'Temizle' : 'Tamam').then(ok => {
        if (!ok || !blockers.length) return;
        wb.begin(); for (let i = b.r1; i <= b.r2; i++) for (let j = b.c1; j <= b.c2; j++) { if (i === r && j === c) continue; const o = wb.cell(si, i, j); if (o && ((o.v != null && o.v !== '') || o.f != null)) wb.patchCell(si, i, j, { v: null, f: null }); } wb.end(); U.toast('✓ Alan temizlendi');
      });
      return true;
    }
    const msg = { '#DIV/0!': 'Sıfıra bölme var: bölen boş ya da 0.', '#VALUE!': 'Sayı beklenen yerde metin var.', '#N/A': 'Aranan değer bulunamadı.', '#NAME?': 'Tanınmayan işlev ya da ad.', '#REF!': 'Kullanılan hücre silinmiş.', '#NUM!': 'Hesaplanamayan sayı.', '#DÖNGÜ!': 'Hesap kendi sonucunu kullanıyor.', '#HESAP!': 'Sonuç boş liste.', '#TAŞMA!': 'Sonuç listesi için yer yok.' }[v.code];
    if (msg) { U.toast(wb.display(si, r, c) + ' → ' + msg); return true; }
    return false;
  }
  // açılır liste kuralı (DÜŞEYARA arama tablosu gibi) varsa onun değerleri, yoksa sütundaki metinler
  choices(T, col) {
    const d = this.wb.dvAt(T.si, T.r0, col.c);
    if (d && d.type === 'list') { const m = new Map(); for (const x of this.wb.dvItems(T.si, d)) if (!m.has(lc(x))) m.set(lc(x), x); return [...m.values()]; }
    return col.kind === 'text' ? this.distinct(T, col) : [];
  }
  distinct(T, col) { const s = new Map(); for (let r = T.r0; r <= T.r2; r++) { const d = this.wb.display(T.si, r, col.c); if (d !== '' && !s.has(lc(d))) s.set(lc(d), d); } return [...s.values()]; }
  commitInput(si, r, c, text) {
    const wb = this.wb; const cell = wb.cell(si, r, c); const st = cell && cell.s;
    if (text.startsWith('=') && text.length > 1) {
      let f; try { f = X.toStore(text.slice(1)); X.parse(f); } catch (e) { U.toast('Formülde hata var: ' + e.message, 'warn'); return false; }
      wb.begin(); wb.patchCell(si, r, c, { f, v: null }); wb.end(); return true;
    }
    const pi = IO.parseInput(text); const bad = wb.dvCheck(si, r, c, pi.v);
    if (bad) { U.alert(bad, 'Kural'); return false; }
    wb.begin(); wb.patchCell(si, r, c, { v: pi.v, f: null });
    if (pi.nf && !(st && st.nf && (X.isDateFormat(st.nf) === X.isDateFormat(pi.nf)))) wb.setStyle(si, r, c, { nf: pi.nf });
    wb.end(); return true;
  }

  // ---------- Kartlar ----------
  viewCards(T) {
    const wb = this.wb, si = T.si; const rows = this.rowsOf(T);
    const titleCol = T.cols.find(c => c.kind === 'text' && !c.computed) || T.cols[0];
    const grid = h('div', { class: 'm-cards' });
    for (const r of rows) {
      const card = h('button', { class: 'm-rcard', onclick: () => this.rowForm(T, r) }, h('div', { class: 'm-rcard-t', text: wb.display(si, r, titleCol.c) || '(adsız)' }));
      const dl = h('dl');
      for (const col of T.cols) { if (col === titleCol) continue; const dd = h('dd', { class: (col.computed ? 'm-fx ' : '') + (X.isErr(wb.value(si, r, col.c)) ? 'err' : '') }); const pill = h('span', { class: 'm-pill', text: wb.display(si, r, col.c) || '—' }); this.paintCell(pill, si, r, col.c, wb.cfStyle(si, r, col.c)); dd.append(pill); dl.append(h('dt', null, this.badge(col), h('span', { text: col.name })), dd); }
      card.append(dl); grid.append(card);
    }
    grid.append(h('button', { class: 'm-rcard add', onclick: () => this.rowForm(T, null) }, h('span', { html: icon('plus') }), h('span', { text: 'Yeni kayıt' })));
    return grid;
  }

  // ---------- kayıt formu (ekle / düzenle) ----------
  rowForm(T, r) {
    const wb = this.wb, si = T.si; const isNew = r == null;
    const body = h('div', { class: 'form m-form' }); const ins = [];
    for (const col of T.cols) {
      const row = h('div', { class: 'frow' }, h('label', { class: 'lbl' }, this.badge(col), h('span', { text: ' ' + col.name })));
      if (col.computed) { row.append(h('div', { class: 'm-ro', text: isNew ? 'Kaydedince hesaplanır' : wb.display(si, r, col.c) || '—' })); body.append(row); continue; }
      const inp = h('input', { class: 'inp', value: isNew ? '' : wb.editText(si, r, col.c), placeholder: col.kind === 'date' ? 'Ör: 30.10.2026' : col.kind === 'money' ? 'Ör: 1250' : col.kind === 'pct' ? 'Ör: %18' : '', inputmode: col.kind === 'text' || col.kind === 'date' ? 'text' : 'decimal' });
      row.append(inp);
      const list = this.choices(T, col);
      if (list.length && list.length <= 30) { const ch = h('div', { class: 'vchips' }, ...list.map(v => h('button', { class: 'chip sm', type: 'button', text: v, onclick: () => { inp.value = v; U.$$('.chip', ch).forEach(x => x.classList.toggle('on', x.textContent === v)); } }))); row.append(ch); }
      ins.push({ col, inp }); body.append(row);
    }
    let rowTone;
    const toneRow = h('div', { class: 'm-tones' }, h('button', { class: 'm-tone none', type: 'button', title: 'Renk yok', text: '✕', onclick: e => pickTone(e.currentTarget, null) }), ...TONES.map(t => h('button', { class: 'm-tone', type: 'button', title: t[0], style: { background: t[1], borderColor: t[2] }, onclick: e => pickTone(e.currentTarget, t) })));
    const pickTone = (b, t) => { rowTone = t || 0; U.$$('.m-tone', toneRow).forEach(x => x.classList.toggle('on', x === b)); };
    body.append(h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Kaydın rengi (isteğe bağlı)' }), toneRow));
    const btns = [];
    if (!isNew) btns.push({ text: 'Sil', kind: 'danger', onClick: async () => { if (await U.confirm('Bu kayıt silinsin mi?', 'Kaydı sil', 'Sil')) { wb.shift(si, 'r', r, -1); d.close(); } return false; } });
    btns.push({ text: 'Vazgeç', value: null }, { text: isNew ? 'Ekle' : 'Kaydet', kind: 'primary', onClick: () => {
      wb.begin();
      if (isNew) r = this.insertRow(T);
      for (const { col, inp } of ins) { const old = wb.editText(si, r, col.c); if (inp.value !== old) this.commitInput(si, r, col.c, inp.value); }
      if (rowTone !== undefined) for (const col of T.cols) wb.setStyle(si, r, col.c, rowTone ? { bg: rowTone[1], fc: rowTone[2] } : { bg: null, fc: null });
      wb.end(); U.toast(isNew ? '✓ Kayıt eklendi' : '✓ Kaydedildi'); return true;
    }, value: 1 });
    const d = U.dialog({ title: icon(isNew ? 'plus' : 'file') + (isNew ? ' Yeni kayıt' : ' Kayıt'), body, buttons: btns });
    setTimeout(() => { const f = ins[0] && ins[0].inp; if (f && !isMobile()) f.focus(); }, 60);
  }
  // tablonun sonuna (Toplam satırının üstüne) boş satır ekler; hesaplanan sütunların formülünü üstteki satırdan kopyalar
  insertRow(T) {
    const wb = this.wb, si = T.si; const at = T.r2 + 1; const last = T.r2 >= T.r0 ? T.r2 : null;
    wb.begin();
    if (T.totals.length) {
      wb.shift(si, 'r', at, 1);
      // Toplam formüllerinin alanını yeni satırı kapsayacak şekilde genişlet
      for (const tr of T.totals) { const rr = tr + 1; const d = wb.dims(si); for (let c = 0; c < d.cols; c++) { const cl = wb.cell(si, rr, c); if (cl && cl.f) { const nf = cl.f.replace(/(:\$?[A-Z]+\$?)(\d+)/g, (m, a, n) => +n === at ? a + (at + 1) : m); if (nf !== cl.f) wb.patchCell(si, rr, c, { f: nf, v: null }); } } }
    }
    if (last != null) for (const col of T.cols) {
      const src = wb.cell(si, last, col.c);
      if (src && src.f != null && col.computed) wb.patchCell(si, at, col.c, { f: X.shiftFormula(src.f, at - last, 0), v: null });
      if (src && src.s) wb.setStyle(si, at, col.c, Object.assign({}, src.s));
    }
    wb.end(); return at;
  }

  // ---------- sütunlar ----------
  colMenu(anchor, T, col) {
    const vs = this.vsCur; const wb = this.wb, si = T.si;
    const items = [
      { head: col.name },
      { text: 'Küçükten büyüğe sırala', icon: 'sortA', onClick: () => { vs.sort = { c: col.c, dir: 1 }; this.render(); } },
      { text: 'Büyükten küçüğe sırala', icon: 'sortD', onClick: () => { vs.sort = { c: col.c, dir: -1 }; this.render(); } },
      { text: 'Bu sütuna göre filtrele…', icon: 'filter', onClick: () => this.filterDialog(T, col) },
      { sep: 1 },
      { text: 'Bu sütunla hesapla…', icon: 'wand', onClick: () => { this.focusCol(T, col.c); this.wiz.openRecipes(); } },
      { text: 'Renk ver…', icon: 'fill', onClick: () => this.colorDialog(T, col) },
      { text: 'Bu sütuna göre özet', icon: 'sigma', onClick: () => { this.vsCur.sumG = col.c; this.view = 'summary'; this.render(); } },
      { sep: 1 },
      { text: 'Adını değiştir', icon: 'tag', disabled: T.hr < 0, onClick: async () => { const v = await U.prompt('Sütun adı', col.name); if (v && v.trim()) { wb.begin(); wb.patchCell(si, T.hr, col.c, { v: v.trim(), f: null }); wb.end(); } } },
      ...(col.computed ? [{ text: 'Hesabı göster', icon: 'calc', onClick: () => { const cl = wb.cell(si, T.r0, col.c); U.alert(cl && cl.f ? 'Bu sütunun her satırında şu hesap var:\n=' + X.toDisplay(cl.f) : 'Bu sütun başka bir hesabın sonucunu gösteriyor.', col.name); } }] : []),
      { text: 'Türü: ' + KINDS.find(k => k[0] === col.kind)[1] + ' ▸', icon: 'hash', onClick: () => setTimeout(() => this.kindMenu(anchor, T, col), 10) },
      { text: 'Sütunu sil', icon: 'trash', danger: 1, onClick: async () => { if (await U.confirm('“' + col.name + '” sütunu ve içindeki bütün değerler silinsin mi?', 'Sütunu sil', 'Sil')) wb.shift(si, 'c', col.c, -1); } }
    ];
    U.menu(anchor, items);
  }
  colorDialog(T, col) {
    const wb = this.wb, si = T.si, sh = wb.sheets[si];
    col = col || T.cols.find(c => c.kind !== 'text') || T.cols[0]; if (!col) return;
    const body = h('div', { class: 'form m-colors' });
    const rangeOf = c => ({ r1: T.r0, c1: c.c, r2: Math.max(T.r2, T.r0) + 500, c2: c.c });
    const rulesOf = c => sh.cf.filter(x => x.range.c1 === c.c && x.range.c2 === c.c);
    const tones = (cur, fn) => h('div', { class: 'm-tones' }, ...TONES.map(t => h('button', { class: 'm-tone' + (cur && cur.bg === t[1] ? ' on' : ''), type: 'button', title: t[0], style: { background: t[1], borderColor: t[2] }, onclick: () => fn(t) })));
    const addRule = rule => { wb.begin(); wb.snapSheet(si); sh.cf.push(rule); sh._cfCache = null; wb.end(); draw(); };
    const describe = x => {
      const v = x.a == null ? '' : typeof x.a === 'number' ? X.formatValue(x.a, col.nf || null) : '“' + x.a + '”';
      const op = (CF_OPS.find(o => o[0] === x.type) || [0, ''])[1];
      if (x.type === 'scale') return 'Renk ölçeği (küçük kırmızı → büyük yeşil)'; if (x.type === 'bar') return 'Veri çubuğu'; if (x.type === 'icons') return 'Ok simgeleri';
      if (x.type === 'between') return X.formatValue(x.a, col.nf || null) + ' ile ' + X.formatValue(x.b, col.nf || null) + ' arasındaysa';
      if (x.type === 'eq') return v + ' ise'; if (x.type === 'top' || x.type === 'bottom') return op.replace('N', x.a);
      return (v ? v + ' ' : '') + op;
    };
    let mode = col.kind === 'text' ? 'values' : 'cond';
    const draw = () => {
      body.innerHTML = '';
      const cs = h('select', { class: 'inp' }, ...T.cols.map(c => h('option', { value: c.c, text: c.name }))); cs.value = col.c;
      cs.onchange = () => { col = T.cols.find(c => c.c === +cs.value); mode = col.kind === 'text' ? 'values' : 'cond'; draw(); };
      body.append(h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Hangi sütun?' }), cs));
      const modes = [['values', 'Değere göre'], ['cond', 'Koşula göre'], ...(col.kind !== 'text' ? [['scale', 'Ölçek / çubuk']] : []), ['fill', 'Tümünü boya']];
      body.append(h('div', { class: 'segs m-modes' }, ...modes.map(([k, t]) => h('button', { class: 'seg' + (mode === k ? ' on' : ''), type: 'button', text: t, onclick: () => { mode = k; draw(); } }))));
      const box = h('div', { class: 'm-cbox' }); body.append(box);
      if (mode === 'values') {
        box.append(h('p', { class: 'hint', text: 'Her değere bir renk seçin. Ör: “Geçti” yeşil, “Kaldı” kırmızı.' }));
        const vals = new Map(); for (let r = T.r0; r <= T.r2; r++) { const v = wb.value(si, r, col.c); if (v == null || v === '' || X.isErr(v)) continue; const k = lc(typeof v === 'number' ? X.fmtGeneral(v) : v); if (!vals.has(k)) vals.set(k, { v, d: wb.display(si, r, col.c) }); }
        if (!vals.size) box.append(h('p', { class: 'hint', text: 'Bu sütunda henüz değer yok.' }));
        [...vals.values()].slice(0, 30).forEach(({ v, d }) => {
          const cur = rulesOf(col).find(x => x.type === 'eq' && lc(typeof x.a === 'number' ? X.fmtGeneral(x.a) : x.a) === lc(typeof v === 'number' ? X.fmtGeneral(v) : v));
          const pill = h('span', { class: 'm-pill', text: d }); if (cur) { pill.style.background = cur.style.bg; pill.style.color = cur.style.fc; }
          box.append(h('div', { class: 'm-vrow' }, pill, tones(cur && cur.style, t => { wb.begin(); wb.snapSheet(si); if (cur) sh.cf.splice(sh.cf.indexOf(cur), 1); sh.cf.push({ type: 'eq', a: v, range: rangeOf(col), style: { bg: t[1], fc: t[2] } }); sh._cfCache = null; wb.end(); draw(); }),
            cur ? h('button', { class: 'ic ghost', type: 'button', html: icon('x'), title: 'Rengi kaldır', onclick: () => { wb.begin(); wb.snapSheet(si); sh.cf.splice(sh.cf.indexOf(cur), 1); wb.end(); draw(); } }) : h('span', { class: 'm-vsp' })));
        });
      } else if (mode === 'cond') {
        const op = h('select', { class: 'inp' }, ...CF_OPS.map(([k, t]) => h('option', { value: k, text: t })));
        op.value = col.kind === 'text' ? 'contains' : 'gt';
        const a = h('input', { class: 'inp', placeholder: col.kind === 'text' ? 'Metin' : 'Değer', inputmode: col.kind === 'text' ? 'text' : 'decimal' });
        const b = h('input', { class: 'inp', placeholder: 'İkinci değer', inputmode: 'decimal' });
        const sync = () => { const k = op.value; a.style.display = /^(above|below|dup|blank)$/.test(k) ? 'none' : ''; a.placeholder = /^(top|bottom)$/.test(k) ? 'Kaç tane? Ör: 3' : col.kind === 'text' ? 'Metin' : 'Değer'; b.style.display = k === 'between' ? '' : 'none'; };
        op.onchange = sync; sync();
        let tone = TONES[0];
        const tn = tones({ bg: tone[1] }, t => { tone = t; U.$$('.m-tone', tn).forEach(x => x.classList.toggle('on', x.title === t[0])); });
        box.append(h('div', { class: 'm-cond' }, h('span', { class: 'm-cond-l', text: col.name }), a, op, b), h('label', { class: 'lbl', text: 'Hangi renk?' }), tn,
          h('button', { class: 'btn primary m-addrule', type: 'button', html: icon('plus') + '<span>Kuralı ekle</span>', onclick: () => {
            const k = op.value; const num = t => { const n = X.parseNumLoose(String(t)); if (n != null) return n; const d = X.parseDateLoose(String(t)); return d ? d.v : null; };
            let av = null, bv = null;
            if (!/^(above|below|dup|blank)$/.test(k)) { if (!a.value.trim()) { U.toast('Bir değer yazın'); a.focus(); return; } av = k === 'contains' || (k === 'eq' && num(a.value) == null) ? a.value.trim() : num(a.value); if (av == null) { U.toast('Sayı yazın'); return; } }
            if (k === 'between') { bv = num(b.value); if (bv == null) { U.toast('İkinci değeri yazın'); return; } }
            addRule({ type: k, a: av, b: bv, range: rangeOf(col), style: { bg: tone[1], fc: tone[2] } }); U.toast('✓ Kural eklendi');
          } }));
      } else if (mode === 'scale') {
        const opt = (t, d, rule, prev) => h('button', { class: 'm-opt', type: 'button', onclick: () => { addRule(Object.assign({ range: rangeOf(col), style: {} }, rule)); U.toast('✓ Eklendi'); } }, h('span', { class: 'm-opt-ic', style: prev }), h('span', null, h('b', { text: t }), h('small', { text: d })));
        box.append(opt('Renk ölçeği', 'Küçük değerler kırmızı, ortalar sarı, büyükler yeşil', { type: 'scale' }, { background: 'linear-gradient(90deg,#f8696b,#ffeb84,#63be7b)' }),
          opt('Ters renk ölçeği', 'Küçük değerler yeşil, büyükler kırmızı (ör. giderler)', { type: 'scale', c1: '#63be7b', c3: '#f8696b' }, { background: 'linear-gradient(90deg,#63be7b,#ffeb84,#f8696b)' }),
          opt('Veri çubuğu', 'Her hücrede değeri kadar dolu bir çubuk', { type: 'bar', color: '#4f8ef7' }, { background: 'linear-gradient(90deg,#4f8ef7 60%,#dbeafe 60%)' }),
          opt('Ok simgeleri', 'Büyükler ▲ yeşil, ortalar ●, küçükler ▼ kırmızı', { type: 'icons' }, { background: 'var(--accent-soft)' }));
      } else {
        box.append(h('p', { class: 'hint', text: 'Bu sütundaki bütün değerleri aynı renge boyar.' }));
        const tn = tones(null, t => { wb.begin(); for (let r = T.r0; r <= T.r2; r++) wb.setStyle(si, r, col.c, { bg: t[1], fc: t[2] }); wb.end(); U.toast('✓ Boyandı'); draw(); });
        box.append(tn, h('button', { class: 'btn', type: 'button', html: icon('eraser') + '<span>Boyamayı kaldır</span>', onclick: () => { wb.begin(); for (let r = T.r0; r <= T.r2; r++) wb.setStyle(si, r, col.c, { bg: null, fc: null }); wb.end(); draw(); } }));
      }
      // bu sütunun kuralları
      const rs = rulesOf(col);
      if (rs.length) {
        const ul = h('div', { class: 'm-rules' }, h('div', { class: 'm-sec2', text: 'Bu sütundaki renk kuralları' }));
        rs.forEach(x => { const sw = h('span', { class: 'm-rsw' }); if (x.style && x.style.bg) sw.style.background = x.style.bg; else if (x.type === 'scale') sw.style.background = 'linear-gradient(90deg,' + (x.c1 || '#f8696b') + ',#ffeb84,' + (x.c3 || '#63be7b') + ')'; else if (x.type === 'bar') sw.style.background = '#4f8ef7';
          ul.append(h('div', { class: 'm-rule' }, sw, h('span', { text: describe(x) }), h('button', { class: 'ic ghost', type: 'button', html: icon('trash'), title: 'Kuralı sil', onclick: () => { wb.begin(); wb.snapSheet(si); sh.cf.splice(sh.cf.indexOf(x), 1); sh._cfCache = null; wb.end(); draw(); } }))); });
        body.append(ul);
      }
    };
    draw();
    U.dialog({ title: icon('fill') + ' Renkler', body, wide: true, buttons: [{ text: 'Tamam', kind: 'primary', value: 1 }] });
  }
  // Bul ve getir: bir değer yazınca (Ör: Ankara) seçilen bilgileri getiren arama tablosu (DÜŞEYARA)
  lookupTool(preset) {
    const wb = this.wb;
    let si = preset ? preset.si : wb.active; let S = this.table(si);
    if (!S.cols.length) { const k = wb.sheets.findIndex((_, i) => this.table(i).cols.length); if (k < 0) { U.toast('Önce verisi olan bir tablo gerekir'); return; } si = k; S = this.table(k); }
    let key = (S.cols.find(c => c.kind === 'text' && !c.computed) || S.cols[0]).c;
    let rets = S.cols.filter(c => c.c !== key).map(c => c.c);
    let mode = 'first'; let picks = [];
    if (preset) { key = preset.key; rets = preset.rets && preset.rets.length ? preset.rets : S.cols.filter(c => c.c !== key).map(c => c.c); picks = preset.picks || []; }
    const body = h('div', { class: 'form m-look' });
    const name = h('input', { class: 'inp' });
    const draw = () => {
      body.innerHTML = '';
      const keyCol = S.cols.find(c => c.c === key);
      const sel = h('select', { class: 'inp' }, ...wb.sheets.map((sh, i) => h('option', { value: i, text: sh.name }))); sel.value = si;
      sel.onchange = () => { si = +sel.value; S = this.table(si); key = (S.cols.find(c => c.kind === 'text') || S.cols[0] || {}).c; rets = S.cols.filter(c => c.c !== key).map(c => c.c); picks = []; draw(); };
      const chip = (t, on, fn) => h('button', { class: 'chip sm' + (on ? ' on' : ''), type: 'button', text: t, onclick: fn });
      body.append(
        h('div', { class: 'm-step' }, h('b', { text: '1' }), h('span', { text: 'Hangi tabloda aransın?' })), sel,
        h('div', { class: 'm-step' }, h('b', { text: '2' }), h('span', { text: 'Hangi sütuna yazacağınız değer aransın?' })),
        h('div', { class: 'vchips' }, ...S.cols.map(c => chip(c.name, c.c === key, () => { key = c.c; rets = S.cols.filter(x => x.c !== key).map(x => x.c); picks = []; draw(); }))),
        h('div', { class: 'm-step' }, h('b', { text: '3' }), h('span', { text: 'Hangi bilgiler gelsin?' })),
        h('div', { class: 'vchips' }, ...S.cols.filter(c => c.c !== key).map(c => chip(c.name, rets.includes(c.c), () => { const i = rets.indexOf(c.c); if (i >= 0) rets.splice(i, 1); else rets.push(c.c); rets.sort((a, b) => a - b); draw(); }))),
        h('div', { class: 'm-step' }, h('b', { text: '4' }), h('span', { text: (keyCol ? keyCol.name : 'Değer') + ' birden çok satırda geçiyorsa' })),
        h('div', { class: 'segs m-modes' }, ...[['first', 'İlk bulunanı getir'], ['sum', 'Sayıların toplamını al'], ['count', 'Kaç tane olduğunu say']].map(([k, t]) => h('button', { class: 'seg' + (mode === k ? ' on' : ''), type: 'button', text: t, onclick: () => { mode = k; draw(); } }))),
        h('div', { class: 'm-step' }, h('b', { text: '5' }), h('span', { text: 'Şimdilik neleri arayalım? (sonra istediğinizi yazarsınız)' })));
      const vals = keyCol ? this.distinct(S, keyCol) : [];
      body.append(h('div', { class: 'vchips' }, chip('Hepsi', vals.length && picks.length === vals.length, () => { picks = picks.length === vals.length ? [] : vals.slice(); draw(); }), ...vals.slice(0, 40).map(v => chip(v, picks.includes(v), () => { const i = picks.indexOf(v); if (i >= 0) picks.splice(i, 1); else picks.push(v); draw(); }))));
      // önizleme
      const cols = this.lookCols(S, key, rets, mode);
      const sample = picks[0] || vals[0];
      if (sample != null && cols.length) {
        const tbl = h('table', { class: 'm-t m-prev' }, h('thead', null, h('tr', null, h('th', { text: keyCol.name }), ...cols.map(c => h('th', { text: c.name })))));
        const tb = h('tbody'); (picks.length ? picks : [sample]).slice(0, 4).forEach(v => { const tr = h('tr', null, h('td', { text: v })); cols.forEach(c => { let x; try { const n = X.parseNumLoose(String(v)); x = X.evaluate(X.parse(c.f(n != null ? String(n) : '"' + String(v).replace(/"/g, '""') + '"')), { model: wb, si, r: 0, c: 0 }); } catch (e) { x = '?'; } tr.append(h('td', { text: X.formatValue(X.isRng(x) ? x.get(0, 0) : x, c.nf || null) })); }); tb.append(tr); });
        tbl.append(tb);
        body.append(h('div', { class: 'm-step' }, h('b', { text: '✓' }), h('span', { text: 'Önizleme' })), h('div', { class: 'm-scrollx m-card' }, tbl));
      }
      if (!name.value || name.dataset.auto) { let nm = 'Arama - ' + (keyCol ? keyCol.name : ''), i = 2; const b = nm; while (wb.sheetIndex(nm) >= 0) nm = b + ' ' + i++; name.value = nm; name.dataset.auto = 1; }
      body.append(h('label', { class: 'lbl', text: 'Yeni tablonun adı' }), name);
    };
    name.oninput = () => { delete name.dataset.auto; };
    draw();
    U.dialog({ title: icon('search') + ' Bul ve getir', body, wide: true, buttons: [{ text: 'Vazgeç', value: null }, { text: 'Oluştur', kind: 'primary', onClick: () => {
      const cols = this.lookCols(S, key, rets, mode); if (!cols.length) { U.toast('Gelecek bilgileri seçin'); return false; }
      const keyCol = S.cols.find(c => c.c === key); let nm = name.value.trim() || 'Arama'; if (wb.sheetIndex(nm) >= 0) { U.toast('Bu adda bir tablo var'); return false; }
      wb.begin(); const sh = wb.addSheet(nm); const ti = wb.sheets.indexOf(sh);
      const hs = { b: 1, bg: '#16a34a', fc: '#ffffff', ha: 'center' };
      wb.patchCell(ti, 0, 0, { v: keyCol.name, f: null }); wb.setStyle(ti, 0, 0, hs);
      cols.forEach((c, j) => { wb.patchCell(ti, 0, j + 1, { v: c.name, f: null }); wb.setStyle(ti, 0, j + 1, hs); });
      const n = Math.max(1, picks.length);
      for (let i = 0; i < n; i++) {
        const R = i + 1; if (picks[i] != null) { const pi = IO.parseInput(String(picks[i])); wb.patchCell(ti, R, 0, { v: pi.v, f: null }); }
        cols.forEach((c, j) => { wb.patchCell(ti, R, j + 1, { f: c.f('$A' + (R + 1)), v: null }); if (c.nf) wb.setStyle(ti, R, j + 1, { nf: c.nf }); });
      }
      // aranacak sütunda açılır liste: kaynak tablodaki değerler
      sh.dv.push({ type: 'list', range: { r1: 1, c1: 0, r2: 500, c2: 0 }, src: '=' + X.quoteSheet(wb.sheets[si].name) + '!$' + X.colName(key) + '$' + (S.r0 + 1) + ':$' + X.colName(key) + '$' + (S.r2 + 500), msg: 'Bu değer “' + wb.sheets[si].name + '” tablosunda yok.' });
      sh.colW = { 0: 140 }; cols.forEach((_, j) => sh.colW[j + 1] = 130); sh.freezeR = 1;
      wb.active = ti; wb.end(); this.view = 'list'; this.render();
      U.toast('✓ Arama tablosu hazır: ilk sütuna yazdığınız değerin bilgileri gelir');
      return true;
    }, value: 1 }] });
  }
  // her getirilecek bilgi için formül üreticisi (anahtar hücresini alır)
  lookCols(S, key, rets, mode) {
    const sh = X.quoteSheet(this.wb.sheets[S.si].name) + '!'; const r1 = S.r0 + 1, r2 = S.r2 + 500;
    const rg = c => sh + '$' + X.colName(c) + '$' + r1 + ':$' + X.colName(c) + '$' + r2;
    if (mode === 'count') return [{ name: 'Kaç tane', nf: null, f: k => `COUNTIF(${rg(key)},${k})` }];
    return rets.map(c => S.cols.find(x => x.c === c)).filter(Boolean).filter(c => mode !== 'sum' || c.kind !== 'text' && c.kind !== 'date').map(col => ({
      name: (mode === 'sum' ? 'Toplam ' : '') + col.name, nf: col.nf || null,
      f: k => mode === 'sum' ? `SUMIF(${rg(key)},${k},${rg(col.c)})`
        : col.c > key ? `IFERROR(VLOOKUP(${k},${sh}$${X.colName(key)}$${r1}:$${X.colName(col.c)}$${r2},${col.c - key + 1},FALSE),"Bulunamadı")`
        : `IFERROR(INDEX(${rg(col.c)},MATCH(${k},${rg(key)},0)),"Bulunamadı")`
    }));
  }
  kindMenu(anchor, T, col) {
    U.menu(anchor, KINDS.map(([k, t, b]) => ({ text: b + '  ' + t, on: k === col.kind, onClick: () => {
      const wb = this.wb; wb.begin(); for (let r = T.r0; r <= Math.max(T.r2, T.r0); r++) wb.setStyle(T.si, r, col.c, { nf: KIND_NF[k] }); wb.end();
    } })));
  }
  focusCol(T, c) { if (c == null) c = 0; const r = T.r0 <= T.r2 ? T.r0 : 0; this.sel = { r, c, r1: r, c1: c, r2: r, c2: c }; }
  addColumn(T) {
    const wb = this.wb, si = T.si;
    const nextC = T.cols.length ? Math.max(...T.cols.map(c => c.c)) + 1 : 0;
    const name = h('input', { class: 'inp', placeholder: 'Ör: Telefon, Fiyat, Teslim tarihi' });
    let kind = 'text';
    const kinds = h('div', { class: 'segs m-kinds' }, ...KINDS.map(([k, t, b]) => h('button', { class: 'seg' + (k === kind ? ' on' : ''), type: 'button', 'data-k': k, onclick: () => { kind = k; U.$$('.seg', kinds).forEach(x => x.classList.toggle('on', x.dataset.k === k)); } }, h('b', { text: b }), h('span', { text: ' ' + t }))));
    const calcBtn = (ic, t, d, fn) => h('button', { class: 'm-opt', type: 'button', onclick: () => { dl.close(); fn(); } }, h('span', { class: 'm-opt-ic', html: icon(ic) }), h('span', null, h('b', { text: t }), h('small', { text: d })));
    const body = h('div', { class: 'form' },
      h('div', { class: 'm-sec2', text: 'Boş sütun (değerleri kendiniz yazarsınız)' }),
      h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Sütun adı' }), name),
      h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Türü' }), kinds),
      h('div', { class: 'm-sec2', text: 'ya da hesaplanan sütun' }),
      calcBtn('wand', 'Hazır hesapla', 'KDV, yüzde, EĞER, tablodan değer getir (DÜŞEYARA)… — her satır için kendiliğinden', () => { this.focusCol(T, T.cols[0] ? T.cols[0].c : 0); this.wiz.openRecipes(); }),
      calcBtn('calc', 'İşlevlerle kendim kurayım', '232 işlev, kutuları doldurarak — her satıra uygulanır', () => { this._colMode = { c: nextC }; const r = T.r0 <= T.r2 ? T.r0 : 1; this.sel = { r, c: nextC, r1: r, c1: nextC, r2: r, c2: nextC }; this.wiz.openFunctions(); }));
    const dl = U.dialog({ title: icon('plus') + ' Sütun ekle', body, buttons: [{ text: 'Vazgeç', value: null }, { text: 'Ekle', kind: 'primary', onClick: () => {
      const nm = name.value.trim(); if (!nm) { U.toast('Sütuna bir ad verin'); name.focus(); return false; }
      wb.begin(); const hr = T.hr >= 0 ? T.hr : 0;
      if (T.hr < 0 && T.cols.length) { wb.shift(si, 'r', 0, 1); }
      const hs = T.cols.length ? (wb.cell(si, hr, T.cols[0].c) || {}).s : { b: 1 };
      wb.patchCell(si, hr, nextC, { v: nm, f: null }); if (hs) wb.setStyle(si, hr, nextC, Object.assign({}, hs, { nf: null }));
      if (KIND_NF[kind]) for (let r = hr + 1; r <= Math.max(T.r2 + (T.hr < 0 && T.cols.length ? 1 : 0), hr + 1); r++) wb.setStyle(si, r, nextC, { nf: KIND_NF[kind] });
      wb.end(); U.toast('✓ “' + nm + '” sütunu eklendi'); return true;
    }, value: 1 }] });
    setTimeout(() => !isMobile() && name.focus(), 60);
  }
  filterDialog(T, col) {
    col = col || T.cols[0]; if (!col) return; const vs = this.vsCur;
    const cs = h('select', { class: 'inp' }, ...T.cols.map(c => h('option', { value: c.c, text: c.name })));
    cs.value = col.c;
    const op = h('select', { class: 'inp' }, ...OPS.map(([v, t]) => h('option', { value: v, text: t })));
    const val = h('input', { class: 'inp', placeholder: 'Değer' });
    const chips = h('div', { class: 'vchips' });
    const fill = () => { chips.innerHTML = ''; const c = T.cols.find(x => x.c === +cs.value); const list = this.distinct(T, c); if (c.kind === 'text') list.slice(0, 14).forEach(v => chips.append(h('button', { class: 'chip sm', type: 'button', text: v, onclick: () => { val.value = v; U.$$('.chip', chips).forEach(x => x.classList.toggle('on', x.textContent === v)); } }))); val.style.display = op.value === 'empty' || op.value === 'filled' ? 'none' : ''; };
    cs.onchange = fill; op.onchange = fill; fill();
    U.dialog({ title: icon('filter') + ' Filtre', body: h('div', { class: 'form' }, h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Hangi sütun?' }), cs), h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Koşul' }), op), h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Değer' }), val, chips)),
      buttons: [{ text: 'Vazgeç', value: null }, { text: 'Uygula', kind: 'primary', value: 1, onClick: () => { vs.filters.push({ c: +cs.value, op: op.value, val: val.value }); this.render(); } }] });
  }

  // ---------- Özet ----------
  viewSummary(T) {
    const wb = this.wb, si = T.si, vs = this.vsCur; const rows = this.rowsOf(T);
    const textCols = T.cols.filter(c => c.kind === 'text' || c.kind === 'date'), numCols = T.cols.filter(c => c.kind !== 'text' && c.kind !== 'date');
    if (vs.sumG == null || !T.cols.find(c => c.c === vs.sumG)) vs.sumG = (textCols.find(c => c.distinctN !== 0 && this.distinct(T, c).length < rows.length) || textCols[0] || T.cols[0]).c;
    if (vs.sumV == null || !T.cols.find(c => c.c === vs.sumV)) vs.sumV = numCols.length ? numCols[numCols.length - 1].c : -1;
    vs.sumOp = vs.sumOp || 'sum';
    const sel = (opts, v, fn) => { const s = h('select', { class: 'inp' }, ...opts.map(([k, t]) => h('option', { value: k, text: t }))); s.value = v; s.onchange = () => fn(s.value); return s; };
    const ctrl = h('div', { class: 'm-sumctl' },
      h('label', null, h('small', { text: 'Gruplara ayır' }), sel(T.cols.map(c => [c.c, c.name]), vs.sumG, v => { vs.sumG = +v; this.render(); })),
      h('label', null, h('small', { text: 'Ne yapılsın?' }), sel(AGGS.filter(a => a[0] !== 'distinct'), vs.sumOp, v => { vs.sumOp = v; this.render(); })),
      vs.sumOp === 'count' ? null : h('label', null, h('small', { text: 'Hangi değer?' }), sel(numCols.map(c => [c.c, c.name]), vs.sumV, v => { vs.sumV = +v; this.render(); })));
    const gcol = T.cols.find(c => c.c === vs.sumG), vcol = T.cols.find(c => c.c === vs.sumV);
    const groups = new Map();
    for (const r of rows) { const k = wb.display(si, r, gcol.c) || '(boş)'; const key = lc(k); if (!groups.has(key)) groups.set(key, { name: k, rows: [] }); groups.get(key).rows.push(r); }
    const op = vs.sumOp, opName = AGGS.find(a => a[0] === op)[1];
    const list = [...groups.values()].map(g => ({ name: g.name, n: g.rows.length, v: op === 'count' || !vcol ? g.rows.length : this.agg(T, vcol, g.rows, op) }));
    list.sort((a, b) => (b.v || 0) - (a.v || 0));
    const fmt = v => v == null ? '—' : op === 'count' || !vcol ? String(v) : X.formatValue(v, vcol.nf || null);
    const max = Math.max(...list.map(x => Math.abs(x.v || 0)), 1); const total = list.reduce((a, b) => a + (b.v || 0), 0);
    const tbl = h('div', { class: 'm-sumlist' });
    list.forEach(x => tbl.append(h('div', { class: 'm-sumrow' }, h('span', { class: 'm-sum-n', text: x.name }), h('small', { text: x.n + ' kayıt' }), h('div', { class: 'm-sum-bar' }, h('i', { style: { width: (Math.abs(x.v || 0) / max * 100).toFixed(1) + '%' } })), h('b', { text: fmt(x.v) }))));
    if (op === 'sum' || op === 'count') tbl.append(h('div', { class: 'm-sumrow total' }, h('span', { class: 'm-sum-n', text: 'Genel toplam' }), h('small', { text: rows.length + ' kayıt' }), h('div'), h('b', { text: fmt(total) })));
    const title = gcol.name + ' bazında ' + lc(opName) + (op !== 'count' && vcol ? ' — ' + vcol.name : '');
    const addBtn = h('button', { class: 'btn', html: icon('table') + '<span>Dosyaya özet tablo olarak ekle</span>', title: 'Excel dosyasında da görünür ve veri değişince güncellenir', onclick: () => this.summaryToSheet(T, gcol, vcol, op) });
    return h('div', { class: 'm-sum' }, h('div', { class: 'm-card' }, ctrl), h('div', { class: 'm-card' }, h('div', { class: 'm-cardhead' }, h('b', { text: title }), h('div', { class: 'spacer' }), addBtn), tbl));
  }
  summaryToSheet(T, gcol, vcol, op) {
    const wb = this.wb, src = wb.sheets[T.si]; const sh = X.quoteSheet(src.name) + '!';
    const rg = c => sh + '$' + X.colName(c) + '$' + (T.r0 + 1) + ':$' + X.colName(c) + '$' + (T.r2 + 1);
    const g = rg(gcol.c), v = vcol ? rg(vcol.c) : g, u = 'UNIQUE(' + g + ')';
    const f = op === 'sum' ? `SUMIF(${g},${u},${v})` : op === 'count' ? `COUNTIF(${g},${u})` : op === 'avg' ? `AVERAGEIF(${g},${u},${v})` : op === 'max' ? `MAXIFS(${v},${g},${u})` : `MINIFS(${v},${g},${u})`;
    let name = 'Özet - ' + gcol.name, i = 2; while (wb.sheetIndex(name) >= 0) name = 'Özet - ' + gcol.name + ' ' + i++;
    const nm = AGGS.find(a => a[0] === op)[1] + (op === 'count' || !vcol ? '' : ' ' + vcol.name);
    wb.begin(); const s = wb.addSheet(name); const si = wb.sheets.indexOf(s);
    const hs = { b: 1, bg: '#16a34a', fc: '#ffffff', ha: 'center' };
    wb.patchCell(si, 0, 0, { v: gcol.name, f: null }); wb.setStyle(si, 0, 0, hs);
    wb.patchCell(si, 0, 1, { v: nm, f: null }); wb.setStyle(si, 0, 1, hs);
    wb.patchCell(si, 1, 0, { f: u, v: null }); wb.patchCell(si, 1, 1, { f, v: null });
    if (op !== 'count' && vcol && vcol.nf) wb.setStyle(si, 1, 1, { nf: vcol.nf });
    s.colW = { 0: 140, 1: 150 }; wb.end();
    this.view = 'list'; this.render(); U.toast('✓ “' + name + '” tablosu eklendi (kendiliğinden güncellenir)');
  }

  // ---------- Grafik ----------
  chartTheme() { return document.documentElement.classList.contains('dark') ? { bg: '#19191c', fg: '#e4e4e7', muted: '#a1a1aa', grid: '#3f3f46' } : null; }
  viewChart(T) {
    const wb = this.wb, si = T.si, sh = wb.sheets[si]; const theme = this.chartTheme();
    const wrap = h('div', { class: 'm-sum' });
    wrap.append(h('div', { class: 'm-card m-chartbar' }, h('div', null, h('b', { text: 'Grafiklerim' }), h('small', { text: sh.charts.length ? sh.charts.length + ' grafik · Excel dosyasına da kaydedilir' : 'Neyi görmek istediğinizi seçerek grafik oluşturun' })),
      h('div', { class: 'spacer' }), h('button', { class: 'btn primary', html: icon('plus') + '<span>Yeni grafik</span>', onclick: () => this.chartWizard(T) })));
    if (!sh.charts.length) {
      wrap.append(h('button', { class: 'm-rcard add m-chartadd', onclick: () => this.chartWizard(T) }, h('span', { html: icon('chart') }), h('span', { text: 'İlk grafiğinizi oluşturun: ne gösterilsin, neye göre, hangi kayıtlar, tahmin…' })));
      return wrap;
    }
    const gal = h('div', { class: 'm-gallery' });
    sh.charts.forEach(spec => {
      const cv = h('canvas', { class: 'm-chartcv sm' });
      requestAnimationFrame(() => { try { C.draw(cv, spec, C.extract(wb, si, spec), theme); } catch (e) { console.warn(e); } });
      const del = async () => { if (await U.confirm('“' + (spec.title || 'Grafik') + '” silinsin mi?', 'Grafiği sil', 'Sil')) { wb.begin(); wb.snapSheet(si); sh.charts.splice(sh.charts.indexOf(spec), 1); wb.end(); } };
      gal.append(h('div', { class: 'm-card' }, h('div', { class: 'm-cardhead' }, h('b', { text: spec.title || 'Grafik' }), h('div', { class: 'spacer' }),
        spec.q ? h('button', { class: 'ic ghost', html: icon('rules'), title: 'Düzenle', 'aria-label': 'Düzenle', onclick: () => this.chartWizard(T, spec) }) : null,
        h('button', { class: 'ic ghost', html: icon('trash'), title: 'Sil', 'aria-label': 'Sil', onclick: del })), cv));
    });
    wrap.append(gal);
    return wrap;
  }
  // Yeni grafik: Bul ve getir gibi adım adım seç, canlı önizle, tahmin ekle
  chartWizard(T, edit, preset) {
    const wb = this.wb, si = T.si, sh = wb.sheets[si]; const theme = this.chartTheme();
    const textCols = T.cols.filter(c => c.kind === 'text' || c.kind === 'date'), numCols = T.cols.filter(c => c.kind !== 'text' && c.kind !== 'date');
    const q = edit ? JSON.parse(JSON.stringify(edit.q)) : preset ? Object.assign({ conds: [], sort: 'none', top: 0, fc: 0, trend: false }, preset) : { lab: (T.cols.find(c => c.kind === 'text') || textCols[0] || T.cols[0]).c, labMode: 'value', vals: numCols.length ? [{ c: numCols[numCols.length - 1].c, op: 'sum' }] : [{ c: -1, op: 'count' }], conds: [], sort: 'none', top: 0, fc: 0, trend: false };
    let type = edit ? edit.type : 'column', title = edit ? edit.title : '', autoTitle = !edit;
    const body = h('div', { class: 'form m-look m-cw' });
    const cv = h('canvas', { class: 'm-chartcv' });
    const fcBox = h('div', { class: 'm-fcbox' });
    const chip = (t, on, fn, cls) => h('button', { class: 'chip sm' + (on ? ' on' : '') + (cls ? ' ' + cls : ''), type: 'button', text: t, onclick: fn });
    const step = (n, t) => h('div', { class: 'm-step' }, h('b', { text: n }), h('span', { text: t }));
    const colName = c => (T.cols.find(x => x.c === c) || {}).name || '?';
    const makeTitle = () => q.vals.map(v => v.c < 0 ? 'Kayıt sayısı' : C.AGG[v.op] + ' ' + colName(v.c)).join(', ') + ' — ' + colName(q.lab) + (q.conds.length ? ' (' + q.conds.map(f => colName(f.c) + ' ' + OPS.find(o => o[0] === f.op)[1] + (f.op === 'empty' || f.op === 'filled' ? '' : ' ' + f.val)).join(', ') + ')' : '');
    const preview = () => {
      if (autoTitle) { title = makeTitle(); const ti = body.querySelector('.m-cw-title'); if (ti) ti.value = title; }
      let data; try { data = C.query(wb, si, q); } catch (e) { data = { labels: [], series: [] }; }
      requestAnimationFrame(() => C.draw(cv, { type, title, labels: true }, data, theme));
      fcBox.innerHTML = '';
      if (data.forecast) {
        const f = data.forecast; const fm = v => X.formatValue(v, f.fmt || '#,##0.##');
        fcBox.append(h('div', { class: 'm-fc' }, h('b', { text: 'Tahmin (TAHMİN.DOĞRUSAL)' }), h('small', { text: 'Her dönem ortalama ' + (f.slope >= 0 ? '+' : '') + fm(f.slope) + ' değişiyor.' }),
          h('div', { class: 'm-fc-list' }, ...f.items.map(it => h('span', null, h('small', { text: it.t }), h('b', { text: fm(it.v) }))))));
      } else if (q.fc > 0) fcBox.append(h('p', { class: 'hint', text: 'Tahmin için en az 2 dönem gerekir.' }));
    };
    const draw = () => {
      body.innerHTML = '';
      const labCol = T.cols.find(c => c.c === q.lab);
      // 1. tür
      body.append(step('1', 'Grafik türü'), h('div', { class: 'vchips' }, ...C.TYPES.filter(t => t[0] !== 'scatter').map(([k, t]) => chip(t, type === k, () => { type = k; draw(); }))));
      // 2. neye göre
      body.append(step('2', 'Neye göre? (grafiğin altındaki etiketler)'), h('div', { class: 'vchips' }, ...T.cols.map(c => chip(c.name, c.c === q.lab, () => { q.lab = c.c; q.labMode = c.kind === 'date' ? 'month' : 'value'; draw(); }))));
      if (labCol && labCol.kind === 'date') body.append(h('div', { class: 'segs m-modes' }, ...[['value', 'Gün gün'], ['month', 'Aylara göre'], ['year', 'Yıllara göre']].map(([k, t]) => h('button', { class: 'seg' + (q.labMode === k ? ' on' : ''), type: 'button', text: t, onclick: () => { q.labMode = k; draw(); } }))));
      // 3. ne gösterilsin
      body.append(step('3', 'Ne gösterilsin? (birden çok seçebilirsiniz)'));
      const isOn = c => q.vals.find(v => v.c === c);
      body.append(h('div', { class: 'vchips' }, chip('Kayıt sayısı', !!isOn(-1), () => { const v = isOn(-1); if (v) { if (q.vals.length > 1) q.vals.splice(q.vals.indexOf(v), 1); } else q.vals.push({ c: -1, op: 'count' }); draw(); }),
        ...T.cols.filter(c => c.c !== q.lab && c.kind !== 'text' && c.kind !== 'date').map(c => chip(c.name, !!isOn(c.c), () => { const v = isOn(c.c); if (v) { if (q.vals.length > 1) q.vals.splice(q.vals.indexOf(v), 1); } else q.vals.push({ c: c.c, op: 'sum' }); draw(); }))));
      q.vals.filter(v => v.c >= 0).forEach(v => body.append(h('div', { class: 'm-vrow m-oprow' }, h('span', { class: 'm-pill', text: colName(v.c) }), h('div', { class: 'segs' }, ...[['sum', 'Topla'], ['avg', 'Ortalama'], ['max', 'En büyük'], ['min', 'En küçük'], ['count', 'Say']].map(([k, t]) => h('button', { class: 'seg' + (v.op === k ? ' on' : ''), type: 'button', text: t, onclick: () => { v.op = k; draw(); } }))))));
      // 4. koşullar
      body.append(step('4', 'Hangi kayıtlar? (EĞER gibi koşul — boş bırakılırsa hepsi)'));
      q.conds.forEach((f, i) => body.append(h('div', { class: 'm-rule' }, h('span', { class: 'm-rsw m-cond-ic', text: 'EĞER' }), h('span', { text: colName(f.c) + ' ' + OPS.find(o => o[0] === f.op)[1] + (f.op === 'empty' || f.op === 'filled' ? '' : ' ' + f.val) }), h('button', { class: 'ic ghost', type: 'button', html: icon('trash'), title: 'Koşulu kaldır', onclick: () => { q.conds.splice(i, 1); draw(); } }))));
      const cs = h('select', { class: 'inp' }, ...T.cols.map(c => h('option', { value: c.c, text: c.name })));
      const op = h('select', { class: 'inp' }, ...OPS.map(([v, t]) => h('option', { value: v, text: t })));
      const val = h('input', { class: 'inp', placeholder: 'Değer' });
      const vch = h('div', { class: 'vchips' });
      const fillV = () => { vch.innerHTML = ''; const c = T.cols.find(x => x.c === +cs.value); if (c && c.kind === 'text') this.distinct(T, c).slice(0, 14).forEach(v => vch.append(chip(v, val.value === v, () => { val.value = v; U.$$('.chip', vch).forEach(x => x.classList.toggle('on', x.textContent === v)); }))); val.style.display = op.value === 'empty' || op.value === 'filled' ? 'none' : ''; op.value = op.value || '='; };
      cs.onchange = fillV; op.onchange = fillV; fillV();
      body.append(h('div', { class: 'm-cond' }, cs, op, val, h('button', { class: 'btn', type: 'button', html: icon('plus') + '<span>Ekle</span>', onclick: () => { if (!/^(empty|filled)$/.test(op.value) && !val.value.trim()) { U.toast('Bir değer yazın ya da seçin'); return; } q.conds.push({ c: +cs.value, op: op.value, val: val.value.trim() }); draw(); } })), vch);
      // 5. sıralama
      body.append(step('5', 'Sıralama'), h('div', { class: 'segs m-modes' }, ...[['none', 'Olduğu gibi'], ['desc', 'Büyükten küçüğe'], ['asc', 'Küçükten büyüğe'], ['label', 'Ada / tarihe göre']].map(([k, t]) => h('button', { class: 'seg' + (q.sort === k ? ' on' : ''), type: 'button', text: t, onclick: () => { q.sort = k; draw(); } }))),
        h('div', { class: 'vchips' }, ...[[0, 'Hepsi'], [3, 'İlk 3'], [5, 'İlk 5'], [10, 'İlk 10']].map(([n, t]) => chip(t, (q.top || 0) === n, () => { q.top = n; draw(); }))));
      // 6. tahmin
      const pie = type === 'pie' || type === 'doughnut';
      body.append(step('6', 'Tahmin (Excel TAHMİN.DOĞRUSAL)'));
      if (pie) body.append(h('p', { class: 'hint', text: 'Pasta ve halka grafikte tahmin yapılmaz; çizgi ya da sütun seçin.' }));
      else body.append(h('div', { class: 'vchips' }, ...[[0, 'Tahmin yok'], [1, 'Sonraki 1 dönem'], [3, 'Sonraki 3 dönem'], [6, 'Sonraki 6 dönem']].map(([n, t]) => chip(t, (q.fc || 0) === n, () => { q.fc = n; if (n && q.sort !== 'label' && q.sort !== 'none') q.sort = 'label'; draw(); }))),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: !!q.trend, onchange: e => { q.trend = e.target.checked; preview(); } }), h('span', { text: 'Eğilim çizgisini göster (genel gidişat)' })));
      if (pie) { q.fc = 0; q.trend = false; }
      // başlık + önizleme
      const ti = h('input', { class: 'inp m-cw-title', value: title, oninput: () => { title = ti.value; autoTitle = false; preview(); } });
      body.append(h('label', { class: 'lbl', text: 'Grafiğin adı' }), ti, step('✓', 'Önizleme'), h('div', { class: 'm-card m-cw-prev' }, cv), fcBox);
      preview();
    };
    draw();
    U.dialog({ title: icon('chart') + (edit ? ' Grafiği düzenle' : ' Yeni grafik'), body, wide: true, cls: 'm-cw-dlg', buttons: [{ text: 'Vazgeç', value: null }, { text: edit ? 'Kaydet' : 'Grafiği ekle', kind: 'primary', value: 1, onClick: () => {
      wb.begin(); wb.snapSheet(si);
      if (edit) { const s2 = sh.charts.find(x => x.id === edit.id); Object.assign(s2, { type, title, q }); }
      else { let x = 0; const d = wb.dims(si); for (let c = 0; c <= d.cols; c++) x += sh.colW[c] || 90; sh.charts.push({ id: 'c' + Date.now().toString(36), type, title, labels: true, q, range: { r1: 0, c1: q.lab, r2: Math.max(T.r2, 0), c2: q.lab }, x: x + 16, y: sh.charts.length * 320, w: 520, h: 300 }); }
      wb.end(); this.view = 'chart'; this.render(); U.toast(edit ? '✓ Grafik güncellendi' : '✓ Grafik eklendi');
    } }] });
  }

  // ---------- tablolar ----------
  sheetMenu(anchor, i) {
    const wb = this.wb;
    U.menu(anchor, [
      { text: 'Adını değiştir', icon: 'tag', onClick: async () => { const v = await U.prompt('Tablo adı', wb.sheets[i].name); if (v && v.trim() && wb.sheetIndex(v.trim()) < 0) wb.renameSheet(i, v.trim()); } },
      { text: 'Tabloyu sil', icon: 'trash', danger: 1, onClick: () => this.deleteTable(i) }
    ]);
  }
  async deleteTable(i) {
    const wb = this.wb, sh = wb.sheets[i]; const n = Math.max(0, this.table(i).r2 - this.table(i).r0 + 1);
    const ok = await U.confirm('“' + sh.name + '” tablosu ve içindeki ' + n + ' kayıt silinsin mi?' + (wb.sheets.length < 2 ? ' Yerine boş bir tablo açılır.' : '') + ' Yanlışlıkla silerseniz sağ üstteki geri al düğmesiyle geri getirebilirsiniz.', 'Tabloyu sil', 'Sil');
    if (!ok) return;
    wb.begin();
    if (wb.sheets.length < 2) { let nm = 'Tablo 1', k = 2; while (wb.sheetIndex(nm) >= 0) nm = 'Tablo ' + k++; wb.addSheet(nm); }
    wb.removeSheet(i); if (wb.active >= wb.sheets.length) wb.active = wb.sheets.length - 1;
    wb.end(); delete this.vs[i]; this.render(); U.toast('✓ Tablo silindi');
  }
  async newTable() {
    const v = await U.prompt('Yeni tablo', '', 'Tablonun adı (Ör: Müşteriler, Giderler)'); if (v == null) return;
    const wb = this.wb; let name = v.trim() || 'Tablo'; let i = 2; const base = name; while (wb.sheetIndex(name) >= 0) name = base + ' ' + i++;
    wb.addSheet(name); this.render(); setTimeout(() => this.addColumn(this.T), 80);
  }

  // ---------- Hazır Hesaplar için gerekenler (klasik görünümle aynı arayüz) ----------
  setSel(s) { this.sel = s; this.render(); }
  refText(s, selSi, baseSi) { let t = s.r1 === s.r2 && s.c1 === s.c2 ? X.addr(s.r1, s.c1) : X.addr(s.r1, s.c1) + ':' + X.addr(s.r2, s.c2); if (selSi != null && baseSi != null && selSi !== baseSi) t = X.quoteSheet(this.wb.sheets[selSi].name) + '!' + t; return t; }
  currentRegion(r, c) {
    const wb = this.wb, si = wb.active; const has = (rr, cc) => rr >= 0 && cc >= 0 && wb.value(si, rr, cc) != null && wb.value(si, rr, cc) !== '';
    let g = { r1: r, c1: c, r2: r, c2: c }; let ch = true, k = 0;
    while (ch && k++ < 200) {
      ch = false;
      const rowHas = rr => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (has(rr, cc)) return true; return false; };
      const colHas = cc => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (has(rr, cc)) return true; return false; };
      if (g.r1 > 0 && rowHas(g.r1 - 1)) { g.r1--; ch = true; } if (rowHas(g.r2 + 1)) { g.r2++; ch = true; }
      if (g.c1 > 0 && colHas(g.c1 - 1)) { g.c1--; ch = true; } if (colHas(g.c2 + 1)) { g.c2++; ch = true; }
    }
    return g;
  }
  detectHeader(g) {
    const wb = this.wb, si = wb.active; if (g.r2 <= g.r1) return false;
    let txt = 0, nonTxt2 = 0;
    for (let c = g.c1; c <= g.c2; c++) { const a = wb.value(si, g.r1, c), b = wb.value(si, g.r1 + 1, c); if (typeof a === 'string') txt++; if (b != null && typeof b !== 'string') nonTxt2++; }
    const st = (wb.cell(si, g.r1, g.c1) || {}).s;
    return txt === g.c2 - g.c1 + 1 && (nonTxt2 > 0 || (st && st.b));
  }
  evalFormula(f, t) {
    try { const ast = X.parse(f); let v = X.evaluate(ast, { model: this.wb, si: t.si, r: t.r, c: t.c }); if (X.isRng(v)) v = v.h && v.w ? v.get(0, 0) : X.E('#VALUE!'); return v == null ? 0 : v; }
    catch (e) { return X.isErr(e) ? e : X.E('#NAME?'); }
  }
  applyFormula(si, r, c, f, n, nf, hdr) {
    const wb = this.wb;
    if (this._calcName && this.fnTarget && wb.sheets[si] && wb.sheets[si].name === 'Hesaplar') {
      const m = /^([A-Z][A-Z0-9.]*)\(/.exec(f); const fd = m && X.FN[m[1]];
      wb.begin(); wb.patchCell(si, r, 0, { v: (fd ? fd.tr + ' — ' : '') + this._calcName, f: null }); wb.patchCell(si, r, 1, { f, v: null }); if (nf) wb.setStyle(si, r, 1, { nf }); wb.end();
      this.fnTarget = null; this._calcName = null; wb.active = si; this.view = 'list'; this.msel = null; this.render(); U.toast('✓ Hesaplar tablosuna eklendi: ' + wb.display(si, r, 1)); return;
    }
    // "işlevlerle kendim kurayım": tek hücre yerine bütün satırlara uygula ve başlık sor
    const cm = this._colMode; this._colMode = null; let askName = false;
    if (cm && cm.c === c && (n || 1) === 1 && !hdr) {
      const T = this.table(si);
      if (T.r2 >= r) {
        // tablo sütununun tamamı (D2:D9) verilmişse her satırda kendi hücresini (D2) kullan, sonra aşağı doldur
        const a = T.r0 + 1, b = T.r2 + 1, rows = T.r2 - r + 1;
        let v = null; try { v = X.evaluate(X.parse(f), { model: wb, si, r, c }); } catch (e) { }
        if (X.isRng(v) && v.h === rows && v.w === 1) { f = f.replace(new RegExp("(^|[^A-Za-z0-9_!$'])\\$?([A-Z]{1,3})\\$?" + a + ':\\$?\\2\\$?' + b + '(?![0-9])', 'g'), (m, p, col) => p + col + a); n = rows; }
        else if (!/\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+/.test(f)) n = rows;
        if (T.hr >= 0 && T.hr < r) { hdr = { r: T.hr, c, text: 'Hesap', s: (wb.cell(si, T.hr, T.cols[0].c) || {}).s }; askName = true; }
      }
    }
    wb.begin();
    if (hdr) { wb.patchCell(si, hdr.r, hdr.c, { v: hdr.text, f: null }); if (hdr.s) wb.setStyle(si, hdr.r, hdr.c, Object.assign({}, hdr.s, { nf: null })); }
    for (let i = 0; i < (n || 1); i++) {
      wb.patchCell(si, r + i, c, { f: i ? X.shiftFormula(f, i, 0) : f, v: null });
      const st = (wb.cell(si, r + i, c) || {}).s;
      if (nf && !(st && st.nf)) wb.setStyle(si, r + i, c, { nf });
    }
    wb.end();
    if (wb.active !== si) wb.active = si;
    this.sel = { r, c, r1: r, c1: c, r2: r + (n || 1) - 1, c2: c }; this.view = this.view === 'chart' || this.view === 'summary' ? 'list' : this.view; this.render();
    if (askName) setTimeout(async () => { const v = await U.prompt('Yeni sütunun adı', 'Hesap'); if (v && v.trim()) { wb.begin(); wb.patchCell(si, hdr.r, c, { v: v.trim(), f: null }); wb.end(); } }, 200);
  }

  // ---------- dosya ----------
  newDoc() { this.saveLocal(); const wb = new M.Workbook(); wb.name = 'Yeni Çalışma'; wb.sheets[0].name = 'Tablo 1'; this.docId = this.newId(); LS.set('cur', this.docId); this.vs = {}; this.attach(wb); this.saveLocal(); setTimeout(() => this.addColumn(this.T), 120); }
  async openFromUrl(u) {
    const busy = this.busy('Açılıyor…');
    try { const blob = await new Promise((res, rej) => { const x = new XMLHttpRequest(); x.open('GET', u); x.responseType = 'blob'; x.onload = () => (x.status === 0 || x.status === 200) ? res(x.response) : rej(new Error(x.status)); x.onerror = () => rej(new Error('okunamadı')); x.send(); }); const name = decodeURIComponent(u.split(/[\\/]/).pop()); const wb = await IO.open(new File([blob], name)); this.docId = this.newId(); LS.set('cur', this.docId); this.vs = {}; this.attach(wb); this.saveLocal(); U.toast('✓ ' + name + ' açıldı'); }
    catch (e) { U.alert('Dosya açılamadı: ' + u, 'Hata'); } finally { busy(); }
  }
  openFile() {
    const inp = h('input', { type: 'file', accept: IO.OPEN_ACCEPT, style: { display: 'none' } });
    inp.onchange = async () => { const f = inp.files[0]; inp.remove(); if (!f) return; const busy = this.busy('Açılıyor: ' + f.name); try { await this.saveLocal(); const wb = await IO.open(f); this.docId = this.newId(); LS.set('cur', this.docId); this.vs = {}; this.attach(wb); this.saveLocal(); U.toast('✓ ' + f.name + ' açıldı'); } catch (e) { console.error(e); U.alert('Dosya açılamadı: ' + (e.message || e), 'Hata'); } finally { busy(); } };
    document.body.append(inp); inp.click();
  }
  saveMenu() {
    const items = ['xlsx', 'xltx', 'xls', 'ods', 'csv'].map(k => ({ text: IO.FORMATS[k].name + ' (.' + IO.FORMATS[k].ext + ')', icon: 'file', onClick: () => this.saveAs(k) }));
    items.push({ sep: 1 }, { text: 'PDF / Yazdır', icon: 'print', onClick: () => IO.print(this.wb, this.wb.active) });
    U.menu($('.m-head .m-save') || this.head, items);
  }
  async saveAs(fmt) {
    const F = IO.FORMATS[fmt]; const name = (this.wb.name || 'Excello') + '.' + F.ext; let handle = null;
    if (G.showSaveFilePicker && !isIOS) { try { handle = await G.showSaveFilePicker({ suggestedName: name, types: [{ description: F.name, accept: { [F.mime]: ['.' + F.ext] } }] }); } catch (e) { if (e.name === 'AbortError') return; handle = null; } }
    const busy = this.busy('Hazırlanıyor…'); let res;
    try { res = await IO.save(this.wb, fmt); } catch (e) { busy(); console.error(e); U.alert('Kaydedilemedi: ' + (e.message || e), 'Hata'); return; }
    busy();
    if (handle) { try { const w = await handle.createWritable(); await w.write(res.blob); await w.close(); this.wb.name = handle.name.replace(/\.[^.]+$/, ''); this.saveLocal(); this.render(); U.toast('✓ Kaydedildi: ' + handle.name); return; } catch (e) { console.warn(e); } }
    const file = new File([res.blob], res.name, { type: res.mime });
    const download = () => { const url = URL.createObjectURL(res.blob); const a = h('a', { href: url, download: res.name }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000); };
    if (navigator.canShare && navigator.canShare({ files: [file] }) && isMobile()) {
      U.dialog({ title: icon('check') + ' Dosya hazır', body: h('div', { class: 'form' }, h('p', { text: res.name }), h('p', { class: 'hint', text: '“Paylaş” ile Dosyalar uygulamasına kaydedebilir, WhatsApp veya e-posta ile gönderebilirsiniz.' })), buttons: [{ text: 'İndir', onClick: download, value: 1 }, { text: 'Paylaş / Kaydet', kind: 'primary', onClick: () => { navigator.share({ files: [file], title: res.name }).catch(() => { }); }, value: 1 }] });
      return;
    }
    download(); U.toast('✓ İndirildi: ' + res.name);
  }
  async recent() {
    let list = []; try { list = (await DB.all()) || []; } catch (e) { }
    list.sort((a, b) => b.updated - a.updated);
    const body = h('div', { class: 'm-recent' });
    if (!list.length) body.append(h('p', { class: 'hint', text: 'Bu cihazda kayıtlı çalışma yok.' }));
    list.slice(0, 50).forEach(doc => {
      const cur = doc.id === this.docId;
      const row = h('div', { class: 'm-recrow' },
        h('button', { class: 'm-opt', onclick: () => { d.close(); if (cur) return; this.saveLocal(); const wb = M.Workbook.fromJSON(doc.data); wb.format = doc.format || 'xlsx'; this.docId = doc.id; LS.set('cur', doc.id); this.vs = {}; this.attach(wb); } }, h('span', { class: 'm-opt-ic', html: icon('file') }), h('span', null, h('b', { text: (doc.name || 'Çalışma') + (cur ? ' (açık)' : '') }), h('small', { text: new Date(doc.updated).toLocaleString('tr-TR') }))),
        h('button', { class: 'ic m-recdel', html: icon('trash'), title: 'Çalışmayı sil', 'aria-label': 'Çalışmayı sil', onclick: async () => {
          if (!await U.confirm('“' + (doc.name || 'Çalışma') + '” çalışması bu cihazdan tamamen silinsin mi? Bu geri alınamaz. (Daha önce Excel olarak kaydettiğiniz dosyalar silinmez.)', 'Çalışmayı sil', 'Sil')) return;
          try { await DB.del(doc.id); } catch (e) { }
          row.remove(); U.toast('✓ Silindi');
          if (cur) { const wb = new M.Workbook(); wb.name = 'Yeni Çalışma'; wb.sheets[0].name = 'Tablo 1'; this.docId = this.newId(); LS.set('cur', this.docId); this.vs = {}; this.attach(wb); this.saveLocal(); }
        } }));
      body.append(row);
    });
    const d = U.dialog({ title: icon('layers') + ' Son çalışmalar', body, buttons: [{ text: 'Kapat', value: null }] });
  }
  busy(msg) { const el = h('div', { class: 'busy' }, h('div', { class: 'spin' }), h('span', { text: msg })); document.body.append(el); return () => el.remove(); }
  scheduleSave() { clearTimeout(this._saveT); this._saveT = setTimeout(() => this.saveLocal(), 1000); }
  async saveLocal() {
    clearTimeout(this._saveT); if (!this.docId) return;
    try { await DB.put({ id: this.docId, name: this.wb.name || 'Çalışma', updated: Date.now(), data: this.wb.toJSON(), format: this.wb.format || 'xlsx' }); LS.set('cur', this.docId); }
    catch (e) { }
  }
  sample() {
    // klasik görünümdeki örnekle aynı veriler
    const wb = new M.Workbook(); wb.name = 'Örnek Satış Tablosu'; const s = wb.sheets[0]; s.name = 'Satışlar';
    const rows = [['Tarih', 'Ürün', 'Şehir', 'Adet', 'Birim Fiyat', 'Tutar'], ['04.06.2026', 'Defter', 'Ankara', 12, 45], ['18.06.2026', 'Kalem', 'İzmir', 40, 12.5], ['06.07.2026', 'Çanta', 'Ankara', 3, 650], ['21.07.2026', 'Defter', 'İstanbul', 25, 45], ['05.08.2026', 'Silgi', 'İzmir', 60, 6], ['19.08.2026', 'Çanta', 'İstanbul', 5, 650], ['08.09.2026', 'Kalem', 'Ankara', 30, 12.5], ['24.09.2026', 'Defter', 'İzmir', 18, 45], ['07.10.2026', 'Çanta', 'İzmir', 4, 650], ['20.10.2026', 'Defter', 'Ankara', 30, 45]];
    rows.forEach((row, r) => row.forEach((v, c) => { const pi = typeof v === 'string' ? IO.parseInput(v) : { v }; s.cells.set(M.key(r, c), { v: pi.v, f: null, s: pi.nf ? { nf: pi.nf } : null }); }));
    for (let r = 1; r < rows.length; r++) { s.cells.set(M.key(r, 5), { v: null, f: `D${r + 1}*E${r + 1}`, s: { nf: '#,##0.00 ₺' } }); s.cells.get(M.key(r, 4)).s = { nf: '#,##0.00 ₺' }; }
    const H = '#16a34a';
    for (let c = 0; c < 6; c++) s.cells.get(M.key(0, c)).s = { b: 1, bg: H, fc: '#ffffff', ha: 'center' };
    s.colW = { 0: 100, 1: 96, 2: 96, 3: 70, 4: 110, 5: 120 }; s.freezeR = 1;
    const u = new M.Sheet('Ürünler'); wb.sheets.push(u);
    [['Ürün', 'Kategori', 'Stok', 'Tedarikçi'], ['Defter', 'Kırtasiye', 120, 'Ada Kağıt'], ['Kalem', 'Kırtasiye', 300, 'Mavi Yazı'], ['Çanta', 'Aksesuar', 15, 'Deri Dünyası'], ['Silgi', 'Kırtasiye', 200, 'Mavi Yazı']]
      .forEach((row, r) => row.forEach((v, c) => u.cells.set(M.key(r, c), { v, f: null, s: r ? null : { b: 1, bg: H, fc: '#ffffff', ha: 'center' } })));
    u.colW = { 0: 100, 1: 110, 2: 80, 3: 130 }; u.freezeR = 1; wb.active = 0;
    wb.recalc(); wb.format = 'xlsx';
    this.saveLocal(); this.docId = this.newId(); LS.set('cur', this.docId); this.vs = {}; this.attach(wb); this.saveLocal();
  }
}

G.addEventListener('DOMContentLoaded', () => { G.app = new Modern(); });
})(window);
