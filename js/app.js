// Excello ana uygulama
(function (G) {
'use strict';
const X = G.XF, M = G.XM, U = G.UI, IO = G.XIO, C = G.XC;
const { h, icon, $, $$ } = U;
const LS = { get(k, d) { try { const v = localStorage.getItem('excello.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } }, set(k, v) { try { localStorage.setItem('excello.' + k, JSON.stringify(v)); } catch (e) { } } };
const isMobile = () => matchMedia('(pointer: coarse)').matches;
const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const NF_MENU = [
  ['Genel', null, '1234,5'], ['Sayı', '#,##0.00', '1.234,50'], ['Tam sayı', '#,##0', '1.235'], ['Para (₺)', '#,##0.00 ₺', '1.234,50 ₺'], ['Dolar ($)', '$#,##0.00', '$1.234,50'], ['Euro (€)', '#,##0.00 €', '1.234,50 €'],
  ['Muhasebe (eksi kırmızı)', '#,##0.00 ₺;[Red]-#,##0.00 ₺', '-1.234,50 ₺'], ['Yüzde', '0%', '%12'], ['Yüzde (2 hane)', '0.00%', '12,34%'], ['Kısa tarih', 'dd.mm.yyyy', '30.10.2026'], ['Uzun tarih', 'd mmmm yyyy dddd', '30 Ekim 2026 Cuma'],
  ['Ay ve yıl', 'mmmm yyyy', 'Ekim 2026'], ['Saat', 'hh:mm', '14:30'], ['Tarih ve saat', 'dd.mm.yyyy hh:mm', '30.10.2026 14:30'], ['Bilimsel', '0.00E+00', '1,23E+03'], ['Metin', '@', 'abc']];

// ---------- IndexedDB ----------
const DB = {
  db: null,
  open() { if (this.db) return Promise.resolve(this.db); return new Promise((res, rej) => { let r; try { r = indexedDB.open('excello', 1); } catch (e) { return rej(e); } r.onupgradeneeded = () => r.result.createObjectStore('docs', { keyPath: 'id' }); r.onsuccess = () => res(this.db = r.result); r.onerror = () => rej(r.error); }); },
  async tx(mode, fn) { const work = (async () => { const db = await this.open(); return new Promise((res, rej) => { const t = db.transaction('docs', mode); const st = t.objectStore('docs'); const rq = fn(st); t.oncomplete = () => res(rq && rq.result); t.onerror = () => rej(t.error); }); })(); return Promise.race([work, new Promise((_, rej) => setTimeout(() => rej(new Error('zaman aşımı')), 3000))]); },
  put(d) { return this.tx('readwrite', s => s.put(d)).catch(e => { try { localStorage.setItem('excello.doc.' + d.id, JSON.stringify(d)); } catch (e2) { throw e; } }); },
  get(id) { return this.tx('readonly', s => s.get(id)).then(r => r || JSON.parse(localStorage.getItem('excello.doc.' + id) || 'null')).catch(() => JSON.parse(localStorage.getItem('excello.doc.' + id) || 'null')); }, all() { return this.tx('readonly', s => s.getAll()); }, del(id) { return this.tx('readwrite', s => s.delete(id)); }
};

class App {
  constructor() {
    this.zoom = LS.get('zoom', isMobile() && innerWidth < 700 ? 0.9 : 1);
    this.selectMode = false; this.geoVer = 0;
    this.wb = new M.Workbook();
    this.sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    this.fInput = $('#fInput'); this.nameBox = $('#nameBox');
    this.applyTheme(LS.get('theme', 'auto'));
    this.grid = new G.XG.Grid(this, $('#gridWrap'));
    this.wiz = new G.XW.Wizards(this);
    this.buildRibbon(); this.bindTop(); this.bindFbar(); this.bindKeys(); this.bindClipboardEvents();
    this.attach(this.wb, true);
    this.boot();
  }
  async boot() {
    const params = new URLSearchParams(location.search);
    const openUrl = params.get('open');
    try {
      const id = LS.get('cur', null);
      const doc = id && await DB.get(id);
      if (doc) { const wb = M.Workbook.fromJSON(doc.data); wb.format = doc.format || 'xlsx'; this.docId = id; this.attach(wb); }
      else this.docId = this.newId();
    } catch (e) { this.docId = this.newId(); }
    if (openUrl) this.openFromUrl(openUrl);
    else if (!LS.get('welcomed', false)) this.welcome();
    if ('serviceWorker' in navigator && /^https?:/.test(location.protocol) && location.hostname !== 'localhost') navigator.serviceWorker.register('sw.js').catch(() => { });
    addEventListener('beforeunload', () => this.saveLocal(true));
    document.addEventListener('visibilitychange', () => { if (document.hidden) this.saveLocal(true); });
  }
  newId() { return 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  attach(wb, first) {
    this.wb = wb; wb.listeners = [];
    wb.on(() => this.onChange());
    this.sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    this.clip = null; this.fileHandle = wb._handle || null;
    $('#docName').value = wb.name || 'Yeni Çalışma';
    this.refreshSheet();
    if (!first) this.updateUndo();
  }
  onChange() {
    this.grid.geoKey = ''; this.grid.render(); this.renderTabs(); this.onSelChange(); this.renderCharts(); this.scheduleSave(); this.updateUndo();
  }
  updateUndo() { $('#btnUndo').disabled = !this.wb.undoS.length; $('#btnRedo').disabled = !this.wb.redoS.length; }
  refreshSheet() {
    const sh = this.wb.sheet;
    this.grid.extraR = 0; this.grid.extraC = 0; this.grid.geoKey = '';
    this.sel = sh._sel || { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    this.grid.geo();
    const sc = this.grid.scroller; sc.scrollLeft = sh._sl || 0; sc.scrollTop = sh._st || 0;
    this.renderTabs(); this.renderCharts(); this.onSelChange(); this.grid.render();
  }
  switchSheet(i) {
    if (i === this.wb.active) return;
    if (this.grid.editing && !this.grid.canInsertRef()) this.grid.commitEdit();
    const sh = this.wb.sheet; sh._sel = this.sel; sh._sl = this.grid.scroller.scrollLeft; sh._st = this.grid.scroller.scrollTop;
    this.wb.active = i; this.refreshSheet(); this.grid.placeEditor();
  }

  // ---------- seçim ----------
  setSel(s, scroll) {
    const nR = this.grid.nR || 1000, nC = this.grid.nC || 50;
    s.r1 = Math.max(0, s.r1); s.c1 = Math.max(0, s.c1); s.r2 = Math.min(Math.max(s.r2, s.r1), X.MAXR - 1); s.c2 = Math.min(Math.max(s.c2, s.c1), X.MAXC - 1);
    this.sel = s; this.wb.sheet._sel = s;
    if (scroll) this.grid.scrollTo(s.r, s.c);
    this.onSelChange(); this.grid.render();
  }
  selBounds(limit) {
    // tam sütun/satır seçimlerini kullanılan alanla sınırla
    const s = this.sel, d = this.wb.dims(this.wb.active);
    const lr = Math.max(d.rows - 1, limit ? 0 : 0), lc = Math.max(d.cols - 1, 0);
    return { r1: s.r1, c1: s.c1, r2: s.r2 >= (this.grid.nR || 200) - 1 ? Math.max(s.r1, Math.min(s.r2, Math.max(lr, s.r1 + (limit || 0)))) : s.r2, c2: s.c2 >= (this.grid.nC || 30) - 1 ? Math.max(s.c1, Math.min(s.c2, Math.max(lc, s.c1))) : s.c2 };
  }
  isWholeCol() { return this.sel.r1 === 0 && this.sel.r2 >= (this.grid.nR || 200) - 1; }
  isWholeRow() { return this.sel.c1 === 0 && this.sel.c2 >= (this.grid.nC || 30) - 1; }
  refText(s, selSi, baseSi) {
    let t;
    const nR = this.grid.nR || 200, nC = this.grid.nC || 30;
    if (s.r1 === 0 && s.r2 >= nR - 1) t = X.colName(s.c1) + ':' + X.colName(s.c2);
    else if (s.c1 === 0 && s.c2 >= nC - 1) t = (s.r1 + 1) + ':' + (s.r2 + 1);
    else t = s.r1 === s.r2 && s.c1 === s.c2 ? X.addr(s.r1, s.c1) : X.addr(s.r1, s.c1) + ':' + X.addr(s.r2, s.c2);
    if (selSi != null && baseSi != null && selSi !== baseSi) t = X.quoteSheet(this.wb.sheets[selSi].name) + '!' + t;
    return t;
  }
  // Hata kodu olan hücre seçilince sade dille ne olduğunu söyle
  errHint(si, r, c) {
    const v = this.wb.value(si, r, c), key = si + ':' + r + ':' + c;
    if (!X.isErr(v)) { this._errKey = null; return; }
    if (this._errKey === key) return; this._errKey = key;
    const cell = this.wb.cell(si, r, c);
    if (v.code === '#TAŞMA!' && cell && cell._blk) {
      const b = cell._blk, wb = this.wb; const blockers = [];
      for (let i = b.r1; i <= b.r2; i++) for (let j = b.c1; j <= b.c2; j++) { if (i === r && j === c) continue; const o = wb.cell(si, i, j); if (o && ((o.v != null && o.v !== '') || o.f != null)) blockers.push(X.addr(i, j)); }
      const list = blockers.slice(0, 4).join(', ') + (blockers.length > 4 ? ' …' : '');
      U.confirm('Bu hücrenin sonucu bir liste ve aşağıya/yana ' + X.addr(b.r1, b.c1) + ':' + X.addr(b.r2, b.c2) + ' alanına yazılmak istiyor, ama ' + (blockers.length ? list + ' hücresi dolu' : 'başka bir liste o alanı kullanıyor') + '. ' + (blockers.length ? 'Engel olan hücreleri temizleyeyim mi?' : 'Formülü başka bir boş yere taşıyın.'), 'Sonuç için yer yok', blockers.length ? 'Temizle' : 'Tamam').then(ok => {
        if (!ok || !blockers.length) return;
        wb.begin(); for (let i = b.r1; i <= b.r2; i++) for (let j = b.c1; j <= b.c2; j++) { if (i === r && j === c) continue; const o = wb.cell(si, i, j); if (o && ((o.v != null && o.v !== '') || o.f != null)) wb.patchCell(si, i, j, { v: null, f: null }); }
        wb.end(); this.grid.render(); this.onSelChange(); U.toast('✓ Alan temizlendi, liste yerleşti');
      });
      return;
    }
    const msg = { '#DIV/0!': 'Sıfıra bölme var: bölen hücre boş ya da 0.', '#VALUE!': 'Yanlış türde değer: sayı beklenen yerde metin var.', '#N/A': 'Aranan değer bulunamadı.', '#NAME?': 'Tanınmayan işlev ya da ad.', '#REF!': 'Başvurulan hücre silinmiş.', '#NUM!': 'Hesaplanamayan sayı (çok büyük ya da geçersiz).', '#DÖNGÜ!': 'Formül kendi sonucunu kullanıyor (döngü).', '#HESAP!': 'Sonuç boş liste.', '#TAŞMA!': 'Sonuç listesi için yer yok: alttaki/yandaki hücreleri boşaltın.' }[v.code];
    if (msg) U.toast(this.wb.display(si, r, c) + ' → ' + msg);
  }
  onSelChange() {
    const s = this.sel, wb = this.wb, si = wb.active;
    this.nameBox.textContent = s.r1 === s.r2 && s.c1 === s.c2 ? X.addr(s.r, s.c) : this.refText(s);
    if (!this.grid.editing) this.fInput.value = wb.editText(si, s.r, s.c);
    const st = (wb.cell(si, s.r, s.c) || {}).s || {};
    const set = (a, on) => $$(`[data-a="${a}"]`).forEach(b => b.classList.toggle('on', !!on));
    set('b', st.b); set('i', st.i); set('u', st.u); set('st', st.st); set('wrap', st.wrap);
    set('alignL', st.ha === 'left'); set('alignC', st.ha === 'center'); set('alignR', st.ha === 'right');
    set('filter', wb.sheet.filter); set('gridlines', wb.sheet.grid !== false); set('selmode', this.selectMode);
    set('merge', wb.mergeAt(si, s.r, s.c));
    const fs = $('#fsVal'); if (fs) fs.textContent = st.fs || 11;
    this.errHint(si, s.r, s.c);
    this.updateStats();
    if (this.pick && this.pick.active) { const ref = this.refText(s, si, this.pick.baseSi); this.pick.onChange(ref); $('#pickRef').textContent = ref; }
    this.updateFxHint();
  }
  updateFxHint() {
    const c = this.wb.cell(this.wb.active, this.sel.r, this.sel.c); const b = $('#btnFx');
    b.classList.toggle('has', !!(c && c.f));
  }
  updateStats() {
    const el = $('#stats'); const s = this.sel;
    if (s.r1 === s.r2 && s.c1 === s.c2) { el.innerHTML = ''; return; }
    const g = this.selBounds(); const wb = this.wb, si = wb.active;
    let sum = 0, n = 0, cnt = 0, mn = Infinity, mx = -Infinity, k = 0;
    for (let r = g.r1; r <= g.r2 && k < 300000; r++) { if (wb.isRowHidden(si, r)) continue; for (let c = g.c1; c <= g.c2; c++, k++) { const v = wb.value(si, r, c); if (v == null || v === '') continue; cnt++; if (typeof v === 'number') { sum += v; n++; if (v < mn) mn = v; if (v > mx) mx = v; } } }
    this.stats = { sum, n, cnt, avg: n ? sum / n : 0, mn, mx };
    const f = v => X.formatValue(v, Math.abs(v) >= 1000 || !Number.isInteger(v) ? '#,##0.##' : '0');
    el.innerHTML = n ? `<span>Ortalama: <b>${f(sum / n)}</b></span><span>Say: <b>${cnt}</b></span><span>Toplam: <b>${f(sum)}</b></span>` : `<span>Say: <b>${cnt}</b></span>`;
  }
  moveSel(dir, extend, jump) {
    const s = this.sel, wb = this.wb, si = wb.active;
    let r = extend ? (s._er ?? s.r) : s.r, c = extend ? (s._ec ?? s.c) : s.c;
    const m = wb.mergeAt(si, r, c);
    const step = (dr, dc) => {
      if (jump) {
        const has = (rr, cc) => wb.value(si, rr, cc) != null;
        let nr = r + dr, nc = c + dc;
        if (!has(r, c) || !has(nr, nc)) { while (nr >= 0 && nc >= 0 && nr < this.grid.nR && nc < this.grid.nC && !has(nr, nc)) { nr += dr; nc += dc; } if (nr < 0 || nc < 0 || nr >= this.grid.nR || nc >= this.grid.nC) { nr -= dr; nc -= dc; } }
        else { while (has(nr + dr, nc + dc) && nr + dr >= 0 && nc + dc >= 0) { nr += dr; nc += dc; } }
        r = Math.max(0, nr); c = Math.max(0, nc); return;
      }
      if (m && !extend) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; if (dr < 0) r = m.r1; if (dc < 0) c = m.c1; }
      do { r += dr; c += dc; } while (r > 0 && c > 0 && (dr && wb.isRowHidden(si, r) || dc && wb.isColHidden(si, c)));
      r = Math.max(0, r); c = Math.max(0, c);
    };
    if (dir === 'down') step(1, 0); else if (dir === 'up') step(-1, 0); else if (dir === 'left') step(0, -1); else if (dir === 'right') step(0, 1);
    else if (dir === 'pgdn') r += Math.floor(this.grid.H / 26); else if (dir === 'pgup') r = Math.max(0, r - Math.floor(this.grid.H / 26));
    if (extend) { const n = this.grid.expandMerges({ r: s.r, c: s.c, r1: Math.min(s.r, r), c1: Math.min(s.c, c), r2: Math.max(s.r, r), c2: Math.max(s.c, c) }); n._er = r; n._ec = c; this.setSel(n, false); this.grid.scrollTo(r, c); }
    else { const mm = wb.mergeAt(si, r, c); this.setSel(mm ? { r: mm.r1, c: mm.c1, r1: mm.r1, c1: mm.c1, r2: mm.r2, c2: mm.c2 } : { r, c, r1: r, c1: c, r2: r, c2: c }, true); }
  }
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
  targetRange() { const s = this.sel; if (s.r1 !== s.r2 || s.c1 !== s.c2) return this.selBounds(); return this.currentRegion(s.r, s.c); }
  focusGrid() { const a = document.activeElement; if (a && (a === this.fInput || a.id === 'docName')) a.blur(); }
  closePopups() { U.closePop(); }

  // ---------- düzenleme ----------
  commitInput(si, r, c, text) {
    const wb = this.wb; const cell = wb.cell(si, r, c); const st = cell && cell.s;
    if (text.startsWith('=') && text.length > 1) {
      let f; try { f = X.toStore(text.slice(1)); X.parse(f); } catch (e) { U.toast('Formülde hata var: ' + e.message, 'warn'); return false; }
      wb.begin(); wb.patchCell(si, r, c, { f, v: null });
      const m = /^([A-Z][A-Z0-9.]*)\(/.exec(f); const fd = m && X.FN[m[1]];
      if (fd && fd.fmt && !(st && st.nf)) wb.setStyle(si, r, c, { nf: fd.fmt });
      wb.end(); return true;
    }
    const pi = IO.parseInput(text);
    const bad = wb.dvCheck(si, r, c, pi.v);
    if (bad) { U.alert(bad, 'Kural'); return false; }
    wb.begin(); wb.patchCell(si, r, c, { v: pi.v, f: null });
    if (pi.nf && !(st && st.nf && (X.isDateFormat(st.nf) === X.isDateFormat(pi.nf)))) wb.setStyle(si, r, c, { nf: pi.nf });
    wb.end(); return true;
  }
  evalFormula(f, t) {
    try { const ast = X.parse(f); let v = X.evaluate(ast, { model: this.wb, si: t.si, r: t.r, c: t.c }); if (X.isRng(v)) v = v.h && v.w ? v.get(0, 0) : X.E('#VALUE!'); return v == null ? 0 : v; }
    catch (e) { return X.isErr(e) ? e : X.E('#NAME?'); }
  }
  applyFormula(si, r, c, f, n, nf, hdr) {
    const wb = this.wb; wb.begin();
    if (hdr) { wb.patchCell(si, hdr.r, hdr.c, { v: hdr.text, f: null }); if (hdr.s) wb.setStyle(si, hdr.r, hdr.c, Object.assign({}, hdr.s, { nf: null })); }
    for (let i = 0; i < (n || 1); i++) {
      wb.patchCell(si, r + i, c, { f: i ? X.shiftFormula(f, i, 0) : f, v: null });
      const st = (wb.cell(si, r + i, c) || {}).s;
      if (nf && !(st && st.nf)) wb.setStyle(si, r + i, c, { nf });
      if (hdr && hdr.band) { const b = (wb.cell(si, r + i, c - 1) || {}).s; if (b && b.bg) wb.setStyle(si, r + i, c, { bg: b.bg }); }
    }
    wb.end();
    if (hdr) { const sh = wb.sheet; if (sh.filter && sh.filter.r1 === hdr.r && sh.filter.c2 === hdr.c - 1) sh.filter.c2 = hdr.c; this.autoFit([c]); this.moveChartsRightOf(si, c); }
    if (wb.active !== si) { wb.active = si; this.refreshSheet(); }
    this.setSel({ r, c, r1: r, c1: c, r2: r + (n || 1) - 1, c2: c }, true);
  }
  // yeni yazılan sütun bir grafiğin altında kalıyorsa grafiği sağa kaydır
  moveChartsRightOf(si, c) {
    const sh = this.wb.sheets[si]; if (!sh.charts.length) return;
    let x0 = 0; for (let k = 0; k < c; k++) x0 += sh.hiddenC.has(k) ? 0 : (sh.colW[k] || G.XG.DEF_W);
    const x1 = x0 + (sh.colW[c] || G.XG.DEF_W);
    const hit = sh.charts.filter(ch => ch.x < x1 && ch.x + ch.w > x0);
    if (!hit.length) return;
    this.wb.begin(); this.wb.snapSheet(si); hit.forEach(ch => { const s2 = sh.charts.find(x => x.id === ch.id); s2.x = x1 + 16; }); this.wb.end();
  }
  forSel(fn, limitRows) {
    const g = this.selBounds(limitRows || 0); const si = this.wb.active;
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(si, r, c, g);
  }
  styleSel(patch) {
    const wb = this.wb; wb.begin();
    const done = new Set();
    this.forSel((si, r, c) => { const m = wb.mergeAt(si, r, c); const rr = m ? m.r1 : r, cc = m ? m.c1 : c; const k = rr * 20000 + cc; if (done.has(k)) return; done.add(k); wb.setStyle(si, rr, cc, typeof patch === 'function' ? patch(wb.cell(si, rr, cc), rr, cc) : patch); }, 100);
    wb.end();
  }
  toggle(prop) { const st = (this.wb.cell(this.wb.active, this.sel.r, this.sel.c) || {}).s || {}; this.styleSel({ [prop]: st[prop] ? null : 1 }); }
  clearSel(mode) {
    const wb = this.wb; wb.begin();
    this.forSel((si, r, c) => { const cell = wb.cell(si, r, c); if (!cell) return; if (mode === 'all') wb.setCell(si, r, c, null); else if (mode === 'formats') wb.patchCell(si, r, c, { s: null }); else wb.patchCell(si, r, c, { v: null, f: null }); });
    if (mode === 'all' || mode === 'formats') { const g = this.selBounds(); const sh = wb.sheet; const inter = x => !(x.range.r2 < g.r1 || x.range.r1 > g.r2 || x.range.c2 < g.c1 || x.range.c1 > g.c2); if (sh.cf.some(inter) || (mode === 'all' && sh.dv.some(inter))) { wb.snapSheet(wb.active); sh.cf = sh.cf.filter(x => !inter(x)); if (mode === 'all') sh.dv = sh.dv.filter(x => !inter(x)); } }
    wb.end();
  }
  setNF(code) { this.styleSel({ nf: code }); }
  changeDecimals(delta) {
    const wb = this.wb, si = wb.active; const v = wb.value(si, this.sel.r, this.sel.c);
    this.styleSel(cell => {
      let nf = cell && cell.s && cell.s.nf;
      if (!nf || /^General$/i.test(nf)) { const val = cell ? wb.value(si, 0, 0) : null; const g = typeof v === 'number' ? X.fmtGeneral(v) : '0'; const d = (g.split(',')[1] || '').length; const n = Math.max(0, d + delta); return { nf: n ? '0.' + '0'.repeat(n) : '0' }; }
      if (X.isDateFormat(nf)) return {};
      const out = nf.split(';').map(sec => sec.replace(/([0#])(\.([0#]*))?(?![0#.,])/, (m, d0, dot, decs) => { const n = Math.max(0, (decs || '').length + delta); return d0 + (n ? '.' + '0'.repeat(n) : ''); })).join(';');
      return { nf: out };
    });
  }
  fontSize(d) { this.styleSel(cell => { const fs = ((cell && cell.s && cell.s.fs) || 11) + d; return { fs: Math.max(6, Math.min(72, fs)) === 11 ? null : Math.max(6, Math.min(72, fs)) }; }); }
  borders(kind, color) {
    const g = this.selBounds(100); const col = color || this.penColor || '#000000';
    const S = kind === 'thick' ? 'medium' : 'thin';
    this.styleSel((cell, r, c) => {
      const bd = Object.assign({}, cell && cell.s && cell.s.bd); const e = { s: S, c: col };
      if (kind === 'none') return { bd: null };
      if (kind === 'all') { bd.t = bd.b = bd.l = bd.r = e; }
      if (kind === 'outer' || kind === 'thick') { if (r === g.r1) bd.t = e; if (r === g.r2) bd.b = e; if (c === g.c1) bd.l = e; if (c === g.c2) bd.r = e; }
      if (kind === 'bottom' && r === g.r2) bd.b = e; if (kind === 'top' && r === g.r1) bd.t = e;
      if (kind === 'left' && c === g.c1) bd.l = e; if (kind === 'right' && c === g.c2) bd.r = e;
      if (kind === 'double' && r === g.r2) bd.b = { s: 'thick', c: col };
      return { bd: Object.keys(bd).length ? bd : null };
    });
  }
  async merge() {
    const wb = this.wb, si = wb.active, sh = wb.sheet; const g = this.selBounds();
    const inter = sh.merges.filter(m => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2));
    if (inter.length) { wb.begin(); wb.snapSheet(si); sh.merges = sh.merges.filter(m => !inter.includes(m)); wb.end(); return; }
    if (g.r1 === g.r2 && g.c1 === g.c2) { U.toast('Birleştirmek için birden fazla hücre seçin'); return; }
    let filled = 0; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (wb.value(si, r, c) != null) filled++;
    if (filled > 1 && !(await U.confirm('Birleştirince yalnızca sol üstteki değer kalır. Devam edilsin mi?', 'Hücreleri birleştir', 'Birleştir'))) return;
    wb.begin(); wb.snapSheet(si);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== g.r1 || c !== g.c1) { const cell = wb.cell(si, r, c); if (cell) cell.v = cell.f = null; }
    sh.merges.push({ ...g });
    const top = wb.cell(si, g.r1, g.c1); wb.setStyle(si, g.r1, g.c1, { ha: 'center', va: (top && top.s && top.s.va) || null });
    wb.end(); this.setSel(Object.assign({ r: g.r1, c: g.c1 }, g));
  }

  // ---------- satır/sütun ----------
  insertRows(after) { const s = this.selBounds(); const n = s.r2 - s.r1 + 1; this.wb.shift(this.wb.active, 'r', after ? s.r2 + 1 : s.r1, n); }
  insertCols(after) { const s = this.selBounds(); const n = s.c2 - s.c1 + 1; this.wb.shift(this.wb.active, 'c', after ? s.c2 + 1 : s.c1, n); }
  deleteRows() { const s = this.selBounds(); this.wb.shift(this.wb.active, 'r', s.r1, -(s.r2 - s.r1 + 1)); this.setSel({ r: s.r1, c: this.sel.c, r1: s.r1, c1: this.sel.c1, r2: s.r1, c2: this.sel.c2 }); }
  deleteCols() { const s = this.selBounds(); this.wb.shift(this.wb.active, 'c', s.c1, -(s.c2 - s.c1 + 1)); this.setSel({ r: this.sel.r, c: s.c1, r1: this.sel.r1, c1: s.c1, r2: this.sel.r2, c2: s.c1 }); }
  hide(axis, show) {
    const wb = this.wb, sh = wb.sheet, s = this.selBounds(); wb.begin(); wb.snapSheet(wb.active);
    if (axis === 'r') { if (show) { if (s.r1 === s.r2) sh.hiddenR.clear(); else for (let r = s.r1; r <= s.r2; r++) sh.hiddenR.delete(r); } else for (let r = s.r1; r <= s.r2; r++) sh.hiddenR.add(r); }
    else { if (show) { if (s.c1 === s.c2) sh.hiddenC.clear(); else for (let c = s.c1; c <= s.c2; c++) sh.hiddenC.delete(c); } else for (let c = s.c1; c <= s.c2; c++) sh.hiddenC.add(c); }
    wb.end();
  }
  autoFit(cols) {
    const wb = this.wb, si = wb.active, sh = wb.sheet; const ctx = this.grid.ctx; const d = wb.dims(si);
    if (!cols) { const g = this.selBounds(); cols = []; const c2 = this.isWholeRow() || (g.c1 === g.c2 && g.r1 === g.r2) ? Math.max(0, d.cols - 1) : g.c2; for (let c = (g.c1 === g.c2 && g.r1 === g.r2) ? 0 : g.c1; c <= c2; c++) cols.push(c); }
    wb.begin(); wb.snapSheet(si);
    for (const c of cols) {
      let w = 0;
      for (let r = 0; r < Math.min(d.rows, 3000); r++) { const cell = sh.cells.get(M.key(r, c)); if (!cell) continue; if (wb.mergeAt(si, r, c)) continue; const t = wb.display(si, r, c); if (!t) continue; ctx.font = this.grid.font(Object.assign({}, cell.s, { fs: (cell.s && cell.s.fs) || 11 })).replace(/[\d.]+px/, m => (parseFloat(m) / this.zoom) + 'px'); w = Math.max(w, ctx.measureText(t).width); }
      if (sh.filter && c >= sh.filter.c1 && c <= sh.filter.c2) w += 20;
      sh.colW[c] = Math.round(Math.max(40, Math.min(480, w + 16)));
    }
    wb.end();
  }

  // ---------- pano ----------
  copy(cut) {
    const g = this.selBounds(); const wb = this.wb, si = wb.active;
    const cells = []; const lines = [];
    for (let r = g.r1; r <= g.r2; r++) { const row = [], tl = []; for (let c = g.c1; c <= g.c2; c++) { const cell = wb.cell(si, r, c); row.push(cell ? { v: cell.v, f: cell.f, s: cell.s, cv: wb.value(si, r, c) } : wb.spillAt(si, r, c) ? { v: wb.value(si, r, c), f: null, s: null, cv: wb.value(si, r, c) } : null); let t = wb.display(si, r, c); if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"'; tl.push(t); } cells.push(row); lines.push(tl.join('\t')); }
    const merges = wb.sheet.merges.filter(m => m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2).map(m => ({ r1: m.r1 - g.r1, c1: m.c1 - g.c1, r2: m.r2 - g.r1, c2: m.c2 - g.c1 }));
    this.clipData = { si, g, cells, cut, merges }; this.clip = Object.assign({ si }, g);
    this.clipTSV = lines.join('\n');
    try { navigator.clipboard && navigator.clipboard.writeText(this.clipTSV).catch(() => { }); } catch (e) { }
    this.grid.render(); U.toast(cut ? 'Kesildi — yapıştırılacak yeri seçin' : 'Kopyalandı');
    return this.clipTSV;
  }
  async paste(mode) {
    let txt = null;
    if (!mode || mode === 'normal') { try { if (navigator.clipboard && navigator.clipboard.readText) txt = await navigator.clipboard.readText(); } catch (e) { } }
    this.pasteFrom(txt, mode);
  }
  pasteFrom(txt, mode) {
    const norm = s => (s || '').replace(/\r/g, '').replace(/\n$/, '');
    if (txt != null && txt !== '' && (!this.clipData || norm(txt) !== norm(this.clipTSV))) return this.pasteText(txt);
    if (!this.clipData) { U.toast('Panoda yapıştırılacak bir şey yok'); return; }
    const cd = this.clipData, wb = this.wb, si = wb.active; const H = cd.cells.length, W = cd.cells[0].length;
    const s = this.sel; let rows = H, cols = W;
    const tr = mode === 'transpose';
    if (tr) { rows = W; cols = H; }
    const sb = this.selBounds();
    if (H === 1 && W === 1 && (sb.r2 > sb.r1 || sb.c2 > sb.c1)) { rows = sb.r2 - sb.r1 + 1; cols = sb.c2 - sb.c1 + 1; }
    wb.begin();
    if (cd.cut) { if (cd.si !== si || true) wb.snapSheet(cd.si); for (let r = cd.g.r1; r <= cd.g.r2; r++) for (let c = cd.g.c1; c <= cd.g.c2; c++) wb.setCell(cd.si, r, c, null); }
    if (mode !== 'values' && mode !== 'formats') wb.snapSheet(si);
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      const sr = tr ? j % H : i % H, sc = tr ? i % W : j % W; const src = cd.cells[sr][sc];
      const r = sb.r1 + i, c = sb.c1 + j; const dst = wb.cell(si, r, c);
      if (mode === 'formats') { wb.patchCell(si, r, c, { s: src && src.s }); continue; }
      if (mode === 'values') { wb.patchCell(si, r, c, { v: src ? (X.isErr(src.cv) ? src.cv.code : src.cv) : null, f: null }); continue; }
      if (!src) { wb.setCell(si, r, c, null); continue; }
      const f = src.f != null ? (cd.cut ? src.f : X.shiftFormula(src.f, r - (cd.g.r1 + sr), c - (cd.g.c1 + sc))) : null;
      wb.setCell(si, r, c, { v: src.v, f, s: src.s });
    }
    if (mode !== 'values' && mode !== 'formats' && cd.merges.length && !tr) { const sh = wb.sheet; for (const m of cd.merges) sh.merges.push({ r1: sb.r1 + m.r1, c1: sb.c1 + m.c1, r2: sb.r1 + m.r2, c2: sb.c1 + m.c2 }); }
    wb.end();
    if (cd.cut) { this.clipData = null; this.clip = null; }
    this.setSel({ r: sb.r1, c: sb.c1, r1: sb.r1, c1: sb.c1, r2: sb.r1 + rows - 1, c2: sb.c1 + cols - 1 });
  }
  pasteText(txt) {
    const rows = parseTSV(txt); if (!rows.length) return;
    const wb = this.wb, si = wb.active, s = this.sel; wb.begin();
    rows.forEach((row, i) => row.forEach((t, j) => { const r = s.r1 + i, c = s.c1 + j; if (t.startsWith('=')) { try { const f = X.toStore(t.slice(1)); X.parse(f); wb.patchCell(si, r, c, { f, v: null }); return; } catch (e) { } } const pi = IO.parseInput(t); wb.patchCell(si, r, c, { v: pi.v, f: null }); if (pi.nf) wb.setStyle(si, r, c, { nf: pi.nf }); }));
    wb.end();
    this.setSel({ r: s.r1, c: s.c1, r1: s.r1, c1: s.c1, r2: s.r1 + rows.length - 1, c2: s.c1 + Math.max(...rows.map(r => r.length)) - 1 });
  }
  bindClipboardEvents() {
    const inField = () => { const a = document.activeElement; return a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.isContentEditable) || U.topDialog(); };
    document.addEventListener('copy', e => { if (inField()) return; e.preventDefault(); const t = this.copy(false); e.clipboardData.setData('text/plain', t); });
    document.addEventListener('cut', e => { if (inField()) return; e.preventDefault(); const t = this.copy(true); e.clipboardData.setData('text/plain', t); });
    document.addEventListener('paste', e => { if (inField()) return; e.preventDefault(); this.pasteFrom(e.clipboardData.getData('text/plain')); });
  }

  // ---------- doldurma ----------
  autofill(src, dst) {
    if (dst.r1 === src.r1 && dst.r2 === src.r2 && dst.c1 === src.c1 && dst.c2 === src.c2) return;
    const wb = this.wb, si = wb.active; wb.begin();
    if (dst.r2 < src.r2 && dst.r1 === src.r1 && dst.c1 === src.c1 && dst.c2 === src.c2) { for (let r = dst.r2 + 1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) wb.patchCell(si, r, c, { v: null, f: null }); wb.end(); return; }
    const vertical = dst.r1 !== src.r1 || dst.r2 !== src.r2;
    const lines = vertical ? range(src.c1, src.c2) : range(src.r1, src.r2);
    for (const ln of lines) {
      const seq = (vertical ? range(src.r1, src.r2) : range(src.c1, src.c2)).map(p => { const r = vertical ? p : ln, c = vertical ? ln : p; return { pos: p, cell: wb.cell(si, r, c) }; });
      const L = seq.length; const start = vertical ? src.r1 : src.c1, end = vertical ? src.r2 : src.c2;
      const tgt = vertical ? range(dst.r1, dst.r2) : range(dst.c1, dst.c2);
      const gen = seriesGen(seq.map(x => x.cell), (vertical ? wb.numFmt(si, src.r1, ln) : wb.numFmt(si, ln, src.c1)));
      for (const p of tgt) {
        if (p >= start && p <= end) continue;
        const i = p - start; const k = ((i % L) + L) % L; const sc = seq[k].cell;
        const r = vertical ? p : ln, c = vertical ? ln : p;
        if (!sc) { wb.setCell(si, r, c, null); continue; }
        if (sc.f != null) { wb.setCell(si, r, c, { f: X.shiftFormula(sc.f, vertical ? p - seq[k].pos : 0, vertical ? 0 : p - seq[k].pos), v: null, s: sc.s }); continue; }
        wb.setCell(si, r, c, { v: gen(i), f: null, s: sc.s });
      }
    }
    wb.end();
  }
  fill(dir) {
    const g = this.selBounds(); if (dir === 'down' && g.r1 === g.r2 || dir === 'right' && g.c1 === g.c2) { U.toast('Doldurmak için birden çok hücre seçin'); return; }
    if (dir === 'series') { const src = this.sel.r2 > this.sel.r1 ? { r1: g.r1, c1: g.c1, r2: Math.min(g.r1 + 1, g.r2), c2: g.c2 } : { r1: g.r1, c1: g.c1, r2: g.r2, c2: Math.min(g.c1 + 1, g.c2) }; const wb = this.wb; const two = this.sel.r2 > this.sel.r1 ? typeof wb.value(wb.active, g.r1 + 1, g.c1) === 'number' : typeof wb.value(wb.active, g.r1, g.c1 + 1) === 'number'; if (!two) { if (this.sel.r2 > this.sel.r1) src.r2 = g.r1; else src.c2 = g.c1; } this.autofill(src, g, true); return; }
    const src = dir === 'down' ? { r1: g.r1, c1: g.c1, r2: g.r1, c2: g.c2 } : { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c1 };
    const wb = this.wb, si = wb.active; wb.begin();
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (r <= src.r2 && c <= src.c2) continue; const sr = dir === 'down' ? src.r1 : r, sc = dir === 'down' ? c : src.c1; const s = wb.cell(si, sr, sc);
      wb.setCell(si, r, c, s ? { v: s.v, f: s.f != null ? X.shiftFormula(s.f, r - sr, c - sc) : null, s: s.s } : null);
    }
    wb.end();
  }

  // ---------- sıralama / filtre ----------
  detectHeader(g) {
    const wb = this.wb, si = wb.active; if (g.r2 <= g.r1) return false;
    const f = wb.sheet.filter; if (f && f.r1 === g.r1) return true;
    let txt = 0, nonTxt2 = 0;
    for (let c = g.c1; c <= g.c2; c++) { const a = wb.value(si, g.r1, c), b = wb.value(si, g.r1 + 1, c); if (typeof a === 'string') txt++; if (b != null && typeof b !== 'string') nonTxt2++; }
    const st = (wb.cell(si, g.r1, g.c1) || {}).s;
    return txt === g.c2 - g.c1 + 1 && (nonTxt2 > 0 || (st && st.b));
  }
  sort(keys, g, header) {
    const wb = this.wb, si = wb.active;
    g = g || this.targetRange(); if (header == null) header = this.detectHeader(g);
    const r0 = header ? g.r1 + 1 : g.r1; if (r0 >= g.r2 + (header ? 1 : 0)) return;
    const rows = []; for (let r = r0; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) { const x = wb.cell(si, r, c); cells.push(x ? { v: x.v, f: x.f, s: x.s } : null); } rows.push({ r, cells, keys: keys.map(k => wb.value(si, r, k.c)) }); }
    rows.sort((a, b) => { for (let i = 0; i < keys.length; i++) { const x = a.keys[i], y = b.keys[i]; const ex = x == null || x === '', ey = y == null || y === ''; if (ex || ey) { if (ex && ey) continue; return ex ? 1 : -1; } const c = X.cmp(x, y); if (c) return keys[i].asc ? c : -c; } return a.r - b.r; });
    wb.begin();
    rows.forEach((row, i) => { const r = r0 + i; row.cells.forEach((x, j) => { const c = g.c1 + j; wb.setCell(si, r, c, x ? { v: x.v, f: x.f != null ? X.shiftFormula(x.f, r - row.r, 0) : null, s: x.s } : null); }); });
    wb.end(); U.toast('Sıralandı');
  }
  quickSort(asc) { const g = this.targetRange(); this.sort([{ c: this.sel.c, asc }], g); }
  sortDialog() {
    const g = this.targetRange(); const wb = this.wb, si = wb.active; let header = this.detectHeader(g);
    const colName = c => header ? (wb.display(si, g.r1, c) || X.colName(c)) + ' (' + X.colName(c) + ')' : 'Sütun ' + X.colName(c);
    const body = h('div', { class: 'form' }); const keysBox = h('div');
    const hdr = h('input', { type: 'checkbox', checked: header }); const mkSel = () => { const s = h('select', { class: 'inp' }); for (let c = g.c1; c <= g.c2; c++) s.append(h('option', { value: c, text: colName(c) })); return s; };
    const rows = [];
    const addKey = c => { const s = mkSel(); s.value = c; const o = h('select', { class: 'inp' }, h('option', { value: '1', text: 'A → Z / Küçükten büyüğe' }), h('option', { value: '0', text: 'Z → A / Büyükten küçüğe' })); const row = h('div', { class: 'frow two' }, h('label', { class: 'lbl', text: rows.length ? 'Sonra şuna göre' : 'Şuna göre sırala' }), h('div', { class: 'pair' }, s, o)); rows.push({ s, o }); keysBox.append(row); };
    addKey(this.sel.c);
    body.append(h('p', { class: 'hint', text: `Alan: ${this.refText(g)}` }), h('label', { class: 'chk' }, hdr, ' İlk satır başlık'), keysBox, h('button', { class: 'btn ghost add', html: icon('plus') + ' Düzey ekle', onclick: () => addKey(g.c1) }));
    U.dialog({ title: icon('sortA') + ' Sırala', body, buttons: [{ text: 'Vazgeç', value: 0 }, { text: 'Sırala', kind: 'primary', onClick: () => { this.sort(rows.map(x => ({ c: +x.s.value, asc: x.o.value === '1' })), g, hdr.checked); }, value: 1 }] });
  }
  toggleFilter() {
    const wb = this.wb, sh = wb.sheet; wb.begin(); wb.snapSheet(wb.active);
    if (sh.filter) { sh.filter = null; sh.filterR.clear(); }
    else { const g = this.targetRange(); sh.filter = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, crit: {} }; U.toast('Başlıklardaki ▾ düğmelerine dokunarak filtreleyin'); }
    wb.end();
  }
  applyFilter() {
    const wb = this.wb, si = wb.active, sh = wb.sheet, f = sh.filter; sh.filterR = new Set(); if (!f) return;
    const d = wb.dims(si); f.r2 = Math.max(f.r2, d.rows - 1);
    for (let r = f.r1 + 1; r <= f.r2; r++) for (const c in f.crit) { const set = f.crit[c]; if (!set) continue; if (!set.includes(wb.display(si, r, +c))) { sh.filterR.add(r); break; } }
  }
  openFilter(c) {
    const wb = this.wb, si = wb.active, sh = wb.sheet, f = sh.filter; if (!f) return;
    const d = wb.dims(si); const r2 = Math.max(f.r2, d.rows - 1);
    const counts = new Map(); for (let r = f.r1 + 1; r <= r2; r++) { const t = wb.display(si, r, c); counts.set(t, (counts.get(t) || 0) + 1); }
    let vals = [...counts.keys()].sort((a, b) => { if (a === '') return 1; if (b === '') return -1; const na = X.parseNumLoose(a), nb = X.parseNumLoose(b); if (na != null && nb != null) return na - nb; return a.localeCompare(b, 'tr'); });
    const cur = f.crit[c] ? new Set(f.crit[c]) : null;
    const search = h('input', { class: 'inp search', placeholder: 'Ara…' }); const list = h('div', { class: 'checklist' }); const boxes = [];
    const all = h('input', { type: 'checkbox', checked: !cur });
    vals.slice(0, 2000).forEach(v => { const cb = h('input', { type: 'checkbox', checked: !cur || cur.has(v) }); boxes.push({ v, cb }); list.append(h('label', { class: 'chk' }, cb, h('span', { text: v === '' ? '(Boş)' : v }), h('small', { text: counts.get(v) }))); });
    all.addEventListener('change', () => boxes.forEach(b => { if (b.cb.closest('label').style.display !== 'none') b.cb.checked = all.checked; }));
    search.addEventListener('input', () => { const q = search.value.toLocaleLowerCase('tr-TR'); boxes.forEach(b => b.cb.closest('label').style.display = b.v.toLocaleLowerCase('tr-TR').includes(q) ? '' : 'none'); });
    const g = { r1: f.r1, c1: f.c1, r2: r2, c2: f.c2 };
    const body = h('div', null,
      h('div', { class: 'btnrow' }, h('button', { class: 'btn', html: icon('sortA') + ' A→Z', onclick: () => { dl.close(); this.sort([{ c, asc: true }], g, true); } }), h('button', { class: 'btn', html: icon('sortD') + ' Z→A', onclick: () => { dl.close(); this.sort([{ c, asc: false }], g, true); } })),
      search, h('label', { class: 'chk all' }, all, h('b', { text: 'Tümünü seç' })), list);
    const name = wb.display(si, f.r1, c) || X.colName(c);
    const dl = U.dialog({ title: icon('filter') + ' ' + U.esc(name), body, buttons: [{ text: 'Filtreyi kaldır', onClick: () => { this.setFilterCrit(c, null); } , value: 0 }, { text: 'Uygula', kind: 'primary', onClick: () => { const sel = boxes.filter(b => b.cb.checked).map(b => b.v); this.setFilterCrit(c, sel.length === boxes.length ? null : sel); }, value: 1 }] });
  }
  setFilterCrit(c, vals) { const wb = this.wb, sh = wb.sheet; wb.begin(); wb.snapSheet(wb.active); if (vals) sh.filter.crit[c] = vals; else delete sh.filter.crit[c]; this.applyFilter(); wb.end(); }
  clearFilters() { const wb = this.wb, sh = wb.sheet; if (!sh.filter) return; wb.begin(); wb.snapSheet(wb.active); sh.filter.crit = {}; sh.filterR.clear(); wb.end(); }
  removeDuplicates() {
    const g = this.targetRange(); const wb = this.wb, si = wb.active; let header = this.detectHeader(g);
    const body = h('div', { class: 'form' }); const hdr = h('input', { type: 'checkbox', checked: header }); const cols = [];
    const box = h('div', { class: 'checklist short' });
    for (let c = g.c1; c <= g.c2; c++) { const cb = h('input', { type: 'checkbox', checked: true }); cols.push({ c, cb }); box.append(h('label', { class: 'chk' }, cb, h('span', { text: (header ? wb.display(si, g.r1, c) + ' ' : '') + '(' + X.colName(c) + ')' }))); }
    body.append(h('p', { class: 'hint', text: 'Seçili sütunlarda aynı değerlere sahip tekrar eden satırlar silinir. Alan: ' + this.refText(g) }), h('label', { class: 'chk' }, hdr, ' İlk satır başlık'), box);
    U.dialog({ title: icon('layers') + ' Yinelenenleri kaldır', body, buttons: [{ text: 'Vazgeç', value: 0 }, { text: 'Kaldır', kind: 'primary', value: 1, onClick: () => {
      const keyCols = cols.filter(x => x.cb.checked).map(x => x.c); if (!keyCols.length) return false;
      const r0 = hdr.checked ? g.r1 + 1 : g.r1; const seen = new Set(); const keep = [];
      for (let r = r0; r <= g.r2; r++) { const k = keyCols.map(c => wb.display(si, r, c).toLocaleLowerCase('tr-TR')).join('\u0001'); if (seen.has(k)) continue; seen.add(k); keep.push(r); }
      const removed = g.r2 - r0 + 1 - keep.length;
      const data = keep.map(r => range(g.c1, g.c2).map(c => { const x = wb.cell(si, r, c); return x ? { v: x.v, f: x.f, s: x.s, r } : null; }));
      wb.begin();
      for (let i = 0; i <= g.r2 - r0; i++) for (let j = 0; j <= g.c2 - g.c1; j++) { const x = data[i] && data[i][j]; wb.setCell(si, r0 + i, g.c1 + j, x ? { v: x.v, f: x.f != null ? X.shiftFormula(x.f, r0 + i - x.r, 0) : null, s: x.s } : null); }
      wb.end(); U.toast(removed ? removed + ' yinelenen satır kaldırıldı' : 'Yinelenen satır bulunamadı');
    } }] });
  }
  findReplace(replaceMode) {
    const wb = this.wb; const body = h('div', { class: 'form' });
    const f = h('input', { class: 'inp', placeholder: 'Aranacak' }), r = h('input', { class: 'inp', placeholder: 'Yeni değer' });
    const mc = h('input', { type: 'checkbox' }), whole = h('input', { type: 'checkbox' }), allSheets = h('input', { type: 'checkbox' });
    body.append(h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Ara' }), f), h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Değiştir (isteğe bağlı)' }), r),
      h('div', { class: 'chkrow' }, h('label', { class: 'chk' }, mc, ' Büyük/küçük harf duyarlı'), h('label', { class: 'chk' }, whole, ' Hücrenin tamamı'), h('label', { class: 'chk' }, allSheets, ' Tüm sayfalar')));
    const res = h('p', { class: 'hint' }); body.append(res);
    const norm = s => mc.checked ? s : s.toLocaleLowerCase('tr-TR');
    const match = t => { const a = norm(t), b = norm(f.value); return whole.checked ? a === b : a.includes(b); };
    const hits = () => { const out = []; const sheets = allSheets.checked ? wb.sheets.map((_, i) => i) : [wb.active]; for (const si of sheets) { const ks = [...wb.sheets[si].cells.keys()].sort((a, b) => a - b); for (const k of ks) { const rr = M.kr(k), cc = M.kc(k); const t = wb.display(si, rr, cc); if (t && f.value && match(t)) out.push({ si, r: rr, c: cc }); } } return out; };
    let idx = -1;
    const next = () => { const hs = hits(); if (!hs.length) { res.textContent = 'Bulunamadı.'; return; } idx = (idx + 1) % hs.length; const x = hs[idx]; if (x.si !== wb.active) this.switchSheet(x.si); this.setSel({ r: x.r, c: x.c, r1: x.r, c1: x.c, r2: x.r, c2: x.c }, true); res.textContent = `${hs.length} sonuçtan ${idx + 1}.`; };
    const replaceAll = () => {
      const hs = hits(); let n = 0; wb.begin();
      const re = new RegExp(f.value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), mc.checked ? 'g' : 'gi');
      for (const x of hs) { const cell = wb.cell(x.si, x.r, x.c); if (!cell || cell.f != null || typeof cell.v !== 'string') continue; const nv = whole.checked ? r.value : cell.v.replace(re, r.value); if (nv !== cell.v) { const pi = IO.parseInput(nv); wb.patchCell(x.si, x.r, x.c, { v: pi.v }); n++; } }
      wb.end(); res.textContent = n + ' hücre değiştirildi.';
    };
    const d = U.dialog({ title: icon('search') + ' Bul ve değiştir', body, cls: 'side', buttons: [{ text: 'Tümünü değiştir', onClick: () => { replaceAll(); return false; } }, { text: 'Sonrakini bul', kind: 'primary', onClick: () => { next(); return false; } }] });
    f.addEventListener('keydown', e => { if (e.key === 'Enter') next(); });
    setTimeout(() => f.focus(), 60);
  }

  // ---------- dondurma / görünüm ----------
  freeze(kind) {
    const wb = this.wb, sh = wb.sheet; wb.begin(); wb.snapSheet(wb.active);
    if (kind === 'row') { sh.freezeR = 1; sh.freezeC = 0; } else if (kind === 'col') { sh.freezeR = 0; sh.freezeC = 1; } else if (kind === 'sel') { sh.freezeR = this.sel.r; sh.freezeC = this.sel.c; } else { sh.freezeR = 0; sh.freezeC = 0; }
    wb.end(); this.grid.scroller.scrollTop = 0; this.grid.scroller.scrollLeft = 0;
  }
  setZoom(z) {
    z = Math.max(0.5, Math.min(2.5, Math.round(z * 100) / 100)); const old = this.zoom; if (z === old) return;
    const sc = this.grid.scroller; const rx = sc.scrollLeft / old, ry = sc.scrollTop / old;
    this.zoom = z; this.grid.geoKey = ''; this.grid.geo(); sc.scrollLeft = rx * z; sc.scrollTop = ry * z;
    this.grid.render(); this.positionCharts(); this.grid.placeEditor(); $('#zoomVal').textContent = Math.round(z * 100) + '%'; LS.set('zoom', z);
  }
  markGeo(final) { this.geoVer++; this.grid.geoKey = ''; this.grid.render(); this.positionCharts(); this.wb.dirty = true; this.scheduleSave(); }
  applyTheme(t) {
    this.theme = t; LS.set('theme', t);
    const dark = t === 'dark' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
    document.documentElement.classList.toggle('dark', dark);
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    const mt = $('meta[name="theme-color"]'); if (mt) mt.content = dark ? '#18181b' : '#ffffff';
    if (this.grid) { this.grid.readTheme(); this.renderCharts(); }
    if (!this._mq) { this._mq = matchMedia('(prefers-color-scheme: dark)'); this._mq.addEventListener && this._mq.addEventListener('change', () => { if (this.theme === 'auto') this.applyTheme('auto'); }); }
  }

  // ---------- seçim (pick) modu ----------
  startPick(o) {
    const dlg = o.dlg && o.dlg(); const baseSi = this.wb.active;
    const prevSel = { ...this.sel };
    this.pick = { active: true, baseSi, onChange: o.onChange, prevSel };
    if (dlg) dlg.hide(true);
    const bar = $('#pickbar'); bar.classList.add('show'); $('#pickLabel').textContent = o.label || 'Hücreleri seçin'; $('#pickRef').textContent = o.initial || '—';
    // mevcut başvuruyu seç
    const t = (o.initial || '').trim(); if (t) { const i = t.lastIndexOf('!'); const sn = i >= 0 ? t.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'") : null; const body = t.slice(i + 1).replace(/\$/g, ''); const [a, b] = body.split(':'); const p = X.parseAddr(a), q = X.parseAddr(b || a); if (p && q) { if (sn) { const si = this.wb.sheetIndex(sn); if (si >= 0) this.switchSheet(si); } this.sel = { r: p.r, c: p.c, r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) }; this.grid.scrollTo(p.r, p.c); } }
    $('#gridWrap').classList.add('picking'); this.grid.render();
    U.toast(isMobile() ? 'Parmağınızı sürükleyerek hücreleri seçin' : 'Fareyle sürükleyerek hücreleri seçin');
    const finish = ok => { bar.classList.remove('show'); $('#gridWrap').classList.remove('picking'); this.pick = null; if (this.wb.active !== baseSi) this.switchSheet(baseSi); this.sel = prevSel; this.onSelChange(); this.grid.render(); if (dlg) dlg.hide(false); o.onDone && o.onDone(ok); };
    $('#pickOk').onclick = () => finish(true); $('#pickCancel').onclick = () => finish(false);
  }

  // ---------- grafikler ----------
  chartDialog(existing) {
    const wb = this.wb, si = wb.active; const g = existing ? existing.range : this.targetRange();
    const spec = existing ? { ...existing } : { id: 'c' + Date.now().toString(36), type: 'column', range: g, title: '', labels: true };
    const data = () => C.extract(wb, si, spec);
    const cv = h('canvas', { class: 'chart-prev' });
    const types = h('div', { class: 'ctypes' });
    const ICON = { column: 'chart', bar: 'chart', line: 'line', area: 'line', pie: 'pie', doughnut: 'pie', stacked: 'chart', scatter: 'dots' };
    C.TYPES.forEach(([k, t]) => types.append(h('button', { class: 'ctype' + (spec.type === k ? ' on' : ''), 'data-k': k, html: icon(ICON[k], k === 'bar' ? 'rot' : '') + `<span>${t}</span>`, onclick: () => { spec.type = k; U.$$('.ctype', types).forEach(b => b.classList.toggle('on', b.dataset.k === k)); draw(); } })));
    const title = h('input', { class: 'inp', placeholder: 'Grafik başlığı (isteğe bağlı)', value: spec.title || '', oninput: () => { spec.title = title.value; draw(); } });
    const lab = h('input', { type: 'checkbox', checked: spec.labels !== false, onchange: () => { spec.labels = lab.checked; draw(); } });
    const rng = h('input', { class: 'inp', value: IO.rangeRef(spec.range) });
    const pick = this.wiz.pickField({ label: 'Grafik verisi', type: 'range' }, rng, () => d, () => { const gg = IO.refToRange(rng.value.replace(/^.*!/, '')); if (gg) { spec.range = gg; draw(); } });
    const draw = () => requestAnimationFrame(() => C.draw(cv, spec, data(), this.chartTheme()));
    const body = h('div', { class: 'form' }, types, h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Veri alanı' }), pick), h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Başlık' }), title), h('label', { class: 'chk' }, lab, ' Değer etiketlerini göster'), cv);
    const d = U.dialog({ title: icon('chart') + (existing ? ' Grafiği düzenle' : ' Grafik ekle'), body, wide: true, buttons: [{ text: 'Vazgeç', value: 0 }, { text: existing ? 'Kaydet' : 'Ekle', kind: 'primary', value: 1, onClick: () => {
      const sh = wb.sheet; wb.begin(); wb.snapSheet(si);
      if (existing) { const i = sh.charts.findIndex(x => x.id === existing.id); if (i >= 0) sh.charts[i] = Object.assign({}, sh.charts[i], spec); }
      else { this.grid.geo(); spec.x = Math.round(this.grid.colX[Math.min(g.c2 + 1, this.grid.nC)] / this.zoom) + 16; spec.y = Math.round(this.grid.rowY[g.r1] / this.zoom); spec.w = Math.min(520, innerWidth - 80); spec.h = 320; sh.charts.push(spec); }
      wb.end();
      if (!existing) { const el = $(`.chart[data-id="${spec.id}"]`); el && el.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' }); }
    } }] });
    draw();
  }
  chartTheme() { const dark = document.documentElement.classList.contains('dark'); return dark ? { bg: '#1f1f23', fg: '#e4e4e7', muted: '#a1a1aa', grid: '#3f3f46' } : null; }
  renderCharts() {
    const layer = this.grid.spacer; const sh = this.wb.sheet; const si = this.wb.active;
    const ids = new Set(sh.charts.map(c => c.id));
    $$('.chart', layer).forEach(el => { if (!ids.has(el.dataset.id) || el.dataset.si !== String(si)) el.remove(); });
    for (const ch of sh.charts) {
      let el = $(`.chart[data-id="${ch.id}"]`, layer);
      if (!el) { el = this.makeChartEl(ch); layer.append(el); }
      el._spec = ch; el.dataset.si = si;
      const cv = $('canvas', el); C.draw(cv, ch, C.extract(this.wb, si, ch), this.chartTheme());
    }
    this.positionCharts();
  }
  positionCharts() {
    const z = this.zoom; const hw = this.grid.hw(), hh = this.grid.hh();
    $$('.chart', this.grid.spacer).forEach(el => { const ch = el._spec; if (!ch) return; el.style.left = (hw + ch.x * z) + 'px'; el.style.top = (hh + ch.y * z) + 'px'; el.style.width = (ch.w * z) + 'px'; el.style.height = (ch.h * z) + 'px'; });
  }
  makeChartEl(ch) {
    const el = h('div', { class: 'chart', 'data-id': ch.id });
    const bar = h('div', { class: 'chart-bar' }, h('span', { class: 'grip', html: icon('dots') }),
      h('button', { class: 'ic ghost', title: 'Düzenle', html: icon('rules'), onclick: e => { e.stopPropagation(); this.chartDialog(el._spec); } }),
      h('button', { class: 'ic ghost', title: 'Resim olarak kaydet', html: icon('share'), onclick: e => { e.stopPropagation(); const s = el._spec; const url = C.toPNG(s, C.extract(this.wb, this.wb.active, s), s.w * 2, s.h * 2); fetch(url).then(r => r.blob()).then(b => this.deliver(b, (s.title || 'grafik') + '.png', 'image/png')); } }),
      h('button', { class: 'ic ghost', title: 'Sil', html: icon('trash'), onclick: e => { e.stopPropagation(); const wb = this.wb; wb.begin(); wb.snapSheet(wb.active); wb.sheet.charts = wb.sheet.charts.filter(x => x.id !== el._spec.id); wb.end(); } }));
    const cv = h('canvas'); const rs = h('div', { class: 'chart-rs' });
    el.append(bar, cv, rs);
    const drag = (mode, e0) => {
      e0.preventDefault(); e0.stopPropagation();
      const p0 = { x: e0.clientX, y: e0.clientY }; const s = el._spec; const o = { x: s.x, y: s.y, w: s.w, h: s.h }; const wb = this.wb;
      wb.begin(); wb.snapSheet(wb.active);
      const sp = wb.sheet.charts.find(x => x.id === s.id) || s;
      const mv = e => { const dx = (e.clientX - p0.x) / this.zoom, dy = (e.clientY - p0.y) / this.zoom; if (mode === 'move') { sp.x = Math.max(0, Math.round(o.x + dx)); sp.y = Math.max(0, Math.round(o.y + dy)); } else { sp.w = Math.max(200, Math.round(o.w + dx)); sp.h = Math.max(150, Math.round(o.h + dy)); } el._spec = sp; this.positionCharts(); if (mode !== 'move') C.draw(cv, sp, C.extract(wb, wb.active, sp), this.chartTheme()); };
      const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); wb.end(); };
      addEventListener('pointermove', mv); addEventListener('pointerup', up);
    };
    bar.addEventListener('pointerdown', e => { if (e.target.closest('button')) return; drag('move', e); });
    rs.addEventListener('pointerdown', e => drag('size', e));
    el.addEventListener('touchstart', e => e.stopPropagation(), { passive: true });
    el.addEventListener('dblclick', e => { e.stopPropagation(); this.chartDialog(el._spec); });
    return el;
  }

  // ---------- tablo stili ----------
  tableStyleDialog() {
    const g = this.targetRange(); const wb = this.wb, si = wb.active;
    const STY = [['#2563eb', '#dbeafe', '#eff6ff'], ['#16a34a', '#dcfce7', '#f0fdf4'], ['#ea580c', '#ffedd5', '#fff7ed'], ['#7c3aed', '#ede9fe', '#f5f3ff'], ['#0d9488', '#ccfbf1', '#f0fdfa'], ['#db2777', '#fce7f3', '#fdf2f8'], ['#334155', '#e2e8f0', '#f8fafc'], ['#ca8a04', '#fef9c3', '#fefce8']];
    let pick = 0;
    const sw = h('div', { class: 'tstyles' });
    STY.forEach((s, i) => sw.append(h('button', { class: 'tsw' + (i === 0 ? ' on' : ''), onclick: e => { pick = i; U.$$('.tsw', sw).forEach((b, j) => b.classList.toggle('on', j === i)); }, html: `<i style="background:${s[0]}"></i><i style="background:${s[1]}"></i><i style="background:${s[2]}"></i><i style="background:${s[1]}"></i>` })));
    const hdr = h('input', { type: 'checkbox', checked: true }), band = h('input', { type: 'checkbox', checked: true }), bord = h('input', { type: 'checkbox', checked: true }), flt = h('input', { type: 'checkbox', checked: true }), tot = h('input', { type: 'checkbox' });
    const body = h('div', { class: 'form' }, h('p', { class: 'hint', text: 'Alan: ' + this.refText(g) }), sw,
      h('div', { class: 'chkrow' }, h('label', { class: 'chk' }, hdr, ' Başlık satırı'), h('label', { class: 'chk' }, band, ' Bantlı satırlar'), h('label', { class: 'chk' }, bord, ' Kenarlıklar'), h('label', { class: 'chk' }, flt, ' Filtre düğmeleri'), h('label', { class: 'chk' }, tot, ' Toplam satırı ekle')));
    U.dialog({ title: icon('table') + ' Tablo stili', body, buttons: [{ text: 'Vazgeç', value: 0 }, { text: 'Uygula', kind: 'primary', value: 1, onClick: () => {
      const [c0, c1, c2] = STY[pick]; const sh = wb.sheet; wb.begin(); wb.snapSheet(si);
      let last = g.r2;
      if (tot.checked) { last = g.r2 + 1; for (let c = g.c1; c <= g.c2; c++) { let num = false; for (let r = g.r1 + (hdr.checked ? 1 : 0); r <= g.r2; r++) if (typeof wb.value(si, r, c) === 'number') { num = true; break; } if (c === g.c1 && !num) wb.patchCell(si, last, c, { v: 'Toplam', f: null }); else if (num) wb.patchCell(si, last, c, { f: `SUM(${X.addr(g.r1 + (hdr.checked ? 1 : 0), c)}:${X.addr(g.r2, c)})`, v: null }); } }
      for (let r = g.r1; r <= last; r++) for (let c = g.c1; c <= g.c2; c++) {
        const isH = hdr.checked && r === g.r1, isT = tot.checked && r === last; const i = r - g.r1 - (hdr.checked ? 1 : 0);
        const bd = bord.checked ? { t: { s: 'thin', c: c1 }, b: { s: 'thin', c: c1 }, l: { s: 'thin', c: c1 }, r: { s: 'thin', c: c1 } } : null;
        if (isT && bd) bd.t = { s: 'medium', c: c0 };
        const patch = isH ? { bg: c0, fc: '#ffffff', b: 1, bd } : isT ? { b: 1, bg: c2, bd } : { bg: band.checked ? (i % 2 ? c2 : '#ffffff') : null, bd };
        if (isT) { const nf = (wb.cell(si, g.r2, c) || {}).s; if (nf && nf.nf) patch.nf = nf.nf; }
        wb.setStyle(si, r, c, patch);
      }
      if (flt.checked && hdr.checked) { sh.filter = { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2, crit: {} }; sh.filterR.clear(); }
      wb.end(); this.autoFit(range(g.c1, g.c2));
    } }] });
  }

  // ---------- koşullu biçim ----------
  rulesDialog() {
    const g = this.targetRange(); const wb = this.wb, si = wb.active;
    const TYPES = [
      ['gt', 'Büyükse', 1], ['lt', 'Küçükse', 1], ['between', 'Arasındaysa', 2], ['eq', 'Eşitse', 1], ['contains', 'Metni içeriyorsa', 1], ['begins', 'Şununla başlıyorsa', 1],
      ['dup', 'Yinelenen değerler', 0], ['unique', 'Benzersiz değerler', 0], ['top', 'En büyük N', 'n'], ['bottom', 'En küçük N', 'n'], ['above', 'Ortalamanın üstü', 0], ['below', 'Ortalamanın altı', 0],
      ['neg', 'Negatif sayılar', 0], ['blank', 'Boş hücreler', 0], ['notblank', 'Dolu hücreler', 0], ['scale', 'Renk skalası', 's'], ['bar', 'Veri çubukları', 'b'], ['icons', 'Ok simgeleri', 'i']];
    const PRE = [['Kırmızı', { bg: '#fee2e2', fc: '#b91c1c' }], ['Yeşil', { bg: '#dcfce7', fc: '#15803d' }], ['Sarı', { bg: '#fef9c3', fc: '#a16207' }], ['Mavi', { bg: '#dbeafe', fc: '#1d4ed8' }], ['Mor', { bg: '#ede9fe', fc: '#6d28d9' }], ['Kalın kırmızı yazı', { fc: '#dc2626', b: 1 }], ['Kalın yeşil yazı', { fc: '#16a34a', b: 1 }], ['Turuncu dolgu', { bg: '#fdba74' }]];
    let type = 'gt', style = PRE[0][1], scale = 0;
    const SCALES = [['#f8696b', '#ffeb84', '#63be7b'], ['#63be7b', '#ffeb84', '#f8696b'], ['#ffffff', null, '#2563eb'], ['#ffffff', null, '#16a34a']];
    const body = h('div', { class: 'form' });
    const grid = h('div', { class: 'cfgrid' });
    TYPES.forEach(([k, t]) => grid.append(h('button', { class: 'chip' + (k === type ? ' on' : ''), 'data-k': k, text: t, onclick: () => { type = k; U.$$('.chip', grid).forEach(b => b.classList.toggle('on', b.dataset.k === k)); upd(); } })));
    const a = h('input', { class: 'inp', placeholder: 'Değer' }), b = h('input', { class: 'inp', placeholder: 've' });
    const vals = h('div', { class: 'frow' }); const styles = h('div', { class: 'cfstyles' });
    PRE.forEach(([t, s], i) => styles.append(h('button', { class: 'cfs' + (i === 0 ? ' on' : ''), style: { background: s.bg || 'transparent', color: s.fc, fontWeight: s.b ? 700 : 400 }, text: t, onclick: e => { style = s; U.$$('.cfs', styles).forEach(x => x.classList.toggle('on', x === e.currentTarget)); } })));
    const scales = h('div', { class: 'cfstyles' }); SCALES.forEach((s, i) => scales.append(h('button', { class: 'cfs scale' + (i === 0 ? ' on' : ''), style: { background: `linear-gradient(90deg, ${s[0]}, ${s[1] || s[0]}, ${s[2]})` }, onclick: e => { scale = i; U.$$('.scale', scales).forEach(x => x.classList.toggle('on', x === e.currentTarget)); } })));
    const rng = h('input', { class: 'inp', value: IO.rangeRef(g) });
    let d;
    const upd = () => {
      const t = TYPES.find(x => x[0] === type); vals.innerHTML = '';
      if (t[2] === 1 || t[2] === 2 || t[2] === 'n') { vals.append(h('label', { class: 'lbl', text: t[2] === 'n' ? 'Kaç tane?' : 'Değer' })); const p = h('div', { class: 'pair' }, a); if (t[2] === 2) p.append(b); vals.append(p); if (t[2] === 'n' && !a.value) a.value = '10'; }
      styles.style.display = /[sbi]/.test(String(t[2])) ? 'none' : ''; scales.style.display = t[2] === 's' ? '' : 'none';
    };
    const list = h('div', { class: 'rulelist' });
    const renderList = () => {
      list.innerHTML = ''; const sh = wb.sheet; if (!sh.cf.length) return;
      list.append(h('div', { class: 'lbl', text: 'Bu sayfadaki kurallar' }));
      sh.cf.forEach((ru, i) => { const t = TYPES.find(x => x[0] === ru.type); list.append(h('div', { class: 'rule' }, h('i', { class: 'sw', style: { background: ru.style && ru.style.bg || (ru.type === 'scale' ? 'linear-gradient(90deg,#f8696b,#63be7b)' : ru.type === 'bar' ? '#4f8ef7' : '#eee') } }), h('span', { text: `${t ? t[1] : ru.type}${ru.a != null && ru.a !== '' ? ' ' + ru.a : ''}${ru.b != null && ru.b !== '' ? ' – ' + ru.b : ''}  ·  ${IO.rangeRef(ru.range)}` }), h('button', { class: 'ic ghost', html: icon('trash'), onclick: () => { wb.begin(); wb.snapSheet(si); wb.sheet.cf.splice(i, 1); wb.end(); renderList(); } }))); });
    };
    body.append(h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Uygulanacak alan' }), this.wiz.pickField({ label: 'Kural alanı', type: 'range' }, rng, () => d)), h('label', { class: 'lbl', text: 'Kural' }), grid, vals, h('label', { class: 'lbl', text: 'Görünüm' }), styles, scales, list);
    d = U.dialog({ title: icon('rules') + ' Renk kuralları', body, wide: true, buttons: [{ text: 'Kapat', value: 0 }, { text: 'Kural ekle', kind: 'primary', onClick: () => {
      const gg = IO.refToRange(rng.value.replace(/^.*!/, '')); if (!gg) { U.toast('Alanı kontrol edin', 'warn'); return false; }
      const num = v => { const n = X.parseNumLoose(v); return n != null ? n : v; };
      const ru = { type, range: gg, style: { ...style } };
      const t = TYPES.find(x => x[0] === type);
      if (t[2] === 1 || t[2] === 2 || t[2] === 'n') { if (a.value.trim() === '') { U.toast('Bir değer girin', 'warn'); return false; } ru.a = num(a.value.trim()); }
      if (t[2] === 2) ru.b = num(b.value.trim());
      if (type === 'scale') { const s = SCALES[scale]; ru.c1 = s[0]; ru.c2 = s[1]; ru.c3 = s[2]; ru.style = {}; }
      if (type === 'bar') { ru.color = '#4f8ef7'; ru.style = {}; } if (type === 'icons') ru.style = {};
      wb.begin(); wb.snapSheet(si); wb.sheet.cf.push(ru); wb.end(); renderList(); U.toast('✓ Kural eklendi'); return false;
    } }] });
    upd(); renderList();
  }

  // ---------- veri doğrulama ----------
  validationDialog() {
    const g = this.targetRange(); const wb = this.wb, si = wb.active; const ex = wb.dvAt(si, g.r1, g.c1);
    const type = h('select', { class: 'inp' }, h('option', { value: 'list', text: 'Açılır liste (seçenekler)' }), h('option', { value: 'whole', text: 'Tam sayı' }), h('option', { value: 'decimal', text: 'Ondalık sayı' }), h('option', { value: 'date', text: 'Tarih' }), h('option', { value: 'textlen', text: 'Metin uzunluğu' }));
    const items = h('textarea', { class: 'inp', rows: 4, placeholder: 'Her satıra bir seçenek yazın\nÖr:\nEvet\nHayır\nBeklemede' });
    const src = h('input', { class: 'inp', placeholder: 'veya tablodan bir liste seçin' });
    const mn = h('input', { class: 'inp', placeholder: 'En az' }), mx = h('input', { class: 'inp', placeholder: 'En çok' });
    const msg = h('input', { class: 'inp', placeholder: 'Hata mesajı (isteğe bağlı)' });
    const rng = h('input', { class: 'inp', value: IO.rangeRef(g) }); let d;
    if (ex) { type.value = ex.type; items.value = (ex.items || []).join('\n'); src.value = ex.src || ''; mn.value = ex.min ?? ''; mx.value = ex.max ?? ''; msg.value = ex.msg || ''; }
    const listBox = h('div', null, h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Seçenekler' }), items), h('div', { class: 'frow' }, this.wiz.pickField({ label: 'Liste kaynağı', type: 'range' }, src, () => d)));
    const numBox = h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Aralık' }), h('div', { class: 'pair' }, mn, mx));
    const upd = () => { listBox.style.display = type.value === 'list' ? '' : 'none'; numBox.style.display = type.value === 'list' ? 'none' : ''; };
    type.addEventListener('change', upd);
    const body = h('div', { class: 'form' }, h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Alan' }), this.wiz.pickField({ label: 'Kural alanı', type: 'range' }, rng, () => d)), h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'İzin verilen' }), type), listBox, numBox, h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Mesaj' }), msg));
    d = U.dialog({ title: icon('list') + ' Açılır liste / giriş kuralı', body, buttons: [{ text: 'Kuralı kaldır', onClick: () => { const gg = IO.refToRange(rng.value.replace(/^.*!/, '')) || g; wb.begin(); wb.snapSheet(si); wb.sheet.dv = wb.sheet.dv.filter(x => x.range.r2 < gg.r1 || x.range.r1 > gg.r2 || x.range.c2 < gg.c1 || x.range.c1 > gg.c2); wb.end(); }, value: 0 }, { text: 'Kaydet', kind: 'primary', onClick: () => {
      const gg = IO.refToRange(rng.value.replace(/^.*!/, '')); if (!gg) return false;
      const o = { range: gg, type: type.value, msg: msg.value.trim() };
      if (o.type === 'list') { o.items = items.value.split(/[\n;]/).map(s => s.trim()).filter(Boolean); if (!o.items.length) { delete o.items; o.src = src.value.trim(); if (!o.src) { U.toast('Seçenek girin veya liste seçin', 'warn'); return false; } } }
      else { const p = v => { const t = v.trim(); if (!t) return null; if (o.type === 'date') { const dd = X.parseDateLoose(t); return dd ? dd.v : null; } return X.parseNumLoose(t); }; o.min = p(mn.value); o.max = p(mx.value); }
      wb.begin(); wb.snapSheet(si); wb.sheet.dv = wb.sheet.dv.filter(x => x.range.r2 < gg.r1 || x.range.r1 > gg.r2 || x.range.c2 < gg.c1 || x.range.c1 > gg.c2); wb.sheet.dv.push(o); wb.end();
      U.toast(o.type === 'list' ? '✓ Açılır liste eklendi — hücreye dokunun' : '✓ Kural eklendi');
    }, value: 1 }] });
    upd();
  }
  openDvList(r, c) {
    const wb = this.wb, si = wb.active; const dv = wb.dvAt(si, r, c); if (!dv) return;
    const its = wb.dvItems(si, dv); const rect = this.grid.cellRect(r, c); const wr = $('#gridWrap').getBoundingClientRect();
    U.menu({ left: wr.left + rect.x, top: wr.top + rect.y, bottom: wr.top + rect.y + rect.h }, [...its.map(t => ({ text: t, onClick: () => { const pi = IO.parseInput(t); wb.begin(); wb.patchCell(si, r, c, { v: pi.v, f: null }); wb.end(); } })), { sep: 1 }, { text: 'Temizle', icon: 'eraser', onClick: () => { wb.begin(); wb.patchCell(si, r, c, { v: null, f: null }); wb.end(); } }]);
  }

  // ---------- sayfa sekmeleri ----------
  renderTabs() {
    const box = $('#sheetTabs'); box.innerHTML = '';
    this.wb.sheets.forEach((s, i) => {
      const b = h('button', { class: 'stab' + (i === this.wb.active ? ' on' : ''), text: s.name, onclick: () => this.switchSheet(i), ondblclick: () => this.sheetMenu(i, b), oncontextmenu: e => { e.preventDefault(); this.sheetMenu(i, b); } });
      let lt; b.addEventListener('touchstart', () => { lt = setTimeout(() => { lt = null; this.sheetMenu(i, b); }, 500); }, { passive: true }); b.addEventListener('touchend', () => clearTimeout(lt)); b.addEventListener('touchmove', () => clearTimeout(lt));
      box.append(b);
    });
  }
  sheetMenu(i, anchor) {
    const wb = this.wb;
    U.menu(anchor, [
      { text: 'Yeniden adlandır', icon: 'italic', onClick: async () => { const n = await U.prompt('Sayfa adı', wb.sheets[i].name); if (n != null && !wb.renameSheet(i, n)) U.toast('Bu ad kullanılamaz', 'warn'); } },
      { text: 'Çoğalt', icon: 'copy', onClick: () => { wb.duplicateSheet(i); this.refreshSheet(); } },
      { text: 'Sola taşı', icon: 'left', disabled: i === 0, onClick: () => wb.moveSheet(i, -1) },
      { text: 'Sağa taşı', icon: 'right', disabled: i === wb.sheets.length - 1, onClick: () => wb.moveSheet(i, 1) },
      { sep: 1 },
      { text: 'Sayfayı sil', icon: 'trash', danger: 1, disabled: wb.sheets.length < 2, onClick: async () => { if (await U.confirm(`"${wb.sheets[i].name}" sayfası silinsin mi?`, 'Sayfayı sil', 'Sil')) { wb.removeSheet(i); this.refreshSheet(); } } },
    ]);
  }

  // ---------- bağlam menüsü ----------
  contextMenu(x, y, hit) {
    const at = { left: x, top: y, bottom: y };
    if (hit && hit.t === 'col') return U.menu(at, [
      { text: 'Sola sütun ekle', icon: 'colAdd', onClick: () => this.insertCols() }, { text: 'Sağa sütun ekle', icon: 'colAdd', onClick: () => this.insertCols(true) }, { text: 'Sütunu sil', icon: 'colDel', danger: 1, onClick: () => this.deleteCols() },
      { sep: 1 }, { text: 'Genişliği sığdır', icon: 'fit', onClick: () => this.autoFit() }, { text: 'Sütunu gizle', icon: 'eyeOff', onClick: () => this.hide('c') }, { text: 'Gizlileri göster', icon: 'eye', onClick: () => this.hide('c', true) },
      { sep: 1 }, { text: 'A → Z sırala', icon: 'sortA', onClick: () => this.quickSort(true) }, { text: 'Z → A sırala', icon: 'sortD', onClick: () => this.quickSort(false) }]);
    if (hit && hit.t === 'row') return U.menu(at, [
      { text: 'Üste satır ekle', icon: 'rowAdd', onClick: () => this.insertRows() }, { text: 'Alta satır ekle', icon: 'rowAdd', onClick: () => this.insertRows(true) }, { text: 'Satırı sil', icon: 'rowDel', danger: 1, onClick: () => this.deleteRows() },
      { sep: 1 }, { text: 'Satırı gizle', icon: 'eyeOff', onClick: () => this.hide('r') }, { text: 'Gizlileri göster', icon: 'eye', onClick: () => this.hide('r', true) }]);
    U.menu(at, [
      { text: 'Kes', icon: 'cut', kbd: 'Ctrl+X', onClick: () => this.copy(true) }, { text: 'Kopyala', icon: 'copy', kbd: 'Ctrl+C', onClick: () => this.copy(false) },
      { text: 'Yapıştır', icon: 'paste', kbd: 'Ctrl+V', onClick: () => this.paste() }, { text: 'Yalnızca değerleri yapıştır', icon: 'paste', onClick: () => this.paste('values') },
      { sep: 1 }, { text: 'Hazır hesap…', icon: 'wand', onClick: () => this.wiz.openRecipes() }, { text: 'Renk kuralı…', icon: 'rules', onClick: () => this.rulesDialog() }, { text: 'Açılır liste…', icon: 'list', onClick: () => this.validationDialog() },
      { sep: 1 }, { text: 'Üste satır ekle', icon: 'rowAdd', onClick: () => this.insertRows() }, { text: 'Sola sütun ekle', icon: 'colAdd', onClick: () => this.insertCols() }, { text: 'Satırı sil', icon: 'rowDel', onClick: () => this.deleteRows() }, { text: 'Sütunu sil', icon: 'colDel', onClick: () => this.deleteCols() },
      { sep: 1 }, { text: 'İçeriği temizle', icon: 'eraser', kbd: 'Del', onClick: () => this.clearSel('contents') }, { text: 'Biçimi temizle', icon: 'eraser', onClick: () => this.clearSel('formats') }]);
  }

  // ---------- şerit ----------
  buildRibbon() {
    const B = (a, i, t, o) => Object.assign({ a, i, t }, o || {});
    const TABS = {
      'Giriş': [
        [B('paste', 'paste', 'Yapıştır', { menu: 1 }), B('cut', 'cut', 'Kes', { sm: 1 }), B('copy', 'copy', 'Kopyala', { sm: 1 }), B('painter', 'fill', 'Biçim kopyala', { sm: 1 })],
        [B('b', 'bold', 'Kalın', { sm: 1 }), B('i', 'italic', 'İtalik', { sm: 1 }), B('u', 'underline', 'Altı çizili', { sm: 1 }), B('st', 'strike', 'Üstü çizili', { sm: 1 }), B('fsDown', 'minus', 'Küçült', { sm: 1 }), { a: 'fsVal', html: '<span id="fsVal">11</span>', t: 'Yazı boyutu', menu: 1, sm: 1 }, B('fsUp', 'plus', 'Büyüt', { sm: 1 }), B('fc', 'color', 'Yazı rengi', { sm: 1, color: 'fc' }), B('bg', 'fill', 'Dolgu', { sm: 1, color: 'bg' })],
        [B('alignL', 'alignL', 'Sola', { sm: 1 }), B('alignC', 'alignC', 'Ortala', { sm: 1 }), B('alignR', 'alignR', 'Sağa', { sm: 1 }), B('valign', 'vMid', 'Dikey', { sm: 1, menu: 1 }), B('wrap', 'wrap', 'Metni kaydır', { sm: 1 }), B('merge', 'merge', 'Birleştir', { sm: 1 })],
        [B('nf', 'hash', 'Sayı biçimi', { menu: 1 }), B('nfTL', null, '₺ Para', { html: '<b class="glyph">₺</b>', sm: 1 }), B('nfPct', null, 'Yüzde', { html: '<b class="glyph">%</b>', sm: 1 }), B('nfThou', null, 'Binlik ayraç', { html: '<b class="glyph">000</b>', sm: 1 }), B('decUp', null, 'Ondalık artır', { html: '<b class="glyph">,0+</b>', sm: 1 }), B('decDown', null, 'Ondalık azalt', { html: '<b class="glyph">,0−</b>', sm: 1 })],
        [B('border', 'border', 'Kenarlık', { menu: 1 }), B('tableStyle', 'table', 'Tablo stili'), B('rules', 'rules', 'Renk kuralları'), B('clear', 'eraser', 'Temizle', { menu: 1 })],
      ],
      'Ekle': [
        [B('rowAdd', 'rowAdd', 'Satır ekle'), B('colAdd', 'colAdd', 'Sütun ekle'), B('rowDel', 'rowDel', 'Satır sil'), B('colDel', 'colDel', 'Sütun sil')],
        [B('chart', 'chart', 'Grafik'), B('tableStyle', 'table', 'Tablo'), B('dv', 'list', 'Açılır liste')],
        [B('sheetAdd', 'file', 'Yeni sayfa'), B('today', 'calendar', 'Bugünün tarihi'), B('now', 'clock', 'Şu anki saat')],
      ],
      'Hesapla': [
        [B('recipes', 'wand', 'Hazır Hesaplar', { big: 1 }), B('sum', 'sigma', 'Toplam', { menu: 1 }), B('fx', null, 'Tüm işlevler', { html: '<span class="fx">fx</span>' })],
        [B('rc:Koşullu', 'filter', 'Koşullu'), B('rc:Para ve Yüzde', 'tag', 'Para ve %'), B('rc:Tarih', 'calendar', 'Tarih'), B('rc:Metin', 'italic', 'Metin')],
        [B('recalc', 'redo', 'Yeniden hesapla')],
      ],
      'Veri': [
        [B('sortA', 'sortA', 'A → Z'), B('sortD', 'sortD', 'Z → A'), B('sortX', 'sortA', 'Özel sırala')],
        [B('filter', 'filter', 'Filtre'), B('filterClear', 'eraser', 'Filtreleri temizle')],
        [B('dedupe', 'layers', 'Yinelenenleri kaldır'), B('find', 'search', 'Bul / Değiştir'), B('dv', 'list', 'Açılır liste / Kural'), B('fillMenu', 'fillDown', 'Doldur', { menu: 1 })],
      ],
      'Görünüm': [
        [B('zoomOut', 'zoomOut', 'Uzaklaş', { sm: 1 }), { a: 'zoom100', html: '<span id="zoomVal">100%</span>', t: 'Yakınlaştırma', sm: 1 }, B('zoomIn', 'zoomIn', 'Yakınlaş', { sm: 1 })],
        [B('freeze', 'freeze', 'Dondur', { menu: 1 }), B('gridlines', 'grid', 'Kılavuz çizgileri'), B('selmode', 'pointer', 'Dokunarak seç'), B('fit', 'fit', 'Sütunları sığdır'), B('hideMenu', 'eyeOff', 'Gizle / göster', { menu: 1 })],
        [B('theme', 'moon', 'Tema', { menu: 1 }), B('print', 'print', 'Yazdır / PDF')],
      ],
    };
    const tabs = $('#tabs'), rib = $('#ribbon');
    Object.keys(TABS).forEach((name, i) => {
      tabs.append(h('button', { class: 'tab' + (i === 0 ? ' on' : ''), text: name, onclick: e => { $$('.tab', tabs).forEach(t => t.classList.toggle('on', t === e.currentTarget)); $$('.panel', rib).forEach(p => p.classList.toggle('on', p.dataset.tab === name)); LS.set('tab', name); } }));
      const panel = h('div', { class: 'panel' + (i === 0 ? ' on' : ''), 'data-tab': name });
      TABS[name].forEach(grp => {
        const g = h('div', { class: 'grp' });
        grp.forEach(bt => {
          const el = h('button', { class: 'rb' + (bt.sm ? ' sm' : '') + (bt.big ? ' big' : '') + (bt.menu ? ' hasmenu' : ''), 'data-a': bt.a, title: bt.t, 'aria-label': bt.t, html: (bt.html || icon(bt.i)) + (bt.color ? `<i class="cbar" data-c="${bt.color}"></i>` : '') + (bt.sm ? '' : `<span class="rl">${bt.t}</span>`) });
          el.addEventListener('mousedown', e => e.preventDefault());
          el.addEventListener('click', e => this.act(bt.a, el, e));
          g.append(el);
        });
        panel.append(g);
      });
      rib.append(panel);
    });
    const lt = LS.get('tab', null); if (lt) { const t = $$('.tab', tabs).find(x => x.textContent === lt); t && t.click(); }
    $('#zoomVal').textContent = Math.round(this.zoom * 100) + '%';
    this.colors = { fc: '#dc2626', bg: '#fef08a' }; this.updateColorBars();
  }
  updateColorBars() { $$('.cbar').forEach(b => b.style.background = this.colors[b.dataset.c]); }
  act(a, el, e) {
    const g = this.grid; if (g.editing && !['fx', 'recipes'].includes(a)) g.commitEdit();
    const wb = this.wb;
    if (a.startsWith('rc:')) return this.wiz.openRecipes(a.slice(3));
    switch (a) {
      case 'paste': return U.menu(el, [{ text: 'Yapıştır', icon: 'paste', onClick: () => this.paste() }, { text: 'Yalnızca değerler', icon: 'hash', onClick: () => this.paste('values') }, { text: 'Yalnızca biçim', icon: 'fill', onClick: () => this.paste('formats') }, { text: 'Satır/sütun çevir', icon: 'redo', onClick: () => this.paste('transpose') }]);
      case 'cut': return this.copy(true);
      case 'copy': return this.copy(false);
      case 'painter': {
        if (this._painter) { const p = this._painter; this._painter = null; el.classList.remove('on'); wb.begin(); this.forSel((si, r, c) => wb.patchCell(si, r, c, { s: Object.keys(p).length ? { ...p } : null }), 100); wb.end(); U.toast('✓ Biçim uygulandı'); return; }
        const c = wb.cell(wb.active, this.sel.r, this.sel.c); this._painter = (c && c.s) || {}; el.classList.add('on');
        U.toast('Biçim alındı — hedef hücreleri seçip “Biçim kopyala”ya tekrar basın'); return;
      }
      case 'b': case 'i': case 'u': case 'st': case 'wrap': return this.toggle(a);
      case 'fsUp': return this.fontSize(1); case 'fsDown': return this.fontSize(-1);
      case 'fsVal': return U.menu(el, [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map(n => ({ text: String(n), onClick: () => this.styleSel({ fs: n === 11 ? null : n }) })));
      case 'fc': case 'bg': return U.colorPicker(el, c => { if (c) { this.colors[a] = c; this.updateColorBars(); } this.styleSel({ [a]: c }); }, a === 'fc' ? 'Otomatik (siyah)' : 'Dolgu yok');
      case 'alignL': case 'alignC': case 'alignR': { const v = { alignL: 'left', alignC: 'center', alignR: 'right' }[a]; const st = (wb.cell(wb.active, this.sel.r, this.sel.c) || {}).s || {}; return this.styleSel({ ha: st.ha === v ? null : v }); }
      case 'valign': return U.menu(el, [{ text: 'Üste', icon: 'vTop', onClick: () => this.styleSel({ va: 'top' }) }, { text: 'Ortaya', icon: 'vMid', onClick: () => this.styleSel({ va: 'middle' }) }, { text: 'Alta', icon: 'vBot', onClick: () => this.styleSel({ va: 'bottom' }) }]);
      case 'merge': return this.merge();
      case 'nf': return U.menu(el, [...NF_MENU.map(([t, code, ex]) => ({ text: t + '   ' + ex, onClick: () => this.setNF(code) })), { sep: 1 }, { text: 'Özel biçim…', icon: 'hash', onClick: () => this.customNF() }]);
      case 'nfTL': return this.setNF('#,##0.00 ₺'); case 'nfPct': return this.setNF('0%'); case 'nfThou': return this.setNF('#,##0');
      case 'decUp': return this.changeDecimals(1); case 'decDown': return this.changeDecimals(-1);
      case 'border': return U.menu(el, [{ text: 'Tüm kenarlıklar', icon: 'grid', onClick: () => this.borders('all') }, { text: 'Dış kenarlık', icon: 'border', onClick: () => this.borders('outer') }, { text: 'Kalın dış kenarlık', icon: 'border', onClick: () => this.borders('thick') }, { text: 'Alt kenarlık', icon: 'minus', onClick: () => this.borders('bottom') }, { text: 'Çift alt (toplam)', icon: 'minus', onClick: () => this.borders('double') }, { text: 'Üst kenarlık', icon: 'minus', onClick: () => this.borders('top') }, { text: 'Sol', icon: 'alignL', onClick: () => this.borders('left') }, { text: 'Sağ', icon: 'alignR', onClick: () => this.borders('right') }, { sep: 1 }, { text: 'Kenarlık rengi…', swatch: this.penColor || '#000', onClick: () => U.colorPicker(el, c => { this.penColor = c || '#000000'; }, 'Siyah') }, { text: 'Kenarlık yok', icon: 'x', onClick: () => this.borders('none') }]);
      case 'tableStyle': return this.tableStyleDialog();
      case 'rules': return this.rulesDialog();
      case 'clear': return U.menu(el, [{ text: 'İçeriği temizle', icon: 'eraser', onClick: () => this.clearSel('contents') }, { text: 'Biçimi temizle', icon: 'fill', onClick: () => this.clearSel('formats') }, { text: 'Tümünü temizle', icon: 'trash', danger: 1, onClick: () => this.clearSel('all') }]);
      case 'rowAdd': return this.insertRows(); case 'colAdd': return this.insertCols(); case 'rowDel': return this.deleteRows(); case 'colDel': return this.deleteCols();
      case 'chart': return this.chartDialog();
      case 'dv': return this.validationDialog();
      case 'sheetAdd': wb.addSheet(); return this.refreshSheet();
      case 'today': case 'now': { const s = this.sel; const v = a === 'today' ? Math.floor(X.nowSerial()) : X.nowSerial(); wb.begin(); wb.patchCell(wb.active, s.r, s.c, { v, f: null }); wb.setStyle(wb.active, s.r, s.c, { nf: a === 'today' ? 'dd.mm.yyyy' : 'hh:mm' }); wb.end(); return; }
      case 'recipes': return this.wiz.openRecipes();
      case 'fx': return this.wiz.openFunctions();
      case 'sum': return U.menu(el, [['SUM', 'Toplam', 'sigma'], ['AVERAGE', 'Ortalama', 'calc'], ['COUNT', 'Say', 'hash'], ['MAX', 'En büyük', 'sortD'], ['MIN', 'En küçük', 'sortA']].map(([f, t, i]) => ({ text: t, icon: i, onClick: () => this.autoSum(f) })));
      case 'recalc': wb.recalc(); this.grid.render(); return U.toast('Yeniden hesaplandı');
      case 'sortA': return this.quickSort(true); case 'sortD': return this.quickSort(false); case 'sortX': return this.sortDialog();
      case 'filter': return this.toggleFilter(); case 'filterClear': return this.clearFilters();
      case 'dedupe': return this.removeDuplicates(); case 'find': return this.findReplace();
      case 'fillMenu': return U.menu(el, [{ text: 'Aşağı doldur', icon: 'fillDown', kbd: 'Ctrl+D', onClick: () => this.fill('down') }, { text: 'Sağa doldur', icon: 'right', kbd: 'Ctrl+R', onClick: () => this.fill('right') }, { text: 'Seri doldur (1, 2, 3…)', icon: 'hash', onClick: () => this.fill('series') }]);
      case 'zoomIn': return this.setZoom(this.zoom + 0.1); case 'zoomOut': return this.setZoom(this.zoom - 0.1); case 'zoom100': return this.setZoom(1);
      case 'freeze': return U.menu(el, [{ text: 'Üst satırı dondur', onClick: () => this.freeze('row') }, { text: 'İlk sütunu dondur', onClick: () => this.freeze('col') }, { text: 'Seçili hücreye göre dondur', onClick: () => this.freeze('sel') }, { text: 'Dondurmayı kaldır', onClick: () => this.freeze('none') }]);
      case 'gridlines': wb.begin(); wb.snapSheet(wb.active); wb.sheet.grid = wb.sheet.grid === false; wb.end(); return;
      case 'selmode': this.selectMode = !this.selectMode; this.onSelChange(); return U.toast(this.selectMode ? 'Dokunarak seçme açık: sürükleyince hücre seçilir (kaydırmak için iki parmak)' : 'Dokunarak seçme kapalı: sürükleyince sayfa kayar');
      case 'fit': return this.autoFit();
      case 'hideMenu': return U.menu(el, [{ text: 'Satırları gizle', onClick: () => this.hide('r') }, { text: 'Satırları göster', onClick: () => this.hide('r', true) }, { text: 'Sütunları gizle', onClick: () => this.hide('c') }, { text: 'Sütunları göster', onClick: () => this.hide('c', true) }]);
      case 'theme': return U.menu(el, [['auto', 'Sistem'], ['light', 'Açık'], ['dark', 'Koyu']].map(([k, t]) => ({ text: t, on: this.theme === k, icon: this.theme === k ? 'check' : null, onClick: () => this.applyTheme(k) })));
      case 'print': return IO.print(wb);
    }
  }
  autoSum(fn) {
    const wb = this.wb, si = wb.active, s = this.sel; const g = this.selBounds();
    if (g.r1 !== g.r2 || g.c1 !== g.c2) {
      wb.begin();
      if (g.r2 > g.r1 || g.c1 === g.c2) { for (let c = g.c1; c <= g.c2; c++) wb.patchCell(si, g.r2 + 1, c, { f: `${fn}(${X.addr(g.r1, c)}:${X.addr(g.r2, c)})`, v: null }); wb.end(); this.setSel({ r: g.r2 + 1, c: g.c1, r1: g.r2 + 1, c1: g.c1, r2: g.r2 + 1, c2: g.c2 }); }
      else { wb.patchCell(si, g.r1, g.c2 + 1, { f: `${fn}(${X.addr(g.r1, g.c1)}:${X.addr(g.r1, g.c2)})`, v: null }); wb.end(); this.setSel({ r: g.r1, c: g.c2 + 1, r1: g.r1, c1: g.c2 + 1, r2: g.r1, c2: g.c2 + 1 }); }
      return;
    }
    // yukarıdaki veya soldaki sayıları bul
    const isN = (r, c) => typeof wb.value(si, r, c) === 'number';
    let r = s.r - 1; while (r >= 0 && !isN(r, s.c) && wb.value(si, r, s.c) == null) r--; let r0 = r; while (r0 >= 0 && isN(r0, s.c)) r0--;
    let f;
    if (r >= 0 && isN(r, s.c)) f = `${fn}(${X.addr(r0 + 1, s.c)}:${X.addr(r, s.c)})`;
    else { let c = s.c - 1; while (c >= 0 && isN(s.r, c)) c--; if (c < s.c - 1) f = `${fn}(${X.addr(s.r, c + 1)}:${X.addr(s.r, s.c - 1)})`; }
    if (!f) { this.wiz.openFunctions(fn); return; }
    wb.begin(); wb.patchCell(si, s.r, s.c, { f, v: null }); wb.end(); U.toast('✓ ' + X.FN[fn].tr + ' eklendi');
  }
  customNF() {
    const cur = this.wb.numFmt(this.wb.active, this.sel.r, this.sel.c) || '';
    const v = this.wb.value(this.wb.active, this.sel.r, this.sel.c);
    const inp = h('input', { class: 'inp', value: cur, placeholder: 'Ör: #,##0.00 "adet"' }); const pv = h('div', { class: 'pv-res' });
    const upd = () => { pv.innerHTML = ''; pv.append(h('span', { text: 'Örnek' }), h('b', { text: X.formatValue(typeof v === 'number' ? v : 1234.5, inp.value || null) })); };
    inp.addEventListener('input', upd);
    const body = h('div', { class: 'form' }, h('p', { class: 'hint', html: 'Kodlar: <code>0</code> rakam, <code>#,##0</code> binlik, <code>0.00</code> ondalık, <code>%</code>, <code>dd.mm.yyyy</code> tarih, <code>"metin"</code>' }), inp, pv);
    U.dialog({ title: 'Özel sayı biçimi', body, buttons: [{ text: 'Vazgeç', value: 0 }, { text: 'Uygula', kind: 'primary', value: 1, onClick: () => this.setNF(inp.value.trim() || null) }] }); upd();
  }

  // ---------- üst çubuk, formül çubuğu, klavye ----------
  bindTop() {
    $('#btnUndo').onclick = () => { this.grid.editing && this.grid.cancelEdit(); this.wb.undo(); };
    $('#btnRedo').onclick = () => { this.grid.editing && this.grid.cancelEdit(); this.wb.redo(); };
    $('#btnSave').onclick = () => this.save();
    $('#btnMenu').onclick = () => this.fileMenu();
    $('#btnAddSheet').onclick = () => { this.wb.addSheet(); this.refreshSheet(); };
    const dn = $('#docName'); dn.addEventListener('change', () => { this.wb.name = dn.value.trim() || 'Çalışma'; this.wb.dirty = true; this.scheduleSave(); });
    dn.addEventListener('keydown', e => { if (e.key === 'Enter') dn.blur(); });
    $('#stats').onclick = e => { if (!this.stats) return; const s = this.stats; const f = v => X.formatValue(v, '#,##0.####'); U.menu(e.currentTarget, [{ head: 'Seçim özeti' }, { text: 'Toplam: ' + f(s.sum), icon: 'sigma', onClick: () => this.autoSum('SUM') }, { text: 'Ortalama: ' + f(s.avg), icon: 'calc', onClick: () => this.autoSum('AVERAGE') }, { text: 'Sayı adedi: ' + s.n, icon: 'hash', onClick: () => this.autoSum('COUNT') }, { text: 'Dolu hücre: ' + s.cnt, icon: 'hash' }, { text: 'En küçük: ' + (s.n ? f(s.mn) : '-'), icon: 'sortA', onClick: () => this.autoSum('MIN') }, { text: 'En büyük: ' + (s.n ? f(s.mx) : '-'), icon: 'sortD', onClick: () => this.autoSum('MAX') }]); };
    this.nameBox.onclick = async () => { const v = await U.prompt('Git', this.nameBox.textContent, 'Hücre ya da alan (ör: B5 veya A1:D20)'); if (!v) return; const g = IO.refToRange(v.trim().toUpperCase()); if (g) this.setSel({ r: g.r1, c: g.c1, ...g }, true); else U.toast('Geçersiz adres', 'warn'); };
    $('#btnFx').onclick = () => this.wiz.openFunctions();
    $('#btnRecipes').onclick = () => this.wiz.openRecipes();
    $('#fOk').onclick = () => this.grid.commitEdit();
    $('#fCancel').onclick = () => this.grid.cancelEdit();
    // dosya bırakma
    addEventListener('dragover', e => { e.preventDefault(); });
    addEventListener('drop', e => { e.preventDefault(); const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) this.openFileObj(f); });
  }
  bindFbar() {
    const fi = this.fInput, g = this.grid;
    fi.addEventListener('focus', () => { if (this.pick && this.pick.active) { fi.blur(); return; } if (!g.editing) { g.startEdit(null, 'edit', 'fbar'); } else g.editing.src = 'fbar'; });
    fi.addEventListener('input', () => { g._refIns = null; g.syncInputs(fi); g.updateRefMarks(); });
    fi.addEventListener('keydown', e => g.editorKey(e, fi));
    fi.addEventListener('blur', () => { setTimeout(() => { if (!g.editing || g.editing.src !== 'fbar') return; const ae = document.activeElement; if (ae === fi || ae === g.editor) return; if (U.topDialog()) return; g.commitEdit(); }, 150); });
  }
  onEditStart() { $('#fbar').classList.add('editing'); }
  bindKeys() {
    document.addEventListener('keydown', e => {
      const g = this.grid; const a = document.activeElement;
      const inField = a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT' || a.isContentEditable);
      if (U.topDialog() || inField) { if (e.key === 'Escape' && U.topDialog() && !inField) U.topDialog().close(null); return; }
      if (this.pick && this.pick.active) { if (e.key === 'Enter') $('#pickOk').click(); if (e.key === 'Escape') $('#pickCancel').click(); }
      const k = e.key, ctrl = e.ctrlKey || e.metaKey;
      if (ctrl) {
        const kl = k.toLowerCase();
        const map = { z: () => this.wb.undo(), y: () => this.wb.redo(), b: () => this.toggle('b'), i: () => this.toggle('i'), u: () => this.toggle('u'), s: () => this.save(), o: () => this.openFile(), f: () => this.findReplace(), h: () => this.findReplace(true), d: () => this.fill('down'), r: () => this.fill('right'), p: () => IO.print(this.wb), a: () => this.setSel({ r: this.sel.r, c: this.sel.c, r1: 0, c1: 0, r2: g.nR - 1, c2: g.nC - 1 }) };
        if (kl === 'z' && e.shiftKey) { e.preventDefault(); this.wb.redo(); return; }
        if (map[kl]) { e.preventDefault(); map[kl](); return; }
        if (k === 'Home') { e.preventDefault(); this.setSel({ r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 }, true); return; }
        if (k === 'End') { e.preventDefault(); const d = this.wb.dims(this.wb.active); const r = Math.max(0, d.rows - 1), c = Math.max(0, d.cols - 1); this.setSel({ r, c, r1: r, c1: c, r2: r, c2: c }, true); return; }
        if (/^Arrow/.test(k)) { e.preventDefault(); this.moveSel(k.slice(5).toLowerCase(), e.shiftKey, true); return; }
        if (k === '+' || k === '=') { e.preventDefault(); this.setZoom(this.zoom + 0.1); return; } if (k === '-') { e.preventDefault(); this.setZoom(this.zoom - 0.1); return; } if (k === '0') { e.preventDefault(); this.setZoom(1); return; }
        return;
      }
      if (/^Arrow/.test(k)) { e.preventDefault(); this.moveSel(k.slice(5).toLowerCase(), e.shiftKey); return; }
      if (k === 'Enter') { e.preventDefault(); this.moveSel(e.shiftKey ? 'up' : 'down'); return; }
      if (k === 'Tab') { e.preventDefault(); this.moveSel(e.shiftKey ? 'left' : 'right'); return; }
      if (k === 'PageDown') { e.preventDefault(); this.moveSel('pgdn'); return; } if (k === 'PageUp') { e.preventDefault(); this.moveSel('pgup'); return; }
      if (k === 'Home') { e.preventDefault(); this.setSel({ r: this.sel.r, c: 0, r1: this.sel.r, c1: 0, r2: this.sel.r, c2: 0 }, true); return; }
      if (k === 'F2') { e.preventDefault(); g.startEdit(null, 'edit'); return; }
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); this.clearSel('contents'); return; }
      if (k === 'Escape') { this.clip = null; this.clipData = null; g.render(); return; }
      if (k.length === 1 && !e.altKey) { e.preventDefault(); g.startEdit(k, 'enter'); }
    });
  }

  // ---------- dosya ----------
  fileMenu() {
    const wb = this.wb;
    const body = h('div', { class: 'drawer' });
    const item = (ic, t, sub, fn) => h('button', { class: 'ditem', onclick: () => { d.close(); fn(); } }, h('span', { class: 'dic', html: icon(ic) }), h('span', null, h('b', { text: t }), sub ? h('small', { text: sub }) : null));
    body.append(
      item('file', 'Yeni', 'Boş çalışma kitabı', () => this.newDoc()),
      item('folder', 'Aç', 'Excel (xlsx, xls, xltx, xlt), ODS, CSV, Numbers', () => this.openFile()),
      item('save', 'Kaydet', (wb.format || 'xlsx').toUpperCase() + ' olarak', () => this.save()),
      h('div', { class: 'dhead', text: 'Farklı kaydet' }),
      h('div', { class: 'fmtgrid' }, ...['xlsx', 'xls', 'xltx', 'xlt', 'ods', 'csv'].map(f => h('button', { class: 'fmt', onclick: () => { d.close(); this.saveAs(f); } }, h('b', { text: '.' + f }), h('small', { text: IO.FORMATS[f].name.replace(/\s*\(.*\)/, '') })))),
      item('print', 'Yazdır / PDF', 'PDF olarak kaydetmek için yazdırma ekranında “PDF” seçin', () => IO.print(wb)),
      item('table', 'Örnek tablo', 'Özellikleri denemek için hazır bir tablo', () => this.sample()),
      item('layers', 'Modern görünüme geç', 'Aynı çalışma; kayıt listesi, kartlar, özet ve grafik', async () => { await this.saveLocal(true); LS.set('ui', 'modern'); location.href = 'index.html'; }),
    );
    const recent = h('div', { class: 'recent' }); body.append(h('div', { class: 'dhead', text: 'Son çalışmalar (bu cihazda)' }), recent);
    DB.all().then(list => {
      list.sort((a, b) => b.updated - a.updated);
      if (!list.length) recent.append(h('p', { class: 'hint', text: 'Henüz yok.' }));
      for (const doc of list.slice(0, 30)) recent.append(h('div', { class: 'ritem' + (doc.id === this.docId ? ' on' : '') }, h('button', { class: 'rname', onclick: () => { d.close(); this.openDoc(doc.id); } }, h('b', { text: doc.name }), h('small', { text: new Date(doc.updated).toLocaleString('tr-TR') })), h('button', { class: 'ic ghost', title: 'Sil', html: icon('trash'), onclick: async () => { if (doc.id === this.docId) { U.toast('Açık olan çalışma silinemez'); return; } await DB.del(doc.id); recent.querySelector(`[data-x]`); d.close(); this.fileMenu(); } })));
    }).catch(() => recent.append(h('p', { class: 'hint', text: 'Kayıt alanına erişilemedi.' })));
    body.append(h('p', { class: 'about', html: 'Excello · internetsiz çalışır · tüm veriler yalnızca bu cihazda saklanır' }));
    const d = U.dialog({ title: '<span class="logo-sm">E</span> Dosya', body, cls: 'leftdrawer' });
  }
  async newDoc() { this.saveLocal(true); const wb = new M.Workbook(); wb.format = 'xlsx'; this.docId = this.newId(); LS.set('cur', this.docId); this.fileHandle = null; this.attach(wb); this.saveLocal(true); }
  async openDoc(id) { await this.saveLocal(true); const doc = await DB.get(id); if (!doc) return; const wb = M.Workbook.fromJSON(doc.data); wb.format = doc.format || 'xlsx'; this.docId = id; LS.set('cur', id); this.fileHandle = null; this.attach(wb); }
  async openFile() {
    if (G.showOpenFilePicker && !isIOS) {
      try { const [hd] = await G.showOpenFilePicker({ types: [{ description: 'Tablolar', accept: { 'application/octet-stream': IO.OPEN_ACCEPT.split(',') } }], excludeAcceptAllOption: false }); const f = await hd.getFile(); await this.openFileObj(f, hd); return; }
      catch (e) { if (e.name === 'AbortError') return; }
    }
    const inp = h('input', { type: 'file', accept: IO.OPEN_ACCEPT, style: { display: 'none' } });
    inp.onchange = () => { const f = inp.files[0]; if (f) this.openFileObj(f); inp.remove(); };
    document.body.append(inp); inp.click();
  }
  async openFileObj(f, handle) {
    const busy = this.busy('Açılıyor: ' + f.name);
    try { await this.saveLocal(true); const wb = await IO.open(f); this.docId = this.newId(); LS.set('cur', this.docId); this.attach(wb); this.fileHandle = handle || null; this.handleFmt = wb.format; this.saveLocal(true); U.toast('✓ ' + f.name + ' açıldı'); }
    catch (e) { console.error(e); U.alert('Dosya açılamadı: ' + (e.message || e), 'Hata'); }
    finally { busy(); }
  }
  async openFromUrl(u) {
    try { const blob = await new Promise((res, rej) => { const x = new XMLHttpRequest(); x.open('GET', u); x.responseType = 'blob'; x.onload = () => (x.status === 0 || x.status === 200) ? res(x.response) : rej(new Error(x.status)); x.onerror = () => rej(new Error('okunamadı')); x.send(); }); const name = decodeURIComponent(u.split(/[\\/]/).pop()); await this.openFileObj(new File([blob], name)); }
    catch (e) { U.alert('Dosya açılamadı: ' + u, 'Hata'); }
  }
  busy(msg) { const el = h('div', { class: 'busy' }, h('div', { class: 'spin' }), h('span', { text: msg })); document.body.append(el); return () => el.remove(); }
  async save() {
    if (this.grid.editing) this.grid.commitEdit();
    const fmt = this.wb.format || 'xlsx';
    if (this.fileHandle && this.handleFmt === fmt) {
      const busy = this.busy('Kaydediliyor…');
      try { const res = await IO.save(this.wb, fmt); const w = await this.fileHandle.createWritable(); await w.write(res.blob); await w.close(); this.wb.dirty = false; U.toast('✓ Kaydedildi: ' + this.fileHandle.name); return; }
      catch (e) { console.warn(e); } finally { busy(); }
    }
    return this.saveAs(fmt);
  }
  async saveAs(fmt) {
    if (this.grid.editing) this.grid.commitEdit();
    const F = IO.FORMATS[fmt]; const name = (this.wb.name || 'Excello') + '.' + F.ext;
    let handle = null;
    if (G.showSaveFilePicker && !isIOS) {
      try { handle = await G.showSaveFilePicker({ suggestedName: name, types: [{ description: F.name, accept: { [F.mime]: ['.' + F.ext] } }] }); }
      catch (e) { if (e.name === 'AbortError') return; handle = null; }
    }
    const busy = this.busy('Hazırlanıyor…'); let res;
    try { res = await IO.save(this.wb, fmt); } catch (e) { busy(); console.error(e); U.alert('Kaydedilemedi: ' + (e.message || e), 'Hata'); return; }
    busy();
    if (handle) {
      try { const w = await handle.createWritable(); await w.write(res.blob); await w.close(); this.fileHandle = handle; this.handleFmt = fmt; this.wb.format = fmt; this.wb.name = handle.name.replace(/\.[^.]+$/, ''); $('#docName').value = this.wb.name; this.wb.dirty = false; this.saveLocal(true); U.toast('✓ Kaydedildi: ' + handle.name); return; }
      catch (e) { console.warn(e); }
    }
    if (fmt !== 'csv') { this.wb.format = fmt; this.saveLocal(true); }
    this.deliver(res.blob, res.name, res.mime);
  }
  deliver(blob, name, mime) {
    // iPad/telefon: paylaş menüsü (Dosyalar'a kaydet); bilgisayar: indir
    const file = new File([blob], name, { type: mime });
    const canShare = navigator.canShare && navigator.canShare({ files: [file] });
    const download = () => { const url = URL.createObjectURL(blob); const a = h('a', { href: url, download: name }); document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000); };
    if (canShare && isMobile()) {
      const d = U.dialog({ title: icon('check') + ' Dosya hazır', body: h('div', { class: 'form' }, h('p', { text: name }), h('p', { class: 'hint', text: '“Paylaş” ile Dosyalar uygulamasına kaydedebilir, WhatsApp veya e-posta ile gönderebilirsiniz.' })), buttons: [{ text: 'İndir', onClick: () => { download(); }, value: 1 }, { text: 'Paylaş / Kaydet', kind: 'primary', onClick: () => { navigator.share({ files: [file], title: name }).catch(() => { }); }, value: 1 }] });
      return;
    }
    download(); U.toast('✓ İndirildi: ' + name);
  }
  scheduleSave() { clearTimeout(this._saveT); const s = $('#saveState'); s.textContent = 'Kaydediliyor…'; s.classList.add('dirty'); this._saveT = setTimeout(() => this.saveLocal(), 1200); }
  async saveLocal(now) {
    clearTimeout(this._saveT);
    if (!this.docId) return;
    try { await DB.put({ id: this.docId, name: this.wb.name || 'Çalışma', updated: Date.now(), data: this.wb.toJSON(), format: this.wb.format || 'xlsx' }); LS.set('cur', this.docId); $('#saveState').textContent = 'Kaydedildi'; $('#saveState').classList.remove('dirty'); }
    catch (e) { $('#saveState').textContent = 'Yerel kayıt yapılamadı'; }
  }
  welcome() {
    LS.set('welcomed', true);
    const body = h('div', { class: 'welcome' },
      h('div', { class: 'wlogo' }, h('span', { class: 'logo' }, 'E'), h('div', null, h('b', { text: 'Excello\'ya hoş geldiniz' }), h('small', { text: 'İnternetsiz çalışan, formül yazdırmayan tablo uygulaması' }))),
      h('ul', null,
        h('li', { html: icon('wand') + '<span><b>Hazır Hesaplar</b> — Toplam, KDV, yüzde, tarih farkı, “eğer… ise” ve daha fazlası. Hücreleri dokunarak seçin, gerisini Excello yapsın.</span>' }),
        h('li', { html: icon('pointer') + '<span><b>Dokun ve seç</b> — Hücreye dokunun; tekrar dokunursanız yazarsınız. Basılı tutup sürükleyerek alan seçin.</span>' }),
        h('li', { html: icon('chart') + '<span><b>Grafik, renk kuralları, tablo stilleri</b> — Hepsi menüde, tek dokunuşla.</span>' }),
        h('li', { html: icon('save') + '<span><b>Excel ile uyumlu</b> — xlsx, xls, xltx/xlt şablon, ods, csv açar ve kaydeder. Çalışmanız otomatik saklanır.</span>' })));
    U.dialog({ title: '', body, buttons: [{ text: 'Örnek tabloyla başla', onClick: () => this.sample(), value: 1 }, { text: 'Boş başla', kind: 'primary', value: 1 }] });
  }
  sample() {
    const wb = new M.Workbook(); wb.name = 'Örnek Satış Tablosu'; const s = wb.sheets[0]; s.name = 'Satışlar';
    const rows = [['Tarih', 'Ürün', 'Şehir', 'Adet', 'Birim Fiyat', 'Tutar'], ['01.10.2026', 'Defter', 'Ankara', 12, 45], ['02.10.2026', 'Kalem', 'İzmir', 40, 12.5], ['03.10.2026', 'Çanta', 'Ankara', 3, 650], ['05.10.2026', 'Defter', 'İstanbul', 25, 45], ['07.10.2026', 'Silgi', 'İzmir', 60, 6], ['09.10.2026', 'Çanta', 'İstanbul', 5, 650], ['12.10.2026', 'Kalem', 'Ankara', 30, 12.5], ['15.10.2026', 'Defter', 'İzmir', 18, 45]];
    rows.forEach((row, r) => row.forEach((v, c) => { const pi = typeof v === 'string' ? IO.parseInput(v) : { v }; s.cells.set(M.key(r, c), { v: pi.v, f: null, s: pi.nf ? { nf: pi.nf } : null }); }));
    for (let r = 1; r < rows.length; r++) { s.cells.set(M.key(r, 5), { v: null, f: `D${r + 1}*E${r + 1}`, s: { nf: '#,##0.00 ₺' } }); s.cells.get(M.key(r, 4)).s = { nf: '#,##0.00 ₺' }; }
    const H = '#16a34a';
    for (let c = 0; c < 6; c++) s.cells.get(M.key(0, c)).s = { b: 1, bg: H, fc: '#ffffff', ha: 'center' };
    s.cells.set(M.key(9, 4), { v: 'Toplam', f: null, s: { b: 1, ha: 'right' } }); s.cells.set(M.key(9, 5), { v: null, f: 'SUM(F2:F9)', s: { b: 1, nf: '#,##0.00 ₺', bd: { t: { s: 'medium', c: H } } } });
    s.colW = { 0: 100, 1: 96, 2: 96, 3: 70, 4: 110, 5: 120 };
    s.cf.push({ type: 'bar', range: { r1: 1, c1: 3, r2: 8, c2: 3 }, color: '#4f8ef7', style: {} });
    s.cf.push({ type: 'gt', a: 1000, range: { r1: 1, c1: 5, r2: 8, c2: 5 }, style: { bg: '#dcfce7', fc: '#15803d' } });
    s.dv.push({ type: 'list', range: { r1: 1, c1: 2, r2: 200, c2: 2 }, items: ['Ankara', 'İstanbul', 'İzmir', 'Bursa'], msg: '' });
    s.charts.push({ id: 'c1', type: 'column', range: { r1: 0, c1: 1, r2: 8, c2: 1 }, title: '', labels: true, x: 640, y: 0, w: 440, h: 280 });
    s.charts[0].range = { r1: 0, c1: 5, r2: 8, c2: 5 }; s.charts[0].lr = { r1: 0, c1: 1, r2: 8, c2: 1 }; s.charts[0].title = 'Satış tutarları';
    s.freezeR = 1;
    const u = new M.Sheet('Ürünler'); wb.sheets.push(u);
    [['Ürün', 'Kategori', 'Stok', 'Tedarikçi'], ['Defter', 'Kırtasiye', 120, 'Ada Kağıt'], ['Kalem', 'Kırtasiye', 300, 'Mavi Yazı'], ['Çanta', 'Aksesuar', 15, 'Deri Dünyası'], ['Silgi', 'Kırtasiye', 200, 'Mavi Yazı']]
      .forEach((row, r) => row.forEach((v, c) => u.cells.set(M.key(r, c), { v, f: null, s: r ? null : { b: 1, bg: H, fc: '#ffffff', ha: 'center' } })));
    u.colW = { 0: 100, 1: 110, 2: 80, 3: 130 }; u.freezeR = 1; wb.active = 0;
    wb.recalc(); wb.format = 'xlsx';
    this.saveLocal(true); this.docId = this.newId(); LS.set('cur', this.docId); this.attach(wb); this.saveLocal(true);
  }
}

// ---------- seri üretimi (otomatik doldurma) ----------
const LISTS = [['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'], ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara'], ['Pazartesi', 'Salı', 'Çarşamba', 'Perşembe', 'Cuma', 'Cumartesi', 'Pazar'], ['Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt', 'Paz'], ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], ['1. Çeyrek', '2. Çeyrek', '3. Çeyrek', '4. Çeyrek']];
function seriesGen(cells, nf) {
  const vals = cells.map(c => c ? c.v : null); const L = vals.length;
  if (vals.every(v => typeof v === 'number')) {
    if (L === 1) { if (X.isDateFormat(nf)) return i => vals[0] + i; return i => vals[0]; }
    const step = (vals[L - 1] - vals[0]) / (L - 1); return i => +(vals[0] + step * i).toPrecision(15);
  }
  if (vals.every(v => typeof v === 'string')) {
    const lo = vals.map(v => v.toLocaleLowerCase('tr-TR'));
    for (const lst of LISTS) { const ll = lst.map(x => x.toLocaleLowerCase('tr-TR')); const idx = lo.map(v => ll.indexOf(v)); if (idx.every(i => i >= 0)) { const step = L > 1 ? (idx[1] - idx[0]) : 1; const up = vals[0] === vals[0].toLocaleUpperCase('tr-TR') && vals[0].length > 1; return i => { const k = ((idx[0] + step * i) % lst.length + lst.length) % lst.length; return up ? lst[k].toLocaleUpperCase('tr-TR') : lst[k]; }; } }
    const m = vals.map(v => /^(.*?)(\d+)(\D*)$/.exec(v));
    if (m.every(x => x) && m.every(x => x[1] === m[0][1] && x[3] === m[0][3])) { const ns = m.map(x => +x[2]); const step = L > 1 ? (ns[L - 1] - ns[0]) / (L - 1) : 1; const w = m[0][2].length; return i => m[0][1] + String(Math.round(ns[0] + step * i)).padStart(m[0][2].startsWith('0') ? w : 0, '0') + m[0][3]; }
  }
  return i => vals[((i % L) + L) % L];
}
function range(a, b) { const o = []; for (let i = a; i <= b; i++) o.push(i); return o; }
function parseTSV(txt) {
  const rows = []; let row = [], cur = '', q = false; const t = txt.replace(/\r\n?/g, '\n');
  for (let i = 0; i < t.length; i++) { const ch = t[i]; if (q) { if (ch === '"') { if (t[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; } if (ch === '"' && cur === '') { q = true; continue; } if (ch === '\t') { row.push(cur); cur = ''; } else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; } else cur += ch; }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

addEventListener('DOMContentLoaded', () => { G.app = new App(); });
})(window);
