// Excello tablo görünümü: tuval çizimi, dokunma/fare/klavye, hücre düzenleyici
(function (G) {
'use strict';
const X = G.XF, M = G.XM;
const DEF_W = 88, DEF_H = 26, HW = 46, HH = 26;
const isTouchDev = () => matchMedia('(pointer: coarse)').matches;

function bsearch(arr, v, n) { let lo = 0, hi = n - 1; if (v <= 0) return 0; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (arr[m] <= v) lo = m; else hi = m - 1; } return lo; }

class Grid {
  constructor(app, wrap) {
    this.app = app; this.wrap = wrap;
    this.cv = wrap.querySelector('#cv'); this.ctx = this.cv.getContext('2d');
    this.scroller = wrap.querySelector('#scroller'); this.spacer = wrap.querySelector('#spacer');
    this.editor = wrap.querySelector('#editor');
    this.extraR = 0; this.extraC = 0; this.geoKey = ''; this.colX = [0]; this.rowY = [0];
    this.editing = null; // {r,c,mode,src}
    this.readTheme();
    this.scroller.addEventListener('scroll', () => { this.growIfNeeded(); this.render(); this.app.positionCharts && this.app.positionCharts(); this.placeEditor(); }, { passive: true });
    new ResizeObserver(() => { this.resize(); }).observe(wrap);
    this.bindMouse(); this.bindTouch(); this.bindEditor();
    this.resize();
  }
  get wb() { return this.app.wb; }
  get sh() { return this.app.wb.sheet; }
  get si() { return this.app.wb.active; }
  get z() { return this.app.zoom; }
  readTheme() {
    const cs = getComputedStyle(document.documentElement); const g = n => cs.getPropertyValue(n).trim();
    this.col = { bg: g('--cv-bg') || '#fff', fg: g('--cv-fg') || '#111', grid: g('--cv-grid') || '#e4e4e7', head: g('--cv-head') || '#f4f4f5', headFg: g('--cv-head-fg') || '#52525b', headSel: g('--cv-head-sel') || '#dcfce7', accent: g('--accent') || '#16a34a', selFill: g('--cv-sel') || 'rgba(22,163,74,.10)', err: '#dc2626', dark: document.documentElement.classList.contains('dark') };
    this.render();
  }
  resize() {
    const r = this.wrap.getBoundingClientRect(); this.W = r.width; this.H = r.height; const dpr = G.devicePixelRatio || 1;
    this.cv.width = Math.round(this.W * dpr); this.cv.height = Math.round(this.H * dpr); this.cv.style.width = this.W + 'px'; this.cv.style.height = this.H + 'px';
    this.dpr = dpr; this.geoKey = ''; this.render();
  }
  // ---------- geometri ----------
  hw() { return Math.round(HW * Math.min(1.3, Math.max(0.8, this.z))); }
  hh() { return Math.round(HH * Math.min(1.3, Math.max(0.8, this.z))); }
  cw(c) { return this.sh.hiddenC.has(c) ? 0 : Math.round((this.sh.colW[c] || DEF_W) * this.z); }
  rh(r) { return this.wb.isRowHidden(this.si, r) ? 0 : Math.round((this.sh.rowH[r] || DEF_H) * this.z); }
  geo() {
    const d = this.wb.dims(this.si);
    const nR = Math.max(200, d.rows + 60, this.sh.freezeR + 50) + this.extraR, nC = Math.max(30, d.cols + 10) + this.extraC;
    const k = [this.si, this.wb.ver, this.z, nR, nC, this.app.geoVer || 0].join('|');
    if (k === this.geoKey) return;
    this.geoKey = k; this.nR = nR; this.nC = nC;
    const cx = new Float64Array(nC + 1); for (let c = 0; c < nC; c++) cx[c + 1] = cx[c] + this.cw(c);
    const ry = new Float64Array(nR + 1); for (let r = 0; r < nR; r++) ry[r + 1] = ry[r] + this.rh(r);
    this.colX = cx; this.rowY = ry;
    this.fr = Math.min(this.sh.freezeR, nR); this.fc = Math.min(this.sh.freezeC, nC);
    const w = this.hw() + cx[nC] + 40, h = this.hh() + ry[nR] + 60;
    if (this._sw !== w || this._sh !== h) { this.spacer.style.width = w + 'px'; this.spacer.style.height = h + 'px'; this._sw = w; this._sh = h; }
  }
  growIfNeeded() {
    const s = this.scroller;
    if (s.scrollTop + s.clientHeight > this._sh - 300) { this.extraR += 200; this.geoKey = ''; }
    if (s.scrollLeft + s.clientWidth > this._sw - 300) { this.extraC += 20; this.geoKey = ''; }
  }
  // satır/sütunun ekran konumu (wrap içi)
  xOf(c) { const hw = this.hw(); return c < this.fc ? hw + this.colX[c] : hw + this.colX[this.fc] + this.colX[c] - this.colX[this.fc] - this.scroller.scrollLeft; }
  yOf(r) { const hh = this.hh(); return r < this.fr ? hh + this.rowY[r] : hh + this.rowY[this.fr] + this.rowY[r] - this.rowY[this.fr] - this.scroller.scrollTop; }
  colAt(x) { const hw = this.hw(); const fw = this.colX[this.fc]; if (x < hw + fw) return bsearch(this.colX, x - hw, this.fc + 1); return Math.min(this.nC - 1, bsearch(this.colX, x - hw - fw + this.scroller.scrollLeft + fw, this.nC)); }
  rowAt(y) { const hh = this.hh(); const fh = this.rowY[this.fr]; if (y < hh + fh) return bsearch(this.rowY, y - hh, this.fr + 1); return Math.min(this.nR - 1, bsearch(this.rowY, y - hh - fh + this.scroller.scrollTop + fh, this.nR)); }
  cellRect(r, c) {
    this.geo(); const m = this.wb.mergeAt(this.si, r, c);
    const r1 = m ? m.r1 : r, c1 = m ? m.c1 : c, r2 = m ? m.r2 : r, c2 = m ? m.c2 : c;
    const x = this.xOf(c1), y = this.yOf(r1);
    return { x, y, w: this.colX[c2 + 1] - this.colX[c1], h: this.rowY[r2 + 1] - this.rowY[r1] };
  }
  hit(x, y) {
    this.geo(); const hw = this.hw(), hh = this.hh();
    if (x < hw && y < hh) return { t: 'corner' };
    if (y < hh) { const c = this.colAt(x); const xr = this.xOf(c) + this.cw(c); const xl = this.xOf(c); if (Math.abs(x - xr) < 7) return { t: 'colB', c }; if (Math.abs(x - xl) < 5 && c > 0) return { t: 'colB', c: c - 1 }; return { t: 'col', c }; }
    if (x < hw) { const r = this.rowAt(y); const yb = this.yOf(r) + this.rh(r); if (Math.abs(y - yb) < 6) return { t: 'rowB', r }; return { t: 'row', r }; }
    const r = this.rowAt(y), c = this.colAt(x);
    const s = this.app.sel;
    if (s && !this.editing) {
      const er = this.cellRect(s.r2, s.c2); const hx = this.xOf(s.c2) + this.cw(s.c2), hy = this.yOf(s.r2) + this.rh(s.r2);
      const tol = isTouchDev() ? 16 : 6;
      if (Math.abs(x - hx) < tol && Math.abs(y - hy) < tol) return { t: 'handle', r, c };
    }
    // filtre düğmesi
    const f = this.sh.filter; const cr = this.cellRect(r, c);
    if (f && r === f.r1 && c >= f.c1 && c <= f.c2 && x > cr.x + cr.w - 22 * Math.max(1, this.z)) return { t: 'filter', r, c };
    const dv = this.wb.dvAt(this.si, r, c);
    if (dv && dv.type === 'list' && s && s.r === r && s.c === c && x > cr.x + cr.w - 24) return { t: 'dv', r, c };
    return { t: 'cell', r, c };
  }
  scrollTo(r, c) {
    this.geo(); const s = this.scroller; const hw = this.hw(), hh = this.hh();
    const fw = this.colX[this.fc], fh = this.rowY[this.fr];
    if (c >= this.fc) { const x = this.colX[c] - fw, w = this.cw(c); const vw = s.clientWidth - hw - fw; if (x < s.scrollLeft) s.scrollLeft = x; else if (x + w > s.scrollLeft + vw) s.scrollLeft = x + w - vw + 2; }
    if (r >= this.fr) { const y = this.rowY[r] - fh, h = this.rh(r); const vh = s.clientHeight - hh - fh; if (y < s.scrollTop) s.scrollTop = y; else if (y + h > s.scrollTop + vh) s.scrollTop = y + h - vh + 2; }
  }

  // ---------- çizim ----------
  render() { if (this._raf) return; this._raf = requestAnimationFrame(() => { this._raf = 0; this.draw(); }); }
  draw() {
    if (!this.app.wb) return;
    this.geo();
    const ctx = this.ctx, W = this.W, H = this.H, col = this.col;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = col.bg; ctx.fillRect(0, 0, W, H);
    const hw = this.hw(), hh = this.hh(); const fw = this.colX[this.fc], fh = this.rowY[this.fr];
    const sL = this.scroller.scrollLeft, sT = this.scroller.scrollTop;
    const rowsScroll = [bsearch(this.rowY, sT + fh, this.nR), bsearch(this.rowY, sT + fh + H - hh - fh, this.nR)];
    const colsScroll = [bsearch(this.colX, sL + fw, this.nC), bsearch(this.colX, sL + fw + W - hw - fw, this.nC)];
    const regions = [];
    const rowsF = this.fr ? [0, this.fr - 1] : null, colsF = this.fc ? [0, this.fc - 1] : null;
    regions.push({ rows: rowsScroll, cols: colsScroll, clip: [hw + fw, hh + fh, W - hw - fw, H - hh - fh] });
    if (rowsF) regions.push({ rows: rowsF, cols: colsScroll, clip: [hw + fw, hh, W - hw - fw, fh] });
    if (colsF) regions.push({ rows: rowsScroll, cols: colsF, clip: [hw, hh + fh, fw, H - hh - fh] });
    if (rowsF && colsF) regions.push({ rows: rowsF, cols: colsF, clip: [hw, hh, fw, fh] });
    for (const rg of regions) { ctx.save(); ctx.beginPath(); ctx.rect(...rg.clip); ctx.clip(); this.drawRegion(rg); this.drawSel(rg); ctx.restore(); }
    this.drawHeaders(rowsScroll, colsScroll, rowsF, colsF);
    if (this.fr) { ctx.fillStyle = col.dark ? '#52525b' : '#a1a1aa'; ctx.fillRect(0, hh + fh - 1, W, 2); }
    if (this.fc) { ctx.fillStyle = col.dark ? '#52525b' : '#a1a1aa'; ctx.fillRect(hw + fw - 1, 0, 2, H); }
  }
  font(st) { const z = this.z; const fs = ((st && st.fs) || 11) * 1.25 * z; return `${st && st.i ? 'italic ' : ''}${st && st.b ? '700 ' : ''}${fs.toFixed(1)}px -apple-system, "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif`; }
  drawRegion(rg) {
    const ctx = this.ctx, col = this.col, sh = this.sh, wb = this.wb, si = this.si;
    const [r0, r1] = rg.rows, [c0, c1] = rg.cols;
    // birleşik hücreler
    const mskip = new Set(), mlist = [];
    for (const m of sh.merges) { if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue; mlist.push(m); for (let r = Math.max(m.r1, r0); r <= Math.min(m.r2, r1); r++) for (let c = Math.max(m.c1, c0); c <= Math.min(m.c2, c1); c++) mskip.add(r * 20000 + c); }
    const xs = [], ws = []; for (let c = c0; c <= c1; c++) { xs[c - c0] = this.xOf(c); ws[c - c0] = this.cw(c); }
    const cache = new Map();
    const info = (r, c) => { const k = r * 20000 + c; let o = cache.get(k); if (o) return o; const cell = sh.cells.get(M.key(r, c)); const cf = (cell || sh.cf.length) ? wb.cfStyle(si, r, c) : null; o = { cell, st: cell && cell.s, cf }; cache.set(k, o); return o; };
    // 1) dolgular
    for (let r = r0; r <= r1; r++) {
      const h = this.rh(r); if (!h) continue; const y = this.yOf(r);
      for (let c = c0; c <= c1; c++) {
        const w = ws[c - c0]; if (!w) continue; const o = info(r, c); const bg = (o.cf && o.cf.bg) || (o.st && o.st.bg);
        if (bg && !mskip.has(r * 20000 + c)) { ctx.fillStyle = bg; ctx.fillRect(xs[c - c0], y, w, h); }
      }
    }
    // 2) kılavuz çizgileri
    if (sh.grid !== false) {
      ctx.strokeStyle = col.grid; ctx.lineWidth = 1; ctx.beginPath();
      const xa = xs[0], xb = xs[c1 - c0] + ws[c1 - c0];
      for (let r = r0; r <= r1; r++) { const h = this.rh(r); if (!h) continue; const y = Math.round(this.yOf(r) + h) - 0.5; ctx.moveTo(xa, y); ctx.lineTo(xb, y); }
      const ya = this.yOf(r0), yb = this.yOf(r1) + this.rh(r1);
      for (let c = c0; c <= c1; c++) { const w = ws[c - c0]; if (!w) continue; const x = Math.round(xs[c - c0] + w) - 0.5; ctx.moveTo(x, ya); ctx.lineTo(x, yb); }
      ctx.stroke();
    }
    // 3) birleşik alanlar
    for (const m of mlist) {
      const x = this.xOf(m.c1), y = this.yOf(m.r1), w = this.colX[m.c2 + 1] - this.colX[m.c1], h = this.rowY[m.r2 + 1] - this.rowY[m.r1];
      const o = info(m.r1, m.c1); ctx.fillStyle = (o.cf && o.cf.bg) || (o.st && o.st.bg) || col.bg; ctx.fillRect(x, y, w - 1, h - 1);
      if (sh.grid !== false) { ctx.strokeStyle = col.grid; ctx.strokeRect(Math.round(x) - 0.5, Math.round(y) - 0.5, Math.round(w), Math.round(h)); }
    }
    // 4) kenarlıklar
    for (let r = r0; r <= r1; r++) {
      const h = this.rh(r); if (!h) continue; const y = this.yOf(r);
      for (let c = c0; c <= c1; c++) {
        const o = info(r, c); if (!o.st || !o.st.bd) continue; const w = ws[c - c0]; if (!w) continue;
        let x = xs[c - c0], ww = w, yy = y, hh2 = h; const m = mskip.has(r * 20000 + c) ? wb.mergeAt(si, r, c) : null;
        if (m) { if (m.r1 !== r || m.c1 !== c) continue; ww = this.colX[m.c2 + 1] - this.colX[m.c1]; hh2 = this.rowY[m.r2 + 1] - this.rowY[m.r1]; }
        this.drawBorders(o.st.bd, x, yy, ww, hh2);
      }
    }
    // 5) metinler
    ctx.textBaseline = 'middle';
    const z = this.z, pad = 4 * z;
    for (let r = r0; r <= r1; r++) {
      const h0 = this.rh(r); if (!h0) continue; const y0 = this.yOf(r);
      for (let c = c0; c <= c1; c++) {
        const w0 = ws[c - c0]; if (!w0) continue;
        let x = xs[c - c0], y = y0, w = w0, h = h0;
        const mk = r * 20000 + c;
        if (mskip.has(mk)) {
          const m = wb.mergeAt(si, r, c);
          if (r !== Math.max(m.r1, r0) || c !== Math.max(m.c1, c0)) continue; // birleşik alan bir kez çizilir
          const o2 = info(m.r1, m.c1); if (!o2.cell) continue;
          this.drawText(o2, m.r1, m.c1, this.xOf(m.c1), this.yOf(m.r1), this.colX[m.c2 + 1] - this.colX[m.c1], this.rowY[m.r2 + 1] - this.rowY[m.r1], pad, null);
          continue;
        }
        const o = info(r, c);
        if (o.cf && o.cf.bar != null) { const bw = Math.max(0, (w - 4) * o.cf.bar); const gr = ctx.createLinearGradient(x, 0, x + bw, 0); gr.addColorStop(0, o.cf.barColor); gr.addColorStop(1, o.cf.barColor + '44'); ctx.fillStyle = gr; ctx.fillRect(x + 2, y + 3, bw, h - 6); }
        if (!o.cell) {
          const sp = wb.spillAt(si, r, c); if (!sp) continue;
          const src = sh.cells.get(sp.src); this.drawText({ st: src && src.s && src.s.nf ? { nf: src.s.nf } : null, cf: o.cf }, r, c, x, y, w, h, pad, null);
          continue;
        }
        // taşma (sola hizalı metin)
        let clipW = null;
        this.drawText(o, r, c, x, y, w, h, pad, (needW) => {
          let ew = w, cc = c + 1;
          while (ew < needW && cc < this.nC && cc - c < 30) { if (wb.value(si, r, cc) != null) break; if (mskip.has(r * 20000 + cc)) break; ew += this.cw(cc); cc++; }
          return ew;
        });
      }
    }
    // filtre düğmeleri
    const f = sh.filter;
    if (f && f.r1 >= r0 && f.r1 <= r1) {
      for (let c = Math.max(f.c1, c0); c <= Math.min(f.c2, c1); c++) {
        const x = xs[c - c0] + ws[c - c0], y = this.yOf(f.r1), h = this.rh(f.r1); const s = Math.min(20 * z, h - 4);
        const active = f.crit && f.crit[c]; ctx.fillStyle = active ? col.accent : (col.dark ? '#3f3f46' : '#f4f4f5'); ctx.strokeStyle = col.dark ? '#52525b' : '#d4d4d8';
        const bx = x - s - 3, by = y + (h - s) / 2; G.XC && rr(ctx, bx, by, s, s, 4); ctx.fill(); ctx.stroke();
        ctx.strokeStyle = active ? '#fff' : col.headFg; ctx.lineWidth = 1.6; ctx.beginPath(); ctx.moveTo(bx + s * 0.3, by + s * 0.42); ctx.lineTo(bx + s * 0.5, by + s * 0.62); ctx.lineTo(bx + s * 0.7, by + s * 0.42); ctx.stroke(); ctx.lineWidth = 1;
      }
    }
  }
  drawBorders(bd, x, y, w, h) {
    const ctx = this.ctx;
    const line = (b, x1, y1, x2, y2) => { if (!b) return; const lw = b.s === 'thick' ? 3 : b.s === 'medium' ? 2 : 1; ctx.strokeStyle = (b.c && !(this.col.dark && b.c === '#000000')) ? b.c : this.col.fg; ctx.lineWidth = lw; ctx.setLineDash(b.s === 'dashed' ? [4, 3] : []); ctx.beginPath(); const o = lw % 2 ? 0.5 : 0; ctx.moveTo(Math.round(x1) - o, Math.round(y1) - o); ctx.lineTo(Math.round(x2) - o, Math.round(y2) - o); ctx.stroke(); };
    line(bd.t, x, y, x + w, y); line(bd.b, x, y + h, x + w, y + h); line(bd.l, x, y, x, y + h); line(bd.r, x + w, y, x + w, y + h);
    ctx.setLineDash([]); ctx.lineWidth = 1;
  }
  drawText(o, r, c, x, y, w, h, pad, overflow) {
    const ctx = this.ctx, wb = this.wb, si = this.si, col = this.col;
    const v = wb.value(si, r, c); if (v == null || v === '') return;
    const st = o.st || {}; const cf = o.cf || {};
    const nf = st.nf; let txt = X.formatValue(v, nf);
    if (txt === '') return;
    const sty = cf.b ? Object.assign({}, st, { b: 1 }) : st;
    ctx.font = this.font(sty);
    let color = cf.fc || X.formatColor(v, nf) || st.fc || (X.isErr(v) ? col.err : null);
    if (!color || (col.dark && !st.bg && !cf.bg && /^#(000000|000|0d0d0d|111111)$/i.test(color))) color = col.fg;
    if (col.dark && (st.bg || cf.bg) && !st.fc && !cf.fc) color = lum(st.bg || cf.bg) > 0.55 ? '#111' : '#f4f4f5';
    ctx.fillStyle = color;
    const isNum = typeof v === 'number';
    let ha = st.ha || (isNum ? 'right' : (typeof v === 'boolean' || X.isErr(v)) ? 'center' : 'left');
    let ix = x;
    if (cf.icon) { const s = Math.min(12 * this.z, h - 6); this.drawIcon(cf.icon, x + pad, y + h / 2, s); ix += s + pad; w -= s + pad; }
    if (st.wrap) {
      const lines = this.wrapLines(txt, w - pad * 2); const lh = ((st.fs || 11) * 1.25 * this.z) * 1.25;
      ctx.save(); ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      let ty = st.va === 'middle' ? y + (h - lines.length * lh) / 2 + lh / 2 : st.va === 'bottom' ? y + h - lines.length * lh + lh / 2 - 2 : y + lh / 2 + 2;
      ctx.textAlign = ha;
      for (const ln of lines) { const tx = ha === 'left' ? ix + pad : ha === 'center' ? ix + w / 2 : ix + w - pad; ctx.fillText(ln, tx, ty); ty += lh; }
      ctx.restore(); return;
    }
    const tw = ctx.measureText(txt).width;
    let clipW = w;
    if (tw > w - pad * 2) {
      if (isNum && !X.isDateFormat(nf)) { const g = X.fmtGeneral(+v.toPrecision(Math.max(1, Math.floor((w - pad * 2) / (ctx.measureText('0').width)) - 1))); txt = ctx.measureText(g).width <= w - pad * 2 && !nf ? g : '#'.repeat(Math.max(1, Math.floor((w - pad * 2) / ctx.measureText('#').width))); }
      else if (ha === 'left' && overflow && typeof v === 'string') clipW = overflow(tw + pad * 2);
    }
    const ty = st.va === 'top' ? y + h * 0.3 : st.va === 'bottom' ? y + h - h * 0.3 : y + h / 2 + 0.5;
    const needClip = clipW < tw + pad * 2 || ha !== 'left';
    if (needClip) { ctx.save(); ctx.beginPath(); ctx.rect(x, y, clipW, h); ctx.clip(); }
    ctx.textAlign = ha;
    const tx = ha === 'left' ? ix + pad : ha === 'center' ? ix + w / 2 : ix + w - pad;
    ctx.fillText(txt, tx, ty);
    if (st.u || st.st) {
      const m = ctx.measureText(txt).width; const lx = ha === 'left' ? tx : ha === 'center' ? tx - m / 2 : tx - m; const fs = (st.fs || 11) * 1.25 * this.z;
      ctx.fillRect(lx, st.u ? ty + fs * 0.45 : ty, m, Math.max(1, this.z));
      if (st.u && st.st) ctx.fillRect(lx, ty, m, Math.max(1, this.z));
    }
    if (needClip) ctx.restore();
  }
  drawIcon(kind, x, cy, s) {
    const ctx = this.ctx; ctx.save();
    ctx.fillStyle = kind === 'up' ? '#16a34a' : kind === 'down' ? '#dc2626' : '#f59e0b';
    ctx.beginPath();
    if (kind === 'up') { ctx.moveTo(x, cy + s / 2); ctx.lineTo(x + s / 2, cy - s / 2); ctx.lineTo(x + s, cy + s / 2); }
    else if (kind === 'down') { ctx.moveTo(x, cy - s / 2); ctx.lineTo(x + s / 2, cy + s / 2); ctx.lineTo(x + s, cy - s / 2); }
    else { ctx.moveTo(x, cy - s / 2); ctx.lineTo(x + s, cy); ctx.lineTo(x, cy + s / 2); }
    ctx.closePath(); ctx.fill(); ctx.restore();
  }
  wrapLines(txt, maxW) {
    const ctx = this.ctx; const out = [];
    for (const para of String(txt).split('\n')) {
      const words = para.split(' '); let line = '';
      for (const wd of words) { const t = line ? line + ' ' + wd : wd; if (ctx.measureText(t).width > maxW && line) { out.push(line); line = wd; } else line = t; }
      out.push(line);
    }
    return out;
  }
  rectOf(g) { const x = this.xOf(g.c1), y = this.yOf(g.r1); return { x, y, w: this.colX[g.c2 + 1] - this.colX[g.c1], h: this.rowY[g.r2 + 1] - this.rowY[g.r1] }; }
  drawSel(rg) {
    const ctx = this.ctx, col = this.col, app = this.app;
    // formül başvuru vurguları
    if (this.refMarks) for (const m of this.refMarks) { if (m.si !== this.si) continue; const q = this.rectOf(m); ctx.fillStyle = m.color + '22'; ctx.fillRect(q.x, q.y, q.w, q.h); ctx.strokeStyle = m.color; ctx.lineWidth = 2; ctx.strokeRect(q.x + 1, q.y + 1, q.w - 2, q.h - 2); }
    // kopya çerçevesi
    if (app.clip && app.clip.si === this.si) { const q = this.rectOf(app.clip); ctx.setLineDash([5, 4]); ctx.strokeStyle = col.accent; ctx.lineWidth = 2; ctx.strokeRect(q.x + 1, q.y + 1, q.w - 2, q.h - 2); ctx.setLineDash([]); }
    const s = app.sel; if (!s) return;
    const q = this.rectOf(s);
    const multi = s.r1 !== s.r2 || s.c1 !== s.c2;
    if (multi) { ctx.fillStyle = col.selFill; ctx.fillRect(q.x, q.y, q.w, q.h); const a = this.cellRect(s.r, s.c); ctx.fillStyle = col.bg; ctx.globalAlpha = 0.0; ctx.fillRect(a.x, a.y, a.w, a.h); ctx.globalAlpha = 1; }
    const pick = app.pick && app.pick.active;
    ctx.strokeStyle = pick ? '#7c3aed' : col.accent; ctx.lineWidth = 2; if (pick) ctx.setLineDash([6, 4]);
    ctx.strokeRect(Math.round(q.x) + 0.5, Math.round(q.y) + 0.5, Math.round(q.w) - 1, Math.round(q.h) - 1); ctx.setLineDash([]);
    if (multi) { const a = this.cellRect(s.r, s.c); ctx.strokeStyle = col.accent; ctx.lineWidth = 1; ctx.strokeRect(a.x + 2.5, a.y + 2.5, a.w - 5, a.h - 5); }
    // doldurma tutamacı
    if (!this.editing && !pick) {
      const hx = q.x + q.w, hy = q.y + q.h; const t = isTouchDev();
      ctx.fillStyle = col.accent; ctx.strokeStyle = col.bg; ctx.lineWidth = 2;
      ctx.beginPath(); if (t) ctx.arc(hx, hy, 7, 0, 7); else ctx.rect(hx - 4, hy - 4, 8, 8); ctx.fill(); ctx.stroke(); ctx.lineWidth = 1;
    }
    // açılır liste oku
    const dv = this.wb.dvAt(this.si, s.r, s.c);
    if (dv && dv.type === 'list' && !this.editing) { const a = this.cellRect(s.r, s.c); const bx = a.x + a.w - 22, by = a.y + 3, bh = a.h - 6; ctx.fillStyle = col.dark ? '#3f3f46' : '#f4f4f5'; rr(ctx, bx, by, 19, bh, 4); ctx.fill(); ctx.strokeStyle = col.headFg; ctx.lineWidth = 1.6; ctx.beginPath(); ctx.moveTo(bx + 5, by + bh / 2 - 2); ctx.lineTo(bx + 9.5, by + bh / 2 + 2.5); ctx.lineTo(bx + 14, by + bh / 2 - 2); ctx.stroke(); ctx.lineWidth = 1; }
  }
  drawHeaders(rowsS, colsS, rowsF, colsF) {
    const ctx = this.ctx, col = this.col, W = this.W, H = this.H; const hw = this.hw(), hh = this.hh(); const s = this.app.sel;
    ctx.fillStyle = col.head; ctx.fillRect(0, 0, W, hh); ctx.fillRect(0, 0, hw, H);
    ctx.font = `500 ${(11.5 * Math.min(1.2, Math.max(0.85, this.z))).toFixed(1)}px -apple-system, "Segoe UI", system-ui, sans-serif`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const fw = this.colX[this.fc], fh = this.rowY[this.fr];
    const doCols = (c0, c1, clipX, clipW) => {
      ctx.save(); ctx.beginPath(); ctx.rect(clipX, 0, clipW, hh); ctx.clip();
      for (let c = c0; c <= c1; c++) { const w = this.cw(c); if (!w) continue; const x = this.xOf(c); const on = s && c >= s.c1 && c <= s.c2;
        if (on) { ctx.fillStyle = col.headSel; ctx.fillRect(x, 0, w, hh); ctx.fillStyle = col.accent; ctx.fillRect(x, hh - 2, w, 2); }
        ctx.fillStyle = on ? col.accent : col.headFg; ctx.fillText(X.colName(c), x + w / 2, hh / 2 + 0.5);
        ctx.fillStyle = col.grid; ctx.fillRect(Math.round(x + w) - 1, 4, 1, hh - 8); }
      ctx.restore();
    };
    const doRows = (r0, r1, clipY, clipH) => {
      ctx.save(); ctx.beginPath(); ctx.rect(0, clipY, hw, clipH); ctx.clip();
      for (let r = r0; r <= r1; r++) { const h = this.rh(r); if (!h) continue; const y = this.yOf(r); const on = s && r >= s.r1 && r <= s.r2;
        if (on) { ctx.fillStyle = col.headSel; ctx.fillRect(0, y, hw, h); ctx.fillStyle = col.accent; ctx.fillRect(hw - 2, y, 2, h); }
        ctx.fillStyle = on ? col.accent : col.headFg; if (h > 9) ctx.fillText(String(r + 1), hw / 2, y + h / 2 + 0.5);
        ctx.fillStyle = col.grid; ctx.fillRect(4, Math.round(y + h) - 1, hw - 8, 1); }
      ctx.restore();
    };
    doCols(colsS[0], colsS[1], hw + fw, W - hw - fw); if (colsF) doCols(colsF[0], colsF[1], hw, fw);
    doRows(rowsS[0], rowsS[1], hh + fh, H - hh - fh); if (rowsF) doRows(rowsF[0], rowsF[1], hh, fh);
    ctx.fillStyle = col.head; ctx.fillRect(0, 0, hw, hh);
    ctx.fillStyle = col.headFg; ctx.beginPath(); ctx.moveTo(hw - 6, hh - 6); ctx.lineTo(hw - 6, hh - 14); ctx.lineTo(hw - 14, hh - 6); ctx.closePath(); ctx.globalAlpha = 0.5; ctx.fill(); ctx.globalAlpha = 1;
    ctx.fillStyle = col.grid; ctx.fillRect(0, hh - 1, W, 1); ctx.fillRect(hw - 1, 0, 1, H);
  }

  // ---------- seçim ----------
  local(e) { const r = this.wrap.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; }
  selectAt(h, extend) {
    const app = this.app; const s = app.sel;
    if (h.t === 'corner') { app.setSel({ r: 0, c: 0, r1: 0, c1: 0, r2: this.nR - 1, c2: this.nC - 1 }); return; }
    if (h.t === 'col') { const a = extend ? s.c : h.c; app.setSel({ r: extend ? s.r : 0, c: a, r1: 0, r2: this.nR - 1, c1: Math.min(a, h.c), c2: Math.max(a, h.c) }); return; }
    if (h.t === 'row') { const a = extend ? s.r : h.r; app.setSel({ r: a, c: extend ? s.c : 0, c1: 0, c2: this.nC - 1, r1: Math.min(a, h.r), r2: Math.max(a, h.r) }); return; }
    if (extend) app.setSel(this.expandMerges({ r: s.r, c: s.c, r1: Math.min(s.r, h.r), c1: Math.min(s.c, h.c), r2: Math.max(s.r, h.r), c2: Math.max(s.c, h.c) }));
    else { const m = this.wb.mergeAt(this.si, h.r, h.c); app.setSel(m ? { r: m.r1, c: m.c1, r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r: h.r, c: h.c, r1: h.r, c1: h.c, r2: h.r, c2: h.c }); }
  }
  expandMerges(s) {
    let ch = true; const ms = this.sh.merges; let k = 0;
    while (ch && k++ < 10) { ch = false; for (const m of ms) { if (m.r2 < s.r1 || m.r1 > s.r2 || m.c2 < s.c1 || m.c1 > s.c2) continue; if (m.r1 < s.r1) { s.r1 = m.r1; ch = true; } if (m.c1 < s.c1) { s.c1 = m.c1; ch = true; } if (m.r2 > s.r2) { s.r2 = m.r2; ch = true; } if (m.c2 > s.c2) { s.c2 = m.c2; ch = true; } } }
    return s;
  }
  // formül düzenlerken hücreye dokunma -> başvuru ekle
  canInsertRef() {
    if (!this.editing) return false;
    const el = this.activeInput(); const t = el.value; if (!t.startsWith('=')) return false;
    const p = el.selectionStart; const before = t.slice(0, p).replace(/\s+$/, '');
    if (this._refIns && this._refIns.end === p) return true;
    return /[=(;,+\-*/^&<>:]$/.test(before);
  }
  insertRef(g, fresh) {
    const el = this.activeInput(); let t = el.value;
    let txt = (g.r1 === g.r2 && g.c1 === g.c2) ? X.addr(g.r1, g.c1) : X.addr(g.r1, g.c1) + ':' + X.addr(g.r2, g.c2);
    if (this.editing.si !== this.si) txt = X.quoteSheet(this.sh.name) + '!' + txt;
    let a, b;
    if (!fresh && this._refIns) { a = this._refIns.start; b = this._refIns.end; } else { a = el.selectionStart; b = el.selectionEnd; }
    t = t.slice(0, a) + txt + t.slice(b); el.value = t; const e = a + txt.length;
    el.setSelectionRange(e, e); this._refIns = { start: a, end: e };
    this.syncInputs(el); this.updateRefMarks();
  }

  bindMouse() {
    const sc = this.scroller; let drag = null;
    sc.addEventListener('mousedown', e => {
      if (e.button !== 0 || e.target.closest('.chart')) return;
      const p = this.local(e); const h = this.hit(p.x, p.y); this.app.closePopups && this.app.closePopups();
      if (h.t === 'colB' || h.t === 'rowB') { drag = { t: h.t, i: h.t === 'colB' ? h.c : h.r, start: h.t === 'colB' ? p.x : p.y, orig: h.t === 'colB' ? (this.sh.colW[h.c] || DEF_W) : (this.sh.rowH[h.r] || DEF_H) }; e.preventDefault(); return; }
      if (h.t === 'filter') { e.preventDefault(); this.app.openFilter(h.c); return; }
      if (h.t === 'dv') { e.preventDefault(); this.app.openDvList(h.r, h.c); return; }
      if (this.app.pick && this.app.pick.active && h.t === 'cell') { e.preventDefault(); this.selectAt(h, e.shiftKey); drag = { t: 'sel' }; return; }
      if (this.editing && h.t === 'cell' && this.canInsertRef()) { e.preventDefault(); this._refAnchor = h; this.insertRef({ r1: h.r, c1: h.c, r2: h.r, c2: h.c }, !this._refIns || this._refIns.end !== this.activeInput().selectionStart); drag = { t: 'ref' }; return; }
      if (this.editing) this.commitEdit();
      if (h.t === 'handle') { e.preventDefault(); drag = { t: 'fill', src: { ...this.app.sel } }; return; }
      this.selectAt(h, e.shiftKey); drag = { t: h.t === 'cell' ? 'sel' : h.t }; e.preventDefault();
      this.app.focusGrid();
    });
    addEventListener('mousemove', e => {
      const p = this.local(e);
      if (!drag) { if (e.target === sc || sc.contains(e.target) && !e.target.closest('.chart')) { const h = this.hit(p.x, p.y); sc.style.cursor = h.t === 'colB' ? 'col-resize' : h.t === 'rowB' ? 'row-resize' : h.t === 'handle' ? 'crosshair' : h.t === 'filter' || h.t === 'dv' ? 'pointer' : 'cell'; } return; }
      this.dragMove(drag, p);
    });
    addEventListener('mouseup', e => { if (drag) { this.dragEnd(drag); drag = null; } });
    sc.addEventListener('dblclick', e => {
      if (e.target.closest('.chart')) return;
      const p = this.local(e); const h = this.hit(p.x, p.y);
      if (h.t === 'colB') { this.app.autoFit([h.c]); return; }
      if (h.t === 'cell') this.startEdit(null, 'edit');
    });
    sc.addEventListener('contextmenu', e => {
      if (e.target.closest('.chart')) return;
      e.preventDefault(); const p = this.local(e); const h = this.hit(p.x, p.y);
      if (h.t === 'cell') { const s = this.app.sel; if (!(h.r >= s.r1 && h.r <= s.r2 && h.c >= s.c1 && h.c <= s.c2)) this.selectAt(h); }
      this.app.contextMenu(e.clientX, e.clientY, h);
    });
    sc.addEventListener('wheel', e => { if (e.ctrlKey) { e.preventDefault(); this.app.setZoom(this.z * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); } }, { passive: false });
  }
  dragMove(d, p) {
    if (d.t === 'colB') { const nw = Math.max(16, d.orig + (p.x - d.start) / this.z); this.sh.colW[d.i] = Math.round(nw); d.changed = true; this.geoKey = ''; this.render(); this.app.positionCharts(); return; }
    if (d.t === 'rowB') { const nh = Math.max(10, d.orig + (p.y - d.start) / this.z); this.sh.rowH[d.i] = Math.round(nh); d.changed = true; this.geoKey = ''; this.render(); return; }
    this.autoScroll(p);
    const x = Math.max(this.hw() + 1, Math.min(this.W - 1, p.x)), y = Math.max(this.hh() + 1, Math.min(this.H - 1, p.y));
    const r = this.rowAt(y), c = this.colAt(x);
    if (d.t === 'sel') this.selectAt({ t: 'cell', r, c }, true);
    else if (d.t === 'col') this.selectAt({ t: 'col', c }, true);
    else if (d.t === 'row') this.selectAt({ t: 'row', r }, true);
    else if (d.t === 'ref') { const a = this._refAnchor; this.insertRef({ r1: Math.min(a.r, r), c1: Math.min(a.c, c), r2: Math.max(a.r, r), c2: Math.max(a.c, c) }, false); this.render(); }
    else if (d.t === 'fill') {
      const s = d.src; let g;
      const dr = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0, dc = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0;
      if (Math.abs(dr) >= Math.abs(dc)) g = { r1: Math.min(s.r1, r), r2: Math.max(s.r2, r), c1: s.c1, c2: s.c2 }; else g = { r1: s.r1, r2: s.r2, c1: Math.min(s.c1, c), c2: Math.max(s.c2, c) };
      d.dst = g; this.app.setSel(Object.assign({ r: s.r, c: s.c }, g));
    }
  }
  dragEnd(d) {
    if ((d.t === 'colB' || d.t === 'rowB') && d.changed) { this.app.markGeo(); }
    if (d.t === 'fill' && d.dst) this.app.autofill(d.src, d.dst);
    if (this._as) { clearInterval(this._as); this._as = null; }
  }
  autoScroll(p) {
    const s = this.scroller; let dx = 0, dy = 0;
    if (p.x > this.W - 20) dx = 20; else if (p.x < this.hw() + 4 && s.scrollLeft > 0) dx = -20;
    if (p.y > this.H - 20) dy = 20; else if (p.y < this.hh() + 4 && s.scrollTop > 0) dy = -20;
    if (dx || dy) { s.scrollLeft += dx; s.scrollTop += dy; }
  }

  bindTouch() {
    const sc = this.scroller; let T = null; let lastTap = { t: 0, r: -1, c: -1 };
    const pt = t => { const r = this.wrap.getBoundingClientRect(); return { x: t.clientX - r.left, y: t.clientY - r.top }; };
    sc.addEventListener('touchstart', e => {
      if (e.target.closest('.chart')) return;
      if (e.touches.length === 2) { if (T && T.timer) clearTimeout(T.timer); const [a, b] = e.touches; T = { mode: 'pinch', d0: Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY), z0: this.z }; e.preventDefault(); return; }
      if (e.touches.length > 2) return;
      const p = pt(e.touches[0]); const h = this.hit(p.x, p.y);
      T = { p0: p, h, t0: Date.now(), mode: null, moved: false };
      this.app.closePopups && this.app.closePopups();
      const pickMode = this.app.pick && this.app.pick.active;
      if (h.t === 'colB' || h.t === 'rowB') { T.mode = 'drag'; T.d = { t: h.t, i: h.t === 'colB' ? h.c : h.r, start: h.t === 'colB' ? p.x : p.y, orig: h.t === 'colB' ? (this.sh.colW[h.c] || DEF_W) : (this.sh.rowH[h.r] || DEF_H) }; e.preventDefault(); return; }
      if (h.t === 'handle') { T.mode = 'drag'; T.d = { t: 'fill', src: { ...this.app.sel } }; e.preventDefault(); return; }
      if (h.t === 'cell' && (pickMode || this.app.selectMode || (this.editing && this.canInsertRef()))) {
        e.preventDefault();
        if (this.editing && !pickMode && this.canInsertRef()) { this._refAnchor = h; this.insertRef({ r1: h.r, c1: h.c, r2: h.r, c2: h.c }, !this._refIns || this._refIns.end !== this.activeInput().selectionStart); T.mode = 'drag'; T.d = { t: 'ref' }; return; }
        if (this.editing) this.commitEdit();
        this.selectAt(h); T.mode = 'drag'; T.d = { t: 'sel' }; return;
      }
      if (h.t === 'col' || h.t === 'row') return; // dokunma ile başlıktan seçim (bırakınca)
      T.timer = setTimeout(() => { if (!T || T.moved || T.mode) return; T.mode = 'long'; if (h.t === 'cell') { const s = this.app.sel; if (!(h.r >= s.r1 && h.r <= s.r2 && h.c >= s.c1 && h.c <= s.c2)) { if (this.editing) this.commitEdit(); this.selectAt(h); } } navigator.vibrate && navigator.vibrate(15); }, 420);
    }, { passive: false });
    sc.addEventListener('touchmove', e => {
      if (!T) return;
      if (T.mode === 'pinch' && e.touches.length === 2) { e.preventDefault(); const [a, b] = e.touches; const d = Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY); this.app.setZoom(T.z0 * d / T.d0); return; }
      const p = pt(e.touches[0]);
      if (!T.moved && Math.hypot(p.x - T.p0.x, p.y - T.p0.y) > 8) { T.moved = true; if (!T.mode) { clearTimeout(T.timer); T.mode = 'scroll'; } else if (T.mode === 'long') { T.mode = 'drag'; T.d = { t: T.h.t === 'cell' ? 'sel' : 'sel' }; } }
      if (T.mode === 'drag' || T.mode === 'long') { e.preventDefault(); if (T.d) this.dragMove(T.d, p); }
    }, { passive: false });
    sc.addEventListener('touchend', e => {
      if (!T) return; clearTimeout(T.timer);
      const t = T; T = null;
      if (t.mode === 'pinch') { this.app.markGeo(true); return; }
      if (t.mode === 'drag') { e.preventDefault(); if (t.d) this.dragEnd(t.d); return; }
      if (t.mode === 'long') { e.preventDefault(); const r = sc.getBoundingClientRect(); this.app.contextMenu(t.p0.x + r.left, t.p0.y + r.top, t.h); return; }
      if (t.mode === 'scroll') return;
      // dokunma
      e.preventDefault();
      const h = t.h;
      if (h.t === 'filter') { this.app.openFilter(h.c); return; }
      if (h.t === 'dv') { this.app.openDvList(h.r, h.c); return; }
      if (h.t !== 'cell') { if (this.editing) this.commitEdit(); this.selectAt(h); return; }
      const s = this.app.sel; const now = Date.now();
      const same = s && s.r === h.r && s.c === h.c && s.r1 === s.r2 && s.c1 === s.c2 || (s && this.wb.mergeAt(this.si, h.r, h.c) && s.r === (this.wb.mergeAt(this.si, h.r, h.c).r1) && s.c === this.wb.mergeAt(this.si, h.r, h.c).c1);
      if (this.editing) { if (this.editing.r === h.r && this.editing.c === h.c) return; this.commitEdit(); this.selectAt(h); return; }
      if (same && now - t.t0 < 400) {
        const dv = this.wb.dvAt(this.si, h.r, h.c);
        if (dv && dv.type === 'list') this.app.openDvList(h.r, h.c); else this.startEdit(null, 'edit');
        return;
      }
      this.selectAt(h); lastTap = { t: now, r: h.r, c: h.c };
    }, { passive: false });
    sc.addEventListener('touchcancel', () => { if (T) clearTimeout(T.timer); T = null; });
  }

  // ---------- düzenleyici ----------
  activeInput() { return this.editing && this.editing.src === 'fbar' ? this.app.fInput : this.editor; }
  syncInputs(from) { const t = from.value; if (from !== this.editor) this.editor.value = t; if (from !== this.app.fInput) this.app.fInput.value = t; this.fitEditor(); }
  bindEditor() {
    const ed = this.editor;
    ed.addEventListener('input', () => { this._refIns = null; this.syncInputs(ed); this.updateRefMarks(); this.app.onEditInput && this.app.onEditInput(ed.value); });
    ed.addEventListener('keydown', e => this.editorKey(e, ed));
    ed.addEventListener('blur', () => { setTimeout(() => { if (!this.editing || this.editing.src !== 'cell') return; const ae = document.activeElement; if (ae === ed || ae === this.app.fInput) return; if (this.app.keepEdit) return; this.commitEdit(); }, 120); });
  }
  editorKey(e, el) {
    if (e.isComposing) return;
    const k = e.key;
    if (k === 'Enter' && !e.altKey && !e.shiftKey) { e.preventDefault(); this.commitEdit('down'); return; }
    if (k === 'Enter' && e.shiftKey) { e.preventDefault(); this.commitEdit('up'); return; }
    if (k === 'Enter' && e.altKey) { e.preventDefault(); const a = el.selectionStart; el.value = el.value.slice(0, a) + '\n' + el.value.slice(el.selectionEnd); el.setSelectionRange(a + 1, a + 1); this.syncInputs(el); return; }
    if (k === 'Tab') { e.preventDefault(); this.commitEdit(e.shiftKey ? 'left' : 'right'); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancelEdit(); return; }
    if (this.editing && this.editing.mode === 'enter' && /^Arrow/.test(k) && !el.value.startsWith('=')) { e.preventDefault(); this.commitEdit({ ArrowDown: 'down', ArrowUp: 'up', ArrowLeft: 'left', ArrowRight: 'right' }[k]); }
  }
  startEdit(initial, mode, src) {
    const s = this.app.sel; if (!s) return;
    if (this.app.pick && this.app.pick.active) return;
    const r = s.r, c = s.c;
    const cell = this.wb.cell(this.si, r, c);
    this.editing = { r, c, si: this.si, mode: mode || 'enter', src: src || 'cell', orig: this.wb.editText(this.si, r, c) };
    const txt = initial != null ? initial : this.editing.orig;
    this.editor.value = txt; this.app.fInput.value = txt;
    const st = (cell && cell.s) || {};
    this.editor.style.font = this.font(st); this.editor.style.textAlign = 'left'; this.editor.style.color = st.fc && !this.col.dark ? st.fc : '';
    this.editor.style.background = st.bg || '';
    this.wrap.classList.add('editing');
    this.placeEditor();
    if (this.editing.src === 'cell') { this.editor.focus({ preventScroll: true }); const L = this.editor.value.length; this.editor.setSelectionRange(L, L); }
    this.updateRefMarks(); this.render();
    this.app.onEditStart && this.app.onEditStart();
  }
  placeEditor() {
    if (!this.editing) return; if (this.editing.si !== this.si) { this.editor.style.display = 'none'; return; }
    const q = this.cellRect(this.editing.r, this.editing.c);
    this.editor.style.display = 'block'; this.editor.style.left = Math.max(this.hw(), q.x) + 'px'; this.editor.style.top = Math.max(this.hh(), q.y) + 'px';
    this.editor.style.minWidth = q.w + 'px'; this.editor.style.minHeight = q.h + 'px'; this._edMinW = q.w; this.fitEditor();
  }
  fitEditor() {
    if (!this.editing) return; const ctx = this.ctx; ctx.font = this.editor.style.font || this.font();
    const lines = this.editor.value.split('\n'); const w = Math.max(...lines.map(l => ctx.measureText(l).width)) + 24;
    const maxW = this.W - parseFloat(this.editor.style.left || 0) - 8;
    this.editor.style.width = Math.min(maxW, Math.max(this._edMinW || 80, w)) + 'px';
    this.editor.style.height = 'auto'; this.editor.style.height = Math.max(parseFloat(this.editor.style.minHeight) || 26, this.editor.scrollHeight) + 'px';
  }
  updateRefMarks() {
    this.refMarks = null;
    if (!this.editing) return;
    const t = this.activeInput().value; if (!t.startsWith('=')) { this.render(); return; }
    let toks; try { toks = X.tokenize(t.slice(1)); } catch (e) { this.render(); return; }
    const colors = ['#2563eb', '#dc2626', '#7c3aed', '#16a34a', '#ea580c', '#0891b2', '#db2777'];
    const marks = []; let ci = 0;
    for (let i = 0; i < toks.length; i++) {
      const a = toks[i]; if (a.t !== 'ref') continue;
      let g = { r1: a.r, c1: a.c, r2: a.r, c2: a.c };
      if (toks[i + 1] && toks[i + 1].t === ':' && toks[i + 2] && toks[i + 2].t === 'ref') { const b = toks[i + 2]; g = { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }; i += 2; }
      const si = a.sheet ? this.wb.sheetIndex(a.sheet) : this.editing.si;
      marks.push(Object.assign(g, { si, color: colors[ci++ % colors.length] }));
    }
    this.refMarks = marks; this.render();
  }
  commitEdit(move) {
    const ed = this.editing; if (!ed) return true;
    const text = this.activeInput().value;
    if (text !== ed.orig) {
      const ok = this.app.commitInput(ed.si, ed.r, ed.c, text);
      if (!ok) { this.activeInput().focus(); return false; }
    }
    this.endEdit();
    if (this.si !== ed.si) { this.wb.active = ed.si; this.app.refreshSheet(); }
    if (move) this.app.moveSel(move);
    return true;
  }
  cancelEdit() { const ed = this.editing; this.endEdit(); if (ed && ed.si !== this.si) { this.wb.active = ed.si; this.app.refreshSheet(); } }
  endEdit() {
    this.editing = null; this.refMarks = null; this._refIns = null;
    this.editor.style.display = 'none'; this.wrap.classList.remove('editing'); this.editor.blur();
    if (document.activeElement === this.app.fInput) this.app.fInput.blur();
    this.app.onSelChange(); this.render(); this.app.focusGrid();
  }
}
function rr(ctx, x, y, w, h, r) { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); }
function lum(hex) { if (!hex || hex[0] !== '#') return 1; const [r, g, b] = M.hex(hex); return (0.299 * r + 0.587 * g + 0.114 * b) / 255; }
G.XG = { Grid, DEF_W, DEF_H };
})(window);
