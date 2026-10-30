// Excello grafik çizici (harici kütüphane yok)
(function (G) {
'use strict';
const X = G.XF;
const COLORS = ['#2563eb', '#16a34a', '#f59e0b', '#dc2626', '#7c3aed', '#0891b2', '#db2777', '#65a30d', '#ea580c', '#475569'];
const TYPES = [['column', 'Sütun'], ['bar', 'Çubuk'], ['line', 'Çizgi'], ['area', 'Alan'], ['pie', 'Pasta'], ['doughnut', 'Halka'], ['stacked', 'Yığılmış sütun'], ['scatter', 'Dağılım']];

// Alandan seri verisi çıkar
function extract(model, si, g, lr) {
  if (g && g.range) { lr = g.lr; g = g.range; }
  const vals = model.values(si, g.r1, g.c1, g.r2, g.c2);
  const disp = (r, c) => model.display(si, g.r1 + r, g.c1 + c);
  const H = vals.length, W = vals[0] ? vals[0].length : 0;
  if (!H || !W) return { labels: [], series: [] };
  const isTxt = v => typeof v === 'string' && v !== '';
  const headerRow = vals[0].some(isTxt) && (H > 1) && vals[0].slice(W > 1 ? 1 : 0).every(v => v == null || isTxt(v));
  const r0 = headerRow ? 1 : 0;
  let labelCol = false;
  if (W > 1) { let t = 0, n = 0; for (let r = r0; r < H; r++) { if (isTxt(vals[r][0])) t++; else if (typeof vals[r][0] === 'number') n++; } const nf = model.numFmt(si, g.r1 + r0, g.c1); labelCol = t >= n || X.isDateFormat(nf); }
  const c0 = labelCol ? 1 : 0;
  const labels = []; for (let r = r0; r < H; r++) labels.push(labelCol ? disp(r, 0) : lr ? model.display(si, g.r1 + r, lr.c1) : String(r - r0 + 1));
  const series = [];
  for (let c = c0; c < W; c++) {
    const name = headerRow ? String(vals[0][c] ?? ('Seri ' + (c - c0 + 1))) : 'Seri ' + (c - c0 + 1);
    const values = []; for (let r = r0; r < H; r++) { const v = vals[r][c]; values.push(typeof v === 'number' ? v : null); }
    series.push({ name, values, fmt: model.numFmt(si, g.r1 + r0, g.c1 + c) });
  }
  return { labels, series };
}

function niceStep(range, n) { const raw = range / n; const p = Math.pow(10, Math.floor(Math.log10(raw))); const m = raw / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }
function short(v) { const a = Math.abs(v); if (a >= 1e9) return X.fmtGeneral(+(v / 1e9).toFixed(1)) + ' Mr'; if (a >= 1e6) return X.fmtGeneral(+(v / 1e6).toFixed(1)) + ' Mn'; if (a >= 1e4) return X.fmtGeneral(+(v / 1e3).toFixed(1)) + ' B'; return X.formatValue(v, Number.isInteger(v) ? '#,##0' : '#,##0.##'); }

function draw(cv, spec, data, theme) {
  const dpr = G.devicePixelRatio || 1; const W = cv.clientWidth || spec.w || 480, H = cv.clientHeight || spec.h || 300;
  cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr);
  const ctx = cv.getContext('2d'); ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  const th = Object.assign({ bg: '#ffffff', fg: '#1f2937', muted: '#6b7280', grid: '#e5e7eb', font: '-apple-system, "Segoe UI", system-ui, sans-serif' }, theme || {});
  ctx.fillStyle = th.bg; ctx.fillRect(0, 0, W, H);
  const colors = spec.colors || COLORS;
  let top = 12;
  if (spec.title) { ctx.fillStyle = th.fg; ctx.font = `600 15px ${th.font}`; ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(spec.title, W / 2, 10, W - 20); top = 34; }
  const type = spec.type || 'column';
  const S = data.series.filter(s => s.values.some(v => v != null));
  if (!S.length) { ctx.fillStyle = th.muted; ctx.font = `13px ${th.font}`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText('Grafik için sayısal veri seçin', W / 2, H / 2); return; }
  // gösterge
  const pieLike = type === 'pie' || type === 'doughnut';
  const legendItems = pieLike ? data.labels.map((l, i) => [l, colors[i % colors.length]]) : S.map((s, i) => [s.name, colors[i % colors.length]]);
  ctx.font = `12px ${th.font}`; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
  let lx = 12, ly = top, legendH = 0;
  if (legendItems.length > 1 || pieLike) {
    const rows = [[]]; let w = 12;
    for (const [t] of legendItems) { const tw = Math.min(140, ctx.measureText(t).width) + 26; if (w + tw > W - 12 && rows[rows.length - 1].length) { rows.push([]); w = 12; } rows[rows.length - 1].push(tw); w += tw; }
    let idx = 0; legendH = rows.length * 18 + 6;
    rows.forEach((row, ri) => { const tot = row.reduce((a, b) => a + b, 0); lx = (W - tot) / 2; for (const tw of row) { const [t, c] = legendItems[idx++]; ctx.fillStyle = c; roundRect(ctx, lx, ly + ri * 18 + 3, 12, 12, 3); ctx.fill(); ctx.fillStyle = th.fg; ctx.fillText(t, lx + 17, ly + ri * 18 + 9, tw - 24); lx += tw; } });
    top += legendH;
  }
  const L = 12, R = W - 14, B = H - 10;
  if (pieLike) {
    const s = S[0]; const vals = s.values.map(v => Math.max(0, v || 0)); const tot = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2, cy = (top + B) / 2, rad = Math.max(10, Math.min(R - L, B - top) / 2 - 6);
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + v / tot * Math.PI * 2;
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, rad, a0, a1); ctx.closePath(); ctx.fillStyle = colors[i % colors.length]; ctx.fill();
      ctx.strokeStyle = th.bg; ctx.lineWidth = 2; ctx.stroke();
      if (v / tot > 0.04) { const am = (a0 + a1) / 2, rr = type === 'doughnut' ? rad * 0.78 : rad * 0.62; ctx.fillStyle = '#fff'; ctx.font = `600 12px ${th.font}`; ctx.textAlign = 'center'; ctx.fillText(X.formatValue(v / tot, '0%'), cx + Math.cos(am) * rr, cy + Math.sin(am) * rr); }
      a0 = a1;
    });
    if (type === 'doughnut') { ctx.beginPath(); ctx.arc(cx, cy, rad * 0.55, 0, Math.PI * 2); ctx.fillStyle = th.bg; ctx.fill(); ctx.fillStyle = th.fg; ctx.font = `700 15px ${th.font}`; ctx.textAlign = 'center'; ctx.fillText(short(tot), cx, cy); }
    return;
  }
  // eksenler
  const horiz = type === 'bar';
  const stacked = type === 'stacked';
  let lo = Infinity, hi = -Infinity;
  if (stacked) { data.labels.forEach((_, i) => { let p = 0, n = 0; S.forEach(s => { const v = s.values[i] || 0; if (v > 0) p += v; else n += v; }); hi = Math.max(hi, p); lo = Math.min(lo, n); }); }
  else S.forEach(s => s.values.forEach(v => { if (v != null) { lo = Math.min(lo, v); hi = Math.max(hi, v); } }));
  if (type !== 'line' && type !== 'scatter') { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (lo === hi) { hi += 1; lo -= type === 'line' ? 1 : 0; }
  const step = niceStep(hi - lo, 5); lo = Math.floor(lo / step) * step; hi = Math.ceil(hi / step) * step;
  ctx.font = `11px ${th.font}`;
  const ticks = []; for (let v = lo; v <= hi + step / 2; v += step) ticks.push(+v.toPrecision(12));
  const tw = Math.max(...ticks.map(t => ctx.measureText(short(t)).width));
  const n = data.labels.length;
  const labelW = horiz ? Math.min(110, Math.max(...data.labels.map(l => ctx.measureText(l).width))) + 8 : 0;
  const x0 = L + (horiz ? labelW : tw + 8), x1 = R, y0 = top + 4, y1 = B - (horiz ? 18 : 20);
  const scale = v => horiz ? x0 + (v - lo) / (hi - lo) * (x1 - x0) : y1 - (v - lo) / (hi - lo) * (y1 - y0);
  ctx.strokeStyle = th.grid; ctx.lineWidth = 1; ctx.fillStyle = th.muted;
  for (const t of ticks) {
    const p = Math.round(scale(t)) + 0.5;
    ctx.beginPath(); if (horiz) { ctx.moveTo(p, y0); ctx.lineTo(p, y1); } else { ctx.moveTo(x0, p); ctx.lineTo(x1, p); } ctx.stroke();
    if (horiz) { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(short(t), p, y1 + 4); } else { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(short(t), x0 - 6, p); }
  }
  const band = (horiz ? (y1 - y0) : (x1 - x0)) / Math.max(1, n);
  // kategori etiketleri
  ctx.fillStyle = th.muted; const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? (y1 - y0) / 16 : (x1 - x0) / 56)))));
  data.labels.forEach((l, i) => { if (i % every) return; const c = (horiz ? y0 : x0) + band * (i + 0.5); if (horiz) { ctx.textAlign = 'right'; ctx.textBaseline = 'middle'; ctx.fillText(l, x0 - 6, c, labelW - 6); } else { ctx.textAlign = 'center'; ctx.textBaseline = 'top'; ctx.fillText(l, c, y1 + 5, band * every - 4); } });
  if (type === 'column' || type === 'bar' || type === 'stacked') {
    const k = stacked ? 1 : S.length; const gap = band * 0.2; const bw = Math.max(1, (band - gap * 2) / k);
    const base = scale(Math.max(lo, Math.min(0, hi)));
    data.labels.forEach((_, i) => {
      let pos = 0, neg = 0;
      S.forEach((s, j) => {
        const v = s.values[i]; if (v == null) return;
        const st = (horiz ? y0 : x0) + band * i + gap + (stacked ? 0 : bw * j);
        let a, b;
        if (stacked) { if (v >= 0) { a = scale(pos); b = scale(pos + v); pos += v; } else { a = scale(neg); b = scale(neg + v); neg += v; } } else { a = base; b = scale(v); }
        ctx.fillStyle = colors[j % colors.length];
        const r = Math.min(4, bw / 3);
        if (horiz) roundRect(ctx, Math.min(a, b), st + 1, Math.abs(b - a), bw - 2, r); else roundRect(ctx, st + 1, Math.min(a, b), bw - 2, Math.abs(b - a), r);
        ctx.fill();
      });
    });
    if (spec.labels && !stacked && S.length === 1 && n <= 24) {
      ctx.fillStyle = th.fg; ctx.font = `600 11px ${th.font}`;
      S[0].values.forEach((v, i) => { if (v == null) return; const c = (horiz ? y0 : x0) + band * (i + 0.5); if (horiz) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(short(v), scale(v) + 4, c); } else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(short(v), c, scale(v) - 3); } });
    }
  } else {
    S.forEach((s, j) => {
      const col = colors[j % colors.length]; const pts = [];
      s.values.forEach((v, i) => { if (v != null) pts.push([x0 + band * (i + 0.5), scale(v)]); });
      if (!pts.length) return;
      if (type === 'scatter') { ctx.fillStyle = col; pts.forEach(([x, y]) => { ctx.beginPath(); ctx.arc(x, y, 4, 0, 7); ctx.fill(); }); return; }
      if (type === 'area') { ctx.beginPath(); ctx.moveTo(pts[0][0], scale(Math.max(lo, 0))); pts.forEach(([x, y]) => ctx.lineTo(x, y)); ctx.lineTo(pts[pts.length - 1][0], scale(Math.max(lo, 0))); ctx.closePath(); ctx.globalAlpha = 0.25; ctx.fillStyle = col; ctx.fill(); ctx.globalAlpha = 1; }
      ctx.beginPath(); pts.forEach(([x, y], i) => i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)); ctx.strokeStyle = col; ctx.lineWidth = 2.5; ctx.lineJoin = 'round'; ctx.stroke();
      if (pts.length <= 40) { ctx.fillStyle = th.bg; pts.forEach(([x, y]) => { ctx.beginPath(); ctx.arc(x, y, 3.5, 0, 7); ctx.fill(); ctx.strokeStyle = col; ctx.lineWidth = 2; ctx.stroke(); }); }
    });
  }
}
function roundRect(ctx, x, y, w, h, r) { r = Math.max(0, Math.min(r, w / 2, h / 2)); ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); }

// PNG olarak (kaydetme / xlsx içine gömme)
function toPNG(spec, data, w, h) {
  const cv = document.createElement('canvas'); cv.style.width = w + 'px'; cv.style.height = h + 'px';
  Object.defineProperty(cv, 'clientWidth', { value: w }); Object.defineProperty(cv, 'clientHeight', { value: h });
  const old = G.devicePixelRatio; draw(cv, spec, data);
  return cv.toDataURL('image/png');
}
G.XC = { extract, draw, toPNG, TYPES, COLORS };
})(window);
