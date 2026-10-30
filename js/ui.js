// Excello arayüz yardımcıları: ikonlar, diyaloglar, açılır menüler, bildirimler
(function (G) {
'use strict';
const P = {
  menu: '<path d="M4 6h16M4 12h16M4 18h16"/>',
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
  save: '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
  folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
  file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M12 12v6M9 15h6"/>',
  bold: '<path d="M7 4h6.5a4 4 0 0 1 0 8H7zM7 12h7.5a4 4 0 0 1 0 8H7z"/>',
  italic: '<path d="M19 4h-9M14 20H5M15 4 9 20"/>',
  underline: '<path d="M6 4v6a6 6 0 0 0 12 0V4M4 20h16"/>',
  strike: '<path d="M16 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H7M4 12h16"/>',
  alignL: '<path d="M3 6h18M3 12h12M3 18h15"/>', alignC: '<path d="M3 6h18M6 12h12M4 18h16"/>', alignR: '<path d="M3 6h18M9 12h12M6 18h15"/>',
  vTop: '<path d="M4 4h16M12 8v12M8 12l4-4 4 4"/>', vMid: '<path d="M4 12h16M12 3v5M12 16v5M9 6l3 2 3-2M9 18l3-2 3 2"/>', vBot: '<path d="M4 20h16M12 4v12M8 12l4 4 4-4"/>',
  wrap: '<path d="M3 6h18M3 12h15a3 3 0 1 1 0 6h-4M16 16l-2 2 2 2M3 18h7"/>',
  merge: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M8 12h8M13 9l3 3-3 3M11 9l-3 3 3 3"/>',
  fill: '<path d="m19 11-8-8-8.5 8.5a2 2 0 0 0 0 2.8l5.7 5.7a2 2 0 0 0 2.8 0z"/><path d="M5 2l5 5M2.5 13H19"/><path d="M22 20a2 2 0 1 1-4 0c0-1.6 2-3 2-4.5 0 1.5 2 2.9 2 4.5z"/>',
  color: '<path d="M7 16 12 4l5 12M9 12h6"/>',
  border: '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 12h18M12 3v18" stroke-dasharray="2 2"/>',
  sigma: '<path d="M18 7V4H6l6 8-6 8h12v-3"/>',
  chart: '<path d="M3 3v18h18"/><path d="M7 16v-4M12 16V8M17 16v-7"/>',
  pie: '<path d="M21.2 15.9A10 10 0 1 1 8 2.8"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
  line: '<path d="M3 3v18h18"/><path d="m7 15 4-5 4 3 5-7"/>',
  table: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>',
  filter: '<path d="M22 3H2l8 9.5V19l4 2v-8.5z"/>',
  sortA: '<path d="m3 8 4-4 4 4M7 4v16M11 12h4M11 16h7M11 20h10"/>',
  sortD: '<path d="m3 16 4 4 4-4M7 20V4M11 4h10M11 8h7M11 12h4"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
  plus: '<path d="M12 5v14M5 12h14"/>', minus: '<path d="M5 12h14"/>',
  trash: '<path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/>',
  rowAdd: '<rect x="3" y="4" width="18" height="6" rx="1"/><path d="M12 14v6M9 17h6"/>',
  colAdd: '<rect x="4" y="3" width="6" height="18" rx="1"/><path d="M14 12h6M17 9v6"/>',
  rowDel: '<rect x="3" y="4" width="18" height="6" rx="1"/><path d="M9 17h6"/>',
  colDel: '<rect x="4" y="3" width="6" height="18" rx="1"/><path d="M14 12h6"/>',
  freeze: '<path d="M12 2v20M2 12h20M5 5l14 14M19 5 5 19"/>',
  zoomIn: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3M11 8v6M8 11h6"/>',
  zoomOut: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3M8 11h6"/>',
  rules: '<circle cx="13.5" cy="6.5" r="1.5"/><circle cx="17.5" cy="10.5" r="1.5"/><circle cx="8.5" cy="7.5" r="1.5"/><circle cx="6.5" cy="12.5" r="1.5"/><path d="M12 2a10 10 0 1 0 0 20c1.1 0 2-.9 2-2 0-.5-.2-1-.5-1.3-.3-.4-.5-.8-.5-1.3 0-1.1.9-2 2-2h2.4A5.6 5.6 0 0 0 22 10c0-4.4-4.5-8-10-8z"/>',
  list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
  copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  cut: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.1 15.9M14.5 14.5 20 20M8.1 8.1 12 12"/>',
  paste: '<rect x="8" y="2" width="8" height="4" rx="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>',
  eraser: '<path d="m7 21-4.3-4.3a1 1 0 0 1 0-1.4l10-10a1 1 0 0 1 1.4 0l5.6 5.6a1 1 0 0 1 0 1.4L11 21M22 21H7M5 11l9 9"/>',
  moon: '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
  x: '<path d="M18 6 6 18M6 6l12 12"/>', check: '<path d="M20 6 9 17l-5-5"/>',
  down: '<path d="m6 9 6 6 6-6"/>', left: '<path d="m15 18-6-6 6-6"/>', right: '<path d="m9 18 6-6-6-6"/>',
  wand: '<path d="m21.6 2.4-1.4-1.4M3 21l12-12M14 4V2M14 12v-2M8 6H6M20 6h-2M17.4 8.6 19 10.2M10.6 3.4 9 1.8M17.4 3.4 19 1.8"/><path d="m15 9-2-2"/>',
  share: '<path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8M16 6l-4-4-4 4M12 2v13"/>',
  print: '<path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>',
  pointer: '<path d="m4 4 7 17 2.5-7.5L21 11z"/>',
  grid: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>',
  calendar: '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
  layers: '<path d="m12 2 10 5-10 5L2 7z"/><path d="m2 17 10 5 10-5M2 12l10 5 10-5"/>',
  fillDown: '<path d="M12 3v14M6 11l6 6 6-6M5 21h14"/>',
  dots: '<circle cx="12" cy="5" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="12" cy="19" r="1"/>',
  fit: '<path d="M21 3 14 10M3 21l7-7M14 3h7v7M10 21H3v-7"/>',
  eye: '<path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
  eyeOff: '<path d="M9.9 4.2A10 10 0 0 1 12 4c7 0 10 8 10 8a13 13 0 0 1-1.7 2.7M6.6 6.6A13.5 13.5 0 0 0 2 12s3 8 10 8a9.7 9.7 0 0 0 5.4-1.6M2 2l20 20M14.1 14.1a3 3 0 1 1-4.2-4.2"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
  info: '<circle cx="12" cy="12" r="9"/><path d="M12 16v-4M12 8h.01"/>',
  calc: '<rect x="4" y="2" width="16" height="20" rx="2"/><path d="M8 6h8M8 10h.01M12 10h.01M16 10h.01M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h4"/>',
  tag: '<path d="M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8z"/><path d="M7.5 7.5h.01"/>',
  hash: '<path d="M4 9h16M4 15h16M10 3 8 21M16 3l-2 18"/>',
};
function icon(n, cls) { return `<svg class="i ${cls || ''}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${P[n] || ''}</svg>`; }

const $ = (s, el) => (el || document).querySelector(s);
const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
function h(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  if (attrs) for (const k in attrs) {
    const v = attrs[k]; if (v == null || v === false) continue;
    if (k === 'class') e.className = v; else if (k === 'html') e.innerHTML = v; else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v); else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) e.append(k.nodeType ? k : document.createTextNode(k));
  return e;
}
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

// ---------- Bildirim ----------
let toastT;
function toast(msg, kind) {
  let t = $('#toast'); if (!t) { t = h('div', { id: 'toast' }); document.body.append(t); }
  t.className = 'show ' + (kind || ''); t.textContent = msg;
  clearTimeout(toastT); toastT = setTimeout(() => t.className = '', 2600);
}

// ---------- Diyalog ----------
const stack = [];
function dialog(opts) {
  // opts: {title, body(el|html), buttons:[{text, kind, value, onClick}], wide, onClose, cls}
  const back = h('div', { class: 'dlg-back' });
  const box = h('div', { class: 'dlg ' + (opts.wide ? 'wide ' : '') + (opts.cls || ''), role: 'dialog' });
  const head = h('div', { class: 'dlg-head' }, h('div', { class: 'dlg-title', html: opts.title || '' }), h('button', { class: 'ic ghost', 'aria-label': 'Kapat', html: icon('x'), onclick: () => close(null) }));
  const body = h('div', { class: 'dlg-body' });
  if (typeof opts.body === 'string') body.innerHTML = opts.body; else if (opts.body) body.append(opts.body);
  box.append(head, body);
  let foot = null;
  if (opts.buttons && opts.buttons.length) {
    foot = h('div', { class: 'dlg-foot' });
    for (const b of opts.buttons) foot.append(h('button', { class: 'btn ' + (b.kind || ''), text: b.text, onclick: async () => { if (b.onClick) { const r = await b.onClick(api); if (r === false) return; } if (b.value !== undefined) close(b.value); } }));
    box.append(foot);
  }
  back.append(box); document.body.append(back);
  requestAnimationFrame(() => back.classList.add('show'));
  back.addEventListener('mousedown', e => { if (e.target === back) close(null); });
  let resolve; const p = new Promise(r => resolve = r);
  function close(v) { if (!back.isConnected) return; back.classList.remove('show'); setTimeout(() => back.remove(), 160); stack.splice(stack.indexOf(api), 1); if (opts.onClose) opts.onClose(v); resolve(v); }
  const api = { el: box, body, back, foot, close, promise: p, hide(b) { back.classList.toggle('hidden', b); } };
  stack.push(api);
  return api;
}
function topDialog() { return stack[stack.length - 1]; }
function alert(msg, title) { return dialog({ title: title || 'Bilgi', body: h('p', { text: msg }), buttons: [{ text: 'Tamam', kind: 'primary', value: true }] }).promise; }
function confirm(msg, title, okText) { return dialog({ title: title || 'Onay', body: h('p', { text: msg }), buttons: [{ text: 'Vazgeç', value: false }, { text: okText || 'Tamam', kind: 'primary', value: true }] }).promise.then(v => !!v); }
function prompt(title, value, label) {
  const inp = h('input', { class: 'inp', value: value || '' });
  const d = dialog({ title, body: h('div', null, label ? h('label', { class: 'lbl', text: label }) : null, inp), buttons: [{ text: 'Vazgeç', value: null }, { text: 'Tamam', kind: 'primary', onClick: () => { d.close(inp.value); return false; } }] });
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') d.close(inp.value); });
  setTimeout(() => { inp.focus(); inp.select(); }, 50);
  return d.promise;
}

// ---------- Açılır menü ----------
let openPop = null;
function closePop() { if (openPop) { openPop.remove(); openPop = null; } }
function popover(anchor, content, opts) {
  closePop();
  const pop = h('div', { class: 'pop ' + ((opts && opts.cls) || '') });
  if (typeof content === 'string') pop.innerHTML = content; else pop.append(content);
  document.body.append(pop); openPop = pop;
  const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : anchor;
  const pw = pop.offsetWidth, ph = pop.offsetHeight, vw = innerWidth, vh = innerHeight;
  let x = r.left, y = r.bottom + 6;
  if (x + pw > vw - 8) x = Math.max(8, vw - pw - 8);
  if (y + ph > vh - 8) y = Math.max(8, r.top - ph - 6);
  pop.style.left = x + 'px'; pop.style.top = y + 'px';
  setTimeout(() => document.addEventListener('pointerdown', outside, true), 0);
  function outside(e) { if (!pop.contains(e.target)) { document.removeEventListener('pointerdown', outside, true); if (openPop === pop) closePop(); } }
  return pop;
}
function menu(anchor, items) {
  // items: [{text, icon, onClick, kbd, sep, danger, disabled}]
  const m = h('div', { class: 'menu' });
  for (const it of items) {
    if (it.sep) { m.append(h('div', { class: 'sep' })); continue; }
    if (it.head) { m.append(h('div', { class: 'mhead', text: it.head })); continue; }
    m.append(h('button', { class: 'mi' + (it.danger ? ' danger' : '') + (it.on ? ' on' : ''), disabled: it.disabled, onclick: () => { closePop(); it.onClick && it.onClick(); } },
      h('span', { class: 'mi-ic', html: it.icon ? icon(it.icon) : (it.swatch ? `<i class="sw" style="background:${it.swatch}"></i>` : '') }), h('span', { class: 'mi-t', text: it.text }), it.kbd ? h('span', { class: 'kbd', text: it.kbd }) : null));
  }
  return popover(anchor, m);
}

const PALETTE = ['#000000', '#3f3f46', '#71717a', '#a1a1aa', '#d4d4d8', '#f4f4f5', '#ffffff',
  '#dc2626', '#ea580c', '#d97706', '#ca8a04', '#16a34a', '#0d9488', '#0284c7', '#2563eb', '#7c3aed', '#c026d3', '#db2777',
  '#fecaca', '#fed7aa', '#fde68a', '#fef08a', '#bbf7d0', '#99f6e4', '#bae6fd', '#bfdbfe', '#ddd6fe', '#f5d0fe', '#fbcfe8',
  '#7f1d1d', '#7c2d12', '#78350f', '#713f12', '#14532d', '#134e4a', '#0c4a6e', '#1e3a8a', '#4c1d95', '#701a75', '#831843'];
function colorPicker(anchor, onPick, noneText) {
  const wrap = h('div', { class: 'cpick' });
  wrap.append(h('button', { class: 'cp-none', text: noneText || 'Renk yok', onclick: () => { closePop(); onPick(null); } }));
  const grid = h('div', { class: 'cp-grid' });
  for (const c of PALETTE) grid.append(h('button', { class: 'cp-sw', style: { background: c }, title: c, onclick: () => { closePop(); onPick(c); } }));
  wrap.append(grid);
  const custom = h('input', { type: 'color', class: 'cp-custom', oninput: e => onPick(e.target.value), onchange: () => closePop() });
  wrap.append(h('label', { class: 'cp-more' }, custom, h('span', { text: 'Özel renk…' })));
  return popover(anchor, wrap);
}

G.UI = { icon, $, $$, h, esc, toast, dialog, topDialog, alert, confirm, prompt, popover, closePop, menu, colorPicker, PALETTE };
})(window);
