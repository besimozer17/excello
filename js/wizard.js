// Excello sihirbazları: Hazır Hesaplar (formül yazmadan) ve İşlev Sihirbazı
(function (G) {
'use strict';
const X = G.XF, U = G.UI, { h, icon } = U;

// ---------- yardımcılar ----------
const RE_REF = /^(('([^']|'')+'|[^\s!'"(),;:+\-*\/^&=<>%]+)!)?\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$|^(('([^']|'')+'|[^\s!'"(),;:+\-*\/^&=<>%]+)!)?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/;
const isRef = s => RE_REF.test(String(s).trim());
const q = s => '"' + String(s).replace(/"/g, '""') + '"';
function valExpr(raw, numOnly) {
  const t = String(raw == null ? '' : raw).trim();
  if (t === '') return numOnly ? '0' : '""';
  if (isRef(t)) return t;
  if (t.startsWith('=')) return t.slice(1);
  const n = X.parseNumLoose(t); if (n != null) return String(n);
  const d = X.parseDateLoose(t); if (d) { const dt = X.serialToDate(Math.floor(d.v)); return `DATE(${dt.getUTCFullYear()},${dt.getUTCMonth() + 1},${dt.getUTCDate()})`; }
  if (numOnly) return t;
  return q(t);
}
function absRef(r) { const t = String(r).trim(); const i = t.lastIndexOf('!'); return t.slice(0, i + 1) + t.slice(i + 1).replace(/(^|:)\$?([A-Za-z]{1,3})\$?(\d*)/g, (m, p, c, d) => p + '$' + c.toUpperCase() + (d ? '$' + d : '')); }
function firstCell(r) { const t = String(r).trim(); const i = t.lastIndexOf('!'); const sheet = i >= 0 ? t.slice(0, i + 1) : ''; return sheet + t.slice(i + 1).split(':')[0]; }
function splitRef(r) {
  const t = String(r || '').trim(); const i = t.lastIndexOf('!'); const [a, b] = t.slice(i + 1).replace(/\$/g, '').split(':');
  const p = X.parseAddr(a || ''), e = X.parseAddr(b || a || ''); return p && e ? { sh: t.slice(0, i + 1), r1: p.r, c1: p.c, r2: e.r, c2: e.c } : null;
}
function lookupExpr(v) {
  const L = splitRef(v.look), R = splitRef(v.ret), key = firstCell(v.key);
  if (L && R && L.sh === R.sh && L.c1 === L.c2 && R.c1 === R.c2 && R.c1 > L.c1 && L.r1 === R.r1 && L.r2 === R.r2)
    return `IFERROR(VLOOKUP(${key},${L.sh}$${X.colName(L.c1)}$${L.r1 + 1}:$${X.colName(R.c1)}$${R.r2 + 1},${R.c1 - L.c1 + 1},FALSE),${valExpr(v.nf)})`;
  return `IFERROR(INDEX(${absRef(v.ret)},MATCH(${key},${absRef(v.look)},0)),${valExpr(v.nf)})`;
}
function refRows(r) {
  const t = String(r).trim(); const body = t.slice(t.lastIndexOf('!') + 1); const [a, b] = body.split(':');
  const p = X.parseAddr(a), e = b ? X.parseAddr(b) : p; if (!p || !e) return 1; return Math.abs(e.r - p.r) + 1;
}
const OPS = [['=', 'eşittir'], ['<>', 'eşit değildir'], ['>', 'büyüktür'], ['<', 'küçüktür'], ['>=', 'büyük veya eşit'], ['<=', 'küçük veya eşit'], ['contains', 'içerir'], ['begins', 'ile başlar']];
function critExpr(op, raw) {
  const t = String(raw == null ? '' : raw).trim();
  if (op === 'contains') return isRef(t) ? `"*"&${t}&"*"` : q('*' + t + '*');
  if (op === 'begins') return isRef(t) ? `${t}&"*"` : q(t + '*');
  if (isRef(t)) return op === '=' ? t : q(op) + '&' + t;
  const n = X.parseNumLoose(t);
  return q((op === '=' ? '' : op) + (n != null ? String(n) : t));
}
function condExpr(cell, op, raw) {
  const v = valExpr(raw);
  if (op === 'contains') return `ISNUMBER(SEARCH(${v},${cell}))`;
  if (op === 'begins') return `LEFT(${cell},LEN(${v}))=${v}`;
  return `${cell}${op}${v}`;
}

// ---------- Hazır hesaplar ----------
const F = (k, label, type, o) => Object.assign({ k, label, type }, o || {});
const RECIPES = [
  { g: 'Hızlı', id: 'sum', icon: 'sigma', t: 'Toplam', d: 'Seçili sayıları toplar', agg: 1, f: [F('rng', 'Toplanacak hücreler', 'range', { sel: 1 })], b: v => `SUM(${v.rng})` },
  { g: 'Hızlı', id: 'avg', icon: 'calc', t: 'Ortalama', d: 'Sayıların ortalaması', agg: 1, f: [F('rng', 'Hücreler', 'range', { sel: 1 })], b: v => `AVERAGE(${v.rng})`, nf: '#,##0.00' },
  { g: 'Hızlı', id: 'count', icon: 'hash', t: 'Say', d: 'Kaç tane sayı var', agg: 1, f: [F('rng', 'Hücreler', 'range', { sel: 1 })], b: v => `COUNT(${v.rng})` },
  { g: 'Hızlı', id: 'counta', icon: 'hash', t: 'Dolu hücre say', d: 'Boş olmayan hücreler', agg: 1, f: [F('rng', 'Hücreler', 'range', { sel: 1 })], b: v => `COUNTA(${v.rng})` },
  { g: 'Hızlı', id: 'max', icon: 'sortD', t: 'En büyük', d: 'En yüksek değer', agg: 1, f: [F('rng', 'Hücreler', 'range', { sel: 1 })], b: v => `MAX(${v.rng})` },
  { g: 'Hızlı', id: 'min', icon: 'sortA', t: 'En küçük', d: 'En düşük değer', agg: 1, f: [F('rng', 'Hücreler', 'range', { sel: 1 })], b: v => `MIN(${v.rng})` },
  { g: 'Hızlı', id: 'run', hd: 'Kümülatif Toplam', icon: 'line', t: 'Yürüyen toplam', d: 'Her satırda o ana kadarki toplam', f: [F('col', 'Sayı sütunu', 'cell', { sel: 1 })], row: 1, b: v => { const a = firstCell(v.col); return `SUM(${absRef(a)}:${a})`; } },

  { g: 'Koşullu', id: 'sumif', icon: 'filter', t: 'Koşula göre topla', d: 'Ör: "Ankara" olan satırların satış toplamı', agg: 1, f: [F('look', 'Hangi sütuna bakılsın?', 'range'), F('op', 'Koşul', 'op'), F('val', 'Değer', 'value', { ph: 'Ör: Ankara veya 100' }), F('sum', 'Hangi sütun toplansın?', 'range')], b: v => `SUMIF(${v.look},${critExpr(v.op, v.val)},${v.sum})` },
  { g: 'Koşullu', id: 'countif', icon: 'hash', t: 'Koşula göre say', d: 'Ör: kaç satırda "Evet" yazıyor', agg: 1, f: [F('look', 'Hangi hücrelerde?', 'range', { sel: 1 }), F('op', 'Koşul', 'op'), F('val', 'Değer', 'value')], b: v => `COUNTIF(${v.look},${critExpr(v.op, v.val)})` },
  { g: 'Koşullu', id: 'avgif', icon: 'calc', t: 'Koşula göre ortalama', d: 'Koşula uyan satırların ortalaması', agg: 1, f: [F('look', 'Hangi sütuna bakılsın?', 'range'), F('op', 'Koşul', 'op'), F('val', 'Değer', 'value'), F('avg', 'Ortalaması alınacak sütun', 'range')], b: v => `IFERROR(AVERAGEIF(${v.look},${critExpr(v.op, v.val)},${v.avg}),0)`, nf: '#,##0.00' },
  { g: 'Koşullu', id: 'if', icon: 'wand', t: 'Eğer… ise…', d: 'Koşula göre farklı sonuç yaz (ör: 50\'den büyükse "Geçti")', row: 1, f: [F('cell', 'Hangi sütuna bakılsın?', 'cell', { sel: 1 }), F('op', 'Koşul', 'op'), F('val', 'Neyle karşılaştırılsın?', 'value', { ph: 'Ör: 50' }), F('yes', 'Koşula uyanlara ne yazılsın?', 'value', { def: 'Geçti' }), F('no', 'Uymayanlara ne yazılsın?', 'value', { def: 'Kaldı' })], hd: 'Durum', b: v => `IF(${condExpr(firstCell(v.cell), v.op, v.val)},${valExpr(v.yes)},${valExpr(v.no)})` },
  { g: 'Koşullu', id: 'lookup', hd: 'Bulunan', icon: 'search', t: 'Tablodan değer getir (DÜŞEYARA)', d: 'Bir koda/isme karşılık gelen bilgiyi başka tablodan ya da başka sayfadan bulur. Ör: ürün adına göre Ürünler sayfasından kategorisini getirir.', row: 1, f: [F('key', 'Neyi arayalım? (bu tablodaki sütun)', 'cell', { sel: 1 }), F('look', 'Nerede arayalım? (diğer listedeki sütun)', 'range'), F('ret', 'Oradan hangi bilgi gelsin?', 'range'), F('nf', 'Bulunamazsa yaz', 'value', { def: 'Bulunamadı' })], b: v => lookupExpr(v) },
  { g: 'Koşullu', id: 'rank', hd: 'Sıra', icon: 'sortD', t: 'Sıralama (kaçıncı?)', d: 'Değerin listedeki sırasını verir', row: 1, f: [F('cell', 'Sıralanacak hücre / sütun', 'cell', { sel: 1 }), F('list', 'Tüm liste', 'range'), F('ord', 'Sıra', 'select', { opts: [['0', 'Büyükten küçüğe (1 = en büyük)'], ['1', 'Küçükten büyüğe']] })], b: v => `RANK(${firstCell(v.cell)},${absRef(v.list || v.cell)},${v.ord || 0})` },

  { g: 'Para ve Yüzde', id: 'arith', hd: 'Sonuç', icon: 'calc', t: 'Dört işlem', d: 'İki hücre veya sütun arasında + − × ÷', row: 1, f: [F('a', '1. hücre / sütun', 'cell', { sel: 1 }), F('op', 'İşlem', 'select', { opts: [['*', '× Çarp'], ['+', '+ Topla'], ['-', '− Çıkar'], ['/', '÷ Böl']] }), F('b', '2. hücre / sütun ya da sayı', 'cellnum')], b: v => { const a = firstCell(v.a), bb = isRef(v.b) ? firstCell(v.b) : valExpr(v.b, true); return v.op === '/' ? `IFERROR(${a}/${bb},0)` : `${a}${v.op || '*'}${bb}`; } },
  { g: 'Para ve Yüzde', id: 'pct', hd: 'Pay %', icon: 'pie', t: 'Yüzde payı', d: 'Bir değerin toplam içindeki yüzdesi', row: 1, nf: '0.00%', f: [F('part', 'Değer (hücre / sütun)', 'cell', { sel: 1 }), F('tot', 'Toplam (hücre ya da sayı)', 'cellnum', { ph: 'Boş bırakırsanız sütun toplamı' })], b: v => { const p = firstCell(v.part); const t = v.tot ? (isRef(v.tot) ? absRef(firstCell(v.tot)) : valExpr(v.tot, true)) : `SUM(${absRef(v.part)})`; return `IFERROR(${p}/${t},0)`; } },
  { g: 'Para ve Yüzde', id: 'chg', hd: 'Değişim %', icon: 'line', t: 'Yüzde değişim', d: 'Eski → yeni değer arasındaki artış/azalış', row: 1, nf: '0.00%', f: [F('old', 'Eski değer', 'cell', { sel: 1 }), F('new', 'Yeni değer', 'cell')], b: v => `IFERROR((${firstCell(v.new)}-${firstCell(v.old)})/${firstCell(v.old)},0)` },
  { g: 'Para ve Yüzde', id: 'kdv', hd: 'KDV Dahil', icon: 'tag', t: 'KDV ekle', d: 'KDV hariç tutara KDV ekler', row: 1, nf: '#,##0.00 ₺', f: [F('a', 'Tutar (hücre / sütun)', 'cell', { sel: 1 }), F('r', 'KDV oranı %', 'select', { opts: [['20', '%20'], ['10', '%10'], ['1', '%1'], ['18', '%18'], ['8', '%8']] })], b: v => `ROUND(${firstCell(v.a)}*(1+${v.r || 20}/100),2)` },
  { g: 'Para ve Yüzde', id: 'kdvx', hd: 'KDV Hariç', icon: 'tag', t: 'KDV\'yi ayır', d: 'KDV dahil fiyattan KDV hariç tutarı bulur', row: 1, nf: '#,##0.00 ₺', f: [F('a', 'KDV dahil tutar', 'cell', { sel: 1 }), F('r', 'KDV oranı %', 'select', { opts: [['20', '%20'], ['10', '%10'], ['1', '%1'], ['18', '%18'], ['8', '%8']] })], b: v => `ROUND(${firstCell(v.a)}/(1+${v.r || 20}/100),2)` },
  { g: 'Para ve Yüzde', id: 'kdvt', hd: 'KDV Tutarı', icon: 'tag', t: 'KDV tutarı', d: 'Tutarın KDV miktarı', row: 1, nf: '#,##0.00 ₺', f: [F('a', 'KDV hariç tutar', 'cell', { sel: 1 }), F('r', 'KDV oranı %', 'select', { opts: [['20', '%20'], ['10', '%10'], ['1', '%1'], ['18', '%18'], ['8', '%8']] })], b: v => `ROUND(${firstCell(v.a)}*${v.r || 20}/100,2)` },
  { g: 'Para ve Yüzde', id: 'disc', hd: 'İndirimli Fiyat', icon: 'tag', t: 'İndirim uygula', d: 'Fiyattan % indirim düşer', row: 1, nf: '#,##0.00 ₺', f: [F('a', 'Fiyat', 'cell', { sel: 1 }), F('p', 'İndirim %', 'cellnum', { def: '10' })], b: v => `ROUND(${firstCell(v.a)}*(1-${isRef(v.p) ? absRef(firstCell(v.p)) : valExpr(v.p, true)}/100),2)` },
  { g: 'Para ve Yüzde', id: 'loan', icon: 'calc', t: 'Kredi taksiti', d: 'Aylık taksit tutarını hesaplar', nf: '#,##0.00 ₺', f: [F('amt', 'Kredi tutarı', 'cellnum', { ph: 'Ör: 100000' }), F('rate', 'Aylık faiz %', 'cellnum', { ph: 'Ör: 3,5' }), F('n', 'Vade (ay)', 'cellnum', { ph: 'Ör: 12' })], b: v => `-PMT(${valExpr(v.rate, true)}/100,${valExpr(v.n, true)},${valExpr(v.amt, true)})` },
  { g: 'Para ve Yüzde', id: 'round', hd: 'Yuvarlanmış', icon: 'hash', t: 'Yuvarla', d: 'Sayıyı belirtilen haneye yuvarlar', row: 1, f: [F('a', 'Hücre / sütun', 'cell', { sel: 1 }), F('d', 'Virgülden sonra kaç hane?', 'select', { opts: [['0', '0 (tam sayı)'], ['1', '1'], ['2', '2'], ['-1', 'Onluğa'], ['-2', 'Yüzlüğe'], ['-3', 'Binliğe']] })], b: v => `ROUND(${firstCell(v.a)},${v.d || 0})` },

  { g: 'Liste ve Özet', id: 'pivot', icon: 'table', t: 'Özet tablo', d: 'Gruplara göre toplar/sayar. Ör: her şehrin toplam satışı. Veri değişince kendiliğinden güncellenir.', multi: 1, f: [F('grp', 'Neye göre gruplansın?', 'range', { kind: 'text' }), F('val', 'Hangi sütun hesaplansın?', 'range', { kind: 'num' }), F('op', 'Ne yapılsın?', 'select', { opts: [['SUM', 'Topla'], ['COUNT', 'Kaç tane (say)'], ['AVG', 'Ortalama'], ['MAX', 'En büyük'], ['MIN', 'En küçük']] })],
    m: (v, t) => { const g = absRef(v.grp), val = absRef(v.val || v.grp); const u = 'UNIQUE(' + g + ')'; const op = v.op || 'SUM'; const f = op === 'SUM' ? 'SUMIF(' + g + ',' + u + ',' + val + ')' : op === 'COUNT' ? 'COUNTIF(' + g + ',' + u + ')' : op === 'AVG' ? 'AVERAGEIF(' + g + ',' + u + ',' + val + ')' : op === 'MAX' ? 'MAXIFS(' + val + ',' + g + ',' + u + ')' : 'MINIFS(' + val + ',' + g + ',' + u + ')'; const nm = { SUM: 'Toplam', COUNT: 'Adet', AVG: 'Ortalama', MAX: 'En büyük', MIN: 'En küçük' }[op]; return [{ dc: 0, f: u, hd: t.colName(v.grp) }, { dc: 1, f, hd: nm + (op === 'COUNT' ? '' : ' ' + t.colName(v.val)), nf: op === 'COUNT' ? null : t.colNf(v.val) }]; } },
  { g: 'Liste ve Özet', id: 'uniq', icon: 'layers', t: 'Benzersiz liste', d: 'Bir sütundaki her değeri bir kez listeler (tekrarsız).', multi: 1, f: [F('col', 'Hangi sütun?', 'range', { sel: 1 })], m: (v, t) => [{ dc: 0, f: 'UNIQUE(' + absRef(v.col) + ')', hd: t.colName(v.col) }] },
  { g: 'Liste ve Özet', id: 'flt', icon: 'filter', t: 'Koşula uyan satırları listele', d: 'Ör: Şehri "Ankara" olan tüm satırları ayrı bir yere listeler. Veri değişince güncellenir.', multi: 1, f: [F('col', 'Hangi sütuna bakılsın?', 'range', { sel: 1 }), F('op', 'Koşul', 'op'), F('val', 'Neyle karşılaştırılsın?', 'value')],
    m: (v, t) => [{ dc: 0, f: 'FILTER(' + absRef(t.dataRef) + ',' + condExpr(absRef(v.col), v.op, v.val) + ',"Uyan satır yok")', hdAll: 1 }] },
  { g: 'Liste ve Özet', id: 'topn', icon: 'sortD', t: 'En büyük / en küçük N kayıt', d: 'Ör: en çok satış yapılan 5 satırı listeler.', multi: 1, f: [F('col', 'Hangi sütuna göre?', 'range', { kind: 'num' }), F('n', 'Kaç kayıt?', 'cellnum', { def: '5' }), F('dir', 'Hangileri?', 'select', { opts: [['-1', 'En büyükler'], ['1', 'En küçükler']] })],
    m: (v, t) => [{ dc: 0, f: 'TAKE(SORTBY(' + absRef(t.dataRef) + ',' + absRef(v.col) + ',' + (v.dir || -1) + '),' + valExpr(v.n, true) + ')', hdAll: 1 }] },
  { g: 'Liste ve Özet', id: 'cntu', icon: 'hash', t: 'Kaç farklı değer var?', d: 'Ör: tabloda kaç farklı şehir var.', agg: 1, f: [F('col', 'Hangi sütun?', 'range', { sel: 1 })], b: v => 'COUNTA(UNIQUE(' + v.col + '))' },
  { g: 'Tarih', id: 'today', icon: 'calendar', t: 'Bugünün tarihi', d: 'Her açılışta güncellenir', nf: 'dd.mm.yyyy', f: [], b: () => 'TODAY()' },
  { g: 'Tarih', id: 'diff', hd: 'Fark', icon: 'calendar', t: 'İki tarih arası', d: 'Gün, ay, yıl veya iş günü farkı', row: 1, f: [F('s', 'Başlangıç tarihi', 'cell', { sel: 1 }), F('e', 'Bitiş tarihi (boşsa bugün)', 'cell'), F('u', 'Birim', 'select', { opts: [['D', 'Gün'], ['M', 'Ay'], ['Y', 'Yıl'], ['W', 'İş günü (hafta içi)']] })], b: v => { const s = firstCell(v.s), e = v.e ? firstCell(v.e) : 'TODAY()'; const u = v.u || 'D'; return u === 'W' ? `NETWORKDAYS(${s},${e})` : u === 'D' ? `IFERROR(DATEDIF(${s},${e},"D"),0)` : `IFERROR(DATEDIF(${s},${e},"${u}"),0)`; }, nf: '0' },
  { g: 'Tarih', id: 'age', hd: 'Yaş', icon: 'calendar', t: 'Yaş hesapla', d: 'Doğum tarihinden bugünkü yaş', row: 1, nf: '0', f: [F('b', 'Doğum tarihi', 'cell', { sel: 1 })], b: v => `DATEDIF(${firstCell(v.b)},TODAY(),"Y")` },
  { g: 'Tarih', id: 'add', hd: 'Yeni Tarih', icon: 'calendar', t: 'Tarihe ekle', d: 'Gün / ay / yıl ekleyerek yeni tarih (vade, son gün)', row: 1, nf: 'dd.mm.yyyy', f: [F('d', 'Tarih', 'cell', { sel: 1 }), F('n', 'Kaç tane?', 'cellnum', { def: '30' }), F('u', 'Birim', 'select', { opts: [['D', 'Gün'], ['M', 'Ay'], ['Y', 'Yıl'], ['W', 'İş günü']] })], b: v => { const d = firstCell(v.d), n = isRef(v.n) ? absRef(firstCell(v.n)) : valExpr(v.n, true); return v.u === 'M' ? `EDATE(${d},${n})` : v.u === 'Y' ? `EDATE(${d},${n}*12)` : v.u === 'W' ? `WORKDAY(${d},${n})` : `${d}+${n}`; } },
  { g: 'Tarih', id: 'part', hd: 'Tarih Bilgisi', icon: 'calendar', t: 'Tarihten bilgi al', d: 'Yıl, ay, gün ya da gün adı', row: 1, f: [F('d', 'Tarih', 'cell', { sel: 1 }), F('w', 'Ne alınsın?', 'select', { opts: [['Y', 'Yıl'], ['M', 'Ay (sayı)'], ['MN', 'Ay adı'], ['D', 'Gün'], ['WD', 'Gün adı'], ['WK', 'Hafta no']] })], b: v => { const d = firstCell(v.d); return { Y: `YEAR(${d})`, M: `MONTH(${d})`, MN: `TEXT(${d},"mmmm")`, D: `DAY(${d})`, WD: `TEXT(${d},"dddd")`, WK: `WEEKNUM(${d},2)` }[v.w || 'Y']; } },

  { g: 'Metin', id: 'join', hd: 'Birleşik', icon: 'layers', t: 'Metinleri birleştir', d: 'Ad + Soyad gibi iki sütunu birleştirir', row: 1, f: [F('a', '1. hücre / sütun', 'cell', { sel: 1 }), F('b', '2. hücre / sütun', 'cell'), F('s', 'Arada ne olsun?', 'select', { opts: [[' ', 'Boşluk'], ['', 'Hiçbir şey'], [', ', 'Virgül'], [' - ', 'Tire'], ['/', 'Eğik çizgi']] })], b: v => `${firstCell(v.a)}&${q(v.s == null ? ' ' : v.s)}&${firstCell(v.b)}` },
  { g: 'Metin', id: 'case', hd: 'Düzeltilmiş', icon: 'italic', t: 'Büyük / küçük harf', d: 'Metni BÜYÜK, küçük ya da Baş Harfleri Büyük yapar', row: 1, f: [F('a', 'Hücre / sütun', 'cell', { sel: 1 }), F('m', 'Biçim', 'select', { opts: [['UPPER', 'BÜYÜK HARF'], ['LOWER', 'küçük harf'], ['PROPER', 'Baş Harfler Büyük']] })], b: v => `${v.m || 'UPPER'}(${firstCell(v.a)})` },
  { g: 'Metin', id: 'trim', hd: 'Temiz Metin', icon: 'eraser', t: 'Fazla boşlukları sil', d: 'Baştaki, sondaki ve çift boşlukları temizler', row: 1, f: [F('a', 'Hücre / sütun', 'cell', { sel: 1 })], b: v => `TRIM(${firstCell(v.a)})` },
  { g: 'Metin', id: 'cut', hd: 'Parça', icon: 'cut', t: 'Metinden parça al', d: 'Baştan veya sondan belirli sayıda karakter', row: 1, f: [F('a', 'Hücre / sütun', 'cell', { sel: 1 }), F('w', 'Nereden?', 'select', { opts: [['LEFT', 'Baştan'], ['RIGHT', 'Sondan']] }), F('n', 'Kaç karakter?', 'cellnum', { def: '3' })], b: v => `${v.w || 'LEFT'}(${firstCell(v.a)},${valExpr(v.n, true)})` },
  { g: 'Metin', id: 'len', hd: 'Karakter Sayısı', icon: 'hash', t: 'Karakter say', d: 'Metnin uzunluğu', row: 1, f: [F('a', 'Hücre / sütun', 'cell', { sel: 1 })], b: v => `LEN(${firstCell(v.a)})` },
  { g: 'Metin', id: 'rnd', icon: 'wand', t: 'Rastgele sayı', d: 'İki sayı arasında rastgele tam sayı', f: [F('a', 'En az', 'cellnum', { def: '1' }), F('b', 'En çok', 'cellnum', { def: '100' })], b: v => `RANDBETWEEN(${valExpr(v.a, true)},${valExpr(v.b, true)})` },
];

class Wizards {
  constructor(app) { this.app = app; }
  get wb() { return this.app.wb; }
  selRef() { return this.app.refText(this.app.sel, this.app.wb.active, this.target ? this.target.si : this.app.wb.active); }

  // Seçili hücrenin bulunduğu tablo: başlıklar ve veri satırları
  tableInfo() {
    const app = this.app, wb = this.wb, si = wb.active, s = app.sel;
    let g = app.currentRegion(s.r, s.c);
    if (g.r1 === g.r2 && g.c1 === g.c2 && wb.value(si, s.r, s.c) == null) { const u = wb.usedRange(si); if (wb.dims(si).rows) g = u; }
    const header = app.detectHeader(g) || (g.r2 > g.r1 && (() => { for (let c = g.c1; c <= g.c2; c++) if (typeof wb.value(si, g.r1, c) !== 'string') return false; return true; })());
    const r0 = header ? g.r1 + 1 : g.r1;
    // alttaki "Toplam" satırlarını veri sayma
    const isTotal = r => { for (let c = g.c1; c <= g.c2; c++) { const v = wb.value(si, r, c), cl = wb.cell(si, r, c); if (typeof v === 'string' && /^(genel\s+)?(toplam|total|ara toplam)/i.test(v.trim())) return true; if (cl && cl.f && /^(SUM|SUBTOTAL|AVERAGE)\(/.test(cl.f)) { try { const m = /\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/.exec(cl.f); if (m && +m[4] - 1 < r) return true; } catch (e) { } } } return false; };
    while (g.r2 > r0 && isTotal(g.r2)) g = Object.assign({}, g, { r2: g.r2 - 1 });
    const cols = [];
    if (wb.dims(si).rows) for (let c = g.c1; c <= g.c2; c++) cols.push({ c, name: header ? wb.display(si, g.r1, c) : '', ref: X.addr(r0, c) + (g.r2 > r0 ? ':' + X.addr(g.r2, c) : '') });
    const dataRef = X.addr(r0, g.c1) + ':' + X.addr(Math.max(g.r2, r0), g.c2);
    const self = this;
    return { g, header, r0, cols, si, dataRef,
      colName(ref) { const c = cols.find(x => x.ref === String(ref || '').trim()); return c ? (c.name || X.colName(c.c)) : 'Değer'; },
      colNf(ref) { const c = cols.find(x => x.ref === String(ref || '').trim()); return c ? self.wb.numFmt(si, r0, c.c) : null; } };
  }
  colLabel(c) { return X.colName(c.c) + ' sütunu' + (c.name ? ' — ' + c.name : ''); }

  // alan seçici kontrolü: sütun listesi + tablodan seç
  pickField(field, inp, dlg, onChange, opts) {
    const pickBtn = h('button', { class: 'btn pickbtn', type: 'button', html: icon('table') + '<span>Tablodan seç</span>', onclick: () => {
      this.miniPick({ label: field.label, initial: inp.value, single: field.single, onDone: ref => { inp.value = ref; sync(); onChange && onChange(); } });
    } });
    const t = this.table;
    if (!(opts && opts.cols) || !t || !t.cols.length) { var sync = () => { }; return h('div', { class: 'pickf' }, inp, pickBtn); }
    const wb = this.wb;
    const tableRef = X.addr(t.r0 - (t.header ? 1 : 0), t.g.c1) + ':' + X.addr(t.g.r2, t.g.c2);
    const dataRef = X.addr(t.r0, t.g.c1) + ':' + X.addr(t.g.r2, t.g.c2);
    const sel = h('select', { class: 'inp colsel' }, h('option', { value: '', text: '— Listeden seçin —' }),
      ...t.cols.map(c => h('option', { value: c.ref, text: this.colLabel(c) + (c.st ? ' (' + ({ num: 'sayı', date: 'tarih', text: 'metin', empty: 'boş' })[c.st.kind] + ')' : '') })),
      h('option', { value: dataRef, text: `Tüm tablo (${X.colName(t.g.c1)}–${X.colName(t.g.c2)} sütunları)` }),
      h('option', { value: '__other', text: 'Başka bir alan (tablodan seçildi)', disabled: true }));
    // Alan: Tümü / Belirli satırlar
    const bAll = h('button', { class: 'seg', type: 'button', text: 'Tümü' });
    const bPart = h('button', { class: 'seg', type: 'button', text: 'Belirli satırlar…' });
    const summary = h('small', { class: 'scope-sum' });
    const parse = () => { const v = inp.value.trim().replace(/^.*!/, '').replace(/\$/g, ''); const [a, b] = v.split(':'); const p = X.parseAddr(a || ''), q = X.parseAddr(b || a || ''); return p && q ? { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) } : null; };
    const colOf = () => { const g = parse(); if (!g) return null; return t.cols.find(c => c.c === g.c1 && g.c1 === g.c2) || (g.c1 === t.g.c1 && g.c2 === t.g.c2 ? { all: 1 } : null); };
    var sync = () => {
      const v = inp.value.trim(); const g = parse(); const col = colOf();
      const exact = t.cols.find(c => c.ref === v) || (v === dataRef ? { ref: dataRef } : null);
      sel.value = exact ? exact.ref : col ? (col.all ? dataRef : col.ref) : (v ? '__other' : '');
      const full = g && g.r1 === t.r0 && g.r2 === t.g.r2;
      bAll.classList.toggle('on', !!(g && col && full)); bPart.classList.toggle('on', !!(g && col && !full));
      bAll.disabled = bPart.disabled = !col;
      summary.textContent = !g ? '' : !col ? 'Seçilen alan: ' + v : full ? `${g.r2 - g.r1 + 1} satırın tamamı (${g.r1 + 1}–${g.r2 + 1}. satırlar)` : `${g.r1 + 1}–${g.r2 + 1}. satırlar (${g.r2 - g.r1 + 1} satır)`;
    };
    const setRows = (r1, r2) => { const g = parse(); if (!g) return; inp.value = X.addr(r1, g.c1) + (r2 > r1 || g.c2 > g.c1 ? ':' + X.addr(r2, g.c2) : ''); sync(); onChange && onChange(); };
    bAll.addEventListener('click', () => setRows(t.r0, t.g.r2));
    bPart.addEventListener('click', () => {
      const g = parse(); if (!g) return; const c = g.c1;
      let a = g.r1, b = g.r2, step = 0;
      const list = h('div', { class: 'rowpick' });
      const head = h('div', { class: 'rp-head' }, h('b', { text: 'Hangi satırlar?' }), h('small', { text: 'Önce ilk satıra, sonra son satıra dokunun' }));
      const paint = () => U.$$('.rp-row', list).forEach(el => { const r = +el.dataset.r; el.classList.toggle('in', r >= a && r <= b); el.classList.toggle('edge', r === a || r === b); });
      for (let r = t.r0; r <= t.g.r2; r++) {
        const txt = wb.display(t.si, r, c) || '(boş)';
        list.append(h('button', { class: 'rp-row', type: 'button', 'data-r': r, onclick: () => { if (step === 0) { a = b = r; step = 1; } else { if (r < a) { b = a; a = r; } else b = r; step = 0; } paint(); } }, h('small', { text: (r + 1) + '.' }), h('span', { text: txt })));
      }
      const foot = h('div', { class: 'rp-foot' },
        h('button', { class: 'btn', type: 'button', text: 'Tümü', onclick: () => { a = t.r0; b = t.g.r2; step = 0; paint(); } }),
        h('button', { class: 'btn primary', type: 'button', text: 'Tamam', onclick: () => { U.closePop(); setRows(a, b); } }));
      U.popover(bPart, h('div', { class: 'rp' }, head, list, foot), { cls: 'rp-pop' }); paint();
    });
    sel.addEventListener('change', () => { if (sel.value === '__other') return; inp.value = sel.value; sync(); onChange && onChange(); });
    inp.addEventListener('input', sync); inp.style.display = 'none'; sync();
    return h('div', { class: 'colf' }, h('div', { class: 'pickf' }, sel, pickBtn), h('div', { class: 'scope' }, h('div', { class: 'segs' }, bAll, bPart), summary), inp);
  }

  // Küçük tablo penceresi: tabloyu görerek dokunup seçme (sütun adları büyük, harfler küçük)
  miniPick(o) {
    const wb = this.wb, base = wb.active;
    let si = base, sel = null, anchor = null, mode = null;
    const parseBody = b => { const [a, c] = b.replace(/\$/g, '').split(':'); const p = X.parseAddr(a || ''), q2 = X.parseAddr(c || a || ''); return p && q2 ? { r1: Math.min(p.r, q2.r), c1: Math.min(p.c, q2.c), r2: Math.max(p.r, q2.r), c2: Math.max(p.c, q2.c) } : null; };
    const init = String(o.initial || '').trim();
    if (init) { const i = init.lastIndexOf('!'); if (i >= 0) { const nm = init.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"); const k = wb.sheets.findIndex(x => x.name === nm); if (k >= 0) si = k; } sel = parseBody(init.slice(i + 1)); }
    // başlık satırı: ilk satırdaki dolu hücrelerin hepsi metinse
    const hasHead = k => { const d = wb.dims(k); if (d.rows < 2) return false; let n = 0; for (let c = 0; c < d.cols; c++) { const v = wb.value(k, 0, c); if (v == null || v === '') continue; if (typeof v !== 'string') return false; n++; } return n > 0; };
    const isTotal = (k, r) => { const d = wb.dims(k); for (let c = 0; c < d.cols; c++) { const v = wb.value(k, r, c), cl = wb.cell(k, r, c); if (typeof v === 'string' && /^(genel\s+)?(toplam|total|ara toplam)/i.test(v.trim())) return true; if (cl && cl.f && /^(SUM|SUBTOTAL)\(/.test(cl.f)) return true; } return false; };
    const colSpan = c => { const r0 = hasHead(si) ? 1 : 0; let r2 = -1; const d = wb.dims(si); for (let r = d.rows - 1; r >= r0; r--) { const v = wb.value(si, r, c); if (v != null && v !== '' && !isTotal(si, r)) { r2 = r; break; } } return { r1: r0, r2: Math.max(r0, r2) }; };
    const headName = c => hasHead(si) ? wb.display(si, 0, c) : '';
    const wrap = h('div', { class: 'mp' });
    const tabs = h('div', { class: 'segs mp-tabs' });
    const info = h('div', { class: 'mp-info' });
    const hint = h('small', { class: 'mp-hint', text: o.single ? 'Bir hücreye dokunun.' : 'Sütun adına dokunun: o sütundaki tüm veriler seçilir. Ya da önce ilk, sonra son hücreye dokunun.' });
    const box = h('div', { class: 'mp-box' });
    if (wb.sheets.length > 1) { wb.sheets.forEach((sh, k) => tabs.append(h('button', { class: 'seg', type: 'button', text: sh.name, 'data-k': k, onclick: () => { if (si === k) return; si = k; sel = null; mode = null; render(); } }))); wrap.append(tabs); }
    wrap.append(hint, box, info);
    const describe = () => {
      if (!sel) return 'Henüz seçim yok';
      const sh = si !== base ? wb.sheets[si].name + ' sayfası · ' : '';
      const n = sel.r2 - sel.r1 + 1;
      if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return sh + X.addr(sel.r1, sel.c1) + ' hücresi' + (headName(sel.c1) ? ' (' + headName(sel.c1) + ')' : '') + ': ' + (wb.display(si, sel.r1, sel.c1) || 'boş');
      if (sel.c1 === sel.c2) return sh + (headName(sel.c1) || X.colName(sel.c1) + ' sütunu') + ' · ' + (sel.r1 + 1) + '–' + (sel.r2 + 1) + '. satırlar (' + n + ' değer)';
      const names = []; for (let c = sel.c1; c <= sel.c2; c++) names.push(headName(c) || X.colName(c));
      return sh + names.join(', ') + ' · ' + n + ' satır';
    };
    const ref = () => { if (!sel) return ''; const t = X.addr(sel.r1, sel.c1) + (sel.r1 === sel.r2 && sel.c1 === sel.c2 ? '' : ':' + X.addr(sel.r2, sel.c2)); return si !== base ? X.quoteSheet(wb.sheets[si].name) + '!' + t : t; };
    const paint = () => {
      U.$$('td[data-r]', box).forEach(td => { const r = +td.dataset.r, c = +td.dataset.c; const on = sel && r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2; td.classList.toggle('in', !!on); td.classList.toggle('a', !!(anchor && mode === 'cell' && anchor.r === r && anchor.c === c)); });
      U.$$('th[data-c]', box).forEach(th => { const c = +th.dataset.c; th.classList.toggle('on', !!(sel && c >= sel.c1 && c <= sel.c2)); });
      U.$$('.seg', tabs).forEach(b => b.classList.toggle('on', +b.dataset.k === si));
      info.textContent = describe();
    };
    const render = () => {
      const d = wb.dims(si); const R = Math.min(Math.max(d.rows + 1, 8), 400), C = Math.min(Math.max(d.cols + 1, 4), 40); const hh = hasHead(si);
      const thead = h('tr', null, h('th', { class: 'mp-corner' }));
      for (let c = 0; c < C; c++) { const nm = headName(c); thead.append(h('th', { 'data-c': c, title: 'Sütunun tamamını seç' }, h('b', { text: nm || X.colName(c) }), nm ? h('small', { text: X.colName(c) }) : null)); }
      const tb = h('tbody');
      for (let r = hh ? 1 : 0; r < R; r++) { const tr = h('tr', null, h('th', { class: 'mp-rn', text: String(r + 1) })); for (let c = 0; c < C; c++) tr.append(h('td', { 'data-r': r, 'data-c': c, text: wb.display(si, r, c) })); tb.append(tr); }
      box.innerHTML = ''; box.append(h('table', { class: 'mp-t' }, h('thead', null, thead), tb)); paint();
      if (sel) { const td = box.querySelector(`td[data-r="${sel.r1}"][data-c="${sel.c1}"]`); if (td) setTimeout(() => td.scrollIntoView({ block: 'center', inline: 'center' }), 30); }
    };
    box.addEventListener('click', e => {
      const th = e.target.closest('th[data-c]'), td = e.target.closest('td[data-r]');
      if (th) {
        const c = +th.dataset.c, sp = colSpan(c);
        if (o.single) { sel = { r1: sp.r1, c1: c, r2: sp.r1, c2: c }; anchor = { r: sp.r1, c }; mode = 'cell'; }
        else if (mode === 'col' && anchor && anchor.c !== c) { const a = colSpan(anchor.c); sel = { r1: Math.min(a.r1, sp.r1), r2: Math.max(a.r2, sp.r2), c1: Math.min(anchor.c, c), c2: Math.max(anchor.c, c) }; mode = null; }
        else { sel = { r1: sp.r1, r2: sp.r2, c1: c, c2: c }; anchor = { c }; mode = 'col'; }
      } else if (td) {
        const r = +td.dataset.r, c = +td.dataset.c;
        if (mode === 'cell' && anchor && !o.single) { sel = { r1: Math.min(anchor.r, r), r2: Math.max(anchor.r, r), c1: Math.min(anchor.c, c), c2: Math.max(anchor.c, c) }; mode = null; }
        else { sel = { r1: r, r2: r, c1: c, c2: c }; anchor = { r, c }; mode = 'cell'; }
      } else return;
      paint();
    });
    render();
    const dl = U.dialog({ title: icon('table') + ' ' + (o.label || 'Alan seç'), body: wrap, wide: true, cls: 'mp-dlg', buttons: [
      { text: 'Vazgeç', value: null },
      { text: 'Tamam', kind: 'primary', onClick: () => { if (!sel) { U.toast('Önce bir alan seçin'); return false; } o.onDone(ref()); }, value: true }] });
    return dl;
  }

  // ---------- Hazır hesaplar ----------
  openRecipes(groupFilter) {
    const app = this.app;
    const body = h('div', { class: 'recipes' });
    const groups = [...new Set(RECIPES.map(r => r.g))];
    const chips = h('div', { class: 'chips' });
    let cur = groupFilter || 'Tümü';
    const list = h('div', { class: 'rgrid' });
    const fill = () => {
      list.innerHTML = '';
      for (const r of RECIPES) if (cur === 'Tümü' || r.g === cur) list.append(h('button', { class: 'rcard', onclick: () => { d.close(); this.openRecipe(r); } }, h('span', { class: 'ric', html: icon(r.icon) }), h('span', { class: 'rtx' }, h('b', { text: r.t }), h('small', { text: r.d }))));
      U.$$('.chip', chips).forEach(c => c.classList.toggle('on', c.textContent === cur));
    };
    ['Tümü', ...groups].forEach(g => chips.append(h('button', { class: 'chip', text: g, onclick: () => { cur = g; fill(); } })));
    body.append(chips, list); fill();
    const d = U.dialog({ title: icon('wand') + ' Hazır Hesaplar', body, wide: true });
    d.body.prepend(h('p', { class: 'hint', text: 'Formül yazmanıza gerek yok: bir hesap seçin, hücreleri tablodan dokunarak gösterin, sonuç otomatik yazılsın.' }));
  }
  // Sütun türü ve değerleri (hazır hesapların tabloyu "okuması" için)
  colStats(col) {
    const t = this.table, wb = this.wb, si = t.si; let num = 0, date = 0, txt = 0, n = 0, mn = Infinity, mx = -Infinity; const cnt = new Map();
    for (let r = t.r0; r <= t.g.r2; r++) {
      const v = wb.value(si, r, col.c); if (v == null || v === '') continue; n++;
      if (typeof v === 'number') { if (X.isDateFormat(wb.numFmt(si, r, col.c))) date++; else num++; if (v < mn) mn = v; if (v > mx) mx = v; } else txt++;
      const d = wb.display(si, r, col.c); cnt.set(d, (cnt.get(d) || 0) + 1);
    }
    const kind = !n ? 'empty' : date >= num && date >= txt ? 'date' : num >= txt ? 'num' : 'text';
    const top = [...cnt.entries()].sort((a, b) => b[1] - a[1]).map(x => x[0]);
    return { kind, n, mn, mx, top, distinct: cnt.size, nf: wb.numFmt(si, t.r0, col.c) };
  }
  statText(st) {
    if (st.kind === 'empty') return 'Boş sütun';
    const f = v => X.formatValue(v, st.nf || null);
    if (st.kind === 'num') return `${st.n} satır · sayı · en küçük ${f(st.mn)} · en büyük ${f(st.mx)}`;
    if (st.kind === 'date') return `${st.n} satır · tarih · ${f(st.mn)} – ${f(st.mx)}`;
    return `${st.n} satır · metin · ${st.distinct} farklı değer` + (st.top.length ? ': ' + st.top.slice(0, 4).join(', ') + (st.distinct > 4 ? '…' : '') : '');
  }
  kindOf(rc, f) {
    if (f.kind) return f.kind;
    if (rc.g === 'Tarih' && f.type === 'cell') return 'date';
    if (rc.g === 'Metin' && f.type === 'cell' && rc.id !== 'len') return 'text';
    if ((rc.g === 'Hızlı' && rc.id !== 'counta') || rc.g === 'Para ve Yüzde') return 'num';
    return null;
  }
  openRecipe(rc) {
    const app = this.app, wb = this.wb, si = wb.active;
    const t = this.table = this.tableInfo();
    t.cols.forEach(c => c.st = this.colStats(c));
    const body = h('div', { class: 'form' });
    body.append(h('p', { class: 'hint', text: rc.d }));
    const preview = h('div', { class: 'preview' });
    let d; const inputs = {}; const colFields = [];
    const s = app.sel; const multi = s.r1 !== s.r2 || s.c1 !== s.c2;
    const used = new Set();
    const findCol = ref => t.cols.find(c => c.ref === String(ref || '').trim());
    const refreshChips = () => colFields.forEach(cf => (cf.refreshers || []).forEach(fn => fn()));
    for (const f of rc.f) {
      const row = h('div', { class: 'frow' }, h('label', { class: 'lbl', text: f.label }));
      let inp;
      if (f.type === 'op') { inp = h('select', { class: 'inp' }, ...OPS.map(([v, tx]) => h('option', { value: v, text: tx }))); inp.value = '='; row.append(inp); }
      else if (f.type === 'select') { inp = h('select', { class: 'inp' }, ...f.opts.map(([v, tx]) => h('option', { value: v, text: tx }))); row.append(inp); }
      else {
        inp = h('input', { class: 'inp', placeholder: f.ph || '', value: f.def || '' });
        const isCol = f.type === 'range' || f.type === 'cell';
        if (isCol) {
          const kind = this.kindOf(rc, f);
          let pre = null;
          if (multi && f.sel) inp.value = this.selRef();
          else {
            const ac = t.cols.find(c => c.c === s.c);
            if (ac && !used.has(ac.c) && (f.sel || kind) && (!kind || ac.st.kind === kind)) pre = ac;
            else if (kind) { const m = t.cols.filter(c => c.st.kind === kind && !used.has(c.c)); pre = kind === 'num' && !f.sel ? m[m.length - 1] : m[0]; }
            if (pre) { inp.value = pre.ref; used.add(pre.c); }
          }
          const info = h('div', { class: 'colinfo' });
          const cf = { f, inp, refreshers: [] };
          cf.showInfo = () => { const c = findCol(inp.value); info.textContent = c ? this.statText(c.st) : (inp.value ? 'Alan: ' + inp.value : ''); cf.refreshers.forEach(fn => fn()); };
          row.append(this.pickField(f, inp, () => d, () => { cf.showInfo(); upd(); }, { cols: true }), info);
          colFields.push(cf);
        } else if (f.type === 'value') {
          // bağlı sütunun gerçek değerleri dokunulabilir düğmeler olarak
          const src = colFields[colFields.length - 1];
          const chips = h('div', { class: 'vchips' });
          const mark = () => U.$$('.chip', chips).forEach(x => x.classList.toggle('on', x.textContent === inp.value));
          inp.addEventListener('input', mark);
          const fill = () => {
            chips.innerHTML = ''; const c = src && findCol(src.inp.value); if (!c || /^(yes|no|nf)$/.test(f.k)) return;
            if (c.st.kind === 'num') { inp.placeholder = (f.ph || 'Değer') + ` — bu sütunda ${X.formatValue(c.st.mn, c.st.nf)} ile ${X.formatValue(c.st.mx, c.st.nf)} arası`; return; }
            c.st.top.slice(0, 12).forEach(v => chips.append(h('button', { class: 'chip sm', type: 'button', text: v, onclick: () => { inp.value = v; mark(); upd(); } })));
            mark();
          };
          if (src) src.refreshers.push(fill);
          row.append(inp, chips);
        } else row.append(this.pickField(f, inp, () => d, upd));
      }
      inp.addEventListener('input', upd); inp.addEventListener('change', upd);
      inputs[f.k] = inp; body.append(row);
    }
    // DÜŞEYARA: aranan değerlerin geçtiği sütunu diğer sayfalarda/listelerde kendiliğinden bul
    if (rc.id === 'lookup' && inputs.key && inputs.look && !inputs.look.value) {
      const kc = findCol(inputs.key.value); const keys = new Set();
      if (kc) for (let r = t.r0; r <= t.g.r2; r++) { const v = wb.value(si, r, kc.c); if (v != null && v !== '') keys.add(String(v).toLocaleLowerCase('tr-TR')); }
      let best = null;
      wb.sheets.forEach((sh, k) => {
        const d = wb.dims(k); if (!d.rows) return;
        for (let c = 0; c < d.cols; c++) {
          if (k === si && c >= t.g.c1 && c <= t.g.c2) continue;
          let hit = 0, last = 0; for (let r = 1; r < d.rows; r++) { const v = wb.value(k, r, c); if (v == null || v === '') continue; last = r; if (keys.has(String(v).toLocaleLowerCase('tr-TR'))) hit++; }
          if (hit && (!best || hit > best.hit)) best = { k, c, hit, last, cols: d.cols };
        }
      });
      if (best) {
        const pre = best.k !== si ? X.quoteSheet(wb.sheets[best.k].name) + '!' : '';
        inputs.look.value = pre + X.addr(1, best.c) + ':' + X.addr(best.last, best.c);
        if (inputs.ret && !inputs.ret.value && best.c + 1 < best.cols) inputs.ret.value = pre + X.addr(1, best.c + 1) + ':' + X.addr(best.last, best.c + 1);
        const hn = best.c + 1 < best.cols ? wb.display(best.k, 0, best.c + 1) : '';
        setTimeout(() => { const hi = U.$$('input', body).find(i => i.value === rc.hd); if (hi && hn) hi.value = hn; [inputs.look, inputs.ret, hi].forEach(i => i && i.dispatchEvent(new Event('input'))); }, 0);
      }
    }
    // hedef: satır bazlı -> tablonun sağındaki ilk boş sütun; toplu -> sütunun altı
    let auto = true;
    const autoTarget = () => {
      if (!t.cols.length) return { si, r: s.r, c: s.c };
      if (rc.multi) {
        // tablonun sağında, yeterince geniş boş alan bul
        const w = rc.id === 'pivot' ? 2 : rc.id === 'uniq' ? 1 : t.g.c2 - t.g.c1 + 1; const top = t.header ? t.g.r1 : t.r0;
        const free = c0 => { for (let c = c0 - 1; c < c0 + w; c++) for (let r = top; r <= t.g.r2 + 60; r++) if (c >= 0 && wb.value(si, r, c) != null && wb.value(si, r, c) !== '') return false; return true; };
        let c = t.g.c2 + 2; while (!free(c) && c < t.g.c2 + 300) c++;
        return { si, r: t.r0, c };
      }
      if (rc.row) { const f0 = rc.f.find(f => f.type === 'cell'); const v = f0 && inputs[f0.k].value; const p = v && X.parseAddr(firstCell(v).replace(/^.*!/, '').replace(/\$/g, '')); return { si, r: p ? p.r : t.r0, c: t.g.c2 + 1 }; }
      if (rc.agg) { const first = rc.f.find(f => f.type === 'range' || f.type === 'cell'); const col = first && findCol(inputs[first.k].value); if (col) return { si, r: t.g.r2 + 1, c: col.c }; }
      const cv = wb.cell(si, s.r, s.c); if (!cv || (cv.v == null && cv.f == null)) return { si, r: s.r, c: s.c };
      return { si, r: t.g.r2 + 2, c: t.g.c1 };
    };
    this.target = autoTarget();
    const tgtInp = h('input', { class: 'inp', value: X.addr(this.target.r, this.target.c) });
    const tgtInfo = h('div', { class: 'colinfo' });
    const hdrInp = h('input', { class: 'inp', value: rc.hd || rc.t.replace(/[….]+$/, '').trim() });
    body.append(h('div', { class: 'frow target' }, h('label', { class: 'lbl', html: icon('right') + ' Sonuç nereye yazılsın?' }),
      this.pickField({ label: 'Sonuç hücresi', type: 'cell', single: 1 }, tgtInp, () => d, () => { const g = X.parseAddr(firstCell(tgtInp.value).replace(/^.*!/, '')); if (g) { this.target = { si: this.wb.active, r: g.r, c: g.c }; auto = false; } upd(); }), tgtInfo));
    if (rc.row && t.header) body.append(h('div', { class: 'frow' }, h('label', { class: 'lbl', text: 'Yeni sütunun başlığı' }), hdrInp));
    body.append(preview);
    const self = this;
    function collect() { const v = {}; for (const f of rc.f) v[f.k] = inputs[f.k].value.trim(); return v; }
    function plan() {
      const vals = collect();
      for (const f of rc.f) if ((f.type === 'range' || f.type === 'cell') && !vals[f.k] && !/boşsa|Boş bırak/i.test(f.label + (f.ph || ''))) return { err: '“' + f.label + '” için listeden bir sütun seçin.' };
      let n = 1;
      if (rc.row) for (const f of rc.f) if (f.type === 'cell' && vals[f.k] && isRef(vals[f.k])) n = Math.max(n, refRows(vals[f.k]));
      if (rc.multi) { let items; try { items = rc.m(vals, t).map(it => Object.assign({}, it, { f: X.toStore(it.f) })); items.forEach(it => X.parse(it.f)); } catch (e) { return { err: 'Girdileri kontrol edin.' }; } return { f: items[items.length - 1].f, n: 1, items }; }
      let f; try { f = X.toStore(rc.b(vals)); X.parse(f); } catch (e) { return { err: 'Girdileri kontrol edin.' }; }
      return { f, n };
    }
    function upd() {
      if (auto) self.target = autoTarget();
      const p = plan(); const tg = self.target;
      tgtInp.value = X.addr(tg.r, tg.c);
      tgtInfo.textContent = rc.row && p.n > 1 ? `${X.colName(tg.c)} sütununa, ${p.n} satıra yazılacak` : `${X.addr(tg.r, tg.c)} hücresine yazılacak`;
      preview.innerHTML = '';
      if (p.err) { preview.append(h('div', { class: 'pv-err', text: p.err })); return; }
      if (p.items) {
        const cols = p.items.map(it => { try { const r = X.evaluate(X.parse(it.f), { model: wb, si: tg.si, r: tg.r, c: tg.c + it.dc }); return { g: X.isRng(r) ? r.grid() : [[r]], nf: it.nf, all: it.hdAll }; } catch (e) { return { g: [[X.isErr(e) ? e : '?']] }; } });
        const H = Math.max(...cols.map(c => c.g.length));
        preview.append(h('div', { class: 'pv-res' }, h('span', { text: 'Sonuç' }), h('b', { text: H + ' satırlık liste' })));
        const ul = h('div', { class: 'pv-rows' });
        for (let i = 0; i < Math.min(H, 6); i++) { const parts = []; cols.forEach(c => (c.g[i] || []).forEach((x, j) => parts.push(X.formatValue(x, c.nf || (c.all ? wb.numFmt(t.si, t.r0, t.g.c1 + j) : null) || null)))); ul.append(h('div', null, h('span', { text: parts.slice(0, 4).join('  ·  ') }))); }
        if (H > 6) ul.append(h('div', { class: 'more', text: '… ve ' + (H - 6) + ' satır daha' }));
        preview.append(ul);
        tgtInfo.textContent = X.addr(tg.r, tg.c) + ' hücresinden başlayarak yazılacak (tablonun sağında)';
        preview.append(h('details', { class: 'pv-f' }, h('summary', { text: 'Arka planda kullanılan formül' }), ...p.items.map(it => h('code', { text: '=' + X.toDisplay(it.f) }))));
        return;
      }
      const v = self.app.evalFormula(p.f, tg);
      preview.append(h('div', { class: 'pv-res' }, h('span', { text: 'Sonuç' + (p.n > 1 ? ' (ilk satır)' : '') }), h('b', { text: X.formatValue(v, rc.nf || null) || '—' })));
      if (p.n > 1) {
        const ul = h('div', { class: 'pv-rows' });
        for (let i = 0; i < Math.min(p.n, 5); i++) { const fi = i ? X.shiftFormula(p.f, i, 0) : p.f; const vi = self.app.evalFormula(fi, { si: tg.si, r: tg.r + i, c: tg.c }); ul.append(h('div', null, h('small', { text: (tg.r + i + 1) + '. satır' }), h('span', { text: X.formatValue(vi, rc.nf || null) || '(boş)' }))); }
        if (p.n > 5) ul.append(h('div', { class: 'more', text: `… ve ${p.n - 5} satır daha` }));
        preview.append(ul);
      }
      preview.append(h('details', { class: 'pv-f' }, h('summary', { text: 'Arka planda kullanılan formül' }), h('code', { text: '=' + X.toDisplay(p.f) })));
    }
    d = U.dialog({ title: icon(rc.icon) + ' ' + rc.t, body, buttons: [{ text: 'Geri', onClick: () => { d.close(); this.openRecipes(); return false; } }, { text: 'Uygula', kind: 'primary', onClick: () => {
      const p = plan(); if (p.err) { U.toast(p.err, 'warn'); return false; }
      const tg = this.target; let hdr = null;
      if (p.items) {
        wb.begin(); const hs = (wb.cell(tg.si, t.g.r1, t.g.c1) || {}).s;
        for (const it of p.items) {
          wb.patchCell(tg.si, tg.r, tg.c + it.dc, { f: it.f, v: null }); if (it.nf) wb.setStyle(tg.si, tg.r, tg.c + it.dc, { nf: it.nf });
          if (t.header && tg.r === t.r0) {
            if (it.hdAll) { for (let c = t.g.c1; c <= t.g.c2; c++) { wb.patchCell(tg.si, t.g.r1, tg.c + (c - t.g.c1), { v: wb.value(tg.si, t.g.r1, c), f: null }); if (hs) wb.setStyle(tg.si, t.g.r1, tg.c + (c - t.g.c1), hs); } }
            else if (it.hd) { wb.patchCell(tg.si, t.g.r1, tg.c + it.dc, { v: it.hd, f: null }); if (hs) wb.setStyle(tg.si, t.g.r1, tg.c + it.dc, hs); }
          }
          if (it.hdAll) for (let c = t.g.c1; c <= t.g.c2; c++) { const nf = wb.numFmt(tg.si, t.r0, c); if (nf) wb.setStyle(tg.si, tg.r, tg.c + (c - t.g.c1), { nf }); }
        }
        wb.end();
        app.setSel({ r: tg.r, c: tg.c, r1: tg.r, c1: tg.c, r2: tg.r, c2: tg.c }, true);
        d.close(); U.toast('✓ ' + rc.t + ' oluşturuldu'); return false;
      }
      if (rc.row && t.header && tg.r >= t.r0 && tg.c > t.g.c2 && hdrInp.value.trim() && wb.value(tg.si, t.g.r1, tg.c) == null) hdr = { r: t.g.r1, c: tg.c, text: hdrInp.value.trim(), s: (wb.cell(tg.si, t.g.r1, t.g.c2) || {}).s };
      app.applyFormula(tg.si, tg.r, tg.c, p.f, p.n, rc.nf, hdr); d.close(); U.toast('✓ ' + rc.t + ' uygulandı' + (p.n > 1 ? ` (${p.n} satır)` : '')); return false;
    } }] });
    colFields.forEach(cf => cf.showInfo());
    upd();
  }

  // ---------- İşlev sihirbazı ----------
  openFunctions(preName, nested) {
    this.nested = nested || null;
    const app = this.app;
    if (preName && X.FN[preName]) return this.openFn(X.FN[preName]);
    // mevcut formülü düzenle
    if (!nested) { const c = this.wb.cell(this.wb.active, app.sel.r, app.sel.c); if (c && c.f) { const m = /^([A-Z][A-Z0-9.]*)\((.*)\)$/s.exec(c.f); if (m && X.FN[m[1]] && balanced(m[2])) return this.openFn(X.FN[m[1]], splitArgs(X.toDisplay(c.f))); } }
    const cats = ['Sık kullanılan', ...new Set(Object.values(X.FN).map(f => f.cat))];
    const fav = ['SUM', 'AVERAGE', 'IF', 'COUNTIF', 'SUMIF', 'VLOOKUP', 'XLOOKUP', 'ROUND', 'MAX', 'MIN', 'COUNTA', 'CONCATENATE', 'TODAY', 'DATEDIF', 'IFERROR', 'TEXT'];
    let cur = 'Sık kullanılan';
    const search = h('input', { class: 'inp search', placeholder: 'İşlev ara (ör: topla, tarih, ara)…' });
    const chips = h('div', { class: 'chips' }); const list = h('div', { class: 'flist' });
    const fill = () => {
      const qy = search.value.trim().toLocaleLowerCase('tr-TR');
      list.innerHTML = '';
      let items = Object.values(X.FN);
      if (qy) items = items.filter(f => (f.tr + ' ' + f.en + ' ' + f.desc).toLocaleLowerCase('tr-TR').includes(qy));
      else if (cur === 'Sık kullanılan') items = fav.map(n => X.FN[n]); else items = items.filter(f => f.cat === cur);
      for (const f of items) list.append(h('button', { class: 'fitem', onclick: () => { closing = true; d.close(); this.openFn(f); } }, h('b', { text: f.desc.replace(/\.$/, '') }), h('small', { text: f.tr + ' · ' + f.cat })));
      if (!items.length) list.append(h('p', { class: 'hint', text: 'Sonuç yok.' }));
      U.$$('.chip', chips).forEach(c => c.classList.toggle('on', !qy && c.textContent === cur));
    };
    cats.forEach(c => chips.append(h('button', { class: 'chip', text: c, onclick: () => { cur = c; search.value = ''; fill(); } })));
    search.addEventListener('input', fill);
    let closing = false;
    const d = U.dialog({ title: '<span class="fx">fx</span> ' + (nested ? 'Kutuya işlev koy' : 'Tüm İşlevler'), body: h('div', null, search, chips, list), wide: true, onClose: () => { if (!closing && this.nested) { const n = this.nested; this.nested = null; n.done(); } } });
    fill();
  }
  openFn(fd, pre) {
    const app = this.app; const wb = this.wb;
    this.target = { si: wb.active, r: app.sel.r, c: app.sel.c };
    this.table = this.tableInfo(); this.table.cols.forEach(c => c.st = this.colStats(c));
    const body = h('div', { class: 'form' });
    body.append(h('p', { class: 'hint', html: `<b>${fd.tr}</b> <span class="en">(${fd.en})</span> — ${U.esc(fd.desc)}` }));
    const rows = h('div'); body.append(rows);
    const preview = h('div', { class: 'preview' });
    const fields = []; let d;
    const sel = this.selRef(); const multi = app.sel.r1 !== app.sel.r2 || app.sel.c1 !== app.sel.c2;
    const addField = (a, idx, val) => {
      const same = fields.filter(f => f.a === a).length;
      let type = a.t, label = a.many ? a.n.replace('…', '').trim() + ' ' + (same + 1) : a.n;
      if (a.pair && a.pairT) { type = a.pairT[same % 2]; label = a.pairL[same % 2] + ' ' + (Math.floor(same / 2) + (a.o ? 2 : 1)); }
      const optional = a.o || (a.many && !a.pair && same > 0);
      const row = h('div', { class: 'frow' }, h('label', { class: 'lbl' }, label, optional ? h('span', { class: 'opt', text: ' (isteğe bağlı)' }) : null, a.d ? h('small', { text: ' · ' + a.d }) : null));
      let inp;
      const valuesMenu = (anchor, srcInp, target) => {
        const c = this.table && this.table.cols.find(x => x.ref === String(srcInp && srcInp.value || '').trim());
        const vals = c ? c.st.top.slice(0, 40) : [];
        if (!vals.length) { U.toast('Önce bir sütun seçin'); return; }
        U.menu(anchor, [{ head: (c.name || X.colName(c.c)) + ' sütunundaki değerler' }, ...vals.map(v => ({ text: v, onClick: () => { target.value = v; target.dispatchEvent(new Event('input')); } }))]);
      };
      const nestBtn = target => h('button', { class: 'btn nestbtn', type: 'button', title: 'Bu kutuya başka bir işlev koy', text: 'ƒ', onclick: () => {
        const keep = { target: this.target, table: this.table };
        d.hide(true);
        this.openFunctions(null, { onPick: txt => { target.value = txt; target.dispatchEvent(new Event('input')); }, done: () => { Object.assign(this, keep); d.hide(false); upd(); } });
      } });
      if (type === 'enum') { inp = h('select', { class: 'inp' }, a.o ? h('option', { value: '', text: '— varsayılan —' }) : null, ...a.opts.map(([v, t]) => h('option', { value: v, text: t }))); if (a.def) inp.value = a.def; row.append(inp); }
      else if (type === 'bool') { inp = h('select', { class: 'inp' }, a.o ? h('option', { value: '', text: '— varsayılan —' }) : null, h('option', { value: 'TRUE', text: 'Evet' }), h('option', { value: 'FALSE', text: 'Hayır' })); if (a.def) inp.value = a.def; row.append(inp); }
      else if (type === 'cond') {
        // [sütun/hücre] [koşul] [değer]
        inp = h('input', { type: 'hidden' });
        const left = h('input', { class: 'inp' }), op = h('select', { class: 'inp' }, ...OPS.map(([v, t]) => h('option', { value: v, text: t }))), right = h('input', { class: 'inp', placeholder: 'Değer (ör: 50 veya Ankara)' });
        let raw = null;
        if (val != null && val !== '') { const m = /^([^<>=]+?)(<>|>=|<=|=|>|<)(.+)$/.exec(val); if (m && isRef(m[1].trim())) { left.value = m[1].trim(); op.value = m[2]; right.value = m[3].trim().replace(/^"(.*)"$/, '$1'); } else raw = val; }
        else if (idx === 0 && multi) left.value = sel;
        const rebuild = () => { inp.value = raw != null ? '=' + raw : left.value.trim() ? '=' + condExpr(left.value.trim(), op.value, right.value) : ''; upd(); };
        [op, right].forEach(x => { x.addEventListener('input', rebuild); x.addEventListener('change', rebuild); });
        const vbtn = h('button', { class: 'btn', type: 'button', html: icon('list'), title: 'Sütundaki değerlerden seç', onclick: e => valuesMenu(e.currentTarget, left, right) });
        if (raw != null) { const rin = h('input', { class: 'inp', value: raw, oninput: () => { raw = rin.value; rebuild(); } }); row.append(rin); }
        else row.append(h('div', { class: 'condb' }, this.pickField({ label: label + ' — sütun ya da hücre', type: 'range' }, left, () => d, rebuild, { cols: true }), h('div', { class: 'pair' }, op, h('div', { class: 'pickf' }, right, vbtn))));
        rebuild();
      }
      else if (type === 'crit') {
        // ölçüt: [koşul] [değer]
        inp = h('input', { type: 'hidden' });
        const COPS = [['=', 'eşittir'], ['<>', 'eşit değildir'], ['>', 'büyüktür'], ['<', 'küçüktür'], ['>=', 'büyük veya eşit'], ['<=', 'küçük veya eşit'], ['*', 'içerir'], ['b', 'ile başlar']];
        const op = h('select', { class: 'inp' }, ...COPS.map(([v, t]) => h('option', { value: v, text: t }))), v2 = h('input', { class: 'inp', placeholder: 'Değer (ör: Ankara veya 100)' });
        if (val != null) { let t = String(val).replace(/^"(.*)"$/, '$1'); const m = /^(<>|>=|<=|=|>|<)(.*)$/.exec(t); if (m) { op.value = m[1]; t = m[2]; } if (/^\*.*\*$/.test(t)) { op.value = '*'; t = t.slice(1, -1); } else if (/\*$/.test(t)) { op.value = 'b'; t = t.slice(0, -1); } v2.value = t; }
        const rebuild = () => { const o = op.value, t = v2.value.trim(); if (!t) { inp.value = ''; upd(); return; }
          if (isRef(t)) inp.value = '=' + (o === '=' ? t : o === '*' ? `"*"&${t}&"*"` : o === 'b' ? `${t}&"*"` : `"${o}"&${t}`);
          else inp.value = o === '=' ? t : o === '*' ? '*' + t + '*' : o === 'b' ? t + '*' : o + t; upd(); };
        [op, v2].forEach(x => { x.addEventListener('input', rebuild); x.addEventListener('change', rebuild); });
        const src = fields.length ? fields[fields.length - 1].srcInp : null;
        const vbtn = h('button', { class: 'btn', type: 'button', html: icon('list'), title: 'Sütundaki değerlerden seç', onclick: e => valuesMenu(e.currentTarget, src, v2) });
        row.append(h('div', { class: 'pair' }, op, h('div', { class: 'pickf' }, v2, vbtn))); rebuild();
      }
      else {
        inp = h('input', { class: 'inp', placeholder: type === 'range' ? 'Listeden sütun seçin' : type === 'num' ? 'Sayı yazın ya da hücre seçin' : type === 'date' ? 'Tarih (ör: 30.10.2026) ya da hücre' : 'Değer yazın ya da hücre seçin', value: val != null ? val : (a.def && type !== 'range' ? a.def : '') });
        if (val == null && idx === 0 && multi && type === 'range') inp.value = sel;
        if (type === 'range') { if (!inp.value) { const ac = this.table.cols.find(c => c.c === app.sel.c); if (ac && idx === 0) inp.value = ac.ref; } row.append(this.pickField({ label, type: 'range' }, inp, () => d, upd, { cols: true })); }
        else row.append(h('div', { class: 'pickf' }, this.pickField({ label, type: 'cellnum' }, inp, () => d, upd), nestBtn(inp)));
      }
      inp.addEventListener('input', upd); inp.addEventListener('change', upd);
      fields.push({ a, inp, row, type, srcInp: type === 'range' ? inp : null }); rows.append(row);
    };
    let pi = 0;
    for (const a of fd.args) {
      if (a.many) { const n = Math.max(a.pair ? (a.o ? 0 : 2) : 1, pre ? pre.length - pi : 0); for (let k = 0; k < n; k++) addField(a, pi, pre ? pre[pi++] : null); }
      else addField(a, pi, pre ? pre[pi++] : null);
    }
    const manyA = fd.args.find(a => a.many);
    if (manyA) body.append(h('button', { class: 'btn ghost add', type: 'button', html: icon('plus') + (manyA.pair ? ' ' + (manyA.pairL ? manyA.pairL[0] : 'Çift') + ' ekle' : ' Bir tane daha ekle'), onclick: () => { addField(manyA, 99, null); if (manyA.pair) addField(manyA, 99, null); upd(); } }));
    body.append(preview);
    const argExpr = (a, raw) => {
      const t = raw.trim(); if (t === '') return '';
      if (a.t === 'enum' || a.t === 'bool') return t;
      if (isRef(t)) return t;
      if (t.startsWith('=')) return t.slice(1);
      if (/^".*"$/.test(t)) return t;
      if (a.t === 'crit') return critText(t);
      if (a.t === 'num' || a.t === 'date') { const n = X.parseNumLoose(t); if (n != null) return String(n); const dd = X.parseDateLoose(t); if (dd) return valExpr(t); return t; }
      if (a.t === 'range') return t;
      // any / text
      if (/^[A-Za-zÇĞİÖŞÜçğıöşü_.]+\(.*\)$/.test(t) || /[<>=+*\/&^]/.test(t) && /[A-Z]+\d/.test(t)) return t;
      return valExpr(t);
    };
    const build = () => {
      const parts = fields.map(f => { const e = argExpr(Object.assign({}, f.a, { t: f.type || f.a.t }), f.inp.value); return e === '' ? '' : X.toStore(e); });
      while (parts.length && parts[parts.length - 1] === '') parts.pop();
      return fd.en + '(' + parts.join(',') + ')';
    };
    const self = this;
    function upd() {
      preview.innerHTML = '';
      let f; try { f = X.toStore(build()); X.parse(f); } catch (e) { preview.append(h('div', { class: 'pv-err', text: 'Girdileri kontrol edin.' })); return; }
      const missing = fields.find(x => !x.a.o && !x.a.many && x.inp.value.trim() === '' && x.type !== 'enum' && x.type !== 'bool');
      const v = missing ? null : self.app.evalFormula(f, self.target);
      preview.append(h('div', { class: 'pv-res' }, h('span', { text: 'Sonuç' }), h('b', { text: missing ? '—' : X.formatValue(v, fd.fmt || null) })));
      preview.append(h('details', { class: 'pv-f', open: true }, h('summary', { text: 'Formül' }), h('code', { text: '=' + X.toDisplay(f) })));
    }
    let fnClosing = false;
    d = U.dialog({ title: '<span class="fx">fx</span> ' + fd.tr, body, onClose: () => { if (!fnClosing && this.nested) { const n = this.nested; this.nested = null; n.done(); } }, buttons: [{ text: 'Geri', onClick: () => { fnClosing = true; const n = this.nested; d.close(); this.openFunctions(null, n); return false; } }, { text: 'Uygula', kind: 'primary', onClick: () => {
      let f; try { f = X.toStore(build()); X.parse(f); } catch (e) { U.toast('Girdileri kontrol edin', 'warn'); return false; }
      if (this.nested) { const n = this.nested; this.nested = null; n.onPick('=' + X.toDisplay(f)); fnClosing = true; d.close(); n.done(); return false; }
      const t = this.target; app.applyFormula(t.si, t.r, t.c, f, 1, fd.fmt); d.close(); U.toast('✓ ' + fd.tr + ' eklendi'); return false;
    } }] });
    upd();
  }
}
function critText(t) { const m = /^(<>|>=|<=|=|>|<)?(.*)$/.exec(t); const n = X.parseNumLoose(m[2]); return q((m[1] || '') + (n != null ? String(n) : m[2])); }
function balanced(s) { let d = 0, qt = false; for (const ch of s) { if (ch === '"') qt = !qt; if (qt) continue; if (ch === '(') d++; else if (ch === ')') { d--; if (d < 0) return false; } } return d === 0; }
function splitArgs(disp) {
  const i = disp.indexOf('('); const inner = disp.slice(i + 1, -1); const out = []; let d = 0, qt = false, cur = '';
  for (const ch of inner) { if (ch === '"') qt = !qt; if (!qt) { if (ch === '(') d++; else if (ch === ')') d--; else if (ch === ';' && d === 0) { out.push(cur); cur = ''; continue; } } cur += ch; }
  out.push(cur); return out.map(s => s.trim());
}
G.XW = { Wizards, RECIPES, isRef, valExpr, absRef };
})(window);
