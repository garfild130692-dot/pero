// Найти и заменить в книге (как в Excel): на листе или во всей книге, по строкам или по столбцам, в формулах
// (текст ввода: «=СУММ(A1:A3)», «1234,5») или в значениях (как на экране), с учётом регистра, ячейка целиком,
// подстановочные знаки * ? (~ — сам знак). Замена всегда — в тексте ввода (как Excel).
import { key, unkey } from './address.js';
import { inputOf, setInputs } from './model.js';
import { shownText } from './filter.js';

// шаблон Excel → RegExp (g — для замены всех вхождений в ячейке)
export function findRegExp(what, { matchCase = false, whole = false } = {}, flags = '') {
  let re = '';
  const s = String(what);
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '~' && i + 1 < s.length && '*?~'.includes(s[i + 1])) { re += `\\${s[++i]}`; continue; }
    re += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(whole ? `^(?:${re})$` : re, (matchCase ? '' : 'i') + flags);
}
const textOf = (cell, lookIn) => (lookIn === 'values' ? shownText(cell) : inputOf(cell));

// Все совпадения: [{ si, r, c, text }] в порядке обхода (листы по порядку; по строкам или по столбцам)
export function findAll(wb, what, opts = {}) {
  const { within = 'sheet', si = wb.active, by = 'rows', lookIn = 'formulas' } = opts;
  if (!String(what)) return [];
  const re = findRegExp(what, opts);
  const out = [];
  const sheets = within === 'book' ? wb.sheets.map((_, i) => i) : [si];
  for (const i of sheets) {
    const sh = wb.sheets[i];
    const hits = [];
    for (const [k, cell] of sh.cells) {
      if (!cell || (cell.v == null && cell.f == null)) continue;
      const text = textOf(cell, lookIn);
      if (text && re.test(text)) { const [r, c] = unkey(k); hits.push({ si: i, r, c, text }); }
    }
    hits.sort((a, b) => (by === 'cols' ? a.c - b.c || a.r - b.r : a.r - b.r || a.c - b.c));
    out.push(...hits);
  }
  return out;
}
// Следующее совпадение после ячейки (si, r, c) с переходом по кругу; back — предыдущее
export function findNext(wb, what, from, opts = {}) {
  const list = findAll(wb, what, opts);
  if (!list.length) return null;
  const by = opts.by || 'rows';
  const order = (x) => [x.si, by === 'cols' ? x.c : x.r, by === 'cols' ? x.r : x.c];
  const cmp = (a, b) => { const p = order(a); const q = order(b); for (let i = 0; i < 3; i++) if (p[i] !== q[i]) return p[i] - q[i]; return 0; };
  if (opts.back) { const prev = [...list].reverse().find((x) => cmp(x, from) < 0); return prev || list[list.length - 1]; }
  return list.find((x) => cmp(x, from) > 0) || list[0];
}
// Заменить в одной ячейке (текст ввода); null — нечего менять
export function replacedInput(cell, what, by, opts = {}) {
  const text = inputOf(cell);
  const reAll = findRegExp(what, opts, 'g');
  if (!findRegExp(what, opts).test(text)) return null;
  if (opts.whole) return String(by);
  return text.replace(reAll, () => String(by)); // замена — буквально (без $1 и т. п.)
}
// Заменить все: { wb, count }; ячейки, текст которых после замены не изменился, не трогаются
export function replaceAll(wb, what, by, opts = {}) {
  const list = findAll(wb, what, { ...opts, lookIn: 'formulas' });
  const items = [];
  for (const x of list) {
    const cell = wb.sheets[x.si].cells.get(key(x.r, x.c));
    const raw = replacedInput(cell, what, by, opts);
    if (raw != null && raw !== inputOf(cell)) items.push({ si: x.si, r: x.r, c: x.c, raw });
  }
  return { wb: items.length ? setInputs(wb, items) : wb, count: items.length };
}
