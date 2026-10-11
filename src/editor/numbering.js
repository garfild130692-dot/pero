// Нумерация заголовков (1, 1.1, 1.1.1 — многоуровневый список, связанный с «Заголовок 1–9») и формат нумерованных
// списков — как в Word.
//
// Заголовки: doc.attrs.headingNumbering = null | {
//   levels: [{ fmt, text, start, on }] × 9 — формат номера уровня (fmt: decimal, lowerLetter…; text: '%1.%2.'), on — уровень
//           нумеруется (в файле Word нумерация может быть только у части стилей заголовков);
//   suff: 'tab' | 'space' | 'nothing'; source: { numId, viaStyle } — нумерация из файла; changed — задана/изменена в Пере;
//   off — нумерация из стилей файла отключена в Пере (при сохранении у заголовков numId 0) }.
// Заголовок без номера — атрибут noNumber (w:numId 0 или не нумеруемый стиль). На экране — CSS-счётчики (headingNumberCss).
// Нумерованный список: атрибуты orderedList numFmt (null — по уровню: 1, a, i) и lvlText ('%1.' по умолчанию; '%1)',
// '(%1)', '%1.%2.' — с номерами родителей).
import { Extension } from '@tiptap/core';

export const NUM_FORMATS = [
  ['decimal', '1, 2, 3'], ['decimalZero', '01, 02, 03'], ['lowerLetter', 'a, b, c'], ['upperLetter', 'A, B, C'],
  ['lowerRoman', 'i, ii, iii'], ['upperRoman', 'I, II, III'], ['russianLower', 'а, б, в'], ['russianUpper', 'А, Б, В'],
];
const CSS_STYLE = {
  decimal: 'decimal', decimalZero: 'decimal-leading-zero', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha',
  lowerRoman: 'lower-roman', upperRoman: 'upper-roman', russianLower: 'pero-ru-lower', russianUpper: 'pero-ru-upper',
};
export const cssCounterStyle = (fmt) => CSS_STYLE[fmt] || 'decimal';

// Русские буквы нумерации Word (без ё, й, ъ, ы, ь)
const RU = 'абвгдежзиклмнопрстуфхцчшщэюя';
export const COUNTER_STYLES_CSS =
  `@counter-style pero-ru-lower { system: alphabetic; symbols: ${[...RU].join(' ')}; }\n` +
  `@counter-style pero-ru-upper { system: alphabetic; symbols: ${[...RU.toUpperCase()].join(' ')}; }`;

const lvls = (texts) => Array.from({ length: 9 }, (_, i) => ({ fmt: 'decimal', text: texts(i), start: 1, on: true }));
const path = (i, sep = '.') => Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(sep);

// Готовые схемы нумерации заголовков (меню «Многоуровневый список»)
export const HEADING_SCHEMES = [
  { key: 'plain', label: '1 / 1.1 / 1.1.1', make: () => lvls((i) => path(i)) },
  { key: 'dots', label: '1. / 1.1. / 1.1.1.', make: () => lvls((i) => `${path(i)}.`) },
  { key: 'article', label: 'Статья 1. / 1.1. / 1.1.1.', make: () => lvls((i) => (i === 0 ? 'Статья %1.' : `${path(i)}.`)) },
  { key: 'chapter', label: 'Глава 1 / 1.1 / 1.1.1', make: () => lvls((i) => (i === 0 ? 'Глава %1' : path(i))) },
  { key: 'section', label: 'Раздел I / 1. / 1.1.', make: () => lvls((i) => (i === 0 ? 'Раздел %1' : `${Array.from({ length: i }, (_, k) => `%${k + 2}`).join('.')}.`)).map((l, i) => (i === 0 ? { ...l, fmt: 'upperRoman' } : l)) },
];

export function makeHeadingNumbering(key, prev = null) {
  const s = HEADING_SCHEMES.find((x) => x.key === key) || HEADING_SCHEMES[0];
  return { levels: s.make(), suff: 'tab', source: prev ? prev.source || null : null, changed: true };
}

// Номер уровня как его покажет Word: text с подставленными номерами (для подписей и предпросмотра)
export function fmtNumber(n, fmt) {
  if (fmt === 'decimalZero') return String(n).padStart(2, '0');
  if (fmt === 'lowerLetter' || fmt === 'upperLetter') {
    const s = String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);
    return fmt === 'upperLetter' ? s.toUpperCase() : s;
  }
  if (fmt === 'russianLower' || fmt === 'russianUpper') {
    const s = RU[(n - 1) % RU.length].repeat(Math.floor((n - 1) / RU.length) + 1);
    return fmt === 'russianUpper' ? s.toUpperCase() : s;
  }
  if (fmt === 'lowerRoman' || fmt === 'upperRoman') {
    const R = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
    let s = ''; let x = n;
    for (const [v, r] of R) while (x >= v) { s += r; x -= v; }
    return fmt === 'upperRoman' ? s.toUpperCase() : s;
  }
  return String(n);
}
export function levelLabel(levels, i, nums) {
  return String(levels[i].text || '').replace(/%(\d)/g, (_, k) => fmtNumber(nums[+k - 1] ?? levels[+k - 1]?.start ?? 1, (levels[+k - 1] || {}).fmt));
}

const cssStr = (s) => `"${String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
// CSS content для шаблона уровня: литералы и counter(pero-hN, стиль)
function contentOf(text, levels) {
  const parts = [];
  let last = 0;
  for (const m of String(text || '').matchAll(/%(\d)/g)) {
    if (m.index > last) parts.push(cssStr(text.slice(last, m.index)));
    const k = +m[1];
    parts.push(`counter(pero-h${k}, ${cssCounterStyle((levels[k - 1] || {}).fmt)})`);
    last = m.index + m[0].length;
  }
  if (last < String(text || '').length) parts.push(cssStr(text.slice(last)));
  return parts.length ? parts.join(' ') : '""';
}

// CSS нумерации заголовков (h1–h6 на экране; уровни 7–9 у нас нет)
export function headingNumberCss(hn) {
  if (!hn || hn.off || !hn.levels) return '';
  const L = hn.levels;
  const sel = (i) => `.page-content h${i + 1}:not([data-nonum])`;
  const resets = L.slice(0, 6).map((l, i) => `pero-h${i + 1} ${(l.start || 1) - 1}`).join(' ');
  const out = [`.page-content { counter-reset: pero-fn pero-en pero-cm ${resets}; }`];
  for (let i = 0; i < 6; i++) {
    const deeper = L.slice(i + 1, 6).map((l, j) => `pero-h${i + j + 2} ${(l.start || 1) - 1}`).join(' ');
    out.push(`${sel(i)} { counter-increment: pero-h${i + 1};${deeper ? ` counter-set: ${deeper};` : ''} }`);
    if (!L[i].on) continue;
    const gap = hn.suff === 'space' ? 'margin-right: 0.3em;' : hn.suff === 'nothing' ? '' : `min-width: ${21.6 + 7.2 * i}pt; padding-right: 0.5em; box-sizing: border-box;`;
    out.push(`${sel(i)}::before { content: ${contentOf(L[i].text, L)}; display: inline-block; ${gap} text-indent: 0; }`);
  }
  return out.join('\n');
}

// Вид маркера нумерованного списка по шаблону уровня
export function markerKind(text, ilvl = 0) {
  const t = String(text || '');
  if (!t || t === `%${ilvl + 1}.`) return null;
  if ((t.match(/%\d/g) || []).length > 1) return /\.$/.test(t) ? 'path-dot' : 'path';
  if (/^\(%\d\)$/.test(t)) return 'both';
  if (/^%\d\)$/.test(t)) return 'paren';
  if (/^%\d$/.test(t)) return 'plain';
  return null;
}

export const LIST_PATTERNS = [['%1.', '1.'], ['%1)', '1)'], ['(%1)', '(1)'], ['%1', '1'], ['%1.%2.', '1.1. (с номером родителя)']];

export const Numbering = Extension.create({
  name: 'numbering',
  addGlobalAttributes() {
    return [
      {
        types: ['heading'],
        attributes: {
          noNumber: { default: null, parseHTML: (el) => (el.hasAttribute('data-nonum') ? true : null), renderHTML: (a) => (a.noNumber ? { 'data-nonum': '' } : {}) },
        },
      },
      {
        types: ['orderedList'],
        attributes: {
          numFmt: { default: null, parseHTML: (el) => el.getAttribute('data-fmt'), renderHTML: (a) => (a.numFmt ? { 'data-fmt': a.numFmt, style: `--lst: ${cssCounterStyle(a.numFmt)}` } : {}) },
          lvlText: { default: null, parseHTML: (el) => el.getAttribute('data-lvl'), renderHTML: (a) => (a.lvlText ? { 'data-lvl': a.lvlText, 'data-mk': markerKind(a.lvlText) || null } : {}) },
        },
      },
    ];
  },
  addCommands() {
    return {
      // схема нумерации заголовков (key из HEADING_SCHEMES), null — без нумерации; levels — своя (диалог)
      setHeadingNumbering: (key, levels = null) => ({ state, tr, dispatch }) => {
        const prev = state.doc.attrs.headingNumbering;
        let next;
        if (levels) next = { ...(prev && !prev.off ? prev : {}), levels, suff: (prev && prev.suff) || 'tab', source: prev ? prev.source || null : null, changed: true, off: false };
        else if (key) next = makeHeadingNumbering(key, prev);
        else next = prev && prev.source && prev.source.viaStyle ? { ...prev, off: true } : null;
        if (dispatch) dispatch(tr.setDocAttribute('headingNumbering', next));
        return true;
      },
      // заголовок под курсором — без номера / с номером
      toggleHeadingNumber: () => ({ state, tr, dispatch }) => {
        const $f = state.selection.$from;
        const node = $f.parent;
        if (node.type.name !== 'heading') return false;
        if (dispatch) dispatch(tr.setNodeMarkup($f.before(), null, { ...node.attrs, noNumber: node.attrs.noNumber ? null : true }));
        return true;
      },
      // формат нумерованного списка под курсором: { numFmt, lvlText, start }
      setListFormat: (fmt) => ({ state, tr, dispatch }) => {
        const $f = state.selection.$from;
        for (let d = $f.depth; d > 0; d--) {
          const n = $f.node(d);
          if (n.type.name !== 'orderedList') continue;
          const attrs = { ...n.attrs };
          if ('numFmt' in fmt) attrs.numFmt = fmt.numFmt || null;
          if ('lvlText' in fmt) attrs.lvlText = fmt.lvlText || null;
          if (fmt.start) attrs.start = fmt.start;
          if (dispatch) dispatch(tr.setNodeMarkup($f.before(d), null, attrs));
          return true;
        }
        return false;
      },
    };
  },
});
