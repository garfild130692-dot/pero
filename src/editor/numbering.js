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
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

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

// ---------- Маркеры и продолжение нумерации ----------
// Маркер маркированного списка — атрибут bulletList bullet (знак; null — по уровню: •, ○, ■, как у Word).
// Нумерованные списки верхнего уровня с одинаковым listId — один список Word (общий w:num): нумерация продолжается
// через абзацы между ними; номер начала на экране считает плагин listContinuation.
export const DEFAULT_BULLETS = ['•', '○', '■'];
export const BULLET_LIBRARY = ['•', '○', '■', '◆', '❖', '➢', '✓', '★', '–', '➤'];
// Знак ↔ lvlText и шрифт w:lvl (знаки шрифтов Symbol/Wingdings — из области частного использования, как пишет Word)
const BULLET_FONTS = { '•': ['\uF0B7', 'Symbol'], '○': ['o', 'Courier New'], '■': ['\uF0A7', 'Wingdings'], '❖': ['\uF076', 'Wingdings'], '➢': ['\uF0D8', 'Wingdings'], '✓': ['\uF0FC', 'Wingdings'], '–': ['–', 'Calibri'] };
export function bulletToLvl(ch) { return BULLET_FONTS[ch] || [ch, 'Segoe UI Symbol']; }
export function bulletFromLvl(text, font) {
  const t = String(text || '');
  for (const [ch, [lt, f]] of Object.entries(BULLET_FONTS)) if (t === lt && (!font || font.toLowerCase() === f.toLowerCase())) return ch;
  const code = t.codePointAt(0);
  if (code >= 0xf000 && code <= 0xf0ff) return '•'; // другой знак шрифта символов — показываем точкой
  return t || '•';
}
// Нумерация: готовые форматы (меню «Нумерованный список ▾»)
export const NUMBER_LIBRARY = [['decimal', '%1.'], ['decimal', '%1)'], ['upperRoman', '%1.'], ['upperLetter', '%1.'], ['lowerLetter', '%1)'],
  ['lowerLetter', '%1.'], ['lowerRoman', '%1.'], ['russianLower', '%1)'], ['russianUpper', '%1.']];

// Номер начала каждого нумерованного списка с listId (продолжение предыдущего)
export function listStarts(doc) {
  const counters = new Map();
  const out = [];
  doc.forEach((node, pos) => {
    if (node.type.name !== 'orderedList' || !node.attrs.listId) return;
    const id = node.attrs.listId;
    const start = counters.has(id) ? counters.get(id) : (node.attrs.start || 1);
    out.push({ pos, node, start, continued: counters.has(id) });
    counters.set(id, start + node.childCount);
  });
  return out;
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
        types: ['bulletList'],
        attributes: {
          bullet: {
            default: null, parseHTML: (el) => el.getAttribute('data-bullet'),
            renderHTML: (a) => (a.bullet ? { 'data-bullet': a.bullet, style: `--bullet: "${String(a.bullet).replace(/["\\]/g, '')}"` } : {}),
          },
        },
      },
      {
        types: ['orderedList'],
        attributes: {
          listId: { default: null, parseHTML: (el) => el.getAttribute('data-list'), renderHTML: (a) => (a.listId ? { 'data-list': a.listId } : {}) },
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
      // маркер маркированного списка под курсором (нет списка — создать)
      setBullet: (ch) => ({ state, chain }) => {
        const list = closestList(state, 'bulletList');
        if (!list) return chain().toggleBulletList().setBullet(ch).run();
        return chain().command(({ tr }) => { tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, bullet: ch && ch !== DEFAULT_BULLETS[list.depth % 3] ? ch : null }); return true; }).run();
      },
      // формат номера из библиотеки (нет нумерованного списка — создать)
      setNumberFormat: (numFmt, lvlText) => ({ state, chain }) => {
        const list = closestList(state, 'orderedList');
        if (!list) return chain().toggleOrderedList().setNumberFormat(numFmt, lvlText).run();
        const def = ['decimal', 'lowerLetter', 'lowerRoman'][list.depth % 3];
        return chain().setListFormat({ numFmt: numFmt === def ? null : numFmt, lvlText: lvlText === '%1.' ? null : lvlText }).run();
      },
      // «Начать заново с 1» / «Задать начальное значение» — новый список
      restartNumbering: (start = 1) => ({ state, tr, dispatch }) => {
        const list = closestList(state, 'orderedList');
        if (!list) return false;
        if (dispatch) dispatch(tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, listId: null, start: Math.max(0, start) }));
        return true;
      },
      // «Продолжить нумерацию» — продолжение ближайшего предыдущего нумерованного списка верхнего уровня
      continueNumbering: () => ({ state, tr, dispatch }) => {
        const list = closestList(state, 'orderedList', true);
        if (!list) return false;
        let prev = null;
        state.doc.forEach((n, pos) => { if (pos < list.pos && n.type.name === 'orderedList') prev = { n, pos }; });
        if (!prev) return false;
        let id = prev.n.attrs.listId;
        if (!id) {
          id = `P${Date.now().toString(36)}`;
          tr.setNodeMarkup(prev.pos, null, { ...prev.n.attrs, listId: id });
        }
        tr.setNodeMarkup(list.pos, null, { ...list.node.attrs, listId: id, start: null });
        if (dispatch) dispatch(tr);
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
          if (fmt.start != null) attrs.start = fmt.start;
          if (dispatch) dispatch(tr.setNodeMarkup($f.before(d), null, attrs));
          return true;
        }
        return false;
      },
    };
  },
  addProseMirrorPlugins() {
    // продолжение нумерации: у списка, продолжающего предыдущий, на экране — номер начала (атрибут start у <ol>)
    const key = new PluginKey('listContinuation');
    const build = (doc) => DecorationSet.create(doc, listStarts(doc).filter((x) => x.continued)
      .map((x) => Decoration.node(x.pos, x.pos + x.node.nodeSize, { start: String(x.start) })));
    return [new Plugin({
      key,
      state: { init: (_c, s) => build(s.doc), apply: (tr, v) => (tr.docChanged ? build(tr.doc) : v) },
      props: { decorations: (s) => key.getState(s) },
    })];
  },
});

// Ближайший список типа type вокруг курсора: { node, pos, depth (вложенность среди списков) }; top — только верхний уровень
function closestList(state, type, top = false) {
  const $f = state.selection.$from;
  let found = null;
  for (let d = $f.depth; d > 0; d--) {
    const n = $f.node(d);
    if (n.type.name !== type) continue;
    if (top && d !== 1) continue;
    let depth = 0;
    for (let k = d - 1; k > 0; k--) if (['bulletList', 'orderedList'].includes($f.node(k).type.name)) depth++;
    found = { node: n, pos: $f.before(d), depth };
    if (!top) break;
  }
  return found;
}
