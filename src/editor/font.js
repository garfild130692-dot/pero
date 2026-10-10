// Шрифт: общие константы и нормализация значений — одни и те же для редактора, импорта и экспорта.

// Шрифт и размер нового документа (как в русском Word по умолчанию)
export const DEFAULT_FONT = 'Calibri';
export const DEFAULT_SIZE = 11;

// Размеры в выпадающем списке и шаги кнопок «Увеличить/уменьшить размер» (как в Word)
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export const MIN_SIZE = 1;
export const MAX_SIZE = 1638;

// Распространённые шрифты Windows — запасной список, если система не отдала перечень установленных
export const COMMON_FONTS = [
  'Arial', 'Arial Narrow', 'Bahnschrift', 'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Century Gothic',
  'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond',
  'Georgia', 'Impact', 'Lucida Console', 'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman',
  'Trebuchet MS', 'Verdana',
];

// Цвета выделения текста (w:highlight): фиксированный набор из 16 значений OOXML
export const HIGHLIGHTS = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};
export const highlightName = (hex) => Object.keys(HIGHLIGHTS).find((k) => HIGHLIGHTS[k] === hex) || null;

// Любой CSS-цвет (#rgb, #rrggbb, rgb(), часть имён) → '#rrggbb' в нижнем регистре; иначе null
const NAMED = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00' };
export function normalizeColor(value) {
  if (!value) return null;
  const v = String(value).trim().toLowerCase();
  if (NAMED[v]) return NAMED[v];
  let m = /^#?([0-9a-f]{6})$/.exec(v);
  if (m) return '#' + m[1];
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
  if (m) return '#' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3];
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(v);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null; // прозрачный
    return '#' + [m[1], m[2], m[3]].map((n) => Math.min(255, +n).toString(16).padStart(2, '0')).join('');
  }
  return null;
}

// Первое семейство из CSS font-family без кавычек: '"Times New Roman", serif' → 'Times New Roman'
export function normalizeFontFamily(value) {
  if (!value) return null;
  const first = String(value).split(',')[0].trim().replace(/^['"]|['"]$/g, '').trim();
  return first || null;
}

// Стек для CSS: нужный шрифт + метрически совместимые замены + общий фолбэк
const SUBSTITUTES = {
  Calibri: ['Carlito'], 'Calibri Light': ['Calibri', 'Carlito'], Cambria: ['Caladea'],
  Arial: ['Liberation Sans'], 'Times New Roman': ['Liberation Serif'], 'Courier New': ['Liberation Mono'],
};
const MONO = /mono|courier|consolas|console/i;
const SERIF = /times|georgia|cambria|garamond|palatino|constantia|serif|book/i;
export function cssFontStack(name) {
  if (!name) return null;
  const q = (s) => `"${s.replace(/["\\]/g, '')}"`;
  const generic = MONO.test(name) ? 'monospace' : SERIF.test(name) && !/sans/i.test(name) ? 'serif' : 'sans-serif';
  return [name, ...(SUBSTITUTES[name] || [])].map(q).concat(generic).join(', ');
}

// Размер в пунктах: '14pt' / '18.67px' / 14 → 14; округление до 0,5 pt (как хранит OOXML в полупунктах)
export function parseSize(value) {
  if (value === null || value === undefined || value === '') return null;
  const m = /^\s*([\d.,]+)\s*(pt|px)?\s*$/i.exec(String(value));
  if (!m) return null;
  let n = parseFloat(m[1].replace(',', '.'));
  if (m[2] && m[2].toLowerCase() === 'px') n *= 0.75;
  if (!isFinite(n)) return null;
  n = Math.round(n * 2) / 2;
  return n >= MIN_SIZE && n <= MAX_SIZE ? n : null;
}

// Следующий/предыдущий размер по шкале FONT_SIZES; за её пределами — шаг 10 pt, как в Word
export function stepSize(size, dir) {
  if (dir > 0) {
    const next = FONT_SIZES.find((s) => s > size);
    return Math.min(next ?? Math.floor(size / 10) * 10 + 10, MAX_SIZE);
  }
  const last = FONT_SIZES[FONT_SIZES.length - 1];
  if (size > last) return Math.max(Math.ceil(size / 10) * 10 - 10, last);
  const prev = [...FONT_SIZES].reverse().find((s) => s < size);
  return prev ?? Math.max(Math.ceil(size) - 1, MIN_SIZE);
}

// Как оформлены заголовки по умолчанию (совпадает со styles.css и стилями Heading1–6 в export.js).
// Оформление заголовка из файла, отличное от этого, переносится в метки текста (см. import.js).
export const HEADING_FONT = 'Calibri Light';
export const HEADING_SIZES = { 1: 16, 2: 13, 3: 12, 4: 11, 5: 11, 6: 11 };
export const HEADING_COLORS = { 1: '#2f5496', 2: '#2f5496', 3: '#1f3763', 4: '#2f5496', 5: '#2f5496', 6: '#2f5496' };
export const headingDefaults = (level) => ({ fontFamily: HEADING_FONT, fontSize: HEADING_SIZES[level] || 11, color: HEADING_COLORS[level] || '#2f5496' });
