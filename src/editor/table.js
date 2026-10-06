// Таблицы: модель границ и общие преобразования — для редактора, импорта и экспорта.
//
// Граница: { style: 'single'|'double'|'dotted'|'dashed'|'none', width: pt, color: '#rrggbb'|null (авто) }.
// Таблица: borders = { top, left, bottom, right, insideH, insideV } — границы по умолчанию для всех ячеек.
// Ячейка: borders = { top?, right?, bottom?, left? } — только переопределённые стороны (null — как у таблицы).

export const NONE = Object.freeze({ style: 'none' });
export const DEFAULT_PEN = Object.freeze({ style: 'single', width: 0.5, color: null });
export const TABLE_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
export const CELL_SIDES = ['top', 'left', 'bottom', 'right'];

// Новая таблица — «Сетка таблицы» (TableGrid): тонкие одинарные линии везде
export const GRID_BORDERS = Object.freeze(Object.fromEntries(TABLE_SIDES.map((s) => [s, DEFAULT_PEN])));
// Таблица без стиля и без границ в файле — как TableNormal в Word: границ нет
export const NO_BORDERS = Object.freeze(Object.fromEntries(TABLE_SIDES.map((s) => [s, NONE])));

export const BORDER_STYLES = [
  ['single', 'Сплошная'], ['dashed', 'Штриховая'], ['dotted', 'Точечная'], ['double', 'Двойная'],
];
export const BORDER_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

// ---------- OOXML ----------
// Ширина линии в OOXML — в восьмых долях пункта (w:sz), допустимо 2…96
export const szToPt = (sz) => Math.min(Math.max(+sz || 4, 2), 96) / 8;
export const ptToSz = (pt) => Math.min(Math.max(Math.round(pt * 8), 2), 96);

const OOXML_STYLE = {
  single: 'single', thick: 'single', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed',
  dotDash: 'dashed', dotDotDash: 'dotted', triple: 'double', thinThickSmallGap: 'double', thickThinSmallGap: 'double',
  thinThickMediumGap: 'double', thickThinMediumGap: 'double', thinThickLargeGap: 'double', thickThinLargeGap: 'double',
};

// Атрибуты элемента границы (val, sz, color) → граница модели
export function borderFromOoxml(val, sz, color) {
  if (!val || ['nil', 'none'].includes(val)) return NONE;
  const c = color && color.toLowerCase() !== 'auto' && /^[0-9a-f]{6}$/i.test(color) ? '#' + color.toLowerCase() : null;
  return { style: OOXML_STYLE[val] || 'single', width: szToPt(sz), color: c };
}

// Граница модели → XML-элемент <w:{name} …/>
export function borderToOoxml(name, b) {
  if (!b || b.style === 'none') return `<w:${name} w:val="nil"/>`;
  const color = b.color ? b.color.slice(1).toUpperCase() : 'auto';
  return `<w:${name} w:val="${b.style}" w:sz="${ptToSz(b.width || 0.5)}" w:space="0" w:color="${color}"/>`;
}

// ---------- CSS ----------
// Граница «нет» на экране — бледная пунктирная сетка (как «Отображать сетку» в Word); в файл не попадает
export function borderCss(b) {
  if (!b || b.style === 'none') return '0.5px dashed var(--gridline, #c9d3df)';
  const color = b.color || '#000';
  const w = b.width || 0.5;
  if (b.style === 'double') return `max(${w * 3}pt, 3px) double ${color}`;
  const css = { single: 'solid', dashed: 'dashed', dotted: 'dotted' }[b.style] || 'solid';
  return `${w}pt ${css} ${color}`;
}

export const sameBorder = (a, b) => JSON.stringify(a || null) === JSON.stringify(b || null);

// Ширина столбцов: в OOXML — твипы, у TipTap — px (1 px = 15 твипов при 96 dpi)
export const TWIPS_PER_PX = 15;
