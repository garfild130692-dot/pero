// Правила условного форматирования из меню (CfMenu.jsx): готовые оформления, гистограммы, шкалы, периоды дат
// и сборка правила по ответу диалога — как их записывает Excel (с формулами для «текст содержит» и дат).
import { addr } from './address.js';
import { fromUser } from './formula.js';
import { parseInput } from './format.js';

// Готовые оформления (как в Excel по умолчанию)
export const CF_PRESETS = [
  ['Светло-красная заливка и тёмно-красный текст', { fill: '#ffc7ce', color: '#9c0006' }],
  ['Жёлтая заливка и тёмно-жёлтый текст', { fill: '#ffeb9c', color: '#9c5700' }],
  ['Зелёная заливка и тёмно-зелёный текст', { fill: '#c6efce', color: '#006100' }],
  ['Светло-красная заливка', { fill: '#ffc7ce' }],
  ['Красный текст', { color: '#9c0006' }],
  ['Красная граница', { border: Object.fromEntries(['left', 'right', 'top', 'bottom'].map((s) => [s, { style: 'thin', color: '#c00000' }])) }],
];
export const BARS = [['Синяя', '#638ec6'], ['Зелёная', '#63c384'], ['Красная', '#ff555a'], ['Оранжевая', '#ffb628'], ['Голубая', '#008aef'], ['Фиолетовая', '#d6007b']];
export const SCALES = [
  ['Зелёный — жёлтый — красный', ['#63be7b', '#ffeb84', '#f8696b']], ['Красный — жёлтый — зелёный', ['#f8696b', '#ffeb84', '#63be7b']],
  ['Зелёный — белый — красный', ['#63be7b', '#fcfcff', '#f8696b']], ['Красный — белый — зелёный', ['#f8696b', '#fcfcff', '#63be7b']],
  ['Синий — белый — красный', ['#5a8ac6', '#fcfcff', '#f8696b']], ['Красный — белый — синий', ['#f8696b', '#fcfcff', '#5a8ac6']],
  ['Белый — зелёный', ['#fcfcff', '#63be7b']], ['Зелёный — белый', ['#63be7b', '#fcfcff']], ['Белый — красный', ['#fcfcff', '#f8696b']],
  ['Красный — белый', ['#f8696b', '#fcfcff']], ['Зелёный — жёлтый', ['#63be7b', '#ffef9c']], ['Жёлтый — зелёный', ['#ffef9c', '#63be7b']],
];
// Периоды дат и формулы, которые пишет для них Excel (A1 — левая верхняя ячейка диапазона)
export const PERIODS = [
  ['yesterday', 'Вчера', (a) => `FLOOR(${a},1)=TODAY()-1`], ['today', 'Сегодня', (a) => `FLOOR(${a},1)=TODAY()`],
  ['tomorrow', 'Завтра', (a) => `FLOOR(${a},1)=TODAY()+1`],
  ['last7Days', 'За последние 7 дней', (a) => `AND(TODAY()-FLOOR(${a},1)<=6,FLOOR(${a},1)<=TODAY())`],
  ['lastWeek', 'На прошлой неделе', (a) => `AND(TODAY()-ROUNDDOWN(${a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a},0)<(WEEKDAY(TODAY())+7))`],
  ['thisWeek', 'На этой неделе', (a) => `AND(TODAY()-ROUNDDOWN(${a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a},0)-TODAY()<=7-WEEKDAY(TODAY()))`],
  ['nextWeek', 'На следующей неделе', (a) => `AND(ROUNDDOWN(${a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a},0)-TODAY()<(15-WEEKDAY(TODAY())))`],
  ['lastMonth', 'В прошлом месяце', (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a})=YEAR(EDATE(TODAY(),0-1)))`],
  ['thisMonth', 'В этом месяце', (a) => `AND(MONTH(${a})=MONTH(TODAY()),YEAR(${a})=YEAR(TODAY()))`],
  ['nextMonth', 'В следующем месяце', (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a})=YEAR(EDATE(TODAY(),0+1)))`],
];
const quote = (s) => `"${String(s).replace(/"/g, '""')}"`;
// значение условия → каноническая формула: «=…» — формула, число/дата — число, иначе текст в кавычках
function valueFormula(text) {
  const t = String(text).trim();
  if (t.startsWith('=')) return fromUser(t.slice(1));
  const p = parseInput(t);
  return typeof p.value === 'number' ? String(p.value) : quote(t);
}

// правило из ответа диалога; rg — диапазон (для относительных формул — его левая верхняя ячейка)
export function buildRule(kind, values, dxf, rg, extra) {
  const a = addr(rg.r1, rg.c1);
  switch (kind) {
    case 'greaterThan': case 'lessThan': case 'equal': return { type: 'cellIs', operator: kind, formulas: [valueFormula(values[0])], dxf };
    case 'between': return { type: 'cellIs', operator: 'between', formulas: [valueFormula(values[0]), valueFormula(values[1])], dxf };
    case 'containsText': return { type: 'containsText', operator: 'containsText', text: values[0], formulas: [`NOT(ISERROR(SEARCH(${quote(values[0])},${a})))`], dxf };
    case 'timePeriod': { const p = PERIODS.find(([k]) => k === extra.period); return { type: 'timePeriod', timePeriod: p[0], formulas: [p[2](a)], dxf }; }
    case 'duplicateValues': return { type: extra.unique ? 'uniqueValues' : 'duplicateValues', dxf };
    case 'top': case 'topPercent': case 'bottom': case 'bottomPercent':
      return { type: 'top10', rank: Math.max(1, Math.round(+String(values[0]).replace(',', '.') || 10)), percent: /Percent/.test(kind), bottom: kind.startsWith('bottom'), dxf };
    case 'above': case 'below': return { type: 'aboveAverage', aboveAverage: kind === 'above', dxf };
    default: return null;
  }
}

