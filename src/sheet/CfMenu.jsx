// Меню «Условное форматирование» (Главная → Стили) и диалог условия — раскладка как в Excel, свои цвета и подписи.
import React, { useState } from 'react';
import { usePopover } from '../FontControls.jsx';
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
const BARS = [['Синяя', '#638ec6'], ['Зелёная', '#63c384'], ['Красная', '#ff555a'], ['Оранжевая', '#ffb628'], ['Голубая', '#008aef'], ['Фиолетовая', '#d6007b']];
const SCALES = [
  ['Зелёный — жёлтый — красный', ['#63be7b', '#ffeb84', '#f8696b']], ['Красный — жёлтый — зелёный', ['#f8696b', '#ffeb84', '#63be7b']],
  ['Зелёный — белый — красный', ['#63be7b', '#fcfcff', '#f8696b']], ['Красный — белый — зелёный', ['#f8696b', '#fcfcff', '#63be7b']],
  ['Синий — белый — красный', ['#5a8ac6', '#fcfcff', '#f8696b']], ['Красный — белый — синий', ['#f8696b', '#fcfcff', '#5a8ac6']],
  ['Белый — зелёный', ['#fcfcff', '#63be7b']], ['Зелёный — белый', ['#63be7b', '#fcfcff']], ['Белый — красный', ['#fcfcff', '#f8696b']],
  ['Красный — белый', ['#f8696b', '#fcfcff']], ['Зелёный — жёлтый', ['#63be7b', '#ffef9c']], ['Жёлтый — зелёный', ['#ffef9c', '#63be7b']],
];
// Периоды дат и формулы, которые пишет для них Excel (A1 — левая верхняя ячейка диапазона)
const PERIODS = [
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

// Виды диалога: заголовок, подпись, поля, как собрать правило
const KINDS = {
  greaterThan: { title: 'Больше', label: 'Форматировать ячейки, которые БОЛЬШЕ:', fields: 1 },
  lessThan: { title: 'Меньше', label: 'Форматировать ячейки, которые МЕНЬШЕ:', fields: 1 },
  between: { title: 'Между', label: 'Форматировать ячейки, значения которых находятся МЕЖДУ:', fields: 2 },
  equal: { title: 'Равно', label: 'Форматировать ячейки, которые РАВНЫ:', fields: 1 },
  containsText: { title: 'Текст, который содержит', label: 'Форматировать ячейки, которые содержат текст:', fields: 1 },
  timePeriod: { title: 'Дата', label: 'Форматировать ячейки, которые содержат дату:', fields: 0, period: true },
  duplicateValues: { title: 'Повторяющиеся значения', label: 'Форматировать ячейки, которые содержат:', fields: 0, dupe: true },
  top: { title: 'Первые элементы', label: 'Форматировать первые ячейки:', fields: 1, rank: true },
  topPercent: { title: 'Первые проценты', label: 'Форматировать первые ячейки (%):', fields: 1, rank: true },
  bottom: { title: 'Последние элементы', label: 'Форматировать последние ячейки:', fields: 1, rank: true },
  bottomPercent: { title: 'Последние проценты', label: 'Форматировать последние ячейки (%):', fields: 1, rank: true },
  above: { title: 'Выше среднего', label: 'Форматировать ячейки, которые ВЫШЕ СРЕДНЕГО для выделенного диапазона', fields: 0 },
  below: { title: 'Ниже среднего', label: 'Форматировать ячейки, которые НИЖЕ СРЕДНЕГО для выделенного диапазона', fields: 0 },
};
// правило из ответа диалога; rg — диапазон (для относительных формул — его левая верхняя ячейка)
function buildRule(kind, values, dxf, rg, extra) {
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

export function CfDialog({ kind, rg, onApply, onClose }) {
  const k = KINDS[kind];
  const [values, setValues] = useState(k.rank ? [kind.includes('Percent') ? '10' : '10', ''] : ['', '']);
  const [preset, setPreset] = useState(0);
  const [period, setPeriod] = useState('today');
  const [unique, setUnique] = useState(false);
  const ok = () => {
    if (k.fields && !String(values[0]).trim()) return;
    if (k.fields === 2 && !String(values[1]).trim()) return;
    const rule = buildRule(kind, values, CF_PRESETS[preset][1], rg, { period, unique });
    if (rule) onApply(rule);
  };
  const set = (i, v) => setValues(values.map((x, j) => (j === i ? v : x)));
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal cf-rule-dlg" role="dialog" aria-label={k.title} onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') ok(); }}>
        <div className="modal-title">{k.title}</div>
        <div className="modal-body">
          <div className="cf-head">{k.label}</div>
          <div className="cf-row">
            {k.dupe && (
              <select value={unique ? '1' : '0'} onChange={(e) => setUnique(e.target.value === '1')}>
                <option value="0">повторяющиеся</option><option value="1">уникальные</option>
              </select>
            )}
            {k.period && <select value={period} onChange={(e) => setPeriod(e.target.value)}>{PERIODS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select>}
            {k.fields >= 1 && <input className="pd-input" autoFocus value={values[0]} onChange={(e) => set(0, e.target.value)} style={{ width: k.rank ? 70 : 150 }} />}
            {k.fields === 2 && <><span className="cf-and">и</span><input className="pd-input" value={values[1]} onChange={(e) => set(1, e.target.value)} style={{ width: 150 }} /></>}
            <span className="cf-with">с</span>
            <select value={preset} onChange={(e) => setPreset(+e.target.value)}>
              {CF_PRESETS.map(([l], i) => <option key={l} value={i}>{l}</option>)}
            </select>
          </div>
          {k.fields > 0 && !k.rank && <div className="cf-hint">Можно ввести число, текст или ссылку на ячейку со знаком «=» (например, =$B$1)</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// Кнопка и выпадающее меню с подменю
export default function CfMenu({ onDialog, onRule, onClear }) {
  // меню — поверх окна (position: fixed), иначе его обрезает лента
  const pop = usePopover();
  const { open, ref } = pop;
  const setOpen = (v) => { pop.setOpen(v); if (!v) setSub(null); };
  const [sub, setSub] = useState(null);
  const pick = (fn) => () => { setOpen(false); setSub(null); fn(); };
  const subMenu = (id, label, items) => (
    <div className="fm-sub" onMouseEnter={() => setSub(id)}>
      <button type="button" className="menu-item" onClick={() => setSub(sub === id ? null : id)}>{label}<span className="fm-arrow">▸</span></button>
      {sub === id && <div className="filter-menu fm-subpop cf-subpop">{items}</div>}
    </div>
  );
  const item = (label, fn, extra = null) => <button key={label} type="button" className="menu-item" onClick={pick(fn)}>{extra}{label}</button>;
  return (
    <div className="rb-popover-anchor" ref={ref}>
      <button type="button" className={`rb-btn big${open ? ' active' : ''}`} title="Условное форматирование: выделить ячейки по правилам" onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6"><rect x="3.5" y="4" width="17" height="16" rx="1" /><path d="M3.5 9.5h17M3.5 15h17M10 4v16" /><rect x="11" y="10.5" width="8.5" height="3.5" fill="#f8696b" stroke="none" /><rect x="11" y="16" width="5" height="3" fill="#63be7b" stroke="none" /></svg>
        <span className="rb-label">Условное<br />форматирование<span className="dd">▾</span></span>
      </button>
      {open && (
        <div className="filter-menu cf-menu" style={pop.style} role="menu">
          {subMenu('hl', 'Правила выделения ячеек', [
            item('Больше…', () => onDialog('greaterThan')), item('Меньше…', () => onDialog('lessThan')), item('Между…', () => onDialog('between')),
            item('Равно…', () => onDialog('equal')), item('Текст содержит…', () => onDialog('containsText')), item('Дата…', () => onDialog('timePeriod')),
            item('Повторяющиеся значения…', () => onDialog('duplicateValues')),
          ])}
          {subMenu('top', 'Правила отбора первых и последних значений', [
            item('Первые 10 элементов…', () => onDialog('top')), item('Первые 10%…', () => onDialog('topPercent')),
            item('Последние 10 элементов…', () => onDialog('bottom')), item('Последние 10%…', () => onDialog('bottomPercent')),
            item('Выше среднего…', () => onDialog('above')), item('Ниже среднего…', () => onDialog('below')),
          ])}
          <div className="menu-sep" />
          {subMenu('bars', 'Гистограммы', BARS.map(([l, c]) => item(l, () => onRule({ type: 'dataBar', bar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: c } }),
            <span className="cf-swatch" style={{ background: `linear-gradient(90deg, ${c}, ${c}22)` }} />)))}
          {subMenu('scales', 'Цветовые шкалы', SCALES.map(([l, cs]) => item(l, () => onRule({
            type: 'colorScale',
            scale: cs.length === 3
              ? { cfvo: [{ type: 'min' }, { type: 'percentile', val: '50' }, { type: 'max' }], colors: cs }
              : { cfvo: [{ type: 'min' }, { type: 'max' }], colors: cs },
          }), <span className="cf-swatch" style={{ background: `linear-gradient(90deg, ${cs.join(', ')})` }} />)))}
          <div className="menu-sep" />
          {subMenu('clear', 'Удалить правила', [item('Удалить правила из выделенных ячеек', () => onClear(false)), item('Удалить правила со всего листа', () => onClear(true))])}
        </div>
      )}
    </div>
  );
}
