// Диалог «Формат ячеек» (Ctrl+1) — вкладки как в Excel: Число, Выравнивание, Шрифт, Граница, Заливка, Защита.
// Коды форматов показываются и вводятся в русской записи (# ##0,00, ДД.ММ.ГГГГ), в модели — английская, как в файле.
import React, { useEffect, useMemo, useState } from 'react';
import { formatCode, fromLocalCode, toLocalCode } from './numfmt.js';
import { FONT_SIZES } from '../editor/font.js';
import { PALETTE, STANDARD, loadSystemFonts } from '../FontControls.jsx';

// ---------- коды категорий «Число» ----------
const CURRENCIES = [['₽', (n) => `#,##0${n} "₽"`], ['$', (n) => `[$$-409]#,##0${n}`], ['€', (n) => `#,##0${n} [$€-x-euro2]`], ['нет', (n) => `#,##0${n}`]];
const NEGATIVE = [
  ['-1 234,10', (c) => c],
  ['1 234,10 (красным)', (c) => `${c};[Red]${c}`],
  ['-1 234,10 (красным)', (c) => `${c};[Red]-${c}`],
];
const DATES = ['dd.mm.yyyy', 'dd.mm.yy', 'd.m.yyyy', 'yyyy-mm-dd', 'mmm yy', 'mmmm yyyy', 'dd.mm.yyyy h:mm', 'dddd'];
const TIMES = ['h:mm', 'h:mm:ss', 'hh:mm', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'mm:ss', '[h]:mm:ss', 'dd.mm.yyyy h:mm'];
const FRACTIONS = [['Простыми дробями до одной цифры (1/4)', '# ?/?'], ['До двух цифр (21/25)', '# ??/??'], ['До трёх цифр (312/943)', '# ???/???'],
  ['Половинными долями (1/2)', '# ?/2'], ['Четвертными (2/4)', '# ?/4'], ['Восьмыми (4/8)', '# ?/8'], ['Шестнадцатыми (8/16)', '# ??/16'],
  ['Десятыми (3/10)', '# ?/10'], ['Сотыми (30/100)', '# ??/100']];
const CUSTOM = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0.00 "₽"', '0%', '0.00%', '0.00E+00', '# ?/?', 'dd.mm.yyyy', 'dd.mm.yy',
  'd mmm yy', 'mmmm yyyy', 'h:mm', 'h:mm:ss', 'dd.mm.yyyy h:mm', '[h]:mm:ss', '@', '#,##0.00;[Red]-#,##0.00', '#,##0;[Red]-#,##0', '0.00;-0.00;"–"'];
const decimals = (n) => (n > 0 ? `.${'0'.repeat(n)}` : '');
const CATS = ['Общий', 'Числовой', 'Денежный', 'Финансовый', 'Дата', 'Время', 'Процентный', 'Дробный', 'Экспоненциальный', 'Текстовый', '(все форматы)'];

// Категория и параметры по коду ячейки (чтобы диалог открывался на том, что задано)
function guess(code) {
  const c = code || 'General';
  if (c === 'General') return { cat: 'Общий' };
  if (c === '@') return { cat: 'Текстовый' };
  const dec = (/\.(0+)/.exec(c) || [, ''])[1].length;
  if (/^0(\.0+)?%$/.test(c)) return { cat: 'Процентный', dec };
  if (/^0(\.0+)?E\+00$/.test(c)) return { cat: 'Экспоненциальный', dec };
  if (/^(#,##)?0(\.0+)?(;\[Red\]-?(#,##)?0(\.0+)?)?$/.test(c)) return { cat: 'Числовой', dec, group: c.includes(','), neg: /\[Red\]-/.test(c) ? 2 : /\[Red\]/.test(c) ? 1 : 0 };
  if (/^#,##0(\.0+)? "₽"$/.test(c)) return { cat: 'Денежный', dec, cur: 0 };
  if (/^_-\*/.test(c)) return { cat: 'Финансовый', dec };
  if (FRACTIONS.some(([, f]) => f === c)) return { cat: 'Дробный', frac: FRACTIONS.findIndex(([, f]) => f === c) };
  if (TIMES.includes(c) && !DATES.includes(c)) return { cat: 'Время', time: TIMES.indexOf(c) };
  if (DATES.includes(c)) return { cat: 'Дата', date: DATES.indexOf(c) };
  return { cat: '(все форматы)' };
}
function codeOf(st) {
  const d = decimals(st.dec);
  switch (st.cat) {
    case 'Общий': return 'General';
    case 'Числовой': { const base = `${st.group ? '#,##' : ''}0${d}`; return NEGATIVE[st.neg][1](base); }
    case 'Денежный': return CURRENCIES[st.cur][1](d);
    case 'Финансовый': return `_-* #,##0${d} "₽"_-;-* #,##0${d} "₽"_-;_-* "-"${'?'.repeat(st.dec)} "₽"_-;_-@_-`;
    case 'Дата': return DATES[st.date];
    case 'Время': return TIMES[st.time];
    case 'Процентный': return `0${d}%`;
    case 'Дробный': return FRACTIONS[st.frac][1];
    case 'Экспоненциальный': return `0${d}E+00`;
    case 'Текстовый': return '@';
    default: return fromLocalCode(st.local);
  }
}

// ---------- границы ----------
const LINES = [['thin', 'тонкая'], ['hair', 'очень тонкая'], ['dotted', 'точками'], ['dashed', 'штрихами'], ['medium', 'средняя'], ['thick', 'толстая'], ['double', 'двойная']];
const LINE_CSS = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', medium: '2px solid', thick: '3px solid', double: '3px double' };

function ColorGrid({ value, onPick, none }) {
  return (
    <div className="fc-colors">
      {none && <button type="button" className={`fc-none${!value ? ' on' : ''}`} onClick={() => onPick(null)}>{none}</button>}
      <div className="swatches">{[...PALETTE.flat(), ...STANDARD].map((c, i) => (
        <button key={i} type="button" className={`swatch${value === c ? ' on' : ''}`} style={{ background: c }} title={c} onClick={() => onPick(c)} />
      ))}</div>
      <label className="fc-other">Другой: <input type="color" value={value || '#000000'} onChange={(e) => onPick(e.target.value)} /></label>
    </div>
  );
}

export default function FormatCellsDialog({ cell, rg, merged, sample, onApply, onClose }) {
  const [fonts, setFonts] = useState([]);
  useEffect(() => { loadSystemFonts().then(setFonts); }, []);
  const s0 = (cell && cell.s) || {};
  const multi = rg.r2 > rg.r1 || rg.c2 > rg.c1;
  const [tab, setTab] = useState('num');
  // число
  const g = guess(s0.numFmt);
  const [num, setNum] = useState({ cat: g.cat, dec: g.dec ?? 2, group: g.group ?? true, neg: g.neg ?? 0, cur: g.cur ?? 0, date: g.date ?? 0, time: g.time ?? 0, frac: g.frac ?? 0, local: toLocalCode(s0.numFmt || 'General') });
  const code = useMemo(() => { try { return codeOf(num); } catch { return 'General'; } }, [num]);
  const shown = useMemo(() => {
    try { const v = sample ?? 1234.567; const r = formatCode(v, code); return { text: typeof v === 'number' ? r.text : String(v), color: r.color }; } catch { return { text: '—', color: null }; }
  }, [code, sample]);
  // выравнивание, шрифт, заливка, защита
  const [al, setAl] = useState({ align: s0.align || '', valign: s0.valign || '', wrap: !!s0.wrap, merge: !!merged });
  const [font, setFont] = useState({ fontFamily: s0.fontFamily || 'Calibri', fontSize: s0.fontSize || 11, bold: !!s0.bold, italic: !!s0.italic, underline: !!s0.underline, strike: !!s0.strike, color: s0.color || null });
  const [fill, setFill] = useState(s0.fill || null);
  const [prot, setProt] = useState({ locked: !s0.unlocked, hidden: !!s0.formulaHidden });
  // границы: заданные пользователем стороны (остальные — без изменений)
  const b0 = (!multi && s0.border) || {};
  const [pen, setPen] = useState({ style: 'thin', color: '#000000' });
  const [sides, setSides] = useState(() => ({ top: b0.top || null, bottom: b0.bottom || null, left: b0.left || null, right: b0.right || null, insideH: null, insideV: null }));
  const [touched, setTouched] = useState(new Set());
  const toggle = (k) => { setSides({ ...sides, [k]: sides[k] ? null : { ...pen } }); setTouched(new Set([...touched, k])); };
  const preset = (which) => {
    const keys = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
    const n = { ...sides };
    if (which === 'none') keys.forEach((k) => { n[k] = null; });
    if (which === 'outer') ['top', 'bottom', 'left', 'right'].forEach((k) => { n[k] = { ...pen }; });
    if (which === 'inner') ['insideH', 'insideV'].forEach((k) => { n[k] = { ...pen }; });
    setSides(n); setTouched(new Set(which === 'inner' ? [...touched, 'insideH', 'insideV'] : which === 'outer' ? [...touched, 'top', 'bottom', 'left', 'right'] : keys));
  };

  const ok = () => {
    const patch = {
      numFmt: code === 'General' ? null : code,
      align: al.align || null, valign: al.valign || null, wrap: al.wrap || null,
      fontFamily: font.fontFamily && font.fontFamily !== 'Calibri' ? font.fontFamily : null, fontSize: +font.fontSize && +font.fontSize !== 11 ? +font.fontSize : null,
      bold: font.bold || null, italic: font.italic || null, underline: font.underline || null, strike: font.strike || null, color: font.color || null,
      fill: fill || null, unlocked: !prot.locked || null, formulaHidden: prot.hidden || null,
    };
    const borders = {};
    for (const k of touched) borders[k] = sides[k];
    onApply({ patch, borders, merge: al.merge !== !!merged ? al.merge : undefined });
  };

  const tabs = [['num', 'Число'], ['align', 'Выравнивание'], ['font', 'Шрифт'], ['border', 'Граница'], ['fill', 'Заливка'], ['prot', 'Защита']];
  const sideCss = (k) => (sides[k] ? `${LINE_CSS[sides[k].style] || '1px solid'} ${sides[k].color || '#000'}` : '1px dashed #ddd');
  const set = (patch) => setNum({ ...num, ...patch });
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal fc-dlg" role="dialog" aria-label="Формат ячеек" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') ok(); }}>
        <div className="modal-title">Формат ячеек</div>
        <div className="fc-tabs" role="tablist">{tabs.map(([k, l]) => <button key={k} type="button" role="tab" className={`fc-tab${tab === k ? ' on' : ''}`} onClick={() => setTab(k)}>{l}</button>)}</div>
        <div className="modal-body fc-body">
          {tab === 'num' && (
            <div className="fc-num">
              <div className="fc-cats" role="listbox" aria-label="Числовые форматы">
                {CATS.map((c) => <button key={c} type="button" className={`fc-cat${num.cat === c ? ' on' : ''}`} onClick={() => set({ cat: c, local: c === '(все форматы)' ? toLocalCode(code) : num.local })}>{c}</button>)}
              </div>
              <div className="fc-opts">
                <div className="fc-sample"><span>Образец</span><b style={{ color: shown.color || undefined }}>{shown.text || ' '}</b></div>
                {['Числовой', 'Денежный', 'Финансовый', 'Процентный', 'Экспоненциальный'].includes(num.cat) && (
                  <label className="pd-field"><span>Число десятичных знаков</span><input type="number" min="0" max="30" className="pd-input" style={{ width: 70 }} value={num.dec} onChange={(e) => set({ dec: Math.max(0, Math.min(30, +e.target.value || 0)) })} /></label>
                )}
                {num.cat === 'Числовой' && (
                  <>
                    <label className="fc-check"><input type="checkbox" checked={num.group} onChange={(e) => set({ group: e.target.checked })} /> Разделитель групп разрядов ( )</label>
                    <div className="fc-sub">Отрицательные числа:</div>
                    <div className="fc-list">{NEGATIVE.map(([l], i) => <button key={l} type="button" className={`fc-item${num.neg === i ? ' on' : ''}`} style={{ color: i ? '#c00000' : undefined }} onClick={() => set({ neg: i })}>{l}</button>)}</div>
                  </>
                )}
                {num.cat === 'Денежный' && (
                  <label className="pd-field"><span>Обозначение</span>
                    <select value={num.cur} onChange={(e) => set({ cur: +e.target.value })}>{CURRENCIES.map(([l], i) => <option key={l} value={i}>{l}</option>)}</select>
                  </label>
                )}
                {num.cat === 'Дата' && <div className="fc-list">{DATES.map((d, i) => <button key={d} type="button" className={`fc-item${num.date === i ? ' on' : ''}`} onClick={() => set({ date: i })}>{formatCode(46054.75, d).text}</button>)}</div>}
                {num.cat === 'Время' && <div className="fc-list">{TIMES.map((d, i) => <button key={d} type="button" className={`fc-item${num.time === i ? ' on' : ''}`} onClick={() => set({ time: i })}>{formatCode(46054.5626, d).text}</button>)}</div>}
                {num.cat === 'Дробный' && <div className="fc-list">{FRACTIONS.map(([l], i) => <button key={l} type="button" className={`fc-item${num.frac === i ? ' on' : ''}`} onClick={() => set({ frac: i })}>{l}</button>)}</div>}
                {num.cat === '(все форматы)' && (
                  <>
                    <label className="pd-field"><span>Тип</span><input className="pd-input fc-code" value={num.local} onChange={(e) => set({ local: e.target.value })} /></label>
                    <div className="fc-list">{CUSTOM.map((c) => { const l = toLocalCode(c); return <button key={c} type="button" className={`fc-item${num.local === l ? ' on' : ''}`} onClick={() => set({ local: l })}>{l}</button>; })}</div>
                  </>
                )}
                <div className="fc-hint">{{
                  Общий: 'Ячейки общего формата не имеют определённого числового формата.',
                  Числовой: 'Для вывода чисел; для денежных значений есть форматы «Денежный» и «Финансовый».',
                  Денежный: 'Для денежных значений.', Финансовый: 'Выравнивает обозначения валюты и десятичные разделители в столбце.',
                  Дата: 'Даты хранятся числами (днями с 1900 года).', Время: 'Время — доля суток.', Процентный: 'Значение умножается на 100 и выводится со знаком %.',
                  Текстовый: 'Значение показывается так, как введено, даже если это число.',
                  '(все форматы)': 'Код формата: 0 — цифра, # — цифра без незначащих нулей, пробел — разряды, «;» — разделы для >0, <0, 0 и текста.',
                }[num.cat] || ''}</div>
              </div>
            </div>
          )}
          {tab === 'align' && (
            <div className="fc-grid2">
              <label className="pd-field"><span>По горизонтали</span>
                <select value={al.align} onChange={(e) => setAl({ ...al, align: e.target.value })}>
                  <option value="">по значению</option><option value="left">по левому краю</option><option value="center">по центру</option><option value="right">по правому краю</option><option value="justify">по ширине</option>
                </select>
              </label>
              <label className="pd-field"><span>По вертикали</span>
                <select value={al.valign} onChange={(e) => setAl({ ...al, valign: e.target.value })}>
                  <option value="top">по верхнему краю</option><option value="center">по центру</option><option value="">по нижнему краю</option>
                </select>
              </label>
              <div className="fc-sub">Отображение</div>
              <label className="fc-check"><input type="checkbox" checked={al.wrap} onChange={(e) => setAl({ ...al, wrap: e.target.checked })} /> Переносить текст</label>
              <label className="fc-check"><input type="checkbox" checked={al.merge} disabled={!multi && !merged} onChange={(e) => setAl({ ...al, merge: e.target.checked })} /> Объединение ячеек</label>
            </div>
          )}
          {tab === 'font' && (
            <div className="fc-grid2">
              <label className="pd-field"><span>Шрифт</span>
                <input className="pd-input" list="fc-fonts" value={font.fontFamily} onChange={(e) => setFont({ ...font, fontFamily: e.target.value })} />
                <datalist id="fc-fonts">{fonts.map((f) => <option key={f} value={f} />)}</datalist>
              </label>
              <label className="pd-field"><span>Начертание</span>
                <select value={`${font.bold ? 'b' : ''}${font.italic ? 'i' : ''}`} onChange={(e) => setFont({ ...font, bold: e.target.value.includes('b'), italic: e.target.value.includes('i') })}>
                  <option value="">обычный</option><option value="i">курсив</option><option value="b">полужирный</option><option value="bi">полужирный курсив</option>
                </select>
              </label>
              <label className="pd-field"><span>Размер</span>
                <input className="pd-input" list="fc-sizes" style={{ width: 70 }} value={font.fontSize} onChange={(e) => setFont({ ...font, fontSize: e.target.value })} />
                <datalist id="fc-sizes">{FONT_SIZES.map((x) => <option key={x} value={x} />)}</datalist>
              </label>
              <label className="pd-field"><span>Подчёркивание</span>
                <select value={font.underline ? '1' : ''} onChange={(e) => setFont({ ...font, underline: !!e.target.value })}><option value="">нет</option><option value="1">одинарное</option></select>
              </label>
              <label className="fc-check"><input type="checkbox" checked={font.strike} onChange={(e) => setFont({ ...font, strike: e.target.checked })} /> Зачёркнутый</label>
              <div className="fc-sub">Цвет</div>
              <ColorGrid value={font.color} none="Авто" onPick={(c) => setFont({ ...font, color: c })} />
              <div className="fc-sample fc-font-sample"><span>Образец</span>
                <b style={{ fontFamily: font.fontFamily, fontSize: `${Math.min(28, +font.fontSize || 11)}pt`, fontWeight: font.bold ? 700 : 400, fontStyle: font.italic ? 'italic' : 'normal', textDecoration: [font.underline && 'underline', font.strike && 'line-through'].filter(Boolean).join(' ') || 'none', color: font.color || '#000' }}>АаБбВв AaBbCc</b>
              </div>
            </div>
          )}
          {tab === 'border' && (
            <div className="fc-border">
              <div className="fc-border-left">
                <div className="fc-sub">Линия</div>
                <div className="fc-lines">{LINES.map(([k, l]) => (
                  <button key={k} type="button" className={`fc-line${pen.style === k ? ' on' : ''}`} title={l} onClick={() => setPen({ ...pen, style: k })}><span style={{ borderTop: `${LINE_CSS[k]} #000` }} /></button>
                ))}</div>
                <label className="pd-field"><span>Цвет</span><input type="color" value={pen.color} onChange={(e) => setPen({ ...pen, color: e.target.value })} /></label>
              </div>
              <div className="fc-border-right">
                <div className="fc-sub">Все</div>
                <div className="tabs-actions">
                  <button type="button" className="btn small" onClick={() => preset('none')}>Нет</button>
                  <button type="button" className="btn small" onClick={() => preset('outer')}>Внешние</button>
                  <button type="button" className="btn small" disabled={!multi} onClick={() => preset('inner')}>Внутренние</button>
                </div>
                <div className="fc-sub">Отдельные (щёлкните сторону)</div>
                <div className="fc-bpreview">
                  <button type="button" className="fc-bside top" onClick={() => toggle('top')} style={{ borderTop: sideCss('top') }} title="Верхняя" />
                  <button type="button" className="fc-bside bottom" onClick={() => toggle('bottom')} style={{ borderTop: sideCss('bottom') }} title="Нижняя" />
                  <button type="button" className="fc-bside left" onClick={() => toggle('left')} style={{ borderLeft: sideCss('left') }} title="Левая" />
                  <button type="button" className="fc-bside right" onClick={() => toggle('right')} style={{ borderLeft: sideCss('right') }} title="Правая" />
                  {multi && <button type="button" className="fc-bside midh" onClick={() => toggle('insideH')} style={{ borderTop: sideCss('insideH') }} title="Внутренние горизонтальные" />}
                  {multi && <button type="button" className="fc-bside midv" onClick={() => toggle('insideV')} style={{ borderLeft: sideCss('insideV') }} title="Внутренние вертикальные" />}
                  <span className="fc-btext">{multi ? 'Текст  Текст' : 'Текст'}</span>
                </div>
              </div>
            </div>
          )}
          {tab === 'fill' && (
            <div>
              <div className="fc-sub">Цвет фона</div>
              <ColorGrid value={fill} none="Нет цвета" onPick={setFill} />
              <div className="fc-sample"><span>Образец</span><b style={{ background: fill || 'transparent', display: 'block', height: 28 }} /></div>
            </div>
          )}
          {tab === 'prot' && (
            <div>
              <label className="fc-check"><input type="checkbox" checked={prot.locked} onChange={(e) => setProt({ ...prot, locked: e.target.checked })} /> Защищаемая ячейка</label>
              <label className="fc-check"><input type="checkbox" checked={prot.hidden} onChange={(e) => setProt({ ...prot, hidden: e.target.checked })} /> Скрыть формулы</label>
              <div className="fc-hint">Блокировка ячеек и скрытие формул действуют только после защиты листа (Рецензирование → Защитить лист).</div>
            </div>
          )}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
