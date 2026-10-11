// Сетка листа: виртуализированная отрисовка видимых ячеек, заголовки строк/столбцов, выделение,
// изменение ширины столбцов и высоты строк мышью, редактор ячейки поверх сетки.
import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { colName, key } from './address.js';
import { formatValue } from './format.js';
import { formatCode } from './numfmt.js';
import { isErr } from './formula.js';
import { DEFAULT_COL_W, DEFAULT_ROW_H, inputOf, isHidden, maxLevel, mergeAt, outlineRuns, spillAnchorAt } from './model.js';
import { cycleRef } from './formula.js';
import { AssistPopup, useFormulaAssist } from './FormulaAssist.jsx';
import { cssFontStack } from '../editor/font.js';

const HEAD_W = 46; // ширина заголовков строк
const HEAD_H = 22; // высота заголовков столбцов

// Накопленные координаты: pos[i] — левый/верхний край i-й строки/столбца
function prefix(count, sizeOf) {
  const pos = new Float64Array(count + 1);
  for (let i = 0; i < count; i++) pos[i + 1] = pos[i] + sizeOf(i);
  return pos;
}
const findIndex = (pos, x) => {
  let lo = 0; let hi = pos.length - 2;
  while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (pos[mid] <= x) lo = mid; else hi = mid - 1; }
  return lo;
};

const BORDER = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted' };
const borderCss = (b) => (b ? `${BORDER[b.style] || '1px solid'} ${b.color || '#000'}` : undefined);

// Ширина текста в пикселях (canvas, с кешем по шрифту и тексту)
const measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
const widthCache = new Map();
function textWidth(text, st) {
  if (!measureCtx) return 0;
  const font = `${st.fontStyle || ''} ${st.fontWeight || ''} ${st.fontSize || '11pt'} ${st.fontFamily || "Calibri, Carlito, 'Segoe UI', sans-serif"}`;
  const k = `${font}|${text}`;
  let w = widthCache.get(k);
  if (w == null) {
    measureCtx.font = font;
    w = Math.max(...String(text).split(String.fromCharCode(10)).map((l) => measureCtx.measureText(l).width));
    if (widthCache.size > 5000) widthCache.clear();
    widthCache.set(k, w);
  }
  return w;
}

// Число, которое не помещается: «Общий» — округляем, иначе — решётки (как в Excel)
function fitNumber(cell, text, width, st) {
  if (textWidth(text, st) <= width) return text;
  const fmt = cell.s && cell.s.numFmt;
  if ((!fmt || fmt === 'General') && Number.isFinite(cell.v)) {
    for (let d = 9; d >= 0; d--) {
      const t = formatValue(Math.round(cell.v * 10 ** d) / 10 ** d);
      if (textWidth(t, st) <= width) return t;
    }
    const e = cell.v.toExponential(2).replace('.', ',').replace('e', 'E');
    if (textWidth(e, st) <= width) return e;
  }
  const hash = textWidth('#', st) || 7;
  return '#'.repeat(Math.max(1, Math.floor(width / hash)));
}

export function cellText(cell) {
  if (!cell) return '';
  return formatValue(cell.v, cell.s && cell.s.numFmt);
}

function cellStyle(cell) {
  const s = (cell && cell.s) || {};
  const v = cell && cell.v;
  const align = s.align || (typeof v === 'number' ? 'right' : typeof v === 'boolean' || isErr(v) ? 'center' : 'left');
  return {
    fontWeight: s.bold ? 'bold' : undefined,
    fontStyle: s.italic ? 'italic' : undefined,
    textDecoration: [s.underline && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ') || undefined,
    // цвет из числового формата ([Красный] у отрицательных и т. п.) важнее цвета шрифта, как в Excel
    color: (s.numFmt && typeof v === 'number' && formatCode(v, s.numFmt).color) || s.color || undefined,
    background: s.fill || undefined,
    fontSize: s.fontSize ? `${s.fontSize}pt` : undefined,
    fontFamily: s.fontFamily ? cssFontStack(s.fontFamily) : undefined,
    justifyContent: align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start',
    textAlign: align,
    alignItems: s.valign === 'top' ? 'flex-start' : s.valign === 'center' ? 'center' : 'flex-end',
    whiteSpace: s.wrap ? 'pre-wrap' : 'pre',
    borderTop: borderCss(s.border && s.border.top),
    borderLeft: borderCss(s.border && s.border.left),
    borderRight: borderCss(s.border && s.border.right),
    borderBottom: borderCss(s.border && s.border.bottom),
  };
}

const OUT_STEP = 14; // ширина уровня в поле структуры (кнопки «+/−» групп)

// refMarks — ссылки редактируемой формулы (цветные рамки); showFormulas — формулы вместо значений (Ctrl+`);
// arrows — стрелки влияющих/зависимых ячеек [{ from: диапазон, to: { r, c } }]; onErrorInfo — значок у ячейки с ошибкой
export default function Grid({ sheet, sel, setSel, editing, setEditing, onCommit, onColWidth, onRowHeight, onRefInsert, onOutline, gridRef,
  refMarks = [], showFormulas = false, arrows = [], onErrorInfo, names = [], onFill, onFilterMenu, cf = null }) {
  const scrollRef = useRef(null);
  const [view, setView] = useState({ top: 0, left: 0, w: 800, h: 500 });
  // размеры области: использованная часть + запас (как «бесконечный» лист)
  const used = useMemo(() => {
    let r = 0; let c = 0;
    for (const k of sheet.cells.keys()) { const i = k.indexOf(','); r = Math.max(r, +k.slice(0, i)); c = Math.max(c, +k.slice(i + 1)); }
    return { r, c };
  }, [sheet.cells]);
  const fr = (sheet.freeze && sheet.freeze.rows) || 0; // закреплённые строки и столбцы
  const fc = (sheet.freeze && sheet.freeze.cols) || 0;
  const rows = Math.max(200, used.r + 60, sel.r2 + 30, sel.r + 30, fr + 30);
  const cols = Math.max(30, used.c + 12, sel.c2 + 8, sel.c + 8, fc + 8);
  // скрытые строки и столбцы — нулевого размера
  const colPos = useMemo(() => prefix(cols, (c) => (isHidden(sheet, 'col', c) ? 0 : sheet.colW[c] ?? DEFAULT_COL_W)), [cols, sheet.colW, sheet.colInfo]);
  const rowPos = useMemo(() => prefix(rows, (r) => (isHidden(sheet, 'row', r) ? 0 : sheet.rowH[r] ?? DEFAULT_ROW_H)), [rows, sheet.rowH, sheet.rowInfo]);
  // поля структуры (группировка): слева от заголовков строк и над заголовками столбцов
  const rowLevels = useMemo(() => maxLevel(sheet, 'row'), [sheet.rowInfo]);
  const colLevels = useMemo(() => maxLevel(sheet, 'col'), [sheet.colInfo]);
  const rowGut = rowLevels ? (rowLevels + 1) * OUT_STEP + 4 : 0;
  const colGut = colLevels ? (colLevels + 1) * OUT_STEP + 4 : 0;
  const HW = HEAD_W + rowGut;
  const HH = HEAD_H + colGut;
  const frozenW = colPos[fc];
  const frozenH = rowPos[fr];

  useLayoutEffect(() => {
    const el = scrollRef.current;
    const update = () => setView({ top: el.scrollTop, left: el.scrollLeft, w: el.clientWidth, h: el.clientHeight });
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // прокрутка к активной ячейке (закреплённая часть не прокручивается)
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const vw = el.clientWidth - HW - frozenW; const vh = el.clientHeight - HH - frozenH;
    if (sel.c >= fc) {
      const x1 = colPos[sel.c] - frozenW; const x2 = colPos[sel.c + 1] - frozenW;
      if (x1 < el.scrollLeft) el.scrollLeft = x1; else if (x2 > el.scrollLeft + vw) el.scrollLeft = x2 - vw;
    }
    if (sel.r >= fr) {
      const y1 = rowPos[sel.r] - frozenH; const y2 = rowPos[sel.r + 1] - frozenH;
      if (y1 < el.scrollTop) el.scrollTop = y1; else if (y2 > el.scrollTop + vh) el.scrollTop = y2 - vh;
    }
  }, [sel.r, sel.c, colPos, rowPos, fr, fc, HW, HH]);

  const r0 = findIndex(rowPos, view.top);
  const r1 = Math.min(rows - 1, findIndex(rowPos, view.top + view.h) + 1);
  const c0 = findIndex(colPos, view.left);
  const c1 = Math.min(cols - 1, findIndex(colPos, view.left + view.w) + 1);

  // ячейка под точкой: в закреплённой части — без прокрутки
  const cellAt = (clientX, clientY) => {
    const rect = scrollRef.current.getBoundingClientRect();
    let x = clientX - rect.left - HW;
    let y = clientY - rect.top - HH;
    if (x >= frozenW) x += scrollRef.current.scrollLeft;
    if (y >= frozenH) y += scrollRef.current.scrollTop;
    return { r: Math.max(0, findIndex(rowPos, Math.max(0, y))), c: Math.max(0, findIndex(colPos, Math.max(0, x))) };
  };

  const drag = useRef(null);
  const onMouseDown = (e) => {
    if (e.button !== 0) return;
    const { r, c } = cellAt(e.clientX, e.clientY);
    if (editing && editing.text.startsWith('=') && !(r === editing.r && c === editing.c)) {
      e.preventDefault();
      onRefInsert({ r1: r, c1: c, r2: r, c2: c }); // щелчок при наборе формулы — вставить ссылку
      drag.current = { mode: 'ref', r, c };
    } else {
      if (editing) onCommit(editing.text, null);
      if (e.shiftKey) setSel({ ...sel, r2: r, c2: c });
      else setSel({ r, c, r2: r, c2: c });
      drag.current = { mode: 'sel' };
    }
    gridRef.current && gridRef.current.focus();
    const move = (ev) => {
      const p = cellAt(ev.clientX, ev.clientY);
      if (drag.current.mode === 'ref') onRefInsert({ r1: Math.min(drag.current.r, p.r), c1: Math.min(drag.current.c, p.c), r2: Math.max(drag.current.r, p.r), c2: Math.max(drag.current.c, p.c) }, true);
      else setSel((s) => ({ ...s, r2: p.r, c2: p.c }));
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); drag.current = null; };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  // маркер заполнения (квадратик в углу выделения): протягивание — ряд/копия (Ctrl — наоборот), назад — очистка;
  // двойной щелчок — вниз до конца данных соседнего столбца
  const [fillPrev, setFillPrev] = useState(null);
  const fillTarget = (p) => {
    const r1 = Math.min(sel.r, sel.r2); const r2 = Math.max(sel.r, sel.r2); const c1 = Math.min(sel.c, sel.c2); const c2 = Math.max(sel.c, sel.c2);
    const dr = p.r > r2 ? p.r - r2 : p.r < r1 ? p.r - r1 : 0;
    const dc = p.c > c2 ? p.c - c2 : p.c < c1 ? p.c - c1 : 0;
    if (!dr && !dc) {
      // внутри выделения — укоротить (очистка хвоста), по большей стороне
      if (r2 > r1 && p.r < r2 && (r2 - p.r) >= (c2 - p.c)) return { r1, c1, r2: p.r, c2 };
      if (c2 > c1 && p.c < c2) return { r1, c1, r2, c2: p.c };
      return { r1, c1, r2, c2 };
    }
    if (Math.abs(dr) >= Math.abs(dc)) return dr > 0 ? { r1, c1, r2: p.r, c2 } : { r1: p.r, c1, r2, c2 };
    return dc > 0 ? { r1, c1, r2, c2: p.c } : { r1, c1: p.c, r2, c2 };
  };
  const startFill = (e) => {
    if (e.button !== 0) return;
    e.preventDefault(); e.stopPropagation();
    const src = { r1: Math.min(sel.r, sel.r2), c1: Math.min(sel.c, sel.c2), r2: Math.max(sel.r, sel.r2), c2: Math.max(sel.c, sel.c2) };
    let dst = src;
    setFillPrev(src);
    const move = (ev) => { dst = fillTarget(cellAt(ev.clientX, ev.clientY)); setFillPrev(dst); };
    const up = (ev) => {
      window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up);
      setFillPrev(null);
      if (dst.r1 !== src.r1 || dst.r2 !== src.r2 || dst.c1 !== src.c1 || dst.c2 !== src.c2) onFill(src, dst, ev.ctrlKey);
    };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };
  const fillDown = (e) => {
    e.preventDefault(); e.stopPropagation();
    const src = { r1: Math.min(sel.r, sel.r2), c1: Math.min(sel.c, sel.c2), r2: Math.max(sel.r, sel.r2), c2: Math.max(sel.c, sel.c2) };
    const filled = (r, c) => c >= 0 && !isEmpty(r, c);
    // соседний столбец слева (иначе справа): сплошные данные ниже выделения
    const side = filled(src.r2 + 1, src.c1 - 1) ? src.c1 - 1 : filled(src.r2 + 1, src.c2 + 1) ? src.c2 + 1 : null;
    if (side == null) return;
    let last = src.r2;
    while (last + 1 < 1048576 && filled(last + 1, side)) last++;
    if (last > src.r2) onFill(src, { ...src, r2: last }, false);
  };

  // изменение ширины столбца / высоты строки за край заголовка
  const startResize = (e, axis, i) => {
    e.preventDefault();
    e.stopPropagation();
    const start = axis === 'col' ? e.clientX : e.clientY;
    const size0 = axis === 'col' ? colPos[i + 1] - colPos[i] : rowPos[i + 1] - rowPos[i];
    const guide = document.createElement('div');
    guide.className = `grid-guide ${axis}`;
    scrollRef.current.parentElement.appendChild(guide);
    const place = (v) => {
      const rect = scrollRef.current.getBoundingClientRect();
      if (axis === 'col') guide.style.left = `${v - rect.left}px`; else guide.style.top = `${v - rect.top}px`;
    };
    place(start);
    let size = size0;
    const move = (ev) => { const d = (axis === 'col' ? ev.clientX : ev.clientY) - start; size = Math.max(8, size0 + d); place(axis === 'col' ? ev.clientX : ev.clientY); };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      guide.remove();
      if (Math.abs(size - size0) >= 1) (axis === 'col' ? onColWidth : onRowHeight)(i, size);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  const selR1 = Math.min(sel.r, sel.r2); const selR2 = Math.max(sel.r, sel.r2);
  const selC1 = Math.min(sel.c, sel.c2); const selC2 = Math.max(sel.c, sel.c2);
  // ячейка с учётом «разлива»: пустая ячейка внутри «разлива» показывает его значение
  const cellAtKey = (r, c) => {
    const x = sheet.cells.get(key(r, c));
    if (sheet.spill && (!x || (x.v == null && x.f == null))) { const g = sheet.spill.get(key(r, c)); if (g !== undefined) return { ...(x || {}), v: g }; }
    return x;
  };
  const isEmpty = (r, c) => { const x = cellAtKey(r, c); return !x || x.v == null || x.v === ''; };
  // «разлив», в котором стоит курсор, обводится рамкой (как в Excel)
  const spillSel = spillAnchorAt(sheet, sel.r, sel.c);

  // ячейки строк ra..rb и столбцов ca..cb (объединённые — одной ячейкой от левой верхней). Текст без переноса
  // «перетекает» в пустые соседние ячейки справа (как в Excel), поэтому смотрим и немного левее области.
  const renderCells = (ra, rb, ca, cb, cMin = 0) => {
    const out = [];
    const drawnMerges = new Set();
    for (let r = ra; r <= rb; r++) {
      if (rowPos[r + 1] === rowPos[r] && !(sheet.merges.length && mergeAt(sheet, r, ca))) continue; // скрытая строка
      for (let c = Math.max(cMin, ca - 20); c <= cb; c++) {
        const m = sheet.merges.length ? mergeAt(sheet, r, c) : null;
        if (c < ca && m) continue;
        let rr = r; let cc = c; let r2 = r; let c2 = c;
        if (m) {
          const mk = `${m.r1},${m.c1}`;
          if (drawnMerges.has(mk)) continue;
          drawnMerges.add(mk);
          rr = m.r1; cc = m.c1; r2 = m.r2; c2 = m.c2;
        }
        const own0 = cellAtKey(rr, cc);
        const cfv = cf && cf.get(key(rr, cc));
        const cell = cfv ? { ...(own0 || {}), s: { ...((own0 && own0.s) || {}), ...cfv.style, border: cfv.style.border ? { ...((own0 && own0.s && own0.s.border) || {}), ...cfv.style.border } : own0 && own0.s && own0.s.border } } : own0;
        if (!cell && !m) continue;
        let text = showFormulas && cell && cell.f != null ? inputOf(cell) : cellText(cell);
        if (!text && !(cell && cell.s && (cell.s.fill || cell.s.border)) && !m && !(cfv && cfv.bar)) continue;
        const own = colPos[c2 + 1] - colPos[cc];
        if (own <= 0 && !m) continue; // скрытый столбец
        const st = cellStyle(cell);
        const box = { left: colPos[cc], top: rowPos[rr], width: own, height: rowPos[r2 + 1] - rowPos[rr] };
        const s0 = (cell && cell.s) || {};
        if (showFormulas && cell && cell.f != null) { st.justifyContent = 'flex-start'; st.textAlign = 'left'; }
        else if (text && typeof cell.v === 'number' && !s0.wrap) text = fitNumber(cell, text, own - 3, st);
        // перетекание текста вправо
        let spill = own;
        if (text && !m && typeof cell.v === 'string' && !s0.wrap && st.textAlign === 'left') {
          const need = textWidth(text, st) + 6;
          let cn = c2 + 1;
          while (spill < need && cn < cols && isEmpty(rr, cn) && !(sheet.merges.length && mergeAt(sheet, rr, cn))) { spill += colPos[cn + 1] - colPos[cn]; cn++; }
        }
        if (c < ca && spill === own) continue;
        // гистограмма условного форматирования — полоса под текстом
        if (cfv && cfv.bar) out.push(<div key={`b${rr},${cc}`} className="gbar" style={{ left: box.left + 1, top: box.top + 2, width: Math.max(0, (own - 3) * cfv.bar.pct), height: Math.max(0, box.height - 4), background: `linear-gradient(90deg, ${cfv.bar.color}, ${cfv.bar.color}33)` }} />);
        if (spill > own) {
          // фон и границы — только своей ячейки, текст — поверх соседних
          out.push(<div key={`f${rr},${cc}`} className="gcell" style={{ ...box, ...st, color: undefined }} />);
          out.push(<div key={`${rr},${cc}`} className="gcell" style={{ ...box, ...st, width: spill, background: undefined, borderTop: undefined, borderLeft: undefined, borderRight: undefined, borderBottom: undefined }}>{text}</div>);
        } else {
          out.push(<div key={`${rr},${cc}`} className={`gcell${m ? ' merged' : ''}`} style={{ ...box, ...st }}>{text}</div>);
        }
      }
    }
    return out;
  };
  // линии сетки — фоном, ячейки — поверх
  const gridLines = (ra, rb, ca, cb) => (
    <svg className="grid-lines" width={colPos[cb + 1] - colPos[ca]} height={rowPos[rb + 1] - rowPos[ra]} style={{ left: colPos[ca], top: rowPos[ra] }}>
      {Array.from({ length: cb - ca + 2 }, (_, i) => { if (i > 0 && colPos[ca + i] === colPos[ca + i - 1]) return null; const x = colPos[ca + i] - colPos[ca] + 0.5; return <line key={`v${i}`} x1={x} x2={x} y1={0} y2="100%" />; })}
      {Array.from({ length: rb - ra + 2 }, (_, i) => { if (i > 0 && rowPos[ra + i] === rowPos[ra + i - 1]) return null; const y = rowPos[ra + i] - rowPos[ra] + 0.5; return <line key={`h${i}`} y1={y} y2={y} x1={0} x2="100%" />; })}
    </svg>
  );

  const ed = editing && { left: colPos[editing.c], top: rowPos[editing.r], width: Math.max(colPos[editing.c + 1] - colPos[editing.c], 80), height: rowPos[editing.r + 1] - rowPos[editing.r] };
  const active = mergeAt(sheet, sel.r, sel.c) || { r1: sel.r, c1: sel.c, r2: sel.r, c2: sel.c };
  const marks = (
    <>
      <div className="gsel" style={{ left: colPos[selC1], top: rowPos[selR1], width: colPos[selC2 + 1] - colPos[selC1], height: rowPos[selR2 + 1] - rowPos[selR1] }} />
      <div className="gactive" style={{ left: colPos[active.c1], top: rowPos[active.r1], width: colPos[active.c2 + 1] - colPos[active.c1], height: rowPos[active.r2 + 1] - rowPos[active.r1] }} />
      {refMarks.map((m, i) => (
        <div key={`rm${i}`} className="gref" style={{ left: colPos[m.c1], top: rowPos[m.r1], width: colPos[Math.min(m.c2, cols - 1) + 1] - colPos[m.c1], height: rowPos[Math.min(m.r2, rows - 1) + 1] - rowPos[m.r1], borderColor: m.color, background: `${m.color}14` }} />
      ))}
      {sheet.autoFilter && onFilterMenu && Array.from({ length: sheet.autoFilter.ref.c2 - sheet.autoFilter.ref.c1 + 1 }, (_, j) => {
        const af = sheet.autoFilter; const c = af.ref.c1 + j; const r = af.ref.r1;
        const w = colPos[c + 1] - colPos[c]; const h = rowPos[r + 1] - rowPos[r];
        if (!w || !h) return null;
        const on = !!(af.cols[j] || (af.raw && af.raw[j]));
        return (
          <button key={`af${j}`} type="button" className={`gfbtn${on ? ' on' : ''}`} title={on ? 'Фильтр включён' : 'Фильтр и сортировка'}
            style={{ left: colPos[c + 1] - 18, top: rowPos[r] + Math.max(0, h - 18) }}
            onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
            onClick={(e) => { e.stopPropagation(); onFilterMenu(j, e.currentTarget.getBoundingClientRect()); }}>{on ? '⏷' : '▾'}</button>
        );
      })}
      {onFill && !editing && !sel.whole && (
        <div className="gfill" title="Протяните, чтобы заполнить (Ctrl — копировать/ряд); двойной щелчок — до конца данных"
          style={{ left: colPos[selC2 + 1] - 4, top: rowPos[selR2 + 1] - 4 }} onMouseDown={startFill} onDoubleClick={fillDown} />
      )}
      {fillPrev && (
        <div className="gfillprev" style={{ left: colPos[fillPrev.c1], top: rowPos[fillPrev.r1], width: colPos[fillPrev.c2 + 1] - colPos[fillPrev.c1], height: rowPos[fillPrev.r2 + 1] - rowPos[fillPrev.r1] }} />
      )}
      {spillSel && (
        <div className="gspill" style={{ left: colPos[spillSel.range.c1], top: rowPos[spillSel.range.r1], width: colPos[spillSel.range.c2 + 1] - colPos[spillSel.range.c1], height: rowPos[spillSel.range.r2 + 1] - rowPos[spillSel.range.r1] }} />
      )}
    </>
  );
  // стрелки влияющих/зависимых ячеек (как «Зависимости формул» Excel): от центра источника к центру ячейки
  const center = (r, c) => [(colPos[c] + colPos[c + 1]) / 2, (rowPos[r] + rowPos[r + 1]) / 2];
  const arrowLayer = arrows.length > 0 && (
    <svg className="garrows" width={colPos[cols]} height={rowPos[rows]}>
      <defs><marker id="garrow-head" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#2f6fbf" /></marker></defs>
      {arrows.map((a, i) => {
        const [x1, y1] = center(a.from.r1, a.from.c1);
        const [x2, y2] = center(a.to.r, a.to.c);
        const range = a.from.r2 > a.from.r1 || a.from.c2 > a.from.c1;
        return (
          <g key={i}>
            {range && <rect x={colPos[a.from.c1] + 0.5} y={rowPos[a.from.r1] + 0.5} width={colPos[a.from.c2 + 1] - colPos[a.from.c1] - 1} height={rowPos[a.from.r2 + 1] - rowPos[a.from.r1] - 1} fill="none" stroke="#2f6fbf" strokeWidth="1.5" />}
            <circle cx={x1} cy={y1} r="3" fill="#2f6fbf" />
            <line x1={x1} y1={y1} x2={x2} y2={y2} stroke="#2f6fbf" strokeWidth="1.5" markerEnd="url(#garrow-head)" />
          </g>
        );
      })}
    </svg>
  );
  const activeCellVal = cellAtKey(sel.r, sel.c);
  const errBadge = onErrorInfo && !editing && activeCellVal && activeCellVal.v && typeof activeCellVal.v === 'object' && 'error' in activeCellVal.v && (
    <button type="button" className="gerr" title="Ошибка в ячейке: подробнее" style={{ left: colPos[sel.c] - 20, top: rowPos[sel.r] + (rowPos[sel.r + 1] - rowPos[sel.r] - 18) / 2 }}
      onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }} onClick={(e) => { e.stopPropagation(); onErrorInfo(sel.r, sel.c, e.currentTarget.getBoundingClientRect()); }}>!</button>
  );
  // редактор — в той части листа (закреплённой или нет), где ячейка
  const editorIn = (top, left) => (editing && editing.text !== null && !editing.fromBar && (editing.r < fr) === top && (editing.c < fc) === left
    ? <CellEditor key={`${editing.r},${editing.c}`} box={ed} editing={editing} setEditing={setEditing} onCommit={onCommit} style={cellStyle(sheet.cells.get(key(editing.r, editing.c)))} names={names} />
    : null);

  // заголовки столбцов ca..cb и строк ra..rb
  const colHeads = (ca, cb) => Array.from({ length: cb - ca + 1 }, (_, i) => {
    const c = ca + i;
    const w = colPos[c + 1] - colPos[c];
    if (!w) return null;
    const on = c >= selC1 && c <= selC2;
    return (
      <div key={c} className={`ghead${on ? ' on' : ''}`} style={{ left: colPos[c], width: w }}
        onMouseDown={(e) => { if (e.target.classList.contains('gresize')) return; setSel(e.shiftKey ? { ...sel, c2: c, r: 0, r2: rows - 1 } : { r: 0, c, r2: rows - 1, c2: c, whole: 'col' }); gridRef.current && gridRef.current.focus(); }}>
        {colName(c)}
        <span className="gresize col" onMouseDown={(e) => startResize(e, 'col', c)} title="Ширина столбца" />
      </div>
    );
  });
  // номера строк под действующим фильтром — синие (как Excel)
  const afRows = sheet.autoFilter && Object.keys(sheet.autoFilter.cols).length ? sheet.autoFilter.ref : null;
  const rowHeads = (ra, rb) => Array.from({ length: rb - ra + 1 }, (_, i) => {
    const r = ra + i;
    const h = rowPos[r + 1] - rowPos[r];
    if (!h) return null;
    const on = r >= selR1 && r <= selR2;
    return (
      <div key={r} className={`ghead row${on ? ' on' : ''}${afRows && r > afRows.r1 && r <= afRows.r2 ? ' filtered' : ''}`} style={{ top: rowPos[r], height: h }}
        onMouseDown={(e) => { if (e.target.classList.contains('gresize')) return; setSel(e.shiftKey ? { ...sel, r2: r, c: 0, c2: cols - 1 } : { r, c: 0, r2: r, c2: cols - 1, whole: 'row' }); gridRef.current && gridRef.current.focus(); }}>
        {r + 1}
        <span className="gresize row" onMouseDown={(e) => startResize(e, 'row', r)} title="Высота строки" />
      </div>
    );
  });

  // поле структуры: линии групп и кнопки «+/−» у итоговой строки (под группой) / столбца (справа)
  const outline = (axis) => {
    const levels = axis === 'row' ? rowLevels : colLevels;
    const pos = axis === 'row' ? rowPos : colPos;
    const [lo, hi] = axis === 'row' ? [r0, r1] : [c0, c1];
    const out = [];
    for (let lvl = 1; lvl <= levels; lvl++) {
      for (const run of outlineRuns(sheet, axis, lvl)) {
        if (run.b + 1 < lo || run.a > hi + 1) continue;
        const collapsed = pos[run.b + 1] === pos[run.a];
        const across = (lvl - 1) * OUT_STEP + 4;
        const len = pos[run.b + 1] - pos[run.a];
        if (len > 0) {
          out.push(<div key={`l${lvl}-${run.a}`} className={`gout-line ${axis}`} style={axis === 'row'
            ? { top: pos[run.a], height: len, left: across + 6 } : { left: pos[run.a], width: len, top: across + 6 }} />);
        }
        const at = pos[run.b + 1] + ((pos[run.b + 2] - pos[run.b + 1]) || 0) / 2 - 6;
        out.push(
          <button key={`b${lvl}-${run.a}`} type="button" className="gout-btn" title={collapsed ? 'Показать детали' : 'Скрыть детали'}
            style={axis === 'row' ? { top: at, left: across } : { left: at, top: across }}
            onMouseDown={(e) => e.preventDefault()} onClick={() => onOutline(axis, collapsed ? { type: 'expand', ...run, level: lvl } : { type: 'collapse', ...run, level: lvl })}>
            {collapsed ? '+' : '−'}
          </button>,
        );
      }
    }
    return out;
  };
  const levelButtons = (axis) => {
    const levels = axis === 'row' ? rowLevels : colLevels;
    return Array.from({ length: levels + 1 }, (_, i) => (
      <button key={i} type="button" className="gout-btn level" title={`Показать уровень ${i + 1}`}
        style={axis === 'row' ? { left: i * OUT_STEP + 4 - 1, top: colGut + (HEAD_H - 13) / 2 } : { top: i * OUT_STEP + 4 - 1, left: rowGut + (HEAD_W - 13) / 2 }}
        onMouseDown={(e) => e.preventDefault()} onClick={() => onOutline(axis, { type: 'level', n: i + 1 })}>{i + 1}</button>
    ));
  };

  return (
    <div className="grid-wrap">
      <div ref={scrollRef} className="grid-scroll" onScroll={(e) => setView((v) => ({ ...v, top: e.currentTarget.scrollTop, left: e.currentTarget.scrollLeft }))}>
        <div className="grid-canvas" style={{ width: colPos[cols] + HW, height: rowPos[rows] + HH }}>
          {/* ячейки */}
          <div className="grid-cells" style={{ left: HW, top: HH, width: colPos[cols], height: rowPos[rows] }}
            onMouseDown={onMouseDown}
            onDoubleClick={(e) => { const p = cellAt(e.clientX, e.clientY); setEditing({ r: p.r, c: p.c, text: null }); }}>
            {gridLines(r0, r1, c0, c1)}
            {renderCells(r0, r1, c0, c1)}
            {marks}
            {arrowLayer}
            {errBadge}
            {editorIn(false, false)}
            {/* закреплённые строки сверху, столбцы слева и угол — поверх прокручиваемой части */}
            {fr > 0 && (
              <div className="grid-pane top" style={{ left: 0, top: view.top, width: colPos[cols], height: frozenH }}>
                {gridLines(0, fr - 1, c0, c1)}{renderCells(0, fr - 1, c0, c1)}{marks}{editorIn(true, false)}
              </div>
            )}
            {fc > 0 && (
              <div className="grid-pane left" style={{ left: view.left, top: 0, width: frozenW, height: rowPos[rows] }}>
                {gridLines(r0, r1, 0, fc - 1)}{renderCells(r0, r1, 0, fc - 1)}{marks}{editorIn(false, true)}
              </div>
            )}
            {fr > 0 && fc > 0 && (
              <div className="grid-pane top left" style={{ left: view.left, top: view.top, width: frozenW, height: frozenH }}>
                {gridLines(0, fr - 1, 0, fc - 1)}{renderCells(0, fr - 1, 0, fc - 1)}{marks}{editorIn(true, true)}
              </div>
            )}
          </div>
          {/* заголовки столбцов */}
          <div className="grid-colhead" style={{ top: view.top + colGut, left: HW, width: colPos[cols], height: HEAD_H }}>
            {colHeads(c0, c1)}
            {fc > 0 && <div className="grid-headpane" style={{ left: view.left, top: 0, width: frozenW, height: HEAD_H }}>{colHeads(0, fc - 1)}</div>}
          </div>
          {/* заголовки строк */}
          <div className="grid-rowhead" style={{ left: view.left + rowGut, top: HH, width: HEAD_W, height: rowPos[rows] }}>
            {rowHeads(r0, r1)}
            {fr > 0 && <div className="grid-headpane" style={{ left: 0, top: view.top, width: HEAD_W, height: frozenH }}>{rowHeads(0, fr - 1)}</div>}
          </div>
          {rowGut > 0 && <div className="grid-outline row" style={{ left: view.left, top: HH, width: rowGut, height: rowPos[rows] }}>{outline('row')}</div>}
          {colGut > 0 && <div className="grid-outline col" style={{ top: view.top, left: HW, height: colGut, width: colPos[cols] }}>{outline('col')}</div>}
          <div className="grid-corner" style={{ left: view.left, top: view.top, width: HW, height: HH }}>
            <div className="grid-corner-all" style={{ left: rowGut, top: colGut }} title="Выделить всё"
              onMouseDown={() => setSel({ r: 0, c: 0, r2: rows - 1, c2: cols - 1 })} />
            {rowGut > 0 && levelButtons('row')}
            {colGut > 0 && levelButtons('col')}
          </div>
        </div>
      </div>
    </div>
  );
}

// Редактор ячейки поверх сетки (Enter — вниз, Tab — вправо, Esc — отмена)
function CellEditor({ box, editing, setEditing, onCommit, style, names }) {
  const ref = useRef(null);
  const assist = useFormulaAssist(editing.text, names);
  // заменить текст и поставить курсор (подсказка, F4)
  const apply = ({ text, caret }) => {
    setEditing({ ...editing, text });
    assist.setCaret(caret);
    requestAnimationFrame(() => { const el = ref.current; if (el) { el.focus(); el.setSelectionRange(caret, caret); } });
  };
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    el.focus();
    const n = el.value.length;
    el.setSelectionRange(n, n);
  }, []);
  useEffect(() => {
    // вставка ссылки щелчком — курсор в конец
    const el = ref.current;
    if (el && document.activeElement !== el) { el.focus(); el.setSelectionRange(el.value.length, el.value.length); }
  }, [editing.text]);
  return (
    <>
    <textarea
      ref={ref}
      className="gedit"
      style={{ left: box.left, top: box.top, minWidth: box.width, height: box.height, fontWeight: style.fontWeight, fontStyle: style.fontStyle, fontFamily: style.fontFamily, fontSize: style.fontSize, color: style.color }}
      value={editing.text}
      spellCheck={false}
      onChange={(e) => { setEditing({ ...editing, text: e.target.value }); assist.setCaret(e.target.selectionStart); }}
      onSelect={(e) => assist.setCaret(e.currentTarget.selectionStart)}
      onKeyDown={(e) => {
        if (assist.onKey(e, apply)) { e.stopPropagation(); return; }
        if (e.key === 'F4') { e.preventDefault(); const r = cycleRef(editing.text, e.currentTarget.selectionStart); if (r) apply(r); e.stopPropagation(); return; }
        if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); onCommit(editing.text, e.shiftKey ? 'up' : 'down'); }
        else if (e.key === 'Enter' && e.altKey) { e.preventDefault(); setEditing({ ...editing, text: `${editing.text}\n` }); }
        else if (e.key === 'Tab') { e.preventDefault(); onCommit(editing.text, e.shiftKey ? 'tabBack' : 'tab'); }
        else if (e.key === 'Escape') { e.preventDefault(); setEditing(null); }
        e.stopPropagation();
      }}
    />
    <AssistPopup assist={assist} style={{ left: box.left, top: box.top + box.height + 2 }} onPick={(i) => { const r = assist.accept(i); if (r) apply(r); }} />
    </>
  );
}
