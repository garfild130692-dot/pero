// Сетка листа: виртуализированная отрисовка видимых ячеек, заголовки строк/столбцов, выделение,
// изменение ширины столбцов и высоты строк мышью, редактор ячейки поверх сетки.
import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { colName, key } from './address.js';
import { formatValue } from './format.js';
import { isErr } from './formula.js';
import { DEFAULT_COL_W, DEFAULT_ROW_H, isHidden, maxLevel, mergeAt, outlineRuns } from './model.js';
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
    color: s.color || undefined,
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

export default function Grid({ sheet, sel, setSel, editing, setEditing, onCommit, onColWidth, onRowHeight, onRefInsert, onOutline, gridRef }) {
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
  const isEmpty = (r, c) => { const x = sheet.cells.get(key(r, c)); return !x || x.v == null || x.v === ''; };

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
        const cell = sheet.cells.get(key(rr, cc));
        if (!cell && !m) continue;
        let text = cellText(cell);
        if (!text && !(cell && cell.s && (cell.s.fill || cell.s.border)) && !m) continue;
        const own = colPos[c2 + 1] - colPos[cc];
        if (own <= 0 && !m) continue; // скрытый столбец
        const st = cellStyle(cell);
        const box = { left: colPos[cc], top: rowPos[rr], width: own, height: rowPos[r2 + 1] - rowPos[rr] };
        const s0 = (cell && cell.s) || {};
        if (text && typeof cell.v === 'number' && !s0.wrap) text = fitNumber(cell, text, own - 3, st);
        // перетекание текста вправо
        let spill = own;
        if (text && !m && typeof cell.v === 'string' && !s0.wrap && st.textAlign === 'left') {
          const need = textWidth(text, st) + 6;
          let cn = c2 + 1;
          while (spill < need && cn < cols && isEmpty(rr, cn) && !(sheet.merges.length && mergeAt(sheet, rr, cn))) { spill += colPos[cn + 1] - colPos[cn]; cn++; }
        }
        if (c < ca && spill === own) continue;
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
    </>
  );
  // редактор — в той части листа (закреплённой или нет), где ячейка
  const editorIn = (top, left) => (editing && editing.text !== null && !editing.fromBar && (editing.r < fr) === top && (editing.c < fc) === left
    ? <CellEditor key={`${editing.r},${editing.c}`} box={ed} editing={editing} setEditing={setEditing} onCommit={onCommit} style={cellStyle(sheet.cells.get(key(editing.r, editing.c)))} />
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
  const rowHeads = (ra, rb) => Array.from({ length: rb - ra + 1 }, (_, i) => {
    const r = ra + i;
    const h = rowPos[r + 1] - rowPos[r];
    if (!h) return null;
    const on = r >= selR1 && r <= selR2;
    return (
      <div key={r} className={`ghead row${on ? ' on' : ''}`} style={{ top: rowPos[r], height: h }}
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
function CellEditor({ box, editing, setEditing, onCommit, style }) {
  const ref = useRef(null);
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
    <textarea
      ref={ref}
      className="gedit"
      style={{ left: box.left, top: box.top, minWidth: box.width, height: box.height, fontWeight: style.fontWeight, fontStyle: style.fontStyle, fontFamily: style.fontFamily, fontSize: style.fontSize, color: style.color }}
      value={editing.text}
      spellCheck={false}
      onChange={(e) => setEditing({ ...editing, text: e.target.value })}
      onKeyDown={(e) => {
        if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); onCommit(editing.text, e.shiftKey ? 'up' : 'down'); }
        else if (e.key === 'Enter' && e.altKey) { e.preventDefault(); setEditing({ ...editing, text: `${editing.text}\n` }); }
        else if (e.key === 'Tab') { e.preventDefault(); onCommit(editing.text, e.shiftKey ? 'left' : 'right'); }
        else if (e.key === 'Escape') { e.preventDefault(); setEditing(null); }
        e.stopPropagation();
      }}
    />
  );
}
