// Вкладка «Таблица»: границы, заливка, выравнивание в ячейке, строки, положение таблицы.
import React, { useEffect, useState } from 'react';
import { selectedRect } from '@tiptap/pm/tables';
import { ColorPalette, SplitColor, usePopover } from './FontControls.jsx';
import { BORDER_STYLES, BORDER_WIDTHS, DEFAULT_PEN, borderCss } from './editor/table.js';
import { cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';

// ---------- Состояние таблицы под курсором ----------
export function tableState(e) {
  let rect;
  try { rect = selectedRect(e.state); } catch { return null; }
  const { table, map, top, left } = rect;
  const cell = table.nodeAt(map.map[top * map.width + left]);
  const row = table.child(top);
  return {
    background: cell.attrs.background || '',
    verticalAlign: cell.attrs.verticalAlign || 'top',
    height: row.attrs.height,
    heightRule: row.attrs.heightRule,
    repeatHeader: !!row.attrs.repeatHeader,
    align: table.attrs.align || 'left',
  };
}

// ---------- Значки ----------
const S = ({ children, size = 18 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{children}</svg>
);
// Квадрат 2×2 ячейки: бледная сетка + выделенные линии режима
const LINES = { top: 'M4 4h16', bottom: 'M4 20h16', left: 'M4 4v16', right: 'M20 4v16', insideH: 'M4 12h16', insideV: 'M12 4v16' };
const MODE_LINES = {
  all: ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'], outer: ['top', 'bottom', 'left', 'right'],
  inner: ['insideH', 'insideV'], none: [], top: ['top'], bottom: ['bottom'], left: ['left'], right: ['right'],
  insideH: ['insideH'], insideV: ['insideV'],
};
function BorderIcon({ mode, size = 18 }) {
  const on = MODE_LINES[mode];
  return (
    <S size={size}>
      {Object.entries(LINES).filter(([k]) => !on.includes(k)).map(([k, d]) => <path key={k} d={d} strokeWidth="1" strokeDasharray="1.5 2" opacity=".45" />)}
      {on.map((k) => <path key={k} d={LINES[k]} strokeWidth="2.2" />)}
    </S>
  );
}
const BucketIcon = () => (
  <svg width="16" height="14" viewBox="0 0 24 20" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" aria-hidden="true">
    <path d="m4 9 7-7 8 8-7 7z" /><path d="M8 5 4 1" /><path d="M20 12s2 2.5 2 4a2 2 0 0 1-4 0c0-1.5 2-4 2-4z" />
  </svg>
);
const VAlignIcon = ({ v }) => (
  <S><rect x="4" y="3" width="16" height="18" rx="1" strokeWidth="1.3" /><path d={{ top: 'M8 7h8M8 10h8', center: 'M8 10.5h8M8 13.5h8', bottom: 'M8 14h8M8 17h8' }[v]} /></S>
);
const PosIcon = ({ a }) => {
  const x = { left: 3, center: 7, right: 11 }[a];
  return <S><path d="M2 3h20M2 21h20" strokeWidth="1.2" /><rect x={x} y="7" width="10" height="10" /><path d={`M${x} 12h10M${x + 5} 7v10`} strokeWidth="1" /></S>;
};
const DistributeIcon = () => <S><rect x="3" y="5" width="18" height="14" /><path d="M9 5v14M15 5v14" /><path d="m5 12 2-1.5v3zM19 12l-2-1.5v3z" fill="currentColor" strokeWidth="1" /></S>;
const HeaderIcon = () => <S><rect x="3" y="4" width="18" height="16" /><path d="M3 9h18" /><path d="M3 4h18v5H3z" fill="currentColor" opacity=".35" /><path d="M10 13l2 2 2-2" /></S>;

const BORDER_MODES = [
  ['bottom', 'Нижняя граница'], ['top', 'Верхняя граница'], ['left', 'Левая граница'], ['right', 'Правая граница'],
  null,
  ['none', 'Нет границы'], ['all', 'Все границы'], ['outer', 'Внешние границы'], ['inner', 'Внутренние границы'],
  null,
  ['insideH', 'Внутренняя горизонтальная граница'], ['insideV', 'Внутренняя вертикальная граница'],
];
const PEN_COLORS = [null, '#7f7f7f', '#1f6f5c', '#2e75b6', '#c00000', '#ffc000', '#00b050', '#7030a0'];

// ---------- Границы ----------
export function BordersButton({ editor }) {
  const pop = usePopover();
  const [mode, setMode] = useState('all');
  const [pen, setPen] = useState(DEFAULT_PEN);
  const apply = (m) => { setMode(m); editor.chain().focus().setCellBorders(m, pen).run(); };
  const label = (BORDER_MODES.find((x) => x && x[0] === mode) || [])[1];
  return (
    <div className="split" ref={pop.ref}>
      <button type="button" className="rb-btn split-main" title={`Границы: ${label}`} onMouseDown={(e) => e.preventDefault()} onClick={() => apply(mode)}>
        <BorderIcon mode={mode} />
      </button>
      <button type="button" className="rb-btn split-arrow" title="Границы: выбрать" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop borders-pop" style={pop.style} role="menu">
          {BORDER_MODES.map((m, i) => (m ? (
            <button key={m[0]} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); apply(m[0]); }}>
              <BorderIcon mode={m[0]} size={16} />{m[1]}
            </button>
          ) : <div key={i} className="menu-sep" />))}
          <div className="menu-sep" />
          <div className="pen">
            <div className="pen-preview" style={{ borderTop: borderCss(pen) }} />
            <label>Тип линии
              <select value={pen.style} onChange={(e) => setPen({ ...pen, style: e.target.value })}>
                {BORDER_STYLES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
            </label>
            <label>Толщина
              <select value={pen.width} onChange={(e) => setPen({ ...pen, width: +e.target.value })}>
                {BORDER_WIDTHS.map((w) => <option key={w} value={w}>{fmtNum(w, 2)} пт</option>)}
              </select>
            </label>
            <div className="pen-colors">
              Цвет
              {PEN_COLORS.map((c) => (
                <button key={c || 'auto'} type="button" className={`swatch${(pen.color || null) === c ? ' current' : ''}`} title={c || 'Авто'}
                  style={{ background: c || '#000' }} onMouseDown={(e) => e.preventDefault()} onClick={() => setPen({ ...pen, color: c })} />
              ))}
              <input type="color" className="pen-custom" title="Другой цвет" value={pen.color || '#000000'} onChange={(e) => setPen({ ...pen, color: e.target.value })} />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Заливка ----------
export function ShadingButton({ editor }) {
  const [last, setLast] = useState('#dcefe9');
  const apply = (c) => { setLast(c); editor.chain().focus().setCellAttribute('background', c).run(); };
  return (
    <SplitColor title="Заливка ячеек" color={last} icon={<BucketIcon />}>
      {{
        apply: () => apply(last),
        palette: (close) => (
          <ColorPalette autoLabel="Нет цвета" autoClass="none" current={last}
            onAuto={() => { close(); editor.chain().focus().setCellAttribute('background', null).run(); }}
            onPick={(x) => { close(); apply(x); }} />
        ),
      }}
    </SplitColor>
  );
}

// ---------- Высота строки ----------
export function RowHeightField({ editor, t }) {
  const shown = t.height ? fmtNum(ptToCm(t.height), 2) : '';
  const [text, setText] = useState(shown);
  useEffect(() => setText(shown), [shown]);
  const commit = () => {
    const v = parseFloat(text.replace(',', '.'));
    if (!text.trim()) editor.chain().focus().setRowHeight(null).run();
    else if (isFinite(v) && v > 0 && v <= 55) editor.chain().focus().setRowHeight(cmToPt(v), t.heightRule).run();
    else setText(shown);
  };
  return (
    <div className="rb-field">
      <span>Высота</span>
      <span className="pd-input small">
        <input value={text} placeholder="авто" aria-label="Высота строки, см" title="Высота строки (минимум), см; пусто — авто"
          onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); commit(); } if (e.key === 'Escape') setText(shown); }}
          onBlur={() => setText(shown)} />
        <em>см</em>
      </span>
    </div>
  );
}

export const TableIcons = { VAlignIcon, PosIcon, DistributeIcon, HeaderIcon };
