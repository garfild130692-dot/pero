// Вертикальная линейка слева от листа (как в Word): показывает лист, на котором курсор, — серым верхнее и нижнее
// поля, деления 0,25 см и сантиметры от верхнего поля; границы полей перетаскиваются мышью (привязка 0,25 см, Alt — без).
// Стоит в слое листа (.page) на высоте текущего листа: его верх берётся из полос между листами (pagination.js).
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { pageInfo, pageTops } from './editor/pagination.js';

export { pageTops };
import { parseSectPr, sectionAt } from './editor/page.js';
import { PT_PER_CM, fmtNum, ptToCm } from './editor/paragraph.js';
import { setPage } from './PageControls.jsx';

const SNAP = PT_PER_CM / 4;
const MIN_TEXT = PT_PER_CM; // между полями — не меньше 1 см текста

// Верх n-го листа (см. pageTops)
export function pageTopOf(pageEl, n) {
  if (!pageEl || n <= 1) return 0;
  return pageTops(pageEl)[n - 1] ?? 0;
}

// Деления: каждые 0,25 см, цифры — каждый сантиметр от верхнего поля (вверх и вниз, как в Word)
function VTicks({ H, T }) { // H, T — pt; рисуем в px (1 pt = 4/3 px)
  const items = [];
  for (let q = -Math.floor(T / PT_PER_CM) * 4; q <= Math.floor((H - T) / PT_PER_CM) * 4; q++) {
    const yp = T + (q / 4) * PT_PER_CM;
    if (yp < 2 || yp > H - 2) continue;
    const y = yp / 0.75;
    if (q % 4 === 0) {
      if (q !== 0) items.push(<text key={q} x={8} y={y + 2.3} textAnchor="middle">{Math.abs(q / 4)}</text>);
    } else {
      const w = q % 2 === 0 ? 4 : 2;
      items.push(<line key={q} y1={y} y2={y} x1={8 - w / 2} x2={8 + w / 2} />);
    }
  }
  return <g className="ruler-ticks">{items}</g>;
}

export default function VRuler({ editor }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({ page: pageInfo(e.state).page, sectPr: sectionAt(e.state).sectPr }),
  });
  const ref = useRef(null);
  const [top, setTop] = useState(0);
  const [drag, setDrag] = useState(null); // { kind: 'top' | 'bottom', y } — y в pt от верха листа
  const g = parseSectPr(s.sectPr);
  const H = g.h / 20;
  const T = g.top / 20;
  const B = g.bottom / 20;

  const place = useCallback(() => {
    const el = ref.current;
    if (el) setTop(pageTopOf(el.closest('.page'), s.page));
  }, [s.page]);
  useLayoutEffect(place, [place, s.sectPr]);
  useEffect(() => {
    document.addEventListener('pero-paginated', place);
    return () => document.removeEventListener('pero-paginated', place);
  }, [place]);

  const startDrag = (kind) => (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const rect = ref.current.getBoundingClientRect();
    const toPt = (clientY, alt) => {
      const y = ((clientY - rect.top) / rect.height) * H;
      return alt ? y : Math.round(y / SNAP) * SNAP;
    };
    const y0 = kind === 'top' ? T : H - B;
    let cur = y0;
    setDrag({ kind, y: cur });
    const move = (e) => {
      const y = toPt(e.clientY, e.altKey);
      cur = kind === 'top' ? Math.min(Math.max(y, 0), H - B - MIN_TEXT) : Math.max(Math.min(y, H), T + MIN_TEXT);
      setDrag({ kind, y: cur });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (Math.abs(cur - y0) < 0.01) { editor.commands.focus(); return; }
      const twips = Math.round((kind === 'top' ? cur : H - cur) * 20);
      setPage(editor, kind === 'top' ? { top: twips } : { bottom: twips });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  let tm = T;
  let bm = B;
  if (drag) { if (drag.kind === 'top') tm = drag.y; else bm = H - drag.y; }
  const at = (y) => ({ top: `${(y / H) * 100}%` });
  const cmTip = (title, v) => `${title}: ${fmtNum(ptToCm(v), 2)} см (перетащите, чтобы изменить)`;

  return (
    <div className="vruler" ref={ref} style={{ top, height: `${H / 0.75}px` }} onMouseDown={(e) => e.preventDefault()}>
      <svg viewBox={`0 0 16 ${H / 0.75}`} aria-hidden="true">
        <rect x={2} y={0} width={12} height={tm / 0.75} className="ruler-margin" />
        <rect x={2} y={tm / 0.75} width={12} height={(H - tm - bm) / 0.75} className="ruler-text" />
        <rect x={2} y={(H - bm) / 0.75} width={12} height={bm / 0.75} className="ruler-margin" />
        <VTicks H={H} T={tm} />
      </svg>
      <div className="vm" style={at(tm)} title={cmTip('Верхнее поле', tm)} onMouseDown={startDrag('top')} />
      <div className="vm" style={at(H - bm)} title={cmTip('Нижнее поле', bm)} onMouseDown={startDrag('bottom')} />
      {drag && <div className="vruler-guide" style={at(drag.y)} />}
    </div>
  );
}
