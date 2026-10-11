// Горизонтальная линейка над страницей: деления в см от левого поля, маркеры отступов абзаца и позиции табуляции.
// Табуляция как в Word: тип — кнопка в углу линейки; щелчок по линейке — новая позиция, перетаскивание — сдвиг,
// увод вниз с линейки — удалить; серые засечки — позиции по умолчанию.
import React, { useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { effectivePara } from './editor/paraFormat.js';
import { PT_PER_CM, cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';
import { parseSectPr, sectionAt } from './editor/page.js';
import { DEFAULT_TAB } from './editor/tabs.js';
import { setPage } from './PageControls.jsx';

const TAB_TYPES = [['left', 'По левому краю'], ['center', 'По центру'], ['right', 'По правому краю'], ['decimal', 'По разделителю'], ['bar', 'Вертикальная черта']];
// Значок позиции табуляции (как на линейке Word)
function TabMark({ align, size = 9 }) {
  const d = {
    left: 'M2 1v7h6', right: 'M7 1v7H1', center: 'M4.5 1v7M1 8h7', decimal: 'M4.5 1v7M1 8h7M6.5 4.5h.1', bar: 'M4.5 0v9',
  }[align] || 'M2 1v7h6';
  return <svg width={size} height={size} viewBox="0 0 9 9" aria-hidden="true"><path d={d} fill="none" stroke="currentColor" strokeWidth="1.6" /></svg>;
}

const SNAP = PT_PER_CM / 4;         // привязка к 0,25 см

// Геометрия страницы документа (из w:sectPr) — в pt
function geometry(sectPr) {
  const g = parseSectPr(sectPr);
  const PAGE_W = g.w / 20;
  const MARGIN_L = g.left / 20;
  const MARGIN_R = g.right / 20;
  return { PAGE_W, MARGIN_L, MARGIN_R, TEXT_W: PAGE_W - MARGIN_L - MARGIN_R };
}

function rulerState(e) {
  const { $from } = e.state.selection;
  let node = $from.parent;
  let inList = false;
  let inTable = false;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d).type.name;
    if (n === 'listItem') inList = true;
    if (n === 'tableCell' || n === 'tableHeader') inTable = true;
  }
  if (!['paragraph', 'heading'].includes(node.type.name)) node = null;
  const p = node ? effectivePara(node, e.state.doc.attrs, inList) : null;
  return {
    sectPr: sectionAt(e.state).sectPr, // раздел под курсором
    active: !!p && !inList && !inTable,
    left: p ? p.indentLeft : 0,
    right: p ? p.indentRight : 0,
    first: p ? p.indentFirst : 0,
    tabs: node ? node.attrs.tabs || [] : [],
    defaultTab: e.state.doc.attrs.defaultTab || DEFAULT_TAB,
  };
}

// Деления: каждые 0,25 см, цифры — каждый сантиметр (от левого поля в обе стороны, как в Word)
function Ticks({ PAGE_W, MARGIN_L }) {
  const items = [];
  const fromCm = -Math.floor(MARGIN_L / PT_PER_CM);
  const toCm = Math.floor((PAGE_W - MARGIN_L) / PT_PER_CM);
  for (let q = fromCm * 4; q <= toCm * 4; q++) {
    const x = MARGIN_L + (q / 4) * PT_PER_CM;
    if (x < 2 || x > PAGE_W - 2) continue;
    if (q % 4 === 0) {
      if (q !== 0) items.push(<text key={q} x={x} y={10.5} textAnchor="middle">{Math.abs(q / 4)}</text>);
    } else {
      const h = q % 2 === 0 ? 4 : 2;
      items.push(<line key={q} x1={x} x2={x} y1={8 - h / 2} y2={8 + h / 2} />);
    }
  }
  return <g className="ruler-ticks">{items}</g>;
}

export default function Ruler({ editor, zoom }) {
  const s = useEditorState({ editor, selector: ({ editor: e }) => rulerState(e) });
  const [drag, setDrag] = useState(null); // { kind, x } — x в pt от левого края текста; поля — { kind: 'marginL'|'marginR', m }
  const geo = geometry(s.sectPr);
  // во время перетаскивания поля — линейка уже с новым полем (деления, маркеры и серая зона сдвигаются)
  const MARGIN_L = drag && drag.kind === 'marginL' ? drag.m : geo.MARGIN_L;
  const MARGIN_R = drag && drag.kind === 'marginR' ? drag.m : geo.MARGIN_R;
  const { PAGE_W } = geo;
  const TEXT_W = PAGE_W - MARGIN_L - MARGIN_R;
  const [tabType, setTabType] = useState('left');
  const boxRef = useRef(null);
  const scale = zoom / 100;

  const startDrag = (kind) => (ev) => {
    ev.preventDefault();
    const rect = boxRef.current.getBoundingClientRect();
    // Привязка к сетке 0,25 см: левые маркеры — от левого поля, правый — от правого (Alt — без привязки)
    const snap = (v, alt) => (alt ? v : Math.round(v / SNAP) * SNAP);
    const toPt = (clientX, alt) => {
      const pt = (clientX - rect.left) / (rect.width / PAGE_W) - MARGIN_L; // ширина в px уже с учётом zoom
      return kind === 'right' ? TEXT_W - snap(TEXT_W - pt, alt) : snap(pt, alt);
    };
    const pos0 = { first: s.left + s.first, hanging: s.left, left: s.left, right: TEXT_W - s.right }[kind];
    let cur = pos0;
    setDrag({ kind, x: cur });
    const move = (e) => {
      cur = Math.min(Math.max(toPt(e.clientX, e.altKey), -MARGIN_L), PAGE_W - MARGIN_L);
      setDrag({ kind, x: cur });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (Math.abs(cur - pos0) < 0.01) { editor.commands.focus(); return; }
      const r = (v) => cmToPt(ptToCm(v)); // округление до твипа
      let patch;
      if (kind === 'first') patch = { indentFirst: r(cur - s.left) };
      else if (kind === 'hanging') patch = { indentLeft: r(cur), indentFirst: r(s.left + s.first - cur) }; // первая строка остаётся на месте
      else if (kind === 'left') patch = { indentLeft: r(cur) };                                             // двигаются обе
      else patch = { indentRight: r(TEXT_W - cur) };
      editor.chain().focus().setParagraphFormat(patch).run();
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  // ---------- Поля страницы: граница серой и белой зоны (как в Word; маркеры отступов поверх неё — у них приоритет) ----------
  const startMarginDrag = (kind) => (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const rect = boxRef.current.getBoundingClientRect();
    const m0 = kind === 'marginL' ? geo.MARGIN_L : geo.MARGIN_R;
    const other = kind === 'marginL' ? geo.MARGIN_R : geo.MARGIN_L;
    let m = m0;
    setDrag({ kind, m });
    const move = (e) => {
      const x = (e.clientX - rect.left) / (rect.width / PAGE_W); // pt от левого края листа
      const raw = kind === 'marginL' ? x : PAGE_W - x;
      const snapped = e.altKey ? raw : Math.round(raw / SNAP) * SNAP;
      m = Math.min(Math.max(snapped, 0), PAGE_W - other - PT_PER_CM); // текста — не меньше 1 см
      setDrag({ kind, m });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (Math.abs(m - m0) < 0.01) { editor.commands.focus(); return; }
      const twips = Math.round(m * 20);
      setPage(editor, kind === 'marginL' ? { left: twips } : { right: twips });
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  // ---------- Табуляция ----------
  const ptAt = (clientX, alt) => {
    const rect = boxRef.current.getBoundingClientRect();
    const pt = (clientX - rect.left) / (rect.width / PAGE_W) - MARGIN_L;
    return alt ? pt : Math.round(pt / SNAP) * SNAP;
  };
  const r2 = (v) => cmToPt(ptToCm(v));
  // щелчок по свободному месту линейки — новая позиция выбранного типа
  const addTab = (ev) => {
    if (!s.active || ev.button !== 0 || ev.target.closest('.rm, .rt')) return;
    ev.preventDefault();
    const pt = ptAt(ev.clientX, ev.altKey);
    if (pt <= 0 || pt > TEXT_W + 0.5) return;
    editor.chain().focus().addTabStop({ pos: r2(pt), align: tabType }).run();
  };
  const startTabDrag = (stop) => (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const rect = boxRef.current.getBoundingClientRect();
    let cur = stop.pos;
    let off = false;
    setDrag({ kind: 'tab', x: cur, align: stop.align, from: stop.pos });
    const move = (e) => {
      off = e.clientY > rect.bottom + 18 || e.clientY < rect.top - 18; // увели с линейки — удалить
      cur = Math.min(Math.max(ptAt(e.clientX, e.altKey), 0), TEXT_W);
      setDrag({ kind: 'tab', x: cur, align: stop.align, from: stop.pos, off });
    };
    const up = () => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setDrag(null);
      if (off) editor.chain().focus().removeTabStop(stop.pos).run();
      else if (Math.abs(cur - stop.pos) > 0.01) editor.chain().focus().moveTabStop(stop.pos, r2(cur)).run();
      else editor.commands.focus();
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };
  const lastCustom = s.tabs.reduce((m, t) => Math.max(m, t.pos), 0);
  const defaults = [];
  for (let x = s.defaultTab; x < TEXT_W; x += s.defaultTab) if (x > lastCustom + 0.3) defaults.push(x);
  const typeIdx = TAB_TYPES.findIndex(([t]) => t === tabType);

  // Позиции маркеров (во время перетаскивания — предварительные)
  let { left, right, first } = s;
  if (drag && !drag.kind.startsWith('margin')) {
    if (drag.kind === 'first') first = drag.x - left;
    else if (drag.kind === 'hanging') { first = left + first - drag.x; left = drag.x; }
    else if (drag.kind === 'left') left = drag.x;
    else right = TEXT_W - drag.x;
  }
  const at = (pt) => ({ left: `${((MARGIN_L + pt) / PAGE_W) * 100}%` });
  const tip = (title, pt) => `${title}: ${fmtNum(ptToCm(pt), 2)} см`;

  return (
    <div className="ruler-bar">
      <div className="ruler" ref={boxRef} style={{ width: `${(PAGE_W / 0.75) * scale}px` }} onMouseDown={addTab}>
        <button type="button" className="ruler-tabtype" title={`Тип табуляции: ${TAB_TYPES[typeIdx][1]} (щёлкните, чтобы сменить; щелчок по линейке — поставить)`}
          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
          onClick={() => setTabType(TAB_TYPES[(typeIdx + 1) % TAB_TYPES.length][0])}>
          <TabMark align={tabType} size={11} />
        </button>
        <svg viewBox={`0 0 ${PAGE_W} 16`} aria-hidden="true">
          <rect x={0} y={2} width={MARGIN_L} height={12} className="ruler-margin" />
          <rect x={MARGIN_L} y={2} width={TEXT_W} height={12} className="ruler-text" />
          <rect x={PAGE_W - MARGIN_R} y={2} width={MARGIN_R} height={12} className="ruler-margin" />
          <Ticks PAGE_W={PAGE_W} MARGIN_L={MARGIN_L} />
        </svg>
        <div className="rmargin" style={{ left: `${(MARGIN_L / PAGE_W) * 100}%` }} title={`Левое поле: ${fmtNum(ptToCm(MARGIN_L), 2)} см (перетащите, чтобы изменить)`} onMouseDown={startMarginDrag('marginL')} />
        <div className="rmargin" style={{ left: `${((PAGE_W - MARGIN_R) / PAGE_W) * 100}%` }} title={`Правое поле: ${fmtNum(ptToCm(MARGIN_R), 2)} см (перетащите, чтобы изменить)`} onMouseDown={startMarginDrag('marginR')} />
        {s.active && (
          <>
            <div className="rm rm-first" style={at(left + first)} title={tip('Отступ первой строки', first)} onMouseDown={startDrag('first')} />
            <div className="rm rm-hanging" style={at(left)} title={tip('Выступ', -first)} onMouseDown={startDrag('hanging')} />
            <div className="rm rm-left" style={at(left)} title={tip('Отступ слева', left)} onMouseDown={startDrag('left')} />
            <div className="rm rm-right" style={at(TEXT_W - right)} title={tip('Отступ справа', right)} onMouseDown={startDrag('right')} />
            {defaults.map((x) => <div key={`d${x}`} className="rt-default" style={at(x)} />)}
            {s.tabs.filter((t) => !(drag && drag.kind === 'tab' && Math.abs(t.pos - (drag.from ?? -1)) < 0.01)).map((t) => (
              <div key={`t${t.pos}`} className={`rt rt-${t.align}`} style={at(t.pos)} title={`${tip(TAB_TYPES.find(([a]) => a === t.align)?.[1] || 'Табуляция', t.pos)}. Перетащите вниз, чтобы удалить`}
                onMouseDown={startTabDrag(t)}>
                <TabMark align={t.align} />
              </div>
            ))}
          </>
        )}
        {drag && drag.kind === 'tab' && !drag.off && <div className={`rt rt-${drag.align}`} style={at(drag.x)}><TabMark align={drag.align} /></div>}
        {drag && !drag.kind.startsWith('margin') && <div className={`ruler-guide${drag.off ? ' off' : ''}`} style={at(drag.x)} />}
        {drag && drag.kind.startsWith('margin') && <div className="ruler-guide" style={{ left: `${((drag.kind === 'marginL' ? drag.m : PAGE_W - drag.m) / PAGE_W) * 100}%` }} />}
      </div>
    </div>
  );
}
