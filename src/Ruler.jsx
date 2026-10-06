// Горизонтальная линейка над страницей: деления в см от левого поля и маркеры отступов абзаца.
import React, { useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { effectivePara } from './editor/paraFormat.js';
import { PT_PER_CM, cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';
import { parseSectPr } from './editor/page.js';

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
    sectPr: e.state.doc.attrs.sectPr,
    active: !!p && !inList && !inTable,
    left: p ? p.indentLeft : 0,
    right: p ? p.indentRight : 0,
    first: p ? p.indentFirst : 0,
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
  const { PAGE_W, MARGIN_L, MARGIN_R, TEXT_W } = geometry(s.sectPr);
  const [drag, setDrag] = useState(null); // { kind, x } — x в pt от левого края текста
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

  // Позиции маркеров (во время перетаскивания — предварительные)
  let { left, right, first } = s;
  if (drag) {
    if (drag.kind === 'first') first = drag.x - left;
    else if (drag.kind === 'hanging') { first = left + first - drag.x; left = drag.x; }
    else if (drag.kind === 'left') left = drag.x;
    else right = TEXT_W - drag.x;
  }
  const at = (pt) => ({ left: `${((MARGIN_L + pt) / PAGE_W) * 100}%` });
  const tip = (title, pt) => `${title}: ${fmtNum(ptToCm(pt), 2)} см`;

  return (
    <div className="ruler-bar">
      <div className="ruler" ref={boxRef} style={{ width: `${(PAGE_W / 0.75) * scale}px` }}>
        <svg viewBox={`0 0 ${PAGE_W} 16`} aria-hidden="true">
          <rect x={0} y={2} width={MARGIN_L} height={12} className="ruler-margin" />
          <rect x={MARGIN_L} y={2} width={TEXT_W} height={12} className="ruler-text" />
          <rect x={PAGE_W - MARGIN_R} y={2} width={MARGIN_R} height={12} className="ruler-margin" />
          <Ticks PAGE_W={PAGE_W} MARGIN_L={MARGIN_L} />
        </svg>
        {s.active && (
          <>
            <div className="rm rm-first" style={at(left + first)} title={tip('Отступ первой строки', first)} onMouseDown={startDrag('first')} />
            <div className="rm rm-hanging" style={at(left)} title={tip('Выступ', -first)} onMouseDown={startDrag('hanging')} />
            <div className="rm rm-left" style={at(left)} title={tip('Отступ слева', left)} onMouseDown={startDrag('left')} />
            <div className="rm rm-right" style={at(TEXT_W - right)} title={tip('Отступ справа', right)} onMouseDown={startDrag('right')} />
          </>
        )}
        {drag && <div className="ruler-guide" style={at(drag.x)} />}
      </div>
    </div>
  );
}
