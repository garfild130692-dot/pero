// «Макет → Абзац»: счётчики отступов (см) и интервалов (пт) прямо на ленте, как в Word.
// Значение — действующее у абзацев выделения (разные — пустое поле); ввод — Enter/уход из поля, стрелки — шаг.
import React, { useEffect, useState } from 'react';
import { effectivePara } from './editor/paraFormat.js';
import { cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';
import { Icon } from './icons.jsx';

const TYPES = ['paragraph', 'heading'];

// действующие значения ключа у абзацев выделения; разные — null
export function paraValue(state, key) {
  const { from, to } = state.selection;
  let v; let mixed = false;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!TYPES.includes(node.type.name)) return true;
    const $p = state.doc.resolve(pos + 1);
    let inList = false;
    for (let d = $p.depth; d > 0; d--) if ($p.node(d).type.name === 'listItem') inList = true;
    const x = effectivePara(node, state.doc.attrs, inList)[key] || 0;
    if (v === undefined) v = x; else if (Math.abs(v - x) > 0.01) mixed = true;
    return false;
  });
  return mixed || v === undefined ? null : v;
}

const parse = (s) => {
  const v = parseFloat(String(s).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isFinite(v) ? v : null;
};

// единица: см — значение хранится в pt, шаг 0,1 см; пт — шаг 6 пт (как в Word)
const UNITS = {
  cm: { show: (pt) => `${fmtNum(ptToCm(pt), 2)} см`, read: (v) => cmToPt(v), step: 0.1, toUnit: ptToCm, min: -55.87, max: 55.87 },
  pt: { show: (pt) => `${fmtNum(pt, 1)} пт`, read: (v) => v, step: 6, toUnit: (x) => x, min: 0, max: 1584 },
};

function Spinner({ editor, attr, unit, icon, label, title, min }) {
  const u = UNITS[unit];
  const value = paraValue(editor.state, attr);
  const [text, setText] = useState(value == null ? '' : u.show(value));
  const [editing, setEditing] = useState(false);
  useEffect(() => { if (!editing) setText(value == null ? '' : u.show(value)); }, [value, editing, u]);
  const apply = (unitValue) => {
    const lo = min ?? u.min;
    const n = Math.min(u.max, Math.max(lo, unitValue));
    editor.chain().focus().setParagraphFormat({ [attr]: Math.round(u.read(n) * 100) / 100 }).run();
  };
  const commit = () => {
    setEditing(false);
    const v = parse(text);
    if (v == null) { setText(value == null ? '' : u.show(value)); return; }
    apply(v);
  };
  const bump = (dir) => {
    const cur = value == null ? 0 : u.toUnit(value);
    // к ближайшему делению шага, как счётчик Word
    const k = Math.round(cur / u.step);
    const next = Math.abs(cur - k * u.step) < 1e-6 ? (k + dir) * u.step : (dir > 0 ? Math.ceil(cur / u.step) : Math.floor(cur / u.step)) * u.step;
    apply(Math.round(next * 100) / 100);
  };
  return (
    <label className="rb-spin" title={title}>
      <span className="rb-spin-label">{icon}{label}</span>
      <span className="rb-spin-box">
        <input
          value={text}
          onFocus={(e) => { setEditing(true); e.target.select(); }}
          onChange={(e) => setText(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Enter') { e.preventDefault(); e.currentTarget.blur(); }
            else if (e.key === 'Escape') { setEditing(false); setText(value == null ? '' : u.show(value)); editor.commands.focus(); }
            else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); setEditing(false); bump(e.key === 'ArrowUp' ? 1 : -1); }
          }}
        />
        <span className="rb-spin-btns">
          <button type="button" tabIndex={-1} aria-label="Больше" onMouseDown={(e) => { e.preventDefault(); bump(1); }}>▴</button>
          <button type="button" tabIndex={-1} aria-label="Меньше" onMouseDown={(e) => { e.preventDefault(); bump(-1); }}>▾</button>
        </span>
      </span>
    </label>
  );
}

export function LayoutParaGroup({ editor, Group, onParagraph }) {
  return (
    <Group title="Абзац" launcher={onParagraph ? { title: 'Параметры абзаца…', onClick: onParagraph } : null}>
      <div className="rb-spin-cols">
        <div className="rb-spin-col">
          <div className="rb-spin-head">Отступ</div>
          <Spinner editor={editor} attr="indentLeft" unit="cm" icon={<Icon.indLeft size={14} />} label="Слева:" title="Отступ слева" />
          <Spinner editor={editor} attr="indentRight" unit="cm" icon={<Icon.indRight size={14} />} label="Справа:" title="Отступ справа" />
        </div>
        <div className="rb-spin-col">
          <div className="rb-spin-head">Интервал</div>
          <Spinner editor={editor} attr="spaceBefore" unit="pt" icon={<Icon.spBefore size={14} />} label="До:" title="Интервал перед абзацем" />
          <Spinner editor={editor} attr="spaceAfter" unit="pt" icon={<Icon.spAfter size={14} />} label="После:" title="Интервал после абзаца" />
        </div>
      </div>
    </Group>
  );
}
