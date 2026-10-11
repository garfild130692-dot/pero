// «Многоуровневый список» (группа «Абзац»): нумерация заголовков готовыми схемами, свой многоуровневый список
// (диалог по уровням), формат нумерованного списка под курсором — как в Word.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { HEADING_SCHEMES, LIST_PATTERNS, NUM_FORMATS, levelLabel, makeHeadingNumbering } from './editor/numbering.js';

const sample = (levels, upto = 3) => {
  const nums = Array(9).fill(0).map((_, i) => levels[i].start || 1);
  return levels.slice(0, upto).map((l, i) => (l.on ? levelLabel(levels, i, nums) : '—')).join('  ·  ');
};

export function MultilevelButton({ editor, Btn, state, onDialog }) {
  const pop = usePopover();
  const c = () => editor.chain().focus();
  const cur = state.headingNumbering;
  const item = (label, fn, checked, title) => (
    <button type="button" className="menu-item" title={title} onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}>
      <span className="check">{checked ? '✓' : ''}</span>{label}
    </button>
  );
  const active = cur && !cur.off;
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn icon="multilevel" title="Многоуровневый список: нумерация заголовков и формат списка" active={pop.open} onClick={() => pop.setOpen(!pop.open)} />
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          <div className="pop-title">Нумерация заголовков</div>
          {item('Нет', () => c().setHeadingNumbering(null).run(), !active)}
          {HEADING_SCHEMES.map((s) => item(s.label, () => c().setHeadingNumbering(s.key).run(), active && sample(cur.levels) === sample(s.make()), sample(s.make())))}
          {item('Определить свой многоуровневый список…', () => onDialog('multilevel'))}
          {state.inHeading && <div className="menu-sep" />}
          {state.inHeading && item(state.headingNoNumber ? 'Нумеровать этот заголовок' : 'Не нумеровать этот заголовок', () => c().toggleHeadingNumber().run())}
          <div className="menu-sep" />
          {item('Формат нумерованного списка…', () => onDialog('listFormat'), false, state.inOrdered ? '' : 'Поставьте курсор в нумерованный список')}
        </div>
      )}
    </div>
  );
}

// «Определение многоуровневого списка» — для нумерации заголовков: по уровням формат номера, стиль, начать с
export function MultilevelDialog({ editor, onClose }) {
  const [levels, setLevels] = useState(() => {
    const hn = editor.state.doc.attrs.headingNumbering;
    return (hn && !hn.off ? hn.levels : makeHeadingNumbering('dots').levels).map((l) => ({ ...l }));
  });
  const [lv, setLv] = useState(0);
  const set = (k, v) => setLevels((ls) => ls.map((l, i) => (i === lv ? { ...l, [k]: v } : l)));
  const L = levels[lv];
  const nums = Array(9).fill(0).map((_, i) => levels[i].start || 1);
  const apply = () => { editor.chain().focus().setHeadingNumbering(null, levels).run(); onClose(); };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd ml-dlg" role="dialog" aria-label="Многоуровневый список" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Определение многоуровневого списка (нумерация заголовков)</div>
        <div className="modal-body ml-body">
          <div className="ml-levels" role="listbox">
            {levels.slice(0, 6).map((l, i) => (
              <div key={i} role="option" aria-selected={i === lv} className={`sym-row bm-row${i === lv ? ' sel' : ''}`} onClick={() => setLv(i)}>
                <span>Уровень {i + 1}</span>
              </div>
            ))}
          </div>
          <div className="ml-edit">
            <label className="pd-check"><input type="checkbox" checked={!!L.on} onChange={(e) => set('on', e.target.checked)} /> Нумеровать «Заголовок {lv + 1}»</label>
            <label className="fd-col"><span>Формат номера (%1 — номер уровня 1, %2 — уровня 2…)</span>
              <input value={L.text} onChange={(e) => set('text', e.target.value)} disabled={!L.on} />
            </label>
            <div className="sym-top">
              <label className="fd-col"><span>Нумерация для этого уровня</span>
                <select value={L.fmt} onChange={(e) => set('fmt', e.target.value)} disabled={!L.on}>
                  {NUM_FORMATS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}
                </select>
              </label>
              <label className="fd-col"><span>Начать с</span>
                <input type="number" min="0" max="999" value={L.start} onChange={(e) => set('start', Math.max(0, +e.target.value || 0))} disabled={!L.on} />
              </label>
            </div>
            <div className="pop-title">Образец</div>
            <div className="ml-preview">
              {levels.slice(0, 4).map((l, i) => (
                <div key={i} style={{ paddingLeft: i * 14, fontWeight: i === lv ? 600 : 400 }}>{l.on ? levelLabel(levels, i, nums) : ''} Заголовок {i + 1}</div>
              ))}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// Формат нумерованного списка под курсором: стиль номера, вид (1. 1) (1) 1.1.), начать с
export function ListFormatDialog({ editor, onClose }) {
  const [init] = useState(() => {
    const $f = editor.state.selection.$from;
    for (let d = $f.depth; d > 0; d--) {
      const n = $f.node(d);
      if (n.type.name === 'orderedList') {
        let depth = 0;
        for (let k = d - 1; k > 0; k--) if ($f.node(k).type.name === 'orderedList') depth++;
        return { fmt: n.attrs.numFmt || ['decimal', 'lowerLetter', 'lowerRoman'][depth % 3], text: n.attrs.lvlText || '%1.', start: n.attrs.start || 1, depth };
      }
    }
    return null;
  });
  const [f, setF] = useState(init || { fmt: 'decimal', text: '%1.', start: 1, depth: 0 });
  const apply = () => {
    const def = ['decimal', 'lowerLetter', 'lowerRoman'][f.depth % 3];
    editor.chain().focus().setListFormat({ numFmt: f.fmt === def ? null : f.fmt, lvlText: f.text === '%1.' ? null : f.text, start: f.start }).run();
    onClose();
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Формат нумерованного списка" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && init) { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Формат нумерованного списка</div>
        <div className="modal-body">
          {!init ? <div className="bm-empty">Поставьте курсор в нумерованный список.</div> : (
            <div className="pd-grid">
              <label className="pd-field"><span>Нумерация</span>
                <select value={f.fmt} onChange={(e) => setF({ ...f, fmt: e.target.value })}>{NUM_FORMATS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
              </label>
              <label className="pd-field"><span>Вид номера</span>
                <select value={f.text} onChange={(e) => setF({ ...f, text: e.target.value })}>{LIST_PATTERNS.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
              </label>
              <label className="pd-field"><span>Начать с</span>
                <span className="pd-input"><input type="number" min="0" max="999" value={f.start} onChange={(e) => setF({ ...f, start: Math.max(0, +e.target.value || 0) })} /></span>
              </label>
            </div>
          )}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" disabled={!init} onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
