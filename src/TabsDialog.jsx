// Диалог «Табуляция» (как в Word): позиции абзаца — список, выравнивание, заполнитель; шаг позиций по умолчанию.
import React, { useState } from 'react';
import { cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';
import { DEFAULT_TAB } from './editor/tabs.js';

const ALIGNS = [['left', 'По левому краю'], ['center', 'По центру'], ['right', 'По правому краю'], ['decimal', 'По разделителю'], ['bar', 'С чертой']];
const LEADERS = [['none', 'Нет'], ['dot', 'Точки ……'], ['hyphen', 'Дефисы ----'], ['underscore', 'Подчёркивание ____']];
const num = (s) => { const v = parseFloat(String(s).replace(',', '.')); return isFinite(v) ? v : null; };
const same = (a, b) => Math.abs(a - b) < 0.05;

export default function TabsDialog({ editor, onClose }) {
  const node = editor.state.selection.$from.parent;
  const [list, setList] = useState(() => [...(node.attrs.tabs || [])]);
  const [pos, setPos] = useState(list[0] ? fmtNum(ptToCm(list[0].pos), 2) : '');
  const [align, setAlign] = useState(list[0] ? list[0].align : 'left');
  const [leader, setLeader] = useState(list[0] ? list[0].leader : 'none');
  const [def, setDef] = useState(fmtNum(ptToCm(editor.state.doc.attrs.defaultTab || DEFAULT_TAB), 2));
  const [error, setError] = useState('');

  const pick = (t) => { setPos(fmtNum(ptToCm(t.pos), 2)); setAlign(t.align); setLeader(t.leader || 'none'); setError(''); };
  const set = () => {
    const v = num(pos);
    if (v === null || v < 0 || v > 55) { setError('Позиция: число от 0 до 55 см'); return null; }
    const stop = { pos: cmToPt(v), align, leader };
    const next = [...list.filter((t) => !same(t.pos, stop.pos)), stop].sort((a, b) => a.pos - b.pos);
    setList(next);
    setError('');
    return next;
  };
  const remove = () => { const v = num(pos); if (v !== null) setList(list.filter((t) => !same(t.pos, cmToPt(v)))); };
  const ok = () => {
    let final = list;
    // как в Word: введённая, но не «установленная» позиция применяется при ОК
    if (pos.trim() && num(pos) !== null && !list.some((t) => same(t.pos, cmToPt(num(pos))))) final = set() || list;
    const d = num(def);
    if (d === null || d <= 0 || d > 55) { setError('Позиции по умолчанию: число больше 0'); return; }
    const chain = editor.chain().focus().setTabStops(final);
    if (Math.abs(cmToPt(d) - (editor.state.doc.attrs.defaultTab || DEFAULT_TAB)) > 0.01) chain.setDefaultTab(cmToPt(d));
    chain.run();
    onClose();
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal tabs-dlg" role="dialog" aria-label="Табуляция" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); ok(); } }}>
        <div className="modal-title">Табуляция</div>
        <div className="modal-body">
          <div className="tabs-grid">
            <div>
              <label className="pd-field"><span>Позиции табуляции, см</span>
                <input value={pos} onChange={(e) => setPos(e.target.value)} autoFocus />
              </label>
              <div className="tabs-list" role="listbox" aria-label="Позиции">
                {list.map((t) => (
                  <button key={t.pos} type="button" className={same(t.pos, cmToPt(num(pos) ?? -1)) ? 'on' : ''} onClick={() => pick(t)}>
                    {fmtNum(ptToCm(t.pos), 2)} см — {ALIGNS.find(([a]) => a === t.align)?.[1].toLowerCase()}
                  </button>
                ))}
                {!list.length && <div className="tabs-empty">Нет позиций</div>}
              </div>
            </div>
            <div>
              <label className="pd-field"><span>По умолчанию, см</span>
                <input value={def} onChange={(e) => setDef(e.target.value)} />
              </label>
              <fieldset><legend>Выравнивание</legend>
                {ALIGNS.map(([v, l]) => <label key={v} className="pd-check"><input type="radio" name="tab-align" checked={align === v} onChange={() => setAlign(v)} /> {l}</label>)}
              </fieldset>
              <fieldset><legend>Заполнитель</legend>
                {LEADERS.map(([v, l]) => <label key={v} className="pd-check"><input type="radio" name="tab-leader" checked={leader === v} onChange={() => setLeader(v)} /> {l}</label>)}
              </fieldset>
            </div>
          </div>
          <div className="tabs-actions">
            <button type="button" className="btn small" onClick={set}>Установить</button>
            <button type="button" className="btn small" onClick={remove}>Удалить</button>
            <button type="button" className="btn small" onClick={() => setList([])}>Удалить все</button>
          </div>
          {error && <div className="pd-error" role="alert">{error}</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
