// Кнопки списков группы «Абзац» как в Word: «Маркированный список ▾» (библиотека маркеров, новый маркер) и
// «Нумерованный список ▾» (библиотека форматов, начать заново, продолжить, начальное значение); диалог начального значения.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { BULLET_LIBRARY, NUMBER_LIBRARY, fmtNumber } from './editor/numbering.js';

export function BulletButton({ editor, Btn, active, onDialog }) {
  const pop = usePopover();
  const c = () => editor.chain().focus();
  return (
    <div className="split rb-popover-anchor" ref={pop.ref}>
      <Btn icon="bullets" title="Маркированный список" active={active} onClick={() => c().toggleBulletList().run()} />
      <button type="button" className="rb-btn split-arrow" title="Библиотека маркеров" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          <div className="pop-title">Библиотека маркеров</div>
          <div className="lib-grid">
            {BULLET_LIBRARY.map((ch) => (
              <button key={ch} type="button" className="lib-cell bullet" title={`Маркер «${ch}»`} onMouseDown={(e) => e.preventDefault()}
                onClick={() => { pop.setOpen(false); c().setBullet(ch).run(); }}>
                {[0, 1, 2].map((i) => <span key={i}><b>{ch}</b><i /></span>)}
              </button>
            ))}
          </div>
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onDialog('bulletSymbol'); }}>
            <span className="check" />Определить новый маркер…
          </button>
        </div>
      )}
    </div>
  );
}

export function NumberButton({ editor, Btn, active, onDialog }) {
  const pop = usePopover();
  const c = () => editor.chain().focus();
  const item = (label, fn, title) => (
    <button type="button" className="menu-item" title={title} onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}>
      <span className="check" />{label}
    </button>
  );
  return (
    <div className="split rb-popover-anchor" ref={pop.ref}>
      <Btn icon="numbers" title="Нумерованный список" active={active} onClick={() => c().toggleOrderedList().run()} />
      <button type="button" className="rb-btn split-arrow" title="Библиотека нумерации" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          <div className="pop-title">Библиотека нумерации</div>
          <div className="lib-grid">
            {NUMBER_LIBRARY.map(([fmt, text]) => (
              <button key={fmt + text} type="button" className="lib-cell" onMouseDown={(e) => e.preventDefault()}
                onClick={() => { pop.setOpen(false); c().setNumberFormat(fmt, text).run(); }}>
                {[1, 2, 3].map((n) => <span key={n}><b>{text.replace('%1', fmtNumber(n, fmt))}</b><i /></span>)}
              </button>
            ))}
          </div>
          <div className="menu-sep" />
          {item('Начать заново с 1', () => c().restartNumbering(1).run(), 'Этот список нумеруется с 1')}
          {item('Продолжить нумерацию', () => c().continueNumbering().run(), 'Продолжить номера предыдущего нумерованного списка')}
          {item('Задать начальное значение…', () => onDialog('numberValue'))}
          {item('Формат нумерованного списка…', () => onDialog('listFormat'))}
        </div>
      )}
    </div>
  );
}

// «Задание начального значения»: новый список с номера N или продолжение предыдущего
export function NumberValueDialog({ editor, onClose }) {
  const [mode, setMode] = useState('new');
  const [start, setStart] = useState(() => {
    const $f = editor.state.selection.$from;
    for (let d = $f.depth; d > 0; d--) if ($f.node(d).type.name === 'orderedList') return $f.node(d).attrs.start ?? 1;
    return 1;
  });
  const apply = () => {
    if (mode === 'new') editor.chain().focus().restartNumbering(start).run();
    else editor.chain().focus().continueNumbering().run();
    onClose();
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Задание начального значения" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Задание начального значения</div>
        <div className="modal-body">
          <label className="pd-check"><input type="radio" checked={mode === 'new'} onChange={() => setMode('new')} /> Начать новый список</label>
          <label className="pd-check"><input type="radio" checked={mode === 'continue'} onChange={() => setMode('continue')} /> Продолжить предыдущий список</label>
          <label className="pd-field" style={{ marginTop: 10 }}><span>Начальное значение</span>
            <span className="pd-input small"><input type="number" min="0" max="999" value={start} disabled={mode !== 'new'} onChange={(e) => setStart(Math.max(0, +e.target.value || 0))} /></span>
          </label>
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
