// «Вставить ▾» на ленте, кнопка «Параметры вставки» у только что вставленного фрагмента и «Специальная вставка»
// (Ctrl+Alt+V) — как в Word. Режимы — editor/paste.js.
import React, { useEffect, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { usePopover } from './FontControls.jsx';
import { PASTE_MODES, pasteKey, pasteWithMode, repaste } from './editor/paste.js';
import * as fileApi from './fileApi.js';

const ModeIcon = ({ mode }) => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" aria-hidden="true">
    <rect x="5" y="4" width="14" height="17" rx="1" />
    {mode === 'keep' && <path d="M8.5 10h7M8.5 13.5h7M8.5 17h4" strokeWidth="2.4" />}
    {mode === 'merge' && <path d="M8.5 10h7M8.5 13.5h7M8.5 17h4" />}
    {mode === 'text' && <text x="12" y="17" textAnchor="middle" fontSize="10" fill="currentColor" stroke="none" fontWeight="700">A</text>}
  </svg>
);

// Вставка буфера в режиме: «исходное» с картинкой в буфере — обычная вставка системы (картинки вставляет App)
export async function pasteAs(editor, mode, nativePaste) {
  const data = await fileApi.clipboardRead();
  editor.commands.focus();
  if ((mode === 'keep' && (data.image || !data.html)) || (!data.html && !data.text)) { nativePaste(); return; }
  pasteWithMode(editor, data, mode);
}

export function PasteButton({ editor, Btn, nativePaste, onSpecial }) {
  const pop = usePopover();
  return (
    <div className="split rb-popover-anchor paste-split" ref={pop.ref}>
      <Btn big icon="paste" label="Вставить" onClick={() => { editor.commands.focus(); nativePaste(); }} title="Вставить (Ctrl+V)" />
      <button type="button" className="rb-btn split-arrow paste-arrow" title="Параметры вставки" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          <div className="pop-title">Параметры вставки</div>
          {PASTE_MODES.map(([k, label]) => (
            <button key={k} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); pasteAs(editor, k, nativePaste); }}>
              <span className="check"><ModeIcon mode={k} /></span>{label}
            </button>
          ))}
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onSpecial(); }}>
            <span className="check" />Специальная вставка… (Ctrl+Alt+V)
          </button>
        </div>
      )}
    </div>
  );
}

// Кнопка «Параметры вставки» у конца вставленного фрагмента; щелчок — меню режимов (переделывает вставку)
export function PasteOptionsTag({ editor }) {
  const p = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const s = pasteKey.getState(e.state);
      return s ? { to: s.to, mode: s.mode } : null;
    },
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  const [open, setOpen] = useState(false);
  const [, setTick] = useState(0);
  useEffect(() => {
    if (!p) { setOpen(false); return undefined; }
    const ws = editor.view.dom.closest('.workspace');
    const re = () => setTick((t) => t + 1);
    if (ws) ws.addEventListener('scroll', re);
    window.addEventListener('resize', re);
    const onKey = (e) => { if (e.key === 'Escape') editor.commands.dismissPasteOptions(); };
    window.addEventListener('keydown', onKey);
    return () => { if (ws) ws.removeEventListener('scroll', re); window.removeEventListener('resize', re); window.removeEventListener('keydown', onKey); };
  }, [p, editor]);
  if (!p) return null;
  let c;
  try { c = editor.view.coordsAtPos(Math.min(p.to, editor.state.doc.content.size)); } catch { return null; }
  return (
    <div className="paste-tag" style={{ left: c.right + 4, top: c.bottom + 2 }} onMouseDown={(e) => e.preventDefault()}>
      <button type="button" className="paste-tag-btn" title="Параметры вставки" onClick={() => setOpen(!open)}>
        <ModeIcon mode={p.mode} /><span>▾</span>
      </button>
      {open && (
        <div className="menu-pop paste-tag-menu" role="menu">
          {PASTE_MODES.map(([k, label]) => (
            <button key={k} type="button" className="menu-item" onClick={() => { setOpen(false); repaste(editor, k); }}>
              <span className="check">{p.mode === k ? '✓' : ''}</span><ModeIcon mode={k} />{label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// «Специальная вставка»: форматы, которые есть в буфере
export function PasteSpecialDialog({ editor, nativePaste, onClose }) {
  const [data, setData] = useState(null);
  const [sel, setSel] = useState(0);
  useEffect(() => { fileApi.clipboardRead().then(setData); }, []);
  const items = !data ? [] : [
    ...(data.html ? [['keep', 'Текст в формате HTML (с исходным форматированием)'], ['merge', 'Текст в формате HTML (с форматированием документа)']] : []),
    ...(data.text ? [['text', 'Неформатированный текст']] : []),
    ...(data.image ? [['image', 'Рисунок']] : []),
  ];
  const apply = (i = sel) => {
    const it = items[i];
    if (!it) return;
    onClose();
    editor.commands.focus();
    if (it[0] === 'image') nativePaste();
    else pasteWithMode(editor, data, it[0]);
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Специальная вставка" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Специальная вставка</div>
        <div className="modal-body">
          <div className="pop-title">Как:</div>
          <div className="bm-list" role="listbox" tabIndex={0}>
            {data && !items.length && <div className="bm-empty">Буфер обмена пуст</div>}
            {items.map(([k, label], i) => (
              <div key={k} role="option" aria-selected={i === sel} className={`sym-row bm-row${i === sel ? ' sel' : ''}`}
                onClick={() => setSel(i)} onDoubleClick={() => apply(i)}>{label}</div>
            ))}
          </div>
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" disabled={!items.length} onClick={() => apply()}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
