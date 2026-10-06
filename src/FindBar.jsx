// Панель «Найти и заменить» (Ctrl+F / Ctrl+H): поле поиска, счётчик, переход, параметры, замена.
import React, { useEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { searchInfo } from './editor/search.js';

export default function FindBar({ editor, mode, setMode }) {
  const [query, setQuery] = useState(() => {
    const { from, to } = editor.state.selection;
    const sel = editor.state.doc.textBetween(from, to, ' ');
    return sel && sel.length < 100 && !sel.includes('\n') ? sel : '';
  });
  const [repl, setRepl] = useState('');
  const [caseSensitive, setCase] = useState(false);
  const [wholeWord, setWhole] = useState(false);
  const [note, setNote] = useState('');
  const findRef = useRef(null);
  const info = useEditorState({ editor, selector: ({ editor: e }) => searchInfo(e.state) });

  useEffect(() => { findRef.current && findRef.current.select(); }, [mode]);
  useEffect(() => {
    editor.commands.setSearch(query, { caseSensitive, wholeWord });
    setNote('');
  }, [editor, query, caseSensitive, wholeWord]);
  useEffect(() => () => editor.commands.clearSearch(), [editor]);

  const close = () => { setMode(null); editor.commands.focus(); };
  const onKey = (e) => {
    if (e.key === 'Escape') { e.preventDefault(); close(); }
    if (e.key === 'Enter' && e.target === findRef.current) { e.preventDefault(); e.shiftKey ? editor.commands.findPrev() : editor.commands.findNext(); }
  };
  const counter = !query ? '' : info.count ? `${info.index + 1} из ${info.count}` : 'Нет совпадений';

  return (
    <div className="findbar" role="search" onKeyDown={onKey}>
      <div className="fb-row">
        <input ref={findRef} className="fb-input" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Найти в документе" aria-label="Найти" />
        <span className={`fb-count${query && !info.count ? ' none' : ''}`}>{counter}</span>
        <button type="button" className="fb-btn" title="Предыдущее (Shift+Enter)" disabled={!info.count} onClick={() => editor.commands.findPrev()}>▲</button>
        <button type="button" className="fb-btn" title="Следующее (Enter)" disabled={!info.count} onClick={() => editor.commands.findNext()}>▼</button>
        <button type="button" className={`fb-btn opt${caseSensitive ? ' on' : ''}`} title="С учётом регистра" onClick={() => setCase(!caseSensitive)}>Аа</button>
        <button type="button" className={`fb-btn opt${wholeWord ? ' on' : ''}`} title="Только слово целиком" onClick={() => setWhole(!wholeWord)}>«ab»</button>
        <button type="button" className="fb-btn" title={mode === 'replace' ? 'Скрыть замену' : 'Заменить (Ctrl+H)'} onClick={() => setMode(mode === 'replace' ? 'find' : 'replace')}>⇄</button>
        <button type="button" className="fb-btn" title="Закрыть (Esc)" onClick={close}>✕</button>
      </div>
      {mode === 'replace' && (
        <div className="fb-row">
          <input className="fb-input" value={repl} onChange={(e) => setRepl(e.target.value)} placeholder="Заменить на" aria-label="Заменить на"
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); editor.commands.replaceCurrent(repl); } }} />
          <button type="button" className="btn small" disabled={!info.count} onClick={() => editor.commands.replaceCurrent(repl)}>Заменить</button>
          <button type="button" className="btn small" disabled={!info.count} onClick={() => {
            const n = info.count;
            editor.commands.replaceAll(repl);
            setNote(`Заменено: ${n}`);
          }}>Заменить все</button>
          {note && <span className="fb-note">{note}</span>}
        </div>
      )}
    </div>
  );
}
