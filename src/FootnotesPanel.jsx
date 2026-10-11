// Панель «Сноски» справа от страницы: сноски по порядку ссылок, правка текста, удаление, переход к ссылке.
// Внизу листов сноски показывает разбиение (pagination.js); щелчок по сноске там открывает эту панель на ней.
import React, { useEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { footnoteRefs } from './editor/footnotes.js';

function NoteCard({ editor, note, active, onActivate }) {
  const [text, setText] = useState(note.text);
  const ref = useRef(null);
  useEffect(() => setText(note.text), [note.text]);
  useEffect(() => { if (active && ref.current) ref.current.focus(); }, [active]);
  const commit = () => { if (text !== note.text) editor.commands.setFootnoteText(note.id, text); };
  return (
    <div className={`comment-card${active ? ' active' : ''}`} onMouseDown={() => { if (!active) { onActivate(note.id); editor.commands.goToFootnote(note.id); } }}>
      <div className="cc-head">
        <span className="cc-badge">{note.n}</span>
        <span className="cc-author">Сноска {note.n}</span>
        <button type="button" className="cc-del" title="Удалить сноску (и знак сноски в тексте)" onMouseDown={(e) => e.stopPropagation()}
          onClick={() => editor.commands.deleteFootnote(note.id)}>✕</button>
      </div>
      {active ? (
        <textarea ref={ref} className="cc-text" value={text} rows={Math.max(2, text.split('\n').length)} placeholder="Текст сноски"
          onChange={(e) => setText(e.target.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Escape') { setText(note.text); editor.commands.focus(); }
            if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); commit(); editor.commands.focus(); }
          }} />
      ) : (
        <div className="cc-body">{note.text || <span className="cc-empty">(пусто)</span>}</div>
      )}
    </div>
  );
}

// Сноски при печати и в PDF: листы режет Chromium, низа страницы у нас там нет — сноски списком в конце документа
export function PrintNotes({ editor }) {
  const list = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const texts = new Map((e.state.doc.attrs.footnotes || []).map((x) => [String(x.id), x.text]));
      return footnoteRefs(e.state.doc).map((r) => ({ n: r.n, text: texts.get(String(r.id)) || '' }));
    },
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  if (!list.length) return null;
  return (
    <div className="print-notes">
      {list.map((n) => <div key={n.n} className="pero-note-p"><sup>{n.n}</sup> {n.text}</div>)}
    </div>
  );
}

export default function FootnotesPanel({ editor, active, onActivate, onClose }) {
  const list = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const texts = new Map((e.state.doc.attrs.footnotes || []).map((x) => [String(x.id), x.text]));
      return footnoteRefs(e.state.doc).map((r) => ({ id: String(r.id), n: r.n, text: texts.get(String(r.id)) || '' }));
    },
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });
  return (
    <aside className="comments-panel" aria-label="Сноски">
      <div className="cp-head">
        <span>Сноски ({list.length})</span>
        <button type="button" className="cc-del" title="Скрыть панель" onClick={onClose}>✕</button>
      </div>
      {!list.length && <div className="cp-empty">Сносок нет. Поставьте курсор и нажмите «Вставить сноску» на вкладке «Ссылки» (Ctrl+Alt+F).</div>}
      {list.map((n) => <NoteCard key={n.id} editor={editor} note={n} active={active === n.id} onActivate={onActivate} />)}
    </aside>
  );
}
