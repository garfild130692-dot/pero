// Панель «Примечания» справа от страницы: список по порядку в документе, правка текста, удаление, переход.
import React, { useEffect, useRef, useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { commentsKey, fmtWordDate as fmtDate, orderedComments } from './editor/comments.js';

function CommentCard({ editor, c, n, active }) {
  const [text, setText] = useState(c.text);
  const ref = useRef(null);
  useEffect(() => setText(c.text), [c.text]);
  useEffect(() => { if (active && !c.text && ref.current) ref.current.focus(); }, [active, c.text]);
  return (
    <div className={`comment-card${active ? ' active' : ''}`} onMouseDown={() => !active && editor.commands.goToComment(c.id)}>
      <div className="cc-head">
        <span className="cc-badge">К{n}</span>
        <span className="cc-author">{c.author || 'Автор не указан'}</span>
        <span className="cc-date">{fmtDate(c.date)}</span>
        <button type="button" className="cc-del" title="Удалить примечание" onMouseDown={(e) => e.stopPropagation()} onClick={() => editor.commands.deleteComment(c.id)}>✕</button>
      </div>
      {active ? (
        <textarea ref={ref} className="cc-text" value={text} rows={Math.max(2, text.split('\n').length)} placeholder="Текст примечания"
          onChange={(e) => setText(e.target.value)}
          onBlur={() => text !== c.text && editor.commands.updateComment(c.id, text)}
          onKeyDown={(e) => { if (e.key === 'Escape') { setText(c.text); editor.commands.focus(); } }} />
      ) : (
        <div className="cc-body">{c.text || <span className="cc-empty">(пусто)</span>}</div>
      )}
    </div>
  );
}

export default function CommentsPanel({ editor, onClose }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({ list: orderedComments(e.state), active: commentsKey.getState(e.state)?.active ?? null }),
  });
  return (
    <aside className="comments-panel" aria-label="Примечания">
      <div className="cp-head">
        <span>Примечания ({s.list.length})</span>
        <button type="button" className="cc-del" title="Скрыть панель" onClick={onClose}>✕</button>
      </div>
      {!s.list.length && <div className="cp-empty">Примечаний нет. Выделите фрагмент и нажмите «Создать примечание» на вкладке «Рецензирование».</div>}
      {s.list.map((c, i) => <CommentCard key={c.id} editor={editor} c={c} n={i + 1} active={s.active === c.id} />)}
    </aside>
  );
}
