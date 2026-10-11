// Диалоги вкладок «Вставка» и «Ссылки»: «Закладка», «Перекрёстная ссылка», «Дата и время» — как в Word.
import React, { useMemo, useState } from 'react';
import { BOOKMARK_NAME, DATE_FORMATS, bookmarks, formatDate, headingList } from './editor/fields.js';

function Modal({ title, className = '', onClose, onEnter, children, actions }) {
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className={`modal pd ${className}`} role="dialog" aria-label={title} onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && onEnter && e.target.tagName !== 'SELECT') { e.preventDefault(); onEnter(); } }}>
        <div className="modal-title">{title}</div>
        <div className="modal-body">{children}</div>
        <div className="modal-actions"><span className="spacer" />{actions}</div>
      </div>
    </div>
  );
}

// «Закладка»: имя, список (по имени / по положению), добавить, удалить, перейти, скрытые закладки (_Toc, _Ref)
export function BookmarkDialog({ editor, onClose }) {
  const [, setTick] = useState(0);
  const all = [...bookmarks(editor.state.doc).values()];
  const [sort, setSort] = useState('name');
  const [hidden, setHidden] = useState(false);
  const [name, setName] = useState(() => {
    const { from } = editor.state.selection;
    const at = all.find((b) => from >= b.start && from <= b.end + 1);
    return at ? at.name : '';
  });
  const [error, setError] = useState('');
  const list = all.filter((b) => hidden || !b.name.startsWith('_'))
    .sort((a, b) => (sort === 'name' ? a.name.localeCompare(b.name, 'ru') : a.start - b.start));
  const exists = all.some((b) => b.name === name);
  const add = () => {
    if (!BOOKMARK_NAME.test(name)) { setError('Имя закладки: с буквы, только буквы, цифры и «_», без пробелов, до 40 знаков'); return; }
    editor.chain().focus().addBookmark(name).run();
    onClose();
  };
  return (
    <Modal title="Закладка" className="bm-dlg" onClose={onClose} onEnter={add}
      actions={<>
        <button className="btn primary" disabled={!name} onClick={add}>Добавить</button>
        <button className="btn" disabled={!exists} onClick={() => { editor.commands.removeBookmark(name); setName(''); setTick((t) => t + 1); }}>Удалить</button>
        <button className="btn" disabled={!exists} onClick={() => { editor.chain().focus().goToBookmark(name).run(); }}>Перейти</button>
        <button className="btn" onClick={onClose}>Закрыть</button>
      </>}>
      <label className="fd-col"><span>Имя закладки</span>
        <input autoFocus value={name} onChange={(e) => { setName(e.target.value.trim()); setError(''); }} />
      </label>
      <div className="bm-list" role="listbox">
        {list.map((b) => (
          <div key={b.name} role="option" aria-selected={b.name === name} className={`sym-row bm-row${b.name === name ? ' sel' : ''}`}
            onClick={() => setName(b.name)} onDoubleClick={() => { editor.chain().focus().goToBookmark(b.name).run(); }}>{b.name}</div>
        ))}
        {!list.length && <div className="bm-empty">В документе нет закладок</div>}
      </div>
      <div className="bm-opts">
        <span>Сортировать по:</span>
        <label><input type="radio" checked={sort === 'name'} onChange={() => setSort('name')} /> имени</label>
        <label><input type="radio" checked={sort === 'pos'} onChange={() => setSort('pos')} /> положению</label>
        <label className="bm-hidden"><input type="checkbox" checked={hidden} onChange={(e) => setHidden(e.target.checked)} /> Скрытые закладки</label>
      </div>
      {error && <div className="pd-error" role="alert">{error}</div>}
    </Modal>
  );
}

const REF_WHAT = {
  heading: [['text', 'Текст заголовка'], ['page', 'Номер страницы'], ['above', 'Выше/ниже']],
  bookmark: [['text', 'Текст закладки'], ['page', 'Номер страницы'], ['above', 'Выше/ниже']],
};

// «Перекрёстная ссылка»: на заголовок или закладку — текст, номер страницы, «выше/ниже»; гиперссылкой (\h)
export function CrossRefDialog({ editor, onClose }) {
  const [type, setType] = useState('heading');
  const [what, setWhat] = useState('text');
  const [link, setLink] = useState(true);
  const [sel, setSel] = useState(0);
  const items = useMemo(() => (type === 'heading'
    ? headingList(editor.state.doc).map((h) => ({ key: h.pos, label: h.text, level: h.level, target: { heading: h.pos } }))
    : [...bookmarks(editor.state.doc).values()].filter((b) => !b.name.startsWith('_')).sort((a, b) => a.name.localeCompare(b.name, 'ru'))
      .map((b) => ({ key: b.name, label: b.name, target: { bookmark: b.name } }))), [type, editor.state.doc]);
  const insert = () => {
    const it = items[sel];
    if (!it) return;
    editor.chain().focus().insertCrossRef(it.target, what, link).run();
  };
  return (
    <Modal title="Перекрёстная ссылка" className="xref-dlg" onClose={onClose} onEnter={insert}
      actions={<>
        <button className="btn primary" disabled={!items.length} onClick={insert}>Вставить</button>
        <button className="btn" onClick={onClose}>Закрыть</button>
      </>}>
      <div className="sym-top">
        <label className="fd-col"><span>Тип ссылки</span>
          <select value={type} onChange={(e) => { setType(e.target.value); setSel(0); }}>
            <option value="heading">Заголовок</option>
            <option value="bookmark">Закладка</option>
          </select>
        </label>
        <label className="fd-col"><span>Вставить ссылку на</span>
          <select value={what} onChange={(e) => setWhat(e.target.value)}>
            {REF_WHAT[type].map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </select>
        </label>
      </div>
      <label className="pd-check"><input type="checkbox" checked={link} onChange={(e) => setLink(e.target.checked)} /> Вставить как гиперссылку</label>
      <div className="pop-title">{type === 'heading' ? 'Для какого заголовка:' : 'Для какой закладки:'}</div>
      <div className="bm-list xref-list" role="listbox">
        {items.map((it, i) => (
          <div key={it.key} role="option" aria-selected={i === sel} className={`sym-row bm-row${i === sel ? ' sel' : ''}`}
            style={it.level ? { paddingLeft: 8 + (it.level - 1) * 14 } : undefined}
            onClick={() => setSel(i)} onDoubleClick={() => { setSel(i); editor.chain().focus().insertCrossRef(it.target, what, link).run(); }}>{it.label}</div>
        ))}
        {!items.length && <div className="bm-empty">{type === 'heading' ? 'В документе нет заголовков' : 'В документе нет закладок'}</div>}
      </div>
    </Modal>
  );
}

// «Дата и время»: форматы как в русском Word, «Обновлять автоматически» — полем DATE/TIME
export function DateDialog({ editor, onClose }) {
  const [now] = useState(() => new Date());
  const [sel, setSel] = useState(0);
  const [auto, setAuto] = useState(false);
  const insert = () => { editor.chain().focus().insertDate(DATE_FORMATS[sel], auto).run(); onClose(); };
  return (
    <Modal title="Дата и время" className="date-dlg" onClose={onClose} onEnter={insert}
      actions={<>
        <button className="btn primary" onClick={insert}>ОК</button>
        <button className="btn" onClick={onClose}>Отмена</button>
      </>}>
      <div className="pop-title">Форматы:</div>
      <div className="bm-list" role="listbox">
        {DATE_FORMATS.map((f, i) => (
          <div key={f} role="option" aria-selected={i === sel} className={`sym-row bm-row${i === sel ? ' sel' : ''}`}
            onClick={() => setSel(i)} onDoubleClick={() => { editor.chain().focus().insertDate(f, auto).run(); onClose(); }}>{formatDate(now, f)}</div>
        ))}
      </div>
      <label className="pd-check"><input type="checkbox" checked={auto} onChange={(e) => setAuto(e.target.checked)} /> Обновлять автоматически (поле, F9)</label>
    </Modal>
  );
}
