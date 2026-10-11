// «Вставка → Экспресс-блоки»: «Поле…» (номера, даты, сведения о документе), свойства документа; «Файл → Свойства».
import React, { useState } from 'react';
import { DATE_FORMATS, bookmarks, fieldNodes, fieldValue, formatDate } from './editor/fields.js';
import { PROP_LABELS } from './editor/docProps.js';
import { usePopover } from './FontControls.jsx';

// [имя, категория, описание, есть формат даты]
export const FIELD_LIST = [
  ['PAGE', 'Нумерация', 'Номер текущей страницы'],
  ['NUMPAGES', 'Сведения о документе', 'Число страниц документа'],
  ['NUMWORDS', 'Сведения о документе', 'Число слов в документе'],
  ['NUMCHARS', 'Сведения о документе', 'Число знаков в документе'],
  ['FILENAME', 'Сведения о документе', 'Имя файла документа'],
  ['AUTHOR', 'Сведения о документе', 'Автор (свойство документа)'],
  ['TITLE', 'Сведения о документе', 'Название (свойство документа)'],
  ['SUBJECT', 'Сведения о документе', 'Тема (свойство документа)'],
  ['KEYWORDS', 'Сведения о документе', 'Ключевые слова (свойство документа)'],
  ['COMMENTS', 'Сведения о документе', 'Примечания (свойство документа)'],
  ['LASTSAVEDBY', 'Сведения о документе', 'Кто последним сохранил документ'],
  ['DATE', 'Дата и время', 'Текущая дата', true],
  ['TIME', 'Дата и время', 'Текущее время', true],
  ['CREATEDATE', 'Дата и время', 'Дата создания документа', true],
];
const CATS = ['(все)', 'Нумерация', 'Дата и время', 'Сведения о документе'];

// Вставить поле с вычисленным результатом (без значения — «(нет данных)» до заполнения свойств)
export function insertField(editor, instr, file) {
  const { state } = editor;
  const pos = state.selection.from;
  const v = fieldValue(instr, pos, { state, marks: bookmarks(state.doc), now: new Date(), file });
  const marks = state.storedMarks || state.selection.$from.marks();
  const nodes = fieldNodes(state.schema, instr, v == null ? '' : v, marks.filter((m) => m.type.name !== 'link'));
  editor.chain().focus().command(({ tr }) => { tr.replaceWith(tr.selection.from, tr.selection.to, nodes); return true; }).run();
}

export function FieldDialog({ editor, file, onProps, onClose }) {
  const [cat, setCat] = useState('(все)');
  const [sel, setSel] = useState('PAGE');
  const [fmt, setFmt] = useState(DATE_FORMATS[0]);
  const [path, setPath] = useState(false);
  const list = FIELD_LIST.filter((f) => cat === '(все)' || f[1] === cat);
  const def = FIELD_LIST.find((f) => f[0] === sel) || FIELD_LIST[0];
  const instr = `${sel}${def[3] ? ` \\@ "${fmt}"` : ''}${sel === 'FILENAME' && path ? ' \\p' : ''}`;
  const ok = () => { insertField(editor, instr, file); onClose(); };
  const now = new Date();
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd field-dlg" role="dialog" aria-label="Поле" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); ok(); } }}>
        <div className="modal-title">Поле</div>
        <div className="modal-body field-body">
          <div>
            <label className="fd-col"><span>Категории</span>
              <select value={cat} onChange={(e) => setCat(e.target.value)}>{CATS.map((c) => <option key={c}>{c}</option>)}</select>
            </label>
            <div className="bm-list" role="listbox">
              {list.map(([n]) => (
                <div key={n} role="option" aria-selected={n === sel} className={`sym-row bm-row${n === sel ? ' sel' : ''}`} onClick={() => setSel(n)} onDoubleClick={() => { setSel(n); setTimeout(ok, 0); }}>{n}</div>
              ))}
            </div>
          </div>
          <div className="field-side">
            <div className="field-desc">{def[2]}</div>
            {def[3] && (
              <label className="fd-col"><span>Формат даты</span>
                <select value={fmt} onChange={(e) => setFmt(e.target.value)}>{DATE_FORMATS.map((f) => <option key={f} value={f}>{formatDate(now, f)}</option>)}</select>
              </label>
            )}
            {sel === 'FILENAME' && <label className="pd-check"><input type="checkbox" checked={path} onChange={(e) => setPath(e.target.checked)} /> Добавить путь к имени файла</label>}
            {['AUTHOR', 'TITLE', 'SUBJECT', 'KEYWORDS', 'COMMENTS'].includes(sel) && (
              <button type="button" className="btn small" onClick={onProps}>Свойства документа…</button>
            )}
            <div className="field-code">Код поля: <code>{instr}</code></div>
            <div className="pd-note">Значения полей обновляются клавишей F9 («Ссылки → Обновить поля»).</div>
          </div>
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={ok}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// «Свойства документа»: название, тема, автор, ключевые слова, примечания (docProps/core.xml)
export function DocPropsDialog({ editor, onClose }) {
  const [p, setP] = useState(() => ({ ...(editor.state.doc.attrs.props || {}) }));
  const set = (k) => (e) => setP((x) => ({ ...x, [k]: e.target.value }));
  const ok = () => {
    editor.chain().focus().command(({ tr }) => { tr.setDocAttribute('props', p).setDocAttribute('propsChanged', true); return true; }).run();
    editor.commands.updateFields(true); // поля AUTHOR, TITLE… — сразу с новыми значениями
    onClose();
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Свойства документа" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Свойства документа</div>
        <div className="modal-body props-body">
          {Object.entries(PROP_LABELS).map(([k, l]) => (
            <label key={k} className="fd-col"><span>{l}</span>
              {k === 'comments' ? <textarea rows={3} value={p[k] || ''} onChange={set(k)} /> : <input value={p[k] || ''} onChange={set(k)} />}
            </label>
          ))}
          {p.created && <div className="pd-note">Создан: {new Date(p.created).toLocaleString('ru-RU')}</div>}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={ok}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// «Экспресс-блоки ▾»: поле…, свойства документа (как поля), текст из файла
export function QuickPartsButton({ editor, Btn, file, onDialog, onInsertFile }) {
  const pop = usePopover();
  const item = (label, fn) => (
    <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}><span className="check" />{label}</button>
  );
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big icon="quickParts" label="Экспресс-блоки" title="Поля, свойства документа, текст из файла" active={pop.open} onClick={() => pop.setOpen(!pop.open)} />
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          <div className="pop-title">Свойство документа</div>
          {[['AUTHOR', 'Автор'], ['TITLE', 'Название'], ['SUBJECT', 'Тема'], ['KEYWORDS', 'Ключевые слова'], ['COMMENTS', 'Примечания']].map(([f, l]) => item(l, () => insertField(editor, f, file)))}
          <div className="menu-sep" />
          {item('Поле…', () => onDialog('field'))}
          {item('Свойства документа…', () => onDialog('docProps'))}
          <div className="menu-sep" />
          {item('Текст из файла…', onInsertFile)}
        </div>
      )}
    </div>
  );
}
