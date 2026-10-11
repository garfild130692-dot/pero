// «Перейти» (Ctrl+G, «Найти ▾ → Перейти»): к странице, разделу, закладке, примечанию, сноске, таблице, рисунку,
// заголовку, полю — по номеру («5», «+2», «-1») или к следующему/предыдущему, как в Word.
import React, { useMemo, useState } from 'react';
import { TextSelection } from '@tiptap/pm/state';
import { paginationKey } from './editor/pagination.js';
import { bookmarks, scanFields } from './editor/fields.js';
import { footnoteRefs } from './editor/footnotes.js';
import { commentAnchors } from './editor/comments.js';

const TARGETS = [
  ['page', 'Страница'], ['section', 'Раздел'], ['bookmark', 'Закладка'], ['comment', 'Примечание'], ['footnote', 'Сноска'],
  ['table', 'Таблица'], ['image', 'Рисунок'], ['heading', 'Заголовок'], ['field', 'Поле'],
];

// Позиции объектов по порядку (для перехода по номеру)
export function targetPositions(state, kind) {
  const { doc } = state;
  const out = [];
  if (kind === 'page') {
    const p = paginationKey.getState(state);
    out.push(1, ...((p && p.breaks) || []));
  } else if (kind === 'section') {
    out.push(1);
    doc.forEach((n, pos) => { if (n.attrs && n.attrs.sectPr) out.push(pos + n.nodeSize + 1); });
    if (out[out.length - 1] > doc.content.size) out.pop();
  } else if (kind === 'comment') {
    for (const a of Object.values(commentAnchors(doc))) out.push(a.start ?? a.ref);
    out.sort((a, b) => a - b);
  } else if (kind === 'footnote') {
    for (const r of footnoteRefs(doc)) out.push(r.pos + 1);
  } else if (kind === 'field') {
    for (const f of scanFields(doc)) if (f.simple || f.depth === 0) out.push((f.begin ?? f.pos) + 1);
  } else {
    const type = { table: 'table', image: 'image', heading: 'heading' }[kind];
    doc.descendants((n, pos) => {
      if (n.type.name === type) { out.push(kind === 'image' ? pos : pos + 1); return false; }
      return true;
    });
  }
  return out;
}

export default function GoToDialog({ editor, onClose }) {
  const [kind, setKind] = useState('page');
  const [val, setVal] = useState('');
  const [mark, setMark] = useState('');
  const [error, setError] = useState('');
  const marks = useMemo(() => [...bookmarks(editor.state.doc).keys()].filter((n) => !n.startsWith('_')).sort((a, b) => a.localeCompare(b, 'ru')), [editor]);

  const jump = (pos) => {
    const { state, view } = editor;
    const p = Math.max(0, Math.min(pos, state.doc.content.size));
    const sel = state.doc.nodeAt(p) && state.doc.nodeAt(p).type.name === 'image'
      ? null : TextSelection.near(state.doc.resolve(p));
    view.dispatch((sel ? state.tr.setSelection(sel) : state.tr).scrollIntoView());
    view.focus();
    setTimeout(() => {
      const dom = view.domAtPos(p).node;
      const el = dom && (dom.nodeType === 1 ? dom : dom.parentElement);
      if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
    }, 0);
  };

  // dir: 0 — по номеру из поля, ±1 — следующий/предыдущий от курсора
  const go = (dir = 0) => {
    setError('');
    if (kind === 'bookmark') {
      const name = mark || marks[0];
      if (!name) { setError('В документе нет закладок'); return; }
      editor.chain().focus().goToBookmark(name).run();
      return;
    }
    const list = targetPositions(editor.state, kind);
    if (!list.length) { setError('Таких объектов в документе нет'); return; }
    const at = editor.state.selection.from;
    let cur = 0;
    list.forEach((p, i) => { if (p <= at) cur = i; });
    let i;
    const v = val.trim();
    if (dir > 0) i = list.findIndex((p) => p > at + 1);
    else if (dir < 0) { i = -1; list.forEach((p, k) => { if (p < at - 1) i = k; }); }
    else if (!v) i = Math.min(cur + 1, list.length - 1);
    else if (/^[+-]\d+$/.test(v)) i = cur + +v;
    else if (/^\d+$/.test(v)) i = +v - 1;
    else { setError('Введите номер, например 5, +2 или -1'); return; }
    if (i == null || i < 0 || i >= list.length) { setError(`Нет такого номера (всего: ${list.length})`); return; }
    jump(list[i]);
  };

  const label = (TARGETS.find(([k]) => k === kind) || [])[1];
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd goto-dlg" role="dialog" aria-label="Перейти" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); go(0); } }}>
        <div className="modal-title">Перейти</div>
        <div className="modal-body goto-body">
          <div className="bm-list" role="listbox">
            {TARGETS.map(([k, l]) => (
              <div key={k} role="option" aria-selected={k === kind} className={`sym-row bm-row${k === kind ? ' sel' : ''}`} onClick={() => { setKind(k); setError(''); }}>{l}</div>
            ))}
          </div>
          <div className="goto-side">
            {kind === 'bookmark' ? (
              <label className="fd-col"><span>Имя закладки</span>
                <select value={mark || marks[0] || ''} onChange={(e) => setMark(e.target.value)}>
                  {!marks.length && <option value="">(нет закладок)</option>}
                  {marks.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
              </label>
            ) : (
              <label className="fd-col"><span>Введите номер: {label ? label.toLowerCase() : ''}</span>
                <input autoFocus value={val} onChange={(e) => { setVal(e.target.value); setError(''); }} placeholder="например 5, +2 или -1" />
              </label>
            )}
            <div className="pd-note">«+» и «−» — относительно текущего места. Пустое поле — следующий.</div>
            {error && <div className="pd-error">{error}</div>}
          </div>
        </div>
        <div className="modal-actions">
          {kind !== 'bookmark' && <button className="btn" onClick={() => go(-1)}>Назад</button>}
          {kind !== 'bookmark' && <button className="btn" onClick={() => go(1)}>Далее</button>}
          <span className="spacer" />
          <button className="btn primary" onClick={() => go(0)}>Перейти</button>
          <button className="btn" onClick={onClose}>Закрыть</button>
        </div>
      </div>
    </div>
  );
}
