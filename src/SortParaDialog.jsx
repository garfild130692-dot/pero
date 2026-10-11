// «Сортировка текста» (группа «Абзац»): абзацы выделения / пункты списка или строки таблицы — по тексту, числу, дате.
import React, { useState } from 'react';
import { sortTarget } from './editor/sort.js';

const TYPES = [['text', 'Текст'], ['number', 'Число'], ['date', 'Дата']];

export default function SortParaDialog({ editor, onClose }) {
  const [target] = useState(() => sortTarget(editor.state));
  const [type, setType] = useState('text');
  const [desc, setDesc] = useState(false);
  const [column, setColumn] = useState(0);
  const [header, setHeader] = useState(() => !!(target && target.kind === 'rows' && target.table.firstChild.attrs.repeatHeader));
  const [error, setError] = useState('');
  const apply = () => {
    const ok = editor.chain().focus().sortParagraphs({ type, desc, column, header }).run();
    if (!ok) { setError('Порядок уже такой — менять нечего.'); return; }
    onClose();
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Сортировка текста" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Сортировка текста</div>
        <div className="modal-body">
          {!target ? <div className="bm-empty">Выделите несколько абзацев, поставьте курсор в список или в таблицу.</div> : (<>
            <div className="pd-grid">
              <label className="pd-field"><span>Сначала по</span>
                {target.kind === 'rows' ? (
                  <select value={column} onChange={(e) => setColumn(+e.target.value)}>
                    {Array.from({ length: target.cols }, (_, i) => <option key={i} value={i}>Столбец {i + 1}</option>)}
                  </select>
                ) : <select disabled><option>{target.parent.type.name.endsWith('List') ? 'Пунктам списка' : 'Абзацам'}</option></select>}
              </label>
              <label className="pd-field"><span>Тип</span>
                <select value={type} onChange={(e) => setType(e.target.value)}>{TYPES.map(([k, l]) => <option key={k} value={k}>{l}</option>)}</select>
              </label>
            </div>
            <div className="sort-order">
              <label className="pd-check"><input type="radio" checked={!desc} onChange={() => setDesc(false)} /> по возрастанию</label>
              <label className="pd-check"><input type="radio" checked={desc} onChange={() => setDesc(true)} /> по убыванию</label>
            </div>
            {target.kind === 'rows' && (
              <label className="pd-check"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> Со строкой заголовка</label>
            )}
            {error && <div className="pd-error">{error}</div>}
          </>)}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" disabled={!target} onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
