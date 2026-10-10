// «Диспетчер имён»: именованные диапазоны и константы книги (wb.names) — список, создание, изменение, удаление.
// Формула имени показывается и вводится в русской записи (=Данные!$C$2:$C$11, =0,2), хранится — в канонической.
import React, { useState } from 'react';
import { fromUser, localize } from './formula.js';
import { isValidName } from './model.js';

export default function NamesDialog({ wb, refOfSelection, onSave, onClose }) {
  const [names, setNames] = useState(() => (wb.names || []).map((n) => ({ ...n })));
  const [sel, setSel] = useState(names.length ? 0 : -1);
  const [edit, setEdit] = useState(null); // { index | -1, name, text, sheet }
  const [error, setError] = useState('');

  const startNew = () => { setError(''); setEdit({ index: -1, name: '', text: `=${refOfSelection}`, sheet: '' }); };
  const startEdit = () => { if (sel < 0) return; const n = names[sel]; setError(''); setEdit({ index: sel, name: n.name, text: `=${localize(n.text)}`, sheet: n.sheet || '' }); };
  const remove = () => { if (sel < 0) return; setNames(names.filter((_, i) => i !== sel)); setSel(-1); };
  const apply = () => {
    const name = edit.name.trim();
    if (!isValidName(name)) { setError('Имя должно начинаться с буквы или «_», без пробелов и не совпадать с адресом ячейки (например, A1)'); return; }
    const scope = edit.sheet || null;
    if (names.some((n, i) => i !== edit.index && n.name.toLowerCase() === name.toLowerCase() && (n.sheet || null) === scope)) { setError('Такое имя уже есть в этой области'); return; }
    const raw = edit.text.trim().replace(/^=/, '');
    if (!raw) { setError('Укажите диапазон или значение'); return; }
    const def = { name, text: fromUser(raw), sheet: scope };
    const next = edit.index >= 0 ? names.map((n, i) => (i === edit.index ? def : n)) : [...names, def];
    setNames(next);
    setSel(edit.index >= 0 ? edit.index : next.length - 1);
    setEdit(null);
  };
  const key = (e) => { if (e.key === 'Escape') { e.stopPropagation(); if (edit) setEdit(null); else onClose(); } };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal names-dlg" role="dialog" aria-label="Диспетчер имён" onMouseDown={(e) => e.stopPropagation()} onKeyDown={key}>
        <div className="modal-title">Диспетчер имён</div>
        <div className="modal-body">
          <div className="names-list" role="listbox" aria-label="Имена">
            <div className="names-row head"><span>Имя</span><span>Диапазон / значение</span><span>Область</span></div>
            {names.map((n, i) => (
              <button key={`${n.sheet}!${n.name}`} type="button" className={`names-row${i === sel ? ' on' : ''}`} onClick={() => setSel(i)} onDoubleClick={() => { setSel(i); setTimeout(startEdit); }}>
                <span>{n.name}</span><span>={localize(n.text)}</span><span>{n.sheet || 'Книга'}</span>
              </button>
            ))}
            {!names.length && <div className="tabs-empty">Имён нет. «Создать» — имя для выделенного диапазона.</div>}
          </div>
          <div className="tabs-actions">
            <button type="button" className="btn small" onClick={startNew}>Создать…</button>
            <button type="button" className="btn small" disabled={sel < 0} onClick={startEdit}>Изменить…</button>
            <button type="button" className="btn small" disabled={sel < 0} onClick={remove}>Удалить</button>
          </div>
          {edit && (
            <div className="names-edit">
              <label className="pd-field"><span>Имя</span><input className="pd-input" autoFocus value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} style={{ width: 260 }} /></label>
              <label className="pd-field"><span>Область</span>
                <select value={edit.sheet} onChange={(e) => setEdit({ ...edit, sheet: e.target.value })} style={{ width: 260 }}>
                  <option value="">Книга</option>
                  {wb.sheets.map((s) => <option key={s.name} value={s.name}>{s.name}</option>)}
                </select>
              </label>
              <label className="pd-field"><span>Диапазон</span><input className="pd-input" value={edit.text} onChange={(e) => setEdit({ ...edit, text: e.target.value })} onKeyDown={(e) => e.key === 'Enter' && apply()} style={{ width: 260 }} /></label>
              <div className="tabs-actions">
                <button type="button" className="btn small" onClick={apply}>{edit.index >= 0 ? 'Применить' : 'Добавить'}</button>
                <button type="button" className="btn small" onClick={() => setEdit(null)}>Отмена</button>
              </div>
            </div>
          )}
          {error && <div className="pd-error" role="alert">{error}</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={() => onSave(names)}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
