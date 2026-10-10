// Диалог «Защитить лист» / «Снять защиту листа»: пароль (необязательный) и его подтверждение
import React, { useState } from 'react';
import { checkPassword, hasPassword, makeProtection } from './protect.js';

// mode 'protect' — onDone(protection); 'unprotect' — onDone(null) после проверки пароля
export default function ProtectDialog({ mode, protection, onDone, onClose }) {
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const protect = mode === 'protect';
  const ok = async () => {
    if (busy) return;
    if (protect && pw !== pw2) { setError('Пароль и подтверждение не совпадают'); return; }
    setBusy(true);
    setError('');
    try {
      if (protect) onDone(await makeProtection(pw));
      else if (await checkPassword(protection, pw)) onDone(null);
      else { setError('Неверный пароль. Проверьте раскладку клавиатуры и регистр букв'); setBusy(false); }
    } catch (e) {
      setError(String(e.message || e));
      setBusy(false);
    }
  };
  const key = (e) => { if (e.key === 'Enter') ok(); if (e.key === 'Escape') onClose(); };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal" role="dialog" aria-label={protect ? 'Защита листа' : 'Снять защиту листа'} onMouseDown={(e) => e.stopPropagation()} onKeyDown={key}>
        <div className="modal-title">{protect ? 'Защита листа' : 'Снять защиту листа'}</div>
        <div className="modal-body">
          {protect && <p style={{ margin: '0 0 10px' }}>Заблокированные ячейки нельзя будет изменить. Разблокировать ячейки для ввода можно
            заранее: «Главная → Формат → Блокировать ячейку». Пароль необязателен.</p>}
          <label className="pd-field"><span>Пароль</span>
            <input className="pd-input" type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} style={{ width: 200 }} />
          </label>
          {protect && (
            <label className="pd-field" style={{ marginTop: 6 }}><span>Подтверждение</span>
              <input className="pd-input" type="password" value={pw2} onChange={(e) => setPw2(e.target.value)} style={{ width: 200 }} />
            </label>
          )}
          {!protect && !hasPassword(protection) && <p style={{ margin: '8px 0 0' }}>Лист защищён без пароля.</p>}
          {busy && <div style={{ marginTop: 8, color: '#666' }}>Проверка пароля…</div>}
          {error && <div className="pd-error" role="alert">{error}</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" disabled={busy} onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
