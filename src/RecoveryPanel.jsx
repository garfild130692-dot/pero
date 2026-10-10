// «Восстановление документов»: копии несохранённых документов, оставшиеся после сбоя или аварийного завершения.
// Открыть — в этом окне (если оно пустое) или в новом; удалить — копия стирается. Закрыть окно — копии остаются
// и будут предложены при следующем запуске.
import React, { useState } from 'react';
import * as fileApi from './fileApi.js';

const fmtTime = (t) => new Date(t).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });

export default function RecoveryPanel({ items, onOpenHere, onClose }) {
  const [list, setList] = useState(items);
  const [busy, setBusy] = useState(false);
  const drop = (id) => setList((l) => { const n = l.filter((x) => x.id !== id); if (!n.length) onClose(); return n; });

  const open = async (it) => {
    setBusy(true);
    try {
      const f = await fileApi.recoveryOpen(it.id);
      drop(it.id);
      if (f) onOpenHere(f);
    } finally { setBusy(false); }
  };
  const discard = async (it) => {
    if (!window.confirm(`Удалить несохранённую копию «${it.name}»? Восстановить её будет нельзя.`)) return;
    await fileApi.recoveryDiscard(it.id);
    drop(it.id);
  };

  return (
    <div className="modal-backdrop">
      <div className="modal recovery" role="dialog" aria-label="Восстановление документов">
        <div className="modal-title">Восстановление документов</div>
        <div className="modal-body">
          <p>Перо было завершено, не сохранив изменения. Найдены автоматически сохранённые копии:</p>
          <ul className="recovery-list">
            {list.map((it) => (
              <li key={it.id}>
                <div className="recovery-info">
                  <div className="recovery-name">{it.name}</div>
                  <div className="recovery-meta">{it.path ? it.path : 'новый, ещё не сохранявшийся'} · {fmtTime(it.time)}</div>
                </div>
                <button type="button" className="btn small primary" disabled={busy} onClick={() => open(it)}>Открыть</button>
                <button type="button" className="btn small danger" disabled={busy} onClick={() => discard(it)}>Удалить</button>
              </li>
            ))}
          </ul>
          <p className="recovery-note">Восстановленный документ откроется как копия — сохраните его, чтобы не потерять изменения.</p>
        </div>
        <div className="modal-actions">
          <button type="button" className="btn" onClick={onClose} title="Копии сохранятся и будут предложены при следующем запуске">Решить позже</button>
        </div>
      </div>
    </div>
  );
}
