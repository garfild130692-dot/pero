// Диалог «Сортировка» (Данные → Сортировка…): уровни «Сортировать по / Затем по», порядок, «Мои данные содержат
// заголовки» — как в Excel. rg — сортируемый диапазон (выделение или текущая область).
import React, { useState } from 'react';
import { colName, key } from './address.js';
import { shownText } from './filter.js';

// Похоже ли на строку заголовков: в первой строке только текст, а ниже есть числа или первая строка отличается видом
export function guessHeader(sheet, rg) {
  if (rg.r2 <= rg.r1) return false;
  let text = 0; let other = 0;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const a = sheet.cells.get(key(rg.r1, c)); const b = sheet.cells.get(key(rg.r1 + 1, c));
    if (a && typeof a.v === 'string') text++;
    if (b && b.v != null && typeof b.v !== 'string') other++;
    if (a && a.s && a.s.bold && !(b && b.s && b.s.bold)) other++;
  }
  return text === rg.c2 - rg.c1 + 1 && other > 0;
}

export default function SortDialog({ sheet, rg, onApply, onClose }) {
  const [header, setHeader] = useState(() => guessHeader(sheet, rg));
  const [levels, setLevels] = useState([{ col: rg.c1, desc: false }]);
  const colLabel = (c) => (header ? shownText(sheet.cells.get(key(rg.r1, c))) || `Столбец ${colName(c)}` : `Столбец ${colName(c)}`);
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  const set = (i, patch) => setLevels(levels.map((l, k) => (k === i ? { ...l, ...patch } : l)));
  const [cur, setCur] = useState(0);
  const add = () => { const used = new Set(levels.map((l) => l.col)); const c = cols.find((x) => !used.has(x)) ?? rg.c1; setLevels([...levels, { col: c, desc: false }]); setCur(levels.length); };
  const remove = () => { if (levels.length <= 1) return; setLevels(levels.filter((_, i) => i !== cur)); setCur(Math.max(0, cur - 1)); };
  const move = (d) => { const j = cur + d; if (j < 0 || j >= levels.length) return; const n = [...levels]; [n[cur], n[j]] = [n[j], n[cur]]; setLevels(n); setCur(j); };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal sort-dlg" role="dialog" aria-label="Сортировка" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Сортировка</div>
        <div className="modal-body">
          <div className="tabs-actions">
            <button type="button" className="btn small" onClick={add} disabled={levels.length >= 64}>Добавить уровень</button>
            <button type="button" className="btn small" onClick={remove} disabled={levels.length <= 1}>Удалить уровень</button>
            <button type="button" className="btn small" onClick={() => move(-1)} disabled={cur === 0} title="Выше">▲</button>
            <button type="button" className="btn small" onClick={() => move(1)} disabled={cur === levels.length - 1} title="Ниже">▼</button>
            <span className="spacer" />
            <label className="sort-head"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> Мои данные содержат заголовки</label>
          </div>
          <div className="sort-levels">
            {levels.map((l, i) => (
              <div key={i} className={`sort-level${i === cur ? ' on' : ''}`} onMouseDown={() => setCur(i)}>
                <span className="sort-cap">{i === 0 ? 'Сортировать по' : 'Затем по'}</span>
                <select value={l.col} onChange={(e) => set(i, { col: +e.target.value })}>
                  {cols.map((c) => <option key={c} value={c}>{colLabel(c)}</option>)}
                </select>
                <select value={l.desc ? '1' : '0'} onChange={(e) => set(i, { desc: e.target.value === '1' })}>
                  <option value="0">По возрастанию (А → Я)</option>
                  <option value="1">По убыванию (Я → А)</option>
                </select>
              </div>
            ))}
          </div>
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={() => onApply(levels, header)}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
