// Меню автофильтра столбца (кнопка ▾ в строке заголовков) и диалог «Пользовательский автофильтр» — как в Excel:
// сортировка, снять фильтр столбца, текстовые/числовые/фильтры по дате, поиск, список значений с флажками.
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { columnValues, shownText } from './filter.js';
import { isDateFormat, parseInput } from './format.js';
import { key } from './address.js';

// Условия: подпись → { op, шаблон значения }; «между» — два условия «И»
const TEXT_OPS = [
  ['равно', 'equal', '%'], ['не равно', 'notEqual', '%'], ['начинается с', 'equal', '%*'], ['не начинается с', 'notEqual', '%*'],
  ['заканчивается на', 'equal', '*%'], ['не заканчивается на', 'notEqual', '*%'], ['содержит', 'equal', '*%*'], ['не содержит', 'notEqual', '*%*'],
];
const NUM_OPS = [
  ['равно', 'equal', '%'], ['не равно', 'notEqual', '%'], ['больше', 'greaterThan', '%'], ['больше или равно', 'greaterThanOrEqual', '%'],
  ['меньше', 'lessThan', '%'], ['меньше или равно', 'lessThanOrEqual', '%'],
];
const DATE_OPS = [['равно', 'equal', '%'], ['до', 'lessThan', '%'], ['после', 'greaterThan', '%'], ['не раньше', 'greaterThanOrEqual', '%'], ['не позже', 'lessThanOrEqual', '%']];
const ALL_OPS = [...TEXT_OPS, ...NUM_OPS.slice(2)];

// Вид столбца: даты, числа или текст — по большинству непустых ячеек данных
export function columnKind(sheet, af, j) {
  let num = 0; let date = 0; let text = 0;
  for (let r = af.ref.r1 + 1; r <= af.ref.r2; r++) {
    const c = sheet.cells.get(key(r, af.ref.c1 + j));
    if (!c || c.v == null || c.v === '') continue;
    if (typeof c.v === 'number') { if (c.s && c.s.numFmt && isDateFormat(c.s.numFmt)) date++; else num++; } else text++;
  }
  return date > num && date >= text ? 'date' : num >= text && num > 0 ? 'number' : 'text';
}

export default function FilterMenu({ sheet, j, x, y, onApply, onSort, onCustom, onClose }) {
  const af = sheet.autoFilter;
  const crit = af.cols[j];
  const { list, blank } = useMemo(() => columnValues(sheet, af, j), [sheet, af, j]);
  const kind = useMemo(() => columnKind(sheet, af, j), [sheet, af, j]);
  const head = shownText(sheet.cells.get(key(af.ref.r1, af.ref.c1 + j))) || `Столбец ${j + 1}`;
  const [search, setSearch] = useState('');
  const [checked, setChecked] = useState(() => {
    const on = new Set();
    for (const it of list) if (!crit || !crit.values || crit.values.some((v) => String(v).toLowerCase() === it.text.toLowerCase())) on.add(it.text.toLowerCase());
    if (blank && (!crit || !crit.values || crit.blank)) on.add('');
    return on;
  });
  const [sub, setSub] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    const down = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); };
    const esc = (e) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('mousedown', down); window.addEventListener('keydown', esc);
    return () => { window.removeEventListener('mousedown', down); window.removeEventListener('keydown', esc); };
  }, [onClose]);
  const q = search.trim().toLowerCase();
  const shown = list.filter((it) => !q || it.text.toLowerCase().includes(q));
  const showBlank = blank && (!q || '(пустые)'.includes(q));
  const allKeys = [...shown.map((it) => it.text.toLowerCase()), ...(showBlank ? [''] : [])];
  const allOn = allKeys.every((k) => checked.has(k));
  const toggle = (k) => { const n = new Set(checked); if (n.has(k)) n.delete(k); else n.add(k); setChecked(n); };
  const toggleAll = () => { const n = new Set(checked); for (const k of allKeys) { if (allOn) n.delete(k); else n.add(k); } setChecked(n); };
  const ok = () => {
    // при поиске — только найденные и отмеченные (как Excel)
    const pool = q ? allKeys : [...list.map((it) => it.text.toLowerCase()), ...(blank ? [''] : [])];
    const on = pool.filter((k) => checked.has(k));
    if (!q && on.length === pool.length) { onApply(null); return; }
    const values = list.filter((it) => on.includes(it.text.toLowerCase())).map((it) => it.text);
    onApply({ values, blank: on.includes('') });
  };
  const ops = kind === 'date' ? DATE_OPS : kind === 'number' ? NUM_OPS : TEXT_OPS;
  const label = kind === 'date' ? 'Фильтры по дате' : kind === 'number' ? 'Числовые фильтры' : 'Текстовые фильтры';
  const style = { left: Math.min(x, window.innerWidth - 270), top: Math.min(y, window.innerHeight - 440) };
  return (
    <div className="filter-menu" ref={ref} style={style} role="menu" onMouseDown={(e) => e.stopPropagation()}>
      <button type="button" className="menu-item" onClick={() => onSort(false)}>{kind === 'text' ? 'Сортировка от А до Я' : kind === 'date' ? 'Сортировка от старых к новым' : 'Сортировка по возрастанию'}</button>
      <button type="button" className="menu-item" onClick={() => onSort(true)}>{kind === 'text' ? 'Сортировка от Я до А' : kind === 'date' ? 'Сортировка от новых к старым' : 'Сортировка по убыванию'}</button>
      <div className="menu-sep" />
      <button type="button" className="menu-item" disabled={!crit} onClick={() => onApply(null)}>Удалить фильтр из столбца «{head}»</button>
      <div className="fm-sub" onMouseEnter={() => setSub(true)} onMouseLeave={() => setSub(false)}>
        <button type="button" className="menu-item" onClick={() => setSub(!sub)}>{label}<span className="fm-arrow">▸</span></button>
        {sub && (
          <div className="filter-menu fm-subpop">
            {ops.map(([text, op, tpl]) => <button key={text} type="button" className="menu-item" onClick={() => onCustom(kind, { op, tpl })}>{text}…</button>)}
            {kind !== 'text' && <button type="button" className="menu-item" onClick={() => onCustom(kind, { between: true })}>между…</button>}
            <div className="menu-sep" />
            <button type="button" className="menu-item" onClick={() => onCustom(kind, null)}>Настраиваемый фильтр…</button>
          </div>
        )}
      </div>
      <input className="fm-search" placeholder="Поиск" value={search} onChange={(e) => setSearch(e.target.value)} autoFocus />
      <div className="fm-list">
        <label className="fm-item"><input type="checkbox" checked={allOn} onChange={toggleAll} />{q ? '(Выделить все результаты поиска)' : '(Выделить всё)'}</label>
        {shown.map((it) => (
          <label key={it.text} className="fm-item"><input type="checkbox" checked={checked.has(it.text.toLowerCase())} onChange={() => toggle(it.text.toLowerCase())} />{it.text}</label>
        ))}
        {showBlank && <label className="fm-item"><input type="checkbox" checked={checked.has('')} onChange={() => toggle('')} />(Пустые)</label>}
        {!shown.length && !showBlank && <div className="fm-empty">Совпадений нет</div>}
      </div>
      <div className="fm-actions">
        <button type="button" className="btn small primary" onClick={ok} disabled={q ? !allKeys.some((k) => checked.has(k)) : false}>ОК</button>
        <button type="button" className="btn small" onClick={onClose}>Отмена</button>
      </div>
    </div>
  );
}

// Значение условия для файла: число — с точкой, дата — серийный номер, текст — по шаблону (* для «содержит» и т. п.)
function condValue(kind, tpl, text) {
  const t = String(text).trim();
  if (kind !== 'text') {
    const p = parseInput(t);
    if (typeof p.value === 'number') return String(p.value);
  }
  return (tpl || '%').replace('%', t);
}
// для показа в диалоге: условие → подпись, текст
function condShow(kind, c) {
  const ops = kind === 'date' ? DATE_OPS : kind === 'number' ? NUM_OPS : ALL_OPS;
  for (const [label, op, tpl] of ops) {
    if (op !== c.op) continue;
    const re = new RegExp(`^${tpl.replace(/\*/g, '\\*').replace('%', '(.*)')}$`);
    const m = re.exec(c.val);
    if (m && (tpl === '%' || !/^[*]|[*]$/.test(m[1]))) {
      let v = m[1];
      if (kind === 'date' && /^\d+(\.\d+)?$/.test(v)) { const d = new Date(Date.UTC(1899, 11, 30) + +v * 86400000); v = d.toLocaleDateString('ru-RU', { timeZone: 'UTC' }); }
      return { label, text: v.replace('.', kind === 'number' ? ',' : '.') };
    }
  }
  return { label: ops[0][0], text: c.val };
}

// «Пользовательский автофильтр»: два условия, И/ИЛИ
export function CustomFilterDialog({ kind, head, initial, preset, onApply, onClose }) {
  const ops = kind === 'date' ? DATE_OPS : kind === 'number' ? NUM_OPS : ALL_OPS;
  const start = () => {
    if (preset && preset.between) return { and: true, rows: [{ label: kind === 'date' ? 'не раньше' : 'больше или равно', text: '' }, { label: kind === 'date' ? 'не позже' : 'меньше или равно', text: '' }] };
    if (preset) return { and: true, rows: [{ label: ops.find(([, op, tpl]) => op === preset.op && tpl === preset.tpl)[0], text: '' }, { label: '', text: '' }] };
    if (initial && initial.custom) {
      const rows = initial.custom.conds.map((c) => condShow(kind, c));
      while (rows.length < 2) rows.push({ label: '', text: '' });
      return { and: initial.custom.and, rows };
    }
    return { and: true, rows: [{ label: ops[0][0], text: '' }, { label: '', text: '' }] };
  };
  const [st, setSt] = useState(start);
  const set = (i, patch) => setSt({ ...st, rows: st.rows.map((r, k) => (k === i ? { ...r, ...patch } : r)) });
  const ok = () => {
    const conds = st.rows.filter((r) => r.label && (r.text.trim() !== '' || /равно/.test(r.label))).map((r) => {
      const [, op, tpl] = ops.find(([l]) => l === r.label);
      return { op, val: condValue(kind, tpl, r.text) };
    });
    onApply(conds.length ? { custom: { and: st.and, conds } } : null);
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal cf-dlg" role="dialog" aria-label="Пользовательский автофильтр" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') ok(); }}>
        <div className="modal-title">Пользовательский автофильтр</div>
        <div className="modal-body">
          <div className="cf-head">Показать строки, у которых «{head}»:</div>
          {st.rows.map((r, i) => (
            <React.Fragment key={i}>
              {i === 1 && (
                <div className="cf-andor">
                  <label><input type="radio" checked={st.and} onChange={() => setSt({ ...st, and: true })} /> И</label>
                  <label><input type="radio" checked={!st.and} onChange={() => setSt({ ...st, and: false })} /> ИЛИ</label>
                </div>
              )}
              <div className="cf-row">
                <select value={r.label} onChange={(e) => set(i, { label: e.target.value })}>
                  {i > 0 && <option value="" />}
                  {ops.map(([l]) => <option key={l} value={l}>{l}</option>)}
                </select>
                <input className="pd-input" value={r.text} autoFocus={i === 0} placeholder={kind === 'date' ? 'ДД.ММ.ГГГГ' : ''} onChange={(e) => set(i, { text: e.target.value })} />
              </div>
            </React.Fragment>
          ))}
          {kind === 'text' && <div className="cf-hint">«?» — любой один знак, «*» — любая последовательность знаков</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
