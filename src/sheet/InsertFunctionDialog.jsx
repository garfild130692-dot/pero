// Мастер функций «Вставка функции» (fx, Shift+F3): поиск и категории → поля аргументов с предпросмотром значений
import React, { useMemo, useState } from 'react';
import { evaluate, fromUser, parse } from './formula.js';
import { CATEGORIES, FN_INFO, FN_LIST, signature } from './fninfo.js';
import { showValue } from './evalSteps.js';

const RECENT_KEY = 'pero.recentFunctions';
const recent = () => { try { return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); } catch { return []; } };
const remember = (en) => { try { localStorage.setItem(RECENT_KEY, JSON.stringify([en, ...recent().filter((x) => x !== en)].slice(0, 10))); } catch { /* */ } };
const DEFAULT_RECENT = ['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNTIF', 'SUMIF', 'MAX', 'MIN', 'ROUND', 'XLOOKUP'];

// значение выражения аргумента для предпросмотра
function preview(text, ctx) {
  if (!text.trim()) return '';
  try { return showValue(evaluate(parse(fromUser(text)), ctx)); } catch { return 'ошибка записи'; }
}

export default function InsertFunctionDialog({ ctx, initial, onInsert, onClose }) {
  const [step, setStep] = useState(initial ? 2 : 1);
  const [query, setQuery] = useState('');
  const [cat, setCat] = useState('Недавно использовавшиеся');
  const [pick, setPick] = useState(initial || null);
  const [args, setArgs] = useState(() => (initial ? FN_INFO[initial].args.map(() => '') : []));
  const list = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (q) return FN_LIST.filter((f) => f.ru.toLowerCase().includes(q) || f.en.toLowerCase().includes(q) || f.description.toLowerCase().includes(q));
    if (cat === 'Полный алфавитный перечень') return FN_LIST;
    if (cat === 'Недавно использовавшиеся') return [...new Set([...recent(), ...DEFAULT_RECENT])].map((en) => FN_INFO[en]).filter(Boolean);
    return FN_LIST.filter((f) => f.category === cat);
  }, [query, cat]);
  // выбранная функция — или первая в списке (поиск сразу выбирает подходящую)
  const current = pick && list.some((f) => f.en === pick) ? pick : list[0] ? list[0].en : null;
  const info = step === 2 ? pick && FN_INFO[pick] : current && FN_INFO[current];
  const choose = (en) => { setPick(en); };
  const next = () => {
    if (!info) return;
    setPick(info.en);
    setArgs(info.args.map(() => ''));
    setStep(2);
  };
  // поля аргументов: обязательные и необязательные; у повторяющихся — новое поле, когда заполнено последнее
  const fields = info ? args.map((v, i) => ({ v, name: i < info.args.length ? info.args[i].name : info.args[info.args.length - 1].name.replace(/\d+$/, '') + (i + 1), optional: i < info.args.length ? info.args[i].optional : true })) : [];
  const setArg = (i, v) => {
    const a = [...args]; a[i] = v;
    if (info.more && i === a.length - 1 && v.trim() && a.length < 60) a.push('');
    setArgs(a);
  };
  const formulaText = info ? `=${info.ru}(${args.slice(0, args.reduce((n, v, i) => (v.trim() ? i + 1 : n), 0)).join(';')})` : '';
  const result = info ? preview(formulaText.slice(1), ctx) : '';
  const insert = () => { remember(info.en); onInsert(formulaText); };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal fn-dlg" role="dialog" aria-label={step === 1 ? 'Вставка функции' : 'Аргументы функции'} onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        {step === 1 ? (
          <>
            <div className="modal-title">Вставка функции</div>
            <div className="modal-body">
              <label className="pd-field"><span>Поиск функции</span>
                <input className="pd-input" autoFocus value={query} placeholder="Например: сумма, дата, ВПР" onChange={(e) => setQuery(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter') next(); }} style={{ width: 320 }} />
              </label>
              <label className="pd-field" style={{ marginTop: 6 }}><span>Категория</span>
                <select value={cat} disabled={!!query.trim()} onChange={(e) => setCat(e.target.value)} style={{ width: 320 }}>
                  {['Недавно использовавшиеся', 'Полный алфавитный перечень', ...CATEGORIES].map((c) => <option key={c}>{c}</option>)}
                </select>
              </label>
              <div className="fn-list" role="listbox" aria-label="Функции">
                {list.map((f) => (
                  <button key={f.en} type="button" className={`fn-row${current === f.en ? ' on' : ''}`} onClick={() => choose(f.en)} onDoubleClick={() => { setPick(f.en); setArgs(f.args.map(() => '')); setStep(2); }}>{f.ru}</button>
                ))}
                {!list.length && <div className="tabs-empty">Ничего не найдено</div>}
              </div>
              {info && <div className="fn-help"><b>{signature(info)}</b><br />{info.description}</div>}
            </div>
            <div className="modal-actions">
              <button type="button" className="btn primary" disabled={!info} onClick={next}>Далее</button>
              <button type="button" className="btn" onClick={onClose}>Отмена</button>
            </div>
          </>
        ) : (
          <>
            <div className="modal-title">Аргументы функции — {info.ru}</div>
            <div className="modal-body">
              <div className="fn-args">
                {fields.map((f, i) => (
                  <label key={i} className="fn-arg">
                    <span className={f.optional ? 'opt' : ''}>{f.name}</span>
                    <input className="pd-input" autoFocus={i === 0} value={f.v} onChange={(e) => setArg(i, e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter') insert(); }} />
                    <span className="fn-val">{f.v.trim() ? `= ${preview(f.v, ctx)}` : ''}</span>
                  </label>
                ))}
                {!fields.length && <div className="tabs-empty">У функции нет аргументов.</div>}
              </div>
              <div className="fn-help">{info.description}</div>
              <div className="fn-help"><b>Формула:</b> {formulaText}<br /><b>Значение:</b> {result}</div>
            </div>
            <div className="modal-actions">
              {!initial && <button type="button" className="btn" onClick={() => setStep(1)}>Назад</button>}
              <span className="spacer" />
              <button type="button" className="btn primary" onClick={insert}>ОК</button>
              <button type="button" className="btn" onClick={onClose}>Отмена</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
