// «Найти и заменить» (Ctrl+F / Ctrl+H) — окно поверх книги, раскладка как в Excel: вкладки «Найти» / «Заменить»,
// «Параметры >>» (искать на листе/в книге, по строкам/столбцам, в формулах/значениях, регистр, ячейка целиком),
// «Найти все» — список совпадений, щелчок — переход.
import React, { useState } from 'react';
import { addr, key } from './address.js';
import { findAll, findNext, replaceAll, replacedInput } from './find.js';
import { setInputs } from './model.js';

export default function FindDialog({ wb, si, sel, mode, onGo, onChange, onClose, onMessage }) {
  const [tab, setTab] = useState(mode || 'find');
  const [what, setWhat] = useState('');
  const [by, setBy] = useState('');
  const [more, setMore] = useState(false);
  const [opts, setOpts] = useState({ within: 'sheet', by: 'rows', lookIn: 'formulas', matchCase: false, whole: false });
  const [list, setList] = useState(null);
  const o = { ...opts, si, lookIn: tab === 'replace' ? 'formulas' : opts.lookIn };
  const here = { si, r: sel.r, c: sel.c };
  const next = (back = false) => {
    const x = findNext(wb, what, here, { ...o, back });
    if (!x) { onMessage(`Не удаётся найти «${what}»`); return null; }
    onGo(x); return x;
  };
  const all = () => {
    const l = findAll(wb, what, o);
    setList(l);
    if (!l.length) onMessage(`Не удаётся найти «${what}»`);
    else onGo(l[0]);
  };
  const replaceOne = () => {
    const cell = wb.sheets[si].cells.get(key(sel.r, sel.c));
    const raw = cell && replacedInput(cell, what, by, o);
    if (raw == null) { next(); return; }
    // заменить в текущей ячейке и перейти к следующему совпадению — уже в новой книге
    const nwb = replaceAt(wb, si, sel.r, sel.c, raw);
    onChange(() => nwb);
    const x = findNext(nwb, what, here, o);
    if (x) onGo(x); else onMessage('Больше совпадений нет');
  };
  const replaceEvery = () => {
    const res = replaceAll(wb, what, by, o);
    if (res.count) onChange(() => res.wb);
    onMessage(res.count ? `Произведено замен: ${res.count}` : `Не удаётся найти «${what}»`);
    setList(null);
  };
  const setOpt = (patch) => { setOpts({ ...opts, ...patch }); setList(null); };
  return (
    <div className="find-dlg" role="dialog" aria-label="Найти и заменить" onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); next(e.shiftKey); } }}>
      <div className="find-head">
        <span>Найти и заменить</span>
        <button type="button" className="find-x" title="Закрыть (Esc)" onClick={onClose}>✕</button>
      </div>
      <div className="fc-tabs">
        <button type="button" className={`fc-tab${tab === 'find' ? ' on' : ''}`} onClick={() => setTab('find')}>Найти</button>
        <button type="button" className={`fc-tab${tab === 'replace' ? ' on' : ''}`} onClick={() => setTab('replace')}>Заменить</button>
      </div>
      <div className="find-body">
        <label className="find-row"><span>Найти:</span><input className="pd-input" autoFocus value={what} onChange={(e) => { setWhat(e.target.value); setList(null); }} /></label>
        {tab === 'replace' && <label className="find-row"><span>Заменить на:</span><input className="pd-input" value={by} onChange={(e) => setBy(e.target.value)} /></label>}
        {more && (
          <div className="find-opts">
            <label>Искать: <select value={opts.within} onChange={(e) => setOpt({ within: e.target.value })}><option value="sheet">на листе</option><option value="book">в книге</option></select></label>
            <label>Просматривать: <select value={opts.by} onChange={(e) => setOpt({ by: e.target.value })}><option value="rows">по строкам</option><option value="cols">по столбцам</option></select></label>
            {tab === 'find' && <label>Область поиска: <select value={opts.lookIn} onChange={(e) => setOpt({ lookIn: e.target.value })}><option value="formulas">формулы</option><option value="values">значения</option></select></label>}
            <label className="fc-check"><input type="checkbox" checked={opts.matchCase} onChange={(e) => setOpt({ matchCase: e.target.checked })} /> Учитывать регистр</label>
            <label className="fc-check"><input type="checkbox" checked={opts.whole} onChange={(e) => setOpt({ whole: e.target.checked })} /> Ячейка целиком</label>
          </div>
        )}
        <div className="find-actions">
          <button type="button" className="btn small" onClick={() => setMore(!more)}>{more ? 'Параметры <<' : 'Параметры >>'}</button>
          <span className="spacer" />
          {tab === 'replace' && <button type="button" className="btn small" disabled={!what} onClick={replaceEvery}>Заменить все</button>}
          {tab === 'replace' && <button type="button" className="btn small" disabled={!what} onClick={replaceOne}>Заменить</button>}
          <button type="button" className="btn small" disabled={!what} onClick={all}>Найти все</button>
          <button type="button" className="btn small primary" disabled={!what} onClick={() => next()}>Найти далее</button>
        </div>
        {list && list.length > 0 && (
          <div className="find-list">
            <div className="find-item head"><span>Лист</span><span>Ячейка</span><span>Значение</span></div>
            {list.map((x) => (
              <button key={`${x.si}:${x.r},${x.c}`} type="button" className={`find-item${x.si === si && x.r === sel.r && x.c === sel.c ? ' on' : ''}`} onClick={() => onGo(x)}>
                <span>{wb.sheets[x.si].name}</span><span>{addr(x.r, x.c)}</span><span>{x.text}</span>
              </button>
            ))}
            <div className="find-count">Найдено ячеек: {list.length}</div>
          </div>
        )}
      </div>
    </div>
  );
}
// замена в одной ячейке — тем же пакетным вводом, что и «Заменить все»
function replaceAt(wb, si, r, c, raw) { return setInputs(wb, [{ si, r, c, raw }]); }
