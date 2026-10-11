// Диалоги таблиц: «Вставка таблицы», «Преобразовать в таблицу», «Преобразовать в текст» — как в Word.
import React, { useState } from 'react';
import { columnsFor, guessSeparator } from './editor/tableConvert.js';

const NumIn = ({ value, onChange, min = 1, max = 63 }) => (
  <span className="pd-input small"><input type="number" min={min} max={max} value={value} onChange={(e) => onChange(Math.max(min, Math.min(max, +e.target.value || min)))} /></span>
);

function Shell({ title, onClose, onOk, okDisabled, children }) {
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label={title} onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); onOk(); } }}>
        <div className="modal-title">{title}</div>
        <div className="modal-body">{children}</div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" disabled={okDisabled} onClick={onOk}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

export function InsertTableDialog({ editor, onClose }) {
  const [cols, setCols] = useState(5);
  const [rows, setRows] = useState(2);
  const [fit, setFit] = useState('auto');
  const [width, setWidth] = useState('3');
  const [header, setHeader] = useState(false);
  const ok = () => {
    const w = parseFloat(String(width).replace(',', '.'));
    editor.chain().focus().insertTableSized(rows, cols, { widthCm: fit === 'fixed' && w > 0 ? w : null, header }).run();
    onClose();
  };
  return (
    <Shell title="Вставка таблицы" onClose={onClose} onOk={ok}>
      <fieldset><legend>Размер таблицы</legend>
        <div className="pd-grid">
          <label className="pd-field"><span>Число столбцов</span><NumIn value={cols} onChange={setCols} /></label>
          <label className="pd-field"><span>Число строк</span><NumIn value={rows} onChange={setRows} max={500} /></label>
        </div>
      </fieldset>
      <fieldset><legend>Ширина столбцов</legend>
        <label className="pd-check"><input type="radio" checked={fit === 'auto'} onChange={() => setFit('auto')} /> По ширине текста (поровну)</label>
        <label className="pd-check"><input type="radio" checked={fit === 'fixed'} onChange={() => setFit('fixed')} /> Постоянная:
          <span className="pd-input small" style={{ marginLeft: 6 }}><input value={width} disabled={fit !== 'fixed'} onChange={(e) => setWidth(e.target.value)} /><em>см</em></span>
        </label>
      </fieldset>
      <label className="pd-check"><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> Первая строка — заголовок (повторять на каждой странице)</label>
    </Shell>
  );
}

const SEPS = [['tab', 'Знак табуляции'], [';', 'Точка с запятой'], [',', 'Запятая'], ['para', 'Знак абзаца'], ['other', 'Другой:']];

function SepPicker({ sep, setSep, other, setOther }) {
  return (
    <fieldset><legend>Разделитель</legend>
      <div className="conv-seps">
        {SEPS.map(([k, l]) => (
          <label key={k} className="pd-check"><input type="radio" checked={sep === k} onChange={() => setSep(k)} /> {l}
            {k === 'other' && <span className="pd-input small" style={{ width: 40, marginLeft: 6 }}><input maxLength={1} value={other} onChange={(e) => { setOther(e.target.value); setSep('other'); }} /></span>}
          </label>
        ))}
      </div>
    </fieldset>
  );
}

export function TextToTableDialog({ editor, onClose }) {
  const [sep, setSep] = useState(() => guessSeparator(editor.state));
  const [other, setOther] = useState('-');
  const real = sep === 'other' ? other || '-' : sep;
  const auto = columnsFor(editor.state, real);
  const [cols, setCols] = useState(null);
  const ok = () => {
    if (editor.chain().focus().textToTable({ sep: real, cols: cols || auto }).run()) onClose();
  };
  return (
    <Shell title="Преобразовать в таблицу" onClose={onClose} onOk={ok} okDisabled={!auto}>
      {!auto && <div className="pd-error">Выделите абзацы с текстом (без таблиц и списков).</div>}
      <label className="pd-field"><span>Число столбцов</span><NumIn value={cols || auto || 1} onChange={setCols} /></label>
      <SepPicker sep={sep} setSep={(s) => { setSep(s); setCols(null); }} other={other} setOther={setOther} />
    </Shell>
  );
}

export function TableToTextDialog({ editor, onClose }) {
  const [sep, setSep] = useState('tab');
  const [other, setOther] = useState('-');
  const ok = () => { if (editor.chain().focus().tableToText({ sep: sep === 'other' ? other || '-' : sep }).run()) onClose(); };
  return (
    <Shell title="Преобразовать в текст" onClose={onClose} onOk={ok}>
      <SepPicker sep={sep} setSep={setSep} other={other} setOther={setOther} />
    </Shell>
  );
}
