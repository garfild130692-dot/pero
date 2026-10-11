// «Проверка вводимых значений» (Данные → Проверка данных) и окно ошибки ввода — как в Excel.
import React, { useState } from 'react';
import { fromUser, localize } from './formula.js';
import { parseInput } from './format.js';

const TYPES = [['any', 'Любое значение'], ['whole', 'Целое число'], ['decimal', 'Действительное'], ['list', 'Список'], ['date', 'Дата'], ['time', 'Время'], ['textLength', 'Длина текста'], ['custom', 'Другой']];
const OPS = [['between', 'между'], ['notBetween', 'вне'], ['equal', 'равно'], ['notEqual', 'не равно'], ['greaterThan', 'больше'], ['lessThan', 'меньше'], ['greaterThanOrEqual', 'больше или равно'], ['lessThanOrEqual', 'меньше или равно']];
const STYLES = [['stop', 'Останов'], ['warning', 'Предупреждение'], ['information', 'Сообщение']];

// формула правила ↔ поле ввода: «=$F$1:$F$3», «да;нет» (список: в поле «;», в файле «,»), число
const showFormula = (f, type) => {
  if (f == null) return '';
  const lit = /^"(.*)"$/s.exec(f);
  if (lit && type === 'list') return lit[1].replace(/""/g, '"').split(',').join(';');
  if (/^-?\d+(\.\d+)?$/.test(f)) return f.replace('.', ',');
  return `=${localize(f)}`;
};
const readFormula = (text, type) => {
  const t = String(text).trim();
  if (!t) return null;
  if (t.startsWith('=')) return fromUser(t.slice(1));
  if (type === 'list') return `"${t.split(';').map((x) => x.trim()).join(',').replace(/"/g, '""')}"`;
  const p = parseInput(t);
  return typeof p.value === 'number' ? String(p.value) : `"${t.replace(/"/g, '""')}"`;
};

export default function DvDialog({ rule, onApply, onClear, onClose }) {
  const r0 = rule || {};
  const [tab, setTab] = useState('set');
  const [st, setSt] = useState({
    type: r0.type || 'any', operator: r0.operator || 'between', f1: showFormula(r0.formula1, r0.type), f2: showFormula(r0.formula2, r0.type),
    allowBlank: rule ? r0.allowBlank : true, inCellDropdown: rule ? r0.inCellDropdown !== false : true,
    showPrompt: rule ? r0.showPrompt : true, promptTitle: r0.promptTitle || '', prompt: r0.prompt || '',
    showError: rule ? r0.showError : true, errorStyle: r0.errorStyle || 'stop', errorTitle: r0.errorTitle || '', error: r0.error || '',
  });
  const set = (patch) => setSt({ ...st, ...patch });
  const two = ['between', 'notBetween'].includes(st.operator) && !['list', 'custom', 'any'].includes(st.type);
  const label1 = st.type === 'list' ? 'Источник (значения через «;» или =диапазон):' : st.type === 'custom' ? 'Формула:' : two ? 'Минимум:' : 'Значение:';
  const ok = () => {
    if (st.type === 'any' && !st.showPrompt) { onClear(); return; }
    onApply({
      type: st.type, operator: st.operator, formula1: st.type === 'any' ? null : readFormula(st.f1, st.type), formula2: two ? readFormula(st.f2, st.type) : null,
      allowBlank: st.allowBlank, inCellDropdown: st.inCellDropdown, showPrompt: st.showPrompt, promptTitle: st.promptTitle, prompt: st.prompt,
      showError: st.showError, errorStyle: st.errorStyle, errorTitle: st.errorTitle, error: st.error,
    });
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal dv-dlg" role="dialog" aria-label="Проверка вводимых значений" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Проверка вводимых значений</div>
        <div className="fc-tabs">
          {[['set', 'Параметры'], ['prompt', 'Сообщение для ввода'], ['error', 'Сообщение об ошибке']].map(([k, l]) => <button key={k} type="button" className={`fc-tab${tab === k ? ' on' : ''}`} onClick={() => setTab(k)}>{l}</button>)}
        </div>
        <div className="modal-body dv-body">
          {tab === 'set' && (
            <>
              <div className="fc-sub">Условие проверки</div>
              <label className="pd-field"><span>Тип данных:</span><select value={st.type} onChange={(e) => set({ type: e.target.value })}>{TYPES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></label>
              {st.type !== 'any' && <label className="fc-check"><input type="checkbox" checked={st.allowBlank} onChange={(e) => set({ allowBlank: e.target.checked })} /> Игнорировать пустые ячейки</label>}
              {st.type === 'list' && <label className="fc-check"><input type="checkbox" checked={st.inCellDropdown} onChange={(e) => set({ inCellDropdown: e.target.checked })} /> Список допустимых значений</label>}
              {!['any', 'list', 'custom'].includes(st.type) && (
                <label className="pd-field"><span>Значение:</span><select value={st.operator} onChange={(e) => set({ operator: e.target.value })}>{OPS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></label>
              )}
              {st.type !== 'any' && <label className="pd-field dv-wide"><span>{label1}</span><input className="pd-input" autoFocus value={st.f1} onChange={(e) => set({ f1: e.target.value })} placeholder={st.type === 'list' ? 'да;нет  или  =$F$1:$F$10' : st.type === 'date' ? '01.01.2026' : ''} /></label>}
              {two && <label className="pd-field dv-wide"><span>Максимум:</span><input className="pd-input" value={st.f2} onChange={(e) => set({ f2: e.target.value })} /></label>}
            </>
          )}
          {tab === 'prompt' && (
            <>
              <label className="fc-check"><input type="checkbox" checked={st.showPrompt} onChange={(e) => set({ showPrompt: e.target.checked })} /> Отображать подсказку, если ячейка является текущей</label>
              <label className="pd-field dv-wide"><span>Заголовок:</span><input className="pd-input" value={st.promptTitle} onChange={(e) => set({ promptTitle: e.target.value })} /></label>
              <label className="pd-field dv-wide"><span>Сообщение:</span><textarea className="pd-input dv-text" value={st.prompt} onChange={(e) => set({ prompt: e.target.value })} /></label>
            </>
          )}
          {tab === 'error' && (
            <>
              <label className="fc-check"><input type="checkbox" checked={st.showError} onChange={(e) => set({ showError: e.target.checked })} /> Выводить сообщение об ошибке</label>
              <label className="pd-field"><span>Вид:</span><select value={st.errorStyle} onChange={(e) => set({ errorStyle: e.target.value })}>{STYLES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></label>
              <label className="pd-field dv-wide"><span>Заголовок:</span><input className="pd-input" value={st.errorTitle} onChange={(e) => set({ errorTitle: e.target.value })} /></label>
              <label className="pd-field dv-wide"><span>Сообщение:</span><textarea className="pd-input dv-text" value={st.error} onChange={(e) => set({ error: e.target.value })} /></label>
            </>
          )}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn" onClick={onClear}>Очистить все</button>
          <span className="spacer" />
          <button type="button" className="btn primary" onClick={ok}>ОК</button>
          <button type="button" className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// Окно ошибки ввода: stop — «Повторить»/«Отмена»; warning — «Продолжить?» Да/Нет/Отмена; information — ОК/Отмена
export function DvErrorDialog({ rule, onCommit, onRetry, onCancel }) {
  const style = rule.errorStyle || 'stop';
  const text = rule.error || 'Значение не соответствует ограничениям проверки данных, установленным для этой ячейки.';
  return (
    <div className="modal-backdrop" onMouseDown={onCancel}>
      <div className="modal dv-err" role="alertdialog" aria-label={rule.errorTitle || 'Перо'} onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Escape') onCancel(); }}>
        <div className="modal-title">{rule.errorTitle || 'Перо'}</div>
        <div className="modal-body dv-err-body">
          <span className={`dv-icon ${style}`} aria-hidden="true">{style === 'stop' ? '✕' : style === 'warning' ? '!' : 'i'}</span>
          <div>{text}{style === 'warning' && <div className="dv-ask">Продолжить?</div>}</div>
        </div>
        <div className="modal-actions">
          {style === 'stop' && <><button type="button" className="btn primary" autoFocus onClick={onRetry}>Повторить</button><button type="button" className="btn" onClick={onCancel}>Отмена</button></>}
          {style === 'warning' && <><button type="button" className="btn" onClick={onCommit}>Да</button><button type="button" className="btn primary" autoFocus onClick={onRetry}>Нет</button><button type="button" className="btn" onClick={onCancel}>Отмена</button></>}
          {style === 'information' && <><button type="button" className="btn primary" autoFocus onClick={onCommit}>ОК</button><button type="button" className="btn" onClick={onCancel}>Отмена</button></>}
        </div>
      </div>
    </div>
  );
}
