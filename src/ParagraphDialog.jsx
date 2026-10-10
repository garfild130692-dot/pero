// Диалог «Абзац»: выравнивание, отступы (см), интервалы (pt), межстрочный интервал; вкладка «Положение на странице».
// Начальные значения — действующие параметры первого абзаца выделения; применяются только изменённые поля.
import React, { useEffect, useState } from 'react';
import { effectivePara } from './editor/paraFormat.js';
import { cmToPt, fmtNum, ptToCm } from './editor/paragraph.js';
import { LAYOUT_KEYS, effectiveLayout } from './editor/paraLayout.js';

const LINE_TYPES = [
  ['single', 'Одинарный'], ['1.5', '1,5 строки'], ['double', 'Двойной'],
  ['atLeast', 'Минимум'], ['exact', 'Точно'], ['multiple', 'Множитель'],
];
const ALIGNS = [['left', 'По левому краю'], ['center', 'По центру'], ['right', 'По правому краю'], ['justify', 'По ширине']];

const num = (s) => {
  const v = parseFloat(String(s).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isFinite(v) ? v : null;
};

function lineToForm(rule, value) {
  if (rule === 'exact' || rule === 'atLeast') return { type: rule, value: fmtNum(value, 2) };
  if (Math.abs(value - 1) < 1e-6) return { type: 'single', value: '' };
  if (Math.abs(value - 1.5) < 1e-6) return { type: '1.5', value: '' };
  if (Math.abs(value - 2) < 1e-6) return { type: 'double', value: '' };
  return { type: 'multiple', value: fmtNum(value, 2) };
}

// Текстовое поле с шагом по стрелкам ↑/↓ (как счётчики в Word)
function NumField({ label, value, onChange, unit, step, min, max, disabled, id }) {
  const bump = (dir) => {
    const v = num(value) ?? 0;
    const n = Math.min(max, Math.max(min, Math.round((v + dir * step) * 100) / 100));
    onChange(fmtNum(n, 2));
  };
  return (
    <label className="pd-field" htmlFor={id}>
      <span>{label}</span>
      <span className="pd-input">
        <input
          id={id}
          value={value}
          disabled={disabled}
          inputMode="decimal"
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'ArrowUp') { e.preventDefault(); bump(1); }
            if (e.key === 'ArrowDown') { e.preventDefault(); bump(-1); }
          }}
        />
        {unit && <em>{unit}</em>}
      </span>
    </label>
  );
}

export default function ParagraphDialog({ editor, onClose }) {
  const [init] = useState(() => {
    const { state } = editor;
    const $from = state.selection.$from;
    let node = $from.parent;
    let inList = false;
    for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.name === 'listItem') inList = true;
    if (!['paragraph', 'heading'].includes(node.type.name)) node = state.schema.nodes.paragraph.create();
    const p = effectivePara(node, state.doc.attrs, inList);
    const line = lineToForm(p.lineRule, p.lineSpacing);
    return {
      align: node.attrs.textAlign || 'left',
      left: fmtNum(ptToCm(p.indentLeft), 2),
      right: fmtNum(ptToCm(p.indentRight), 2),
      firstType: p.indentFirst > 0 ? 'first' : p.indentFirst < 0 ? 'hanging' : 'none',
      firstBy: p.indentFirst ? fmtNum(ptToCm(Math.abs(p.indentFirst)), 2) : '',
      before: fmtNum(p.spaceBefore, 1),
      after: fmtNum(p.spaceAfter, 1),
      lineType: line.type,
      lineValue: line.value,
      ...effectiveLayout(node), // widowControl, keepNext, keepLines, pageBreakBefore
    };
  });
  const [tab, setTab] = useState('indent');
  const [f, setF] = useState(init);
  const set = (k) => (v) => setF((s) => ({ ...s, [k]: v }));
  const [error, setError] = useState('');

  useEffect(() => { setError(''); }, [f]);

  const apply = () => {
    const changed = (k) => f[k] !== init[k];
    const patch = {};
    const bad = (msg) => { setError(msg); return true; };

    if (changed('left')) { const v = num(f.left); if (v === null || Math.abs(v) > 55) return bad('Отступ слева: число от −55 до 55 см'); patch.indentLeft = cmToPt(v); }
    if (changed('right')) { const v = num(f.right); if (v === null || Math.abs(v) > 55) return bad('Отступ справа: число от −55 до 55 см'); patch.indentRight = cmToPt(v); }
    if (changed('firstType') || changed('firstBy')) {
      if (f.firstType === 'none') patch.indentFirst = 0;
      else {
        const v = num(f.firstBy);
        if (v === null || v < 0 || v > 55) return bad('Первая строка: число от 0 до 55 см');
        patch.indentFirst = cmToPt(f.firstType === 'hanging' ? -v : v);
      }
    }
    if (changed('before')) { const v = num(f.before); if (v === null || v < 0 || v > 1584) return bad('Интервал перед: число от 0 до 1584 пт'); patch.spaceBefore = Math.round(v * 20) / 20; }
    if (changed('after')) { const v = num(f.after); if (v === null || v < 0 || v > 1584) return bad('Интервал после: число от 0 до 1584 пт'); patch.spaceAfter = Math.round(v * 20) / 20; }
    if (changed('lineType') || changed('lineValue')) {
      const t = f.lineType;
      if (t === 'single') Object.assign(patch, { lineRule: 'auto', lineSpacing: 1 });
      else if (t === '1.5') Object.assign(patch, { lineRule: 'auto', lineSpacing: 1.5 });
      else if (t === 'double') Object.assign(patch, { lineRule: 'auto', lineSpacing: 2 });
      else {
        const v = num(f.lineValue);
        if (t === 'multiple' && (v === null || v < 0.06 || v > 132)) return bad('Множитель: число от 0,06 до 132');
        if (t !== 'multiple' && (v === null || v < 0.7 || v > 1584)) return bad('Значение: число от 0,7 до 1584 пт');
        Object.assign(patch, { lineRule: t === 'multiple' ? 'auto' : t, lineSpacing: v });
      }
    }
    const layout = {};
    for (const k of LAYOUT_KEYS) if (changed(k)) layout[k] = f[k];
    const chain = editor.chain().focus();
    if (Object.keys(patch).length) chain.setParagraphFormat(patch);
    if (Object.keys(layout).length) chain.setParagraphLayout(layout);
    if (changed('align')) chain.setTextAlign(f.align);
    chain.run();
    onClose();
    return true;
  };

  // при смене вида интервала — подставить значение по умолчанию, как Word
  const onLineType = (t) => setF((s) => ({
    ...s,
    lineType: t,
    lineValue: t === 'multiple' ? (s.lineType === 'multiple' ? s.lineValue : '3') : t === 'exact' || t === 'atLeast' ? (['exact', 'atLeast'].includes(s.lineType) ? s.lineValue : '12') : '',
  }));
  const onFirstType = (t) => setF((s) => ({ ...s, firstType: t, firstBy: t === 'none' ? '' : s.firstBy || '1,25' }));

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Абзац" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Абзац</div>
        <div className="pd-tabs" role="tablist">
          <button type="button" role="tab" className={tab === 'indent' ? 'active' : ''} onClick={() => setTab('indent')}>Отступы и интервалы</button>
          <button type="button" role="tab" className={tab === 'position' ? 'active' : ''} onClick={() => setTab('position')}>Положение на странице</button>
        </div>
        <div className="modal-body">
          {tab === 'position' && (
            <fieldset>
              <legend>Разбивка на страницы</legend>
              <div className="pd-checks">
                {[
                  ['widowControl', 'Запрет висячих строк', 'Первая и последняя строки абзаца не остаются на странице поодиночке'],
                  ['keepNext', 'Не отрывать от следующего', 'Абзац и следующий за ним — на одной странице'],
                  ['keepLines', 'Не разрывать абзац', 'Все строки абзаца — на одной странице'],
                  ['pageBreakBefore', 'С новой страницы', 'Абзац начинается с новой страницы'],
                ].map(([k, label, hint]) => (
                  <label key={k} className="pd-check" title={hint}>
                    <input type="checkbox" checked={!!f[k]} onChange={(e) => set(k)(e.target.checked)} /> {label}
                  </label>
                ))}
              </div>
            </fieldset>
          )}
          {tab === 'indent' && (<>
          <fieldset>
            <legend>Общие</legend>
            <label className="pd-field">
              <span>Выравнивание</span>
              <select value={f.align} onChange={(e) => set('align')(e.target.value)}>
                {ALIGNS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
            </label>
          </fieldset>
          <fieldset>
            <legend>Отступ</legend>
            <div className="pd-grid">
              <NumField id="pd-left" label="Слева" unit="см" value={f.left} onChange={set('left')} step={0.1} min={-55} max={55} />
              <label className="pd-field">
                <span>Первая строка</span>
                <select value={f.firstType} onChange={(e) => onFirstType(e.target.value)}>
                  <option value="none">(нет)</option>
                  <option value="first">Отступ</option>
                  <option value="hanging">Выступ</option>
                </select>
              </label>
              <NumField id="pd-right" label="Справа" unit="см" value={f.right} onChange={set('right')} step={0.1} min={-55} max={55} />
              <NumField id="pd-first" label="на" unit="см" value={f.firstBy} onChange={set('firstBy')} step={0.1} min={0} max={55} disabled={f.firstType === 'none'} />
            </div>
          </fieldset>
          <fieldset>
            <legend>Интервал</legend>
            <div className="pd-grid">
              <NumField id="pd-before" label="Перед" unit="пт" value={f.before} onChange={set('before')} step={6} min={0} max={1584} />
              <label className="pd-field">
                <span>Междустрочный</span>
                <select value={f.lineType} onChange={(e) => onLineType(e.target.value)}>
                  {LINE_TYPES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                </select>
              </label>
              <NumField id="pd-after" label="После" unit="пт" value={f.after} onChange={set('after')} step={6} min={0} max={1584} />
              <NumField id="pd-line" label="значение" unit={f.lineType === 'multiple' ? '' : 'пт'} value={f.lineValue} onChange={set('lineValue')}
                step={f.lineType === 'multiple' ? 0.5 : 1} min={0} max={1584} disabled={!['multiple', 'exact', 'atLeast'].includes(f.lineType)} />
            </div>
          </fieldset>
          </>)}
          {error && <div className="pd-error" role="alert">{error}</div>}
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
