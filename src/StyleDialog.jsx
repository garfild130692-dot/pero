// Диалог стиля абзаца: создать из текущего абзаца, изменить существующий или «Обычный».
import React, { useState } from 'react';
import { COMMON_FONTS, normalizeColor, parseSize } from './editor/font.js';
import { effectiveFont } from './editor/extensions.js';
import { effectivePara } from './editor/paraFormat.js';
import { cmToPt, fmtNum, paraBase, ptToCm } from './editor/paragraph.js';
import { newStyleId, styleOf } from './editor/styles.js';

const num = (s) => {
  const v = parseFloat(String(s).replace(',', '.'));
  return isFinite(v) ? v : null;
};

// Действующее оформление текущего абзаца — начальные значения нового стиля
function currentLook(editor) {
  const { state } = editor;
  const node = state.selection.$from.parent;
  const p = ['paragraph', 'heading'].includes(node.type.name) ? effectivePara(node, state.doc.attrs) : paraBase(state.doc.attrs.para);
  let marks = [];
  node.forEach((c) => { if (!marks.length && c.isText) marks = c.marks; });
  const ts = (marks.find((m) => m.type.name === 'textStyle') || {}).attrs || {};
  const f = effectiveFont(ts, node, state.doc.attrs);
  const has = (t) => marks.some((m) => m.type.name === t);
  return { p, run: { fontFamily: f.fontFamily, fontSize: f.fontSize, color: ts.color || null, bold: has('bold'), italic: has('italic'), underline: has('underline') }, align: node.attrs.textAlign || null };
}

export default function StyleDialog({ editor, mode, onClose }) {
  // mode: 'new' | 'normal' | id стиля
  const attrs = editor.state.doc.attrs;
  const isNormal = mode === 'normal';
  const existing = !isNormal && mode !== 'new' ? styleOf(attrs, mode) : null;
  const [f, setF] = useState(() => {
    const docBase = paraBase(attrs.para);
    let p; let run; let align; let name;
    if (isNormal) { p = docBase; run = { fontFamily: attrs.fontFamily, fontSize: attrs.fontSize }; align = null; name = 'Обычный'; }
    else if (existing) {
      p = { ...docBase, ...existing.para };
      run = { fontFamily: attrs.fontFamily, fontSize: attrs.fontSize, ...existing.run };
      align = existing.align; name = existing.name;
    } else { const c = currentLook(editor); p = c.p; run = c.run; align = c.align; name = ''; }
    return {
      name,
      fontFamily: run.fontFamily || '', fontSize: fmtNum(run.fontSize || 11, 1),
      bold: !!run.bold, italic: !!run.italic, underline: !!run.underline,
      color: run.color || '', align: align || 'left',
      before: fmtNum(p.spaceBefore, 1), after: fmtNum(p.spaceAfter, 1),
      line: p.lineRule === 'auto' ? fmtNum(p.lineSpacing, 2) : '1',
      first: fmtNum(ptToCm(p.indentFirst), 2), left: fmtNum(ptToCm(p.indentLeft), 2),
    };
  });
  const [err, setErr] = useState('');
  const set = (k) => (e) => { setErr(''); setF({ ...f, [k]: e.target.type === 'checkbox' ? e.target.checked : e.target.value }); };

  const save = () => {
    const size = parseSize(f.fontSize);
    const before = num(f.before); const after = num(f.after); const line = num(f.line);
    const first = num(f.first); const left = num(f.left);
    if (!isNormal && !f.name.trim()) return setErr('Укажите имя стиля');
    if (!size) return setErr('Размер шрифта — число от 1 до 1638');
    if ([before, after].some((v) => v === null || v < 0 || v > 1584)) return setErr('Интервалы — от 0 до 1584 пт');
    if (line === null || line < 0.06 || line > 132) return setErr('Междустрочный множитель — от 0,06 до 132');
    if ([first, left].some((v) => v === null || Math.abs(v) > 55)) return setErr('Отступы — от −55 до 55 см');
    const paraVals = {
      spaceBefore: Math.round(before * 20) / 20, spaceAfter: Math.round(after * 20) / 20,
      lineRule: 'auto', lineSpacing: Math.round(line * 240) / 240, indentFirst: cmToPt(first), indentLeft: cmToPt(left),
    };
    if (isNormal) {
      editor.chain().focus().setNormalStyle({ fontFamily: f.fontFamily.trim() || attrs.fontFamily, fontSize: size, para: paraVals }).run();
      return onClose();
    }
    // в стиле храним только отличия от «Обычного»
    const docBase = paraBase(attrs.para);
    const para = {};
    for (const [k, v] of Object.entries(paraVals)) if (k !== 'lineRule' && Math.abs(v - docBase[k]) > 1e-6) para[k] = v;
    if (para.lineSpacing != null) para.lineRule = 'auto';
    const run = {};
    if (f.fontFamily.trim() && f.fontFamily.trim() !== attrs.fontFamily) run.fontFamily = f.fontFamily.trim();
    if (size !== attrs.fontSize) run.fontSize = size;
    const color = normalizeColor(f.color);
    if (color) run.color = color;
    for (const k of ['bold', 'italic', 'underline']) if (f[k]) run[k] = true;
    const id = existing ? existing.id : newStyleId(attrs.styles);
    const chain = editor.chain().focus().saveStyle({ id, name: f.name.trim(), para, run, align: f.align === 'left' ? null : f.align });
    if (!existing) chain.setParagraphStyle(id);
    chain.run();
    return onClose();
  };

  const remove = () => { editor.chain().focus().deleteStyle(existing.id).run(); onClose(); };
  const title = isNormal ? 'Изменить стиль «Обычный»' : existing ? `Изменить стиль «${existing.name}»` : 'Создать стиль';
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label={title} onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); save(); } }}>
        <div className="modal-title">{title}</div>
        <div className="modal-body">
          {!isNormal && (
            <label className="pd-field"><span>Имя</span>
              <span className="pd-input wide"><input autoFocus value={f.name} onChange={set('name')} placeholder="Например: Основной текст" /></span>
            </label>
          )}
          <fieldset><legend>Шрифт</legend>
            <div className="pd-grid">
              <label className="pd-field"><span>Шрифт</span><span className="pd-input"><input list="style-fonts" value={f.fontFamily} onChange={set('fontFamily')} /></span></label>
              <label className="pd-field"><span>Размер</span><span className="pd-input"><input value={f.fontSize} onChange={set('fontSize')} /><em>пт</em></span></label>
            </div>
            <datalist id="style-fonts">{COMMON_FONTS.map((x) => <option key={x} value={x} />)}</datalist>
            {!isNormal && (
              <div className="style-flags">
                <label><input type="checkbox" checked={f.bold} onChange={set('bold')} /> <b>Полужирный</b></label>
                <label><input type="checkbox" checked={f.italic} onChange={set('italic')} /> <i>Курсив</i></label>
                <label><input type="checkbox" checked={f.underline} onChange={set('underline')} /> <u>Подчёркнутый</u></label>
                <label>Цвет <input type="color" value={normalizeColor(f.color) || '#000000'} onChange={set('color')} /></label>
                {f.color && <button type="button" className="btn small" onClick={() => setF({ ...f, color: '' })}>Авто</button>}
              </div>
            )}
          </fieldset>
          <fieldset><legend>Абзац</legend>
            <div className="pd-grid">
              {!isNormal && (
                <label className="pd-field"><span>Выравнивание</span>
                  <select value={f.align} onChange={set('align')}>
                    <option value="left">По левому краю</option><option value="center">По центру</option>
                    <option value="right">По правому краю</option><option value="justify">По ширине</option>
                  </select>
                </label>
              )}
              <label className="pd-field"><span>Междустрочный</span><span className="pd-input"><input value={f.line} onChange={set('line')} /><em>×</em></span></label>
              <label className="pd-field"><span>Перед</span><span className="pd-input"><input value={f.before} onChange={set('before')} /><em>пт</em></span></label>
              <label className="pd-field"><span>После</span><span className="pd-input"><input value={f.after} onChange={set('after')} /><em>пт</em></span></label>
              <label className="pd-field"><span>Отступ слева</span><span className="pd-input"><input value={f.left} onChange={set('left')} /><em>см</em></span></label>
              <label className="pd-field"><span>Первая строка</span><span className="pd-input"><input value={f.first} onChange={set('first')} /><em>см</em></span></label>
            </div>
          </fieldset>
          {err && <div className="pd-error" role="alert">{err}</div>}
        </div>
        <div className="modal-actions">
          {existing && <button className="btn danger" onClick={remove}>Удалить стиль</button>}
          <span className="spacer" />
          <button className="btn primary" onClick={save}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
