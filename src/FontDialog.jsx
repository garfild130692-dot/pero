// Диалог «Шрифт» (Ctrl+D, уголок группы «Шрифт»): вкладки «Шрифт» (гарнитура, начертание, размер, цвет,
// подчёркивание, видоизменение) и «Дополнительно» (межзнаковый интервал). Начальные значения — по выделению
// (пусто — в выделении разные); применяются только изменённые поля, как в Word.
import React, { useEffect, useMemo, useState } from 'react';
import { COMMON_FONTS, FONT_SIZES, cssFontStack, parseSize } from './editor/font.js';
import { UNDERLINES } from './editor/extensions.js';
import { fontState, loadSystemFonts } from './FontControls.jsx';
import { fmtNum } from './editor/paragraph.js';

const STYLES = [['regular', 'Обычный'], ['italic', 'Курсив'], ['bold', 'Полужирный'], ['boldItalic', 'Полужирный курсив']];
const num = (s) => {
  const v = parseFloat(String(s).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isFinite(v) ? v : null;
};

// Значение у всего выделения: true/false или '' (разные)
function markState(editor, name) {
  const { state } = editor;
  const { from, to, empty } = state.selection;
  const type = state.schema.marks[name];
  if (empty) return !!type.isInSet(state.storedMarks || state.selection.$from.marks());
  let on = 0; let off = 0;
  state.doc.nodesBetween(from, to, (n) => { if (n.isText) { if (type.isInSet(n.marks)) on++; else off++; } });
  return on && off ? '' : on > 0;
}
// Атрибут метки у выделения: одно значение или '' (разные)
function attrState(editor, mark, attr) {
  const { state } = editor;
  const { from, to, empty } = state.selection;
  const type = state.schema.marks[mark];
  const vals = new Set();
  const take = (marks) => { const m = type.isInSet(marks); vals.add(m ? m.attrs[attr] ?? null : null); };
  if (empty) take(state.storedMarks || state.selection.$from.marks());
  else state.doc.nodesBetween(from, to, (n) => { if (n.isText) take(n.marks); });
  return vals.size === 1 ? [...vals][0] : '';
}

export default function FontDialog({ editor, onClose }) {
  const [init] = useState(() => {
    const fs = fontState(editor);
    const b = markState(editor, 'bold'); const i = markState(editor, 'italic');
    const u = markState(editor, 'underline');
    const kind = u === true ? attrState(editor, 'underline', 'kind') : null;
    const sp = attrState(editor, 'textStyle', 'spacing');
    return {
      fontFamily: fs.fontFamily, fontSize: fs.fontSize === '' ? '' : String(fs.fontSize).replace('.', ','),
      style: b === '' || i === '' ? '' : b && i ? 'boldItalic' : b ? 'bold' : i ? 'italic' : 'regular',
      color: fs.color, // '' — авто (или разные)
      underline: u === '' ? '' : u ? (kind || 'single') : 'none',
      underlineColor: u === true ? attrState(editor, 'underline', 'color') || '' : '',
      strike: markState(editor, 'strike'), superscript: markState(editor, 'superscript'), subscript: markState(editor, 'subscript'),
      caps: attrState(editor, 'textStyle', 'caps') === '' ? '' : !!attrState(editor, 'textStyle', 'caps'),
      smallCaps: attrState(editor, 'textStyle', 'smallCaps') === '' ? '' : !!attrState(editor, 'textStyle', 'smallCaps'),
      spacingType: sp === '' ? '' : !sp ? 'normal' : sp > 0 ? 'expanded' : 'condensed',
      spacingBy: sp ? fmtNum(Math.abs(sp), 2) : '',
    };
  });
  const [f, setF] = useState(init);
  const set = (k) => (v) => setF((s) => ({ ...s, [k]: v }));
  const [tab, setTab] = useState('font');
  const [error, setError] = useState('');
  const [fonts, setFonts] = useState(COMMON_FONTS);
  useEffect(() => { loadSystemFonts().then(setFonts); }, []);
  useEffect(() => { setError(''); }, [f]);
  const fontList = useMemo(() => [...new Set(fonts)].sort((a, b) => a.localeCompare(b, 'ru')), [fonts]);

  const apply = () => {
    const changed = (k) => f[k] !== init[k];
    const bad = (msg) => { setError(msg); return false; };
    let size = null;
    if (changed('fontSize') && f.fontSize !== '') { size = parseSize(f.fontSize); if (!size) return bad('Размер: число от 1 до 1638'); }
    let spacing;
    if ((changed('spacingType') || changed('spacingBy')) && f.spacingType !== '') {
      if (f.spacingType === 'normal') spacing = null;
      else {
        const v = num(f.spacingBy);
        if (v === null || v <= 0 || v > 1584) return bad('Интервал: число от 0 до 1584 пт');
        spacing = Math.round((f.spacingType === 'condensed' ? -v : v) * 4) / 4;
      }
    }
    const ch = editor.chain().focus();
    const mark = (name, on) => (on ? ch.setMark(name) : ch.unsetMark(name));
    if (changed('fontFamily') && f.fontFamily.trim()) ch.setFontFamily(f.fontFamily.trim());
    if (size) ch.setFontSize(size);
    if (changed('style') && f.style) { mark('bold', /bold/i.test(f.style)); mark('italic', /italic/i.test(f.style)); }
    if (changed('color')) (f.color ? ch.setColor(f.color) : ch.unsetColor());
    if ((changed('underline') || changed('underlineColor')) && f.underline !== '') {
      if (f.underline === 'none') ch.unsetMark('underline');
      else ch.setUnderlineKind(f.underline, f.underlineColor || null);
    }
    if (changed('strike') && f.strike !== '') mark('strike', f.strike);
    if (changed('superscript') && f.superscript !== '') { if (f.superscript) ch.unsetMark('subscript'); mark('superscript', f.superscript); }
    if (changed('subscript') && f.subscript !== '') { if (f.subscript) ch.unsetMark('superscript'); mark('subscript', f.subscript); }
    const run = {};
    if (changed('caps') && f.caps !== '') run.caps = f.caps || null;
    if (changed('smallCaps') && f.smallCaps !== '') run.smallCaps = f.smallCaps || null;
    if (spacing !== undefined) run.spacing = spacing;
    if (Object.keys(run).length) ch.setRunAttrs(run);
    ch.run();
    onClose();
    return true;
  };

  // взаимоисключающие пары — как в Word
  const check = (k, other) => (e) => setF((s) => ({ ...s, [k]: e.target.checked, ...(e.target.checked && other ? { [other]: false } : {}) }));
  const Check = ({ k, other, label }) => (
    <label className="pd-check">
      <input type="checkbox" checked={f[k] === true} ref={(el) => { if (el) el.indeterminate = f[k] === ''; }} onChange={check(k, other)} /> {label}
    </label>
  );

  const sp = num(f.spacingBy) || 0;
  const previewSize = parseSize(f.fontSize) || 11;
  const preview = {
    fontFamily: f.fontFamily ? cssFontStack(f.fontFamily) : undefined,
    fontSize: `${Math.min(previewSize, 36)}pt`,
    fontWeight: /bold/i.test(f.style) ? 'bold' : undefined,
    fontStyle: /italic/i.test(f.style) ? 'italic' : undefined,
    color: f.color || undefined,
    textDecorationLine: [f.underline && f.underline !== 'none' && 'underline', f.strike === true && 'line-through'].filter(Boolean).join(' ') || undefined,
    textDecorationStyle: { double: 'double', wave: 'wavy', wavyDouble: 'wavy', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashLong: 'dashed', dotDash: 'dashed', dotDotDash: 'dashed' }[f.underline],
    textDecorationColor: f.underlineColor || undefined,
    textTransform: f.caps === true ? 'uppercase' : undefined,
    fontVariant: f.smallCaps === true ? 'small-caps' : undefined,
    letterSpacing: f.spacingType === 'expanded' ? `${sp}pt` : f.spacingType === 'condensed' ? `${-sp}pt` : undefined,
    verticalAlign: f.superscript === true ? 'super' : f.subscript === true ? 'sub' : undefined,
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd font-dlg" role="dialog" aria-label="Шрифт" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Шрифт</div>
        <div className="pd-tabs" role="tablist">
          <button type="button" role="tab" className={tab === 'font' ? 'active' : ''} onClick={() => setTab('font')}>Шрифт</button>
          <button type="button" role="tab" className={tab === 'adv' ? 'active' : ''} onClick={() => setTab('adv')}>Дополнительно</button>
        </div>
        <div className="modal-body">
          {tab === 'font' && (<>
            <div className="fd-top">
              <label className="fd-col wide"><span>Шрифт</span>
                <input list="fd-fonts" value={f.fontFamily} onChange={(e) => set('fontFamily')(e.target.value)} autoFocus />
                <datalist id="fd-fonts">{fontList.map((x) => <option key={x} value={x} />)}</datalist>
              </label>
              <label className="fd-col"><span>Начертание</span>
                <select value={f.style} onChange={(e) => set('style')(e.target.value)}>
                  {f.style === '' && <option value="">(разное)</option>}
                  {STYLES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                </select>
              </label>
              <label className="fd-col small"><span>Размер</span>
                <input list="fd-sizes" value={f.fontSize} onChange={(e) => set('fontSize')(e.target.value)} />
                <datalist id="fd-sizes">{FONT_SIZES.map((x) => <option key={x} value={String(x).replace('.', ',')} />)}</datalist>
              </label>
            </div>
            <div className="fd-top">
              <label className="fd-col"><span>Цвет текста</span>
                <span className="fd-color">
                  <input type="color" value={f.color || '#000000'} onChange={(e) => set('color')(e.target.value)} />
                  <button type="button" className={`btn small${f.color ? '' : ' on'}`} onClick={() => set('color')('')}>Авто</button>
                </span>
              </label>
              <label className="fd-col"><span>Подчёркивание</span>
                <select value={f.underline} onChange={(e) => set('underline')(e.target.value)}>
                  {f.underline === '' && <option value="">(разное)</option>}
                  <option value="none">(нет)</option>
                  {UNDERLINES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                </select>
              </label>
              <label className="fd-col"><span>Цвет подчёркивания</span>
                <span className="fd-color">
                  <input type="color" disabled={!f.underline || f.underline === 'none'} value={f.underlineColor || '#000000'} onChange={(e) => set('underlineColor')(e.target.value)} />
                  <button type="button" className={`btn small${f.underlineColor ? '' : ' on'}`} disabled={!f.underline || f.underline === 'none'} onClick={() => set('underlineColor')('')}>Авто</button>
                </span>
              </label>
            </div>
            <fieldset>
              <legend>Видоизменение</legend>
              <div className="fd-checks">
                <Check k="strike" label="Зачёркнутый" />
                <Check k="superscript" other="subscript" label="Надстрочный" />
                <Check k="subscript" other="superscript" label="Подстрочный" />
                <Check k="smallCaps" other="caps" label="Малые прописные" />
                <Check k="caps" other="smallCaps" label="Все прописные" />
              </div>
            </fieldset>
          </>)}
          {tab === 'adv' && (
            <fieldset>
              <legend>Межзнаковый интервал</legend>
              <div className="pd-grid">
                <label className="pd-field"><span>Интервал</span>
                  <select value={f.spacingType} onChange={(e) => setF((s) => ({ ...s, spacingType: e.target.value, spacingBy: e.target.value === 'normal' ? '' : s.spacingBy || '1' }))}>
                    {f.spacingType === '' && <option value="">(разный)</option>}
                    <option value="normal">Обычный</option>
                    <option value="expanded">Разреженный</option>
                    <option value="condensed">Уплотнённый</option>
                  </select>
                </label>
                <label className="pd-field"><span>на</span>
                  <span className="pd-input"><input value={f.spacingBy} disabled={!['expanded', 'condensed'].includes(f.spacingType)} inputMode="decimal"
                    onChange={(e) => set('spacingBy')(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); set('spacingBy')(fmtNum(Math.max(0.1, (num(f.spacingBy) || 0) + (e.key === 'ArrowUp' ? 0.1 : -0.1)), 2)); } }} /><em>пт</em></span>
                </label>
              </div>
            </fieldset>
          )}
          <fieldset>
            <legend>Образец</legend>
            <div className="fd-preview"><span style={preview}>{f.fontFamily || 'Образец текста'}</span></div>
          </fieldset>
          {error && <div className="pd-error" role="alert">{error}</div>}
        </div>
        <div className="modal-actions">
          <span className="spacer" />
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
