// Встроенные стили Word, которых нет среди «Обычный / Заголовок 1–6»: абзацные («Без интервала», «Название»,
// «Подзаголовок», «Цитата», «Выделенная цитата») и символьные («Выделение», «Строгий», «Сильное выделение»…).
// Определение попадает в документ (doc.attrs.styles / doc.attrs.charStyles, origin 'pero', builtin) при первом
// применении; при сохранении в исходный пакет стиль с тем же внутренним именем Word берётся из файла (свой id не пишется).
//
// Символьный стиль — метка charStyle { id } (w:rStyle); вид — CSS по data-cstyle (charStyleCss).
import { Mark, mergeAttributes } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { cssFontStack } from './font.js';
import { wordRange } from './runFormat.js';

export const BUILTIN_PARA = [
  { id: 'NoSpacing', name: 'No Spacing', label: 'Без интервала', para: { spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1 }, run: {} },
  { id: 'Title', name: 'Title', label: 'Название', para: { spaceAfter: 0, lineRule: 'auto', lineSpacing: 1 }, run: { fontFamily: 'Calibri Light', fontSize: 28 } },
  { id: 'Subtitle', name: 'Subtitle', label: 'Подзаголовок', para: {}, run: { color: '#5a5a5a' } },
  { id: 'Quote', name: 'Quote', label: 'Цитата', para: { spaceBefore: 10, spaceAfter: 8, indentLeft: 43.2, indentRight: 43.2 }, align: 'center', run: { italic: true, color: '#404040' } },
  { id: 'IntenseQuote', name: 'Intense Quote', label: 'Выделенная цитата', para: { spaceBefore: 18, spaceAfter: 18, indentLeft: 43.2, indentRight: 43.2 }, align: 'center', run: { italic: true, color: '#4472c4' } },
];
export const BUILTIN_CHAR = [
  { id: 'Emphasis', name: 'Emphasis', label: 'Выделение', run: { italic: true } },
  { id: 'Strong', name: 'Strong', label: 'Строгий', run: { bold: true } },
  { id: 'IntenseEmphasis', name: 'Intense Emphasis', label: 'Сильное выделение', run: { italic: true, color: '#4472c4' } },
  { id: 'SubtleEmphasis', name: 'Subtle Emphasis', label: 'Слабое выделение', run: { italic: true, color: '#404040' } },
  { id: 'BookTitle', name: 'Book Title', label: 'Название книги', run: { bold: true, italic: true } },
];

// Русская подпись встроенного стиля по внутреннему имени Word (имена стилей файла показываются как есть)
const LABELS = Object.fromEntries([...BUILTIN_PARA, ...BUILTIN_CHAR].map((s) => [s.name.toLowerCase(), s.label]));
export const styleLabel = (name) => LABELS[String(name || '').toLowerCase()] || name;
export const isBuiltinName = (name) => !!LABELS[String(name || '').toLowerCase()];

// Определение стиля в документе по внутреннему имени (из файла — с его id) или новое встроенное
export function ensureStyle(list, def) {
  const found = (list || []).find((s) => String(s.name).toLowerCase() === def.name.toLowerCase() || s.id === def.id);
  if (found) return { list, id: found.id };
  const { label, ...rest } = def;
  return { list: [...(list || []), { ...rest, origin: 'pero', builtin: true, modified: true }], id: def.id };
}

// CSS символьных стилей
export function charStyleCss(list) {
  return (list || []).map((s) => {
    const r = s.run || {};
    const d = [];
    if (r.fontFamily) d.push(`font-family: ${cssFontStack(r.fontFamily)}`);
    if (r.fontSize) d.push(`font-size: ${r.fontSize}pt`);
    if (r.color) d.push(`color: ${r.color}`);
    if (r.bold) d.push('font-weight: bold');
    if (r.italic) d.push('font-style: italic');
    if (r.underline) d.push('text-decoration: underline');
    if (r.strike) d.push('text-decoration: line-through');
    return d.length ? `.page-content span[data-cstyle="${String(s.id).replace(/"/g, '')}"] { ${d.join('; ')} }` : '';
  }).join('\n');
}

export const CharStyle = Mark.create({
  name: 'charStyle',
  excludes: 'charStyle', // у текста один символьный стиль (как w:rStyle)
  addAttributes() {
    return { id: { default: null, parseHTML: (el) => el.getAttribute('data-cstyle'), renderHTML: (a) => ({ 'data-cstyle': a.id }) } };
  },
  parseHTML() { return [{ tag: 'span[data-cstyle]' }]; },
  renderHTML({ HTMLAttributes }) { return ['span', mergeAttributes(HTMLAttributes), 0]; },
  addCommands() {
    return {
      // применить встроенный абзацный стиль (определение добавляется в документ при первом применении)
      applyBuiltinParaStyle: (def) => ({ state, chain }) => {
        const { list, id } = ensureStyle(state.doc.attrs.styles, def);
        return chain().command(({ tr }) => { if (list !== state.doc.attrs.styles) tr.setDocAttribute('styles', list); return true; }).setParagraphStyle(id).run();
      },
      // символьный стиль: на выделение (без выделения — слово под курсором); уже весь в этом стиле — снять
      applyCharStyle: (def) => ({ state, tr, dispatch }) => {
        const { list, id } = ensureStyle(state.doc.attrs.charStyles, def);
        let { from, to } = state.selection;
        if (from === to) { const w = wordRange(state); if (w) ({ from, to } = w); }
        const type = state.schema.marks.charStyle;
        if (from === to) { if (dispatch) dispatch(tr.addStoredMark(type.create({ id })).setDocAttribute('charStyles', list)); return true; }
        let all = true;
        state.doc.nodesBetween(from, to, (n) => { if (n.isText && !n.marks.some((m) => m.type === type && m.attrs.id === id)) all = false; });
        if (all) tr.removeMark(from, to, type); else tr.addMark(from, to, type.create({ id }));
        if (list !== state.doc.attrs.charStyles) tr.setDocAttribute('charStyles', list);
        if (dispatch) dispatch(tr.setSelection(TextSelection.create(tr.doc, state.selection.from, state.selection.to)));
        return true;
      },
    };
  },
});
