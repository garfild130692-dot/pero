// Пользовательские стили абзацев.
//
// Определения — в атрибуте документа styles: [{ id, name, para, run, align, origin, modified }]
//   para: отличия от «Обычного» (spaceBefore, spaceAfter, lineRule+lineSpacing, indentLeft/Right/First) в pt;
//   run: { fontFamily, fontSize, color, bold, italic, underline } — только заданное; align: left|center|right|justify.
//   origin: 'file' | 'pero'; modified — изменён в Пере (при сохранении в исходный пакет определение переписывается).
// Абзац ссылается на стиль атрибутом styleId; вид стиля задаёт CSS (styleCss), изменение стиля меняет все его абзацы.
import { Extension } from '@tiptap/core';
import { cssFontStack } from './font.js';
import { lineCss } from './paraFormat.js';

export const styleOf = (docAttrs, id) => (id && docAttrs && (docAttrs.styles || []).find((s) => s.id === id)) || null;

// CSS всех стилей документа — правила по data-pstyle
export function styleCss(styles) {
  return (styles || []).map((s) => {
    const r = s.run || {};
    const p = s.para || {};
    const decl = [];
    if (r.fontFamily) decl.push(`font-family: ${cssFontStack(r.fontFamily)}`);
    if (r.fontSize) decl.push(`font-size: ${r.fontSize}pt`);
    if (r.color) decl.push(`color: ${r.color}`);
    if (r.bold) decl.push('font-weight: bold');
    if (r.italic) decl.push('font-style: italic');
    if (r.underline) decl.push('text-decoration: underline');
    if (s.align) decl.push(`text-align: ${s.align}`);
    if (p.spaceBefore != null) decl.push(`margin-top: ${p.spaceBefore}pt`);
    if (p.spaceAfter != null) decl.push(`margin-bottom: ${p.spaceAfter}pt`);
    if (p.lineSpacing != null) decl.push(`line-height: ${lineCss(p.lineRule || 'auto', p.lineSpacing)}`);
    if (p.indentLeft != null) decl.push(`margin-left: ${p.indentLeft}pt`);
    if (p.indentRight != null) decl.push(`margin-right: ${p.indentRight}pt`);
    if (p.indentFirst != null) decl.push(`text-indent: ${p.indentFirst}pt`);
    const sel = `.page-content p[data-pstyle="${String(s.id).replace(/"/g, '')}"]`;
    return decl.length ? `${sel} { ${decl.join('; ')} }` : '';
  }).join('\n');
}

// Новый id стиля (латиница — так безопаснее для styles.xml)
export function newStyleId(styles) {
  let n = 1;
  const ids = new Set((styles || []).map((s) => s.id));
  while (ids.has(`PeroStyle${n}`)) n++;
  return `PeroStyle${n}`;
}

export const ParagraphStyles = Extension.create({
  name: 'paragraphStyles',

  addGlobalAttributes() {
    return [{
      types: ['paragraph'],
      attributes: {
        styleId: {
          default: null,
          keepOnSplit: true, // Enter — новый абзац того же стиля
          parseHTML: (el) => el.getAttribute('data-pstyle'),
          renderHTML: (a) => (a.styleId ? { 'data-pstyle': a.styleId } : {}),
        },
      },
    }];
  },

  addCommands() {
    const setDoc = (tr, name, value) => tr.setDocAttribute(name, value);
    return {
      // Применить стиль к абзацам выделения (null — «Обычный»); заголовки становятся абзацами
      setParagraphStyle: (id) => ({ state, tr, dispatch }) => {
        const { from, to } = state.selection;
        const para = state.schema.nodes.paragraph;
        state.doc.nodesBetween(from, to, (node, pos) => {
          if (node.type.name === 'heading') {
            const { level, ...rest } = node.attrs; // прямые параметры абзаца заголовка сохраняются
            tr.setNodeMarkup(pos, para, { ...rest, styleId: id });
            return false;
          }
          if (node.type.name === 'paragraph') { tr.setNodeMarkup(pos, undefined, { ...node.attrs, styleId: id }); return false; }
          return true;
        });
        if (dispatch) dispatch(tr);
        return true;
      },
      // Добавить/заменить определение стиля (по id)
      saveStyle: (style) => ({ state, tr, dispatch }) => {
        const list = [...(state.doc.attrs.styles || [])];
        const i = list.findIndex((s) => s.id === style.id);
        if (i >= 0) list[i] = { ...list[i], ...style, modified: true };
        else list.push({ origin: 'pero', modified: true, ...style });
        if (dispatch) dispatch(setDoc(tr, 'styles', list));
        return true;
      },
      // Удалить стиль: абзацы этого стиля становятся «Обычными» (их прямое форматирование остаётся)
      deleteStyle: (id) => ({ state, tr, dispatch }) => {
        state.doc.descendants((node, pos) => {
          if (node.type.name === 'paragraph' && node.attrs.styleId === id) tr.setNodeMarkup(pos, undefined, { ...node.attrs, styleId: null });
        });
        setDoc(tr, 'styles', (state.doc.attrs.styles || []).filter((s) => s.id !== id));
        if (dispatch) dispatch(tr);
        return true;
      },
      // Изменить «Обычный»: умолчания документа (шрифт, размер, абзац)
      setNormalStyle: ({ fontFamily, fontSize, para }) => ({ state, tr, dispatch }) => {
        if (fontFamily) setDoc(tr, 'fontFamily', fontFamily);
        if (fontSize) setDoc(tr, 'fontSize', fontSize);
        if (para) setDoc(tr, 'para', { ...state.doc.attrs.para, ...para });
        setDoc(tr, 'defaultsModified', true);
        if (dispatch) dispatch(tr);
        return true;
      },
    };
  },
});
