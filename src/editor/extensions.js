// Набор расширений редактора — единая схема документа для UI и для тестов.
import { Extension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Document from '@tiptap/extension-document';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import {
  DEFAULT_FONT, DEFAULT_SIZE, cssFontStack, normalizeColor, normalizeFontFamily, parseSize, stepSize, MIN_SIZE, MAX_SIZE,
} from './font.js';
import { NEW_DOC_PARA } from './paragraph.js';
import { ParaFormat } from './paraFormat.js';
import { PeroTable, PeroTableCell, PeroTableHeader, PeroTableRow, TableFormat } from './tableFormat.js';
import { Opaque } from './opaque.js';
import { PeroImage } from './image.js';

// Как оформлены заголовки (совпадает со styles.css и стилями Heading1–6 в export.js)
export const HEADING_FONT = 'Calibri Light';
export const HEADING_SIZES = { 1: 16, 2: 13, 3: 12, 4: 11, 5: 11, 6: 11 };

// Умолчания шрифта и абзаца храним в атрибутах документа (аналог docDefaults + стиль «Обычный»).
// В текст и абзацы попадает только прямое форматирование — то, что отличается от умолчаний.
const PeroDocument = Document.extend({
  addAttributes() {
    return {
      fontFamily: { default: DEFAULT_FONT },
      fontSize: { default: DEFAULT_SIZE },
      para: { default: NEW_DOC_PARA }, // { spaceBefore, spaceAfter, lineRule, lineSpacing, indentLeft, indentRight, indentFirst }
      sectPr: { default: null },       // исходный XML параметров последнего раздела (страница, поля, колонтитулы)
    };
  },
});

// Загрузить документ целиком. setContent меняет только содержимое, атрибуты узла doc (умолчания шрифта)
// остаются от прошлого документа — выставляем их отдельно; без записи в историю и без события «изменён».
export function loadDocument(editor, json) {
  editor.commands.setContent(json, { emitUpdate: false });
  const attrs = { fontFamily: DEFAULT_FONT, fontSize: DEFAULT_SIZE, para: NEW_DOC_PARA, sectPr: null, ...((json && json.attrs) || {}) };
  const { tr } = editor.state;
  for (const [k, v] of Object.entries(attrs)) if (editor.state.doc.attrs[k] !== v) tr.setDocAttribute(k, v);
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false).setMeta('preventUpdate', true));
}

// Действующие шрифт и размер текста с учётом прямого форматирования, заголовка и умолчаний документа
export function effectiveFont(textStyleAttrs, parent, docAttrs) {
  const a = textStyleAttrs || {};
  const isHeading = parent && parent.type.name === 'heading';
  return {
    fontFamily: a.fontFamily || (isHeading ? HEADING_FONT : docAttrs.fontFamily || DEFAULT_FONT),
    fontSize: a.fontSize || (isHeading ? HEADING_SIZES[parent.attrs.level] || DEFAULT_SIZE : docAttrs.fontSize || DEFAULT_SIZE),
  };
}

// Шрифт, размер и цвет — атрибуты метки textStyle (в DOM — <span style>).
// Свои атрибуты вместо FontFamily/FontSize/Color из TipTap: значения нормализуем сразу при вставке из буфера.
const RunStyle = Extension.create({
  name: 'runStyle',

  addGlobalAttributes() {
    return [{
      types: ['textStyle'],
      attributes: {
        fontFamily: {
          default: null,
          parseHTML: (el) => normalizeFontFamily(el.style.fontFamily),
          renderHTML: (a) => (a.fontFamily ? { style: `font-family: ${cssFontStack(a.fontFamily)}` } : {}),
        },
        fontSize: {
          default: null,
          parseHTML: (el) => parseSize(el.style.fontSize),
          renderHTML: (a) => (a.fontSize ? { style: `font-size: ${a.fontSize}pt` } : {}),
        },
        color: {
          default: null,
          parseHTML: (el) => normalizeColor(el.style.color),
          renderHTML: (a) => (a.color ? { style: `color: ${a.color}` } : {}),
        },
      },
    }];
  },

  addCommands() {
    const set = (attrs) => ({ chain }) => chain().setMark('textStyle', attrs).removeEmptyTextStyle().run();
    return {
      setFontFamily: (name) => set({ fontFamily: normalizeFontFamily(name) }),
      setFontSize: (size) => {
        const n = parseSize(size);
        return n ? set({ fontSize: n }) : () => false;
      },
      setColor: (color) => set({ color: normalizeColor(color) }),
      unsetColor: () => set({ color: null }),

      // Увеличить/уменьшить размер — у каждого фрагмента выделения отдельно, как в Word.
      // mode 'scale' — по шкале размеров (Ctrl+Shift+>/<), 'point' — на 1 pt (Ctrl+]/[).
      changeFontSize: (dir, mode = 'scale') => ({ state, tr, dispatch }) => {
        const type = state.schema.marks.textStyle;
        const next = (size) => (mode === 'point'
          ? Math.min(Math.max(size + dir, MIN_SIZE), MAX_SIZE)
          : stepSize(size, dir));
        const apply = (marks, parent) => {
          const m = type.isInSet(marks);
          const attrs = m ? m.attrs : {};
          const cur = effectiveFont(attrs, parent, state.doc.attrs).fontSize;
          return type.create({ ...attrs, fontSize: next(cur) });
        };
        const { selection } = tr;
        if (selection.empty) {
          const marks = state.storedMarks || selection.$from.marks();
          if (dispatch) tr.addStoredMark(apply(marks, selection.$from.parent));
          return true;
        }
        if (dispatch) {
          for (const { $from, $to } of selection.ranges) {
            state.doc.nodesBetween($from.pos, $to.pos, (node, pos, parent) => {
              if (!node.isText) return;
              const from = Math.max(pos, $from.pos);
              const to = Math.min(pos + node.nodeSize, $to.pos);
              if (from < to) tr.addMark(from, to, apply(node.marks, parent));
            });
          }
        }
        return true;
      },

      // Очистить форматирование символов: снять все метки, кроме гиперссылок (Word их сохраняет)
      clearFormatting: () => ({ state, tr, dispatch }) => {
        const types = Object.values(state.schema.marks).filter((t) => t.name !== 'link');
        if (dispatch) {
          for (const { $from, $to } of tr.selection.ranges) types.forEach((t) => tr.removeMark($from.pos, $to.pos, t));
          tr.setStoredMarks((state.storedMarks || tr.selection.$from.marks()).filter((m) => m.type.name === 'link'));
        }
        return true;
      },
    };
  },

  addKeyboardShortcuts() {
    const c = () => this.editor.commands;
    return {
      'Mod-Shift-.': () => c().changeFontSize(1),
      'Mod-Shift-,': () => c().changeFontSize(-1),
      'Mod-]': () => c().changeFontSize(1, 'point'),
      'Mod-[': () => c().changeFontSize(-1, 'point'),
      'Mod-Space': () => c().clearFormatting(),
    };
  },
});

// Надстрочный и подстрочный взаимоисключающие (в OOXML это одно свойство w:vertAlign); клавиши — как в Word
const Sup = Superscript.extend({
  excludes: 'subscript',
  addKeyboardShortcuts() {
    // одна привязка: 'Mod-+' вдобавок к этой срабатывала бы на то же нажатие второй раз
    return { 'Mod-Shift-=': () => this.editor.commands.toggleSuperscript() };
  },
});
const Sub = Subscript.extend({
  excludes: 'superscript',
  addKeyboardShortcuts() {
    return { 'Mod-=': () => this.editor.commands.toggleSubscript() };
  },
});

export const extensions = [
  StarterKit.configure({
    document: false, // свой Document с атрибутами умолчаний
    // то, чего нет в Word как отдельных сущностей, отключаем
    code: false,
    codeBlock: false,
    blockquote: false,
    horizontalRule: false,
    heading: { levels: [1, 2, 3, 4, 5, 6] },
    link: { openOnClick: false, autolink: true, defaultProtocol: 'https' },
  }),
  PeroDocument,
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  ParaFormat,
  PeroTable.configure({ resizable: true }),
  PeroTableRow,
  PeroTableCell,
  PeroTableHeader,
  TableFormat,
  TextStyle,
  RunStyle,
  Opaque,
  PeroImage,
  Highlight.configure({ multicolor: true }),
  Sup,
  Sub,
];
