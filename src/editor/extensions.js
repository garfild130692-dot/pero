// Набор расширений редактора — единая схема документа для UI и для тестов.
import { Extension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import Document from '@tiptap/extension-document';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyle } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import Underline from '@tiptap/extension-underline';
import {
  DEFAULT_FONT, DEFAULT_SIZE, HEADING_FONT, HEADING_SIZES, cssFontStack, normalizeColor, normalizeFontFamily, parseSize, stepSize, MIN_SIZE, MAX_SIZE,
} from './font.js';
import { NEW_DOC_PARA } from './paragraph.js';
import { ParaFormat } from './paraFormat.js';
import { ParaLayout } from './paraLayout.js';
import { DEFAULT_TAB, TabNode, Tabs } from './tabs.js';
import { PeroTable, PeroTableCell, PeroTableHeader, PeroTableRow, TableFormat } from './tableFormat.js';
import { Opaque } from './opaque.js';
import { PeroImage } from './image.js';
import { EMPTY_HF } from './page.js';
import { Pagination } from './pagination.js';
import { Sections } from './sections.js';
import { SearchReplace } from './search.js';
import { ParagraphStyles, styleOf } from './styles.js';
import { Comments } from './comments.js';
import { Deletion, Insertion, TrackChanges } from './track.js';
import { RunFormat } from './runFormat.js';
import { FormattingMarks } from './marks.js';
import { Fields } from './fields.js';
import { Footnotes } from './footnotes.js';
import { Numbering } from './numbering.js';
import { Design } from './design.js';
import { PasteOptions } from './paste.js';
import { CharStyle } from './builtinStyles.js';

// Как оформлены заголовки — font.js (совпадает со styles.css и стилями Heading1–6 в export.js)
export { HEADING_FONT, HEADING_SIZES };

// Умолчания шрифта и абзаца храним в атрибутах документа (аналог docDefaults + стиль «Обычный»).
// В текст и абзацы попадает только прямое форматирование — то, что отличается от умолчаний.
const PeroDocument = Document.extend({
  addAttributes() {
    return {
      fontFamily: { default: DEFAULT_FONT },
      fontSize: { default: DEFAULT_SIZE },
      para: { default: NEW_DOC_PARA }, // { spaceBefore, spaceAfter, lineRule, lineSpacing, indentLeft, indentRight, indentFirst }
      sectPr: { default: null },       // исходный XML параметров последнего раздела (страница, поля, колонтитулы)
      hf: { default: EMPTY_HF },       // колонтитулы: { header, footer, own } (см. page.js)
      styles: { default: [] },         // пользовательские стили абзацев (см. styles.js)
      defaultsModified: { default: false }, // «Обычный» изменён в Пере — при сохранении переписать docDefaults
      comments: { default: [] },       // примечания: [{ id, author, initials, date, text }] (см. comments.js)
      commentsEdited: { default: false }, // примечания менялись — при сохранении comments.xml пишется заново
      footnotes: { default: [] },      // тексты сносок: [{ id, text, changed }] (см. footnotes.js)
      headingNumbering: { default: null }, // нумерация заголовков (см. numbering.js)
      charStyles: { default: [] },     // символьные стили: [{ id, name, run, origin }] (см. builtinStyles.js)
      watermark: { default: null },     // подложка (см. design.js)
      watermarkChanged: { default: false },
      pageColor: { default: null },     // цвет страницы '#rrggbb'
      pageColorChanged: { default: false },
      footnotesEdited: { default: false }, // сноски менялись — при сохранении footnotes.xml пишется заново
      defaultTab: { default: DEFAULT_TAB }, // шаг позиций табуляции по умолчанию, pt (settings.xml w:defaultTabStop)
    };
  },
});

// Загрузить документ целиком. setContent меняет только содержимое, атрибуты узла doc (умолчания шрифта)
// остаются от прошлого документа — выставляем их отдельно; без записи в историю и без события «изменён».
export function loadDocument(editor, json) {
  editor.commands.setContent(json, { emitUpdate: false });
  const attrs = {
    fontFamily: DEFAULT_FONT, fontSize: DEFAULT_SIZE, para: NEW_DOC_PARA, sectPr: null, hf: EMPTY_HF, styles: [], defaultsModified: false,
    comments: [], commentsEdited: false, footnotes: [], footnotesEdited: false, headingNumbering: null, charStyles: [], watermark: null, watermarkChanged: false, pageColor: null, pageColorChanged: false, defaultTab: DEFAULT_TAB,
    ...((json && json.attrs) || {}),
  };
  const { tr } = editor.state;
  for (const [k, v] of Object.entries(attrs)) if (editor.state.doc.attrs[k] !== v) tr.setDocAttribute(k, v);
  if (tr.docChanged) editor.view.dispatch(tr.setMeta('addToHistory', false).setMeta('preventUpdate', true));
}

// Действующие шрифт и размер текста с учётом прямого форматирования, заголовка и умолчаний документа
export function effectiveFont(textStyleAttrs, parent, docAttrs) {
  const a = textStyleAttrs || {};
  const isHeading = parent && parent.type.name === 'heading';
  const st = !isHeading && parent ? (styleOf(docAttrs, parent.attrs.styleId) || {}).run || {} : {}; // стиль абзаца
  return {
    fontFamily: a.fontFamily || st.fontFamily || (isHeading ? HEADING_FONT : docAttrs.fontFamily || DEFAULT_FONT),
    fontSize: a.fontSize || st.fontSize || (isHeading ? HEADING_SIZES[parent.attrs.level] || DEFAULT_SIZE : docAttrs.fontSize || DEFAULT_SIZE),
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
        // межзнаковый интервал, pt (w:spacing): > 0 — разреженный, < 0 — уплотнённый
        spacing: {
          default: null,
          parseHTML: (el) => { const v = parseFloat(el.style.letterSpacing); return v ? Math.round(v * 4) / 4 : null; },
          renderHTML: (a) => (a.spacing ? { style: `letter-spacing: ${a.spacing}pt` } : {}),
        },
        // все прописные (w:caps) и малые прописные (w:smallCaps) — вид, сам текст не меняется
        caps: {
          default: null,
          parseHTML: (el) => (el.style.textTransform === 'uppercase' ? true : null),
          renderHTML: (a) => (a.caps ? { style: 'text-transform: uppercase' } : {}),
        },
        smallCaps: {
          default: null,
          parseHTML: (el) => (el.style.fontVariant === 'small-caps' || el.style.fontVariantCaps === 'small-caps' ? true : null),
          renderHTML: (a) => (a.smallCaps ? { style: 'font-variant: small-caps' } : {}),
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
      setRunAttrs: (attrs) => set(attrs), // { spacing, caps, smallCaps, … } — диалог «Шрифт»
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

// Подчёркивание с видом и цветом (как w:u w:val / w:color): single, double, thick, dotted, dottedHeavy, dash,
// dashLong, dotDash, dotDotDash, wave, wavyDouble, words; клавиши как в Word: Ctrl+U, Ctrl+Shift+D (двойное),
// Ctrl+Shift+W (только слова)
export const UNDERLINES = [
  ['single', 'одинарное'], ['words', 'только слова'], ['double', 'двойное'], ['thick', 'толстое'], ['dotted', 'точками'],
  ['dottedHeavy', 'толстое точками'], ['dash', 'штрихами'], ['dashLong', 'длинными штрихами'], ['dotDash', 'штрих-пунктир'],
  ['dotDotDash', 'штрих-пунктир с двумя точками'], ['wave', 'волнистое'], ['wavyDouble', 'двойное волнистое'],
];
const UNDERLINE_CSS = {
  single: 'solid', words: 'solid', double: 'double', thick: 'solid', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashLong: 'dashed',
  dotDash: 'dashed', dotDotDash: 'dashed', wave: 'wavy', wavyDouble: 'wavy',
};
const PeroUnderline = Underline.extend({
  addAttributes() {
    return {
      kind: { default: null, parseHTML: (el) => el.getAttribute('data-u') || null, renderHTML: () => ({}) },
      color: { default: null, parseHTML: (el) => normalizeColor(el.style.textDecorationColor), renderHTML: () => ({}) },
    };
  },
  renderHTML({ mark, HTMLAttributes }) {
    const k = mark.attrs.kind || 'single';
    const style = [`text-decoration-line: underline`, `text-decoration-style: ${UNDERLINE_CSS[k] || 'solid'}`,
      /Heavy|thick/.test(k) ? 'text-decoration-thickness: 2px' : '', mark.attrs.color ? `text-decoration-color: ${mark.attrs.color}` : '',
      k === 'words' ? 'text-decoration-skip-ink: auto' : ''].filter(Boolean).join('; ');
    return ['u', { ...HTMLAttributes, ...(k !== 'single' ? { 'data-u': k } : {}), style }, 0];
  },
  addCommands() {
    return {
      ...this.parent?.(),
      // вид подчёркивания у выделения (null — убрать)
      setUnderlineKind: (kind, color = null) => ({ commands }) => (kind ? commands.setMark('underline', { kind: kind === 'single' ? null : kind, color }) : commands.unsetMark('underline')),
    };
  },
  addKeyboardShortcuts() {
    const t = (kind) => () => (this.editor.isActive('underline', { kind }) ? this.editor.commands.unsetMark('underline') : this.editor.commands.setUnderlineKind(kind));
    return { 'Mod-u': () => this.editor.commands.toggleUnderline(), 'Mod-Shift-d': t('double'), 'Mod-Shift-w': t('words') };
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
    underline: false, // своё: с видом и цветом
  }),
  PeroUnderline,
  PeroDocument,
  TextAlign.configure({ types: ['heading', 'paragraph'] }),
  ParaFormat,
  ParaLayout,
  TabNode,
  Tabs,
  ParagraphStyles,
  Comments,
  Insertion,
  Deletion,
  TrackChanges,
  RunFormat,
  FormattingMarks,
  Fields,
  Footnotes,
  Numbering,
  Design,
  PasteOptions,
  CharStyle,
  PeroTable.configure({ resizable: true }),
  PeroTableRow,
  PeroTableCell,
  PeroTableHeader,
  TableFormat,
  TextStyle,
  RunStyle,
  Opaque,
  PeroImage,
  Sections,
  Pagination,
  SearchReplace,
  Highlight.configure({ multicolor: true }),
  Sup,
  Sub,
];
