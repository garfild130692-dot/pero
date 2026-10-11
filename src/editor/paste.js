// Параметры вставки как в Word: «Сохранить исходное форматирование» (обычная вставка), «Объединить форматирование»
// (остаются полужирный, курсив, подчёркивание, ссылки, списки, таблицы и заголовки; шрифты, размеры, цвета, интервалы
// и отступы — как в документе), «Сохранить только текст» (текст с оформлением места вставки).
// После вставки рядом с ней — кнопка «Параметры вставки»: выбор другого варианта переделывает уже вставленное.
import { Extension } from '@tiptap/core';
import { Fragment, Slice } from '@tiptap/pm/model';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';

export const pasteKey = new PluginKey('pasteOptions');
export const PASTE_MODES = [
  ['keep', 'Сохранить исходное форматирование', 'И'],
  ['merge', 'Объединить форматирование', 'О'],
  ['text', 'Сохранить только текст', 'Т'],
];
const MERGE_MARKS = new Set(['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'link']);
const KEEP_ATTRS = { heading: ['level'], orderedList: ['start'], image: null, opaque: null, tab: null, tableCell: ['colspan', 'rowspan', 'colwidth'], tableHeader: ['colspan', 'rowspan', 'colwidth'] };

let pendingMode = null; // режим вставки, которую сейчас выполняет pasteWithMode
let lastData = null;    // что было в буфере при последней вставке (для «Параметров вставки»)

// Фрагмент без «чужого» оформления: метки — только смысловые, атрибуты блоков — по умолчанию (уровень заголовка,
// начало нумерации, размеры ячеек и данные рисунков/значков — сохраняются)
function mergeFragment(frag, schema) {
  const out = [];
  frag.forEach((node) => {
    if (node.isText) { out.push(node.mark(node.marks.filter((m) => MERGE_MARKS.has(m.type.name)))); return; }
    const keep = KEEP_ATTRS[node.type.name];
    let attrs = keep === null ? node.attrs : {};
    if (keep) for (const k of keep) attrs[k] = node.attrs[k];
    const marks = node.marks.filter((m) => MERGE_MARKS.has(m.type.name));
    out.push(node.type.create(attrs, mergeFragment(node.content, schema), marks));
  });
  return Fragment.from(out);
}

// Только текст: абзацы из текста с оформлением места вставки
function textSlice(text, state) {
  const marks = state.storedMarks || state.selection.$from.marks();
  const { schema } = state;
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
  const paras = lines.map((l) => schema.nodes.paragraph.create(null, l ? schema.text(l, marks) : null));
  return new Slice(Fragment.from(paras), 1, 1);
}

export const PasteOptions = Extension.create({
  name: 'pasteOptions',
  addCommands() {
    return {
      dismissPasteOptions: () => ({ tr, dispatch }) => { if (dispatch) dispatch(tr.setMeta(pasteKey, { clear: true })); return true; },
    };
  },
  addProseMirrorPlugins() {
    return [new Plugin({
      key: pasteKey,
      state: {
        init: () => null,
        apply(tr, prev, oldState) {
          const meta = tr.getMeta(pasteKey);
          if (meta) return meta.clear ? null : meta;
          if (tr.getMeta('uiEvent') === 'paste' && lastData) {
            const from = tr.mapping.map(oldState.selection.from, -1);
            return { from, to: tr.selection.to, data: lastData, mode: tr.getMeta('pasteMode') || 'keep' };
          }
          if (!prev) return null;
          if (!tr.docChanged) return prev;
          // догоняющие правки других плагинов (табуляции, исправления) — только сдвиг; правка пользователя — кнопка
          // исчезает (как в Word, когда начинаешь печатать)
          if (tr.getMeta('appendedTransaction')) return { ...prev, from: tr.mapping.map(prev.from, -1), to: tr.mapping.map(prev.to) };
          return null;
        },
      },
      props: {
        handlePaste(view, event) {
          const cd = event && event.clipboardData;
          lastData = cd ? { html: cd.getData('text/html') || '', text: cd.getData('text/plain') || '' } : lastData;
          return false; // вставку делает ProseMirror (картинки — App)
        },
        transformPasted(slice, view) {
          const mode = pendingMode;
          if (mode === 'merge') return new Slice(mergeFragment(slice.content, view.state.schema), slice.openStart, slice.openEnd);
          if (mode === 'text') return textSlice(slice.content.textBetween(0, slice.content.size, '\n', '\n'), view.state);
          return slice;
        },
      },
      appendTransaction(trs, oldState, newState) {
        // запомнить режим только что выполненной вставки (для отметки в меню)
        const t = trs.find((x) => x.getMeta('uiEvent') === 'paste');
        if (!t || !pendingMode) return null;
        const p = pasteKey.getState(newState);
        if (!p) return null;
        return newState.tr.setMeta(pasteKey, { ...p, mode: pendingMode }).setMeta('addToHistory', false);
      },
    })];
  },
});

// Вставить данные буфера ({ html, text }) в выбранном режиме вместо выделения. Не команда TipTap: вставку
// диспетчеризует сам ProseMirror (pasteHTML/pasteText), команда поверх неё дала бы «несовпадающую транзакцию».
export function pasteWithMode(editor, data, mode) {
  if (!data) return false;
  const { view } = editor;
  lastData = data;
  pendingMode = mode;
  try {
    if (mode === 'text' || !data.html) view.pasteText(data.text || '');
    else view.pasteHTML(data.html);
  } finally { pendingMode = null; }
  return true;
}

// Переделать последнюю вставку в другом режиме (кнопка «Параметры вставки»)
export function repaste(editor, mode) {
  const { state, view } = editor;
  const p = pasteKey.getState(state);
  if (!p || !p.data) return false;
  view.dispatch(state.tr.setSelection(TextSelection.create(state.doc, p.from, p.to)));
  view.focus();
  return pasteWithMode(editor, p.data, mode);
}

// Для тестов: преобразование фрагмента без ProseMirror-вида
export const _test = { mergeFragment, textSlice };
