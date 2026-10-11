// Регистр (Shift+F3, меню «Регистр») и формат по образцу (кисть, Ctrl+Shift+C / Ctrl+Shift+V) — как в Word.
// Логика — чистые функции над состоянием ProseMirror (state → tr), их проверяют тесты без DOM.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';

// ---------- Регистр ----------
export const CASES = [
  ['sentence', 'Как в предложениях'], ['lower', 'все строчные'], ['upper', 'ВСЕ ПРОПИСНЫЕ'],
  ['title', 'Начинать С Прописных'], ['toggle', 'иЗМЕНИТЬ РЕГИСТР'],
];

const isLetter = (ch) => /\p{L}/u.test(ch);
const isWordChar = (ch) => /[\p{L}\p{N}'’]/u.test(ch);

// Преобразование строки. prev — символ перед строкой (для начала слова/предложения), sentenceStart — с начала абзаца.
export function transformCase(text, mode, prev = '', sentenceStart = true) {
  let out = '';
  let wordStart = !isWordChar(prev);
  let sent = sentenceStart;
  for (const ch of text) {
    let r = ch;
    if (mode === 'upper') r = ch.toUpperCase();
    else if (mode === 'lower') r = ch.toLowerCase();
    else if (mode === 'toggle') r = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase();
    else if (mode === 'title') r = wordStart ? ch.toUpperCase() : ch.toLowerCase();
    else if (mode === 'sentence') r = sent && isLetter(ch) ? ch.toUpperCase() : ch.toLowerCase();
    out += r;
    if (isLetter(ch) || /\p{N}/u.test(ch)) sent = false;
    if (/[.!?…]/.test(ch)) sent = true;
    wordStart = !isWordChar(ch);
  }
  return out;
}

// Shift+F3 как в Word: строчные → ВСЕ ПРОПИСНЫЕ → Начинать С Прописных → строчные; смешанный — ПРОПИСНЫЕ
export function nextCase(text) {
  const letters = [...text].filter(isLetter).join('');
  if (!letters) return 'upper';
  if (letters === letters.toUpperCase()) return transformCase(text, 'title') === text ? 'lower' : 'title';
  if (text === transformCase(text, 'title')) return 'lower';
  return 'upper';
}

// Слово под курсором (Word меняет регистр и кистью красит слово, если ничего не выделено)
export function wordRange(state) {
  const { $from } = state.selection;
  const parent = $from.parent;
  if (!parent.isTextblock) return null;
  const text = parent.textBetween(0, parent.content.size, '\0', '\0');
  let a = $from.parentOffset;
  let b = a;
  while (a > 0 && isWordChar(text[a - 1])) a--;
  while (b < text.length && isWordChar(text[b])) b++;
  if (a === b) return null;
  const start = $from.start();
  return { from: start + a, to: start + b };
}

// Сменить регистр в выделении (или слове под курсором). mode — из CASES или 'cycle' (Shift+F3)
export function changeCaseTr(state, mode) {
  let { from, to } = state.selection;
  if (from === to) {
    const w = wordRange(state);
    if (!w) return null;
    ({ from, to } = w);
  }
  if (mode === 'cycle') mode = nextCase(state.doc.textBetween(from, to, ' ', ' '));
  const tr = state.tr;
  const parts = []; // [pos, длина, новый текст, marks]
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return;
    const a = Math.max(from, pos);
    const b = Math.min(to, pos + node.nodeSize);
    if (a >= b) return;
    // контекст: текст абзаца перед фрагментом
    const $a = state.doc.resolve(a);
    const before = $a.parent.textBetween(0, $a.parentOffset, '\n', ' ');
    const prev = before.slice(-1);
    const sentenceStart = !/[\p{L}\p{N}]/u.test(before.replace(/^[\s\S]*[.!?…]/u, ''));
    const s = node.text.slice(a - pos, b - pos);
    const t = transformCase(s, mode, prev, sentenceStart);
    if (t !== s) parts.push([a, b, t, node.marks]);
  });
  if (!parts.length) return null;
  for (let i = parts.length - 1; i >= 0; i--) {
    const [a, b, t, marks] = parts[i];
    tr.replaceWith(a, b, state.schema.text(t, marks));
  }
  // выделение — тот же фрагмент (длина могла измениться: ß → SS)
  tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(from, -1), tr.mapping.map(to, 1)));
  return tr;
}

// ---------- Формат по образцу ----------
// Метки, которые переносит кисть (ссылки, примечания, исправления — не формат)
export const RUN_MARKS = ['bold', 'italic', 'underline', 'strike', 'textStyle', 'highlight', 'superscript', 'subscript'];
// Атрибуты абзаца, которые кисть не переносит (раздел — не формат абзаца)
const KEEP_ATTRS = ['sectPr'];

// Запомнить формат: метки первого символа выделения (или у курсора) и абзац. Абзацный формат переносится, если
// выделения нет или выделен абзац целиком (как в Word — когда в образец попал знак абзаца).
export function captureFormat(state) {
  const { selection, doc } = state;
  let marks = null;
  if (!selection.empty) {
    doc.nodesBetween(selection.from, selection.to, (node) => {
      if (marks || !node.isText) return !marks;
      marks = node.marks;
      return false;
    });
  }
  if (!marks) marks = state.storedMarks || selection.$from.marks();
  const run = marks.filter((m) => RUN_MARKS.includes(m.type.name)).map((m) => m.toJSON());
  const $f = selection.$from;
  const block = $f.parent.isTextblock ? $f.parent : null;
  const whole = selection.empty || ($f.parentOffset === 0 && selection.$to.parentOffset === selection.$to.parent.content.size)
    || $f.parent !== selection.$to.parent;
  return { run, para: block && whole ? { type: block.type.name, attrs: { ...block.attrs } } : null };
}

// Применить запомненный формат к выделению (пустое — слово под курсором)
export function applyFormatTr(state, fmt) {
  if (!fmt) return null;
  let { from, to } = state.selection;
  const tr = state.tr;
  const { schema } = state;
  if (from === to) {
    const w = wordRange(state);
    if (w) ({ from, to } = w);
  }
  if (from < to) {
    for (const name of RUN_MARKS) if (schema.marks[name]) tr.removeMark(from, to, schema.marks[name]);
    for (const m of fmt.run) {
      const type = schema.marks[m.type];
      if (type) tr.addMark(from, to, type.create(m.attrs));
    }
  } else {
    tr.setStoredMarks(fmt.run.map((m) => schema.markFromJSON(m)));
  }
  if (fmt.para) {
    const type = schema.nodes[fmt.para.type];
    const sel = state.selection;
    state.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
      if (!node.isTextblock) return true;
      if (!['paragraph', 'heading'].includes(node.type.name) || !type) return false;
      const attrs = { ...fmt.para.attrs };
      for (const k of KEEP_ATTRS) if (k in node.attrs) attrs[k] = node.attrs[k];
      tr.setNodeMarkup(pos, type, attrs);
      return false;
    });
  }
  return tr.docChanged || tr.storedMarksSet ? tr : null;
}

// Кисть: { fmt, mode: 0 — выключена, 1 — один раз, 2 — пока не нажата Esc или кнопка (двойной щелчок) }
export const painterKey = new PluginKey('formatPainter');

export const RunFormat = Extension.create({
  name: 'runFormat',
  addStorage() {
    return { copied: null }; // Ctrl+Shift+C → Ctrl+Shift+V
  },
  addCommands() {
    return {
      changeCase: (mode) => ({ state, dispatch }) => {
        const tr = changeCaseTr(state, mode);
        if (!tr) return false;
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
      copyFormat: () => ({ state }) => {
        this.storage.copied = captureFormat(state);
        return true;
      },
      pasteFormat: (fmt) => ({ state, dispatch }) => {
        const tr = applyFormatTr(state, fmt || this.storage.copied);
        if (!tr) return false;
        if (dispatch) dispatch(tr);
        return true;
      },
      // кисть: mode 1 — один раз, 2 — несколько раз, 0 — выключить
      setPainter: (mode) => ({ state, tr, dispatch }) => {
        if (dispatch) dispatch(tr.setMeta(painterKey, mode ? { fmt: captureFormat(state), mode } : { fmt: null, mode: 0 }));
        return true;
      },
    };
  },
  addKeyboardShortcuts() {
    const c = () => this.editor.commands;
    return {
      'Shift-F3': () => c().changeCase('cycle'),
      'Mod-Shift-c': () => c().copyFormat(),
      'Mod-Shift-v': () => c().pasteFormat(),
      Escape: () => (painterKey.getState(this.editor.state).mode ? c().setPainter(0) : false),
    };
  },
  addProseMirrorPlugins() {
    return [new Plugin({
      key: painterKey,
      state: {
        init: () => ({ fmt: null, mode: 0 }),
        apply: (tr, v) => tr.getMeta(painterKey) || v,
      },
      props: {
        attributes: (state) => (painterKey.getState(state).mode ? { class: 'pero-painter' } : {}),
      },
      // после отпускания мыши (выделение готово) — нанести формат
      view: (view) => {
        const onUp = () => setTimeout(() => {
          const p = painterKey.getState(view.state);
          if (!p.mode) return;
          const tr = applyFormatTr(view.state, p.fmt) || view.state.tr;
          if (p.mode === 1) tr.setMeta(painterKey, { fmt: null, mode: 0 });
          view.dispatch(tr);
        }, 0);
        view.dom.addEventListener('mouseup', onUp);
        return { destroy: () => view.dom.removeEventListener('mouseup', onUp) };
      },
    })];
  },
});
