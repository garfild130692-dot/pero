// Непечатаемые знаки (¶, Ctrl+Shift+8) и ввод особых символов как в Word: неразрывный пробел (Ctrl+Shift+Пробел),
// неразрывный дефис (Ctrl+Shift+-, U+2011 ↔ w:noBreakHyphen), мягкий перенос (Ctrl+-, U+00AD ↔ w:softHyphen).
// Знаки — декорации: ¶ и ↵ — виджеты нулевой ширины, точки пробелов — ::after поверх символа; раскладка строк
// (а значит, и разбиение на страницы) от показа знаков не меняется. При печати не видны.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export const marksKey = new PluginKey('formattingMarks');

const CHAR_CLASS = { ' ': 'pm-sp', ' ': 'pm-nbsp', '­': 'pm-shy', '‑': 'pm-nbh' };
const CHAR_RE = /[  ­‑]/g;

const markEl = (cls, text) => () => {
  const el = document.createElement('span');
  el.className = `pm-mark ${cls}`;
  el.textContent = text;
  el.contentEditable = 'false';
  return el;
};

// Декорации одного текстового блока (pos — позиция блока)
function blockDecos(node, pos, out) {
  node.forEach((child, offset) => {
    const at = pos + 1 + offset;
    if (child.isText) {
      CHAR_RE.lastIndex = 0;
      let m;
      while ((m = CHAR_RE.exec(child.text))) out.push(Decoration.inline(at + m.index, at + m.index + 1, { class: CHAR_CLASS[m[0]] }));
    } else if (child.type.name === 'hardBreak') {
      out.push(Decoration.widget(at, markEl('pm-br', '↵'), { side: -1, key: 'br', marks: [] }));
    }
  });
  out.push(Decoration.widget(pos + node.nodeSize - 1, markEl('pm-para', '¶'), { side: 1, key: 'para', marks: [] }));
}

export function marksDecorations(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) { blockDecos(node, pos, out); return false; }
    return true;
  });
  return DecorationSet.create(doc, out);
}

// Пересчитать только изменённые текстовые блоки
function updateDecorations(set, tr) {
  let result = set.map(tr.mapping, tr.doc);
  const ranges = [];
  tr.mapping.maps.forEach((map, i) => map.forEach((_s, _e, from, to) => {
    const rest = tr.mapping.slice(i + 1);
    ranges.push([rest.map(from, -1), rest.map(to, 1)]);
  }));
  const seen = new Set();
  for (const [from, to] of ranges) {
    const a = Math.max(0, Math.min(from, tr.doc.content.size));
    const b = Math.max(a, Math.min(to, tr.doc.content.size));
    tr.doc.nodesBetween(a, b, (node, pos) => {
      if (!node.isTextblock) return true;
      if (seen.has(pos)) return false;
      seen.add(pos);
      result = result.remove(result.find(pos, pos + node.nodeSize));
      const out = [];
      blockDecos(node, pos, out);
      result = result.add(tr.doc, out);
      return false;
    });
  }
  return result;
}

let stored = false;
try { stored = localStorage.getItem('pero.showMarks') === '1'; } catch { /* нет хранилища — выключено */ }

export const FormattingMarks = Extension.create({
  name: 'formattingMarks',
  addCommands() {
    return {
      toggleFormattingMarks: () => ({ state, tr, dispatch }) => {
        const on = !marksKey.getState(state).on;
        try { localStorage.setItem('pero.showMarks', on ? '1' : '0'); } catch { /* не сохраняем */ }
        if (dispatch) dispatch(tr.setMeta(marksKey, { on }));
        return true;
      },
    };
  },
  addKeyboardShortcuts() {
    const ins = (ch) => () => this.editor.commands.insertContent(ch);
    return {
      'Mod-Shift-8': () => this.editor.commands.toggleFormattingMarks(),
      'Mod-Shift-Space': ins(' '),
      'Mod-Shift--': ins('‑'),
      'Mod--': ins('­'),
    };
  },
  addProseMirrorPlugins() {
    return [new Plugin({
      key: marksKey,
      state: {
        init: (_c, state) => ({ on: stored, set: stored ? marksDecorations(state.doc) : DecorationSet.empty }),
        apply: (tr, v, _old, state) => {
          const meta = tr.getMeta(marksKey);
          if (meta) return { on: meta.on, set: meta.on ? marksDecorations(state.doc) : DecorationSet.empty };
          if (!v.on || !tr.docChanged) return v;
          return { on: true, set: updateDecorations(v.set, tr) };
        },
      },
      props: {
        decorations: (state) => marksKey.getState(state).set,
        attributes: (state) => (marksKey.getState(state).on ? { class: 'show-marks' } : {}),
      },
    })];
  },
});
