// Разрывы как в Word: страницы (Ctrl+Enter), колонки (Ctrl+Shift+Enter), раздела (со следующей страницы, на текущей
// странице, с чётной/нечётной) и «Пустая страница». Разрыв страницы/колонки — значок opaque (w:br), абзац после него —
// новый; разрыв раздела — sectPr абзаца (копия параметров текущего раздела с w:type).
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { sectionAt, sectionList } from './page.js';

export const SECTION_TYPES = [['nextPage', 'Со следующей страницы'], ['continuous', 'На текущей странице'], ['evenPage', 'С чётной страницы'], ['oddPage', 'С нечётной страницы']];
const BR = { page: ['pageBreak', '<w:r><w:br w:type="page"/></w:r>'], column: ['columnBreak', '<w:r><w:br w:type="column"/></w:r>'] };
const DEFAULT_SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

// Тип начала раздела (w:type): nextPage по умолчанию
export const sectionType = (sectXml) => ((/<w:type\b[^>]*w:val="([^"]+)"/.exec(sectXml || '') || [])[1]) || 'nextPage';

// sectPr с типом начала раздела (w:type — по схеме перед pgSz)
export function withSectionType(sectXml, type) {
  let s = String(sectXml || DEFAULT_SECT).replace(/<w:type\b[^>]*\/>/g, '');
  if (type === 'nextPage') return s; // по умолчанию — без элемента
  const tag = `<w:type w:val="${type}"/>`;
  const at = s.search(/<w:pgSz\b/);
  return at >= 0 ? s.slice(0, at) + tag + s.slice(at) : s.replace(/(<w:sectPr\b[^>]*>)/, `$1${tag}`);
}

// Позиция текстового блока курсора (глубина) — разрыв вставляется в абзац (в таблице — в абзац ячейки)
function blockDepth($pos) {
  let d = $pos.depth;
  while (d > 0 && !$pos.node(d).isTextblock) d--;
  return d;
}

export const Breaks = Extension.create({
  name: 'breaks',
  addCommands() {
    return {
      // Разрыв страницы/колонки: значок в конце абзаца, текст после курсора — с нового абзаца (как в Word)
      insertBreak: (kind = 'page') => ({ state, tr, dispatch }) => {
        const [k, xml] = BR[kind] || BR.page;
        if (!state.selection.empty) tr.deleteSelection();
        const $p = tr.selection.$from;
        if (!$p.parent.isTextblock) return false;
        const node = state.schema.nodes.opaque.create({ kind: k, xml });
        tr.insert($p.pos, node);
        tr.split($p.pos + 1);
        tr.setSelection(TextSelection.create(tr.doc, $p.pos + 3));
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
      // «Пустая страница»: разрыв, пустой абзац, разрыв
      insertBlankPage: () => ({ chain }) => chain().insertBreak('page').insertBreak('page').command(({ tr }) => {
        // курсор — на пустой странице (абзац между разрывами)
        const $p = tr.selection.$from;
        const before = $p.before(blockDepth($p)) - 1;
        if (before > 0) tr.setSelection(TextSelection.near(tr.doc.resolve(before), -1));
        return true;
      }).run(),
      // Разрыв раздела: абзац до курсора закрывает раздел (копия параметров текущего, его начало — как было);
      // тип нового разрыва (w:type) — у следующего раздела, т.е. у текущих параметров (как в Word)
      insertSectionBreak: (type = 'nextPage') => ({ state, tr, dispatch }) => {
        if (!state.selection.empty) tr.deleteSelection();
        const $p = tr.selection.$from;
        const d = blockDepth($p);
        if (d !== 1) return false; // разрыв раздела — только в абзаце верхнего уровня (не в таблице/списке)
        const sec = sectionAt(state);
        const cur = sec.sectPr || DEFAULT_SECT;
        const pos = $p.pos;
        const blockPos = $p.before(d);
        tr.split(pos);
        const node = tr.doc.nodeAt(blockPos);
        tr.setNodeMarkup(blockPos, undefined, { ...node.attrs, sectPr: cur });
        // следующий раздел (где теперь курсор) начинается так, как выбрано
        if (sec.pos == null) tr.setDocAttribute('sectPr', withSectionType(cur, type));
        else {
          const at = tr.mapping.map(sec.pos);
          const n = tr.doc.nodeAt(at);
          if (n && n.attrs.sectPr) tr.setNodeMarkup(at, undefined, { ...n.attrs, sectPr: withSectionType(n.attrs.sectPr, type) });
        }
        tr.setSelection(TextSelection.create(tr.doc, pos + 2));
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
    };
  },
  addProseMirrorPlugins() {
    // подпись разрыва раздела — по типу начала следующего раздела (data-section-type у абзаца с sectPr)
    const key = new PluginKey('sectionBreakLabels');
    const build = (doc) => {
      const list = sectionList(doc);
      const decos = [];
      let i = 0;
      doc.forEach((n, pos) => {
        if (!n.attrs || !n.attrs.sectPr) return;
        i++;
        decos.push(Decoration.node(pos, pos + n.nodeSize, { 'data-section-type': sectionType(list[i]) }));
      });
      return DecorationSet.create(doc, decos);
    };
    return [new Plugin({
      key,
      state: { init: (_c, st) => build(st.doc), apply: (tr, v) => (tr.docChanged ? build(tr.doc) : v) },
      props: { decorations: (st) => key.getState(st) },
    })];
  },
  addKeyboardShortcuts() {
    return {
      'Mod-Enter': () => this.editor.commands.insertBreak('page'),
      'Mod-Shift-Enter': () => this.editor.commands.insertBreak('column'),
    };
  },
});
