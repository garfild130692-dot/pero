// Поиск и замена: подсветка всех совпадений, переход по ним, замена текущего и всех.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export const searchKey = new PluginKey('search');
const OBJ = '￼'; // нетекстовые узлы (рисунки, значки, разрывы) — совпадения через них не идут

// Регулярное выражение поиска; «слово целиком» — по буквам/цифрам Юникода (\b в JS кириллицу не понимает)
export function buildRegex(query, { caseSensitive = false, wholeWord = false } = {}) {
  if (!query) return null;
  const esc = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const src = wholeWord ? `(?<![\\p{L}\\p{N}_])${esc}(?![\\p{L}\\p{N}_])` : esc;
  return new RegExp(src, `gu${caseSensitive ? '' : 'i'}`);
}

// Все совпадения в документе: [{ from, to }] (позиции документа)
export function findAll(doc, query, opts) {
  const re = buildRegex(query, opts);
  if (!re) return [];
  const out = [];
  doc.descendants((block, blockPos) => {
    if (!block.isTextblock) return true;
    let text = '';
    const map = []; // индекс символа → позиция документа
    block.forEach((child, off) => {
      const pos = blockPos + 1 + off;
      if (child.isText) {
        for (let i = 0; i < child.text.length; i++) map.push(pos + i);
        text += child.text;
      } else {
        map.push(pos);
        text += OBJ;
      }
    });
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(text))) {
      if (!m[0].length) { re.lastIndex++; continue; }
      const last = map[m.index + m[0].length - 1];
      out.push({ from: map[m.index], to: last + 1 });
    }
    return false;
  });
  return out;
}

const emptyState = { query: '', opts: {}, results: [], index: -1, deco: DecorationSet.empty };

function decorate(doc, results, index) {
  return DecorationSet.create(doc, results.map((r, i) => Decoration.inline(r.from, r.to, { class: i === index ? 'search-hit current' : 'search-hit' })));
}

// индекс ближайшего совпадения не раньше pos
const nearest = (results, pos) => {
  const i = results.findIndex((r) => r.from >= pos);
  return i < 0 ? (results.length ? 0 : -1) : i;
};

export const SearchReplace = Extension.create({
  name: 'searchReplace',

  addProseMirrorPlugins() {
    return [new Plugin({
      key: searchKey,
      state: {
        init: () => emptyState,
        apply(tr, old, _oldState, newState) {
          const meta = tr.getMeta(searchKey);
          if (meta) {
            const query = meta.query ?? old.query;
            const opts = meta.opts ?? old.opts;
            const results = findAll(newState.doc, query, opts);
            let index = meta.index ?? (results.length ? nearest(results, newState.selection.from) : -1);
            if (index >= results.length) index = results.length ? 0 : -1;
            return { query, opts, results, index, deco: decorate(newState.doc, results, index) };
          }
          if (tr.docChanged && old.query) {
            const results = findAll(newState.doc, old.query, old.opts);
            const index = results.length ? Math.min(Math.max(old.index, 0), results.length - 1) : -1;
            return { ...old, results, index, deco: decorate(newState.doc, results, index) };
          }
          return old;
        },
      },
      props: { decorations: (state) => searchKey.getState(state).deco },
    })];
  },

  addCommands() {
    const go = (dir) => ({ state, tr, dispatch, view }) => {
      const s = searchKey.getState(state);
      if (!s.results.length) return false;
      const index = s.index < 0 ? 0 : (s.index + dir + s.results.length) % s.results.length;
      const r = s.results[index];
      if (dispatch) {
        tr.setMeta(searchKey, { index }).setSelection(TextSelection.create(tr.doc, r.from, r.to)).scrollIntoView();
        dispatch(tr);
        // прокрутка к совпадению по центру области
        setTimeout(() => {
          const dom = view && view.domAtPos(r.from).node;
          const el = dom && (dom.nodeType === 1 ? dom : dom.parentElement);
          if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center' });
        }, 0);
      }
      return true;
    };
    return {
      setSearch: (query, opts = {}) => ({ tr, dispatch }) => {
        if (dispatch) dispatch(tr.setMeta(searchKey, { query, opts }));
        return true;
      },
      clearSearch: () => ({ tr, dispatch }) => {
        if (dispatch) dispatch(tr.setMeta(searchKey, { query: '', opts: {} }));
        return true;
      },
      findNext: () => go(1),
      findPrev: () => go(-1),
      // Заменить текущее совпадение и перейти к следующему
      replaceCurrent: (text) => ({ state, tr, dispatch }) => {
        const s = searchKey.getState(state);
        const r = s.results[s.index];
        if (!r) return false;
        if (dispatch) {
          const marks = state.doc.resolve(r.from + 1).marks();
          tr.replaceWith(r.from, r.to, text ? state.schema.text(text, marks) : []);
          const after = r.from + text.length;
          tr.setMeta(searchKey, { index: undefined });
          tr.setSelection(TextSelection.create(tr.doc, after));
          dispatch(tr);
        }
        return true;
      },
      // Заменить все: одной транзакцией (одна отмена), с конца, чтобы позиции не сдвигались
      replaceAll: (text) => ({ state, tr, dispatch }) => {
        const s = searchKey.getState(state);
        if (!s.results.length) return false;
        if (dispatch) {
          for (const r of [...s.results].reverse()) {
            const marks = state.doc.resolve(r.from + 1).marks();
            tr.replaceWith(r.from, r.to, text ? state.schema.text(text, marks) : []);
          }
          tr.setMeta(searchKey, {});
          dispatch(tr);
        }
        return s.results.length;
      },
    };
  },
});

export const searchInfo = (state) => {
  const s = searchKey.getState(state);
  return s ? { count: s.results.length, index: s.index } : { count: 0, index: -1 };
};
