// Режим исправлений (рецензирование): метки вставки/удаления, запись правок, принятие и отклонение.
//
// Метки insertion / deletion { id, author, date } — как w:ins / w:del в OOXML. В режиме исправлений:
// набор текста → вставка; Backspace/Delete/вырезание → текст не удаляется, а помечается удалённым
// (свои вставки удаляются по-настоящему, как в Word); вставка из буфера → вставка.
// Не отслеживаются: разбиение/слияние абзацев, изменения форматирования, таблиц, рисунков.
import { Extension, Mark } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import { wordDate } from './comments.js';

export const trackKey = new PluginKey('trackChanges');
const REV = ['insertion', 'deletion'];

const revAttrs = () => ({
  id: { default: null },
  author: { default: '' },
  date: { default: '' },
});

export const Insertion = Mark.create({
  name: 'insertion',
  inclusive: false,
  excludes: '',
  addAttributes: revAttrs,
  parseHTML: () => [{ tag: 'ins' }],
  renderHTML: ({ HTMLAttributes }) => ['ins', { class: 'rev-ins', title: `Вставлено: ${HTMLAttributes.author || ''}` }, 0],
});

export const Deletion = Mark.create({
  name: 'deletion',
  inclusive: false,
  excludes: '',
  addAttributes: revAttrs,
  parseHTML: () => [{ tag: 'del' }],
  renderHTML: ({ HTMLAttributes }) => ['del', { class: 'rev-del', title: `Удалено: ${HTMLAttributes.author || ''}` }, 0],
});

// Следующий свободный id правки (по документу)
function nextRevId(doc) {
  let max = 0;
  doc.descendants((n) => { for (const m of n.marks) if (REV.includes(m.type.name)) max = Math.max(max, +m.attrs.id || 0); });
  return max + 1;
}

// Пометить [from, to) удалённым: свои вставки — удалить по-настоящему, уже удалённое — не трогать
function markDeleted(state, tr, from, to, attrs) {
  const del = state.schema.marks.deletion;
  const ins = state.schema.marks.insertion;
  const realDelete = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const s = Math.max(from, pos);
    const e = Math.min(to, pos + node.nodeSize);
    if (s >= e) return false;
    const own = ins.isInSet(node.marks);
    if (own && own.attrs.author === attrs.author) realDelete.push([s, e]);
    else if (!node.isText) realDelete.push([s, e]); // рисунки, значки — без отслеживания
    else if (!del.isInSet(node.marks)) tr.addMark(s, e, del.create(attrs));
    return false;
  });
  for (const [s, e] of realDelete.reverse()) tr.delete(tr.mapping.map(s), tr.mapping.map(e));
}

// Непрерывный диапазон текста с той же меткой правки, что у символа в pos (для «принять/отклонить это»)
function revRangeAt(doc, pos) {
  const $p = doc.resolve(pos);
  const parent = $p.parent;
  if (!parent.isTextblock) return null;
  const start = $p.start();
  let found = null;
  const cells = [];
  parent.forEach((child, off) => cells.push({ child, from: start + off, to: start + off + child.nodeSize }));
  const idx = cells.findIndex((c) => pos >= c.from && pos <= c.to && c.child.marks.some((m) => REV.includes(m.type.name)));
  if (idx < 0) return null;
  const mark = cells[idx].child.marks.find((m) => REV.includes(m.type.name));
  let a = idx;
  let b = idx;
  const same = (c) => c && c.child.marks.some((m) => m.type === mark.type && m.attrs.id === mark.attrs.id);
  while (same(cells[a - 1])) a--;
  while (same(cells[b + 1])) b++;
  found = { from: cells[a].from, to: cells[b].to, type: mark.type.name, mark };
  return found;
}

// Все правки документа по порядку: [{ from, to, type }]
export function allRevisions(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    const m = node.marks.find((x) => REV.includes(x.type.name));
    if (!m) return false;
    const last = out[out.length - 1];
    if (last && last.to === pos && last.type === m.type.name && last.id === m.attrs.id) last.to = pos + node.nodeSize;
    else out.push({ from: pos, to: pos + node.nodeSize, type: m.type.name, id: m.attrs.id });
    return false;
  });
  return out;
}

// Принять (accept=true) или отклонить правки в диапазонах
function resolveRevisions(state, tr, ranges, accept) {
  const ins = state.schema.marks.insertion;
  const del = state.schema.marks.deletion;
  const remove = [];
  for (const r of ranges) {
    if ((r.type === 'insertion') === accept) tr.removeMark(r.from, r.to, r.type === 'insertion' ? ins : del);
    else remove.push(r);
  }
  for (const r of remove.sort((x, y) => y.from - x.from)) tr.delete(tr.mapping.map(r.from), tr.mapping.map(r.to));
}

export const TrackChanges = Extension.create({
  name: 'trackChanges',
  addOptions() { return { author: 'Пользователь' }; },

  addProseMirrorPlugins() {
    const ext = this;
    const revision = (state) => {
      const s = trackKey.getState(state);
      return { id: String(s.nextId), author: s.author || ext.options.author, date: wordDate() };
    };
    return [new Plugin({
      key: trackKey,
      state: {
        init: (_c, state) => ({ enabled: false, author: ext.options.author, nextId: nextRevId(state.doc) }),
        apply(tr, old) {
          const meta = tr.getMeta(trackKey);
          let next = old;
          if (meta) next = { ...next, ...meta };
          if (tr.getMeta('trackUsedId')) next = { ...next, nextId: next.nextId + 1 };
          return next;
        },
      },
      props: {
        // набор текста — вставка (поверх выделения — удаление выделенного + вставка)
        handleTextInput(view, from, to, text) {
          const { state } = view;
          if (!trackKey.getState(state).enabled) return false;
          const attrs = revision(state);
          const tr = state.tr;
          // продолжение своей вставки — тот же id (одна правка), иначе новый
          const before = state.doc.resolve(from).nodeBefore;
          const prevIns = before && state.schema.marks.insertion.isInSet(before.marks);
          const insAttrs = prevIns && prevIns.attrs.author === attrs.author ? prevIns.attrs : attrs;
          if (from < to) markDeleted(state, tr, from, to, attrs);
          const at = tr.mapping.map(to);
          const marks = (state.storedMarks || state.doc.resolve(from).marks()).filter((m) => !REV.includes(m.type.name));
          tr.insert(at, state.schema.text(text, [...marks, state.schema.marks.insertion.create(insAttrs)]));
          tr.setSelection(TextSelection.create(tr.doc, at + text.length));
          if (insAttrs === attrs) tr.setMeta('trackUsedId', true);
          view.dispatch(tr.scrollIntoView());
          return true;
        },
        handleDOMEvents: {
          // вырезание: скопировать как обычно, а текст пометить удалённым
          cut(view, event) {
            const { state } = view;
            if (!trackKey.getState(state).enabled || state.selection.empty) return false;
            event.preventDefault();
            const sel = state.selection;
            const slice = sel.content();
            const { dom, text } = view.serializeForClipboard ? view.serializeForClipboard(slice) : { dom: null, text: state.doc.textBetween(sel.from, sel.to, '\n') };
            if (event.clipboardData) {
              if (dom) event.clipboardData.setData('text/html', dom.innerHTML);
              event.clipboardData.setData('text/plain', text);
            }
            const tr = state.tr;
            markDeleted(state, tr, sel.from, sel.to, revision(state));
            tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(sel.from)));
            view.dispatch(tr.setMeta('trackUsedId', true));
            return true;
          },
        },
      },
      // вставка из буфера/перетаскивание в режиме исправлений — пометить вставленный текст
      appendTransaction(trs, oldState, newState) {
        if (!trackKey.getState(newState).enabled) return null;
        const t = trs.find((x) => x.docChanged && ['paste', 'drop'].includes(x.getMeta('uiEvent')));
        if (!t) return null;
        const attrs = revision(newState);
        const tr = newState.tr;
        t.mapping.maps.forEach((map, i) => map.forEach((_os, _oe, ns, ne) => {
          const rest = t.mapping.slice(i + 1);
          const a = rest.map(ns, -1);
          const b = rest.map(ne, 1);
          if (b > a) tr.addMark(a, b, newState.schema.marks.insertion.create(attrs));
        }));
        return tr.setMeta('trackUsedId', true);
      },
    })];
  },

  addKeyboardShortcuts() {
    const del = (dir) => () => {
      const { state, view } = this.editor;
      if (!trackKey.getState(state).enabled) return false;
      const attrs = { id: String(trackKey.getState(state).nextId), author: trackKey.getState(state).author || this.options.author, date: wordDate() };
      const sel = state.selection;
      const tr = state.tr;
      if (!sel.empty) {
        markDeleted(state, tr, sel.from, sel.to, attrs);
        tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(dir < 0 ? sel.from : sel.to)));
        view.dispatch(tr.setMeta('trackUsedId', true));
        return true;
      }
      // пропустить уже удалённые символы в сторону удаления
      const $p = sel.$from;
      if (!$p.parent.isTextblock) return false;
      let pos = sel.from;
      const delMark = state.schema.marks.deletion;
      for (;;) {
        const node = dir < 0 ? state.doc.resolve(pos).nodeBefore : state.doc.resolve(pos).nodeAfter;
        if (!node || !node.isText || !delMark.isInSet(node.marks)) break;
        pos += dir * node.nodeSize;
      }
      const $q = state.doc.resolve(pos);
      if (dir < 0 ? pos <= $q.start() : pos >= $q.end()) return false; // граница абзаца — без отслеживания
      const [from, to] = dir < 0 ? [pos - 1, pos] : [pos, pos + 1];
      markDeleted(state, tr, from, to, attrs);
      const cursor = dir < 0 ? tr.mapping.map(from) : tr.mapping.map(to);
      tr.setSelection(TextSelection.create(tr.doc, cursor));
      view.dispatch(tr.setMeta('trackUsedId', true));
      return true;
    };
    return { Backspace: del(-1), Delete: del(1) };
  },

  addCommands() {
    return {
      setTrackChanges: (enabled) => ({ tr, dispatch }) => { if (dispatch) dispatch(tr.setMeta(trackKey, { enabled })); return true; },
      setTrackAuthor: (author) => ({ tr, dispatch }) => { if (dispatch) dispatch(tr.setMeta(trackKey, { author })); return true; },
      // Принять/отклонить правку под курсором (или все в выделении)
      resolveChange: (accept) => ({ state, tr, dispatch }) => {
        const sel = state.selection;
        let ranges;
        if (sel.empty) {
          const r = revRangeAt(state.doc, sel.from) || revRangeAt(state.doc, Math.max(0, sel.from - 1));
          ranges = r ? [r] : [];
        } else ranges = allRevisions(state.doc).filter((r) => r.to > sel.from && r.from < sel.to)
          .map((r) => ({ ...r, from: Math.max(r.from, sel.from), to: Math.min(r.to, sel.to) }));
        if (!ranges.length) return false;
        if (dispatch) { resolveRevisions(state, tr, ranges, accept); dispatch(tr); }
        return true;
      },
      resolveAllChanges: (accept) => ({ state, tr, dispatch }) => {
        const ranges = allRevisions(state.doc);
        if (!ranges.length) return false;
        if (dispatch) { resolveRevisions(state, tr, ranges, accept); dispatch(tr); }
        return true;
      },
      // Перейти к следующей/предыдущей правке и выделить её
      goToChange: (dir) => ({ state, tr, dispatch }) => {
        const list = allRevisions(state.doc);
        if (!list.length) return false;
        const at = dir > 0 ? state.selection.to : state.selection.from;
        const r = dir > 0 ? (list.find((x) => x.from >= at) || list[0]) : ([...list].reverse().find((x) => x.to <= at) || list[list.length - 1]);
        if (dispatch) dispatch(tr.setSelection(TextSelection.create(tr.doc, r.from, r.to)).scrollIntoView());
        return true;
      },
    };
  },
});

// Число правок — по блокам верхнего уровня с кешем по узлу (правка не переходит границу блока): строка состояния и
// лента спрашивают это на каждое нажатие, а обход всего текста 200+ страниц — десятки миллисекунд
const revCache = new WeakMap();
function revisionCount(doc) {
  let n = 0;
  doc.forEach((block) => {
    let c = revCache.get(block);
    if (c == null) { c = allRevisions(block).length; revCache.set(block, c); }
    n += c;
  });
  return n;
}

export const trackInfo = (state) => {
  const s = trackKey.getState(state);
  return { enabled: !!(s && s.enabled), count: revisionCount(state.doc) };
};
