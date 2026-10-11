// «Сортировка» (группа «Абзац») как в Word: абзацы выделения (или пункты списка) — по тексту, числу или дате,
// по возрастанию/убыванию; в таблице — строки по столбцу (строка заголовка остаётся на месте).
import { Extension } from '@tiptap/core';
import { Fragment } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';

// Ключ сортировки из текста: число — первое число («1 234,5», «-7»), дата — дд.мм.гггг (или гггг-мм-дд)
export function sortKey(text, type) {
  const t = String(text || '').trim();
  if (type === 'number') {
    const m = /-?\d[\d\s ]*(?:[.,]\d+)?/.exec(t);
    return m ? parseFloat(m[0].replace(/[\s ]/g, '').replace(',', '.')) : null;
  }
  if (type === 'date') {
    let m = /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/.exec(t);
    if (m) { const y = +m[3] < 100 ? 2000 + +m[3] : +m[3]; return Date.UTC(y, m[2] - 1, +m[1]); }
    m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t);
    return m ? Date.UTC(+m[1], m[2] - 1, +m[3]) : null;
  }
  return t;
}

const collator = new Intl.Collator('ru', { numeric: true, sensitivity: 'base' });
// Сравнение как у Word: пустые и без ключа (не число/дата) — в конце при возрастании
export function compareKeys(a, b, type) {
  if (type === 'text') return collator.compare(a, b);
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;
  return a - b;
}

// Устойчивая сортировка узлов по тексту (textOf)
export function sortNodes(nodes, textOf, { type = 'text', desc = false } = {}) {
  return nodes.map((n, i) => ({ n, i, k: sortKey(textOf(n), type) }))
    .sort((x, y) => {
      if (type !== 'text' && (x.k == null) !== (y.k == null)) return x.k == null ? 1 : -1; // без ключа — всегда в конце
      const c = compareKeys(x.k, y.k, type);
      return (desc ? -c : c) || x.i - y.i;
    })
    .map((x) => x.n);
}

// Что сортировать под выделением: { kind: 'rows' | 'blocks', parentPos, from, to (индексы детей), cols? }
export function sortTarget(state) {
  const { $from, $to } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type.name === 'table') {
      const table = $from.node(d);
      let cols = 0;
      table.firstChild.forEach((c) => { cols += c.attrs.colspan || 1; });
      return { kind: 'rows', pos: $from.before(d), table, cols };
    }
  }
  // общий родитель выделения: список — его пункты, иначе — блоки (абзацы) между началом и концом
  let depth = $from.sharedDepth($to.pos);
  while (depth > 0 && $from.node(depth).isTextblock) depth--;
  for (let d = depth; d > 0; d--) {
    if (['bulletList', 'orderedList'].includes($from.node(d).type.name)) { depth = d; break; }
  }
  const parent = $from.node(depth);
  let a = $from.index(depth); let b = $to.index(depth);
  if (a === b && parent.childCount > 1 && !['bulletList', 'orderedList'].includes(parent.type.name)) return null; // один абзац — нечего сортировать
  if (['bulletList', 'orderedList'].includes(parent.type.name) && a === b) { a = 0; b = parent.childCount - 1; } // курсор в списке — весь список
  return { kind: 'blocks', parentPos: depth > 0 ? $from.before(depth) : -1, parent, from: a, to: b };
}

export function sortTr(state, { type = 'text', desc = false, column = 0, header = false } = {}) {
  const t = sortTarget(state);
  if (!t) return null;
  const tr = state.tr;
  if (t.kind === 'rows') {
    const rows = [];
    t.table.forEach((r) => rows.push(r));
    const head = header ? rows.slice(0, 1) : [];
    const body = header ? rows.slice(1) : rows;
    // текст столбца строки (с учётом объединённых по горизонтали ячеек)
    const cellText = (row) => {
      let c = 0; let out = '';
      row.forEach((cell) => { if (c <= column && column < c + (cell.attrs.colspan || 1)) out = cell.textContent; c += cell.attrs.colspan || 1; });
      return out;
    };
    const sorted = sortNodes(body, cellText, { type, desc });
    if (sorted.every((n, i) => n === body[i])) return null;
    tr.replaceWith(t.pos + 1, t.pos + t.table.nodeSize - 1, Fragment.from([...head, ...sorted]));
    return tr;
  }
  const kids = [];
  for (let i = t.from; i <= t.to; i++) kids.push(t.parent.child(i));
  const sorted = sortNodes(kids, (n) => n.textContent, { type, desc });
  if (sorted.every((n, i) => n === kids[i])) return null;
  let start = t.parentPos + 1;
  for (let i = 0; i < t.from; i++) start += t.parent.child(i).nodeSize;
  const end = start + kids.reduce((s, n) => s + n.nodeSize, 0);
  tr.replaceWith(start, end, Fragment.from(sorted));
  tr.setSelection(TextSelection.create(tr.doc, start + 1, Math.max(start + 1, end - 1)));
  return tr;
}

export const Sorting = Extension.create({
  name: 'sorting',
  addCommands() {
    return {
      sortParagraphs: (opts) => ({ state, dispatch }) => {
        const tr = sortTr(state, opts);
        if (!tr) return false;
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
      // «Горизонтальная линия»: пустой абзац с нижней границей после текущего
      insertHorizontalLine: () => ({ state, tr, dispatch }) => {
        const $f = state.selection.$from;
        let d = $f.depth;
        while (d > 1 && !$f.node(d).isTextblock) d--;
        const at = $f.after(d);
        const line = state.schema.nodes.paragraph.create({ pBorders: { bottom: { style: 'single', width: 0.75, color: '#a0a0a0', space: 1 } } });
        tr.insert(at, [line, state.schema.nodes.paragraph.create()]);
        tr.setSelection(TextSelection.create(tr.doc, at + line.nodeSize + 1));
        if (dispatch) dispatch(tr.scrollIntoView());
        return true;
      },
    };
  },
});
