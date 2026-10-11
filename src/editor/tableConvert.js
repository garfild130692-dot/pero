// «Вставить таблицу…» (строки, столбцы, ширина) и преобразования «Текст → таблица» / «Таблица → текст», как в Word.
// Разделитель: табуляция (узел tab), «;», «,», свой знак или абзацы (каждый абзац — ячейка, число столбцов задаётся).
import { Extension } from '@tiptap/core';
import { Fragment } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';

export const PX_PER_CM = 96 / 2.54;

// Таблица rows × cols; widthCm — ширина столбца (null — авто, по ширине текста); header — первая строка повторяется
export function makeTable(schema, rows, cols, { widthCm = null, header = false, cells = null } = {}) {
  const colwidth = widthCm ? [Math.round(widthCm * PX_PER_CM)] : null;
  const rowsOut = [];
  for (let r = 0; r < rows; r++) {
    const cs = [];
    for (let c = 0; c < cols; c++) {
      const inline = cells && cells[r] && cells[r][c] ? cells[r][c] : [];
      cs.push(schema.nodes.tableCell.create({ colwidth }, schema.nodes.paragraph.create(null, inline.length ? inline : null)));
    }
    rowsOut.push(schema.nodes.tableRow.create(r === 0 && header ? { repeatHeader: true } : null, cs));
  }
  return schema.nodes.table.create(null, rowsOut);
}

// Разбить содержимое абзаца (inline-узлы) по разделителю: sep 'tab' — узел табуляции, иначе знак в тексте
export function splitInline(para, sep) {
  const parts = [[]];
  para.forEach((n) => {
    if (sep === 'tab') {
      if (n.type.name === 'tab') parts.push([]); else parts[parts.length - 1].push(n);
      return;
    }
    if (!n.isText) { parts[parts.length - 1].push(n); return; }
    const pieces = n.text.split(sep);
    pieces.forEach((t, i) => {
      if (i > 0) parts.push([]);
      if (t) parts[parts.length - 1].push(n.type.schema.text(t, n.marks));
    });
  });
  return parts;
}

// Абзацы выделения верхнего уровня: { from, to, paras }
function selectedParas(state) {
  const { $from, $to } = state.selection;
  if ($from.depth < 1 || $to.depth < 1) return null;
  const a = $from.index(0); const b = $to.index(0);
  const paras = [];
  let from = 0;
  state.doc.forEach((n, pos, i) => { if (i === a) from = pos; });
  let to = from;
  for (let i = a; i <= b; i++) {
    const n = state.doc.child(i);
    if (n.type.name !== 'paragraph') return null; // только простые абзацы (как и в Word — без таблиц внутри)
    paras.push(n);
    to += n.nodeSize;
  }
  return { from, to, paras };
}

// Определить разделитель по тексту выделения (как предлагает Word): табуляция → «;» → абзацы
export function guessSeparator(state) {
  const s = selectedParas(state);
  if (!s) return 'tab';
  let tab = 0; let semi = 0;
  for (const p of s.paras) p.forEach((n) => { if (n.type.name === 'tab') tab++; if (n.isText) semi += (n.text.match(/;/g) || []).length; });
  return tab ? 'tab' : semi ? ';' : 'para';
}

// Сколько столбцов получится при разделителе sep (для диалога)
export function columnsFor(state, sep) {
  const s = selectedParas(state);
  if (!s) return 0;
  if (sep === 'para') return 1;
  return Math.max(1, ...s.paras.map((p) => splitInline(p, sep).length));
}

export function textToTableTr(state, { sep = 'tab', cols = null } = {}) {
  const s = selectedParas(state);
  if (!s || !s.paras.length) return null;
  let cells;
  if (sep === 'para') {
    const n = Math.max(1, cols || 1);
    cells = [];
    s.paras.forEach((p, i) => {
      if (i % n === 0) cells.push([]);
      const inline = []; p.forEach((x) => inline.push(x));
      cells[cells.length - 1].push(inline);
    });
  } else cells = s.paras.map((p) => splitInline(p, sep));
  const ncols = sep === 'para' ? Math.max(1, cols || 1) : Math.max(cols || 0, ...cells.map((r) => r.length));
  const table = makeTable(state.schema, cells.length, ncols, { cells });
  const tr = state.tr.replaceWith(s.from, s.to, table);
  tr.setSelection(TextSelection.near(tr.doc.resolve(s.from + 3)));
  return tr;
}

// Таблица под курсором → абзацы: строка — абзац, ячейки через разделитель (sep 'para' — каждая ячейка абзацем)
export function tableToTextTr(state, { sep = 'tab' } = {}) {
  const $f = state.selection.$from;
  let d = $f.depth;
  while (d > 0 && $f.node(d).type.name !== 'table') d--;
  if (d <= 0) return null;
  const table = $f.node(d);
  const pos = $f.before(d);
  const { schema } = state;
  const paras = [];
  const cellInline = (cell) => {
    const out = [];
    cell.forEach((blk, _o, i) => {
      if (i > 0 && out.length) out.push(schema.text(' '));
      if (blk.isTextblock) blk.forEach((x) => out.push(x)); else if (blk.textContent) out.push(schema.text(blk.textContent));
    });
    return out;
  };
  table.forEach((row) => {
    if (sep === 'para') { row.forEach((cell) => paras.push(schema.nodes.paragraph.create(null, cellInline(cell)))); return; }
    const inline = [];
    row.forEach((cell, _o, i) => {
      if (i > 0) inline.push(sep === 'tab' ? schema.nodes.tab.create() : schema.text(sep));
      inline.push(...cellInline(cell));
    });
    paras.push(schema.nodes.paragraph.create(null, inline));
  });
  const tr = state.tr.replaceWith(pos, pos + table.nodeSize, Fragment.from(paras));
  tr.setSelection(TextSelection.near(tr.doc.resolve(pos + 1)));
  return tr;
}

export const TableConvert = Extension.create({
  name: 'tableConvert',
  addCommands() {
    const run = (fn) => (opts) => ({ state, dispatch }) => {
      const tr = fn(state, opts);
      if (!tr) return false;
      if (dispatch) dispatch(tr.scrollIntoView());
      return true;
    };
    return {
      insertTableSized: (rows, cols, opts = {}) => ({ state, tr, dispatch }) => {
        const table = makeTable(state.schema, rows, cols, opts);
        if (dispatch) {
          tr.replaceSelectionWith(table);
          dispatch(tr.scrollIntoView());
        }
        return true;
      },
      textToTable: run(textToTableTr),
      tableToText: run(tableToTextTr),
    };
  },
});
