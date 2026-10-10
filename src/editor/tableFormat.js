// Таблицы: атрибуты таблицы/строки/ячейки, отрисовка и команды форматирования.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { Table, TableCell, TableHeader, TableRow, TableView } from '@tiptap/extension-table';
import { TableMap, selectedRect } from '@tiptap/pm/tables';
import { CELL_SIDES, DEFAULT_PEN, GRID_BORDERS, NONE, NO_BORDERS, TABLE_SIDES, borderCss } from './table.js';

// Стили таблиц: [ключ, название, образец] — готовые оформления (границы таблицы, заливка шапки и строк)
export const TABLE_STYLES = [
  ['grid', 'Сетка таблицы', '▦'],
  ['none', 'Без границ', '⬚'],
  ['horizontal', 'Только горизонтальные линии', '☰'],
  ['header', 'Сетка с заливкой шапки', '▤'],
  ['banded', 'Чередование строк', '▥'],
  ['frame', 'Строгая рамка', '▣'],
];
const HEADER_BG = '#dcefe9';
const BAND_BG = '#f2f7f5';
const line = (width = 0.5) => ({ ...DEFAULT_PEN, width });
function tableStyleDef(key) {
  const H = { ...NO_BORDERS, top: line(), bottom: line(), insideH: line() };
  switch (key) {
    case 'none': return { borders: NO_BORDERS };
    case 'horizontal': return { borders: H };
    case 'header': return { borders: GRID_BORDERS, header: HEADER_BG };
    case 'banded': return { borders: H, header: HEADER_BG, band: BAND_BG };
    case 'frame': return { borders: { top: line(1.5), bottom: line(1.5), left: line(1.5), right: line(1.5), insideH: line(0.5), insideV: line(0.5) } };
    default: return { borders: GRID_BORDERS };
  }
}
import { normalizeColor } from './font.js';

const CSS_VAR = { top: '--tb-top', left: '--tb-left', bottom: '--tb-bottom', right: '--tb-right', insideH: '--tb-ih', insideV: '--tb-iv' };

// <table> рисует TableView (нужен для изменения ширины столбцов мышью) — он не переносит атрибуты узла,
// поэтому границы по умолчанию и положение таблицы выставляем сами: CSS-переменные + отступы.
class PeroTableView extends TableView {
  constructor(node, cellMinWidth, view) {
    super(node, cellMinWidth, view);
    this.apply(node);
  }

  update(node) {
    if (!super.update(node)) return false;
    this.apply(node);
    return true;
  }

  apply(node) {
    const b = node.attrs.borders || GRID_BORDERS;
    for (const s of TABLE_SIDES) this.table.style.setProperty(CSS_VAR[s], borderCss(b[s]));
    const align = node.attrs.align;
    this.table.style.marginLeft = align === 'center' || align === 'right' ? 'auto' : '';
    this.table.style.marginRight = align === 'center' ? 'auto' : '';
  }
}

export const PeroTable = Table.extend({
  addOptions() {
    return { ...this.parent(), View: PeroTableView };
  },
  addAttributes() {
    return {
      ...this.parent?.(),
      borders: { default: GRID_BORDERS, rendered: false },
      align: { default: null, rendered: false }, // null | 'center' | 'right' — положение таблицы на странице
    };
  },
});

export const PeroTableRow = TableRow.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      height: { default: null, renderHTML: (a) => (a.height ? { style: `height: ${a.height}pt` } : {}) }, // pt
      heightRule: { default: null, rendered: false }, // null — «минимум», 'exact' — «точно»
      repeatHeader: {
        default: false,
        renderHTML: (a) => (a.repeatHeader ? { 'data-repeat-header': 'true' } : {}),
      },
      // «Не разрывать строку» (w:cantSplit): строка целиком на одном листе; по умолчанию Word делит строку между листами
      cantSplit: {
        default: false,
        renderHTML: (a) => (a.cantSplit ? { 'data-cant-split': 'true' } : {}),
      },
    };
  },
});

// Атрибуты оформления ячейки (общие для td и th)
const cellAttrs = (parent) => ({
  ...parent,
  borders: {
    default: null,
    renderHTML: (a) => {
      if (!a.borders) return {};
      const css = CELL_SIDES.filter((s) => a.borders[s]).map((s) => `--cb-${s}: ${borderCss(a.borders[s])}`);
      return css.length ? { style: css.join('; ') } : {};
    },
  },
  background: {
    default: null,
    parseHTML: (el) => normalizeColor(el.style.backgroundColor),
    renderHTML: (a) => (a.background ? { style: `background-color: ${a.background}` } : {}),
  },
  verticalAlign: {
    default: null, // null — сверху (как в Word по умолчанию)
    parseHTML: (el) => ({ middle: 'center', bottom: 'bottom' })[el.style.verticalAlign] || null,
    renderHTML: (a) => (a.verticalAlign ? { style: `vertical-align: ${a.verticalAlign === 'center' ? 'middle' : 'bottom'}` } : {}),
  },
});

export const PeroTableCell = TableCell.extend({ addAttributes() { return cellAttrs(this.parent?.()); } });
export const PeroTableHeader = TableHeader.extend({ addAttributes() { return cellAttrs(this.parent?.()); } });

// ---------- Команды ----------
// Обход ячеек выделенного прямоугольника: fn(cell, pos, { top, bottom, left, right } — сторона на краю прямоугольника)
function eachCellInRect(state, fn) {
  const rect = selectedRect(state);
  const { map, tableStart } = rect;
  for (const rel of map.cellsInRect(rect)) {
    const cell = rect.table.nodeAt(rel);
    const r = map.findCell(rel);
    fn(cell, tableStart + rel, {
      top: r.top === rect.top, bottom: r.bottom === rect.bottom, left: r.left === rect.left, right: r.right === rect.right,
    }, r);
  }
  return rect;
}

// Какие стороны ячейки получает режим «Границы»; undefined — не трогать
const MODES = {
  all: () => true,
  none: () => true,
  outer: (edge) => edge,
  inner: (edge) => !edge,
  top: (edge, side) => edge && side === 'top',
  bottom: (edge, side) => edge && side === 'bottom',
  left: (edge, side) => edge && side === 'left',
  right: (edge, side) => edge && side === 'right',
  insideH: (edge, side) => !edge && (side === 'top' || side === 'bottom'),
  insideV: (edge, side) => !edge && (side === 'left' || side === 'right'),
};
const OPPOSITE = { top: 'bottom', bottom: 'top', left: 'right', right: 'left' };

// Края таблицы для каждой ячейки (с учётом объединений): data-tb-edge="t l b r" — по нему ячейка берёт внешнюю
// или внутреннюю границу. Декорации, а не правка DOM: ProseMirror не перечитывает ячейки из-за чужих атрибутов.
const edgesKey = new PluginKey('tableEdges');
function edgeDecorations(doc) {
  const decos = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== 'table') return true;
    const map = TableMap.get(node);
    const seen = new Set();
    for (let i = 0; i < map.map.length; i++) {
      const off = map.map[i];
      if (seen.has(off)) continue;
      seen.add(off);
      const rect = map.findCell(off);
      const e = [];
      if (rect.top === 0) e.push('t');
      if (rect.left === 0) e.push('l');
      if (rect.bottom === map.height) e.push('b');
      if (rect.right === map.width) e.push('r');
      const cell = node.nodeAt(off);
      const from = pos + 1 + off;
      decos.push(Decoration.node(from, from + cell.nodeSize, { 'data-tb-edge': e.join(' ') }));
    }
    return false;
  });
  return DecorationSet.create(doc, decos);
}

export const TableFormat = Extension.create({
  addProseMirrorPlugins() {
    return [new Plugin({
      key: edgesKey,
      state: {
        init: (_c, state) => edgeDecorations(state.doc),
        apply: (tr, old) => (tr.docChanged ? edgeDecorations(tr.doc) : old),
      },
      props: { decorations: (state) => edgesKey.getState(state) },
    })];
  },

  name: 'tableFormat',

  addCommands() {
    return {
      // Границы выделенных ячеек (как кнопка «Границы» в Word). pen — граница для режимов кроме 'none'.
      // Стороны на краю выделения дублируются в соседние ячейки снаружи: иначе их линия перекрывала бы новую.
      setCellBorders: (mode, pen) => ({ state, tr, dispatch }) => {
        const want = MODES[mode];
        if (!want) return false;
        let rect;
        try { rect = selectedRect(state); } catch { return false; }
        if (!dispatch) return true;
        const value = mode === 'none' ? NONE : pen;
        const { map, tableStart, table } = rect;
        const patches = new Map(); // pos → borders
        const patch = (pos, cell, side, v) => {
          const cur = patches.get(pos) || { ...(cell.attrs.borders || {}) };
          cur[side] = v;
          patches.set(pos, cur);
        };
        eachCellInRect(state, (cell, pos, edge, r) => {
          for (const side of CELL_SIDES) {
            if (!want(edge[side], side)) continue;
            patch(pos, cell, side, value);
            if (!edge[side]) continue;
            // соседи за краем выделения
            const cols = side === 'top' || side === 'bottom' ? [r.left, r.right] : null;
            const rows = cols ? null : [r.top, r.bottom];
            const nr = side === 'top' ? r.top - 1 : side === 'bottom' ? r.bottom : null;
            const nc = side === 'left' ? r.left - 1 : side === 'right' ? r.right : null;
            if ((nr !== null && (nr < 0 || nr >= map.height)) || (nc !== null && (nc < 0 || nc >= map.width))) continue;
            const seen = new Set();
            if (cols) for (let c = cols[0]; c < cols[1]; c++) seen.add(map.map[nr * map.width + c]);
            else for (let rr = rows[0]; rr < rows[1]; rr++) seen.add(map.map[rr * map.width + nc]);
            for (const rel of seen) patch(tableStart + rel, table.nodeAt(rel), OPPOSITE[side], value);
          }
        });
        for (const [pos, borders] of patches) {
          const node = tr.doc.nodeAt(pos);
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders });
        }
        dispatch(tr);
        return true;
      },

      // Вставка строки/столбца с оформлением соседей (как Word): prosemirror-tables создаёт ячейки по умолчанию,
      // а мы копируем в них заливку, границы, выравнивание (и высоту строки) из строки/столбца, от которых вставляли.
      insertRowFormatted: (dir) => ({ state, chain }) => {
        let rect;
        try { rect = selectedRect(state); } catch { return false; }
        const src = dir > 0 ? rect.bottom - 1 : rect.top;
        const at = dir > 0 ? rect.bottom : rect.top; // индекс новой строки
        return chain()[dir > 0 ? 'addRowAfter' : 'addRowBefore']()
          .command(({ tr }) => copyFormatting(tr, rect.tableStart, 'row', at, dir > 0 ? src : src + 1))
          .run();
      },
      insertColumnFormatted: (dir) => ({ state, chain }) => {
        let rect;
        try { rect = selectedRect(state); } catch { return false; }
        const src = dir > 0 ? rect.right - 1 : rect.left;
        const at = dir > 0 ? rect.right : rect.left;
        return chain()[dir > 0 ? 'addColumnAfter' : 'addColumnBefore']()
          .command(({ tr }) => copyFormatting(tr, rect.tableStart, 'col', at, dir > 0 ? src : src + 1))
          .run();
      },

      // Строки выделения: высота (pt, null — авто) и правило
      setRowHeight: (height, rule = null) => ({ state, tr, dispatch }) => updateRows(state, tr, dispatch, () => ({ height, heightRule: height ? rule : null })),
      toggleRepeatHeader: () => ({ state, tr, dispatch }) => {
        let on = null;
        return updateRows(state, tr, dispatch, (row) => {
          if (on === null) on = !row.attrs.repeatHeader;
          return { repeatHeader: on };
        });
      },

      // Положение таблицы на странице: 'left' | 'center' | 'right'
      setTableAlign: (align) => ({ state, tr, dispatch }) => updateTable(state, tr, dispatch, () => ({ align: align === 'left' ? null : align })),
      // Границы таблицы по умолчанию (например, «Сетка таблицы»)
      setTableBorders: (borders) => ({ state, tr, dispatch }) => updateTable(state, tr, dispatch, () => ({ borders })),

      // Стиль таблицы: границы таблицы, заливка шапки/чередующихся строк; свои границы и заливка ячеек сбрасываются
      applyTableStyle: (key) => ({ state, tr, dispatch }) => {
        let rect;
        try { rect = selectedRect(state); } catch { return false; }
        if (!dispatch) return true;
        const def = tableStyleDef(key);
        const { table, tableStart } = rect;
        tr.setNodeMarkup(tableStart - 1, undefined, { ...table.attrs, borders: def.borders });
        let rowPos = tableStart;
        table.forEach((row, _o, r) => {
          let cellPos = rowPos + 1;
          row.forEach((cell) => {
            const bg = r === 0 && def.header ? def.header : def.band && r > 0 && r % 2 === 0 ? def.band : null;
            tr.setNodeMarkup(cellPos, undefined, { ...cell.attrs, borders: null, background: bg });
            cellPos += cell.nodeSize;
          });
          rowPos += row.nodeSize;
        });
        dispatch(tr);
        return true;
      },

      // «Выровнять ширину столбцов»: все столбцы одинаковые, общая ширина сохраняется
      distributeColumns: () => ({ state, tr, dispatch }) => {
        let rect;
        try { rect = selectedRect(state); } catch { return false; }
        const { map, table, tableStart } = rect;
        // общая ширина — сумма известных colwidth; если ширин нет — ширина текста страницы (16,5 см ≈ 624 px)
        let total = 0;
        for (let c = 0; c < map.width; c++) total += colWidth(table, map, c) || 0;
        if (!total) total = 623.6;
        const w = Math.round((total / map.width) * 100) / 100;
        if (!dispatch) return true;
        const done = new Set();
        for (const rel of map.map) {
          if (done.has(rel)) continue;
          done.add(rel);
          const cell = table.nodeAt(rel);
          tr.setNodeMarkup(tableStart + rel, undefined, { ...cell.attrs, colwidth: Array(cell.attrs.colspan).fill(w) });
        }
        dispatch(tr);
        return true;
      },
    };
  },
});

// Скопировать оформление ячеек строки/столбца src в новую строку/столбец at (индексы по сетке таблицы)
function copyFormatting(tr, tableStart, axis, at, src) {
  const table = tr.doc.nodeAt(tableStart - 1);
  if (!table || table.type.name !== 'table') return true;
  const map = TableMap.get(table);
  const len = axis === 'row' ? map.width : map.height;
  const idx = (i, k) => (axis === 'row' ? k * map.width + i : i * map.width + k);
  for (let i = 0; i < len; i++) {
    const rel = map.map[idx(i, at)];
    const r = map.findCell(rel);
    if ((axis === 'row' ? r.top : r.left) !== at) continue; // клетка — часть объединения, не новая
    const from = table.nodeAt(map.map[idx(i, src)]);
    const cell = table.nodeAt(rel);
    const { borders, background, verticalAlign } = from.attrs;
    tr.setNodeMarkup(tableStart + rel, undefined, { ...cell.attrs, borders, background, verticalAlign });
  }
  if (axis === 'row') {
    let pos = tableStart;
    for (let i = 0; i < at; i++) pos += table.child(i).nodeSize;
    const { height, heightRule } = table.child(src).attrs;
    tr.setNodeMarkup(pos, undefined, { ...table.child(at).attrs, height, heightRule });
  }
  return true;
}

// Ширина столбца c (px) по первой ячейке, у которой она известна
function colWidth(table, map, c) {
  for (let r = 0; r < map.height; r++) {
    const rel = map.map[r * map.width + c];
    const cell = table.nodeAt(rel);
    const rc = map.findCell(rel);
    const w = cell.attrs.colwidth && cell.attrs.colwidth[c - rc.left];
    if (w) return w;
  }
  return null;
}

function updateRows(state, tr, dispatch, fn) {
  let rect;
  try { rect = selectedRect(state); } catch { return false; }
  if (!dispatch) return true;
  const { table, tableStart } = rect;
  let pos = tableStart;
  for (let i = 0; i < table.childCount; i++) {
    const row = table.child(i);
    if (i >= rect.top && i < rect.bottom) tr.setNodeMarkup(pos, undefined, { ...row.attrs, ...fn(row) });
    pos += row.nodeSize;
  }
  dispatch(tr);
  return true;
}

function updateTable(state, tr, dispatch, fn) {
  let rect;
  try { rect = selectedRect(state); } catch { return false; }
  if (!dispatch) return true;
  const pos = rect.tableStart - 1;
  tr.setNodeMarkup(pos, undefined, { ...rect.table.attrs, ...fn(rect.table) });
  dispatch(tr);
  return true;
}

export { TableMap };
