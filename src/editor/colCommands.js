// Команды колонок текста («Макет → Колонки»): колонки раздела под курсором, выделенного текста (отдельный раздел
// «на текущей странице», как Word), до конца документа, всего документа. Раздел с несколькими колонками — в контейнере
// columns (columns.js), с одной — без него.
import { Extension } from '@tiptap/core';
import { parseCols, patchCols } from './columns.js';

const TEXT_BLOCKS = new Set(['paragraph', 'heading']);
const DEFAULT_SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';

// тип начала раздела (w:type)
const withType = (xml, type) => {
  const s = String(xml || DEFAULT_SECT).replace(/<w:type\b[^>]*\/>/g, '');
  if (type === 'nextPage') return s;
  const at = s.search(/<w:pgSz\b/);
  const tag = `<w:type w:val="${type}"/>`;
  return at >= 0 ? s.slice(0, at) + tag + s.slice(at) : s.replace(/(<w:sectPr\b[^>]*>)/, `$1${tag}`);
};

// Разделы документа: [{ from, to, close }] — границы узлов верхнего уровня; close — узел с sectPr (null — последний)
export function sectionsOf(doc) {
  const out = [];
  let from = 0;
  doc.forEach((n, pos) => {
    if (n.attrs && n.attrs.sectPr) { out.push({ from, to: pos + n.nodeSize, close: pos }); from = pos + n.nodeSize; }
  });
  out.push({ from, to: doc.content.size, close: null });
  return out;
}
const sectionIndexAt = (doc, pos) => {
  const list = sectionsOf(doc);
  const i = list.findIndex((s) => pos >= s.from && pos <= s.to);
  return i < 0 ? list.length - 1 : i;
};
// sectPr раздела (у контейнера — его атрибут, у последнего — атрибут документа)
const sectXml = (doc, s) => (s.close == null ? doc.attrs.sectPr : doc.nodeAt(s.close).attrs.sectPr);
const topNodes = (doc, s) => { const out = []; doc.nodesBetween(s.from, s.to, (n) => { out.push(n); return false; }); return out; };

// Колонки раздела k: sectPr + контейнер (обернуть / развернуть)
function setSectionCols(tr, k, cols) {
  const { schema } = tr.doc.type;
  const s = sectionsOf(tr.doc)[k];
  if (!s) return;
  const nodes = topNodes(tr.doc, s);
  if (!nodes.length) return;
  const box = nodes.length === 1 && nodes[0].type.name === 'columns' ? nodes[0] : null;
  const last = s.close == null;
  const next = patchCols(sectXml(tr.doc, s) || DEFAULT_SECT, cols);
  if (last) tr.setDocAttribute('sectPr', next);
  if (cols && cols.num > 1) {
    if (box) { if (!last) tr.setNodeMarkup(s.from, undefined, { ...box.attrs, sectPr: next }); return; }
    const inner = nodes.map((n, i) => (i === nodes.length - 1 && !last ? n.type.create({ ...n.attrs, sectPr: null }, n.content, n.marks) : n));
    tr.replaceWith(s.from, s.to, schema.nodes.columns.create({ sectPr: last ? null : next }, inner));
    return;
  }
  // одна колонка: разрыв раздела — снова у последнего абзаца (после таблицы — пустой абзац)
  if (!box) { if (!last) { const n = tr.doc.nodeAt(s.close); tr.setNodeMarkup(s.close, undefined, { ...n.attrs, sectPr: next }); } return; }
  tr.replaceWith(s.from, s.to, unwrapped(schema, box, last ? null : next));
}

function unwrapped(schema, box, xml) {
  const inner = [];
  box.forEach((n) => inner.push(n));
  if (xml) {
    const lastNode = inner[inner.length - 1];
    if (lastNode && TEXT_BLOCKS.has(lastNode.type.name) && !lastNode.attrs.sectPr) {
      inner[inner.length - 1] = lastNode.type.create({ ...lastNode.attrs, sectPr: xml }, lastNode.content, lastNode.marks);
    } else inner.push(schema.nodes.paragraph.create({ sectPr: xml }));
  }
  return inner;
}

// Развернуть контейнер раздела k (колонки остаются в sectPr) — чтобы делить раздел по блокам
function unwrapSection(tr, k) {
  const s = sectionsOf(tr.doc)[k];
  const box = s && tr.doc.nodeAt(s.from);
  if (!box || box.type.name !== 'columns') return;
  tr.replaceWith(s.from, s.to, unwrapped(tr.doc.type.schema, box, s.close == null ? null : box.attrs.sectPr));
}

// Разделить раздел перед узлом верхнего уровня at: первая часть закрывается копией sectPr (прежний тип начала), вторая
// начинается «на текущей странице». Возвращает false, если делить нечего (at — граница раздела)
function splitSectionBefore(tr, at) {
  const { schema } = tr.doc.type;
  const k = sectionIndexAt(tr.doc, at);
  const s = sectionsOf(tr.doc)[k];
  if (!s || at <= s.from || at >= s.to || tr.doc.resolve(at).depth !== 0) return false;
  const xml = sectXml(tr.doc, s) || DEFAULT_SECT;
  const before = tr.doc.resolve(at).nodeBefore;
  if (before && TEXT_BLOCKS.has(before.type.name) && !before.attrs.sectPr) tr.setNodeMarkup(at - before.nodeSize, undefined, { ...before.attrs, sectPr: xml });
  else tr.insert(at, schema.nodes.paragraph.create({ sectPr: xml }));
  const s2 = sectionsOf(tr.doc)[k + 1];
  if (s2.close == null) tr.setDocAttribute('sectPr', withType(tr.doc.attrs.sectPr, 'continuous'));
  else { const n = tr.doc.nodeAt(s2.close); tr.setNodeMarkup(s2.close, undefined, { ...n.attrs, sectPr: withType(n.attrs.sectPr, 'continuous') }); }
  return true;
}

// Контейнеры по sectPr: раздел с несколькими колонками — в контейнере, с одной — без
function normalizeSections(tr) {
  for (let k = sectionsOf(tr.doc).length - 1; k >= 0; k--) {
    const s = sectionsOf(tr.doc)[k];
    const box = tr.doc.nodeAt(s.from);
    const xml = s.close == null ? tr.doc.attrs.sectPr : (box && box.type.name === 'columns' ? box.attrs.sectPr : sectXml(tr.doc, s));
    setSectionCols(tr, k, parseCols(xml));
  }
}

// Колонки раздела под курсором (меню, диалог)
export function columnsAt(state) {
  const { doc } = state;
  const s = sectionsOf(doc)[sectionIndexAt(doc, state.selection.from)];
  const box = doc.nodeAt(s.from);
  const xml = s.close == null ? doc.attrs.sectPr : (box && box.type.name === 'columns' ? box.attrs.sectPr : sectXml(doc, s));
  return { cols: parseCols(xml), xml };
}

// Транзакция «колонки»: cols — { num, space, sep, equal, cols } | null (одна колонка); scope — section | selection | forward | all
export function columnsTr(state, cols, scope = 'section') {
  const tr = state.tr;
  const top = (pos) => { const $p = tr.doc.resolve(pos); return $p.depth ? $p.before(1) : pos; };
  const topEnd = (pos) => { const $p = tr.doc.resolve(pos); return $p.depth ? $p.after(1) : pos; };
  const { from, to, empty } = state.selection;
  if (scope === 'all') {
    for (let k = sectionsOf(tr.doc).length - 1; k >= 0; k--) setSectionCols(tr, k, cols);
  } else if (scope === 'selection' && !empty) {
    // выделенные блоки — отдельным разделом «на текущей странице»
    let k = sectionIndexAt(tr.doc, from);
    unwrapSection(tr, k);
    if (splitSectionBefore(tr, top(tr.mapping.map(from)))) k += 1;
    // конец выделения — исходная позиция через все шаги транзакции
    const sk = sectionsOf(tr.doc)[k];
    const cut = Math.max(sk.from, Math.min(topEnd(tr.mapping.map(to, -1)), sk.to));
    if (cut < sk.to) splitSectionBefore(tr, cut);
    setSectionCols(tr, k, cols);
    normalizeSections(tr);
  } else if (scope === 'forward') {
    let k = sectionIndexAt(tr.doc, from);
    unwrapSection(tr, k);
    if (splitSectionBefore(tr, top(tr.mapping.map(from)))) k += 1;
    for (let j = sectionsOf(tr.doc).length - 1; j >= k; j--) setSectionCols(tr, j, cols);
    normalizeSections(tr);
  } else setSectionCols(tr, sectionIndexAt(tr.doc, from), cols);
  return tr;
}

export const ColumnsCommands = Extension.create({
  name: 'columnsCommands',
  addCommands() {
    return {
      setColumns: (cols, scope = 'section') => ({ state, dispatch }) => {
        if (dispatch) dispatch(columnsTr(state, cols, scope).scrollIntoView());
        return true;
      },
    };
  },
});
