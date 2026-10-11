// Колонки текста (w:cols раздела). Блоки раздела с несколькими колонками лежат в узле-контейнере `columns`
// (документ ProseMirror плоский — иначе абзацы раздела не обернуть в многоколоночный блок). sectPr такого раздела —
// атрибут контейнера (в файле он в pPr последнего абзаца раздела; у последнего раздела документа — doc.attrs.sectPr,
// у контейнера null), поэтому код, который ищет разделы по узлам верхнего уровня (sectionList, sectionAt), работает
// как раньше. Вид — CSS-колонки; разбиение на листы — pagination.js (ряды колонок высотой с лист).
import { Node, mergeAttributes } from '@tiptap/core';

const TWIPS_PER_PX = 15;

// Колонки раздела из XML sectPr: { num, space (твипы), sep, equal, cols: [{ w, space }] } или null (одна колонка)
export function parseCols(sectXml) {
  const m = /<w:cols\b([^>]*?)(\/>|>([\s\S]*?)<\/w:cols>)/.exec(sectXml || '');
  if (!m) return null;
  const a = (name, src = m[1]) => { const x = new RegExp(`\\bw:${name}="([^"]*)"`).exec(src); return x ? x[1] : null; };
  const onOff = (v) => v != null && !['0', 'false', 'off'].includes(v);
  const cols = [];
  for (const c of (m[3] || '').matchAll(/<w:col\b([^>]*)\/>/g)) cols.push({ w: +a('w', c[1]) || 0, space: +a('space', c[1]) || 0 });
  const equal = a('equalWidth') == null ? true : onOff(a('equalWidth'));
  const num = equal ? (+a('num') || 1) : (cols.length || +a('num') || 1);
  if (num <= 1) return null;
  return { num, space: a('space') != null ? +a('space') : 720, sep: onOff(a('sep')), equal, cols: equal ? [] : cols };
}

// Записать колонки в sectPr (null — одна колонка). По схеме w:cols — после pgNumType, перед formProt/vAlign/…/docGrid
export function patchCols(sectXml, c) {
  let s = String(sectXml || '').replace(/<w:cols\b[^>]*\/>|<w:cols\b[^>]*>[\s\S]*?<\/w:cols>/g, '');
  let tag;
  if (!c || c.num <= 1) tag = `<w:cols w:space="${c && c.space != null ? c.space : 708}"/>`;
  else if (c.equal !== false || !c.cols || !c.cols.length) {
    tag = `<w:cols w:num="${c.num}"${c.sep ? ' w:sep="1"' : ''} w:space="${Math.round(c.space ?? 708)}"/>`;
  } else {
    tag = `<w:cols w:num="${c.cols.length}"${c.sep ? ' w:sep="1"' : ''} w:space="${Math.round(c.space ?? 708)}" w:equalWidth="0">${
      c.cols.map((x) => `<w:col w:w="${Math.round(x.w)}" w:space="${Math.round(x.space || 0)}"/>`).join('')}</w:cols>`;
  }
  const m = /<w:(?:formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(s);
  if (m) return s.slice(0, m.index) + tag + s.slice(m.index);
  return s.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`);
}

// Промежуток между колонками, px (разная ширина колонок — средний промежуток: CSS делает колонки только равными)
export function colsGap(c) {
  if (!c) return 0;
  return (c.equal !== false || !c.cols.length ? c.space : c.cols.slice(0, -1).reduce((s, x) => s + x.space, 0) / Math.max(1, c.cols.length - 1)) / TWIPS_PER_PX;
}
// Свойства колонок на DOM контейнера
export function applyColumnsStyle(dom, c) {
  const set = (k, v) => { if (dom.style.getPropertyValue(k) !== v) dom.style.setProperty(k, v); };
  set('column-count', String(c ? c.num : 1));
  set('column-gap', `${colsGap(c).toFixed(2)}px`);
  // разделитель: при печати — column-rule (через --pc-rule), на экране — слой линий разбиения (с начала раздела, не с верха листа)
  set('--pc-rule', c && c.sep ? '0.75pt solid currentColor' : 'none');
  if (c && c.sep) dom.dataset.sep = '1'; else delete dom.dataset.sep;
}

// ---------- JSON документа (импорт/экспорт) ----------
// Импорт: блоки раздела с колонками — в контейнер; sectPr последнего абзаца раздела переходит на контейнер
export function wrapColumnSections(content, bodySectPr) {
  const out = [];
  let start = 0;
  const flush = (end, sectPr, closing) => {
    const part = content.slice(start, end);
    if (!part.length) return;
    if (parseCols(sectPr)) {
      const inner = part.map((n) => n);
      if (closing) {
        const last = inner[inner.length - 1];
        inner[inner.length - 1] = { ...last, attrs: { ...(last.attrs || {}), sectPr: null } };
      }
      out.push({ type: 'columns', attrs: { sectPr: closing ? sectPr : null }, content: inner });
    } else out.push(...part);
  };
  content.forEach((n, i) => {
    if (n.attrs && n.attrs.sectPr) { flush(i + 1, n.attrs.sectPr, true); start = i + 1; }
  });
  flush(content.length, bodySectPr, false);
  return out;
}

// Экспорт: контейнер → его блоки; sectPr контейнера — в последний абзац (после таблицы — пустой абзац, как у Word)
export function unwrapColumnSections(content) {
  const out = [];
  for (const n of content || []) {
    if (n.type !== 'columns') { out.push(n); continue; }
    const inner = [...(n.content || [])];
    const sectPr = n.attrs && n.attrs.sectPr;
    if (sectPr) {
      const last = inner[inner.length - 1];
      if (last && (last.type === 'paragraph' || last.type === 'heading') && !(last.attrs && last.attrs.sectPr)) {
        inner[inner.length - 1] = { ...last, attrs: { ...(last.attrs || {}), sectPr } };
      } else inner.push({ type: 'paragraph', attrs: { sectPr } });
    }
    out.push(...inner);
  }
  return out;
}

// ---------- Узел ----------
export const Columns = Node.create({
  name: 'columns',
  group: 'sectionBlock', // не block: колонки не вкладываются в списки, таблицы и друг в друга
  content: 'block+',
  defining: true,
  isolating: false,

  addAttributes() {
    return {
      sectPr: { default: null, parseHTML: () => null, renderHTML: () => ({}) },
    };
  },

  parseHTML() {
    return [{ tag: 'div.pero-cols' }];
  },

  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { class: 'pero-cols' }), 0];
  },

  // NodeView: разбиение на листы ставит контейнеру стили раскладки (высота рядов, сдвиг первого ряда) прямо в DOM —
  // эти изменения атрибутов не должны перерисовывать узел
  addNodeView() {
    return ({ node, editor, getPos }) => {
      const dom = document.createElement('div');
      dom.className = 'pero-cols';
      // число колонок, промежуток, разделитель — свойства по отдельности: остальные (высота рядов, сдвиг первого ряда)
      // ставит разбиение на листы (applyColumnsStyle), их не трогаем
      const apply = (n) => {
        const pos = typeof getPos === 'function' ? getPos() : null;
        const xml = n.attrs.sectPr || (pos != null && isLastSection(editor.state.doc, pos) ? editor.state.doc.attrs.sectPr : null);
        applyColumnsStyle(dom, parseCols(xml));
        if (n.attrs.sectPr) dom.dataset.sect = '1'; else delete dom.dataset.sect; // раздел закрывается разрывом (не последний)
      };
      apply(node);
      return {
        dom,
        contentDOM: dom,
        update(n) { if (n.type.name !== 'columns') return false; apply(n); return true; },
        ignoreMutation: (m) => m.type === 'attributes' && m.target === dom,
      };
    };
  },
});

// Контейнер — в последнем разделе документа (после него нет узлов с sectPr)?
export function isLastSection(doc, pos) {
  let last = true;
  doc.forEach((n, p) => { if (p > pos && n.attrs && n.attrs.sectPr) last = false; });
  const node = doc.nodeAt(pos);
  return last && !(node && node.attrs.sectPr);
}

// sectPr раздела контейнера (последний раздел — из атрибутов документа)
export const columnsSectPr = (doc, pos) => {
  const n = doc.nodeAt(pos);
  return (n && n.attrs.sectPr) || (isLastSection(doc, pos) ? doc.attrs.sectPr : null);
};
