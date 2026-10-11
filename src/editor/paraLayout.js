// Положение абзаца на странице и его рамка: «не отрывать от следующего», «не разрывать», «с новой страницы»,
// запрет висячих строк; границы абзаца (w:pBdr) и заливка (w:shd). Атрибуты узлов paragraph и heading.
//
// Атрибуты: keepNext, keepLines, pageBreakBefore, widowControl — true/false, null = как по умолчанию
// (у заголовков keepNext и keepLines включены — как стили Heading в Word; запрет висячих строк включён у всех);
// pBorders — null | { top, left, bottom, right, between: { style, width (pt), color, space (pt) } };
// shading — null | '#rrggbb'.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { borderFromOoxml } from './table.js';

const TYPES = ['paragraph', 'heading'];
// noHyphens ↔ w:suppressAutoHyphens, noLineNumbers ↔ w:suppressLineNumbers
export const LAYOUT_KEYS = ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl', 'noHyphens', 'noLineNumbers'];
export const PBDR_SIDES = ['top', 'left', 'bottom', 'right', 'between'];

// Умолчания Word (проверено через COM: новый абзац — WidowControl, заголовок — KeepWithNext + KeepTogether)
export function layoutDefaults(isHeading) {
  return { keepNext: !!isHeading, keepLines: !!isHeading, pageBreakBefore: false, widowControl: true, noHyphens: false, noLineNumbers: false };
}
export function effectiveLayout(node) {
  const d = layoutDefaults(node.type.name === 'heading');
  const out = {};
  for (const k of LAYOUT_KEYS) out[k] = node.attrs[k] ?? d[k];
  return out;
}

// ---------- OOXML ----------
const onOff = (el, wAttr) => {
  if (!el) return null;
  const v = String(wAttr(el, 'val') ?? 'true').toLowerCase();
  return !['0', 'false', 'off'].includes(v);
};
// Свойства pPr, заданные в этом элементе (для цепочки стилей): { keepNext?, keepLines?, …, pBorders?, shading? }
export function readLayoutProps(pPr, child, wAttr) {
  const out = {};
  if (!pPr) return out;
  const map = { keepNext: 'keepNext', keepLines: 'keepLines', pageBreakBefore: 'pageBreakBefore', widowControl: 'widowControl', noHyphens: 'suppressAutoHyphens', noLineNumbers: 'suppressLineNumbers' };
  for (const [k, tag] of Object.entries(map)) {
    const v = onOff(child(pPr, tag), wAttr);
    if (v !== null) out[k] = v;
  }
  const bdr = child(pPr, 'pBdr');
  if (bdr) {
    const b = {};
    for (const side of PBDR_SIDES) {
      const alt = { left: 'start', right: 'end' }[side];
      const el = child(bdr, side) || (alt && child(bdr, alt));
      if (!el) continue;
      const m = borderFromOoxml(wAttr(el, 'val'), wAttr(el, 'sz'), wAttr(el, 'color'));
      if (m.style === 'none') continue;
      b[side] = { ...m, space: +wAttr(el, 'space') || 0 };
    }
    out.pBorders = Object.keys(b).length ? b : null;
  }
  const shd = child(pPr, 'shd');
  if (shd) {
    const fill = String(wAttr(shd, 'fill') || '').toLowerCase();
    out.shading = /^[0-9a-f]{6}$/.test(fill) && wAttr(shd, 'val') !== 'nil' ? `#${fill}` : null;
  }
  return out;
}

// Значения для атрибутов узла: только отличия от умолчаний (у заголовка свои умолчания)
export function layoutAttrs(values, isHeading) {
  const d = layoutDefaults(isHeading);
  const out = {};
  for (const k of LAYOUT_KEYS) if (values[k] != null && values[k] !== d[k]) out[k] = values[k];
  if (values.pBorders) out.pBorders = values.pBorders;
  if (values.shading) out.shading = values.shading;
  return out;
}

const ptToSz = (pt) => Math.max(2, Math.round((pt || 0.5) * 8));
// XML свойств вёрстки в порядке схемы pPr: keepNext, keepLines, pageBreakBefore, [framePr], widowControl
export function layoutXmlHead(a) {
  let x = '';
  const flag = (tag, v) => { if (v === true) x += `<w:${tag}/>`; else if (v === false) x += `<w:${tag} w:val="0"/>`; };
  flag('keepNext', a.keepNext);
  flag('keepLines', a.keepLines);
  flag('pageBreakBefore', a.pageBreakBefore);
  flag('widowControl', a.widowControl);
  return x;
}
// …и после numPr: suppressLineNumbers, pBdr, shd
export function layoutXmlTail(a) {
  let x = '';
  if (a.noLineNumbers === true) x += '<w:suppressLineNumbers/>';
  else if (a.noLineNumbers === false) x += '<w:suppressLineNumbers w:val="0"/>';
  const b = a.pBorders;
  if (b) {
    const sides = PBDR_SIDES.filter((s) => b[s]);
    if (sides.length) {
      x += `<w:pBdr>${sides.map((s) => {
        const v = b[s];
        const color = v.color ? v.color.slice(1).toUpperCase() : 'auto';
        return `<w:${s} w:val="${v.style || 'single'}" w:sz="${ptToSz(v.width)}" w:space="${Math.round(v.space || 0)}" w:color="${color}"/>`;
      }).join('')}</w:pBdr>`;
    }
  }
  if (a.shading) x += `<w:shd w:val="clear" w:color="auto" w:fill="${a.shading.slice(1).toUpperCase()}"/>`;
  return x;
}

// ---------- Экран ----------
const BORDER_CSS = { single: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed' };
const sideCss = (b) => (b ? `${b.style === 'double' ? Math.max(b.width * 3, 2.25) : b.width}pt ${BORDER_CSS[b.style] || 'solid'} ${b.color || '#000'}` : 'none');
// Рамка и фон рисуются псевдоэлементом ::before (см. styles.css): вокруг текста, интервалы до/после — снаружи
function boxStyle(a) {
  const b = a.pBorders || {};
  const off = (s) => (b[s] ? (b[s].space || 0) + (b[s].style === 'double' ? b[s].width * 3 : b[s].width) : 0);
  return [
    `--pbx-top: ${sideCss(b.top)}`, `--pbx-left: ${sideCss(b.left)}`, `--pbx-bottom: ${sideCss(b.bottom)}`,
    `--pbx-right: ${sideCss(b.right)}`, `--pbx-between: ${sideCss(b.between)}`,
    `--pbx-bg: ${a.shading || 'transparent'}`,
    `--pbx-t: -${off('top')}pt`, `--pbx-b: -${off('bottom')}pt`, // интервалы до/после — margin, снаружи рамки
    `--pbx-l: -${off('left')}pt`, `--pbx-r: -${off('right')}pt`,
  ].join('; ');
}
// Ключ «одинаковой рамки»: соседние абзацы с одной рамкой и отступами Word рисует одним блоком
const boxKey = (n) => (n.attrs.pBorders || n.attrs.shading
  ? JSON.stringify([n.attrs.pBorders, n.attrs.shading, n.attrs.indentLeft, n.attrs.indentRight, n.type.name]) : null);

const groupKey = new PluginKey('paraBoxGroups');
// Промежуток между соседними абзацами (pt): больший из «после» верхнего и «перед» нижнего, как в Word и CSS
function gapPt(a, b, docAttrs) {
  const dp = (docAttrs && docAttrs.para) || {};
  const after = a.attrs.spaceAfter ?? (a.type.name === 'heading' ? 0 : dp.spaceAfter ?? 8);
  const before = b.attrs.spaceBefore ?? (b.type.name === 'heading' ? (b.attrs.level === 1 ? 12 : 2) : dp.spaceBefore ?? 0);
  return Math.max(after, before);
}
function groupDecorations(doc) {
  const decos = [];
  const scan = (parent, base) => {
    let prev = null;
    parent.forEach((node, off) => {
      const pos = base + off;
      const k = TYPES.includes(node.type.name) ? boxKey(node) : null;
      if (k && prev && prev.key === k) {
        // верхний блок дотягивается до нижнего через промежуток между абзацами
        decos.push(Decoration.node(prev.pos, prev.pos + prev.size, { class: 'pb-next', style: `--pbx-join: ${gapPt(prev.node, node, doc.attrs)}pt` }));
        decos.push(Decoration.node(pos, pos + node.nodeSize, { class: 'pb-prev' }));
      }
      prev = k ? { key: k, pos, size: node.nodeSize, node } : null;
      if (!node.isTextblock && node.childCount) scan(node, pos + 1);
    });
  };
  scan(doc, 0);
  return DecorationSet.create(doc, decos);
}

export const ParaLayout = Extension.create({
  name: 'paraLayout',

  addGlobalAttributes() {
    const flag = (name, dataAttr) => ({
      default: null,
      keepOnSplit: true,
      parseHTML: () => null,
      renderHTML: (a) => (a[name] == null ? {} : { [dataAttr]: String(a[name]) }),
    });
    return [{
      types: TYPES,
      attributes: {
        keepNext: flag('keepNext', 'data-keep-next'),
        keepLines: flag('keepLines', 'data-keep-lines'),
        pageBreakBefore: { ...flag('pageBreakBefore', 'data-page-break-before'), keepOnSplit: false },
        widowControl: flag('widowControl', 'data-widow-control'),
        noHyphens: flag('noHyphens', 'data-no-hyphens'),
        noLineNumbers: flag('noLineNumbers', 'data-no-line-numbers'),
        pBorders: {
          default: null,
          keepOnSplit: true,
          parseHTML: () => null,
          renderHTML: (a) => (a.pBorders || a.shading ? { 'data-pbox': '', style: boxStyle(a) } : {}),
        },
        shading: { default: null, keepOnSplit: true, parseHTML: () => null, renderHTML: () => ({}) }, // рисуется вместе с pBorders
      },
    }, {
      // Исходный стиль заголовка из файла { id, level } («Название», «Заголовок 1» с другим id…): при сохранении
      // заголовок того же уровня пишется с ним, а не с нашим Heading N — оформление стиля файла не теряется
      types: ['heading'],
      attributes: {
        fileStyle: { default: null, keepOnSplit: false, parseHTML: () => null, renderHTML: () => ({}) },
      },
    }];
  },

  addCommands() {
    const each = (state, tr, fn) => {
      const { from, to } = tr.selection;
      let changed = false;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!TYPES.includes(node.type.name)) return true;
        const next = fn(node);
        if (next) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next }); changed = true; }
        return false;
      });
      return changed;
    };
    const paras = (state) => {
      const out = [];
      state.doc.nodesBetween(state.selection.from, state.selection.to, (n) => { if (TYPES.includes(n.type.name)) { out.push(n); return false; } return true; });
      return out;
    };
    return {
      // { keepNext, keepLines, pageBreakBefore, widowControl } — true/false; значение по умолчанию сохраняется как null
      setParagraphLayout: (values) => ({ state, tr, dispatch }) => {
        const changed = each(state, tr, (node) => {
          const d = layoutDefaults(node.type.name === 'heading');
          const patch = {};
          for (const k of LAYOUT_KEYS) if (k in values) patch[k] = values[k] === d[k] ? null : values[k];
          return patch;
        });
        if (dispatch && changed) dispatch(tr);
        return changed;
      },
      // Границы как кнопка Word: 'bottom' | 'top' | 'left' | 'right' | 'between' — переключить сторону;
      // 'outer' — рамка, 'all' — рамка + между абзацами, 'none' — убрать. pen — { style, width, color }
      setParagraphBorders: (mode, pen) => ({ state, tr, dispatch }) => {
        const line = { style: 'single', width: 0.5, color: null, ...(pen || {}) };
        const space = (s) => (s === 'left' || s === 'right' ? 4 : s === 'between' ? 1 : 1); // как Word: 4 pt по бокам, 1 pt сверху/снизу
        const list = paras(state);
        const sides = { outer: ['top', 'left', 'bottom', 'right'], all: ['top', 'left', 'bottom', 'right', 'between'] }[mode] || [mode];
        // переключатель: если у всех абзацев уже есть сторона — снять
        const allHave = mode !== 'outer' && mode !== 'all' && mode !== 'none' && list.length > 0 && list.every((n) => n.attrs.pBorders && n.attrs.pBorders[mode]);
        const changed = each(state, tr, (node) => {
          if (mode === 'none') return node.attrs.pBorders ? { pBorders: null } : null;
          const b = { ...(node.attrs.pBorders || {}) };
          for (const s of sides) {
            if (allHave) delete b[s];
            else b[s] = { ...line, space: space(s) };
          }
          return { pBorders: Object.keys(b).length ? b : null };
        });
        if (dispatch && changed) dispatch(tr);
        return changed;
      },
      setParagraphShading: (color) => ({ state, tr, dispatch }) => {
        const changed = each(state, tr, () => ({ shading: color || null }));
        if (dispatch && changed) dispatch(tr);
        return changed;
      },
    };
  },

  addProseMirrorPlugins() {
    return [new Plugin({
      key: groupKey,
      state: {
        init: (_c, state) => groupDecorations(state.doc),
        apply: (tr, old) => (tr.docChanged ? groupDecorations(tr.doc) : old),
      },
      props: { decorations: (state) => groupKey.getState(state) },
    })];
  },
});

// w:suppressAutoHyphens — в pPr после tabs, перед spacing
export const hyphensXml = (a) => (a.noHyphens === true ? '<w:suppressAutoHyphens/>' : a.noHyphens === false ? '<w:suppressAutoHyphens w:val="0"/>' : '');
