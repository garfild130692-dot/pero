// Табуляция как в Word.
//
// Символ табуляции — inline-узел `tab` (а не «\t» в тексте): его ширину выставляет раскладка по позициям табуляции.
// Позиции — атрибут абзаца `tabs`: [{ pos (pt от левого поля текста), align: left|center|right|decimal|bar,
// leader: none|dot|hyphen|underscore|heavy|middleDot }] — действующий список (стиль + прямые, «clear» уже учтены).
// Шаг позиций по умолчанию — атрибут документа `defaultTab` (pt, из settings.xml w:defaultTabStop; в русском Word 1,25 см).
// Узел tab с атрибутом `ptab` — «абсолютная» табуляция (w:ptab, колонтитулы): выравнивание относительно полей.
//
// Правила Word (раскладка, plugin tabLayout):
//  • позиция отсчитывается от левого поля (в таблице — от левого края ячейки);
//  • выступ создаёт неявную позицию на отступе слева;
//  • позиции по умолчанию — через defaultTab от поля, но только правее последней заданной позиции;
//  • left — текст начинается на позиции; right — заканчивается; center — центр на позиции; decimal — разделитель на позиции;
//  • bar — только вертикальная черта, позицией для символа табуляции не является.
import { Extension, Node, mergeAttributes } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';

const TYPES = ['paragraph', 'heading'];
export const DEFAULT_TAB = 35.4; // 708 твипов = 1,25 см
export const TAB_ALIGNS = ['left', 'center', 'right', 'decimal', 'bar'];
export const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];

// ---------- OOXML ----------
const ALIGN_IN = { left: 'left', start: 'left', num: 'left', center: 'center', right: 'right', end: 'right', decimal: 'decimal', bar: 'bar' };
// <w:tabs> → { set: [stops], clear: [pos] } (pos в pt)
export function readTabs(tabsEl, children, wAttr) {
  const out = { set: [], clear: [] };
  for (const t of children(tabsEl, 'tab')) {
    const pos = (+wAttr(t, 'pos') || 0) / 20;
    const val = wAttr(t, 'val') || 'left';
    if (val === 'clear') { out.clear.push(pos); continue; }
    const leader = wAttr(t, 'leader') || 'none';
    out.set.push({ pos, align: ALIGN_IN[val] || 'left', leader: TAB_LEADERS.includes(leader) ? leader : 'none' });
  }
  return out;
}
// Применить слой (стиль или прямые) к действующему списку
export function applyTabLayer(list, layer) {
  if (!layer) return list;
  const same = (a, b) => Math.abs(a - b) < 0.05;
  let out = list.filter((s) => !layer.clear.some((p) => same(p, s.pos)) && !layer.set.some((n) => same(n.pos, s.pos)));
  out = [...out, ...layer.set];
  return out.sort((a, b) => a.pos - b.pos);
}
// Действующий список → <w:tabs> (inherited — позиции стиля: убранные в абзаце пишем как clear)
export function tabsXml(list, inherited = []) {
  const same = (a, b) => Math.abs(a - b) < 0.05;
  const items = [];
  for (const s of inherited) if (!(list || []).some((t) => same(t.pos, s.pos))) items.push({ pos: s.pos, clear: true });
  for (const s of list || []) {
    const st = inherited.find((t) => same(t.pos, s.pos));
    if (st && st.align === s.align && st.leader === s.leader) continue; // как в стиле — не повторяем
    items.push(s);
  }
  if (!items.length) return '';
  items.sort((a, b) => a.pos - b.pos);
  return `<w:tabs>${items.map((s) => (s.clear
    ? `<w:tab w:val="clear" w:pos="${Math.round(s.pos * 20)}"/>`
    : `<w:tab w:val="${s.align}"${s.leader && s.leader !== 'none' ? ` w:leader="${s.leader}"` : ''} w:pos="${Math.round(s.pos * 20)}"/>`)).join('')}</w:tabs>`;
}

// ---------- Узел «табуляция» ----------
export const TabNode = Node.create({
  name: 'tab',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      // абсолютная табуляция (w:ptab): { alignment: left|center|right, relativeTo: margin|indent, leader }
      ptab: { default: null, parseHTML: () => null, renderHTML: () => ({}) },
    };
  },

  parseHTML() { return [{ tag: 'span.pero-tab' }]; },
  renderHTML({ HTMLAttributes }) { return ['span', mergeAttributes(HTMLAttributes, { class: 'pero-tab' }), '\t']; },
  renderText() { return '\t'; },

  // Своя DOM-часть: ширину выставляет раскладка, ProseMirror её изменения не трогает
  addNodeView() {
    return () => {
      const dom = document.createElement('span');
      dom.className = 'pero-tab';
      dom.contentEditable = 'false';
      dom.textContent = '​';
      return { dom, ignoreMutation: () => true };
    };
  },
});

// ---------- Раскладка ширин ----------
const PX_PER_PT = 4 / 3;

// Ширина текста от after до before (или до конца абзаца), только первая строка
function segmentWidth(p, after, before, scale) {
  const r = document.createRange();
  r.setStartAfter(after);
  if (before) r.setEndBefore(before); else r.setEnd(p, p.childNodes.length);
  const rects = r.getClientRects();
  if (!rects.length) return 0;
  const top = rects[0].top;
  let left = Infinity; let right = -Infinity;
  for (const x of rects) { if (Math.abs(x.top - top) > 2) break; left = Math.min(left, x.left); right = Math.max(right, x.right); }
  return right > left ? (right - left) / scale : 0;
}
// Ширина текста до десятичного разделителя (для decimal)
function decimalWidth(p, after, before, scale) {
  const r = document.createRange();
  r.setStartAfter(after);
  if (before) r.setEndBefore(before); else r.setEnd(p, p.childNodes.length);
  const text = r.toString();
  const m = /\d([,.])\d/.exec(text);
  if (!m) return segmentWidth(p, after, before, scale);
  // ищем текстовый узел и смещение разделителя
  let left = m.index + 1;
  const walker = document.createTreeWalker(r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  const start = document.createRange();
  start.setStartAfter(after);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!r.intersectsNode(n)) continue;
    const from = n === r.startContainer ? r.startOffset : 0;
    const len = n.length - from;
    if (left <= len) {
      start.setEnd(n, from + left);
      const rect = start.getBoundingClientRect();
      return rect.width / scale;
    }
    left -= len;
  }
  return segmentWidth(p, after, before, scale);
}

const layoutCache = new WeakMap(); // DOM абзаца → ключ последней раскладки

export function layoutTabs(view, { force = false } = {}) {
  const { doc } = view.state;
  const defTab = doc.attrs.defaultTab || DEFAULT_TAB;
  const docKey = `${defTab}|${doc.attrs.fontFamily}|${doc.attrs.fontSize}|${JSON.stringify(doc.attrs.para)}`;
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    let hasTab = false;
    node.forEach((c) => { if (c.type.name === 'tab') hasTab = true; });
    if (!hasTab) return false;
    const p = view.nodeDOM(pos);
    if (!p || p.nodeType !== 1) return false;
    const container = p.closest('td, th') || view.dom;
    const cr = container.getBoundingClientRect();
    const scale = container.offsetWidth ? cr.width / container.offsetWidth : 1;
    if (!scale || !isFinite(scale)) return false;
    // шрифт абзаца — фактический: CSS документа App применяет позже обновления ProseMirror
    const fs = getComputedStyle(p);
    const key = `${docKey}|${Math.round(cr.width / scale)}|${fs.fontFamily}|${fs.fontSize}`;
    const prev = layoutCache.get(p);
    if (!force && prev && prev.node === node && prev.key === key) return false;
    layoutParagraph(p, node, container, cr, scale, defTab);
    layoutCache.set(p, { node, key });
    return false;
  });
}

function layoutParagraph(p, node, container, cr, scale, defTab) {
  const cs = getComputedStyle(container);
  const refLeft = cr.left + (parseFloat(cs.paddingLeft) + parseFloat(cs.borderLeftWidth || 0)) * scale;
  const refRight = cr.right - (parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth || 0)) * scale;
  const ps = getComputedStyle(p);
  const indentLeftPt = ((p.getBoundingClientRect().left - refLeft) / scale) / PX_PER_PT;
  const indentFirstPt = parseFloat(ps.textIndent) / PX_PER_PT || 0;
  const textW = (refRight - refLeft) / scale / PX_PER_PT;

  const stops = (node.attrs.tabs || []).filter((s) => s.align !== 'bar').map((s) => ({ ...s }));
  // выступ — неявная позиция на отступе слева
  if (indentFirstPt < -0.5 && !stops.some((s) => Math.abs(s.pos - indentLeftPt) < 0.5)) {
    stops.push({ pos: indentLeftPt, align: 'left', leader: 'none', implicit: true });
  }
  stops.sort((a, b) => a.pos - b.pos);
  const lastCustom = (node.attrs.tabs || []).reduce((m, s) => Math.max(m, s.pos), -Infinity);
  const nextStop = (xPt) => {
    const s = stops.find((t) => t.pos > xPt + 0.3);
    if (s) return s;
    let k = Math.floor(xPt / defTab + 1e-6) + 1;
    while (k * defTab <= lastCustom + 0.3) k++;
    return { pos: k * defTab, align: 'left', leader: 'none' };
  };

  const tabs = Array.from(p.querySelectorAll('.pero-tab')).filter((t) => t.closest('p, h1, h2, h3, h4, h5, h6, li') === p || p.contains(t));
  for (const t of tabs) { t.style.width = '0px'; t.className = 'pero-tab'; }
  tabs.forEach((t, i) => {
    const next = tabs[i + 1] || null;
    const ptab = t.pmViewDesc && t.pmViewDesc.node ? t.pmViewDesc.node.attrs.ptab : null;
    const x = (t.getBoundingClientRect().left - refLeft) / scale; // px от поля
    let stop;
    if (ptab) {
      const at = { left: 0, center: textW / 2, right: textW }[ptab.alignment] ?? 0;
      stop = { pos: at, align: ptab.alignment === 'center' ? 'center' : ptab.alignment === 'right' ? 'right' : 'left', leader: ptab.leader || 'none' };
    } else stop = nextStop(x / PX_PER_PT);
    const target = stop.pos * PX_PER_PT;
    let w;
    if (stop.align === 'right') w = target - x - segmentWidth(p, t, next, scale);
    else if (stop.align === 'center') w = target - x - segmentWidth(p, t, next, scale) / 2;
    else if (stop.align === 'decimal') w = target - x - decimalWidth(p, t, next, scale);
    else w = target - x;
    // не помещается до позиции (или позиция за правым краем строки) — Word двигает к следующей; приближение: минимальная ширина
    const lineRight = (refRight - refLeft) / scale - (parseFloat(ps.marginRight) || 0);
    if (!(w > 0) || x + w > lineRight + 1) w = Math.max(0, Math.min(w, lineRight - x));
    // текст после правой/центральной позиции у самого поля: субпиксельное округление не должно переносить его на новую строку
    if (stop.align !== 'left') w = Math.max(0, w - 0.5);
    t.style.width = `${Math.max(0, w)}px`;
    if (stop.leader && stop.leader !== 'none') t.classList.add(`tab-leader-${stop.leader}`);
  });
  void indentFirstPt;
}

const tabLayoutKey = new PluginKey('tabLayout');
const toTabNodes = new PluginKey('tabChars');

export const Tabs = Extension.create({
  name: 'tabs',
  priority: 50, // Tab в списке и таблице обрабатывают их расширения (приоритет выше)

  addGlobalAttributes() {
    return [{
      types: TYPES,
      attributes: {
        tabs: { default: null, keepOnSplit: true, parseHTML: () => null, renderHTML: () => ({}) },
      },
    }];
  },

  addCommands() {
    const each = (state, tr, fn) => {
      let changed = false;
      state.doc.nodesBetween(tr.selection.from, tr.selection.to, (node, pos) => {
        if (!TYPES.includes(node.type.name)) return true;
        const next = fn(node.attrs.tabs || []);
        const norm = next && next.length ? [...next].sort((a, b) => a.pos - b.pos) : null;
        if (JSON.stringify(norm) !== JSON.stringify(node.attrs.tabs || null)) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, tabs: norm }); changed = true; }
        return false;
      });
      return changed;
    };
    const same = (a, b) => Math.abs(a - b) < 0.05;
    return {
      insertTab: () => ({ commands }) => commands.insertContent({ type: 'tab' }),
      setTabStops: (list) => ({ state, tr, dispatch }) => { const c = each(state, tr, () => list); if (dispatch && c) dispatch(tr); return c; },
      addTabStop: (stop) => ({ state, tr, dispatch }) => {
        const s = { align: 'left', leader: 'none', ...stop, pos: Math.round(stop.pos * 20) / 20 };
        const c = each(state, tr, (l) => [...l.filter((t) => !same(t.pos, s.pos)), s]);
        if (dispatch && c) dispatch(tr);
        return c;
      },
      removeTabStop: (pos) => ({ state, tr, dispatch }) => { const c = each(state, tr, (l) => l.filter((t) => !same(t.pos, pos))); if (dispatch && c) dispatch(tr); return c; },
      moveTabStop: (from, to) => ({ state, tr, dispatch }) => {
        const c = each(state, tr, (l) => l.map((t) => (same(t.pos, from) ? { ...t, pos: Math.round(to * 20) / 20 } : t)).filter((t, i, a) => a.findIndex((u) => same(u.pos, t.pos)) === i));
        if (dispatch && c) dispatch(tr);
        return c;
      },
      setDefaultTab: (pt) => ({ tr, dispatch }) => {
        if (!(pt > 0)) return false;
        if (dispatch) dispatch(tr.setDocAttribute('defaultTab', Math.round(pt * 20) / 20));
        return true;
      },
    };
  },

  addKeyboardShortcuts() {
    return {
      Tab: () => this.editor.commands.insertTab(),
    };
  },

  addProseMirrorPlugins() {
    return [
      // «\t» в тексте (вставка из буфера, старые данные) → узлы tab
      new Plugin({
        key: toTabNodes,
        appendTransaction(trs, _old, state) {
          if (!trs.some((t) => t.docChanged)) return null;
          const type = state.schema.nodes.tab;
          let tr = null;
          state.doc.descendants((node, pos) => {
            if (!node.isText || !node.text.includes('\t')) return true;
            tr = tr || state.tr;
            const parts = node.text.split('\t');
            const content = [];
            parts.forEach((part, i) => {
              if (i > 0) content.push(type.create(null, null, node.marks));
              if (part) content.push(state.schema.text(part, node.marks));
            });
            const from = tr.mapping.map(pos);
            tr.replaceWith(from, tr.mapping.map(pos + node.nodeSize), content);
            return false;
          });
          return tr;
        },
      }),
      // раскладка ширин сразу после обновления DOM — до замера страниц (pagination ждёт 40 мс)
      new Plugin({
        key: tabLayoutKey,
        view(view) {
          if (typeof window === 'undefined') return {};
          let timer = 0;
          const run = (force) => { if (!view.isDestroyed) { try { layoutTabs(view, { force }); } catch { /* раскладка — только вид */ } } };
          const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => { clearTimeout(timer); timer = setTimeout(() => run(true), 30); }) : null;
          if (ro) ro.observe(view.dom);
          // шрифты подгружаются позже первой раскладки
          const onFonts = () => run(true);
          if (document.fonts) {
            if (document.fonts.ready) document.fonts.ready.then(onFonts);
            document.fonts.addEventListener('loadingdone', onFonts);
          }
          let later = 0;
          run(true);
          return {
            update(v, prev) {
              if (v.state.doc === prev.doc) return;
              run(false); // сразу — до замера страниц
              clearTimeout(later);
              later = setTimeout(() => run(false), 80); // ещё раз — после того как React применит стили документа
            },
            destroy() {
              clearTimeout(timer); clearTimeout(later);
              if (ro) ro.disconnect();
              if (document.fonts) document.fonts.removeEventListener('loadingdone', onFonts);
            },
          };
        },
      }),
    ];
  },
});
