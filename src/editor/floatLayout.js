// Показ плавающих рисунков (anchor.js): после разбиения на страницы для каждого рисунка с обтеканием считается его
// место — по области привязки (колонка/поле/лист/абзац) и выравниванию или смещению, — и в DOM рисунка ставятся стили:
//   • «за текстом» / «перед текстом» — position: absolute (на раскладку не влияет; за текстом — z-index под текстом);
//   • «вокруг рамки», «по контуру», «сквозное» — float со стороны, к которой рисунок ближе; сдвиг вниз от строки
//     привязки — margin-top, а shape-outside: inset() убирает этот сдвиг из обтекания (текст идёт и над рисунком);
//     отступы от текста — distL/R/T/B;
//   • «сверху и снизу» — float шириной во всю колонку (рядом с рисунком текста нет).
// CSS обтекает только с одной стороны: рисунок посреди колонки, который Word обтекает с двух сторон, у нас — с одной.
// Поменялась раскладка — разбиение пересчитывается (событие pero-remeasure), не больше нескольких кругов подряд.
import { Extension } from '@tiptap/core';
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state';
import { pagePx, parseSectPr, sectionList } from './page.js';

const BLOCKS = 'p, h1, h2, h3, h4, h5, h6';

// Верх листов (px без масштаба от верха .page) — по полосам разбиения
function pageTops(pageEl) {
  const tops = [0];
  if (!pageEl) return tops;
  const bands = [...pageEl.querySelectorAll(':scope > .pero-bands > .pero-page-band')].sort((a, b) => a.offsetTop - b.offsetTop);
  for (const b of bands) {
    const [foot, gap] = b.children;
    tops.push(b.offsetTop + (foot ? foot.offsetHeight : 0) + (gap ? gap.offsetHeight : 0));
  }
  return tops;
}

// Раздел позиции (номер в sectionList)
function sectionIndex(doc, pos) {
  let k = 0;
  doc.forEach((n, p) => { if (p + n.nodeSize <= pos && n.attrs && n.attrs.sectPr) k++; });
  return k;
}

const setStyle = (el, css) => {
  if (el.dataset.floatCss === css) return false;
  el.dataset.floatCss = css;
  el.style.cssText = css;
  return true;
};

// Расставить плавающие рисунки; возвращает true, если что-то сдвинулось (нужен новый замер страниц).
// pending — позиции рисунков, которые пользователь только что поставил (меню «Положение», перетаскивание): если их место
// выше строки привязки или на другом листе, рисунок переносится к абзацу на этой высоте (как Word перепривязывает).
export function placeFloats(view, pending = new Set()) {
  const { doc } = view.state;
  const items = [];
  doc.descendants((n, pos) => { if (n.type.name === 'image' && n.attrs.anchor) items.push({ n, pos }); });
  if (!items.length) return false;
  const root = view.dom;
  const pageEl = root.closest('.page');
  const rr = root.getBoundingClientRect();
  const scale = root.offsetWidth ? rr.width / root.offsetWidth : 1;
  const pr = pageEl ? pageEl.getBoundingClientRect() : rr;
  const tops = pageTops(pageEl);
  const sects = sectionList(doc).map((x) => pagePx(parseSectPr(x)));
  const rcs = getComputedStyle(root);
  let changed = false;
  for (const { n, pos } of items) {
    const dom = view.nodeDOM(pos);
    if (!dom || dom.nodeType !== 1) continue;
    const a = n.attrs.anchor;
    const img = dom.querySelector('img');
    const w = (n.attrs.width || (img ? img.offsetWidth : 96));
    const h = (n.attrs.height || (img ? img.offsetHeight : 96));
    const para = dom.closest(BLOCKS) || dom.parentElement;
    const cell = dom.closest('td, th');
    // всё дальше — px без масштаба, от левого верхнего угла .page
    const X = (cx) => (cx - pr.left) / scale;
    const Y = (cy) => (cy - pr.top) / scale;
    const block = para.closest('.page-content > *') || para;
    const secL = parseFloat(getComputedStyle(block).getPropertyValue('--sec-l')) || 0;
    const secR = parseFloat(getComputedStyle(block).getPropertyValue('--sec-r')) || 0;
    const g = sects[Math.min(sectionIndex(doc, pos), sects.length - 1)] || sects[0];
    let colL; let colR;
    if (cell) { const c = cell.getBoundingClientRect(); colL = X(c.left); colR = X(c.right); }
    else { colL = X(rr.left) + parseFloat(rcs.paddingLeft) + secL; colR = X(rr.right) - parseFloat(rcs.paddingRight) - secR; }
    const pageL = colL - g.left; const pageR = colR + g.right;
    // лист абзаца
    const pRect = para.getBoundingClientRect();
    const pTop = Y(pRect.top);
    let k = 0; while (k + 1 < tops.length && tops[k + 1] <= pTop + 1) k++;
    const pageTop = tops[k]; const pageBottom = pageTop + g.height;
    // по горизонтали
    const hr = a.h.rel;
    const [h0, h1] = hr === 'page' ? [pageL, pageR] : hr === 'leftMargin' || hr === 'insideMargin' ? [pageL, colL]
      : hr === 'rightMargin' || hr === 'outsideMargin' ? [colR, pageR] : hr === 'character' ? [X(pRect.left), colR] : [colL, colR];
    const al = a.h.align;
    const x = al === 'center' ? (h0 + h1 - w) / 2 : al === 'right' || al === 'outside' ? h1 - w : al === 'left' || al === 'inside' ? h0 : h0 + a.h.off;
    // по вертикали
    const vr = a.v.rel;
    const [v0, v1] = vr === 'page' ? [pageTop, pageBottom] : vr === 'margin' ? [pageTop + g.top, pageBottom - g.bottom]
      : vr === 'topMargin' ? [pageTop, pageTop + g.top] : vr === 'bottomMargin' ? [pageBottom - g.bottom, pageBottom] : [pTop, pTop + (pRect.height / scale)];
    const va = a.v.align;
    const y = va === 'center' ? (v0 + v1 - h) / 2 : va === 'bottom' ? v1 - h : va === 'top' ? v0 : v0 + a.v.off;
    const d = a.dist || { t: 0, b: 0, l: 0, r: 0 };
    // для перетаскивания: место рисунка от колонки и от абзаца
    dom.dataset.relx = (x - colL).toFixed(2);
    dom.dataset.rely = (y - pTop).toFixed(2);
    if (pending.has(pos) && a.wrap !== 'none') {
      const moved = reanchor(view, pos, n, a, { x, y, colL, scale, pr, pTop });
      if (moved) return 'moved';
    }
    let css;
    if (a.wrap === 'none') {
      // абсолютно: от блока, относительно которого позиционирован рисунок
      if (!dom.classList.contains('pf-abs')) dom.classList.add('pf-abs');
      const op = dom.offsetParent || root;
      const o = op.getBoundingClientRect();
      css = `position:absolute;left:${(x - X(o.left) - op.clientLeft).toFixed(1)}px;top:${(y - Y(o.top) - op.clientTop).toFixed(1)}px;z-index:${a.behind ? -1 : 3}`;
    } else {
      // float: статическое место — начало строки привязки (замер без сдвигов)
      if (dom.dataset.floatCss !== 'probe') { dom.style.cssText = 'float:left;margin:0'; dom.dataset.floatCss = 'probe'; }
      const st = dom.getBoundingClientRect();
      const lineTop = Y(st.top);
      const ps = getComputedStyle(para);
      const contL = X(pRect.left) + parseFloat(ps.paddingLeft) + parseFloat(ps.borderLeftWidth || 0);
      const contR = X(pRect.right) - parseFloat(ps.paddingRight) - parseFloat(ps.borderRightWidth || 0);
      const mt = Math.max(0, y - lineTop); // выше строки привязки float не поднять — у самой строки
      // +1 px: строку, которая задевает рисунок на доли пикселя, Word не обтекает (строка 14,3 pt над рисунком со сдвигом 14,2 pt)
      const inset = Math.max(0, mt - d.t + (mt > 0 ? 1 : 0));
      const shape = `shape-outside:inset(${inset.toFixed(1)}px 0 0 0) margin-box`;
      if (a.wrap === 'topBottom') {
        css = `float:left;margin:${mt.toFixed(1)}px ${Math.max(0, contR - x - w).toFixed(1)}px ${d.b}px ${(x - contL).toFixed(1)}px;${shape}`;
      } else if (x - colL < colR - x - w - 1) {
        // текст — со стороны, где места больше; поровну (рисунок по центру) — слева, как начинает строки Word
        css = `float:left;margin:${mt.toFixed(1)}px ${d.r}px ${d.b}px ${(x - contL).toFixed(1)}px;${shape}`;
      } else {
        css = `float:right;margin:${mt.toFixed(1)}px ${(contR - x - w).toFixed(1)}px ${d.b}px ${d.l}px;${shape}`;
      }
    }
    if (setStyle(dom, css)) changed = true;
  }
  return changed;
}

// Перенести рисунок к абзацу, который на высоте y (того же листа): float не поднимается выше своей строки привязки.
// Возвращает true, если документ изменён.
function reanchor(view, pos, node, a, { x, y, colL, scale, pr, pTop }) {
  const dom = view.nodeDOM(pos);
  // строка привязки — верх рисунка без его сдвига вниз (margin-top)
  const line = dom ? dom.getBoundingClientRect().top - (parseFloat(getComputedStyle(dom).marginTop) || 0) * scale : null;
  const lineY = line == null ? pTop : (line - pr.top) / scale;
  if (y >= lineY - 2) return false; // ниже строки привязки — переносить незачем
  const hit = view.posAtCoords({ left: pr.left + (colL + 2) * scale, top: pr.top + (y + 2) * scale });
  if (!hit) return false;
  const $p = view.state.doc.resolve(hit.pos);
  let d = $p.depth;
  while (d > 0 && !$p.node(d).isTextblock) d--;
  if (d < 1) return false;
  const start = $p.start(d);
  if (start <= pos && pos < $p.end(d)) return false; // уже в этом абзаце
  const blockDom = view.nodeDOM($p.before(d));
  const newTop = blockDom && blockDom.getBoundingClientRect ? (blockDom.getBoundingClientRect().top - pr.top) / scale : y;
  const v = a.v.align ? a.v : { rel: 'paragraph', align: null, off: Math.round((y - newTop) * 100) / 100 };
  const tr = view.state.tr.delete(pos, pos + node.nodeSize);
  const at = tr.mapping.map(start);
  tr.insert(at, node.type.create({ ...node.attrs, anchor: { ...a, v, h: a.h.align ? a.h : { rel: 'column', align: null, off: Math.round((x - colL) * 100) / 100 } } }));
  tr.setSelection(NodeSelection.create(tr.doc, at));
  view.dispatch(tr.setMeta(floatLayoutKey, { done: pos }));
  return true;
}

export const floatLayoutKey = new PluginKey('floatLayout');

export const FloatLayout = Extension.create({
  name: 'floatLayout',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: floatLayoutKey,
      // позиции рисунков, ждущих перепривязки (meta { pending: pos } — поставил пользователь; { done } — перенесён)
      state: {
        init: () => new Set(),
        apply(tr, set) {
          const meta = tr.getMeta(floatLayoutKey);
          let next = tr.docChanged ? new Set([...set].map((p) => tr.mapping.map(p))) : set;
          if (meta && meta.pending != null) next = new Set([...next, meta.pending]);
          if (meta && meta.done != null) next = new Set([]);
          return next;
        },
      },
      view(view) {
        let timer = 0;
        let rounds = 0; // кругов «рисунки сдвинулись → новый замер» подряд
        const run = () => {
          clearTimeout(timer);
          timer = setTimeout(() => {
            if (view.isDestroyed) return;
            if (typeof matchMedia === 'function' && matchMedia('print').matches) return;
            const pending = floatLayoutKey.getState(view.state);
            const moved = placeFloats(view, pending);
            if (moved === 'moved') return; // документ изменён — разбиение и новый круг придут сами
            if (pending.size) view.dispatch(view.state.tr.setMeta(floatLayoutKey, { done: -1 }).setMeta('addToHistory', false));
            if (moved && rounds < 3) { rounds++; view.dom.dispatchEvent(new Event('pero-remeasure')); } else if (!moved) rounds = 0;
          }, 10);
        };
        view.dom.addEventListener('pero-paginated', run);
        run();
        return {
          update: (v, prev) => { if (v.state.doc !== prev.doc) rounds = 0; },
          destroy: () => { clearTimeout(timer); view.dom.removeEventListener('pero-paginated', run); },
        };
      },
    })];
  },
});
