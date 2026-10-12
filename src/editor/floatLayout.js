// Показ плавающих рисунков (anchor.js): после разбиения на страницы для каждого рисунка с обтеканием считается его
// место — по области привязки (колонка/поле/лист/абзац) и выравниванию или смещению, — и в DOM рисунка ставятся стили:
//   • «за текстом» / «перед текстом» — position: absolute (на раскладку не влияет; за текстом — z-index под текстом);
//   • «вокруг рамки», «по контуру», «сквозное» — float со стороны, к которой рисунок ближе; сдвиг вниз от строки
//     привязки — margin-top, а shape-outside: inset() убирает этот сдвиг из обтекания (текст идёт и над рисунком);
//     отступы от текста — distL/R/T/B;
//   • «сверху и снизу» — float шириной во всю колонку (рядом с рисунком текста нет).
// CSS обтекает только с одной стороны: рисунок посреди колонки, который Word обтекает с двух сторон, у нас — с одной.
// Рисунок выше своей строки привязки (Word так хранит перетащенные рисунки; float туда не поднять): на строке нужной
// высоты — виджет-заместитель (пустой float того же размера, даёт обтекание), сам рисунок — абсолютно на своём месте.
// Поменялась раскладка — разбиение пересчитывается (событие pero-remeasure), не больше нескольких кругов подряд.
import { Extension } from '@tiptap/core';
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { pagePx, parseSectPr, sectionList } from './page.js';
import { rotatedPad } from './picFormat.js';
import { shapeOf } from './shapes.js';

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
  // рисунки с обтеканием и фигуры/надписи с wp:anchor (opaque drawing, shapes.js)
  doc.descendants((n, pos) => {
    if (n.type.name === 'image' && n.attrs.anchor) items.push({ n, pos, a: n.attrs.anchor, rot: n.attrs.rotate });
    else if (n.type.name === 'opaque' && n.attrs.kind === 'drawing') {
      const s = shapeOf(n.attrs.xml);
      if (s && s.anchor) items.push({ n, pos, a: s.anchor, w: s.w, h: s.h, pad: s.effect });
    }
  });
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
  const proxies = [];
  for (const it of items) {
    const { n, pos, a } = it;
    const dom = view.nodeDOM(pos);
    if (!dom || dom.nodeType !== 1) continue;
    const img = dom.querySelector('img');
    const w = it.w || n.attrs.width || (img ? img.offsetWidth : 96);
    const h = it.h || n.attrs.height || (img ? img.offsetHeight : 96);
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
    // повёрнутый рисунок: Word ставит на место рамку без поворота, а узел у нас шире на запас под габарит
    // запас вокруг рамки по сторонам: у фигуры — wp:effectExtent, у повёрнутого рисунка — габарит поворота
    const rp = it.pad ? null : rotatedPad(w, h, it.rot);
    const pad = it.pad || { l: rp.x, r: rp.x, t: rp.y, b: rp.y };
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
      css = `position:absolute;left:${(x - pad.l - X(o.left) - op.clientLeft).toFixed(1)}px;top:${(y - pad.t - Y(o.top) - op.clientTop).toFixed(1)}px;z-index:${a.behind ? -1 : 3}`;
    } else {
      // float: статическое место — начало строки привязки (замер без сдвигов)
      if (dom.dataset.floatCss !== 'probe') { dom.style.cssText = 'float:left;margin:0'; dom.dataset.floatCss = 'probe'; }
      const st = dom.getBoundingClientRect();
      const lineTop = Y(st.top);
      const ps = getComputedStyle(para);
      const contL = X(pRect.left) + parseFloat(ps.paddingLeft) + parseFloat(ps.borderLeftWidth || 0);
      const contR = X(pRect.right) - parseFloat(ps.paddingRight) - parseFloat(ps.borderRightWidth || 0);
      // float в строке с верхом lt: отступы сверху и по сторонам, shape-outside
      const floatCss = (lt) => {
        const mt = Math.max(0, y - pad.t - lt);
        // +1 px: строку, которая задевает рисунок на доли пикселя, Word не обтекает (строка 14,3 pt над рисунком со сдвигом 14,2 pt)
        const inset = Math.max(0, mt - d.t + (mt > 0 ? 1 : 0));
        const shape = `shape-outside:inset(${inset.toFixed(1)}px 0 0 0) margin-box`;
        if (a.wrap === 'topBottom') return `float:left;margin:${mt.toFixed(1)}px ${Math.max(0, contR - x - w - pad.r).toFixed(1)}px ${d.b}px ${(x - pad.l - contL).toFixed(1)}px;${shape}`;
        // текст — со стороны, где места больше; поровну (рисунок по центру) — слева, как начинает строки Word
        if (x - colL < colR - x - w - 1) return `float:left;margin:${mt.toFixed(1)}px ${d.r}px ${d.b}px ${(x - pad.l - contL).toFixed(1)}px;${shape}`;
        return `float:right;margin:${mt.toFixed(1)}px ${(contR - x - w - pad.r).toFixed(1)}px ${d.b}px ${d.l}px;${shape}`;
      };
      const hit = y - pad.t < lineTop - 2 ? view.posAtCoords({ left: pr.left + (colL + 2) * scale, top: pr.top + (y - pad.t + 2) * scale }) : null;
      if (hit && hit.pos < pos) {
        // выше строки привязки: заместитель на строке этой высоты, сам рисунок — абсолютно
        let lt = y - pad.t;
        try { lt = Y(view.coordsAtPos(hit.pos).top); } catch { /* верх строки — по высоте рисунка */ }
        proxies.push({ pos: hit.pos, css: `${floatCss(lt)};width:${(w + pad.l + pad.r).toFixed(1)}px;height:${(h + pad.t + pad.b).toFixed(1)}px` });
        const op = dom.offsetParent || root;
        const o = op.getBoundingClientRect();
        css = `position:absolute;left:${(x - pad.l - X(o.left) - op.clientLeft).toFixed(1)}px;top:${(y - pad.t - Y(o.top) - op.clientTop).toFixed(1)}px`;
      } else css = floatCss(lineTop);
    }
    if (setStyle(dom, css)) changed = true;
  }
  // заместители поменялись — новый набор виджетов (одна транзакция с метаданными)
  const st = floatLayoutKey.getState(view.state);
  const sig = JSON.stringify(proxies);
  if (sig !== JSON.stringify(st.proxies)) {
    view.dispatch(view.state.tr.setMeta(floatLayoutKey, { proxies }).setMeta('addToHistory', false));
    changed = true;
  }
  return changed;
}

const proxyDom = (css) => () => {
  const el = document.createElement('span');
  el.className = 'pero-float-proxy';
  el.contentEditable = 'false';
  el.style.cssText = css;
  return el;
};
const proxySet = (doc, list) => DecorationSet.create(doc, list.map((p) => Decoration.widget(p.pos, proxyDom(p.css), { side: -1, key: `fp:${p.css}`, ignoreSelection: true, marks: [] })));

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
      // pending — позиции рисунков, ждущих перепривязки (meta { pending: pos } — поставил пользователь; { done } — всё);
      // proxies — заместители рисунков, стоящих выше строки привязки ({ pos, css }), deco — их виджеты
      state: {
        init: () => ({ pending: new Set(), proxies: [], deco: DecorationSet.empty }),
        apply(tr, st) {
          const meta = tr.getMeta(floatLayoutKey);
          let { pending, proxies, deco } = st;
          if (tr.docChanged) {
            pending = new Set([...pending].map((p) => tr.mapping.map(p)));
            deco = deco.map(tr.mapping, tr.doc);
            proxies = deco.find().map((d) => ({ pos: d.from, css: d.spec.key.slice(3) }));
          }
          if (meta && meta.pending != null) pending = new Set([...pending, meta.pending]);
          if (meta && meta.done != null) pending = new Set();
          if (meta && meta.proxies) { proxies = meta.proxies; deco = proxySet(tr.doc, proxies); }
          return pending === st.pending && proxies === st.proxies && deco === st.deco ? st : { pending, proxies, deco };
        },
      },
      props: { decorations: (state) => floatLayoutKey.getState(state).deco },
      view(view) {
        let timer = 0;
        let rounds = 0; // кругов «рисунки сдвинулись → новый замер» подряд
        const run = () => {
          clearTimeout(timer);
          timer = setTimeout(() => {
            if (view.isDestroyed) return;
            if (typeof matchMedia === 'function' && matchMedia('print').matches) return;
            const { pending } = floatLayoutKey.getState(view.state);
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
