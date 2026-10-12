// Раскладка разделов с колонками на экране (вызывается из разбиения на страницы, pagination.js).
//
// Контейнер columns показывается одним из способов (план):
//   • natural — помещается на остаток листа и за ним «на текущей странице» следующий раздел: CSS-колонки
//     выравниваются сами (column-fill: balance), как у Word перед непрерывным разрывом;
//   • fill    — помещается, но дальше новый лист/конец документа: колонки заполняются по очереди (Word их не
//     выравнивает) — column-fill: auto и высота = остаток листа;
//   • rows    — на несколько листов: ряды колонок высотой с область текста листа (column-height + column-wrap),
//     промежуток между рядами = полоса между листами (там колонтитулы). Начался с середины листа (d > 0) — контейнер
//     сдвинут вверх на d (первый ряд совпадает с листом), а верх каждой колонки первого ряда занят распоркой высотой d.
//     Последний ряд перед «на текущей странице» выравнивается распорками в местах начала колонок, высота контейнера
//     ограничивается, следующий текст идёт сразу под колонками.
// Распорки колонок — виджеты своего плагина (colLayoutKey); при печати их нет (печать — обычные CSS-колонки, листы режет
// Chromium, column-fill по тому же правилу выравнивания).
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { scratchRange } from './domRange.js';

export const colLayoutKey = new PluginKey('colLayout');

function spacerDom(h, block, colBreak) {
  const el = document.createElement(block ? 'div' : 'span');
  el.className = block ? `pero-col-spacer pero-col-spacer-block${colBreak ? ' pero-col-break' : ''}` : 'pero-col-spacer';
  el.contentEditable = 'false';
  el.dataset.peroSpacer = 'col';
  el.style.height = `${Math.max(0, h)}px`;
  return el;
}

// specs: [{ pos, h, block }]
function decoSet(doc, specs) {
  return DecorationSet.create(doc, specs.map((s, i) => Decoration.widget(s.pos, () => spacerDom(s.h, s.block, s.colBreak), {
    side: -1, key: `col${i}:${s.pos}:${Math.round(s.h)}:${s.block ? 'b' : 'l'}${s.colBreak ? 'c' : ''}`, ignoreSelection: true, marks: [], colSpec: s,
  })));
}

export const ColLayout = Extension.create({
  name: 'colLayout',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: colLayoutKey,
      state: {
        init: () => ({ specs: [], set: DecorationSet.empty }),
        apply(tr, st) {
          const meta = tr.getMeta(colLayoutKey);
          if (meta) return { specs: meta, set: decoSet(tr.doc, meta) };
          if (!tr.docChanged || !st.specs.length) return st;
          // распорки, чьё место удалено вместе с текстом (замена документа и т.п.), отпадают вместе с декорациями
          const set = st.set.map(tr.mapping, tr.doc);
          return { specs: set.find().map((d) => ({ ...d.spec.colSpec, pos: d.from })), set };
        },
      },
      props: { decorations: (state) => colLayoutKey.getState(state).set },
    })];
  },
});

const setSpecs = (view, specs) => view.dispatch(view.state.tr.setMeta(colLayoutKey, specs).setMeta('addToHistory', false));
export const colSpecs = (view) => colLayoutKey.getState(view.state).specs;
export const clearColSpacers = (view) => { if (colSpecs(view).length) setSpecs(view, []); };

// Режим контейнера (стили — прямо в DOM: NodeView колонок не перерисовывает узел из-за них)
export function applyMode(dom, plan, balance) {
  const s = dom.style;
  dom.classList.toggle('pc-rows', plan.mode === 'rows');
  if (plan.mode !== 'rows') delete dom.dataset.rows;
  dom.dataset.printFill = balance ? 'balance' : 'auto';
  s.setProperty('column-fill', plan.mode === 'natural' ? 'balance' : 'auto');
  s.setProperty('height', plan.mode === 'fill' || plan.mode === 'balance' ? `${plan.height || plan.avail}px`
    : plan.mode === 'rows' && plan.height ? `${plan.height}px` : '');
  if (plan.mode === 'rows') {
    s.setProperty('--pc-h', `${plan.H}px`);
    s.setProperty('--pc-g', `${plan.G}px`);
    s.setProperty('--pc-mt', `${-plan.d}px`);
  } else {
    for (const k of ['--pc-h', '--pc-g', '--pc-mt']) s.removeProperty(k);
  }
}

// Строки содержимого контейнера в координатах контейнера (px без масштаба): { top, bottom, x, pos, first, el }.
// В колонках порядок строк — не порядок по высоте (верх второй колонки выше низа первой), поэтому строки абзаца
// собираются по колонкам, а начало строки ищется по ключу «ряд → колонка → высота» (он растёт по порядку текста).
// pos — куда ставить распорку: перед абзацем (first) или в начало строки. Строки таблиц — без распорок (pos null).
function containerLines(view, dom, scale, geo) {
  const cr = dom.getBoundingClientRect();
  const rel = (r) => ({ top: (r.top - cr.top) / scale, bottom: (r.bottom - cr.top) / scale, left: (r.left - cr.left) / scale });
  const key = (x) => geo.rowOf(x.top) * 1e7 + geo.colOf(x.left) * 1e5 + x.top;
  const out = [];
  for (const b of dom.querySelectorAll('p, h1, h2, h3, h4, h5, h6, tr')) {
    if (b.tagName === 'TR') {
      if (b.dataset.peroSpacer) continue;
      const r = rel(b.getBoundingClientRect());
      out.push({ top: r.top, bottom: r.bottom, x: r.left, pos: null, first: true, el: b });
      continue;
    }
    if (b.closest('td, th')) continue; // абзацы ячеек — строкой таблицы
    let start = null;
    try { start = view.posAtDOM(b, 0); } catch { start = null; }
    // прямоугольники текста и inline-значков абзаца
    const rects = [];
    const walk = (n) => {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) {
          if (!c.length) continue;
          const rg = scratchRange();
          rg.selectNodeContents(c);
          for (const x of rg.getClientRects()) if (x.width > 0 || x.height > 0) rects.push({ ...rel(x), bottom: (x.bottom - cr.top) / scale, src: c });
        } else if (c.nodeType === 1) {
          if (c.dataset && c.dataset.peroSpacer) continue;
          if (c.classList.contains('pero-ln') || c.classList.contains('ProseMirror-trailingBreak')) continue;
          if (c.contentEditable === 'false' || c.tagName === 'IMG') {
            const x = c.getBoundingClientRect();
            if (x.height > 0) rects.push({ ...rel(x), src: c });
            continue;
          }
          walk(c);
        }
      }
    };
    walk(b);
    if (!rects.length) {
      const r = rel(b.getBoundingClientRect());
      out.push({ top: r.top, bottom: r.bottom, x: r.left, pos: start != null ? start - 1 : null, first: true, el: b });
      continue;
    }
    rects.sort((u, v) => key(u) - key(v));
    const lines = [];
    for (const r of rects) {
      const last = lines[lines.length - 1];
      if (last && geo.colOf(r.left) === last.col && geo.rowOf(r.top) === last.row && r.top < last.bottom - 1) {
        last.bottom = Math.max(last.bottom, r.bottom);
        if (r.left < last.x) last.x = r.left;
        continue;
      }
      lines.push({ top: r.top, bottom: r.bottom, x: r.left, col: geo.colOf(r.left), row: geo.rowOf(r.top), src: r.src });
    }
    // позиция начала строки — по требованию (posOf): поиск дорогой, а нужен лишь для нескольких строк
    lines.forEach((l, i) => out.push({ top: l.top, bottom: l.bottom, x: l.x, first: i === 0, el: b, start, src: l.src, key, cr, scale }));
  }
  return out;
}

// Документная позиция строки: перед абзацем (первая строка) или начало строки; кешируется в объекте строки
function posOf(view, l) {
  if (l.pos !== undefined) return l.pos;
  if (l.first) l.pos = l.start != null ? l.start - 1 : null;
  else l.pos = l.src ? lineStart(view, l, l.key, l.cr, l.scale) : null;
  return l.pos;
}

// Документная позиция первого символа строки: делением пополам по ключу порядка в текстовом узле строки
function lineStart(view, line, key, cr, scale) {
  const src = line.src;
  try {
    if (src.nodeType !== 3) return view.posAtDOM(src.parentNode, Array.prototype.indexOf.call(src.parentNode.childNodes, src));
    const rg = scratchRange();
    const want = key({ top: line.top, left: line.x }) - 1;
    const kAt = (i) => {
      rg.setStart(src, i); rg.setEnd(src, i + 1);
      const x = rg.getClientRects()[0];
      return x ? key({ top: (x.top - cr.top) / scale, left: (x.left - cr.left) / scale }) : -Infinity;
    };
    let lo = 0; let hi = src.length - 1;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (kAt(mid) >= want) hi = mid; else lo = mid + 1; }
    return view.posAtDOM(src, lo);
  } catch { return null; }
}

// Линии между колонками (раздел с разделителем) — слой поверх листа: в каждом ряду от начала текста раздела (первый ряд —
// ниже зарезервированного сверху) до низа самой длинной колонки ряда. Координаты — от .page, px без масштаба.
export function placeColumnRules(view, scale) {
  const page = view.dom.closest('.page');
  if (!page) return;
  let layer = page.querySelector(':scope > .pero-col-rules');
  const conts = [...view.dom.querySelectorAll(':scope > .pero-cols[data-sep]')];
  if (!conts.length) { if (layer) layer.remove(); return; }
  if (!layer) { layer = document.createElement('div'); layer.className = 'pero-col-rules'; page.appendChild(layer); }
  const pr = page.getBoundingClientRect();
  const out = [];
  for (const dom of conts) {
    const cs = getComputedStyle(dom);
    const N = Math.max(1, parseInt(cs.columnCount, 10) || 1);
    if (N < 2) continue;
    const gap = parseFloat(cs.columnGap) || 0;
    const colW = (dom.clientWidth - gap * (N - 1)) / N;
    const cr = dom.getBoundingClientRect();
    const rows = dom.classList.contains('pc-rows');
    const H = rows ? parseFloat(dom.style.getPropertyValue('--pc-h')) || 0 : Infinity;
    const G = rows ? parseFloat(dom.style.getPropertyValue('--pc-g')) || 0 : 0;
    const geo = { rowOf: rows ? (y) => Math.max(0, Math.floor((y + 0.5) / (H + G))) : () => 0, colOf: (x) => colIndexOf(x, colW, gap) };
    const lines = containerLines(view, dom, scale, geo).filter((l) => l.bottom > l.top);
    const R = rows ? (+dom.dataset.rows || 1) : 1;
    for (let r = 0; r < R; r++) {
      const inRow = lines.filter((l) => geo.rowOf(l.top) === r);
      if (!inRow.length) continue;
      const top = Math.min(...inRow.map((l) => l.top));
      const bottom = Math.max(...inRow.map((l) => l.bottom));
      for (let k = 1; k < N; k++) {
        const x = (cr.left - pr.left) / scale + k * colW + (k - 0.5) * gap;
        out.push(`<div class="pero-col-rule" style="left:${x.toFixed(1)}px;top:${((cr.top - pr.top) / scale + top).toFixed(1)}px;height:${(bottom - top).toFixed(1)}px"></div>`);
      }
    }
  }
  const html = out.join('');
  if (layer.innerHTML !== html) layer.innerHTML = html;
}

// Колонка строки по x (ширина колонки и промежуток — из стилей контейнера)
function colIndexOf(x, colW, gap) { return Math.max(0, Math.floor((x + gap / 2) / (colW + gap))); }

// Выровнять колонки одного ряда: минимальная высота, при которой строки ряда по очереди помещаются в N колонок, с
// запретом висячих строк (абзац, разорванный между колонками, — не меньше двух строк с каждой стороны; иначе переносится
// целиком). Разрыв перед абзацем — нулевая распорка с break-before: column, внутри абзаца — распорка до низа ряда.
// Возвращает { height — высота выровненного ряда, place(rowBottom) — поставить распорки, когда низ колонок известен }.
function balanceRow(view, dom, scale, geo, r, rowTop, N, push) {
  let lines = containerLines(view, dom, scale, geo);
  const seq = lines.filter((l) => geo.rowOf(l.top) === r).sort((a, b) => geo.colOf(a.x) - geo.colOf(b.x) || a.top - b.top);
  // строки ряда одной последовательностью: каждая следующая колонка — после конца предыдущей
  const flat = [];
  let base = 0; let curCol = -1; let colEnd = 0; let colBegin = 0;
  for (const l of seq) {
    const c = geo.colOf(l.x);
    if (c !== curCol) { base += colEnd - colBegin; curCol = c; colBegin = l.top - rowTop; colEnd = colBegin; }
    flat.push({ a: base + (l.top - rowTop - colBegin), b: base + (l.bottom - rowTop - colBegin), line: l });
    colEnd = Math.max(colEnd, l.bottom - rowTop);
  }
  const none = (height) => ({ height, place: () => {} });
  if (!flat.length) return none(0);
  const total = flat[flat.length - 1].b;
  if (N <= 1) return none(total);
  // границы абзацев в последовательности (для запрета висячих строк)
  const firstOf = new Map(); const lastOf = new Map();
  flat.forEach((f, i) => { const el = f.line.el; if (!el) return; if (!firstOf.has(el)) firstOf.set(el, i); lastOf.set(el, i); });
  const fits = (h) => {
    const cuts = []; let start = 0; let i = 0;
    while (i < flat.length) {
      if (flat[i].b - start > h + 0.5 && i > (cuts.length ? cuts[cuts.length - 1] : 0)) {
        let c = i;
        const el = flat[c].line.el;
        if (el && !flat[c].line.first && flat[c].line.el.tagName !== 'TR') {
          const before = c - firstOf.get(el); const after = lastOf.get(el) - c + 1;
          if (before < 2) c = firstOf.get(el); // в колонке осталась бы одна строка абзаца — весь абзац дальше
          else if (after < 2) c = before >= 3 ? c - 1 : firstOf.get(el); // одна строка в новой колонке — ещё одну туда же
        }
        const prev = cuts.length ? cuts[cuts.length - 1] : 0;
        if (c <= prev) c = i;
        cuts.push(c);
        if (cuts.length > N - 1) return null;
        start = flat[c].a; i = c + 1;
        continue;
      }
      i++;
    }
    return cuts;
  };
  let lo = Math.max(...flat.map((f) => f.b - f.a)); let hi = Math.max(lo, total);
  if (!fits(hi)) return none(total);
  for (let it = 0; it < 24 && hi - lo > 0.5; it++) { const mid = (lo + hi) / 2; if (fits(mid)) hi = mid; else lo = mid; }
  const cuts = fits(hi) || [];
  // высота ряда — самая высокая колонка при этих разрывах
  let height = 0;
  [0, ...cuts, flat.length].reduce((from, to) => { if (to > from) height = Math.max(height, flat[to - 1].b - flat[from].a); return to; });
  const targets = cuts.map((c) => ({ pos: posOf(view, flat[c].line), first: flat[c].line.first, el: flat[c].line.el }));
  const place = (rowBottom) => { for (let j = 0; j < targets.length; j++) {
    const t = targets[j];
    if (t.pos == null) break;
    lines = containerLines(view, dom, scale, geo);
    // та же строка после предыдущих распорок: в том же абзаце, с той же позицией
    const cur = lines.find((l) => l.el === t.el && l.first === t.first && posOf(view, l) === t.pos) || null;
    if (!cur) break;
    if (geo.colOf(cur.x) > j) continue; // браузер сам начал с этой строки колонку j + 1 — распорка не нужна
    if (cur.first) push({ pos: t.pos, h: 0, block: true, colBreak: true });
    // от верха букв до низа ряда минус запас: строка начинается выше букв на половину интерлиньяжа — распорка
    // точно по остатку не влезла бы и сама ушла в следующую колонку
    else push({ pos: t.pos, h: Math.max(0, rowBottom - cur.top - 4), block: false });
  } };
  return { height: Math.ceil(height), place };
}

// Колонки в пределах листа перед «на текущей странице» (режим balance): заполнение по очереди с высотой остатка листа,
// распорки выравнивания, высота контейнера — выровненная. Возвращает { height }.
export function layoutBalance(view, pos, plan, { scale }) {
  const dom = view.nodeDOM(pos);
  const node = view.state.doc.nodeAt(pos);
  if (!dom || !node) return { height: 0 };
  const N = Math.max(1, parseInt(getComputedStyle(dom).columnCount, 10) || 1);
  const gap = parseFloat(getComputedStyle(dom).columnGap) || 0;
  const colW = (dom.clientWidth - gap * (N - 1)) / N;
  const geo = { rowOf: () => 0, colOf: (x) => colIndexOf(x, colW, gap) };
  const specs = colSpecs(view).filter((x) => x.pos < pos || x.pos > pos + node.nodeSize);
  if (specs.length !== colSpecs(view).length) setSpecs(view, specs);
  dom.style.setProperty('height', `${plan.avail}px`);
  const push = (spec) => { specs.push(spec); setSpecs(view, [...specs]); };
  // высота колонок — выровненная; распорки (поправки разрывов: висячие строки) — относительно неё
  const b = balanceRow(view, dom, scale, geo, 0, 0, N, push);
  const height = Math.max(1, Math.min(plan.avail, b.height + 1));
  dom.style.setProperty('height', `${height}px`);
  b.place(height);
  return { height };
}

// Расставить распорки контейнера в режиме rows: верх колонок первого ряда (d > 0) и выравнивание последнего ряда.
// Каждая распорка ставится по замеру после предыдущей (содержимое сдвигается). Возвращает { R, rowPos, height }.
export function layoutRows(view, pos, plan, balance, { scale }) {
  const dom = view.nodeDOM(pos);
  const node = view.state.doc.nodeAt(pos);
  if (!dom || !node) return { R: 1, rowPos: [] };
  const N = Math.max(1, parseInt(getComputedStyle(dom).columnCount, 10) || 1);
  const gap = parseFloat(getComputedStyle(dom).columnGap) || 0;
  const colW = (dom.clientWidth - gap * (N - 1)) / N;
  const { H, G, d } = plan;
  const rowOf = (y) => Math.max(0, Math.floor((y + 0.5) / (H + G)));
  const geo = { rowOf, colOf: (x) => colIndexOf(x, colW, gap) };
  // свои прежние распорки — долой (план поменялся)
  const specs = colSpecs(view).filter((s) => s.pos < pos || s.pos > pos + node.nodeSize);
  if (specs.length !== colSpecs(view).length) setSpecs(view, specs);
  dom.style.setProperty('height', '');
  const push = (spec) => { specs.push(spec); setSpecs(view, [...specs]); };
  // строка, с которой начинается колонка k ряда r (первая по высоте)
  const firstIn = (lines, r, k) => lines.filter((l) => rowOf(l.top) === r && colIndexOf(l.x, colW, gap) === k).sort((a, b) => a.top - b.top)[0];
  // распорка перед строкой: блок — если строка первая в абзаце/строка таблицы, иначе — в строке
  const spacerAt = (line, h) => {
    const p = posOf(view, line);
    if (p == null) return false;
    push({ pos: p, h, block: !!line.first });
    return true;
  };
  if (d > 0.5) {
    // колонка 1: распорка в начале содержимого
    push({ pos: pos + 1, h: d, block: true });
    for (let k = 1; k < N; k++) {
      const lines = containerLines(view, dom, scale, geo);
      const l = firstIn(lines, 0, k);
      if (!l) break;
      if (l.top > d + 1) continue; // уже ниже зарезервированного (колонка началась с распорки) — не нужно
      spacerAt(l, d);
    }
  }
  plan.height = 0;
  if (balance) {
    // последний ряд перед «на текущей странице» — выровнять, высота контейнера — до низа выровненного ряда
    const lines = containerLines(view, dom, scale, geo);
    const R = Math.max(1, ...lines.map((l) => rowOf(l.top) + 1));
    const rowTop = (R - 1) * (H + G);
    // колонки ряда — высотой с лист (column-height); распорки — до низа ряда, высота контейнера — до выровненного низа
    const b = balanceRow(view, dom, scale, geo, R - 1, rowTop, N, push);
    b.place(rowTop + H);
    if (b.height) plan.height = rowTop + b.height + 1;
  }
  if (plan.height) dom.style.setProperty('height', `${plan.height}px`);
  // ряды и первые позиции рядов (для номера страницы под курсором)
  const lines = containerLines(view, dom, scale, geo);
  const R = Math.max(1, ...lines.map((l) => rowOf(l.top) + 1));
  const rowPos = [];
  for (let r = 0; r < R; r++) {
    const first = lines.filter((l) => rowOf(l.top) === r).sort((a, b) => colIndexOf(a.x, colW, gap) - colIndexOf(b.x, colW, gap) || a.top - b.top)[0];
    const p = first ? posOf(view, first) : null;
    rowPos.push(p != null ? p - pos : 1); // от начала контейнера: при правках выше позиции сдвигаются
  }
  dom.dataset.rows = String(R);
  return { R, rowPos };
}
