// Разбиение на страницы на экране — как в Word: абзацы делятся по строкам, таблицы — по строкам таблицы,
// строка таблицы (если не «не разрывать строку») — посреди ячеек.
//
// Плагин меряет документ и ставит «распорки»-виджеты, которые резервируют высоту остатка листа + полосы между
// листами (нижнее поле, промежуток, верхнее поле следующего листа):
//   • между блоками — div (перед абзацем/таблицей, которые переносятся целиком);
//   • внутри абзаца — inline-block шириной 100% в начале первой строки следующего листа;
//   • в таблице — строка-распорка <tr> перед строкой следующего листа (+ копии строк-заголовков);
//   • в разделённой строке таблицы — такие же inline-распорки в каждой ячейке (шапка над продолжением — в слое полос).
// Сами полосы (колонтитулы, промежуток) рисуются отдельным слоем .pero-bands поверх листа по замеренным
// координатам распорок, поэтому им не мешают позиционированные родители (рамки абзацев, ячейки таблиц).
//
// Замер — по «чистой» раскладке: перед замером распорки снимаются, после — ставятся заново (синхронно, без
// перерисовки между шагами; позиция прокрутки сохраняется). Поэтому координаты точные и не зависят от прошлых распорок.
//
// Правила Word (сверено через COM на tests/fixtures/pages.docx): «запрет висячих строк» (минимум 2 строки абзаца
// на листе), «не разрывать абзац», «не отрывать от следующего» (цепочкой), «с новой страницы», разрыв страницы,
// разрыв раздела; интервал после абзаца может «свисать» за нижнее поле; интервал перед абзацем в начале листа
// (кроме первого) не показывается. Печать/PDF распорки не используют: там страницы режет Chromium по @page (см. App).
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { EMPTY_HF, hfText, pagePx, parseSectPr } from './page.js';

export const paginationKey = new PluginKey('pagination');
export const PAGE_GAP = 18; // px между листами на экране

const LIST_TYPES = ['bulletList', 'orderedList'];
const EPS = 0.5;

// Текст колонтитула → HTML-узлы (абзацы с выравниванием)
export function renderHf(container, paras, page, total) {
  container.textContent = '';
  for (const p of paras || []) {
    const div = document.createElement('div');
    div.className = `hf-line hf-${p.align || 'left'}`;
    div.textContent = hfText(p.text, page, total) || ' ';
    container.appendChild(div);
  }
}

// Масштаб листа на экране — из transform обёртки (PageScaler). Не делить размеры на offsetWidth/offsetHeight:
// они округлены до пикселя, и ошибка копилась к концу длинного документа (~0,4 px на лист).
function pageScale(el) {
  const sc = el.closest('.page-scaled');
  if (!sc) return 1;
  const t = getComputedStyle(sc).transform;
  const m = /^matrix\(([^,]+),/.exec(t || '');
  return m ? parseFloat(m[1]) || 1 : 1;
}

// ---------- Распорки ----------
function blockSpacer(h, fill) {
  const el = document.createElement('div');
  el.className = 'pero-page-spacer';
  el.contentEditable = 'false';
  el.dataset.peroSpacer = 'block';
  el.dataset.fill = String(fill);
  el.style.height = `${Math.max(0, h)}px`;
  return el;
}
function lineSpacer(h, fill, group) {
  const el = document.createElement('span');
  el.className = 'pero-line-spacer';
  el.contentEditable = 'false';
  el.dataset.peroSpacer = group == null ? 'line' : 'cell';
  if (group != null) el.dataset.group = String(group);
  el.dataset.fill = String(fill);
  el.style.height = `${Math.max(0, h)}px`;
  return el;
}
function rowSpacer(h, fill, cols) {
  const tr = document.createElement('tr');
  tr.className = 'pero-row-spacer';
  tr.contentEditable = 'false';
  tr.dataset.peroSpacer = 'row';
  tr.dataset.fill = String(fill);
  const td = document.createElement('td');
  td.colSpan = Math.max(1, cols);
  td.style.height = `${Math.max(0, h)}px`;
  tr.appendChild(td);
  return tr;
}
// Копия строки-заголовка таблицы на новом листе (только вид)
function headerCopy(tr) {
  const c = tr.cloneNode(true);
  c.contentEditable = 'false';
  c.dataset.peroSpacer = 'header';
  c.classList.add('pero-header-copy');
  c.querySelectorAll('[contenteditable]').forEach((x) => x.removeAttribute('contenteditable'));
  c.querySelectorAll('[data-pero-spacer]').forEach((x) => x.remove());
  return c;
}

// ---------- Замер (по чистой раскладке, без распорок) ----------
// Строки текстового блока: [{ top, bottom, x, y }] — top/bottom в координатах от верха области текста,
// x/y — экранная точка в начале строки (для posAtCoords)
function textLines(dom, rootTop, scale) {
  const rects = [];
  const walk = (n) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) {
        if (!c.length) continue;
        const r = document.createRange();
        r.selectNodeContents(c);
        for (const x of r.getClientRects()) if (x.width > 0 || x.height > 0) rects.push({ top: x.top, bottom: x.bottom, left: x.left, height: x.height, src: c });
      } else if (c.nodeType === 1) {
        if (c.dataset && c.dataset.peroSpacer) continue;
        if (c.classList.contains('ProseMirror-trailingBreak')) continue;
        const inline = /^(inline|inline-block|inline-flex)$/.test(getComputedStyle(c).display);
        if (inline && (c.contentEditable === 'false' || c.tagName === 'IMG')) {
          const x = c.getBoundingClientRect();
          if (x.height > 0) rects.push({ top: x.top, bottom: x.bottom, left: x.left, height: x.height, src: c });
          continue;
        }
        walk(c);
      }
    }
  };
  walk(dom);
  rects.sort((a, b) => a.top - b.top || a.left - b.left);
  const lines = [];
  for (const r of rects) {
    const last = lines[lines.length - 1];
    if (last && r.top < last.gBottom - 1) {
      last.gBottom = Math.max(last.gBottom, r.bottom);
      if (r.left < last.x) { last.x = r.left; last.y = r.top + r.height / 2; last.src = r.src; }
      continue;
    }
    lines.push({ gTop: r.top, gBottom: r.bottom, x: r.left, y: r.top + r.height / 2, src: r.src });
  }
  const cs = getComputedStyle(dom);
  const dr = dom.getBoundingClientRect();
  const contentTop = dr.top + (parseFloat(cs.paddingTop) + parseFloat(cs.borderTopWidth || 0)) * scale;
  const contentBottom = dr.bottom - (parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth || 0)) * scale;
  const y = (v) => (v - rootTop) / scale;
  // граница между строками — середина между соседними строками (одинаковая половина интерлиньяжа)
  return lines.map((l, i) => ({
    top: y(i === 0 ? contentTop : (lines[i - 1].gBottom + l.gTop) / 2),
    bottom: y(i === lines.length - 1 ? contentBottom : (l.gBottom + lines[i + 1].gTop) / 2),
    x: l.x,
    y: l.y,
    src: l.src,
    gTop: l.gTop,
    gBottom: l.gBottom,
  }));
}

// Строки всех абзацев ячейки по порядку
function cellLines(td, rootTop, scale) {
  const out = [];
  for (const b of td.querySelectorAll('p, h1, h2, h3, h4, h5, h6')) out.push(...textLines(b, rootTop, scale));
  return out.sort((a, b) => a.top - b.top);
}

// Единицы разбиения в порядке документа
function collectUnits(view, rootTop, scale) {
  const units = [];
  const box = (dom) => {
    const r = dom.getBoundingClientRect();
    const cs = getComputedStyle(dom);
    const mt = parseFloat(cs.marginTop) || 0;
    const top = (r.top - rootTop) / scale - mt;
    const bottom = (r.bottom - rootTop) / scale;
    const textBottom = bottom - (parseFloat(cs.paddingBottom) || 0) - (parseFloat(cs.borderBottomWidth) || 0);
    // sb — интервал перед абзацем: в начале листа (кроме первого) Word его не показывает
    return { top, bottom: bottom + (parseFloat(cs.marginBottom) || 0), textBottom, sb: mt + (parseFloat(cs.paddingTop) || 0) };
  };
  const visit = (node, pos) => {
    if (LIST_TYPES.includes(node.type.name)) {
      node.forEach((item, off) => {
        const itemPos = pos + 1 + off;
        item.forEach((child, off2) => visit(child, itemPos + 1 + off2));
      });
      return;
    }
    const dom = view.nodeDOM(pos);
    if (!dom || dom.nodeType !== 1) return;
    if (node.type.name === 'table') {
      const table = dom.tagName === 'TABLE' ? dom : dom.querySelector('table');
      if (!table) return;
      const trs = Array.from(table.querySelectorAll(':scope > tbody > tr, :scope > thead > tr')).filter((tr) => !tr.dataset.peroSpacer);
      // строки, внутри которых нельзя рвать (продолжение вертикально объединённых ячеек)
      const covered = new Array(trs.length).fill(false);
      trs.forEach((tr, i) => { for (const td of tr.cells) for (let k = 1; k < (td.rowSpan || 1); k++) if (i + k < trs.length) covered[i + k] = true; });
      const rows = [];
      node.forEach((row, off, i) => {
        const tr = trs[i];
        if (!tr) return;
        const r = tr.getBoundingClientRect();
        rows.push({
          pos: pos + 1 + off, node: row, dom: tr, top: (r.top - rootTop) / scale, bottom: (r.bottom - rootTop) / scale,
          header: !!row.attrs.repeatHeader, cantSplit: !!row.attrs.cantSplit || Array.from(tr.cells).some((td) => (td.rowSpan || 1) > 1), covered: covered[i],
        });
      });
      // заголовок повторяется — только непрерывные строки с начала таблицы
      let h = 0; while (h < rows.length && rows[h].header) h++;
      rows.forEach((r, i) => { r.header = i < h; });
      const cols = Math.max(1, ...trs.map((tr) => Array.from(tr.cells).reduce((s, c) => s + (c.colSpan || 1), 0)));
      units.push({ kind: 'table', pos, node, dom, table, ...box(dom), rows, cols, keepNext: false });
      return;
    }
    const text = node.isTextblock;
    const lay = text ? node.attrs : {};
    const heading = node.type.name === 'heading';
    let pageBreakAt = null; // позиция сразу после разрыва страницы внутри абзаца
    let breakAtEnd = false;
    if (text) {
      node.forEach((c, off) => {
        if (c.type.name === 'opaque' && c.attrs.kind === 'pageBreak') {
          if (off + c.nodeSize < node.content.size) pageBreakAt = pos + 1 + off + c.nodeSize; else breakAtEnd = true;
        }
      });
    }
    units.push({
      kind: text ? 'para' : 'block', pos, node, dom, ...box(dom),
      keepNext: text && (lay.keepNext ?? heading),
      keepLines: text && (lay.keepLines ?? heading),
      widow: text && (lay.widowControl ?? true),
      forceBefore: text && !!lay.pageBreakBefore,
      forceAfter: !!(node.attrs && node.attrs.sectPr) || breakAtEnd,
      pageBreakAt,
    });
  };
  view.state.doc.forEach((node, off) => visit(node, off));
  return units;
}

// Начало строки (документная позиция первого символа строки) — по DOM-узлу, с которого строка начинается.
// posAtCoords в таблицах ненадёжен (слои изменения ширины столбцов): точка текста ячейки давала позицию таблицы.
function posAtLine(view, line) {
  const src = line.src;
  try {
    if (src && src.nodeType === 3) {
      // первый символ текстового узла, попавший в эту строку (верх символа монотонно растёт со смещением)
      const r = document.createRange();
      const topOf = (i) => { r.setStart(src, i); r.setEnd(src, i + 1); const x = r.getClientRects()[0]; return x ? x.top : -Infinity; };
      let lo = 0; let hi = src.length - 1;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (topOf(mid) >= line.gTop - 1) hi = mid; else lo = mid + 1; }
      return view.posAtDOM(src, lo);
    }
    if (src && src.parentNode) return view.posAtDOM(src.parentNode, Array.prototype.indexOf.call(src.parentNode.childNodes, src));
  } catch { /* узел вне документа — запасной путь */ }
  const hit = view.posAtCoords({ left: line.x + 1, top: line.y });
  return hit ? hit.pos : null;
}

// Разбиение: список разрывов по единицам (координаты — от верха области текста, px без масштаба)
function paginate(view, units, contentH, rootTop, scale) {
  const breaks = [];
  let pageTop = 0;
  const limit = () => pageTop + contentH;
  const firstOnPage = (u) => u.top <= pageTop + EPS;

  // разрыв перед единицей i с учётом цепочки «не отрывать от следующего»; возвращает индекс, с которого продолжать
  const breakBefore = (i) => {
    let j = i;
    while (j > 0 && units[j - 1].keepNext && units[j - 1].kind !== 'table' && units[j - 1].top > pageTop + EPS) j--;
    if (j !== i && firstOnPage(units[j])) j = i; // вся цепочка с начала листа — рвём перед исходной единицей
    const u = units[j];
    const before = j > 0 ? units[j - 1].bottom : 0;
    breaks.push({ kind: 'block', pos: u.pos, fill: limit() - before, sup: u.sb || 0 });
    pageTop = u.top + (u.sb || 0);
    return j;
  };
  const afterForced = (i) => {
    const n = units[i + 1];
    if (!n) return;
    breaks.push({ kind: 'block', pos: n.pos, fill: limit() - units[i].bottom, sup: n.sb || 0 });
    pageTop = n.top + (n.sb || 0);
  };

  for (let i = 0; i < units.length; i++) {
    let u = units[i];
    if (u.forceBefore && u.top > pageTop + EPS) { i = breakBefore(i); u = units[i]; }

    if (u.kind === 'table') {
      if (u.bottom <= limit() + EPS) continue;
      const headers = u.rows.filter((r) => r.header);
      const headerH = headers.length ? headers[headers.length - 1].bottom - headers[0].top : 0;
      let rewind = -1;
      for (let k = 0; k < u.rows.length; k++) {
        const r = u.rows[k];
        if (r.bottom <= limit() + EPS) continue;
        const atTop = r.top <= pageTop + EPS;
        // строку можно делить посреди ячеек (как Word по умолчанию)
        if (!r.cantSplit && !r.header) {
          const cells = Array.from(r.dom.cells).map((td) => ({ td, lines: cellLines(td, rootTop, scale) }));
          const splits = cells.map((c) => ({ ...c, k: c.lines.findIndex((l) => l.bottom > limit() + EPS) })).filter((c) => c.k >= 0);
          const someFit = splits.length < cells.length || splits.some((c) => c.k > 0);
          if (splits.length && (someFit || atTop)) {
            const items = [];
            let minTop = Infinity;
            const rowContentBottom = Math.max(...cells.map((c) => (c.lines.length ? c.lines[c.lines.length - 1].bottom : r.top)));
            for (const c of splits) {
              const line = c.lines[c.k > 0 || !atTop ? c.k : Math.min(1, c.lines.length - 1)];
              const pos = posAtLine(view, line);
              if (pos == null) continue;
              items.push({ pos, fill: limit() - line.top });
              const cellBottom = c.lines[c.lines.length - 1].bottom;
              minTop = Math.min(minTop, line.top + (rowContentBottom - cellBottom));
            }
            if (items.length) {
              breaks.push({ kind: 'cells', items, headerH: headers.length ? headerH : 0, table: u.pos, headers: headers.length > 0 });
              pageTop = minTop - (headers.length ? headerH : 0);
              continue; // строка продолжается на новом листе; дальнейшие строки — по новому листу
            }
          }
        }
        // целиком на новый лист: ищем строку, перед которой можно рвать (не внутри объединения)
        let b = k;
        while (b > 0 && u.rows[b].covered) b--;
        if (u.rows[b].top <= pageTop + EPS) continue; // строка выше листа — переполнение, ничего не поделать
        if (b === 0 || u.rows[b].header) {
          // таблица целиком на следующий лист
          if (firstOnPage(u)) continue;
          const back = breakBefore(i);
          if (back !== i) { rewind = back; break; }
          k = -1; // заново по строкам на новом листе
          continue;
        }
        const row = u.rows[b];
        const repeat = headers.length && !row.header ? headers : [];
        breaks.push({ kind: 'row', pos: row.pos, fill: limit() - row.top, cols: u.cols, headers: repeat.map((x) => x.dom) });
        pageTop = row.top - (repeat.length ? headerH : 0);
        k = b;
      }
      if (rewind >= 0) i = rewind - 1;
      continue;
    }

    if (u.kind === 'block') {
      if (u.bottom > limit() + EPS && !firstOnPage(u)) i = breakBefore(i);
      if (units[i].forceAfter) afterForced(i);
      continue;
    }

    // абзац
    const explicitBreak = u.pageBreakAt != null;
    if (u.textBottom <= limit() + EPS && !explicitBreak) {
      if (u.forceAfter) afterForced(i);
      continue;
    }
    const lines = textLines(u.dom, rootTop, scale);
    if (!lines.length) continue;
    // разрыв страницы внутри абзаца: текст после него — со следующего листа
    let forcedLine = -1;
    if (explicitBreak) {
      try {
        const c = view.coordsAtPos(u.pageBreakAt, 1);
        const yy = (c.top - rootTop) / scale;
        forcedLine = lines.findIndex((l) => l.bottom > yy + 1);
      } catch { forcedLine = -1; }
    }
    let first = 0; // первая строка абзаца на текущем листе
    let rewound = false;
    for (;;) {
      let k = lines.findIndex((l, idx) => idx >= first && l.bottom > limit() + EPS);
      if (forcedLine > first && (k < 0 || forcedLine < k)) {
        const pos = posAtLine(view, lines[forcedLine]);
        if (pos != null) breaks.push({ kind: 'line', pos, fill: limit() - lines[forcedLine].top });
        pageTop = lines[forcedLine].top;
        first = forcedLine;
        forcedLine = -1;
        continue;
      }
      if (k < 0) break;
      const n = lines.length;
      const fitsWhole = lines[n - 1].bottom - lines[0].top <= contentH;
      if (first === 0 && (k === 0 || (u.keepLines && fitsWhole) || (u.widow && k < 2 && n >= 2))) {
        // абзац целиком на следующий лист (если он не первый на листе)
        if (firstOnPage(u)) { if (k === 0) k = 1; } else {
          const back = breakBefore(i);
          if (back !== i) { i = back - 1; rewound = true; break; }
          first = 0;
          continue;
        }
      }
      if (u.widow && n - k < 2 && k - first > 2) k = n - 2; // висячая строка внизу: переносим ещё одну
      if (k <= first) k = first + 1;
      if (k >= n) break;
      const pos = posAtLine(view, lines[k]);
      if (pos != null) breaks.push({ kind: 'line', pos, fill: limit() - lines[k].top });
      pageTop = lines[k].top;
      first = k;
    }
    if (rewound) continue;
    if (u.forceAfter) afterForced(i);
  }
  const lastBottom = units.length ? units[units.length - 1].bottom : 0;
  return { breaks, trailing: Math.max(0, pageTop + contentH - lastBottom) };
}

function measure(view) {
  const { doc } = view.state;
  const geo = pagePx(parseSectPr(doc.attrs.sectPr));
  const contentH = geo.height - geo.top - geo.bottom;
  const band = geo.bottom + PAGE_GAP + geo.top;
  const hf = doc.attrs.hf || EMPTY_HF;
  const prev = paginationKey.getState(view.state);

  // 1. снять распорки (прокрутку запоминаем: документ временно короче)
  const scroller = view.dom.closest('.workspace') || document.scrollingElement;
  const scrollTop = scroller ? scroller.scrollTop : 0;
  if (prev.deco !== DecorationSet.empty) view.dispatch(view.state.tr.setMeta(paginationKey, { ...prev, deco: DecorationSet.empty }).setMeta('addToHistory', false));

  // 2. замер чистой раскладки
  const rootRect = view.dom.getBoundingClientRect();
  const scale = pageScale(view.dom); // масштаб страницы (transform)
  let result = { breaks: [], trailing: 0 };
  if (scale && isFinite(scale)) result = paginate(view, collectUnits(view, rootRect.top, scale), contentH, rootRect.top, scale);
  const { breaks, trailing } = result;
  const total = breaks.length + 1;
  const sig = breaks.map((b) => `${b.kind}:${b.pos}:${Math.round(b.fill)}:${(b.items || []).map((x) => `${x.pos}/${Math.round(x.fill)}`).join(',')}:${(b.headers ? 1 : 0)}`).join('|');
  const key = JSON.stringify([sig, Math.round(trailing), total, geo, hf]);

  // 3. поставить распорки (те же — если ничего не изменилось)
  let deco = prev.deco;
  if (prev.key !== key || prev.deco === DecorationSet.empty) {
    const decos = [];
    breaks.forEach((b, i) => {
      const k = `${b.kind}${i}:${Math.round(b.fill)}`;
      if (b.kind === 'block') decos.push(Decoration.widget(b.pos, () => blockSpacer(b.fill + band - (b.sup || 0), b.fill), { side: -1, key: k, ignoreSelection: true }));
      else if (b.kind === 'line') decos.push(Decoration.widget(b.pos, () => lineSpacer(b.fill + band, b.fill), { side: -1, key: k, ignoreSelection: true, marks: [] }));
      else if (b.kind === 'cells') {
        b.items.forEach((it, j) => decos.push(Decoration.widget(it.pos, () => lineSpacer(it.fill + band + b.headerH, it.fill, b.headers ? `h${i}` : `c${i}`), {
          side: -1, key: `${k}c${j}:${Math.round(it.fill)}:${Math.round(b.headerH)}`, ignoreSelection: true, marks: [],
        })));
      } else {
        decos.push(Decoration.widget(b.pos, () => rowSpacer(b.fill + band, b.fill, b.cols), { side: -2, key: k, ignoreSelection: true }));
        (b.headers || []).forEach((tr, j) => decos.push(Decoration.widget(b.pos, () => headerCopy(tr), { side: -1, key: `${k}h${j}:${tr.textContent.length}`, ignoreSelection: true })));
      }
    });
    decos.push(Decoration.widget(doc.content.size, () => {
      const el = document.createElement('div');
      el.className = 'pero-page-fill';
      el.contentEditable = 'false';
      el.style.height = `${trailing}px`;
      return el;
    }, { side: 1, key: `tail:${Math.round(trailing)}`, ignoreSelection: true }));
    deco = DecorationSet.create(view.state.doc, decos);
  }
  view.dispatch(view.state.tr.setMeta(paginationKey, { key, deco, breaks: pageStarts(breaks), total }).setMeta('addToHistory', false));
  if (scroller && scroller.scrollTop !== scrollTop) scroller.scrollTop = scrollTop;
  placeBands(view, geo, hf, total);
}

// Полосы стоят там, где сейчас распорки? (раскладка могла сдвинуться позже замера: шрифты, стили документа)
function bandsAligned(view) {
  const page = view.dom.closest('.page');
  const layer = page && page.querySelector(':scope > .pero-bands');
  if (!layer) return true;
  const pr = page.getBoundingClientRect();
  const scale = pageScale(page);
  const bands = Array.from(layer.querySelectorAll(':scope > .pero-page-band')).map((b) => b.offsetTop);
  const seen = new Set();
  const tops = [];
  for (const s of view.dom.querySelectorAll('[data-pero-spacer="block"], [data-pero-spacer="line"], [data-pero-spacer="row"], [data-pero-spacer="cell"]')) {
    if (s.dataset.group) { if (seen.has(s.dataset.group)) continue; seen.add(s.dataset.group); }
    tops.push((s.getBoundingClientRect().top - pr.top) / scale + (+s.dataset.fill || 0));
  }
  tops.sort((a, b) => a - b);
  return tops.length === bands.length && tops.every((t, i) => Math.abs(t - bands[i]) < 1);
}

// Позиции начала листов (для номера страницы под курсором)
const pageStarts = (breaks) => breaks.map((b) => (b.kind === 'cells' ? Math.min(...b.items.map((x) => x.pos)) : b.pos));

// ---------- Слой полос между листами ----------
function placeBands(view, geo, hf, total) {
  const page = view.dom.closest('.page');
  if (!page) return;
  let layer = page.querySelector(':scope > .pero-bands');
  if (!layer) {
    layer = document.createElement('div');
    layer.className = 'pero-bands';
    page.appendChild(layer);
  }
  const pr = page.getBoundingClientRect();
  const scale = pageScale(page);
  // одна полоса на разрыв: распорки ячеек одной разделённой строки дают общую полосу
  const seen = new Set();
  const spacers = [];
  for (const s of view.dom.querySelectorAll('[data-pero-spacer="block"], [data-pero-spacer="line"], [data-pero-spacer="row"], [data-pero-spacer="cell"]')) {
    const top = (s.getBoundingClientRect().top - pr.top) / scale + (+s.dataset.fill || 0);
    if (s.dataset.group) {
      if (seen.has(s.dataset.group)) continue;
      seen.add(s.dataset.group);
    }
    spacers.push({ top, header: s.dataset.group && s.dataset.group.startsWith('h') ? s : null, cell: s.dataset.peroSpacer === 'cell' ? s : null });
  }
  spacers.sort((a, b) => a.top - b.top);
  const bandH = geo.bottom + PAGE_GAP + geo.top;
  const sig = JSON.stringify([spacers.map((s) => Math.round(s.top)), geo, hf, total]);
  if (layer.dataset.sig === sig) return;
  layer.dataset.sig = sig;
  layer.textContent = '';
  spacers.forEach((s, i) => {
    const b = document.createElement('div');
    b.className = 'pero-page-band';
    b.style.top = `${s.top}px`;
    b.style.height = `${bandH}px`;
    const foot = document.createElement('div');
    foot.className = 'pero-hf pero-footer';
    foot.style.height = `${geo.bottom}px`;
    foot.style.paddingBottom = `${geo.footer}px`;
    renderHf(foot, hf.footer, i + 1, total);
    const gap = document.createElement('div');
    gap.className = 'pero-page-gap';
    gap.style.height = `${PAGE_GAP}px`;
    gap.dataset.label = `Страница ${i + 2}`;
    const head = document.createElement('div');
    head.className = 'pero-hf pero-header';
    head.style.height = `${geo.top}px`;
    head.style.paddingTop = `${geo.header}px`;
    renderHf(head, hf.header, i + 2, total);
    b.append(foot, gap, head);
    layer.appendChild(b);
    // разделённая строка таблицы: линия под её первой частью и над продолжением (как рисует Word)
    if (s.cell) {
      const table = s.cell.closest('table');
      if (table) {
        const tr = table.getBoundingClientRect();
        const style = table.style.getPropertyValue('--tb-ih') || '0.5pt solid #000';
        const line = (y) => {
          const l = document.createElement('div');
          l.className = 'pero-split-line';
          l.style.cssText = `top: ${y}px; left: ${(tr.left - pr.left) / scale}px; width: ${tr.width / scale}px; border-top: ${style}`;
          layer.appendChild(l);
        };
        line(s.top);
        if (!s.header) line(s.top + bandH);
      }
    }
    // шапка таблицы над продолжением разделённой строки
    if (s.header) {
      const table = s.header.closest('table');
      if (table) {
        const tr = table.getBoundingClientRect();
        const copy = table.cloneNode(false);
        const cg = table.querySelector(':scope > colgroup');
        if (cg) copy.appendChild(cg.cloneNode(true));
        const body = document.createElement('tbody');
        // только настоящие строки-заголовки (не их копии от других разрывов)
        for (const row of table.querySelectorAll(':scope > tbody > tr[data-repeat-header="true"]:not([data-pero-spacer])')) body.appendChild(headerCopy(row));
        copy.appendChild(body);
        // обёртка с классом области текста — у копии те же стили, что у таблицы в документе (поля, линии, шрифт)
        const wrap = document.createElement('div');
        wrap.className = 'page-content pero-header-overlay';
        wrap.style.left = `${(tr.left - pr.left) / scale}px`;
        wrap.style.width = `${tr.width / scale}px`;
        wrap.style.top = `${s.top + bandH}px`;
        wrap.appendChild(copy);
        layer.appendChild(wrap);
      }
    }
  });
}

export const Pagination = Extension.create({
  name: 'pagination',

  addProseMirrorPlugins() {
    return [new Plugin({
      key: paginationKey,
      state: {
        init: () => ({ key: null, deco: DecorationSet.empty, breaks: [], total: 1 }),
        apply(tr, old) {
          const meta = tr.getMeta(paginationKey);
          if (meta) return meta;
          if (!tr.docChanged) return old;
          return { ...old, deco: old.deco.map(tr.mapping, tr.doc), breaks: old.breaks.map((p) => tr.mapping.map(p)) };
        },
      },
      props: {
        decorations(state) { return paginationKey.getState(state).deco; },
      },
      view(view) {
        if (typeof window === 'undefined') return {};
        // setTimeout, а не requestAnimationFrame: rAF не срабатывает, пока окно скрыто/свёрнуто,
        // а задержка заодно склеивает частые изменения при наборе текста
        let timer = 0;
        let verify = 0;
        let measuring = false;
        const schedule = () => {
          if (measuring) return; // собственные изменения размера во время замера
          clearTimeout(timer);
          timer = setTimeout(() => {
            if (view.isDestroyed) return;
            measuring = true;
            const t0 = performance.now();
            try { measure(view); } finally { measuring = false; }
            if (import.meta.env && import.meta.env.DEV) window.__peroMeasureMs = Math.round(performance.now() - t0); // отладка: время замера
            // самопроверка: раскладка могла сдвинуться позже (шрифты, коэффициент строки документа) — замерить заново
            clearTimeout(verify);
            verify = setTimeout(() => { if (!view.isDestroyed && !bandsAligned(view)) schedule(); }, 300);
          }, 40);
        };
        const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(schedule) : null;
        if (ro) ro.observe(view.dom);
        if (document.fonts) document.fonts.addEventListener('loadingdone', schedule);
        schedule();
        return {
          update(v, prevState) { if (v.state.doc !== prevState.doc) schedule(); },
          destroy() {
            clearTimeout(timer); clearTimeout(verify);
            if (ro) ro.disconnect();
            if (document.fonts) document.fonts.removeEventListener('loadingdone', schedule);
          },
        };
      },
    })];
  },
});

// Номер страницы под курсором и всего страниц (для строки состояния)
export function pageInfo(state) {
  const p = paginationKey.getState(state);
  if (!p) return { page: 1, total: 1 };
  const at = state.selection.from;
  return { page: 1 + p.breaks.filter((b) => b <= at).length, total: p.total };
}
