// Нумерация строк («Макет → Номера строк»): w:lnNumType раздела { countBy, start (в файле — с нуля), distance, restart },
// запрет у абзаца — noLineNumbers (paraLayout.js) ↔ w:suppressLineNumbers. Как в Word: номера у строк основного текста,
// кроме таблиц (их строки не считаются); пустой абзац — тоже строка; показываются номера, кратные шагу.
//
// Показ: после разбиения на страницы (событие pero-paginated) строки нумеруемых абзацев меряются, у каждой показываемой
// строки — виджет с номером (numberWidget); одна транзакция только с метаданными — разбиение от неё не пересчитывается
// (оно ждёт изменений документа и размеров). Номера считаются по экранным листам: при печати, если Chromium разобьёт
// страницы иначе, «заново на каждой странице» может начаться на строку раньше/позже.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { pagePx, parseSectPr, sectionList } from './page.js';
import { posAtLine, textLines } from './pagination.js';

const TWIPS_PER_PX = 15;
export const LN_AUTO_DISTANCE = 360; // «авто» — 0,25 дюйма (как Word для одной колонки)

// Печать: Chromium обрезает всё, что выходит за область страницы в поля, — а номера стоят в левом поле. Если нумерация
// есть, левое поле @page уменьшается на pad, а у текста — такой же внутренний отступ: текст на бумаге на месте,
// номера внутри области печати. pad — ширина колонки номеров (не больше самого узкого левого поля).
export function lnPrintPad(sects, mirror = false) {
  const lns = sects.map(parseLnNum);
  if (!lns.some(Boolean)) return 0;
  let pad = 0; let room = Infinity;
  sects.forEach((x, i) => {
    const p = pagePx(parseSectPr(x));
    room = Math.min(room, p.left, mirror ? p.right : Infinity);
    if (lns[i]) pad = Math.max(pad, ((lns[i].distance ?? LN_AUTO_DISTANCE) / TWIPS_PER_PX) + 48); // + место под 5 цифр
  });
  return Math.max(0, Math.floor(Math.min(pad, room - 2)));
}

// Параметры нумерации раздела из XML sectPr; нет — null
export function parseLnNum(sectXml) {
  const tag = (/<w:lnNumType\b[^>]*\/?>/.exec(sectXml || '') || [])[0];
  if (!tag) return null;
  const a = (n) => { const m = new RegExp(`\\bw:${n}="([^"]*)"`).exec(tag); return m ? m[1] : null; };
  const countBy = +a('countBy') || 0;
  if (countBy <= 0) return null; // без шага Word номера не показывает
  return {
    countBy,
    start: (+a('start') || 0) + 1,
    distance: a('distance') != null ? +a('distance') : null,
    restart: a('restart') || 'newPage',
  };
}

// Записать/убрать w:lnNumType в sectPr (по схеме — после pgSz/pgMar/paperSrc/pgBorders, перед pgNumType/cols/…)
export function patchLnNum(sectXml, ln) {
  let s = String(sectXml || '').replace(/<w:lnNumType\b[^>]*\/>/g, '');
  if (!ln) return s;
  const attrs = [`w:countBy="${ln.countBy || 1}"`];
  if ((ln.start || 1) !== 1) attrs.push(`w:start="${(ln.start || 1) - 1}"`);
  if (ln.distance != null) attrs.push(`w:distance="${Math.round(ln.distance)}"`);
  attrs.push(`w:restart="${ln.restart || 'newPage'}"`);
  const tag = `<w:lnNumType ${attrs.join(' ')}/>`;
  const m = /<w:(?:pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/.exec(s);
  if (m) return s.slice(0, m.index) + tag + s.slice(m.index);
  return s.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`);
}

const TEXT = new Set(['paragraph', 'heading']);

// Нумеруемые абзацы: { pos, sect } — абзацы верхнего уровня и в списках, не в таблицах, без запрета
function numberedBlocks(doc) {
  const lns = sectionList(doc).map(parseLnNum);
  if (!lns.some(Boolean)) return { lns, list: [] };
  const list = [];
  let s = 0;
  doc.forEach((top, topPos) => {
    if (lns[Math.min(s, lns.length - 1)]) {
      const visit = (node, pos) => {
        if (TEXT.has(node.type.name)) { if (!node.attrs.noLineNumbers) list.push({ pos, sect: s }); return; }
        if (node.type.name === 'table') return; // строки таблиц не нумеруются
        node.forEach((c, off) => visit(c, pos + 1 + off));
      };
      visit(top, topPos);
    }
    if (top.attrs && top.attrs.sectPr) s++;
  });
  return { lns, list };
}

export const lineNumbersKey = new PluginKey('lineNumbers');

// Номер — пустой виджет внутри слова в середине своей строки, в нём абсолютный элемент: по высоте — в статической
// позиции (там, где строка), по горизонтали — явно от поля (left; при печати своё значение --lxp). Внутри слова нет мест
// переноса — виджет не уходит на соседнюю строку (пустой элемент на границе строк «прилипал» к концу предыдущей); середина
// строки — запас на случай, когда при печати Chromium переносит строку на слово раньше/позже экрана (ширины глифов
// в PDF чуть другие). Цифры — генерируемый текст (::before, attr(data-n)): их не видят орфография, поиск и копирование.
const numberWidget = (it) => Decoration.widget(it.pos, () => {
  const el = document.createElement('span');
  el.className = 'pero-ln';
  el.contentEditable = 'false';
  const n = document.createElement('span');
  n.dataset.n = String(it.n);
  n.dataset.mid = String(it.mid);
  n.dataset.sect = String(it.sect);
  el.appendChild(n);
  return el;
}, { side: -1, key: `ln${it.n}:${it.sect}:${it.mid}`, ignoreSelection: true });

function decorationsOf(doc, items) {
  return items && items.length ? DecorationSet.create(doc, items.map(numberWidget)) : DecorationSet.empty;
}

// Верх листов (px без масштаба, от верха .page) — по полосам разбиения, как вертикальная линейка
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

// Замер строк нумеруемых абзацев (после разбиения на страницы): [{ pos, n, dx, gTop }] — только показываемые номера
export function measureLineNumbers(view) {
  const doc = view.state.doc;
  const { lns, list } = numberedBlocks(doc);
  if (!list.length) return [];
  const root = view.dom;
  const pageEl = root.closest('.page');
  const rr = root.getBoundingClientRect();
  const scale = root.offsetWidth ? rr.width / root.offsetWidth : 1;
  const tops = pageTops(pageEl);
  const pr = pageEl ? pageEl.getBoundingClientRect() : rr;
  const pageOf = (yClient) => {
    const y = (yClient - pr.top) / scale;
    let k = 0;
    while (k + 1 < tops.length && tops[k + 1] <= y + 1) k++;
    return k;
  };
  const marginLeft = rr.left + parseFloat(getComputedStyle(root).paddingLeft) * scale; // левый край текста (без сдвига раздела)
  const out = [];
  let n = 0; let lastPage = -1; let lastSect = -1;
  for (const b of list) {
    const ln = lns[Math.min(b.sect, lns.length - 1)];
    const p = view.nodeDOM(b.pos);
    if (!ln || !p || p.nodeType !== 1) continue;
    const block = p.closest('.page-content > *') || p;
    const secL = (parseFloat(getComputedStyle(block).getPropertyValue('--sec-l')) || 0) * scale;
    const dist = (ln.distance ?? LN_AUTO_DISTANCE) / TWIPS_PER_PX;
    const lines = textLines(p, rr.top, scale);
    const node = doc.nodeAt(b.pos);
    const pr0 = p.getBoundingClientRect();
    const plist = lines.length
      ? lines.map((l) => ({ pos: posAtLine(view, l), gTop: l.gTop, gBottom: l.gBottom }))
      : [{ pos: b.pos + 1, gTop: pr0.top, gBottom: pr0.bottom }];
    const end = b.pos + node.nodeSize - 1; // конец содержимого абзаца
    const isLetter = (p0) => /[^\s\ufffc]/.test(doc.textBetween(p0, p0 + 1, '', ''));
    plist.forEach((l, k) => { l.end = k + 1 < plist.length && plist[k + 1].pos != null ? plist[k + 1].pos : end; });
    for (const l of plist) {
      const page = pageOf(l.gTop);
      if (b.sect !== lastSect) {
        if (ln.restart === 'newSection' || lastSect < 0 || !lns[Math.min(lastSect, lns.length - 1)]) n = ln.start - 1;
        lastSect = b.sect; lastPage = page;
      }
      if (page !== lastPage) { if (ln.restart === 'newPage') n = ln.start - 1; lastPage = page; }
      n++;
      if (n % ln.countBy || l.pos == null) continue;
      const start = Math.max(b.pos + 1, Math.min(end, l.pos));
      // внутри слова у середины строки (между двумя буквами); нет такого места — после первой буквы, иначе начало строки
      let pos = start;
      const midp = Math.floor((start + Math.max(start, l.end)) / 2);
      for (let d = 0; d <= Math.max(0, l.end - start); d++) {
        const c1 = midp + d; const c2 = midp - d;
        if (c1 > start && c1 < l.end && isLetter(c1 - 1) && isLetter(c1)) { pos = c1; break; }
        if (c2 > start && c2 < l.end && isLetter(c2 - 1) && isLetter(c2)) { pos = c2; break; }
      }
      if (pos === start && start < end && isLetter(start)) pos = start + 1;
      out.push({ pos, n, sect: b.sect, mid: Math.round(((l.gTop + l.gBottom) / 2 - rr.top) / scale * 10) / 10 });
    }
  }
  return out;
}

const sameItems = (a, b) => a.length === b.length && a.every((x, i) => x.pos === b[i].pos && x.n === b[i].n && x.sect === b[i].sect && x.mid === b[i].mid);

// Подогнать номера: правый край — у поля минус расстояние (left от блока, относительно которого номер позиционирован;
// если это область текста — при печати у неё внутренний отступ lnPrintPad, поэтому своё значение --lxp), середина —
// на середине строки (margin-top от статической позиции — тот же на экране и при печати).
function alignNumbers(view) {
  const root = view.dom;
  const doc = view.state.doc;
  const lns = sectionList(doc).map(parseLnNum);
  const rr = root.getBoundingClientRect();
  const scale = root.offsetWidth ? rr.width / root.offsetWidth : 1;
  const marginLeft = rr.left + parseFloat(getComputedStyle(root).paddingLeft) * scale;
  const pad = lnPrintPad(sectionList(doc), !!doc.attrs.mirrorMargins);
  for (const s of root.querySelectorAll('.pero-ln > span')) {
    const ln = lns[Math.min(+s.dataset.sect, lns.length - 1)];
    if (!ln) continue;
    const dist = (ln.distance ?? LN_AUTO_DISTANCE) / TWIPS_PER_PX;
    const block = s.closest('.page-content > *');
    const secL = block ? parseFloat(getComputedStyle(block).getPropertyValue('--sec-l')) || 0 : 0;
    const cb = s.offsetParent || root;
    const cr = cb.getBoundingClientRect();
    const cbLeft = cr.left + (parseFloat(getComputedStyle(cb).borderLeftWidth) || 0) * scale; // край блока заполнения
    const lx = Math.round(((marginLeft - cbLeft) / scale + secL - dist) * 10) / 10;
    const lxp = cb === root ? Math.round((pad + secL - dist) * 10) / 10 : lx;
    if (s.style.getPropertyValue('--lx') !== `${lx}px`) s.style.setProperty('--lx', `${lx}px`);
    if (s.style.getPropertyValue('--lxp') !== `${lxp}px`) s.style.setProperty('--lxp', `${lxp}px`);
    const r = s.getBoundingClientRect();
    const mt = parseFloat(s.style.marginTop) || 0;
    const mid = ((r.top + r.bottom) / 2 - rr.top) / scale - mt;
    const nt = Math.round((+s.dataset.mid - mid) * 10) / 10;
    if (Math.abs(nt - mt) > 0.2) s.style.marginTop = `${nt}px`;
  }
}

export const LineNumbers = Extension.create({
  name: 'lineNumbers',
  addCommands() {
    return {
      // нумерация строк раздела под курсором: ln — { countBy, start, distance, restart } | null (выключить)
      setLineNumbering: (ln) => ({ state, tr, dispatch }) => {
        const top = state.selection.$from.index(0);
        let pos = null;
        state.doc.forEach((n, p, i) => { if (pos == null && i >= top && n.attrs && n.attrs.sectPr) pos = p; });
        if (dispatch) {
          if (pos == null) tr.setDocAttribute('sectPr', patchLnNum(state.doc.attrs.sectPr || DEFAULT_SECT, ln));
          else { const n = state.doc.nodeAt(pos); tr.setNodeMarkup(pos, undefined, { ...n.attrs, sectPr: patchLnNum(n.attrs.sectPr, ln) }); }
          dispatch(tr);
        }
        return true;
      },
      // «Запретить в текущем абзаце» — переключатель у абзацев выделения
      toggleNoLineNumbers: () => ({ state, tr, dispatch }) => {
        const { from, to } = state.selection;
        let first = null;
        state.doc.nodesBetween(from, to, (n, pos) => {
          if (!TEXT.has(n.type.name)) return true;
          if (first == null) first = !n.attrs.noLineNumbers;
          tr.setNodeMarkup(pos, undefined, { ...n.attrs, noLineNumbers: first ? true : null });
          return false;
        });
        if (dispatch && first != null) dispatch(tr);
        return first != null;
      },
    };
  },
  addProseMirrorPlugins() {
    return [new Plugin({
      key: lineNumbersKey,
      state: {
        init: () => ({ items: [], set: DecorationSet.empty }),
        apply(tr, st, old, state) {
          const items = tr.getMeta(lineNumbersKey);
          if (items) return { items, set: decorationsOf(state.doc, items) };
          if (!numberedBlocks(state.doc).list.length) return st.items.length ? { items: [], set: DecorationSet.empty } : st;
          if (!tr.docChanged) return st;
          // до нового замера — сдвиг; виджеты потерялись (их место удалено) — список сбрасывается, замер вернёт номера
          const set = st.set.map(tr.mapping, state.doc);
          return { items: set.find().length === st.items.length ? st.items : [], set };
        },
      },
      props: { decorations: (state) => lineNumbersKey.getState(state).set },
      view(view) {
        // таймер, а не requestAnimationFrame: в свёрнутом/скрытом окне rAF не срабатывает (разбиение тоже на таймерах)
        let timer = 0;
        const run = () => {
          clearTimeout(timer);
          timer = setTimeout(() => {
            if (view.isDestroyed) return;
            // при печати (и её эмуляции) полос и распорок нет — номера остаются по экранному разбиению
            if (typeof matchMedia === 'function' && matchMedia('print').matches) return;
            const items = measureLineNumbers(view);
            if (!sameItems(items, lineNumbersKey.getState(view.state).items)) {
              view.dispatch(view.state.tr.setMeta(lineNumbersKey, items).setMeta('addToHistory', false));
            }
            alignNumbers(view);
          }, 20);
        };
        view.dom.addEventListener('pero-paginated', run);
        run();
        return {
          update: (v, prev) => { if (v.state.doc !== prev.doc && !numberedBlocks(v.state.doc).list.length) run(); },
          destroy: () => { clearTimeout(timer); view.dom.removeEventListener('pero-paginated', run); },
        };
      },
    })];
  },
});

const DEFAULT_SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
