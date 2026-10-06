// Разбиение на страницы на экране.
//
// Плагин меряет «единицы» документа (абзацы верхнего уровня, заголовки, таблицы, абзацы в пунктах списков) и
// перед единицей, которая не помещается на текущий лист, ставит распорку-виджет: заполнение до конца листа +
// «полоса» (нижнее поле с нижним колонтитулом, промежуток между листами, верхнее поле следующего листа с верхним
// колонтитулом). Полоса — absolute на всю ширину листа, поэтому распорки работают и внутри списков.
// Ограничения: абзац и таблица не делятся между листами — переносятся целиком (Word делит по строкам).
// Печать/PDF распорки не используют: там страницы режет сам Chromium по @page (см. App).
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { EMPTY_HF, hfText, pagePx, parseSectPr } from './page.js';

export const paginationKey = new PluginKey('pagination');
export const PAGE_GAP = 18; // px между листами на экране

const LIST_TYPES = ['bulletList', 'orderedList'];

// Единицы разбиения: { pos, node, forceAfter }
function collectUnits(doc) {
  const units = [];
  const visit = (node, pos) => {
    if (LIST_TYPES.includes(node.type.name)) {
      node.forEach((item, off) => {
        const itemPos = pos + 1 + off;
        item.forEach((child, off2) => visit(child, itemPos + 1 + off2));
      });
      return;
    }
    let forceAfter = !!node.attrs.sectPr; // разрыв раздела — со следующей страницы
    if (node.isTextblock) node.forEach((c) => { if (c.type.name === 'opaque' && c.attrs.kind === 'pageBreak') forceAfter = true; });
    units.push({ pos, node, forceAfter });
  };
  doc.forEach((node, off) => visit(node, off));
  return units;
}

// Текст колонтитула → HTML-узлы (абзацы с выравниванием)
export function renderHf(container, paras, page, total) {
  container.textContent = '';
  for (const p of paras || []) {
    const div = document.createElement('div');
    div.className = `hf-line hf-${p.align || 'left'}`;
    div.textContent = hfText(p.text, page, total) || ' ';
    container.appendChild(div);
  }
}

function makeSpacer({ fill, page, total, geo, hf }) {
  const el = document.createElement('div');
  el.className = 'pero-page-spacer';
  el.contentEditable = 'false';
  const band = geo.bottom + PAGE_GAP + geo.top;
  el.style.height = `${fill + band}px`;
  // заполнитель до конца листа, за ним — полоса (absolute без top: встаёт на своё место в потоке, сразу после него)
  const filler = document.createElement('div');
  filler.style.height = `${fill}px`;
  const b = document.createElement('div');
  b.className = 'pero-page-band';
  b.style.height = `${band}px`;
  // нижнее поле листа page
  const foot = document.createElement('div');
  foot.className = 'pero-hf pero-footer';
  foot.style.height = `${geo.bottom}px`;
  foot.style.paddingBottom = `${geo.footer}px`;
  renderHf(foot, hf.footer, page, total);
  const gap = document.createElement('div');
  gap.className = 'pero-page-gap';
  gap.style.height = `${PAGE_GAP}px`;
  gap.dataset.label = `Страница ${page + 1}`;
  // верхнее поле листа page + 1
  const head = document.createElement('div');
  head.className = 'pero-hf pero-header';
  head.style.height = `${geo.top}px`;
  head.style.paddingTop = `${geo.header}px`;
  renderHf(head, hf.header, page + 1, total);
  b.append(foot, gap, head);
  el.append(filler, b);
  return el;
}

function measure(view) {
  const { doc } = view.state;
  const geo = pagePx(parseSectPr(doc.attrs.sectPr));
  const contentH = geo.height - geo.top - geo.bottom;
  const hf = doc.attrs.hf || EMPTY_HF;
  const rootRect = view.dom.getBoundingClientRect();
  const scale = view.dom.offsetHeight ? rootRect.height / view.dom.offsetHeight : 1; // CSS zoom страницы
  if (!scale || !isFinite(scale)) return;
  const spacers = Array.from(view.dom.querySelectorAll(':scope .pero-page-spacer')).map((s) => {
    const r = s.getBoundingClientRect();
    return { top: (r.top - rootRect.top) / scale, h: r.height / scale };
  });
  const gapsAbove = (y) => spacers.reduce((sum, s) => (s.top < y - 0.5 ? sum + s.h : sum), 0);

  const units = [];
  for (const u of collectUnits(doc)) {
    const dom = view.nodeDOM(u.pos);
    if (!dom || dom.nodeType !== 1) continue;
    const r = dom.getBoundingClientRect();
    const cs = getComputedStyle(dom);
    const top = (r.top - rootRect.top) / scale;
    const g = gapsAbove(top);
    units.push({
      ...u,
      top: top - (parseFloat(cs.marginTop) || 0) - g,
      bottom: (r.bottom - rootRect.top) / scale + (parseFloat(cs.marginBottom) || 0) - g,
    });
  }

  const breaks = [];
  let pageStart = 0;
  let prevBottom = 0;
  let force = false;
  for (const u of units) {
    const overflow = u.bottom > pageStart + contentH + 0.5 && u.top > pageStart + 0.5;
    if (force || overflow) {
      breaks.push({ pos: u.pos, fill: Math.max(0, pageStart + contentH - prevBottom) });
      pageStart = u.top;
    }
    prevBottom = Math.max(prevBottom, u.bottom);
    // слишком высокий блок (не делим) — следующий начнётся с нового листа
    force = u.forceAfter || u.bottom > pageStart + contentH + 0.5;
  }
  const trailing = Math.max(0, pageStart + contentH - prevBottom);
  const total = breaks.length + 1;
  const key = JSON.stringify([breaks, Math.round(trailing), total, geo, hf]);
  const prev = paginationKey.getState(view.state);
  if (prev && prev.key === key) return;

  const decos = breaks.map((b, i) => Decoration.widget(b.pos, () => makeSpacer({ fill: b.fill, page: i + 1, total, geo, hf }), {
    side: -1, key: `pb${i}:${Math.round(b.fill)}:${total}:${JSON.stringify(hf)}:${JSON.stringify(geo)}`, ignoreSelection: true,
  }));
  decos.push(Decoration.widget(doc.content.size, () => {
    const el = document.createElement('div');
    el.className = 'pero-page-fill';
    el.contentEditable = 'false';
    el.style.height = `${trailing}px`;
    return el;
  }, { side: 1, key: `tail:${Math.round(trailing)}`, ignoreSelection: true }));
  view.dispatch(view.state.tr.setMeta(paginationKey, {
    key, deco: DecorationSet.create(doc, decos), breaks: breaks.map((b) => b.pos), total,
  }).setMeta('addToHistory', false));
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
        const schedule = () => {
          clearTimeout(timer);
          timer = setTimeout(() => { if (!view.isDestroyed) measure(view); }, 40);
        };
        const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(schedule) : null;
        if (ro) ro.observe(view.dom);
        schedule();
        return {
          update(v, prevState) { if (v.state.doc !== prevState.doc) schedule(); },
          destroy() { clearTimeout(timer); if (ro) ro.disconnect(); },
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
