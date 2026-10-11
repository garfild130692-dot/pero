// Расстановка переносов («Макет → Расстановка переносов»). У Electron нет словарей переносов Chromium (hyphens: auto
// ничего не делает), поэтому места переноса русских слов считаем сами — по правилам П. Христова — и показываем
// виджетами с мягким переносом (U+00AD) внутри слова: текст документа не меняется, браузер переносит слово по ним
// (и на экране, и при печати). Настройки — атрибуты документа ↔ settings.xml (w:autoHyphenation, w:doNotHyphenateCaps);
// запрет у абзаца — атрибут noHyphens (paraLayout.js) ↔ w:suppressAutoHyphens.
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

const ALL = '[абвгдеёжзийклмнопрстуфхцчшщъыьэюя]';
const V = '[аеёиоуыэюя]';                     // гласные
const C = '[бвгджзклмнпрстфхцчшщ]';           // согласные
const X = '[йъь]';                            // знаки, от которых не отрывают
// правила Христова: перенос между левой и правой частью шаблона
const RULES = [
  [X, `${ALL}${ALL}`],
  [V, `${V}${ALL}`],
  [`${V}${C}`, `${C}${V}`],
  [`${C}${V}`, `${C}${V}`],
  [`${V}${C}`, `${C}${C}${V}`],
  [`${V}${C}${C}`, `${C}${C}${V}`],
  [`${V}${C}`, `${C}${C}${C}${V}`], // «производ-ственный»
].map(([l, r]) => [new RegExp(`${l}$`, 'i'), new RegExp(`^${r}`, 'i')]);
const HAS_V = new RegExp(V, 'i');
const MIN_LEFT = 2;
const MIN_RIGHT = 2;

const cache = new Map();
// Места переноса в слове (индексы, перед которыми можно перенести); только кириллица
export function hyphenPoints(word) {
  const w = word.toLowerCase();
  let r = cache.get(w);
  if (r) return r;
  r = [];
  for (let i = MIN_LEFT; i <= w.length - MIN_RIGHT; i++) {
    const left = w.slice(0, i);
    const right = w.slice(i);
    if (!HAS_V.test(left) || !HAS_V.test(right)) continue; // в каждой части — гласная
    if (/^[йъь]/.test(right)) continue;
    if (/с$/.test(left) && /^тв/.test(right)) continue; // не «водс-твен»: суффикс -ств- не разрываем
    if (RULES.some(([lr, rr]) => lr.test(left) && rr.test(right))) r.push(i);
  }
  // соседние места на расстоянии 1 (одна буква между переносами) — оставляем оба: браузер выберет одно
  if (cache.size > 20000) cache.clear();
  cache.set(w, r);
  return r;
}

// Слово с мягкими переносами — для проверки и для тестов
export const hyphenate = (word, mark = '-') => {
  const pts = hyphenPoints(word);
  let out = ''; let prev = 0;
  for (const p of pts) { out += word.slice(prev, p) + mark; prev = p; }
  return out + word.slice(prev);
};

const WORD = /[а-яё]{4,}/giu;
const UPPER = /^[А-ЯЁ]+$/u;
const TYPES = new Set(['paragraph', 'heading']);

// Позиции мягких переносов в текстовом блоке (pos — начало узла)
function blockPoints(node, pos, caps) {
  if (node.attrs.noHyphens) return [];
  // строка блока: каждая буква и каждый inline-узел — один символ (для inline-содержимого позиции совпадают)
  const text = node.textBetween(0, node.content.size, null, '￼');
  const out = [];
  for (const m of text.matchAll(WORD)) {
    if (!caps && UPPER.test(m[0])) continue;
    for (const p of hyphenPoints(m[0])) out.push(pos + 1 + m.index + p);
  }
  return out;
}

const widget = () => {
  const el = document.createElement('span');
  el.className = 'pero-shy';
  el.textContent = '­';
  return el;
};
const deco = (pos) => Decoration.widget(pos, widget, { side: -1, key: 'shy', ignoreSelection: true });

function allDecos(doc) {
  const caps = doc.attrs.hyphenateCaps !== false;
  const list = [];
  doc.descendants((node, pos) => {
    if (TYPES.has(node.type.name)) { for (const p of blockPoints(node, pos, caps)) list.push(deco(p)); return false; }
    return true;
  });
  return DecorationSet.create(doc, list);
}

export const hyphenKey = new PluginKey('hyphenation');

export const Hyphenation = Extension.create({
  name: 'hyphenation',
  addCommands() {
    return {
      // { auto, caps } — параметры документа (null — не менять)
      setHyphenation: ({ auto = null, caps = null } = {}) => ({ tr, state, dispatch }) => {
        if (dispatch) {
          if (auto != null && auto !== !!state.doc.attrs.autoHyphenation) tr.setDocAttribute('autoHyphenation', auto);
          if (caps != null && caps !== (state.doc.attrs.hyphenateCaps !== false)) tr.setDocAttribute('hyphenateCaps', caps);
          tr.setDocAttribute('hyphenChanged', true);
          dispatch(tr);
        }
        return true;
      },
    };
  },
  addProseMirrorPlugins() {
    return [new Plugin({
      key: hyphenKey,
      state: {
        init: (_, state) => (state.doc.attrs.autoHyphenation ? allDecos(state.doc) : DecorationSet.empty),
        apply(tr, set, old, state) {
          const on = !!state.doc.attrs.autoHyphenation;
          if (!on) return DecorationSet.empty;
          const was = !!old.doc.attrs.autoHyphenation;
          if (!was || (old.doc.attrs.hyphenateCaps !== false) !== (state.doc.attrs.hyphenateCaps !== false)) return allDecos(state.doc);
          if (!tr.docChanged) return set;
          // пересчитать только изменённые текстовые блоки
          let next = set.map(tr.mapping, state.doc);
          const ranges = [];
          tr.steps.forEach((step, i) => {
            const map = tr.mapping.slice(i + 1);
            step.getMap().forEach((_a, _b, from, to) => ranges.push([map.map(from, -1), map.map(to, 1)]));
          });
          const caps = state.doc.attrs.hyphenateCaps !== false;
          const seen = new Set();
          for (const [from, to] of ranges) {
            state.doc.nodesBetween(Math.max(0, from), Math.min(state.doc.content.size, to), (node, pos) => {
              if (!TYPES.has(node.type.name)) return true;
              if (seen.has(pos)) return false;
              seen.add(pos);
              next = next.remove(next.find(pos, pos + node.nodeSize));
              next = next.add(state.doc, blockPoints(node, pos, caps).map(deco));
              return false;
            });
          }
          return next;
        },
      },
      props: { decorations: (state) => hyphenKey.getState(state) },
    })];
  },
});
