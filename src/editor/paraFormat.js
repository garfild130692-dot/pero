// Форматирование абзаца: интервалы, межстрочный интервал, отступы. Атрибуты узлов paragraph и heading.
import { Extension } from '@tiptap/core';
import { INDENT_STEP, PARA_KEYS, normalizeLine, paraBase } from './paragraph.js';

const TYPES = ['paragraph', 'heading'];

// CSS-длина → pt (для вставки из буфера: Word кладёт в HTML pt/cm, браузеры — px)
const UNIT = { pt: 1, px: 0.75, cm: 72 / 2.54, mm: 72 / 25.4, in: 72 };
function cssPt(v) {
  const m = /^\s*(-?[\d.]+)\s*(pt|px|cm|mm|in)\s*$/i.exec(v || '');
  return m ? Math.round(parseFloat(m[1]) * UNIT[m[2].toLowerCase()] * 20) / 20 : null;
}

// Межстрочный интервал в CSS. «Одинарный» в Word = естественная высота строки шрифта; в CSS её даёт
// line-height: normal, но множить normal нельзя — поэтому коэффициент шрифта (--line-ratio) меряет App.
export function lineCss(rule, value) {
  if (rule === 'exact') return `${value}pt`;
  if (rule === 'atLeast') return `max(${value}pt, calc(var(--line-ratio, 1.15) * 1em))`;
  return `calc(var(--line-ratio, 1.15) * ${value})`;
}

// Действующие параметры абзаца: прямое форматирование → (заголовок) → умолчания документа
export function effectivePara(node, docAttrs, inList = false) {
  const base = paraBase(docAttrs && docAttrs.para, { headingLevel: node.type.name === 'heading' ? node.attrs.level : 0, inList });
  // стиль абзаца — между умолчаниями документа и прямым форматированием
  const st = node.attrs.styleId && docAttrs && (docAttrs.styles || []).find((s) => s.id === node.attrs.styleId);
  if (st && !inList) for (const [k, v] of Object.entries(st.para || {})) if (v != null) base[k] = v;
  const out = {};
  for (const k of PARA_KEYS) out[k] = node.attrs[k] ?? base[k];
  return out;
}

const inListAt = ($pos) => {
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === 'listItem') return true;
  return false;
};

// Обойти абзацы/заголовки выделения: fn(node, pos, inList) → новые attrs или null
function eachPara(state, tr, fn) {
  const { from, to } = tr.selection;
  let changed = false;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!TYPES.includes(node.type.name)) return true;
    const next = fn(node, pos, inListAt(state.doc.resolve(pos + 1)));
    if (next) { tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next }); changed = true; }
    return false;
  });
  return changed;
}

export const ParaFormat = Extension.create({
  name: 'paraFormat',

  addGlobalAttributes() {
    const attr = (parse, css) => ({
      default: null,
      keepOnSplit: true, // Enter в абзаце с отступом — новый абзац с тем же отступом, как в Word
      parseHTML: parse,
      renderHTML: (a) => css(a),
    });
    return [{
      types: TYPES,
      attributes: {
        spaceBefore: attr((el) => cssPt(el.style.marginTop), (a) => (a.spaceBefore != null ? { style: `margin-top: ${a.spaceBefore}pt` } : {}) /* margin: Word берёт больший из соседних интервалов */),
        spaceAfter: attr((el) => cssPt(el.style.marginBottom), (a) => (a.spaceAfter != null ? { style: `margin-bottom: ${a.spaceAfter}pt` } : {})),
        lineRule: attr((el) => {
          const v = el.style.lineHeight;
          if (!v) return null;
          return /%$/.test(v) || /^[\d.]+$/.test(v) ? 'auto' : cssPt(v) != null ? 'exact' : null;
        }, () => ({})), // рисуется вместе с lineSpacing
        lineSpacing: attr((el) => {
          const v = el.style.lineHeight;
          if (!v) return null;
          if (/%$/.test(v)) return Math.round(parseFloat(v) / 100 * 240) / 240;
          if (/^[\d.]+$/.test(v)) return Math.round(parseFloat(v) * 240) / 240;
          return cssPt(v);
        }, (a) => (a.lineSpacing != null ? { style: `line-height: ${lineCss(a.lineRule || 'auto', a.lineSpacing)}` } : {})),
        indentLeft: attr((el) => cssPt(el.style.marginLeft), (a) => (a.indentLeft != null ? { style: `margin-left: calc(var(--sec-l) + ${a.indentLeft}pt)` } : {}) /* + сдвиг раздела */),
        indentRight: attr((el) => cssPt(el.style.marginRight), (a) => (a.indentRight != null ? { style: `margin-right: calc(var(--sec-r) + ${a.indentRight}pt)` } : {})),
        indentFirst: attr((el) => cssPt(el.style.textIndent), (a) => (a.indentFirst != null ? { style: `text-indent: ${a.indentFirst}pt` } : {})),
        // Разрыв раздела после этого абзаца: исходный XML w:sectPr (страница, поля, колонтитулы раздела).
        // Не переносится при Enter — иначе раздел «размножился» бы.
        sectPr: {
          default: null,
          keepOnSplit: false,
          parseHTML: () => null,
          renderHTML: (a) => (a.sectPr ? { 'data-section-break': 'true' } : {}), // подпись — плагин sectionBreakLabels (тип следующего раздела)
        },
      },
    }];
  },

  addCommands() {
    return {
      // Применить параметры ко всем абзацам выделения; null — вернуть к умолчанию
      setParagraphFormat: (attrs) => ({ state, tr, dispatch }) => {
        const patch = { ...attrs };
        if ('lineSpacing' in patch && patch.lineSpacing != null) {
          const n = normalizeLine(patch.lineRule || 'auto', patch.lineSpacing);
          if (!n) return false;
          Object.assign(patch, n);
        }
        const changed = eachPara(state, tr, () => patch);
        if (dispatch && changed) dispatch(tr);
        return changed;
      },

      setLineSpacing: (multiple) => ({ commands }) => commands.setParagraphFormat({ lineRule: 'auto', lineSpacing: multiple }),

      // «Увеличить/уменьшить отступ»: в списке — уровень списка, иначе левый отступ шагом 1,25 см (с привязкой к шагу)
      changeIndent: (dir) => ({ state, tr, dispatch, commands }) => {
        if (inListAt(state.selection.$from)) {
          return dir > 0 ? commands.sinkListItem('listItem') : commands.liftListItem('listItem');
        }
        const changed = eachPara(state, tr, (node, pos, inList) => {
          if (inList) return null;
          const cur = effectivePara(node, state.doc.attrs).indentLeft;
          const steps = dir > 0 ? Math.floor(cur / INDENT_STEP + 1e-6) + 1 : Math.ceil(cur / INDENT_STEP - 1e-6) - 1;
          const next = Math.max(0, Math.round(steps * INDENT_STEP * 20) / 20);
          return next === cur ? null : { indentLeft: next };
        });
        if (dispatch && changed) dispatch(tr);
        return true;
      },

      // Ctrl+0: добавить/убрать интервал 12 pt перед абзацем (по первому абзацу выделения)
      toggleSpaceBefore: () => ({ state, commands }) => {
        const node = state.selection.$from.parent;
        const cur = TYPES.includes(node.type.name) ? effectivePara(node, state.doc.attrs).spaceBefore : 0;
        return commands.setParagraphFormat({ spaceBefore: cur > 0 ? 0 : 12 });
      },
      toggleSpaceAfter: () => ({ state, commands }) => {
        const node = state.selection.$from.parent;
        const cur = TYPES.includes(node.type.name) ? effectivePara(node, state.doc.attrs).spaceAfter : 0;
        const def = ((state.doc.attrs.para || {}).spaceAfter) || 8;
        return commands.setParagraphFormat({ spaceAfter: cur > 0 ? 0 : def });
      },
    };
  },

  addKeyboardShortcuts() {
    const c = () => this.editor.commands;
    return {
      'Mod-1': () => c().setLineSpacing(1),
      'Mod-2': () => c().setLineSpacing(2),
      'Mod-5': () => c().setLineSpacing(1.5),
      'Mod-0': () => c().toggleSpaceBefore(),
      'Mod-m': () => c().changeIndent(1),
      'Mod-Shift-m': () => c().changeIndent(-1),
    };
  },
});
