// «Непрозрачные» фрагменты OOXML внутри абзаца: то, что Перо пока не редактирует, но обязано сохранить
// (рисунки, ссылки на сноски, комментарии, закладки, поля, формулы, символы, разрывы страниц).
// Узел хранит исходный XML и при сохранении записывается обратно без изменений.
import { Node, mergeAttributes } from '@tiptap/core';
import { mathDom } from './omml.js';
import { shapeHtml, shapeOf } from './shapes.js';
import { transformCss } from './picFormat.js';

// Подписи и вид значков. hidden — служебные метки без видимого представления (закладки, части полей)
export const OPAQUE_KINDS = {
  drawing: { label: 'Рисунок', title: 'Рисунок или фигура — сохраняется без изменений' },
  object: { label: 'Объект', title: 'Внедрённый объект — сохраняется без изменений' },
  footnote: { label: '', title: 'Сноска — текст внизу страницы (щелчок по нему — изменить)' },
  endnote: { label: '', title: 'Ссылка на концевую сноску — сохраняется без изменений' },
  comment: { label: '', title: 'Комментарий — сохраняется без изменений' },
  commentRange: { hidden: true },
  marker: { hidden: true },       // прочие служебные метки (диапазоны разрешений и т.п.)
  bookmark: { hidden: true },
  field: { hidden: true },        // части сложного поля (начало, код, разделитель, конец); результат — обычный текст
  fieldSimple: { label: null, title: 'Поле — сохраняется без изменений' }, // показывает свой текст
  math: { label: 'Формула', title: 'Формула Word — показывается, сохраняется без изменений' }, // MathML (omml.js)
  symbol: { label: null, title: 'Символ' },       // показывает сам символ
  pageBreak: { label: 'Разрыв страницы', title: 'Разрыв страницы' },
  columnBreak: { label: 'Разрыв колонки', title: 'Разрыв колонки' },
  other: { label: '…', title: 'Элемент, который Перо пока не показывает, — сохраняется без изменений' },
};

// Какие виды считаются «сохраняются, но не редактируются» — для строки состояния
export const OPAQUE_SUMMARY = {
  drawing: 'рисунки', object: 'объекты', endnote: 'концевые сноски', comment: 'комментарии',
  math: 'формулы', fieldSimple: 'поля', field: 'поля', header: 'колонтитулы',
  columns: 'колонки разной ширины (на экране — одинаковой)',
};

export const Opaque = Node.create({
  name: 'opaque',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: false,

  addAttributes() {
    return {
      kind: { default: 'other', parseHTML: (el) => el.getAttribute('data-opaque') },
      xml: { default: '', parseHTML: (el) => el.getAttribute('data-xml') || '' },
      text: { default: '', parseHTML: (el) => el.getAttribute('data-text') || '' }, // видимый текст (поле, символ)
      font: { default: null, parseHTML: (el) => el.getAttribute('data-font') },   // шрифт символа (Symbol, Wingdings…)
    };
  },

  parseHTML() {
    return [{ tag: 'span[data-opaque]' }];
  },

  renderHTML({ node, HTMLAttributes }) {
    const { kind, xml, text, font } = node.attrs;
    const k = OPAQUE_KINDS[kind] || OPAQUE_KINDS.other;
    const math = kind === 'math' ? mathDom(xml) : null; // формула — MathML; не разобралась — значок
    const attrs = mergeAttributes(HTMLAttributes, {
      'data-opaque': kind,
      'data-xml': xml,
      'data-text': text || null,
      'data-font': font || null,
      class: `opaque opaque-${kind}${math ? ` math-shown${/^\s*<m:oMathPara\b/.test(xml) ? ' math-block' : ''}` : ''}`,
      title: k.title || null,
      contenteditable: 'false',
      style: font ? `font-family: "${font}"` : null,
    });
    // xml/text — в атрибутах: так значок переживает копирование и вставку внутри документа
    if (math) return ['span', attrs, math];
    // фигура или надпись (wps:wsp) — нарисована; положение и обтекание — floatLayout.js, как у рисунков
    const shape = kind === 'drawing' ? shapeOf(xml) : null;
    if (shape) {
      attrs.class += ` pero-shape${shape.anchor ? ` pero-float pf-${shape.anchor.wrap}` : ''}`;
      attrs.title = null;
      const box = document.createElement('span');
      box.className = 'pero-shape-box';
      const e = shape.effect; // запас wp:effectExtent (у повёрнутой фигуры Word включает в него габарит)
      box.style.cssText = `width:${shape.w.toFixed(1)}px;height:${shape.h.toFixed(1)}px;margin:${[e.t, e.r, e.b, e.l].map((v) => `${v.toFixed(1)}px`).join(' ')}` +
        (shape.rotate ? `;transform:${transformCss({ rotate: shape.rotate })}` : '');
      box.innerHTML = shapeHtml(shape);
      return ['span', attrs, box];
    }
    const label = k.hidden ? '' : k.label === null ? text : k.label;
    return ['span', attrs, label];
  },

  // в обычный текст (копирование, счётчик слов) — только видимый текст поля/символа
  renderText({ node }) {
    return node.attrs.kind === 'fieldSimple' || node.attrs.kind === 'symbol' ? node.attrs.text : '';
  },
});
