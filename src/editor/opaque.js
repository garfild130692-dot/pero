// «Непрозрачные» фрагменты OOXML внутри абзаца: то, что Перо пока не редактирует, но обязано сохранить
// (рисунки, ссылки на сноски, комментарии, закладки, поля, формулы, символы, разрывы страниц).
// Узел хранит исходный XML и при сохранении записывается обратно без изменений.
import { Node, mergeAttributes } from '@tiptap/core';

// Подписи и вид значков. hidden — служебные метки без видимого представления (закладки, части полей)
export const OPAQUE_KINDS = {
  drawing: { label: 'Рисунок', title: 'Рисунок или фигура — сохраняется без изменений' },
  object: { label: 'Объект', title: 'Внедрённый объект — сохраняется без изменений' },
  footnote: { label: '', title: 'Ссылка на сноску — текст сноски сохраняется без изменений' },
  endnote: { label: '', title: 'Ссылка на концевую сноску — сохраняется без изменений' },
  comment: { label: '', title: 'Комментарий — сохраняется без изменений' },
  commentRange: { hidden: true },
  marker: { hidden: true },       // прочие служебные метки (диапазоны разрешений и т.п.)
  bookmark: { hidden: true },
  field: { hidden: true },        // части сложного поля (начало, код, разделитель, конец); результат — обычный текст
  fieldSimple: { label: null, title: 'Поле — сохраняется без изменений' }, // показывает свой текст
  math: { label: 'Формула', title: 'Формула — сохраняется без изменений' },
  symbol: { label: null, title: 'Символ' },       // показывает сам символ
  pageBreak: { label: 'Разрыв страницы', title: 'Разрыв страницы' },
  columnBreak: { label: 'Разрыв колонки', title: 'Разрыв колонки' },
  other: { label: '…', title: 'Элемент, который Перо пока не показывает, — сохраняется без изменений' },
};

// Какие виды считаются «сохраняются, но не редактируются» — для строки состояния
export const OPAQUE_SUMMARY = {
  drawing: 'рисунки', object: 'объекты', footnote: 'сноски', endnote: 'концевые сноски', comment: 'комментарии',
  math: 'формулы', fieldSimple: 'поля', field: 'поля', header: 'колонтитулы',
  columns: 'колонки текста (на экране — одной колонкой)',
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
    const attrs = mergeAttributes(HTMLAttributes, {
      'data-opaque': kind,
      'data-xml': xml,
      'data-text': text || null,
      'data-font': font || null,
      class: `opaque opaque-${kind}`,
      title: k.title || null,
      contenteditable: 'false',
      style: font ? `font-family: "${font}"` : null,
    });
    // xml/text — в атрибутах: так значок переживает копирование и вставку внутри документа
    const label = k.hidden ? '' : k.label === null ? text : k.label;
    return ['span', attrs, label];
  },

  // в обычный текст (копирование, счётчик слов) — только видимый текст поля/символа
  renderText({ node }) {
    return node.attrs.kind === 'fieldSimple' || node.attrs.kind === 'symbol' ? node.attrs.text : '';
  },
});
