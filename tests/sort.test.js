// Сортировка абзацев, пунктов списка и строк таблицы (src/editor/sort.js), горизонтальная линия
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { sortKey, sortNodes, sortTr } from '../src/editor/sort.js';

const schema = getSchema(extensions);
const p = (text) => ({ type: 'paragraph', content: text ? [{ type: 'text', text }] : [] });
const li = (text) => ({ type: 'listItem', content: [p(text)] });
const cell = (text) => ({ type: 'tableCell', content: [p(text)] });
const row = (...c) => ({ type: 'tableRow', content: c.map(cell) });
const state = (content, from, to = from) => {
  const s = EditorState.create({ schema, doc: schema.nodeFromJSON({ type: 'doc', content }) });
  return s.apply(s.tr.setSelection(TextSelection.create(s.doc, from, to)));
};
const texts = (doc) => { const out = []; doc.descendants((n) => { if (n.isTextblock) out.push(n.textContent); }); return out; };

describe('сортировка', () => {
  it('ключи: текст по-русски, числа с запятой и пробелами, даты дд.мм.гггг', () => {
    expect(sortKey(' 1 234,5 руб.', 'number')).toBe(1234.5);
    expect(sortKey('до 05.03.2024', 'date')).toBe(Date.UTC(2024, 2, 5));
    expect(sortNodes(['Ёж', 'ель', 'Яблоко', 'арбуз'], (x) => x).join(' ')).toBe('арбуз Ёж ель Яблоко'); // ё = е, как в Word
    expect(sortNodes(['10', 'нет', '9', '-1'], (x) => x, { type: 'number' })).toEqual(['-1', '9', '10', 'нет']);
    expect(sortNodes(['10', 'нет', '9'], (x) => x, { type: 'number', desc: true })).toEqual(['10', '9', 'нет']);
  });

  it('абзацы выделения по алфавиту; остальные на месте', () => {
    const s = state([p('Заголовок'), p('Сидоров'), p('Андреев'), p('Петров'), p('Подпись')], 12, 35);
    const r = s.apply(sortTr(s, { type: 'text' }));
    expect(texts(r.doc)).toEqual(['Заголовок', 'Андреев', 'Петров', 'Сидоров', 'Подпись']);
  });

  it('курсор в списке — весь список; по убыванию', () => {
    const s = state([{ type: 'bulletList', content: [li('б'), li('в'), li('а')] }], 4);
    const r = s.apply(sortTr(s, { desc: true }));
    expect(texts(r.doc)).toEqual(['в', 'б', 'а']);
  });

  it('таблица: строки по 2-му столбцу как числа, строка заголовка на месте', () => {
    const s = state([{ type: 'table', content: [row('Товар', 'Цена'), row('Молоко', '89,90'), row('Хлеб', '45'), row('Сыр', '1 250')] }], 5);
    const r = s.apply(sortTr(s, { type: 'number', column: 1, header: true }));
    expect(texts(r.doc)).toEqual(['Товар', 'Цена', 'Хлеб', '45', 'Молоко', '89,90', 'Сыр', '1 250']);
  });
});
