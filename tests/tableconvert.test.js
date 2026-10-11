// «Вставить таблицу…», «Текст → таблица», «Таблица → текст» (src/editor/tableConvert.js)
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { columnsFor, guessSeparator, makeTable, tableToTextTr, textToTableTr } from '../src/editor/tableConvert.js';

const schema = getSchema(extensions);
const p = (...parts) => ({ type: 'paragraph', content: parts.map((x) => (x === '\t' ? { type: 'tab' } : typeof x === 'string' ? { type: 'text', text: x } : x)) });
const state = (content, from, to = from) => {
  const s = EditorState.create({ schema, doc: schema.nodeFromJSON({ type: 'doc', content }) });
  return s.apply(s.tr.setSelection(TextSelection.create(s.doc, from, to)));
};
const grid = (table) => { const out = []; table.forEach((r) => { const row = []; r.forEach((c) => row.push(c.textContent)); out.push(row); }); return out; };

describe('таблицы: вставка и преобразование', () => {
  it('таблица заданного размера: ширина столбцов в см, строка заголовка', () => {
    const t = makeTable(schema, 2, 3, { widthCm: 2.54, header: true });
    expect(t.childCount).toBe(2);
    expect(t.firstChild.childCount).toBe(3);
    expect(t.firstChild.attrs.repeatHeader).toBe(true);
    expect(t.firstChild.firstChild.attrs.colwidth).toEqual([96]);
  });

  it('текст с табуляцией → таблица (оформление сохраняется), разделитель угадывается', () => {
    const s = state([p('ФИО', '\t', 'Должность'), p('Иванов', '\t', { type: 'text', text: 'директор', marks: [{ type: 'bold' }] }), p('Петров')], 2, 35);
    expect(guessSeparator(s)).toBe('tab');
    expect(columnsFor(s, 'tab')).toBe(2);
    const r = s.apply(textToTableTr(s, { sep: 'tab' }));
    const t = r.doc.firstChild;
    expect(t.type.name).toBe('table');
    expect(grid(t)).toEqual([['ФИО', 'Должность'], ['Иванов', 'директор'], ['Петров', '']]);
    let bold = false;
    t.child(1).child(1).descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'bold')) bold = true; });
    expect(bold).toBe(true);
  });

  it('«;» и абзацы → таблица; таблица → текст обратно', () => {
    const s = state([p('a;b;c'), p('d;e')], 2, 10);
    expect(guessSeparator(s)).toBe(';');
    const t = s.apply(textToTableTr(s, { sep: ';' })).doc.firstChild;
    expect(grid(t)).toEqual([['a', 'b', 'c'], ['d', 'e', '']]);
    const s2 = state([p('1'), p('2'), p('3'), p('4')], 2, 11);
    expect(grid(s2.apply(textToTableTr(s2, { sep: 'para', cols: 2 })).doc.firstChild)).toEqual([['1', '2'], ['3', '4']]);
    // обратно: строки — абзацы, ячейки через «;» / табуляцию
    const st = EditorState.create({ schema, doc: schema.nodeFromJSON({ type: 'doc', content: [makeTable(schema, 2, 2, { cells: [[[schema.text('x')], [schema.text('y')]], [[schema.text('z')], []]] }).toJSON()] }) });
    const s3 = st.apply(st.tr.setSelection(TextSelection.create(st.doc, 4)));
    const back = s3.apply(tableToTextTr(s3, { sep: ';' })).doc;
    expect([back.child(0).textContent, back.child(1).textContent]).toEqual(['x;y', 'z;']);
    const tabs = s3.apply(tableToTextTr(s3, { sep: 'tab' })).doc;
    expect(tabs.child(0).child(1).type.name).toBe('tab');
  });
});
