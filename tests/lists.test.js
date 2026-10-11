// Маркеры, библиотека номеров, продолжение и перезапуск нумерации (src/editor/numbering.js, import/export)
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'fs';
import { Editor, getSchema } from '@tiptap/core';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { bulletFromLvl, bulletToLvl, listStarts } from '../src/editor/numbering.js';

const SRC = 'tests/fixtures/lists.docx'; // tests/make_lists_fixture.ps1 (Word COM)
const schema = getSchema(extensions);
const lists = (doc) => doc.content.filter((n) => /List$/.test(n.type)).map((n) => ({ type: n.type, ...n.attrs }));

describe('списки', () => {
  it('знаки маркеров ↔ Symbol/Wingdings', () => {
    expect(bulletToLvl('•')).toEqual(['', 'Symbol']);
    expect(bulletFromLvl('', 'Wingdings')).toBe('➢');
    expect(bulletFromLvl('o', 'Courier New')).toBe('○');
    expect(bulletFromLvl('', 'Wingdings')).toBe('•'); // неизвестный знак шрифта символов
    expect(bulletToLvl('★')).toEqual(['★', 'Segoe UI Symbol']);
  });

  it('из Word: маркеры ➢ и ✓, продолжение нумерации через абзац, новый список; сохранение сохраняет', async () => {
    const src = readFileSync(SRC);
    const { doc } = await importDocx(src);
    const want = [
      { type: 'bulletList', bullet: '➢' }, { type: 'bulletList', bullet: '✓' },
      { type: 'orderedList', start: 1, listId: 'L1' }, { type: 'orderedList', start: 1, listId: 'L1' }, { type: 'orderedList', start: 1 },
    ];
    expect(lists(doc)).toEqual(want);
    const node = schema.nodeFromJSON(doc);
    expect(listStarts(node).map((x) => [x.start, x.continued])).toEqual([[1, false], [3, true]]);
    for (const bytes of [await exportDocx(doc, { source: src }), await exportDocx(doc)]) {
      expect(lists((await importDocx(bytes)).doc)).toEqual(want);
      const d = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
      const ids = [...d.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);
      expect(ids[4]).toBe(ids[6]); // продолжение — тот же w:num
      expect(ids[8]).not.toBe(ids[4]); // новый список — свой
    }
  });

  it('команды: маркер, формат из библиотеки, начать заново, продолжить', () => {
    const editor = new Editor({ extensions, content: {
      type: 'doc',
      content: [
        { type: 'orderedList', content: [1, 2].map((n) => ({ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: `a${n}` }] }] })) },
        { type: 'paragraph', content: [{ type: 'text', text: 'между' }] },
        { type: 'orderedList', content: [1, 2, 3].map((n) => ({ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: `b${n}` }] }] })) },
        { type: 'paragraph', content: [{ type: 'text', text: 'конец' }] },
      ],
    } });
    const posOf = (t) => { let r = 0; editor.state.doc.descendants((n, p) => { if (n.isText && n.text === t) r = p; }); return r + 1; };
    editor.commands.setTextSelection(posOf('b2'));
    editor.commands.continueNumbering();
    let ls = listStarts(editor.state.doc);
    expect(ls.map((x) => x.start)).toEqual([1, 3]);
    editor.commands.restartNumbering(5);
    expect(editor.state.doc.child(2).attrs).toMatchObject({ listId: null, start: 5 });
    editor.commands.setNumberFormat('upperRoman', '%1)');
    expect(editor.state.doc.child(2).attrs).toMatchObject({ numFmt: 'upperRoman', lvlText: '%1)' });
    editor.commands.setTextSelection(posOf('конец'));
    editor.commands.setBullet('➢');
    expect(editor.state.doc.lastChild.type.name).toBe('bulletList');
    expect(editor.state.doc.lastChild.attrs.bullet).toBe('➢');
    editor.commands.setBullet('•'); // знак уровня по умолчанию — атрибут не нужен
    expect(editor.state.doc.lastChild.attrs.bullet).toBe(null);
    editor.destroy();
  });
});
