// Колонки текста: w:cols ↔ модель, контейнер columns при импорте/экспорте, команды (src/editor/columns.js, colCommands.js).
// Раскладка на экране сверяется с Word отдельно: tests/columns_check.mjs (эталон — tests/make_columns_fixture.ps1).
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import JSZip from 'jszip';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { parseCols, patchCols } from '../src/editor/columns.js';
import { columnsTr, sectionsOf } from '../src/editor/colCommands.js';

const schema = getSchema(extensions);
const SECT = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
const para = (t) => ({ type: 'paragraph', content: [{ type: 'text', text: t }] });
const shape = (doc) => doc.content.content.map((n) => (n.type.name === 'columns' ? `[${n.childCount}${n.attrs.sectPr ? '|' : ''}]` : `${n.textContent}${n.attrs.sectPr ? '|' : ''}`)).join(' ');

describe('колонки текста', () => {
  it('w:cols ↔ параметры (равные и разной ширины), место в sectPr', () => {
    expect(parseCols('<w:sectPr><w:cols w:space="708"/></w:sectPr>')).toBe(null);
    expect(parseCols('<w:cols w:num="2" w:sep="1" w:space="709"/>')).toEqual({ num: 2, space: 709, sep: true, equal: true, cols: [] });
    const un = parseCols('<w:cols w:num="2" w:space="708" w:equalWidth="0"><w:col w:w="2551" w:space="709"/><w:col w:w="6094" w:space="708"/></w:cols>');
    expect(un).toMatchObject({ num: 2, equal: false, cols: [{ w: 2551, space: 709 }, { w: 6094, space: 708 }] });
    const x = patchCols(SECT, { num: 3, space: 425, sep: false });
    expect(x).toContain('<w:cols w:num="3" w:space="425"/><w:docGrid');
    expect(parseCols(patchCols(x, un))).toEqual(un);
    expect(parseCols(patchCols(x, null))).toBe(null);
  });

  it('файл Word: разделы с колонками — в контейнерах, сохранение не меняет sectPr', async () => {
    const src = readFileSync('tests/fixtures/columns.docx');
    const { doc } = await importDocx(src);
    expect(doc.content.map((n) => n.type)).toEqual(['paragraph', 'paragraph', 'columns', 'paragraph', 'paragraph', 'columns', 'paragraph', 'paragraph', 'columns']);
    expect(parseCols(doc.content[2].attrs.sectPr)).toMatchObject({ num: 2, sep: true });
    expect(doc.content[8].attrs.sectPr).toBe(null); // последний раздел — sectPr документа
    const xml = async (bytes) => (await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string')).match(/<w:cols\b(?:[^>]*?\/>|[^>]*>[\s\S]*?<\/w:cols>)/g);
    expect(await xml(await exportDocx(doc, { source: src }))).toEqual(await xml(src));
    const back = (await importDocx(await exportDocx(doc))).doc;
    expect(back.content.map((n) => n.type)).toEqual(doc.content.map((n) => n.type));
  });

  it('команды: раздел целиком, снова одна колонка, выделенный текст, до конца документа', () => {
    const doc = schema.nodeFromJSON({ type: 'doc', attrs: { sectPr: SECT }, content: [para('a'), para('b'), para('c'), para('d')] });
    const at = (d, from, to = from) => EditorState.create({ schema, doc: d, selection: TextSelection.create(d, from, to) });
    const two = { num: 2, space: 708, sep: false };
    // весь (последний) раздел
    let d = columnsTr(at(doc, 2), two).doc;
    expect(shape(d)).toBe('[4]');
    expect(parseCols(d.attrs.sectPr).num).toBe(2);
    d = columnsTr(at(d, 3), null).doc;
    expect(shape(d)).toBe('a b c d');
    // выделение b..c — отдельный раздел «на текущей странице»
    d = columnsTr(at(doc, 4, 7), two, 'selection').doc;
    expect(shape(d)).toBe('a| [2|] d');
    const box = d.child(1);
    expect(parseCols(box.attrs.sectPr).num).toBe(2);
    expect(box.attrs.sectPr).toContain('<w:type w:val="continuous"/>');
    expect(d.attrs.sectPr).toContain('<w:type w:val="continuous"/>');
    expect(parseCols(d.child(0).attrs.sectPr)).toBe(null);
    expect(sectionsOf(d)).toHaveLength(3);
    // до конца документа — с абзаца c
    d = columnsTr(at(doc, 7), two, 'forward').doc;
    expect(shape(d)).toBe('a b| [2]');
    expect(parseCols(d.attrs.sectPr).num).toBe(2);
  });

  it('выделенный текст с колонками сохраняется и читается обратно тем же строением', async () => {
    const doc = schema.nodeFromJSON({ type: 'doc', attrs: { sectPr: SECT }, content: [para('a'), para('b'), para('c'), para('d')] });
    const d = columnsTr(EditorState.create({ schema, doc, selection: TextSelection.create(doc, 4, 7) }), { num: 3, space: 425, sep: true }, 'selection').doc;
    const back = (await importDocx(await exportDocx(d.toJSON()))).doc;
    expect(back.content.map((n) => (n.type === 'columns' ? `[${n.content.length}]` : n.content[0].text))).toEqual(['a', '[2]', 'd']);
    expect(parseCols(back.content[1].attrs.sectPr)).toMatchObject({ num: 3, sep: true, space: 425 });
  });
});
