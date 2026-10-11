// Поля сведений о документе, свойства документа (core.xml), «Текст из файла» (src/editor/fields.js, docProps.js, insertFile.js)
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { getSchema } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { bookmarks, fieldValue } from '../src/editor/fields.js';
import { patchCoreProps, readCoreProps } from '../src/editor/docProps.js';
import { prepareInsert } from '../src/insertFile.js';

const schema = getSchema(extensions);

describe('экспресс-блоки', () => {
  it('поля сведений о документе', () => {
    const doc = schema.nodeFromJSON({
      type: 'doc', attrs: { props: { author: 'Перо Тест', title: 'Договор', created: '2026-10-01T09:30:00Z' } },
      content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Раз два три' }] }],
    });
    const state = EditorState.create({ schema, doc });
    const ctx = { state, marks: bookmarks(doc), now: new Date(2026, 9, 11), file: { name: 'Договор.docx', path: 'D:\\Документы\\Договор.docx' } };
    expect(fieldValue('AUTHOR', 1, ctx)).toBe('Перо Тест');
    expect(fieldValue('DOCPROPERTY Title', 1, ctx)).toBe('Договор');
    expect(fieldValue('SUBJECT', 1, ctx)).toBe('');
    expect(fieldValue('FILENAME', 1, ctx)).toBe('Договор.docx');
    expect(fieldValue('FILENAME \\p', 1, ctx)).toBe('D:\\Документы\\Договор.docx');
    expect(fieldValue('NUMWORDS', 1, ctx)).toBe('3');
    expect(fieldValue('CREATEDATE \\@ "dd.MM.yyyy"', 1, ctx)).toBe('01.10.2026');
  });

  it('свойства документа ↔ core.xml (исходные элементы и прочее — как было)', async () => {
    const core = '<cp:coreProperties xmlns:cp="c" xmlns:dc="d" xmlns:dcterms="t" xmlns:xsi="x"><dc:creator>Старый</dc:creator><cp:revision>3</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>';
    const x = patchCoreProps(core, { author: 'Новый', title: 'Название & Ко' }, new Date('2026-10-11T10:00:00Z'));
    expect(x).toContain('<dc:creator>Новый</dc:creator><cp:revision>3</cp:revision>');
    expect(x).toContain('<dc:title>Название &amp; Ко</dc:title>');
    expect(x).toContain('2026-10-11T10:00:00Z');
    expect(readCoreProps(x)).toMatchObject({ author: 'Новый', title: 'Название & Ко' });
    const src = readFileSync('tests/fixtures/footnotes.docx');
    const { doc } = await importDocx(src);
    expect(doc.attrs.props.author).toBe('Перо Тест');
    doc.attrs = { ...doc.attrs, props: { ...doc.attrs.props, title: 'Поставка' }, propsChanged: true };
    const back = (await importDocx(await exportDocx(doc, { source: src }))).doc;
    expect(back.attrs.props).toMatchObject({ author: 'Перо Тест', title: 'Поставка' });
  });

  it('текст из файла: сноски с новыми номерами, примечания и фигуры не переносятся, рисунки — как новые', async () => {
    const notes = (await importDocx(readFileSync('tests/fixtures/footnotes.docx'))).doc;
    const p = prepareInsert(notes, { footnotes: [{ id: '1', text: 'своя' }, { id: '2', text: 'своя 2' }] });
    expect(p.footnotes.map((f) => [f.id, f.text])).toEqual([['3', 'Типовая форма компании.'], ['4', 'Календарных дней с даты счёта.']]);
    const ids = [];
    const walk = (n) => { if (n.type === 'opaque' && n.attrs.kind === 'footnote') ids.push(/w:id="(\d+)"/.exec(n.attrs.xml)[1]); (n.content || []).forEach(walk); };
    p.content.forEach(walk);
    expect(ids).toEqual(['3', '4']);
    const pres = (await importDocx(readFileSync('tests/fixtures/preserve.docx'))).doc;
    const q = prepareInsert(pres, {});
    expect(q.dropped).toBeGreaterThan(0);
    const kinds = new Set();
    const walk2 = (n) => { if (n.type === 'opaque') kinds.add(n.attrs.kind); if (n.type === 'image') expect(n.attrs.xml).toBe(null); (n.content || []).forEach(walk2); };
    q.content.forEach(walk2);
    expect(kinds.has('comment')).toBe(false);
    expect(q.content.every((n) => !(n.attrs && n.attrs.sectPr))).toBe(true);
  });
});
