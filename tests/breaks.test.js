// Разрывы страницы, колонки, раздела и «Пустая страница» (src/editor/breaks.js) + сохранение в docx
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { extensions } from '../src/editor/extensions.js';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { sectionType, withSectionType } from '../src/editor/breaks.js';

const make = (texts) => new Editor({ extensions, content: { type: 'doc', content: texts.map((t) => ({ type: 'paragraph', content: t ? [{ type: 'text', text: t }] : [] })) } });
const blocks = (ed) => {
  const out = [];
  ed.state.doc.forEach((n) => {
    let s = '';
    n.forEach((c) => { s += c.isText ? c.text : c.type.name === 'opaque' ? `[${c.attrs.kind}]` : ''; });
    out.push(n.attrs.sectPr ? `${s}{раздел}` : s);
  });
  return out;
};

describe('разрывы', () => {
  it('тип начала раздела в sectPr (w:type перед pgSz)', () => {
    const x = withSectionType('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>', 'continuous');
    expect(x).toBe('<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>');
    expect(sectionType(x)).toBe('continuous');
    expect(sectionType(withSectionType(x, 'nextPage'))).toBe('nextPage');
  });

  it('разрыв страницы делит абзац, «Пустая страница» — два разрыва и пустой абзац', () => {
    const ed = make(['ПервыйВторой']);
    ed.commands.setTextSelection(7);
    ed.commands.insertBreak('page');
    expect(blocks(ed)).toEqual(['Первый[pageBreak]', 'Второй']);
    ed.commands.insertBlankPage();
    expect(blocks(ed)).toEqual(['Первый[pageBreak]', '[pageBreak]', '[pageBreak]', 'Второй']);
    ed.destroy();
  });

  it('разрыв раздела «на текущей странице»: раздел до курсора, тип — у следующего; сохраняется в docx', async () => {
    const ed = make(['Раздел один', 'Раздел два']);
    ed.commands.setTextSelection(12); // конец первого абзаца
    ed.commands.insertSectionBreak('continuous');
    expect(blocks(ed)).toEqual(['Раздел один{раздел}', '', 'Раздел два']);
    expect(sectionType(ed.state.doc.child(0).attrs.sectPr)).toBe('nextPage');
    expect(sectionType(ed.state.doc.attrs.sectPr)).toBe('continuous');
    const bytes = await exportDocx(ed.getJSON());
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>.*?<\/w:sectPr><\/w:pPr><w:r><w:t xml:space="preserve">Раздел один/);
    expect(xml).toMatch(/<w:sectPr><w:type w:val="continuous"\/><w:pgSz/);
    const back = (await importDocx(bytes)).doc;
    expect(back.content[0].attrs.sectPr).toBeTruthy();
    expect(sectionType(back.attrs.sectPr)).toBe('continuous');
    ed.destroy();
  });
});
