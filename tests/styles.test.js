// Встроенные стили Word (абзацные и символьные), панель/галерея — модель (src/editor/builtinStyles.js, import/export)
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'fs';
import { Editor } from '@tiptap/core';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { BUILTIN_CHAR, BUILTIN_PARA, charStyleCss, ensureStyle } from '../src/editor/builtinStyles.js';

const quote = BUILTIN_PARA.find((s) => s.id === 'Quote');
const emph = BUILTIN_CHAR.find((s) => s.id === 'Emphasis');
const part = async (bytes, name) => (await JSZip.loadAsync(bytes)).file(name).async('string');

describe('встроенные стили', () => {
  it('определение добавляется один раз; стиль файла с тем же именем — его id', () => {
    const a = ensureStyle([], quote);
    expect(a.id).toBe('Quote');
    expect(a.list[0]).toMatchObject({ id: 'Quote', name: 'Quote', builtin: true, origin: 'pero' });
    expect(ensureStyle(a.list, quote).list).toBe(a.list);
    expect(ensureStyle([{ id: 'af2', name: 'Quote', origin: 'file' }], quote).id).toBe('af2');
    expect(charStyleCss([{ id: 'Strong', run: { bold: true } }])).toBe('.page-content span[data-cstyle="Strong"] { font-weight: bold }');
  });

  it('применение: стиль абзаца «Цитата», символьный «Выделение» (без выделения — слово, повтор — снять)', () => {
    const para = (t) => ({ type: 'paragraph', content: [{ type: 'text', text: t }] });
    const ed = new Editor({ extensions, content: { type: 'doc', content: [para('Первый абзац текста'), para('Второй')] } });
    ed.commands.setTextSelection(3);
    ed.commands.applyBuiltinParaStyle(quote);
    expect(ed.state.doc.child(0).attrs.styleId).toBe('Quote');
    expect(ed.state.doc.attrs.styles.map((s) => s.name)).toEqual(['Quote']);
    ed.commands.setTextSelection(10); // внутри «абзац»
    ed.commands.applyCharStyle(emph);
    const marked = [];
    ed.state.doc.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'charStyle')) marked.push(n.text); });
    expect(marked).toEqual(['абзац']);
    expect(ed.state.doc.attrs.charStyles.map((s) => s.id)).toEqual(['Emphasis']);
    ed.commands.applyCharStyle(emph);
    let any = false;
    ed.state.doc.descendants((n) => { if (n.isText && n.marks.some((m) => m.type.name === 'charStyle')) any = true; });
    expect(any).toBe(false);
    ed.destroy();
  });

  it('сохранение: новый документ — определения встроенных стилей; в файл со стилем «Emphasis» под своим id — ссылка на него', async () => {
    const doc = {
      type: 'doc',
      attrs: {
        styles: [{ ...ensureStyle([], quote).list[0] }],
        charStyles: [{ ...ensureStyle([], emph).list[0] }],
      },
      content: [{ type: 'paragraph', attrs: { styleId: 'Quote' }, content: [{ type: 'text', text: 'Цитата ' }, { type: 'text', text: 'важно', marks: [{ type: 'charStyle', attrs: { id: 'Emphasis' } }] }] }],
    };
    const bytes = await exportDocx(doc);
    const st = await part(bytes, 'word/styles.xml');
    expect(st).toContain('<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>');
    expect(st).toMatch(/<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"\/>.*<w:i\/>/);
    expect(await part(bytes, 'word/document.xml')).toContain('<w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t xml:space="preserve">важно</w:t>');
    const back = (await importDocx(bytes)).doc;
    expect(back.content[0].attrs.styleId).toBe('Quote');
    expect(back.content[0].content[1].marks).toEqual([{ type: 'charStyle', attrs: { id: 'Emphasis' } }]);
    expect(back.attrs.charStyles).toEqual([{ id: 'Emphasis', name: 'Emphasis', run: { italic: true }, origin: 'file' }]);

    // исходный файл, где «Emphasis» — стиль с русским id «a9»: новый не дописывается, текст ссылается на «a9»
    const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/sample.docx'));
    const sx = await zip.file('word/styles.xml').async('string');
    zip.file('word/styles.xml', sx.replace('w:styleId="Emphasis"', 'w:styleId="a9"')); // как в русском Word: свой id
    const src = await zip.generateAsync({ type: 'uint8array' });
    const a = (await importDocx(src)).doc;
    a.attrs.charStyles = [...a.attrs.charStyles, ensureStyle([], emph).list[0]];
    a.content.push({ type: 'paragraph', content: [{ type: 'text', text: 'выделено', marks: [{ type: 'charStyle', attrs: { id: 'Emphasis' } }] }] });
    const out = await exportDocx(a, { source: src });
    expect(await part(out, 'word/document.xml')).toContain('<w:rStyle w:val="a9"/></w:rPr><w:t xml:space="preserve">выделено</w:t>');
    expect(await part(out, 'word/styles.xml')).not.toContain('w:styleId="Emphasis"');
  });
});
