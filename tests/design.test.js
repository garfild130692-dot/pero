// Вкладка «Конструктор» (src/editor/design.js): подложка, цвет страницы, границы страниц — чтение и запись как у Word
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { addWatermark, parsePgBorders, patchPgBorders, readWatermark, stripWatermark } from '../src/editor/design.js';

const SRC = 'tests/fixtures/design.docx'; // tests/make_design_fixture.ps1 (Word COM)
const files = async (bytes, re) => {
  const zip = await JSZip.loadAsync(bytes);
  return Promise.all(Object.keys(zip.files).filter((n) => re.test(n)).sort().map(async (n) => [n, await zip.file(n).async('string')]));
};
const part = async (bytes, name) => (await JSZip.loadAsync(bytes)).file(name).async('string');

describe('конструктор', () => {
  it('файл Word: подложка, цвет страницы, рамка', async () => {
    const { doc } = await importDocx(readFileSync(SRC));
    expect(doc.attrs.watermark).toEqual({ text: 'ОБРАЗЕЦ', font: 'Calibri', color: '#c0c0c0', semi: true, layout: 'diagonal', size: null });
    expect(doc.attrs.pageColor).toBe('#ffffcc');
    expect(parsePgBorders(doc.attrs.sectPr)).toEqual({ offsetFrom: 'page', style: 'single', width: 1, color: '#0070c0', space: 24 });
  });

  it('XML: подложка в абзаце колонтитула, удаление; рамка — после pgMar', () => {
    const run = '<w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1"><v:textpath string="X"/></v:shape></w:pict></w:r>';
    const hdr = '<w:hdr><w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr></w:p></w:hdr>';
    const withWm = addWatermark(hdr, run);
    expect(withWm).toBe(`<w:hdr><w:p><w:pPr><w:pStyle w:val="a3"/></w:pPr>${run}</w:p></w:hdr>`);
    expect(addWatermark('<w:hdr><w:p/></w:hdr>', run)).toBe(`<w:hdr><w:p>${run}</w:p></w:hdr>`);
    expect(stripWatermark(withWm)).toBe(hdr);
    expect(readWatermark(withWm)).toMatchObject({ text: 'X', layout: 'horizontal' });
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134"/><w:cols w:space="708"/></w:sectPr>';
    const b = { offsetFrom: 'text', style: 'double', width: 0.75, color: null, space: 4 };
    const x = patchPgBorders(sect, b);
    expect(x).toMatch(/<w:pgMar w:top="1134"\/><w:pgBorders><w:top w:val="double" w:sz="6" w:space="4" w:color="auto"\/>.*<\/w:pgBorders><w:cols/);
    expect(parsePgBorders(x)).toEqual(b);
    expect(patchPgBorders(x, null)).toBe(sect);
  });

  it('сохранение поверх файла Word: другая подложка во всех колонтитулах, без цвета, рамка как была', async () => {
    const src = readFileSync(SRC);
    const { doc } = await importDocx(src);
    doc.attrs = { ...doc.attrs, watermark: { ...doc.attrs.watermark, text: 'КОПИЯ', layout: 'horizontal' }, watermarkChanged: true, pageColor: null, pageColorChanged: true };
    const bytes = await exportDocx(doc, { source: src });
    const hdrs = await files(bytes, /^word\/header\d+\.xml$/);
    expect(hdrs.length).toBe(3);
    for (const [, x] of hdrs) { expect(x).toContain('string="КОПИЯ"'); expect(x).not.toContain('ОБРАЗЕЦ'); }
    const back = (await importDocx(bytes)).doc;
    expect(back.attrs.watermark).toMatchObject({ text: 'КОПИЯ', layout: 'horizontal', semi: true });
    expect(back.attrs.pageColor).toBe(null);
    expect(await part(bytes, 'word/settings.xml')).not.toContain('displayBackgroundShape');
    expect(parsePgBorders(back.attrs.sectPr)).toMatchObject({ offsetFrom: 'page', width: 1 });
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/design_resaved.docx', bytes);
    // подложку убрали
    doc.attrs = { ...doc.attrs, watermark: null };
    const none = await files(await exportDocx(doc, { source: src }), /^word\/header\d+\.xml$/);
    for (const [, x] of none) expect(x).not.toContain('PowerPlusWaterMarkObject');
  });

  it('новый документ: свой колонтитул с подложкой, цвет страницы, рамка', async () => {
    const doc = {
      type: 'doc',
      attrs: {
        watermark: { text: 'ЧЕРНОВИК', font: 'Arial', color: '#ff0000', semi: false, layout: 'diagonal', size: null }, watermarkChanged: true,
        pageColor: '#eaf3ef', pageColorChanged: true,
        sectPr: patchPgBorders('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
          { offsetFrom: 'text', style: 'single', width: 0.5, color: null, space: 4 }),
      },
      content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Текст' }] }],
    };
    const bytes = await exportDocx(doc);
    const d = await part(bytes, 'word/document.xml');
    expect(d).toMatch(/<w:background w:color="EAF3EF"\/><w:body>/);
    expect(d).toMatch(/<w:headerReference w:type="default" r:id="rIdPeroH"\/>/);
    expect(await part(bytes, 'word/settings.xml')).toContain('<w:displayBackgroundShape/>');
    expect(await part(bytes, 'word/header_pero.xml')).toContain('string="ЧЕРНОВИК"');
    const back = (await importDocx(bytes)).doc;
    expect(back.attrs.watermark).toEqual({ text: 'ЧЕРНОВИК', font: 'Arial', color: '#ff0000', semi: false, layout: 'diagonal', size: null });
    expect(back.attrs.pageColor).toBe('#eaf3ef');
    expect(parsePgBorders(back.attrs.sectPr)).toEqual({ offsetFrom: 'text', style: 'single', width: 0.5, color: null, space: 4 });
    writeFileSync('tests/out/design_new.docx', bytes);
  });
});
