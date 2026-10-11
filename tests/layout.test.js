// Вкладка «Макет»: счётчики абзаца (src/ParaSpinners.jsx), «Параметры страницы» (переплёт, бумага, зеркальные поля)
import { describe, expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { paraValue } from '../src/ParaSpinners.jsx';
import { readFileSync } from 'fs';
import { MIRROR_AFTER, parseSectPr, patchSectPr, patchSettingsFlag } from '../src/editor/page.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import JSZip from 'jszip';
import { hyphenate } from '../src/editor/hyphen.js';

const schema = getSchema(extensions);

describe('«Макет → Абзац»', () => {
  it('действующее значение абзацев выделения, разные — null', () => {
    const doc = schema.nodeFromJSON({
      type: 'doc', attrs: { para: { spaceAfter: 8 } },
      content: [
        { type: 'paragraph', attrs: { indentLeft: 28.35 }, content: [{ type: 'text', text: 'раз' }] },
        { type: 'paragraph', content: [{ type: 'text', text: 'два' }] },
      ],
    });
    const at = (from, to) => EditorState.create({ schema, doc, selection: TextSelection.create(doc, from, to) });
    expect(paraValue(at(2, 2), 'indentLeft')).toBe(28.35);
    expect(paraValue(at(2, 2), 'spaceAfter')).toBe(8);
    expect(paraValue(at(2, 7), 'indentLeft')).toBe(null);
    expect(paraValue(at(2, 7), 'spaceAfter')).toBe(8);
  });
});


describe('«Параметры страницы»', () => {
  const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="567"/></w:sectPr>';
  it('переплёт входит в левое поле раскладки, при записи — отдельно', () => {
    const g = parseSectPr(sect);
    expect([g.marLeft, g.gutter, g.left]).toEqual([1701, 567, 2268]);
    // линейка тянет поле вместе с переплётом — в файл уходит поле без него
    expect(patchSectPr(sect, { left: 2000 })).toContain('w:left="1433"');
    const x = patchSectPr(sect, { left: 1800, gutter: 0 });
    expect(x).toContain('w:left="1800"');
    expect(x).toContain('w:gutter="0"');
  });
  it('код бумаги: свой формат — убирается, известный — ставится', () => {
    expect(patchSectPr(sect, { w: 10000, h: 15000 })).not.toContain('w:code');
    expect(patchSectPr(sect, { w: 16838, h: 23811, code: 8 })).toContain('w:code="8"');
    expect(patchSectPr(sect, { top: 1000 })).toContain('w:code="9"');
  });
  it('w:mirrorMargins — на месте по схеме', () => {
    const s = '<w:settings><w:zoom w:percent="100"/><w:proofState w:spelling="clean"/><w:defaultTabStop w:val="708"/></w:settings>';
    expect(patchSettingsFlag(s, 'mirrorMargins', true, MIRROR_AFTER)).toBe('<w:settings><w:zoom w:percent="100"/><w:mirrorMargins/><w:proofState w:spelling="clean"/><w:defaultTabStop w:val="708"/></w:settings>');
    expect(patchSettingsFlag(patchSettingsFlag(s, 'mirrorMargins', true, MIRROR_AFTER), 'mirrorMargins', false, MIRROR_AFTER)).toBe(s);
  });
  it('зеркальные поля сохраняются (новый документ и поверх исходного)', async () => {
    const src = readFileSync('tests/fixtures/sample.docx');
    const { doc } = await importDocx(src);
    expect(doc.attrs.mirrorMargins).toBeUndefined();
    doc.attrs = { ...doc.attrs, mirrorMargins: true, mirrorChanged: true };
    const out = await exportDocx(doc, { source: src });
    const sx = await (await JSZip.loadAsync(out)).file('word/settings.xml').async('string');
    expect(sx).toContain('<w:mirrorMargins/>');
    expect((await importDocx(out)).doc.attrs.mirrorMargins).toBe(true);
    const fresh = await exportDocx(doc);
    expect((await importDocx(fresh)).doc.attrs.mirrorMargins).toBe(true);
  });
});

describe('расстановка переносов', () => {
  it('русские слова по слогам, не меньше двух букв с краю, без отрыва й/ь/ъ', () => {
    expect(hyphenate('электрификация')).toBe('элек-три-фи-ка-ция');
    expect(hyphenate('производственных')).toBe('про-из-вод-ствен-ных');
    expect(hyphenate('подъезд')).toBe('подъ-езд');
    expect(hyphenate('майский')).toBe('май-ский');
    expect(hyphenate('окно')).toBe('ок-но');
    expect(hyphenate('оса')).toBe('оса');
  });
  it('параметры документа и запрет у абзаца сохраняются', async () => {
    const src = readFileSync('tests/fixtures/sample.docx');
    const { doc } = await importDocx(src);
    doc.attrs = { ...doc.attrs, autoHyphenation: true, hyphenateCaps: false, hyphenChanged: true };
    const p = doc.content.find((n) => n.type === 'paragraph');
    p.attrs = { ...p.attrs, noHyphens: true };
    for (const out of [await exportDocx(doc, { source: src }), await exportDocx(doc)]) {
      const zip = await JSZip.loadAsync(out);
      const sx = await zip.file('word/settings.xml').async('string');
      expect(sx).toMatch(/<w:defaultTabStop\b[^>]*\/><w:autoHyphenation\/>/);
      expect(sx).toContain('<w:doNotHyphenateCaps/>');
      expect(await zip.file('word/document.xml').async('string')).toContain('<w:suppressAutoHyphens/>');
      const back = (await importDocx(out)).doc;
      expect(back.attrs.autoHyphenation).toBe(true);
      expect(back.attrs.hyphenateCaps).toBe(false);
      expect(back.content.find((n) => n.type === 'paragraph').attrs.noHyphens).toBe(true);
    }
  });
});
