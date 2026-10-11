// Сноски (src/editor/footnotes.js): чтение, создание, правка, удаление, сохранение (с исходным пакетом и без)
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { execFileSync } from 'child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { footnoteRefs, insertFootnoteTr, setFootnoteTextTr } from '../src/editor/footnotes.js';

const schema = getSchema(extensions);
const PY = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const SRC = 'tests/fixtures/footnotes.docx'; // tests/make_footnotes_fixture.ps1 (Word COM)
const load = async (bytes) => EditorState.create({ schema, doc: schema.nodeFromJSON((await importDocx(bytes)).doc) });
const notesOf = (s) => {
  const texts = new Map((s.doc.attrs.footnotes || []).map((n) => [String(n.id), n.text]));
  return footnoteRefs(s.doc).map((r) => texts.get(String(r.id)));
};
const part = async (bytes, name) => { const f = (await JSZip.loadAsync(bytes)).file(name); return f ? f.async('string') : null; };

describe('сноски', () => {
  it('читаются из файла Word', async () => {
    const s = await load(readFileSync(SRC));
    expect(notesOf(s)).toEqual(['Типовая форма компании.', 'Календарных дней с даты счёта.']);
    expect(footnoteRefs(s.doc).map((r) => r.n)).toEqual([1, 2]);
  });

  it('правка, новая сноска, удаление → сохранение поверх исходного файла', async () => {
    const src = readFileSync(SRC);
    let s = await load(src);
    const [r1, r2] = footnoteRefs(s.doc);
    s = s.apply(setFootnoteTextTr(s, r2.id, 'Рабочих дней.\nВторой абзац.'));
    s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, s.doc.content.size - 2)));
    const ins = insertFootnoteTr(s);
    s = s.apply(ins.tr);
    expect(ins.id).toBe('3');
    s = s.apply(setFootnoteTextTr(s, ins.id, 'Акт ТОРГ-12.'));
    expect(notesOf(s)).toEqual(['Типовая форма компании.', 'Рабочих дней.\nВторой абзац.', 'Акт ТОРГ-12.']);
    const bytes = await exportDocx(s.doc.toJSON(), { source: src });
    const orig = await part(src, 'word/footnotes.xml');
    const xml = await part(bytes, 'word/footnotes.xml');
    // неизменённая сноска и разделители — исходной разметкой
    const first = /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(orig)[0];
    expect(xml).toContain(first);
    expect(xml).toContain(/<w:footnote w:type="separator"[\s\S]*?<\/w:footnote>/.exec(orig)[0]);
    const back = await load(bytes);
    expect(notesOf(back)).toEqual(['Типовая форма компании.', 'Рабочих дней.\nВторой абзац.', 'Акт ТОРГ-12.']);
    // удаление: знак в тексте убран — сноска не сохраняется
    let d = back.apply(back.tr.delete(r1.pos, r1.pos + 1).setDocAttribute('footnotesEdited', true));
    expect(notesOf(d)).toEqual(['Рабочих дней.\nВторой абзац.', 'Акт ТОРГ-12.']);
    const bytes2 = await exportDocx(d.doc.toJSON(), { source: bytes });
    expect(await part(bytes2, 'word/footnotes.xml')).not.toContain('Типовая форма');
    d = await load(bytes2);
    expect(notesOf(d)).toEqual(['Рабочих дней.\nВторой абзац.', 'Акт ТОРГ-12.']);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/footnotes_resaved.docx', bytes);
  });

  it('новый документ со сноской: часть footnotes.xml, связь, тип; python-docx открывает', async () => {
    let s = EditorState.create({ schema, doc: schema.nodeFromJSON({ type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Новый текст' }] }] }) });
    s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, 6)));
    const ins = insertFootnoteTr(s);
    s = s.apply(ins.tr);
    s = s.apply(setFootnoteTextTr(s, ins.id, 'Сноска нового документа.'));
    const bytes = await exportDocx(s.doc.toJSON());
    expect(await part(bytes, 'word/footnotes.xml')).toContain('<w:footnote w:type="separator" w:id="-1">');
    expect(await part(bytes, 'word/_rels/document.xml.rels')).toMatch(/Type="[^"]*\/footnotes" Target="footnotes.xml"/);
    expect(await part(bytes, '[Content_Types].xml')).toContain('PartName="/word/footnotes.xml"');
    expect(await part(bytes, 'word/document.xml')).toContain('<w:footnoteReference w:id="1"/>');
    expect(notesOf(await load(bytes))).toEqual(['Сноска нового документа.']);
    writeFileSync('tests/out/footnotes_new.docx', bytes);
    const out = execFileSync(PY, ['-c', 'import docx,sys; d=docx.Document(sys.argv[1]); print(d.paragraphs[0].text)', 'tests/out/footnotes_new.docx'], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8').trim();
    expect(out).toBe('Новый текст');
  });
});
