// Поля, закладки, перекрёстные ссылки, оглавление (src/editor/fields.js) — на состоянии ProseMirror и через docx
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { execFileSync } from 'child_process';
import { mkdirSync, writeFileSync } from 'fs';
import { getSchema } from '@tiptap/core';
import { EditorState, TextSelection } from '@tiptap/pm/state';
import { extensions } from '../src/editor/extensions.js';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import {
  addBookmarkTr, bookmarks, fieldNodes, formatDate, headingList, insertCrossRefTr, insertDateTr, insertTocTr, removeBookmarkTr,
  removeTocTr, scanFields, tocOptions, updateFieldsTr, REF_ERROR,
} from '../src/editor/fields.js';

const schema = getSchema(extensions);
const PY = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const p = (text, attrs) => ({ type: 'paragraph', ...(attrs ? { attrs } : {}), content: text ? [{ type: 'text', text }] : [] });
const h = (level, text) => ({ type: 'heading', attrs: { level }, content: [{ type: 'text', text }] });
const stateOf = (content, sel = 1) => {
  const s = EditorState.create({ schema, doc: schema.nodeFromJSON({ type: 'doc', content }) });
  return s.apply(s.tr.setSelection(TextSelection.create(s.doc, sel)));
};
const run = (s, tr) => s.apply(tr);
const txt = (n) => n.textBetween(0, n.content.size, '', (leaf) => (leaf.type.name === 'tab' ? '	' : ''));
const withField = (s, instr, text) => run(s, s.tr.insert(s.selection.from, fieldNodes(schema, instr, text)));

describe('поля', () => {
  it('формат даты как в Word', () => {
    const d = new Date(2026, 9, 11, 9, 5, 7);
    expect(formatDate(d, 'dd.MM.yyyy')).toBe('11.10.2026');
    expect(formatDate(d, "d MMMM yyyy 'г.'")).toBe('11 октября 2026 г.');
    expect(formatDate(d, 'MMMM yy')).toBe('октябрь 26');
    expect(formatDate(d, 'dddd, d MMMM')).toBe('воскресенье, 11 октября');
    expect(formatDate(d, 'H:mm:ss')).toBe('9:05:07');
    expect(formatDate(d, 'hh:mm AM/PM')).toBe('09:05 AM');
  });

  it('разбор и обновление: DATE, PAGE, NUMPAGES, REF, ошибка ссылки', () => {
    let s = stateOf([p('Дата: '), p('Стр. ')], 7);
    s = withField(s, 'DATE \\@ "dd.MM.yyyy"', '01.01.2000');
    s = run(s, s.tr.setSelection(TextSelection.create(s.doc, s.doc.child(0).nodeSize + 6)));
    s = withField(s, 'PAGE', '9');
    s = run(s, s.tr.insert(s.doc.content.size - 1, fieldNodes(schema, 'REF нет_такой \\h', 'старое')));
    const f = scanFields(s.doc);
    expect(f.map((x) => x.instr.trim())).toEqual(['DATE \\@ "dd.MM.yyyy"', 'PAGE', 'REF нет_такой \\h']);
    s = run(s, updateFieldsTr(s, null, new Date(2026, 9, 11)));
    expect(s.doc.child(0).textContent).toBe('Дата: 11.10.2026');
    expect(s.doc.child(1).textContent).toBe(`Стр. 1${REF_ERROR}`);
    expect(updateFieldsTr(s, null, new Date(2026, 9, 11))).toBe(null); // уже актуально — без изменений
  });

  it('закладки и перекрёстные ссылки на закладку и заголовок', () => {
    let s = stateOf([h(1, 'Введение'), p('Срок поставки — 10 дней.'), p('См. ')]);
    // закладка на «10 дней»
    const para = s.doc.child(0).nodeSize + 1;
    s = run(s, s.tr.setSelection(TextSelection.create(s.doc, para + 16, para + 23)));
    s = run(s, addBookmarkTr(s, 'Срок'));
    expect([...bookmarks(s.doc).keys()]).toEqual(['Срок']);
    expect(s.doc.child(1).textContent).toBe('Срок поставки — 10 дней.');
    // ссылка на текст закладки
    s = run(s, s.tr.setSelection(TextSelection.create(s.doc, s.doc.content.size - 1)));
    s = run(s, insertCrossRefTr(s, { bookmark: 'Срок' }, 'text', true));
    expect(s.doc.child(2).textContent).toBe('См. 10 дней');
    // ссылка на заголовок: закладка _Ref вокруг текста заголовка
    s = run(s, insertCrossRefTr(s, { heading: 0 }, 'text', true));
    expect(s.doc.child(2).textContent).toBe('См. 10 днейВведение');
    const ref = [...bookmarks(s.doc).keys()].find((n) => n.startsWith('_Ref'));
    expect(ref).toMatch(/^_Ref\d{9}$/);
    expect(headingList(s.doc)[0].ref).toBe(ref);
    // правка текста закладки → F9
    const b = bookmarks(s.doc).get('Срок');
    s = run(s, s.tr.insertText('15', b.start + 1, b.start + 3));
    s = run(s, updateFieldsTr(s));
    expect(s.doc.child(2).textContent).toBe('См. 15 днейВведение');
    s = run(s, removeBookmarkTr(s, 'Срок'));
    expect(bookmarks(s.doc).has('Срок')).toBe(false);
  });

  it('дата: полем или текстом', () => {
    const s = stateOf([p('')]);
    const a = run(s, insertDateTr(s, 'dd.MM.yyyy', true, new Date(2026, 0, 2)));
    expect(scanFields(a.doc)[0].instr.trim()).toBe('DATE \\@ "dd.MM.yyyy"');
    expect(a.doc.textContent).toBe('02.01.2026');
    const b = run(s, insertDateTr(s, 'H:mm', true, new Date(2026, 0, 2, 7, 3)));
    expect(scanFields(b.doc)[0].instr.trim()).toBe('TIME \\@ "H:mm"');
  });
});

describe('оглавление', () => {
  const docWithHeads = () => stateOf([p(''), h(1, 'Глава 1'), p('текст'), h(2, 'Раздел 1.1'), h(4, 'Глубже'), h(1, 'Глава 2')]);

  it('параметры поля', () => {
    expect(tocOptions('TOC \\o "1-3" \\h \\z \\u')).toEqual({ from: 1, to: 3, links: true });
    expect(tocOptions('TOC \\o "2-2"')).toEqual({ from: 2, to: 2, links: false });
  });

  it('вставка: заголовок «Оглавление», пункты со ссылками и PAGEREF, закладки _Toc, стили toc N', () => {
    const s0 = docWithHeads();
    const s = run(s0, insertTocTr(s0));
    const json = s.doc.toJSON();
    expect(s.doc.child(0).textContent).toBe('Оглавление');
    const entries = [1, 2, 3].map((i) => s.doc.child(i));
    expect(entries.map(txt)).toEqual(['Глава 1\t1', 'Раздел 1.1\t1', 'Глава 2\t1']);
    const styles = s.doc.attrs.styles;
    expect(styles.map((x) => x.name)).toEqual(['TOC Heading', 'toc 1', 'toc 2']);
    expect(entries[1].attrs.styleId).toBe('TOC2');
    expect(entries[0].attrs.tabs[0]).toMatchObject({ align: 'right', leader: 'dot' });
    const marks = bookmarks(s.doc);
    const toc = [...marks.keys()].filter((n) => n.startsWith('_Toc'));
    expect(toc.length).toBe(3);
    const fields = scanFields(s.doc);
    expect(fields.map((f) => f.instr.trim().split(' ')[0])).toEqual(['TOC', 'PAGEREF', 'PAGEREF', 'PAGEREF']);
    const link = entries[0].child(3).marks.find((m) => m.type.name === 'link');
    expect(link.attrs.href).toBe(`#${fields[1].instr.trim().split(' ')[1]}`);
    void json;
    // обновление: новый заголовок попадает в оглавление, закладки не дублируются
    const last = s.doc.content.size;
    let s2 = run(s, s.tr.insert(last, schema.nodeFromJSON(h(2, 'Раздел 2.1'))));
    s2 = run(s2, updateFieldsTr(s2));
    expect([1, 2, 3, 4].map((i) => txt(s2.doc.child(i)))).toEqual(['Глава 1\t1', 'Раздел 1.1\t1', 'Глава 2\t1', 'Раздел 2.1\t1']);
    expect([...bookmarks(s2.doc).keys()].filter((n) => n.startsWith('_Toc')).length).toBe(4);
    expect(s2.doc.attrs.styles.length).toBe(3);
    // повторное обновление ничего не добавляет
    const s3 = run(s2, updateFieldsTr(s2));
    expect(s3.doc.childCount).toBe(s2.doc.childCount);
    expect([...bookmarks(s3.doc).keys()].length).toBe(4);
    // удаление — вместе с «Оглавление»
    const s4 = run(s3, removeTocTr(s3));
    expect(s4.doc.child(0).type.name).toBe('heading');
    expect(scanFields(s4.doc).length).toBe(0);
  });

  it('сохранение в docx: поле TOC, гиперссылки на закладки, стили toc; Word-структура читается обратно', async () => {
    const s0 = docWithHeads();
    const s = run(s0, insertTocTr(s0));
    const bytes = await exportDocx(s.doc.toJSON());
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml').async('string');
    expect(xml).toMatch(/<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u <\/w:instrText>/);
    expect(xml).toMatch(/<w:hyperlink w:anchor="_Toc\d{9}">/);
    expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d{9}"\/>/);
    const styles = await zip.file('word/styles.xml').async('string');
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/>');
    expect(styles).toMatch(/w:styleId="TOC1">[\s\S]*?<w:tab w:val="right" w:leader="dot" w:pos="\d+"\/>/);
    const back = (await importDocx(bytes)).doc;
    const st = EditorState.create({ schema, doc: schema.nodeFromJSON(back) });
    expect(scanFields(st.doc).map((f) => f.instr.trim().split(' ')[0])).toEqual(['TOC', 'PAGEREF', 'PAGEREF', 'PAGEREF']);
    expect([...bookmarks(st.doc).keys()].filter((n) => n.startsWith('_Toc')).length).toBe(3);
    // python-docx открывает
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/toc.docx', bytes);
    const out = execFileSync(PY, ['-c', 'import docx,sys; d=docx.Document(sys.argv[1]); print(len(d.paragraphs), d.paragraphs[1].style.name)', 'tests/out/toc.docx']).toString().trim();
    expect(out).toBe('9 toc 1');
  });
});

describe('эталон Word (tests/fixtures/toc.docx, make_toc_fixture.ps1)', () => {
  it('поля и закладки читаются; F9 перестраивает оглавление Word его стилями и сохраняется поверх исходного файла', async () => {
    const { readFileSync } = await import('fs');
    const src = readFileSync('tests/fixtures/toc.docx');
    const { doc } = await importDocx(src);
    let s = EditorState.create({ schema, doc: schema.nodeFromJSON(doc) });
    expect(scanFields(s.doc).map((f) => f.instr.trim())).toEqual([
      'TOC \\o "1-3"', 'PAGEREF _Toc \\h', 'PAGEREF _Toc \\h', 'PAGEREF _Toc \\h',
      'REF Срок \\h', 'PAGEREF Срок \\h', 'PAGE', 'DATE \\@ "dd.MM.yyyy"',
    ].map((x, i) => (i >= 1 && i <= 3 ? expect.stringMatching(/^PAGEREF _Toc\d+ \\h$/) : x)));
    expect([...bookmarks(s.doc).keys()]).toContain('Срок');
    const tocStyle = (s.doc.attrs.styles || []).find((x) => x.name === 'toc 1').id;
    s = s.apply(updateFieldsTr(s, null, new Date(2026, 9, 11)));
    const entries = [0, 1, 2].map((i) => s.doc.child(i));
    expect(entries.map(txt)).toEqual(['Общие положения\t1', 'Предмет договора\t1', 'Ответственность сторон\t1']);
    expect(entries[0].attrs.styleId).toBe(tocStyle);
    expect(s.doc.child(3).type.name).toBe('heading'); // абзац с концом поля Word — слит с последним пунктом
    expect([...bookmarks(s.doc).keys()].filter((n) => n.startsWith('_Toc')).length).toBe(3); // закладки Word переиспользованы
    expect(s.doc.child(s.doc.childCount - 1).textContent).toBe('Стр. 1, дата: 11.10.2026');
    const bytes = await exportDocx(s.doc.toJSON(), { source: src });
    writeFileSync('tests/out/toc_resaved.docx', bytes);
    const out = execFileSync(PY, ['-c', 'import docx,sys; d=docx.Document(sys.argv[1]); print(d.paragraphs[0].style.name, len(d.paragraphs))', 'tests/out/toc_resaved.docx']).toString().trim();
    expect(out).toBe('toc 1 10');
  });
});
