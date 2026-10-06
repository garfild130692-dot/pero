import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { execFileSync } from 'child_process';
import { getSchema } from '@tiptap/core';
import { Node } from '@tiptap/pm/model';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { extensions } from '../src/editor/extensions.js';
import { normalizeColor, normalizeFontFamily, parseSize, stepSize } from '../src/editor/font.js';

// На Windows `python3` часто ведёт на заглушку/Python из Store без пакетов — берём `python`.
// Переопределяется переменной окружения PYTHON.
const PYTHON = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3');

const schema = getSchema(extensions);
const text = (doc) => Node.fromJSON(schema, doc).textBetween(0, Node.fromJSON(schema, doc).content.size, '\n', ' ');

describe('docx', () => {
  it('импорт даёт валидный документ редактора', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/sample.docx'));
    const node = Node.fromJSON(schema, doc);
    node.check(); // бросит исключение, если структура не соответствует схеме
    const t = text(doc);
    for (const s of ['Отчёт по продажам', 'жирный', 'Пункт B', 'Вложенный B.1', 'Шаг 2', 'R2C0', 'Конец документа']) {
      expect(t).toContain(s);
    }
  });

  it('round-trip: импорт → экспорт → импорт сохраняет текст и структуру', async () => {
    const a = (await importDocx(readFileSync('tests/fixtures/sample.docx'))).doc;
    const bytes = await exportDocx(a);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/roundtrip.docx', bytes);
    const b = (await importDocx(bytes)).doc;
    Node.fromJSON(schema, b).check();
    expect(b).toEqual(a);
  });

  it('результат открывает сторонняя библиотека (python-docx)', async () => {
    const out = execFileSync(PYTHON, ['-c', `
import docx, sys
d = docx.Document('tests/out/roundtrip.docx')
print('|'.join(p.text for p in d.paragraphs))
print(len(d.tables), d.tables[0].cell(2,0).text)
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8'); // на Windows stdout иначе в cp1251
    expect(out).toContain('Отчёт по продажам');
    expect(out).toContain('1 R2C0');
  });

  it('шрифт документа по умолчанию берётся из темы оформления (python-docx: minorHAnsi = Cambria)', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/sample.docx'));
    expect(doc.attrs).toMatchObject({ fontFamily: 'Cambria', fontSize: 11 });
  });
});

// Ищем текстовый фрагмент и возвращаем его метки в компактном виде { тип: attrs | true }
function marksOf(doc, text) {
  let found = null;
  const walk = (n) => {
    if (found) return;
    if (n.type === 'text' && n.text === text) found = n;
    (n.content || []).forEach(walk);
  };
  walk(doc);
  if (!found) throw new Error(`нет фрагмента «${text}»`);
  return Object.fromEntries((found.marks || []).map((m) => [m.type, m.attrs || true]));
}

describe('шрифт', () => {
  const ts = (a) => ({ fontFamily: null, fontSize: null, color: null, ...a });

  it('импорт: умолчания документа, прямое форматирование, стили', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/fonts.docx'));
    Node.fromJSON(schema, doc).check();
    expect(doc.attrs).toMatchObject({ fontFamily: 'Times New Roman', fontSize: 14 });
    expect(marksOf(doc, 'Обычный, ')).toEqual({});
    expect(marksOf(doc, 'Arial 12')).toEqual({ textStyle: ts({ fontFamily: 'Arial', fontSize: 12 }) });
    expect(marksOf(doc, 'красный')).toEqual({ textStyle: ts({ color: '#c00000' }) });
    expect(marksOf(doc, 'жёлтый')).toEqual({ highlight: { color: '#ffff00' } });
    expect(marksOf(doc, 'заливка')).toEqual({ highlight: { color: '#ffc000' } }); // w:shd
    expect(marksOf(doc, '10,5')).toEqual({ textStyle: ts({ fontSize: 10.5 }) });
    expect(text(doc)).toContain('O, как обычный, '); // Times New Roman 14 явно = умолчание → без меток
    expect(marksOf(doc, 'print("Перо")')).toEqual({ textStyle: ts({ fontFamily: 'Courier New', fontSize: 10 }) }); // стиль абзаца
    expect(marksOf(doc, 'акцент')).toEqual({ bold: true, textStyle: ts({ color: '#1f6f5c' }) }); // символьный стиль
    const sup = doc.content[1].content.findIndex((n) => n.text === ', x') + 1;
    expect(doc.content[1].content[sup].marks).toEqual([{ type: 'superscript' }]);
    const sub = doc.content[1].content.findIndex((n) => n.text === ', H') + 1;
    expect(doc.content[1].content[sub].marks).toEqual([{ type: 'subscript' }]);
  });

  it('round-trip сохраняет шрифты, цвета и выделение', async () => {
    const a = (await importDocx(readFileSync('tests/fixtures/fonts.docx'))).doc;
    const bytes = await exportDocx(a);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/fonts.docx', bytes);
    const b = (await importDocx(bytes)).doc;
    Node.fromJSON(schema, b).check();
    expect(b).toEqual(a);
  });

  it('python-docx видит шрифты, размеры, цвета и умолчания документа', () => {
    const out = execFileSync(PYTHON, ['-c', `
import docx
d = docx.Document('tests/out/fonts.docx')
x = d.styles.element.xpath
print('default', x('.//w:rPrDefault//w:rFonts/@w:ascii')[0], x('.//w:rPrDefault//w:sz/@w:val')[0])
print('shd', d.element.xpath('//w:r[w:t="заливка"]/w:rPr/w:shd/@w:fill'))
for p in d.paragraphs:
    for r in p.runs:
        f = r.font
        print(repr(r.text), f.name, f.size.pt if f.size else None, f.color.rgb if f.color and f.color.type else None,
              f.highlight_color, f.superscript, f.subscript)
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out).toContain('default Times New Roman 28');
    expect(out).toContain("'Arial 12' Arial 12.0");
    expect(out).toContain("'красный' None None C00000");
    expect(out).toMatch(/'жёлтый' None None None YELLOW/);
    expect(out).toContain("shd ['FFC000']");
    expect(out).toContain("'2' None None None None True False");
    expect(out).toContain("'2' None None None None False True");
    expect(out).toContain("'10,5' None 10.5");
  });

  it('нормализация значений', () => {
    expect(normalizeColor('rgb(192, 0, 0)')).toBe('#c00000');
    expect(normalizeColor('#ABC')).toBe('#aabbcc');
    expect(normalizeColor('rgba(0,0,0,0)')).toBe(null);
    expect(normalizeFontFamily('"Times New Roman", serif')).toBe('Times New Roman');
    expect(parseSize('18.67px')).toBe(14);
    expect(parseSize('10,5')).toBe(10.5);
    expect([stepSize(11, 1), stepSize(11, -1), stepSize(72, 1), stepSize(80, -1), stepSize(8, -1)]).toEqual([12, 10.5, 80, 72, 7]);
  });
});

// Атрибуты абзаца с указанным текстом (только заданные)
function paraOf(doc, text) {
  let found = null;
  const walk = (n) => {
    if (found) return;
    if ((n.type === 'paragraph' || n.type === 'heading') && (n.content || []).map((t) => t.text).join('') === text) found = n;
    (n.content || []).forEach(walk);
  };
  walk(doc);
  if (!found) throw new Error(`нет абзаца «${text}»`);
  return Object.fromEntries(Object.entries(found.attrs).filter(([k, v]) => v != null && !['textAlign', 'level'].includes(k)));
}

describe('абзац', () => {
  it('импорт: умолчания из «Обычного», прямое форматирование, стиль, список', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/para.docx'));
    Node.fromJSON(schema, doc).check();
    expect(doc.attrs.para).toEqual({
      spaceBefore: 0, spaceAfter: 0, lineRule: 'auto', lineSpacing: 1.5, indentLeft: 0, indentRight: 0, indentFirst: 35.45,
    });
    expect(paraOf(doc, 'Обычный абзац по умолчанию.')).toEqual({});
    expect(paraOf(doc, 'Перед 12, после 6.')).toEqual({ spaceBefore: 12, spaceAfter: 6 });
    expect(paraOf(doc, 'Одинарный.')).toEqual({ lineRule: 'auto', lineSpacing: 1 });
    expect(paraOf(doc, 'Двойной.')).toEqual({ lineRule: 'auto', lineSpacing: 2 });
    expect(paraOf(doc, 'Точно 18 pt.')).toEqual({ lineRule: 'exact', lineSpacing: 18 });
    expect(paraOf(doc, 'Минимум 20 pt.')).toEqual({ lineRule: 'atLeast', lineSpacing: 20 });
    expect(paraOf(doc, 'Слева 2 см, справа 1 см.')).toEqual({ indentLeft: 56.7, indentRight: 28.35 });
    expect(paraOf(doc, 'Выступ 1 см при отступе слева 1 см.')).toEqual({ indentLeft: 28.35, indentFirst: -28.35 });
    expect(paraOf(doc, 'Без отступа первой строки.')).toEqual({ indentFirst: 0 });
    expect(paraOf(doc, 'Абзац пользовательского стиля.')).toEqual({ spaceBefore: 6, indentLeft: 85.05 });
    expect(paraOf(doc, 'Пункт списка с отступом')).toEqual({}); // отступы элемента списка задаёт список
  });

  it('round-trip сохраняет интервалы и отступы', async () => {
    const a = (await importDocx(readFileSync('tests/fixtures/para.docx'))).doc;
    const bytes = await exportDocx(a);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/para.docx', bytes);
    const b = (await importDocx(bytes)).doc;
    Node.fromJSON(schema, b).check();
    expect(b).toEqual(a);
  });

  it('python-docx видит интервалы, отступы и умолчания документа', () => {
    const out = execFileSync(PYTHON, ['-c', `
import docx
d = docx.Document('tests/out/para.docx')
x = d.styles.element.xpath
print('default', x('.//w:pPrDefault//w:spacing/@w:line')[0], x('.//w:pPrDefault//w:ind/@w:firstLine')[0])
for p in d.paragraphs:
    f = p.paragraph_format
    pt = lambda v: None if v is None else round(v.pt, 2)
    ls = f.line_spacing
    print(repr(p.text), pt(f.space_before), pt(f.space_after), ls if isinstance(ls, float) or ls is None else pt(ls),
          f.line_spacing_rule, pt(f.left_indent), pt(f.right_indent), pt(f.first_line_indent))
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out).toContain('default 360 709');
    expect(out).toContain("'Перед 12, после 6.' 12.0 6.0");
    expect(out).toMatch(/'Точно 18 pt\.' None None 18\.0 EXACTLY/);
    expect(out).toMatch(/'Минимум 20 pt\.' None None 20\.0 AT_LEAST/);
    expect(out).toMatch(/'Двойной\.' None None 2\.0 DOUBLE/);
    expect(out).toContain("'Слева 2 см, справа 1 см.' None None None None 56.7 28.35 None");
    expect(out).toContain("'Выступ 1 см при отступе слева 1 см.' None None None None 28.35 None -28.35");
    expect(out).toContain("'Без отступа первой строки.' None None None None None None 0.0");
  });

  it('новый документ: умолчания как у Word 2013+', async () => {
    const bytes = await exportDocx({ type: 'doc', content: [{ type: 'paragraph' }] });
    const xml = await (await JSZip.loadAsync(bytes)).file('word/styles.xml').async('string');
    expect(xml).toMatch(/<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="160" w:line="259" w:lineRule="auto"\/><w:ind w:left="0" w:right="0" w:firstLine="0"\/><\/w:pPr><\/w:pPrDefault>/);
  });
});

describe('таблицы', () => {
  const tables = (doc) => doc.content.filter((n) => n.type === 'table');
  const cell = (t, r, c) => t.content[r].content[c].attrs;
  const line = (style, width, color = null) => ({ style, width, color });

  it('импорт: ширины, границы (прямые, из стиля, «нет»), заливка, выравнивание, строки', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/tables.docx'));
    Node.fromJSON(schema, doc).check();
    const [t1, t2, t3, t4] = tables(doc);
    // 1) «Сетка таблицы»
    expect(t1.attrs.align).toBe('center');
    expect(t1.attrs.borders.insideV).toEqual(line('single', 0.5));
    expect(t1.content.map((r) => r.content[0].attrs.colwidth)).toEqual([[113.4], [113.4], [113.4]]); // 3 см
    expect(cell(t1, 0, 2).colwidth[0] * 15).toBeCloseTo(4535, 5); // 8 см в твипах
    expect(t1.content[0].attrs).toEqual({ height: null, heightRule: null, repeatHeader: true });
    expect(t1.content[1].attrs).toEqual({ height: 28.35, heightRule: 'exact', repeatHeader: false });
    expect(cell(t1, 0, 0).background).toBe('#dcefe9');
    expect([cell(t1, 1, 1).verticalAlign, cell(t1, 1, 2).verticalAlign]).toEqual(['center', 'bottom']);
    expect(cell(t1, 2, 2).borders).toEqual({ bottom: line('single', 3, '#c00000') });
    // 2) без стиля — границ нет
    expect(Object.values(t2.attrs.borders).every((b) => b.style === 'none')).toBe(true);
    // 3) прямые tblBorders
    expect(t3.attrs.borders.top).toEqual(line('double', 0.5, '#1f6f5c'));
    expect(t3.attrs.borders.insideH).toEqual(line('dashed', 1));
    // 4) границы из стиля таблицы
    expect(t4.attrs.borders.insideH).toEqual(line('single', 1, '#4f81bd'));
  });

  it('round-trip сохраняет оформление таблиц', async () => {
    const a = (await importDocx(readFileSync('tests/fixtures/tables.docx'))).doc;
    const bytes = await exportDocx(a);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/tables.docx', bytes);
    const b = (await importDocx(bytes)).doc;
    Node.fromJSON(schema, b).check();
    expect(b).toEqual(a);
  });

  it('python-docx видит ширины, высоту строк, заголовок и положение таблицы', () => {
    const out = execFileSync(PYTHON, ['-c', `
import docx
d = docx.Document('tests/out/tables.docx')
t = d.tables[0]
print('widths', [round(c.width.cm, 2) for c in t.columns])
print('align', t.alignment)
print('row1', round(t.rows[1].height.cm, 2), t.rows[1].height_rule)
print('header', bool(t.rows[0]._tr.xpath('./w:trPr/w:tblHeader')))
print('valign', t.cell(1, 1).vertical_alignment, t.cell(1, 2).vertical_alignment)
print('shd', t.cell(0, 0)._tc.xpath('./w:tcPr/w:shd/@w:fill'))
print('t2', d.tables[1]._tbl.xpath('./w:tblPr/w:tblBorders/w:top/@w:val'))
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out).toContain('widths [3.0, 5.0, 8.0]');
    expect(out).toContain('align CENTER');
    expect(out).toContain('row1 1.0 EXACTLY');
    expect(out).toContain('header True');
    expect(out).toContain('valign CENTER (1) BOTTOM (3)');
    expect(out).toContain("shd ['DCEFE9']");
    expect(out).toContain("t2 ['nil']");
  });

  it('новая таблица — «Сетка таблицы»', async () => {
    const table = { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', content: [{ type: 'paragraph' }] }] }] };
    const doc = Node.fromJSON(schema, { type: 'doc', content: [table] }).toJSON();
    expect(doc.content[0].attrs.borders.top).toEqual(line('single', 0.5));
    const xml = await (await JSZip.loadAsync(await exportDocx(doc))).file('word/document.xml').async('string');
    expect(xml).toContain('<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
  });
});

describe('сохранность исходного файла', () => {
  const SRC = 'tests/fixtures/preserve.docx'; // создан самим Word (tests/make_preserve_fixture.ps1)
  const kinds = (doc) => {
    const out = [];
    const walk = (n) => { if (n.type === 'opaque') out.push(n.attrs.kind); (n.content || []).forEach(walk); };
    walk(doc);
    return out;
  };
  const files = async (bytes) => {
    const z = await JSZip.loadAsync(bytes);
    const out = {};
    for (const n of Object.keys(z.files)) if (!z.files[n].dir) out[n] = await z.file(n).async('string');
    return out;
  };

  it('импорт: непонятное сохраняется значками, разделы — исходным XML, потерь нет', async () => {
    const { doc, warnings, preserved } = await importDocx(readFileSync(SRC));
    Node.fromJSON(schema, doc).check();
    expect(warnings).toEqual({});
    expect(kinds(doc)).toEqual(['footnote', 'commentRange', 'commentRange', 'comment', 'bookmark', 'bookmark',
      'field', 'field', 'field', 'field', 'math', 'symbol', 'pageBreak']);
    expect(preserved.header).toBe(1);
    expect(doc.attrs.sectPr).toMatch(/^<w:sectPr\b[\s\S]*<w:pgSz\b[^>]*w:orient="landscape"/);
    const sect = doc.content.find((n) => n.attrs.sectPr);
    expect(sect.content[0].text).toBe('Конец первого раздела');
    expect(sect.attrs.sectPr).toContain('w:headerReference');
    // результат поля — обычный редактируемый текст между частями поля
    expect(text(doc)).toContain('06.10.2026');
  });

  it('сохранение с исходным пакетом: всё, чего Перо не касается, — байт в байт', async () => {
    const src = readFileSync(SRC);
    const a = (await importDocx(src)).doc;
    const bytes = await exportDocx(a, { source: src });
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/preserve.docx', bytes);
    const b = await importDocx(bytes);
    expect(b.doc).toEqual(a);
    expect(b.warnings).toEqual({});
    const before = await files(src);
    const after = await files(bytes);
    const touched = ['word/document.xml', 'word/styles.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml', 'docProps/core.xml'];
    for (const [name, content] of Object.entries(before)) {
      if (!touched.includes(name)) expect(after[name], name).toBe(content);
    }
    // корневой тег документа — исходный (все пространства имён и mc:Ignorable)
    expect(after['word/document.xml'].match(/<w:document\b[^>]*>/)[0]).toBe(before['word/document.xml'].match(/<w:document\b[^>]*>/)[0]);
    // исходные связи (колонтитулы, сноски, картинка…) на месте
    for (const id of before['word/_rels/document.xml.rels'].match(/Id="rId\d+"/g)) expect(after['word/_rels/document.xml.rels']).toContain(id);
  });

  it('python-docx видит разделы, колонтитул и рисунок после сохранения', () => {
    const out = execFileSync(PYTHON, ['-c', `
import docx
d = docx.Document('tests/out/preserve.docx')
print('sections', len(d.sections), d.sections[1].orientation)
print('header', d.sections[0].header.paragraphs[0].text)
print('shapes', len(d.inline_shapes))
print('margin', round(d.sections[0].left_margin.cm, 2))
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out).toContain('sections 2 LANDSCAPE (1)');
    expect(out).toContain('header Верхний колонтитул Перо');
    expect(out).toContain('shapes 1');
    expect(out).toContain('margin 3.0');
  });

  it('правки в документе со своими стилями: заголовок и ссылка — на стили файла, недостающие дописываются', async () => {
    const src = readFileSync(SRC);
    const a = (await importDocx(src)).doc;
    a.content.unshift(
      { type: 'heading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Новый заголовок' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'сайт', marks: [{ type: 'link', attrs: { href: 'https://example.org' } }] }] },
    );
    const f = await files(await exportDocx(a, { source: src }));
    const doc = f['word/document.xml'];
    expect(doc).toContain('<w:rStyle w:val="af"/>');           // «Hyperlink» в этом файле — af
    expect(doc).toContain('<w:pStyle w:val="Heading2"/>');     // заголовка 2 не было — дописан
    expect(f['word/styles.xml']).toMatch(/<w:style [^>]*w:styleId="Heading2"[\s\S]*?<w:basedOn w:val="a"\/>/); // на «Обычный» файла
    expect(f['word/_rels/document.xml.rels']).toContain('Target="https://example.org" TargetMode="External"');
    const b = (await importDocx(await exportDocx(a, { source: src }))).doc;
    expect(b.content[0]).toMatchObject({ type: 'heading', attrs: { level: 2 } });
  });

  it('нумерация: наши списки добавляются к исходной с непересекающимися номерами', async () => {
    const src = readFileSync('tests/fixtures/sample.docx');
    const a = (await importDocx(src)).doc;
    const before = await files(src);
    const f = await files(await exportDocx(a, { source: src }));
    const count = (xml, re) => (xml.match(re) || []).length;
    const origAbs = count(before['word/numbering.xml'], /<w:abstractNum\b/g);
    const origNums = count(before['word/numbering.xml'], /<w:num\b/g);
    expect(count(f['word/numbering.xml'], /<w:abstractNum\b/g)).toBe(origAbs + 2);
    expect(count(f['word/numbering.xml'], /<w:num\b/g)).toBeGreaterThan(origNums);
    // все abstractNum — до первого num (порядок схемы)
    const n = f['word/numbering.xml'];
    expect(n.lastIndexOf('<w:abstractNum ')).toBeLessThan(n.search(/<w:num[\s>]/));
    const b = (await importDocx(await exportDocx(a, { source: src }))).doc;
    expect(b).toEqual(a);
  });
});

describe('рисунки', () => {
  const PNG = `data:image/png;base64,${readFileSync('build/icon.png').toString('base64')}`; // 512×512
  const images = (doc) => {
    const out = [];
    const walk = (n) => { if (n.type === 'image') out.push(n); (n.content || []).forEach(walk); };
    walk(doc);
    return out;
  };
  const zipText = async (bytes, name) => (await JSZip.loadAsync(bytes)).file(name)?.async('string');

  it('импорт: рисунок из файла Word — узел image с картинкой, размером и исходной разметкой', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/preserve.docx'));
    const [img] = images(doc);
    expect(img.attrs.src).toMatch(/^data:image\/png;base64,iVBOR/);
    expect([img.attrs.width, img.attrs.height]).toEqual([512, 512]);
    expect(img.attrs.target).toBe('word/media/image1.png');
    expect(img.attrs.xml).toContain('<w:drawing');
    expect(img.attrs.floating).toBe(false);
  });

  it('изменение размера рисунка из файла: исходная разметка, новый размер, та же картинка', async () => {
    const src = readFileSync('tests/fixtures/preserve.docx');
    const { doc } = await importDocx(src);
    const [img] = images(doc);
    Object.assign(img.attrs, { width: 256, height: 128, alt: 'Значок Пера' });
    const bytes = await exportDocx(doc, { source: src });
    writeFileSync('tests/out/image-resized.docx', bytes);
    const xml = await zipText(bytes, 'word/document.xml');
    expect(xml).toContain(`cx="${256 * 9525}" cy="${128 * 9525}"`);
    expect(xml).toContain('descr="Значок Пера"');
    expect(xml).not.toContain('rIdPeroImg'); // новая картинка не создавалась
    const [again] = images((await importDocx(bytes)).doc);
    expect([again.attrs.width, again.attrs.height, again.attrs.alt]).toEqual([256, 128, 'Значок Пера']);
  });

  it('новый рисунок в новом документе: файл в word/media, связь, тип содержимого', async () => {
    const doc = { type: 'doc', content: [{ type: 'paragraph', content: [
      { type: 'text', text: 'Схема: ' }, { type: 'image', attrs: { src: PNG, width: 200, height: 200, alt: 'Схема' } },
    ] }] };
    const bytes = await exportDocx(doc);
    writeFileSync('tests/out/image-new.docx', bytes);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file('word/media/pero_image1.png')).toBeTruthy();
    expect(await zip.file('[Content_Types].xml').async('string')).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(await zip.file('word/_rels/document.xml.rels').async('string')).toContain('Target="media/pero_image1.png"');
    const [img] = images((await importDocx(bytes)).doc);
    expect([img.attrs.width, img.attrs.height, img.attrs.alt]).toEqual([200, 200, 'Схема']);
    const out = execFileSync(PYTHON, ['-c', `
import docx
d = docx.Document('tests/out/image-new.docx')
s = d.inline_shapes[0]
print(len(d.inline_shapes), round(s.width.cm, 2), round(s.height.cm, 2))
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out.trim()).toBe('1 5.29 5.29'); // 200 px = 5,29 см
  });

  it('рисунок из другого документа в документе с исходным пакетом — новая картинка, исходные не трогаются', async () => {
    const [foreign] = images((await importDocx(readFileSync('tests/fixtures/preserve.docx'))).doc);
    const src = readFileSync('tests/fixtures/sample.docx');
    const { doc } = await importDocx(src);
    doc.content.push({ type: 'paragraph', content: [foreign] });
    const bytes = await exportDocx(doc, { source: src });
    const xml = await zipText(bytes, 'word/document.xml');
    expect(xml).toContain('r:embed="rIdPeroImg1"');
    const [img] = images((await importDocx(bytes)).doc);
    expect(img.attrs.target).toBe('word/media/pero_image1.png');
    // повторное сохранение в получившийся пакет — уже «свой» рисунок, без дублей
    const again = await exportDocx((await importDocx(bytes)).doc, { source: bytes });
    const z = await JSZip.loadAsync(again);
    const names = Object.keys(z.files).filter((n) => n.includes('media/') && !z.files[n].dir);
    expect(names).toEqual(['word/media/pero_image1.png']);
  });
});
