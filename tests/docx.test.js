import { describe, it, expect } from 'vitest';
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
    expect(doc.attrs).toEqual({ fontFamily: 'Cambria', fontSize: 11 });
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
    expect(doc.attrs).toEqual({ fontFamily: 'Times New Roman', fontSize: 14 });
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
