import { describe, it, expect } from 'vitest';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { execFileSync } from 'child_process';
import { getSchema } from '@tiptap/core';
import { Node } from '@tiptap/pm/model';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { extensions } from '../src/editor/extensions.js';

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
    const out = execFileSync('python3', ['-c', `
import docx, sys
d = docx.Document('tests/out/roundtrip.docx')
print('|'.join(p.text for p in d.paragraphs))
print(len(d.tables), d.tables[0].cell(2,0).text)
`]).toString();
    expect(out).toContain('Отчёт по продажам');
    expect(out).toContain('1 R2C0');
  });
});
