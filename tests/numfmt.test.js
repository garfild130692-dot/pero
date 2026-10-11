// Числовые форматы: показ как у Excel (tests/make_numfmt_fixture.mjs — Excel открывает книгу с кодами и возвращает
// текст ячеек; коды и значения — tests/numfmt_cases.json)
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { formatValue } from '../src/sheet/format.js';
import { formatCode } from '../src/sheet/numfmt.js';

const { values } = JSON.parse(readFileSync(new URL('./numfmt_cases.json', import.meta.url), 'utf8'));
const excel = JSON.parse(readFileSync(new URL('./fixtures/numfmt_excel.json', import.meta.url), 'utf8'));

describe('числовые форматы', () => {
  it('текст ячеек — как у Excel (разделы, условия, валюты, дроби, даты, время, заполнение)', () => {
    const diffs = [];
    for (const { code, texts } of excel) {
      // «####» зависит от ширины столбца, заполнение * — тоже: сравниваются без учёта ширины
      const norm = (x) => { const s = String(x).replace(/ /g, ' '); return /^#+$/.test(s.trim()) ? '#' : code.includes('*') ? s.replace(/ {2,}/g, ' ').trim() : s; };
      values.forEach((v, j) => { const ours = formatValue(v, code); if (norm(ours) !== norm(texts[j])) diffs.push(`${code} / ${v}: Excel «${texts[j]}» Перо «${ours}»`); });
    }
    expect(diffs).toEqual([]);
  });
  it('цвет раздела', () => {
    expect(formatCode(-5, '0;[Red]-0')).toEqual({ text: '-5', color: '#ff0000' });
    expect(formatCode(5, '[Синий]0')).toEqual({ text: '5', color: '#0000ff' });
  });
});
