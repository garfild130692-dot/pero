// Условное форматирование: показ сверяется с тем, что показывает сам Excel (DisplayFormat, tests/make_cf_fixture.ps1),
// правила читаются, пишутся и переживают круг файл → Перо → файл
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { importXlsx, exportXlsx } from '../src/sheet/xlsx.js';
import { cfView } from '../src/sheet/cf.js';
import * as m from '../src/sheet/model.js';

const excel = JSON.parse(readFileSync(new URL('./fixtures/cf_excel.json', import.meta.url), 'utf8').replace(/^﻿/, ''));
// близкие цвета (шкала: Excel округляет промежуточные цвета по-своему) — разница каналов до 3
const near = (a, b) => {
  if (!a || !b) return a === b;
  const p = (c) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16));
  return p(a).every((x, i) => Math.abs(x - p(b)[i]) <= 3);
};

describe('условное форматирование', () => {
  it('показ как у Excel: больше/между, текст содержит, повторы, первые 3, выше среднего, шкала, формула', async () => {
    const { wb } = await importXlsx(readFileSync(new URL('./fixtures/cf.xlsx', import.meta.url)));
    const view = cfView(wb, 0);
    const diffs = [];
    for (const [k, want] of Object.entries(excel)) {
      const got = (view.get(k) || { style: {} }).style;
      const fill = got.fill || null;
      const color = got.color || '#000000';
      const bold = !!got.bold;
      if (!near(fill, want.fill) || !near(color, want.color) || bold !== want.bold) diffs.push(`${k}: Excel ${JSON.stringify(want)} / Перо ${JSON.stringify({ fill, color, bold })}`);
    }
    expect(diffs).toEqual([]);
  });

  it('новые правила: запись в файл, чтение обратно; правка поверх файла Excel сохраняет его правила', async () => {
    let wb = m.newWorkbook();
    for (let r = 0; r < 5; r++) wb = m.setInput(wb, 0, r, 0, String(r * 30));
    wb = m.addCfRule(wb, 0, [{ r1: 0, c1: 0, r2: 4, c2: 0 }], { type: 'cellIs', operator: 'greaterThan', formulas: ['50'], dxf: { fill: '#ffc7ce', color: '#9c0006' } });
    wb = m.addCfRule(wb, 0, [{ r1: 0, c1: 0, r2: 4, c2: 0 }], { type: 'dataBar', bar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '#638ec6' } });
    const v = cfView(wb, 0);
    expect(v.get('2,0').style).toEqual({ fill: '#ffc7ce', color: '#9c0006' });
    expect(v.get('0,0').bar.pct).toBeCloseTo(0.1, 5);
    expect(v.get('4,0').bar.pct).toBeCloseTo(1, 5);
    const back = (await importXlsx(await exportXlsx(wb))).wb;
    const plain = (cf) => cf.map((b) => ({ ranges: b.ranges, rules: b.rules.map(({ raw, edited, ...r }) => r) }));
    expect(plain(back.sheets[0].cf)).toEqual(plain(wb.sheets[0].cf));
    // поверх файла Excel: своё правило дописывается, его правила — исходной разметкой, dxfs — в конец
    const src = readFileSync(new URL('./fixtures/cf.xlsx', import.meta.url));
    let w2 = (await importXlsx(src)).wb;
    w2 = m.addCfRule(w2, 0, [{ r1: 0, c1: 2, r2: 11, c2: 2 }], { type: 'cellIs', operator: 'lessThan', formulas: ['3'], dxf: { fill: '#ff0000' } });
    const out = await exportXlsx(w2, { source: src });
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(out);
    const xml = await z.file('xl/worksheets/sheet1.xml').async('string');
    expect(xml).toContain('<cfRule type="cellIs" dxfId="7" priority="1" operator="lessThan"><formula>3</formula></cfRule>');
    expect(xml).toContain('<cfRule type="expression" dxfId="0" priority="9" stopIfTrue="1"><formula>$A1&gt;$C1*10</formula></cfRule>');
    expect(await z.file('xl/styles.xml').async('string')).toMatch(/<dxfs count="8">[\s\S]*<dxf><fill><patternFill><bgColor rgb="FFFF0000"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/);
    const w3 = (await importXlsx(out)).wb;
    expect(w3.sheets[0].cf.flatMap((b) => b.rules).length).toBe(9);
  });

  it('удаление правил из выделения урезает диапазоны; вставка строк сдвигает диапазоны и формулы', async () => {
    const { wb } = await importXlsx(readFileSync(new URL('./fixtures/cf.xlsx', import.meta.url)));
    const w2 = m.clearCf(wb, 0, { r1: 0, c1: 0, r2: 3, c2: 0 });
    expect(w2.sheets[0].cf[0].ranges).toEqual([{ r1: 4, c1: 0, r2: 11, c2: 0 }]);
    expect(m.clearCf(wb, 0, null).sheets[0].cf).toBeUndefined();
    const w3 = m.shiftCells(wb, 0, 'row', 0, 2);
    const e = w3.sheets[0].cf.find((b) => b.rules[0].type === 'expression');
    expect(e.ranges).toEqual([{ r1: 2, c1: 4, r2: 13, c2: 4 }]);
    expect(e.rules[0].formulas).toEqual(['$A3>$C3*10']);
  });
});

describe('правила из меню', () => {
  it('формулы как у Excel: текст содержит, даты; значения — число, текст, ссылка', async () => {
    const { buildRule, CF_PRESETS } = await import('../src/sheet/cfRules.js');
    const rg = { r1: 1, c1: 2, r2: 9, c2: 2 };
    const d = CF_PRESETS[0][1];
    expect(buildRule('containsText', ['ка'], d, rg, {})).toMatchObject({ type: 'containsText', text: 'ка', formulas: ['NOT(ISERROR(SEARCH("ка",C2)))'] });
    expect(buildRule('timePeriod', [], d, rg, { period: 'today' })).toMatchObject({ timePeriod: 'today', formulas: ['FLOOR(C2,1)=TODAY()'] });
    expect(buildRule('greaterThan', ['1,5'], d, rg, {}).formulas).toEqual(['1.5']);
    expect(buildRule('equal', ['банк'], d, rg, {}).formulas).toEqual(['"банк"']);
    expect(buildRule('between', ['=$B$1', '01.02.2026'], d, rg, {}).formulas).toEqual(['$B$1', '46054']);
    expect(buildRule('topPercent', ['15'], d, rg, {})).toMatchObject({ type: 'top10', rank: 15, percent: true, bottom: false });
  });
});
