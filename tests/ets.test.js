// FORECAST.ETS* — сверка с ответами Excel (tests/make_ets_fixture.ps1 → tests/fixtures/ets_excel.json).
// Подбор параметров α, β, γ у Excel не документирован, наш — свой (минимум ошибки прогноза на шаг), поэтому
// сравнивается прогноз (не дальше 5 % от Excel) и найденная сезонность; параметры и показатели точности — нет.
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { newWorkbook, setInput, valueAt } from '../src/sheet/model.js';
import { localize } from '../src/sheet/formula.js';

const excel = JSON.parse(readFileSync(new URL('./fixtures/ets_excel.json', import.meta.url), 'utf8').replace(/^﻿/, ''));
let wb = newWorkbook();
excel.data.forEach((row, i) => {
  wb = setInput(wb, 0, i, 0, String(i + 1));
  row.forEach((v, k) => { wb = setInput(wb, 0, i, k + 1, String(v).replace('.', ',')); });
});
const calc = (f) => { wb = setInput(wb, 0, 50, 0, `=${localize(f.slice(1))}`); return valueAt(wb.sheets[0], 50, 0); };

describe('FORECAST.ETS', () => {
  for (const col of ['B', 'C', 'D', 'E']) {
    it(`ряд ${col}: прогноз в пределах 5 % от Excel`, () => {
      for (const [f, want] of excel[col]) {
        if (!/^=FORECAST\.ETS\(/.test(f)) continue;
        const got = calc(f);
        expect(typeof got, f).toBe('number');
        expect(Math.abs(got - want) / Math.abs(want), f).toBeLessThan(0.05);
      }
    });
  }
  it('сезонность как у Excel (ряд B: Excel — 2, Перо — 4, настоящий сезон ряда)', () => {
    const season = (col) => calc(excel[col][0][0]);
    expect(excel.C[0][1]).toBe(7); expect(season('C')).toBe(7);
    expect(excel.D[0][1]).toBe(12); expect(season('D')).toBe(12);
    expect(excel.E[0][1]).toBe(0); expect(season('E')).toBe(0);
    expect(season('B')).toBe(4);
  });
  it('параметры, шаг, доверительный интервал, ошибки', () => {
    const a = calc('=FORECAST.ETS.STAT(B1:B36,A1:A36,1)');
    expect(a).toBeGreaterThanOrEqual(0.001); expect(a).toBeLessThanOrEqual(0.999);
    expect(calc('=FORECAST.ETS.STAT(B1:B36,A1:A36,8)')).toBe(1);
    expect(calc('=FORECAST.ETS.STAT(B1:B36,A1:A36,9)')).toEqual({ error: '#NUM!' });
    const ci1 = calc('=FORECAST.ETS.CONFINT(37,D1:D36,A1:A36)'); const ci9 = calc('=FORECAST.ETS.CONFINT(45,D1:D36,A1:A36)');
    expect(ci1).toBeGreaterThan(0); expect(ci9).toBeGreaterThan(ci1);
    expect(calc('=FORECAST.ETS(0,B1:B36,A1:A36)')).toEqual({ error: '#NUM!' }); // раньше начала ряда
    expect(calc('=FORECAST.ETS(37,B1:B36,A1:A35)')).toEqual({ error: '#N/A' }); // разная длина
    expect(calc('=FORECAST.ETS(37,B1:B36,A1:A36,30)')).toEqual({ error: '#NUM!' }); // меньше двух сезонов
  });
});
