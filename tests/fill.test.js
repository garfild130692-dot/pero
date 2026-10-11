// Маркер заполнения: ряды как в Excel (fill.js + model.fillSeries)
import { describe, it, expect } from 'vitest';
import { newWorkbook, setInput, setStyle, fillSeries, valueAt } from '../src/sheet/model.js';
import { formatValue } from '../src/sheet/format.js';

// заполнить столбец A значениями src и протянуть вниз до n ячеек; вернуть то, что видно на экране
function down(src, n, opts) {
  let wb = newWorkbook();
  src.forEach((t, i) => { wb = setInput(wb, 0, i, 0, t); });
  wb = fillSeries(wb, 0, { r1: 0, c1: 0, r2: src.length - 1, c2: 0 }, { r1: 0, c1: 0, r2: n - 1, c2: 0 }, opts);
  return Array.from({ length: n }, (_, r) => { const c = wb.sheets[0].cells.get(`${r},0`); return c ? formatValue(c.v, c.s && c.s.numFmt) : null; });
}

describe('маркер заполнения', () => {
  it('числа: одно — копия (Ctrl — ряд), два и больше — линейный тренд', () => {
    expect(down(['5'], 4)).toEqual(['5', '5', '5', '5']);
    expect(down(['5'], 4, { ctrl: true })).toEqual(['5', '6', '7', '8']);
    expect(down(['1', '3'], 5)).toEqual(['1', '3', '5', '7', '9']);
    expect(down(['0,1', '0,2'], 5)).toEqual(['0,1', '0,2', '0,3', '0,4', '0,5']);
    expect(down(['1', '2', '4'], 5)).toEqual(['1', '2', '4', '5,3333333333', '6,8333333333']); // тренд, как Excel
    expect(down(['1', '3'], 4, { ctrl: true })).toEqual(['1', '3', '1', '3']); // Ctrl — копия
  });
  it('даты: одна — по дням, одно число месяца — по месяцам, иначе шаг в днях', () => {
    expect(down(['31.01.2026'], 3)).toEqual(['31.01.2026', '01.02.2026', '02.02.2026']);
    expect(down(['31.01.2026'], 3, { ctrl: true })).toEqual(['31.01.2026', '31.01.2026', '31.01.2026']);
    expect(down(['31.01.2026', '28.02.2026'], 4)).toEqual(['31.01.2026', '28.02.2026', '31.03.2026', '30.04.2026']); // концы месяцев
    expect(down(['31.03.2026', '30.04.2026'], 6)).toEqual(['31.03.2026', '30.04.2026', '31.05.2026', '30.06.2026', '31.07.2026', '31.08.2026']);
    expect(down(['15.01.2026', '15.02.2026'], 5)).toEqual(['15.01.2026', '15.02.2026', '15.03.2026', '15.04.2026', '15.05.2026']);
    expect(down(['01.01.2025', '01.01.2026'], 3)).toEqual(['01.01.2025', '01.01.2026', '01.01.2027']);
    expect(down(['01.01.2026', '08.01.2026'], 3)).toEqual(['01.01.2026', '08.01.2026', '15.01.2026']);
  });
  it('месяцы и дни недели (регистр сохраняется), текст с номером', () => {
    expect(down(['Январь'], 4)).toEqual(['Январь', 'Февраль', 'Март', 'Апрель']);
    expect(down(['ноя'], 4)).toEqual(['ноя', 'дек', 'янв', 'фев']);
    expect(down(['Май'], 2)).toEqual(['Май', 'Июнь']);
    expect(down(['ПН', 'СР'], 4)).toEqual(['ПН', 'СР', 'ПТ', 'ВС']);
    expect(down(['Monday'], 3)).toEqual(['Monday', 'Monday', 'Monday']); // в русском Excel английских списков нет
    expect(down(['Кв 1'], 6)).toEqual(['Кв 1', 'Кв 2', 'Кв 3', 'Кв 4', 'Кв 1', 'Кв 2']);
    expect(down(['2 квартал'], 4)).toEqual(['2 квартал', '3 квартал', '4 квартал', '1 квартал']);
    expect(down(['кв.1'], 3)).toEqual(['кв.1', 'кв.2', 'кв.3']);
    expect(down(['Кв1', 'Кв3'], 4)).toEqual(['Кв1', 'Кв3', 'Кв1', 'Кв3']);
    expect(down(['Q1'], 6)).toEqual(['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6']);
    expect(down(['Отдел 7', 'Отдел 7'], 3)).toEqual(['Отдел 7', 'Отдел 7', 'Отдел 7']);
    expect(down(['янв', 'мар'], 6)).toEqual(['янв', 'мар', 'май', 'июл', 'сен', 'ноя']);
    expect(down(['понедельник'], 3)).toEqual(['понедельник', 'вторник', 'среда']);
    expect(down(['Item001', 'Item003'], 4)).toEqual(['Item001', 'Item003', 'Item005', 'Item007']);
    expect(down(['а', 'б'], 5)).toEqual(['а', 'б', 'а', 'б', 'а']);
    expect(down(['а1', 'б'], 4)).toEqual(['а1', 'б', 'а2', 'б']);
  });
  it('формулы сдвигаются, стили копируются, заполнение вверх и вправо, протягивание назад очищает', () => {
    let wb = newWorkbook();
    wb = setInput(wb, 0, 0, 0, '1'); wb = setInput(wb, 0, 1, 0, '2'); wb = setInput(wb, 0, 2, 0, '3');
    wb = setInput(wb, 0, 0, 1, '=A1*10');
    wb = setStyle(wb, 0, { r1: 0, c1: 1, r2: 0, c2: 1 }, { bold: true });
    wb = fillSeries(wb, 0, { r1: 0, c1: 1, r2: 0, c2: 1 }, { r1: 0, c1: 1, r2: 2, c2: 1 });
    expect([1, 2].map((r) => wb.sheets[0].cells.get(`${r},1`).f)).toEqual(['A2*10', 'A3*10']);
    expect(valueAt(wb.sheets[0], 2, 1)).toBe(30);
    expect(wb.sheets[0].cells.get('2,1').s.bold).toBe(true);
    // вправо: ряд по строке
    let w2 = setInput(setInput(newWorkbook(), 0, 5, 2, '10'), 0, 5, 3, '20');
    w2 = fillSeries(w2, 0, { r1: 5, c1: 2, r2: 5, c2: 3 }, { r1: 5, c1: 2, r2: 5, c2: 6 });
    expect([4, 5, 6].map((c) => valueAt(w2.sheets[0], 5, c))).toEqual([30, 40, 50]);
    // вверх: ряд продолжается назад
    let w3 = setInput(setInput(newWorkbook(), 0, 5, 0, '10'), 0, 6, 0, '20');
    w3 = fillSeries(w3, 0, { r1: 5, c1: 0, r2: 6, c2: 0 }, { r1: 2, c1: 0, r2: 6, c2: 0 });
    expect([2, 3, 4].map((r) => valueAt(w3.sheets[0], r, 0))).toEqual([-20, -10, 0]);
    // назад внутрь выделения — очистка
    const w4 = fillSeries(wb, 0, { r1: 0, c1: 0, r2: 2, c2: 0 }, { r1: 0, c1: 0, r2: 0, c2: 0 });
    expect([valueAt(w4.sheets[0], 0, 0), valueAt(w4.sheets[0], 1, 0), valueAt(w4.sheets[0], 2, 0)]).toEqual([1, null, null]);
  });
});

describe('«Формат ячеек»', () => {
  it('стиль, границы (внешние/внутренние), объединение — одной операцией', async () => {
    const m = await import('../src/sheet/model.js');
    let wb = newWorkbook();
    wb = m.applyCellFormat(wb, 0, { r1: 0, c1: 0, r2: 1, c2: 1 }, {
      patch: { numFmt: '#,##0.00', bold: true, fill: '#ffff00' },
      borders: { top: { style: 'thin', color: '#000000' }, bottom: { style: 'thin', color: '#000000' }, insideH: { style: 'dashed', color: '#ff0000' } },
    });
    const st = (r, c) => wb.sheets[0].cells.get(`${r},${c}`).s;
    expect(st(0, 0)).toMatchObject({ numFmt: '#,##0.00', bold: true, fill: '#ffff00', border: { top: { style: 'thin' }, bottom: { style: 'dashed' } } });
    expect(st(1, 1).border).toEqual({ top: { style: 'dashed', color: '#ff0000' }, bottom: { style: 'thin', color: '#000000' } });
    wb = m.applyCellFormat(wb, 0, { r1: 0, c1: 0, r2: 1, c2: 1 }, { merge: true });
    expect(wb.sheets[0].merges).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 1 }]);
    wb = m.applyCellFormat(wb, 0, { r1: 0, c1: 0, r2: 1, c2: 1 }, { patch: { bold: null }, borders: { top: null } });
    expect(st(0, 0).bold).toBeUndefined();
    expect(st(0, 0).border.top).toBeUndefined();
  });
});
