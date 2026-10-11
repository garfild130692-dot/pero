// Автофильтр и сортировка по нескольким столбцам (filter.js, model.setAutoFilter/setColumnFilter/sortRows)
import { describe, it, expect } from 'vitest';
import * as m from '../src/sheet/model.js';
import { columnValues, filterCount } from '../src/sheet/filter.js';

const { newWorkbook, setInput, valueAt } = m;
const DATA = [
  ['Регион', 'Товар', 'Сумма', 'Дата'],
  ['Север', 'Стол', '100', '15.01.2026'],
  ['Юг', 'Стул', '250', '20.02.2026'],
  ['Север', 'Шкаф', '', '03.03.2026'],
  ['Запад', 'Стол', '75', '15.01.2026'],
  ['Юг', 'Стол', '300', '01.04.2026'],
];
function book() {
  let wb = newWorkbook();
  DATA.forEach((row, r) => row.forEach((t, c) => { if (t) wb = setInput(wb, 0, r, c, t); }));
  wb = setInput(wb, 0, 8, 2, '=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;C2:C6)');
  wb = setInput(wb, 0, 9, 2, '=СУММ(C2:C6)');
  return m.setAutoFilter(wb, 0, { r1: 1, c1: 1, r2: 1, c2: 1 }); // одна ячейка — текущая область
}
const hidden = (wb) => Object.entries(wb.sheets[0].rowInfo || {}).filter(([, i]) => i.hidden).map(([r]) => +r);

describe('автофильтр', () => {
  it('текущая область, список значений, пустые, итоги видимых строк', () => {
    let wb = book();
    expect(wb.sheets[0].autoFilter.ref).toEqual({ r1: 0, c1: 0, r2: 5, c2: 3 });
    const { list, blank } = columnValues(wb.sheets[0], wb.sheets[0].autoFilter, 2);
    expect(list.map((x) => x.text)).toEqual(['75', '100', '250', '300']); // числа — по возрастанию
    expect(blank).toBe(true);
    wb = m.setColumnFilter(wb, 0, 0, { values: ['Юг'] });
    expect(hidden(wb)).toEqual([1, 3, 4]);
    expect(valueAt(wb.sheets[0], 8, 2)).toBe(550); // ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9) — только видимые
    expect(valueAt(wb.sheets[0], 9, 2)).toBe(725); // СУММ — все
    expect(filterCount(wb.sheets[0], wb.sheets[0].autoFilter)).toEqual({ shown: 2, total: 5 });
    // второй столбец: список значений — только из строк, прошедших другие условия
    expect(columnValues(wb.sheets[0], wb.sheets[0].autoFilter, 1).list.map((x) => x.text)).toEqual(['Стол', 'Стул']);
    wb = m.setColumnFilter(wb, 0, 1, { values: ['Стол'] });
    expect(hidden(wb)).toEqual([1, 2, 3, 4]);
    // снять условие столбца, затем весь фильтр
    wb = m.setColumnFilter(wb, 0, 0, null);
    expect(hidden(wb)).toEqual([2, 3]);
    wb = m.setAutoFilter(wb, 0, null);
    expect(hidden(wb)).toEqual([]);
    expect(wb.sheets[0].autoFilter).toBeUndefined();
  });
  it('условия: больше/между, текст «содержит», пустые, даты', () => {
    let wb = book();
    wb = m.setColumnFilter(wb, 0, 2, { custom: { and: true, conds: [{ op: 'greaterThanOrEqual', val: '100' }, { op: 'lessThan', val: '300' }] } });
    expect(hidden(wb)).toEqual([3, 4, 5]);
    wb = m.setColumnFilter(wb, 0, 2, { custom: { and: false, conds: [{ op: 'equal', val: '*ка*' }] } });
    expect(hidden(wb)).toEqual([1, 2, 3, 4, 5]); // в столбце сумм текста нет
    wb = m.setColumnFilter(wb, 0, 2, null);
    wb = m.setColumnFilter(wb, 0, 1, { custom: { and: false, conds: [{ op: 'equal', val: 'ст*' }] } });
    expect(hidden(wb)).toEqual([3]);
    wb = m.setColumnFilter(wb, 0, 1, null);
    wb = m.setColumnFilter(wb, 0, 2, { values: [], blank: true });
    expect(hidden(wb)).toEqual([1, 2, 4, 5]);
    wb = m.setColumnFilter(wb, 0, 2, null);
    wb = m.setColumnFilter(wb, 0, 3, { dates: [{ year: 2026, month: 1 }] });
    expect(hidden(wb)).toEqual([2, 3, 5]);
  });
  it('дописанные снизу строки попадают в фильтр при повторном применении; вставка строк сдвигает фильтр', () => {
    let wb = m.setColumnFilter(book(), 0, 0, { values: ['Север'] });
    wb = setInput(wb, 0, 6, 0, 'Юг');
    wb = m.reapplyFilter(wb, 0);
    expect(hidden(wb)).toEqual([2, 4, 5, 6]);
    expect(wb.sheets[0].autoFilter.ref.r2).toBe(6);
    const w2 = m.shiftCells(wb, 0, 'row', 0, 2);
    expect(w2.sheets[0].autoFilter.ref).toEqual({ r1: 2, c1: 0, r2: 8, c2: 3 });
    const w3 = m.shiftCells(wb, 0, 'col', 0, 1);
    expect(w3.sheets[0].autoFilter.ref.c1).toBe(1);
    expect(Object.keys(w3.sheets[0].autoFilter.cols)).toEqual(['0']);
  });
});

describe('сортировка по нескольким столбцам', () => {
  it('регион по возрастанию, затем сумма по убыванию; заголовок на месте; пустые внизу', () => {
    let wb = newWorkbook();
    DATA.forEach((row, r) => row.forEach((t, c) => { if (t) wb = setInput(wb, 0, r, c, t); }));
    wb = m.sortRows(wb, 0, { r1: 0, c1: 0, r2: 5, c2: 3 }, [{ col: 0, desc: false }, { col: 2, desc: true }], { header: true });
    const col = (c) => [0, 1, 2, 3, 4, 5].map((r) => valueAt(wb.sheets[0], r, c));
    expect(col(0)).toEqual(['Регион', 'Запад', 'Север', 'Север', 'Юг', 'Юг']);
    expect(col(2)).toEqual(['Сумма', 75, 100, null, 300, 250]);
  });
});

describe('автофильтр в файле', () => {
  it('фильтр из Excel читается, повторное применение даёт те же строки, запись — как у Excel', async () => {
    const { readFileSync } = await import('node:fs');
    const { importXlsx, exportXlsx } = await import('../src/sheet/xlsx.js');
    const JSZip = (await import('jszip')).default;
    const src = readFileSync(new URL('./fixtures/filter.xlsx', import.meta.url));
    const { wb } = await importXlsx(src);
    const sh = wb.sheets[0];
    expect(sh.autoFilter.ref).toEqual({ r1: 0, c1: 0, r2: 5, c2: 3 });
    expect(sh.autoFilter.cols).toEqual({
      0: { values: ['Юг'], blank: false },
      2: { custom: { and: true, conds: [{ op: 'greaterThanOrEqual', val: '100' }, { op: 'lessThan', val: '300' }] } },
    });
    expect(hidden(wb)).toEqual([1, 3, 4, 5]);
    expect(sh.rowInfo[1].filtered).toBe(true);
    expect(hidden(m.reapplyFilter(wb, 0))).toEqual([1, 3, 4, 5]); // наши условия скрывают то же, что Excel
    // правка условия → сохранение поверх файла: новый autoFilter и имя _FilterDatabase
    const w2 = m.setColumnFilter(wb, 0, 0, { values: ['Север', 'Запад'] });
    const out = await exportXlsx(w2, { source: src });
    const z = await JSZip.loadAsync(out);
    const xml = await z.file('xl/worksheets/sheet1.xml').async('string');
    expect(xml).toContain('<autoFilter ref="A1:D6"><filterColumn colId="0"><filters><filter val="Север"/><filter val="Запад"/></filters></filterColumn>');
    expect(await z.file('xl/workbook.xml').async('string')).toMatch(/<definedName name="_xlnm\._FilterDatabase" localSheetId="0" hidden="1">Лист1!\$A\$1:\$D\$6<\/definedName>/);
    const back = (await importXlsx(out)).wb;
    expect(back.sheets[0].autoFilter.cols[0]).toEqual({ values: ['Север', 'Запад'], blank: false });
  });
});
