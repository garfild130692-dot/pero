// Проверка данных: списки, ограничения, сообщения, файл (dv.js, model.setValidation)
import { describe, it, expect } from 'vitest';
import * as m from '../src/sheet/model.js';
import { dvAt, listItems, validate } from '../src/sheet/dv.js';
import { exportXlsx, importXlsx } from '../src/sheet/xlsx.js';

function book() {
  let wb = m.newWorkbook();
  ['Москва', 'Казань', 'Омск'].forEach((t, r) => { wb = m.setInput(wb, 0, r, 5, t); }); // F1:F3 — источник списка
  wb = m.setValidation(wb, 0, { r1: 0, c1: 0, r2: 9, c2: 0 }, { type: 'list', formula1: '"да,нет,может быть"', inCellDropdown: true, allowBlank: true, showError: true, errorStyle: 'stop' });
  wb = m.setValidation(wb, 0, { r1: 0, c1: 1, r2: 9, c2: 1 }, { type: 'list', formula1: '$F$1:$F$3', inCellDropdown: true, allowBlank: true, showError: true });
  wb = m.setValidation(wb, 0, { r1: 0, c1: 2, r2: 9, c2: 2 }, { type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true, showError: true, showPrompt: true, promptTitle: 'Количество', prompt: 'От 1 до 10' });
  wb = m.setValidation(wb, 0, { r1: 0, c1: 3, r2: 9, c2: 3 }, { type: 'textLength', operator: 'lessThanOrEqual', formula1: '5', allowBlank: false, showError: true, errorStyle: 'warning' });
  return wb;
}

describe('проверка данных', () => {
  it('списки и ограничения', () => {
    const wb = book();
    expect(listItems(wb, 0, dvAt(wb.sheets[0], 3, 0), 3, 0)).toEqual(['да', 'нет', 'может быть']);
    expect(listItems(wb, 0, dvAt(wb.sheets[0], 3, 1), 3, 1)).toEqual(['Москва', 'Казань', 'Омск']);
    expect(validate(wb, 0, 0, 0, 'Нет')).toBe(false); // список текстом — с учётом регистра (как Excel)
    expect(validate(wb, 0, 0, 1, 'омск')).toBe(true); // список из диапазона — без
    expect(validate(wb, 0, 0, 0, 'возможно')).toBe(false);
    expect(validate(wb, 0, 0, 1, 'Омск')).toBe(true);
    expect(validate(wb, 0, 0, 2, 5)).toBe(true);
    expect(validate(wb, 0, 0, 2, 5.5)).toBe(false); // целое
    expect(validate(wb, 0, 0, 2, 11)).toBe(false);
    expect(validate(wb, 0, 0, 2, null)).toBe(true); // пустые разрешены
    expect(validate(wb, 0, 0, 3, 'абвгде')).toBe(false);
    expect(validate(wb, 0, 0, 3, null)).toBe(true); // пустую ячейку Excel не проверяет
    expect(validate(wb, 0, 0, 4, 'что угодно')).toBe(true); // нет правила
    // новое правило на часть диапазона вычитается из старого
    const w2 = m.setValidation(wb, 0, { r1: 0, c1: 0, r2: 4, c2: 0 }, null);
    expect(dvAt(w2.sheets[0], 2, 0)).toBeNull();
    expect(dvAt(w2.sheets[0], 6, 0).type).toBe('list');
  });
  it('файл: запись, чтение, Excel-атрибуты; вставка строк сдвигает диапазоны и ссылки', async () => {
    const wb = book();
    const JSZip = (await import('jszip')).default;
    const out = await exportXlsx(wb);
    const xml = await (await JSZip.loadAsync(out)).file('xl/worksheets/sheet1.xml').async('string');
    expect(xml).toContain('<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="A1:A10"><formula1>"да,нет,может быть"</formula1></dataValidation>');
    expect(xml).toContain('<dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Количество" prompt="От 1 до 10" sqref="C1:C10"><formula1>1</formula1><formula2>10</formula2></dataValidation>');
    const back = (await importXlsx(out)).wb;
    const plain = (dv) => dv.map(({ raw, edited, ...r }) => r);
    expect(plain(back.sheets[0].dv)).toEqual(plain(wb.sheets[0].dv));
    const w3 = m.shiftCells(wb, 0, 'row', 0, 2);
    const r = dvAt(w3.sheets[0], 5, 1);
    expect(r.ranges).toEqual([{ r1: 2, c1: 1, r2: 11, c2: 1 }]);
    expect(r.formula1).toBe('$F$3:$F$5');
  });
});
