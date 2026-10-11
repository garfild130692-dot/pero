// Найти и заменить в книге (find.js)
import { describe, it, expect } from 'vitest';
import * as m from '../src/sheet/model.js';
import { findAll, findNext, replaceAll } from '../src/sheet/find.js';

function book() {
  let wb = m.newWorkbook();
  wb = m.addSheet(wb);
  const put = (si, r, c, t) => { wb = m.setInput(wb, si, r, c, t); };
  put(0, 0, 0, 'Север'); put(0, 0, 1, '1234,5'); put(0, 1, 0, 'Юг'); put(0, 1, 1, '=СУММ(B1;10)');
  put(0, 2, 0, 'северо-запад'); put(0, 2, 1, '15.01.2026');
  put(1, 0, 0, 'Север'); put(1, 3, 2, 'СЕВЕР');
  return { ...wb, active: 0 };
}

describe('найти и заменить', () => {
  it('поиск: регистр, целиком, подстановочные знаки, формулы и значения, лист/книга, порядок', () => {
    const wb = book();
    expect(findAll(wb, 'север').map((x) => `${x.r},${x.c}`)).toEqual(['0,0', '2,0']);
    expect(findAll(wb, 'север', { matchCase: true }).length).toBe(1);
    expect(findAll(wb, 'север', { whole: true }).length).toBe(1);
    expect(findAll(wb, 'с*р', { whole: true }).length).toBe(1);
    expect(findAll(wb, 'север', { within: 'book' }).map((x) => `${x.si}:${x.r},${x.c}`)).toEqual(['0:0,0', '0:2,0', '1:0,0', '1:3,2']);
    expect(findAll(wb, 'СУММ').length).toBe(1); // в формулах — текст формулы
    expect(findAll(wb, 'СУММ', { lookIn: 'values' }).length).toBe(0);
    expect(findAll(wb, '1244,5', { lookIn: 'values' }).length).toBe(1); // значение формулы на экране
    expect(findAll(wb, '15.01', { lookIn: 'values' }).length).toBe(1);
    expect(findAll(wb, '~*').length).toBe(0);
    // по столбцам: B раньше, чем A второй строки
    expect(findAll(wb, '?', { by: 'cols' }).map((x) => `${x.r},${x.c}`).slice(0, 3)).toEqual(['0,0', '1,0', '2,0']);
    expect(findNext(wb, 'север', { si: 0, r: 0, c: 0 })).toMatchObject({ r: 2, c: 0 });
    expect(findNext(wb, 'север', { si: 0, r: 2, c: 0 })).toMatchObject({ r: 0, c: 0 }); // по кругу
    expect(findNext(wb, 'север', { si: 0, r: 0, c: 0 }, { back: true })).toMatchObject({ r: 2, c: 0 });
  });
  it('замена: в тексте и формулах, одной операцией, с пересчётом', () => {
    const { wb, count } = replaceAll(book(), 'север', 'Восток', { within: 'book' });
    expect(count).toBe(4);
    expect([m.valueAt(wb.sheets[0], 0, 0), m.valueAt(wb.sheets[0], 2, 0), m.valueAt(wb.sheets[1], 3, 2)]).toEqual(['Восток', 'Востоко-запад', 'Восток']);
    const r2 = replaceAll(book(), '10', '20');
    expect(r2.wb.sheets[0].cells.get('1,1').f).toBe('SUM(B1,20)');
    expect(m.valueAt(r2.wb.sheets[0], 1, 1)).toBe(1254.5);
    const r3 = replaceAll(book(), 'юг', '$1', { whole: true });
    expect(m.valueAt(r3.wb.sheets[0], 1, 0)).toBe('$1');
  });
});
