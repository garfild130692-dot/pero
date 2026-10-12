// Скорость на больших файлах (этап 17): ловит возврат квадратичных мест и полных проходов. Пределы — с запасом ~3×
// от замеров на машине разработки (нагрузка машины меняет цифры в разы — тест про порядок, а не про миллисекунды).
// Подробные замеры: tests/perf_check.mjs (документ в окне Перо), tests/perf_book.mjs (книга на 100 000 строк).
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { importXlsx, exportXlsx } from '../src/sheet/xlsx.js';
import * as m from '../src/sheet/model.js';

const time = async (fn) => { const t = performance.now(); const r = await fn(); return [r, performance.now() - t]; };

describe('скорость на больших файлах', () => {
  it('документ на 234 страницы: разбор и сохранение', async () => {
    const bytes = readFileSync('tests/fixtures/big.docx');
    const [{ doc }, open] = await time(() => importDocx(bytes));
    const [, save] = await time(() => exportDocx(doc, { source: bytes }));
    expect(open).toBeLessThan(3000);
    expect(save).toBeLessThan(3000);
  });

  it('книга на 20 000 строк: открытие без полного пересчёта, правка, правка справочника ВПР, сохранение', async () => {
    // книга средствами модели: число, цена, формулы по строке, ВПР по справочнику, скользящая сумма, итог по столбцу
    const N = 20000;
    let wb = m.addSheet(m.newWorkbook());
    const cells = new Map();
    for (let i = 0; i < N; i++) {
      const r = i + 1;
      cells.set(`${i},0`, { v: i % 50 + 1 });
      cells.set(`${i},1`, { v: (i * 37) % 900 + 10 });
      cells.set(`${i},2`, { f: `A${r}*B${r}`, v: null });
      cells.set(`${i},3`, { f: `VLOOKUP(MOD(A${r},20),Лист2!$A$1:$B$20,2,FALSE)`, v: null });
      cells.set(`${i},4`, { f: `SUM(C${Math.max(1, r - 9)}:C${r})`, v: null });
    }
    cells.set('0,6', { f: 'SUM(C:C)', v: null });
    const ref = new Map();
    for (let i = 0; i < 20; i++) { ref.set(`${i},0`, { v: i }); ref.set(`${i},1`, { v: i * 10 }); }
    wb = { ...wb, sheets: [{ ...wb.sheets[0], cells }, { ...wb.sheets[1], cells: ref }] };
    wb = m.recalc(wb, { full: true });
    const bytes = await exportXlsx(wb);
    const [{ wb: opened }, open] = await time(() => importXlsx(bytes));
    expect(m.recalcInfo.partial).toBe(true); // значения из файла, граф — по ссылкам формул
    const [edited, edit] = await time(() => m.setInput(opened, 0, N / 2, 0, '7'));
    expect(m.recalcInfo.partial).toBe(true);
    expect(edited.sheets[0].cells.get(`${N / 2},2`).v).toBe(7 * ((N / 2 * 37) % 900 + 10));
    const [, refEdit] = await time(() => m.setInput(edited, 1, 3, 1, '1000')); // справочник: тысяча ВПР
    const [, save] = await time(() => exportXlsx(edited, { source: bytes }));
    expect(open).toBeLessThan(6000);
    expect(edit).toBeLessThan(1500);
    expect(refEdit).toBeLessThan(3000);
    expect(save).toBeLessThan(6000);
  }, 60000);
});
