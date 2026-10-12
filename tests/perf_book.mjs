// Замеры большой книги в Node (этап 17): разбор, полный пересчёт, правка ячейки, правка справочника ВПР, сохранение.
// Книга — tests/make_bigbook_fixture.py [строк] → tests/out/bigbook.xlsx (100 000 × 10, формулы, ВПР, итоги по столбцам).
// Запуск: node --max-old-space-size=8000 tests/perf_book.mjs [книга.xlsx]  (REF=1 — ещё правка справочника)
import { readFileSync } from 'node:fs';
import { importXlsx, exportXlsx } from '../src/sheet/xlsx.js';
import * as m from '../src/sheet/model.js';
const file = process.argv[2] || 'tests/out/bigbook.xlsx';
const buf = readFileSync(file);
let t = performance.now();
const wbRaw = (await importXlsx(buf, { recalc: false })).wb;
console.log('разбор без пересчёта', Math.round(performance.now() - t), 'мс');
t = performance.now();
let wb = m.recalc(wbRaw, { full: true });
console.log('полный пересчёт', Math.round(performance.now() - t), 'мс', 'итог', wb.sheets[2].cells.get('0,1')?.v, wb.sheets[2].cells.get('2,1')?.v);
for (const [name, r, c, v] of [['правка числа внизу', wb.sheets[0].cells.size / 10 - 1, 3, '7'], ['правка числа вверху (D2)', 1, 3, '9']].concat(process.env.REF ? [['правка справочника', 0, 1, 'Склад X']] : [])) {
  t = performance.now();
  const w2 = name.includes('справочника') ? m.setInput(wb, 1, r, c, v) : m.setInput(wb, 0, r, c, v);
  console.log(name, Math.round(performance.now() - t), 'мс', JSON.stringify(m.recalcInfo || {}));
}
t = performance.now();
const out = await exportXlsx(wb, { source: buf });
console.log('сохранение', Math.round(performance.now() - t), 'мс', Math.round(out.byteLength / 1024), 'КБ');
console.log('память', Math.round(process.memoryUsage().heapUsed / 1e6), 'МБ');
