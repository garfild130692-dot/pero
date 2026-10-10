// Сверка формул Перо с Excel: tests/fixtures/formulas.xlsx создан Excel (make_formulas_fixture.ps1) по formula_cases.json,
// значения в файле посчитаны самим Excel. Перо пересчитывает книгу и сравнивает каждую формулу.
//   совпало        — значение то же (числа — до 15 значащих цифр);
//   нет в Перо     — Перо не знает функцию или запись (#ИМЯ?) — допустимо, пока функция не сделана;
//   расхождение    — Перо считает, но не так, как Excel, — тест падает.
// Отчёт по каждой формуле — tests/out/formulas_report.md.
import { describe, it, expect } from 'vitest';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { importXlsx } from '../src/sheet/xlsx.js';
import { recalc } from '../src/sheet/model.js';
import { isErr } from '../src/sheet/formula.js';

const VOLATILE = 'Летучие (только тип)';
// Известные расхождения — до соответствующей части этапа 12 (по мере работы список сокращается)
const KNOWN = new Set([]);
const show = (v) => (v == null ? '(пусто)' : isErr(v) ? v.error : typeof v === 'string' ? JSON.stringify(v) : String(v));
const kind = (v) => (v == null ? 'empty' : isErr(v) ? 'error' : typeof v);
function same(a, b) {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.error === b.error;
  if (typeof a === 'number' && typeof b === 'number') return a === b || Math.abs(a - b) <= 1e-12 * Math.max(1, Math.abs(b)) * 1000;
  // пустая ячейка-результат в Excel — 0
  if (a == null && b === 0) return true;
  return a === b;
}

export async function compareWithExcel() {
  const { wb } = await importXlsx(readFileSync('tests/fixtures/formulas.xlsx'), { recalc: false });
  const ours = recalc(wb);
  const si = wb.sheets.findIndex((s) => s.name === 'Формулы');
  const excelSheet = wb.sheets[si]; const oursSheet = ours.sheets[si];
  const rows = [];
  for (let r = 1; ; r++) {
    const text = excelSheet.cells.get(`${r},1`);
    if (!text) break;
    const group = excelSheet.cells.get(`${r},0`).v;
    const excel = (excelSheet.cells.get(`${r},2`) || {}).v ?? null;
    const mine = (oursSheet.cells.get(`${r},2`) || {}).v ?? null;
    const rejected = excelSheet.cells.get(`${r},3`);
    let status;
    if (rejected || (isErr(excel) && excel.error === '#NAME?')) status = 'нет в Excel';
    else if (group === VOLATILE) status = kind(mine) === kind(excel) ? 'совпало' : isErr(mine) && mine.error === '#NAME?' ? 'нет в Перо' : 'расхождение';
    else if (same(mine, excel)) status = 'совпало';
    else if (isErr(mine) && mine.error === '#NAME?') status = 'нет в Перо';
    else status = KNOWN.has(text.v) ? 'известное расхождение' : 'расхождение';
    rows.push({ group, text: text.v, excel, mine, status });
  }
  return rows;
}

describe('формулы как в Excel', () => {
  it('значения совпадают с Excel (кроме ещё не сделанных функций)', async () => {
    const rows = await compareWithExcel();
    const count = (st) => rows.filter((x) => x.status === st).length;
    const lines = ['# Сверка формул с Excel', '',
      `Всего ${rows.length}: совпало ${count('совпало')}, нет в Перо ${count('нет в Перо')}, известные расхождения ${count('известное расхождение')}, расхождение ${count('расхождение')}, нет в Excel ${count('нет в Excel')}.`, ''];
    for (const g of [...new Set(rows.map((x) => x.group))]) {
      const list = rows.filter((x) => x.group === g);
      lines.push(`## ${g} — ${list.filter((x) => x.status === 'совпало').length} из ${list.length}`, '', '| Формула | Excel | Перо | |', '|---|---|---|---|');
      for (const x of list) lines.push(`| \`=${x.text.replace(/\|/g, '\\|')}\` | ${show(x.excel)} | ${show(x.mine)} | ${x.status === 'совпало' ? '✓' : x.status} |`);
      lines.push('');
    }
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/formulas_report.md', lines.join('\n'));
    const bad = rows.filter((x) => x.status === 'расхождение').map((x) => `=${x.text}: Excel ${show(x.excel)}, Перо ${show(x.mine)}`);
    expect(bad).toEqual([]);
  });
});
