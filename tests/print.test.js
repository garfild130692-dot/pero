// Разбиение листа на печатные страницы — как у Excel (tests/make_pages_xlsx_fixture.ps1: HPageBreaks/VPageBreaks
// самого Excel для разных настроек страницы)
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { importXlsx, exportXlsx } from '../src/sheet/xlsx.js';
import { paginate } from '../src/sheet/printPages.js';
import * as m from '../src/sheet/model.js';

const ex = JSON.parse(readFileSync(new URL('./fixtures/pages_excel.json', import.meta.url), 'utf8').replace(/^﻿/, ''));
const load = async (name) => (await importXlsx(readFileSync(new URL(`./fixtures/pages_${name}.xlsx`, import.meta.url)))).wb;

describe('страницы печати', () => {
  for (const name of Object.keys(ex)) {
    it(`разрывы как у Excel: ${name}`, async () => {
      const p = paginate((await load(name)).sheets[0]);
      expect(p.rowBreaks.map((r) => r + 1)).toEqual(ex[name].rows);
      expect(p.colBreaks.map((c) => c + 1)).toEqual(ex[name].cols);
      if (name === 'fitwide') expect(Math.round(p.scale * 100)).toBe(19); // Excel записал в файл scale="19"
    });
  }
  it('порядок страниц, сквозные строки, ручные разрывы в файле и при вставке строк', async () => {
    const wb = await load('titles');
    const p = paginate(wb.sheets[0]);
    expect(p.pages.length).toBe(4 * 6);
    expect(p.pages[1].rows[0]).toBe(45); // вниз, затем поперёк
    expect(p.pages[1].titleRows).toEqual([0, 1]); // сквозные строки на второй странице
    expect(p.pages[0].titleRows).toEqual([]);
    const man = await load('manual');
    expect(man.sheets[0].breaks).toEqual({ rows: [29], cols: [4] });
    let w = m.setPageBreak(man, 0, 59, 0);
    expect(w.sheets[0].breaks.rows).toEqual([29, 59]);
    w = m.shiftCells(w, 0, 'row', 10, 2);
    expect(w.sheets[0].breaks.rows).toEqual([31, 61]);
    const back = (await importXlsx(await exportXlsx(w))).wb;
    expect(back.sheets[0].breaks).toEqual({ rows: [31, 61], cols: [4] });
    expect(m.setPageBreak(w, 0, 0, 0, 'reset').sheets[0].breaks).toBeUndefined();
  });
});
