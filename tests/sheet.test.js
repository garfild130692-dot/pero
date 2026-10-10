import { describe, it, expect } from 'vitest';
import { addr, colIndex, colName, parseRange } from '../src/sheet/address.js';
import { canonical, fromUser, localize, offsetFormula, parse, shiftFormula, evaluate, isErr } from '../src/sheet/formula.js';
import { formatValue, parseInput, changeDecimals } from '../src/sheet/format.js';
import {
  newWorkbook, setInput, setStyle, shiftCells, sortRange, mergeRange, copyBlock, pasteBlock, pasteText,
  renameSheet, addSheet, deleteSheet, fill, autoSumFormula,
} from '../src/sheet/model.js';

const val = (wb, a, si = 0) => {
  const { r1, c1 } = parseRange(a);
  const cell = wb.sheets[si].cells.get(`${r1},${c1}`);
  return cell ? cell.v : null;
};
const formulaOf = (wb, a, si = 0) => {
  const { r1, c1 } = parseRange(a);
  return (wb.sheets[si].cells.get(`${r1},${c1}`) || {}).f;
};
const put = (wb, entries, si = 0) => Object.entries(entries).reduce((acc, [a, v]) => {
  const { r1, c1 } = parseRange(a);
  return setInput(acc, si, r1, c1, v);
}, wb);

describe('адреса', () => {
  it('столбцы и ссылки', () => {
    expect([colName(0), colName(25), colName(26), colName(701), colName(702)]).toEqual(['A', 'Z', 'AA', 'ZZ', 'AAA']);
    expect([colIndex('A'), colIndex('AA'), colIndex('XFD')]).toEqual([0, 26, 16383]);
    expect(addr(9, 27)).toBe('AB10');
    expect(parseRange('C3:A1')).toEqual({ r1: 0, c1: 0, r2: 2, c2: 2 });
  });
});

describe('формулы', () => {
  const ev = (f, cells = {}) => evaluate(parse(f), {
    sheet: 'Лист1',
    get: (_s, r, c) => cells[`${r},${c}`] ?? null,
  });
  it('арифметика и приоритеты как в Excel', () => {
    expect(ev('1+2*3')).toBe(7);
    expect(ev('(1+2)*3')).toBe(9);
    expect(ev('-2^2')).toBe(4);       // в Excel унарный минус сильнее степени
    expect(ev('2^3^2')).toBe(64);     // слева направо
    expect(ev('50%*10')).toBe(5);
    expect(ev('"а"&"б"&1')).toBe('аб1');
    expect(ev('1/0')).toEqual({ error: '#DIV/0!' });
    expect(ev('"abc"+1')).toEqual({ error: '#VALUE!' });
    expect(ev('"10"+1')).toBe(11);
    expect(ev('3>2')).toBe(true);
    expect(ev('"Б"="б"')).toBe(true);
  });
  it('функции: агрегаты, логика, текст, условия, поиск, даты', () => {
    const cells = { '0,0': 1, '1,0': 2, '2,0': 3, '3,0': 'x', '0,1': 'яблоко', '1,1': 'груша', '2,1': 'яблоко' };
    expect(ev('SUM(A1:A4)', cells)).toBe(6);
    expect(ev('AVERAGE(A1:A3)', cells)).toBe(2);
    expect(ev('COUNT(A1:A4)', cells)).toBe(3);
    expect(ev('COUNTA(A1:B4)', cells)).toBe(7);
    expect(ev('IF(A1>0,"плюс","минус")', cells)).toBe('плюс');
    expect(ev('IF(FALSE,1/0,5)')).toBe(5);                         // ленивое вычисление
    expect(ev('IFERROR(1/0,"нет")')).toBe('нет');
    expect(ev('ROUND(2.345,2)')).toBe(2.35);
    expect(ev('ROUNDDOWN(-2.7,0)')).toBe(-2);
    expect(ev('MOD(-3,2)')).toBe(1);
    expect(ev('SUMIF(B1:B3,"яблоко",A1:A3)', cells)).toBe(4);
    expect(ev('COUNTIF(A1:A3,">1")', cells)).toBe(2);
    expect(ev('COUNTIF(B1:B3,"гр*")', cells)).toBe(1);
    expect(ev('VLOOKUP("груша",B1:B3,1,FALSE)', cells)).toBe('груша');
    expect(ev('MATCH(3,A1:A3,0)', cells)).toBe(3);
    expect(ev('LEFT("Перо",2)&UPPER("ок")')).toBe('ПеОК');
    expect(ev('YEAR(DATE(2026,10,6))')).toBe(2026);
    expect(ev('НЕИЗВЕСТНАЯ(1)')).toEqual({ error: '#NAME?' });
  });
  it('русские имена функций и «;» приводятся к хранимому виду', () => {
    expect(canonical('СУММ(A1:A3; 5)+ЕСЛИ(ИСТИНА;1;2)')).toBe('SUM(A1:A3, 5)+IF(TRUE,1,2)');
    expect(canonical("'Мой лист'!A1+Лист2!$B$2")).toBe("'Мой лист'!A1+Лист2!$B$2");
    // русская запись в строке формул: десятичная запятая, «;», русские имена — и обратно
    expect(fromUser('ЕСЛИ(A1>1,5;"да, нет";ОКРУГЛ(B1*0,25;2))')).toBe('IF(A1>1.5,"да, нет",ROUND(B1*0.25,2))');
    expect(localize('IF(A1>1.5,"да, нет",ROUND(B1*0.25,2))+COUNT(A:A)')).toBe('ЕСЛИ(A1>1,5;"да, нет";ОКРУГЛ(B1*0,25;2))+СЧЁТ(A:A)');
    expect(fromUser(localize('SUM(A1:A3, 5)*TRUE'))).toBe('SUM(A1:A3, 5)*TRUE');
  });
  it('копирование со смещением и сдвиг при вставке/удалении строк', () => {
    expect(offsetFormula('A1+$B$1+B$2+$C3', 2, 1)).toBe('B3+$B$1+C$2+$C5');
    expect(offsetFormula('A1', -1, 0)).toBe('#REF!');
    const opts = { sheet: 'Лист1', current: 'Лист1', axis: 'row' };
    expect(shiftFormula('SUM(A1:A5)+A6', { ...opts, at: 2, count: 2 })).toBe('SUM(A1:A7)+A8');
    expect(shiftFormula('SUM(A1:A5)+A3', { ...opts, at: 2, count: -1 })).toBe('SUM(A1:A4)+#REF!');
    expect(shiftFormula('Лист2!A5', { ...opts, at: 0, count: 1 })).toBe('Лист2!A5'); // другой лист — не трогаем
  });
});

describe('ввод и форматы', () => {
  it('разбор ввода по-русски', () => {
    expect(parseInput('12,5').value).toBe(12.5);
    expect(parseInput('1 234,56').value).toBe(1234.56);
    expect(parseInput('15%')).toEqual({ value: 0.15, format: '0%' });
    expect(parseInput('06.10.2026')).toMatchObject({ value: 46301, format: 'dd.mm.yyyy' });
    expect(parseInput('12:30').value).toBeCloseTo(0.520833, 5);
    expect(parseInput('ИСТИНА').value).toBe(true);
    expect(parseInput("'007")).toMatchObject({ value: '007', text: true });
    expect(parseInput('текст').value).toBe('текст');
  });
  it('показ значений', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1 234,50');
    expect(formatValue(0.256, '0.0%')).toBe('25,6%');
    expect(formatValue(46301, 'dd.mm.yyyy')).toBe('06.10.2026');
    expect(formatValue(1500, '#,##0.00 "₽"')).toBe('1 500,00 ₽');
    expect(formatValue(0.1 + 0.2)).toBe('0,3');
    expect(formatValue(true)).toBe('ИСТИНА');
    expect(changeDecimals('0.00', 1)).toBe('0.000');
    expect(changeDecimals('General', 1, 3.5)).toBe('0.00');
  });
});

describe('модель книги', () => {
  it('ввод, формулы, пересчёт, циклы', () => {
    let wb = put(newWorkbook(), { A1: '10', A2: '20', A3: '=СУММ(A1:A2)', B1: '=A3*2' });
    expect([val(wb, 'A3'), val(wb, 'B1'), formulaOf(wb, 'A3')]).toEqual([30, 60, 'SUM(A1:A2)']);
    wb = put(wb, { A1: '15' });
    expect([val(wb, 'A3'), val(wb, 'B1')]).toEqual([35, 70]);
    wb = put(wb, { C1: '=C2+1', C2: '=C1+1' });
    expect(isErr(val(wb, 'C1'))).toBe(true);
  });
  it('вставка и удаление строк сдвигают данные и формулы', () => {
    let wb = put(newWorkbook(), { A1: '1', A2: '2', A3: '3', A4: '=SUM(A1:A3)' });
    wb = shiftCells(wb, 0, 'row', 1, 1);   // вставить строку перед 2-й
    expect([val(wb, 'A3'), formulaOf(wb, 'A5'), val(wb, 'A5')]).toEqual([2, 'SUM(A1:A4)', 6]);
    wb = shiftCells(wb, 0, 'row', 0, -1);  // удалить первую
    expect([formulaOf(wb, 'A4'), val(wb, 'A4')]).toEqual(['SUM(A1:A3)', 5]);
  });
  it('сортировка, объединение, буфер, заполнение, автосумма', () => {
    let wb = put(newWorkbook(), { A1: 'груша', B1: '3', A2: 'арбуз', B2: '1', A3: 'яблоко', B3: '2' });
    wb = sortRange(wb, 0, { r1: 0, c1: 0, r2: 2, c2: 1 }, 1, true);
    expect([val(wb, 'A1'), val(wb, 'A2'), val(wb, 'A3')]).toEqual(['арбуз', 'яблоко', 'груша']);
    wb = sortRange(wb, 0, { r1: 0, c1: 0, r2: 2, c2: 1 }, 0, false);
    expect(val(wb, 'A1')).toBe('яблоко');
    wb = mergeRange(wb, 0, { r1: 0, c1: 0, r2: 0, c2: 1 });
    expect(wb.sheets[0].merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 1 }]);
    expect(val(wb, 'B1')).toBe(null);
    let w2 = put(newWorkbook(), { A1: '1', B1: '=A1*10' });
    w2 = pasteBlock(w2, 0, 4, 0, copyBlock(w2.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 1 }));
    expect(formulaOf(w2, 'B5')).toBe('A5*10');
    w2 = pasteText(w2, 0, 9, 0, '5\t=A10*2\r\n7\t\r\n');
    expect([val(w2, 'A10'), val(w2, 'B10'), val(w2, 'A11')]).toEqual([5, 10, 7]);
    w2 = put(w2, { C1: '1', C2: '2', C3: '3' });
    w2 = fill(w2, 0, { r1: 0, c1: 1, r2: 2, c2: 1 }, 'down');
    expect([formulaOf(w2, 'B3'), val(w2, 'B3')]).toEqual(['A3*10', 0]);
    expect(autoSumFormula(w2.sheets[0], 3, 2)).toBe('SUM(C1:C3)');
  });
  it('листы: ссылки между листами, переименование, удаление', () => {
    let wb = addSheet(put(newWorkbook(), { A1: '5' }));
    wb = put(wb, { A1: '=Лист1!A1*2' }, 1);
    expect(val(wb, 'A1', 1)).toBe(10);
    wb = renameSheet(wb, 0, 'Данные за год');
    expect(formulaOf(wb, 'A1', 1)).toBe("'Данные за год'!A1*2");
    expect(val(wb, 'A1', 1)).toBe(10);
    wb = deleteSheet(wb, 0);
    expect(wb.sheets).toHaveLength(1);
    expect(isErr(val(wb, 'A1'))).toBe(true);
    wb = setStyle(wb, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, { bold: true, fill: '#ffff00' });
    expect(wb.sheets[0].cells.get('0,0').s).toEqual({ bold: true, fill: '#ffff00' });
  });
});

describe('xlsx', async () => {
  const { readFileSync, writeFileSync, mkdirSync } = await import('fs');
  const { execFileSync } = await import('child_process');
  const { importXlsx, exportXlsx } = await import('../src/sheet/xlsx.js');
  const PYTHON = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
  const cell = (wb, a, si = 0) => { const { r1, c1 } = parseRange(a); return wb.sheets[si].cells.get(`${r1},${c1}`) || {}; };

  it('импорт книги Excel: значения, формулы (в т.ч. общие), стили, объединение, ширины, листы', async () => {
    const { wb, warnings } = await importXlsx(readFileSync('tests/fixtures/book.xlsx'));
    expect(warnings).toEqual({});
    expect(wb.sheets.map((s) => s.name)).toEqual(['Продажи', 'Свод']);
    expect(cell(wb, 'A2').v).toBe('Ручка');
    expect(cell(wb, 'D4')).toMatchObject({ f: 'B4*C4', v: 305 });                   // общая формула развёрнута
    expect(cell(wb, 'D6')).toMatchObject({ f: 'SUM(D2:D5)', v: 1120 }); // 455 + 360 + 305 + 0
    expect(cell(wb, 'C7').v).toBe(3);
    expect(cell(wb, 'D7').v).toBe('много');
    expect(cell(wb, 'E7').v).toEqual({ error: '#DIV/0!' });
    expect(cell(wb, 'B1', 1)).toMatchObject({ f: 'Продажи!D6*2', v: 2240 });
    expect(cell(wb, 'A1').s).toMatchObject({ bold: true, fill: '#dcefe9', align: 'center' });
    expect(cell(wb, 'B2').s.numFmt).toBe(String.raw`#,##0.00\ "₽"`); // Excel экранирует пробел
    expect(cell(wb, 'E2').s.numFmt).toBe(String.raw`dd\.mm\.yyyy`);
    expect(formatValue(cell(wb, 'E2').v, cell(wb, 'E2').s.numFmt)).toBe('06.10.2026');
    expect(formatValue(cell(wb, 'B2').v, cell(wb, 'B2').s.numFmt)).toBe('45,50 ₽');
    expect(formatValue(cell(wb, 'A8').v, cell(wb, 'A8').s.numFmt)).toBe('12,5%');
    expect(cell(wb, 'A2').s.border.top.style).toBe('thin');
    expect(wb.sheets[0].merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 2 }]);
    expect(wb.sheets[0].colW[0]).toBe(145);
    expect(wb.sheets[0].rowH[0]).toBe(32);
  });

  it('экспорт → импорт: то же содержимое; openpyxl видит формулы, значения, стили', async () => {
    const { wb } = await importXlsx(readFileSync('tests/fixtures/book.xlsx'));
    const bytes = await exportXlsx(wb);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/book.xlsx', bytes);
    const back = (await importXlsx(bytes)).wb;
    expect(back.sheets.map((s) => s.name)).toEqual(['Продажи', 'Свод']);
    for (let i = 0; i < 2; i++) {
      expect([...back.sheets[i].cells.entries()]).toEqual([...wb.sheets[i].cells.entries()]);
      expect(back.sheets[i].merges).toEqual(wb.sheets[i].merges);
      expect(back.sheets[i].colW).toEqual(wb.sheets[i].colW);
      expect(back.sheets[i].rowH).toEqual(wb.sheets[i].rowH);
    }
    const out = execFileSync(PYTHON, ['-c', `
import openpyxl
wb = openpyxl.load_workbook('tests/out/book.xlsx')
s = wb['Продажи']
print(s['D4'].value, s['D6'].value, s['A1'].font.b, s['A1'].fill.fgColor.rgb, s['B2'].number_format, list(map(str, s.merged_cells.ranges)))
print(wb['Свод']['B1'].value)
v = openpyxl.load_workbook('tests/out/book.xlsx', data_only=True)
print(v['Продажи']['D6'].value, v['Свод']['B1'].value)
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(out).toContain(String.raw`=B4*C4 =SUM(D2:D5) True FFDCEFE9 #,##0.00\ "₽" ['A6:C6']`);
    expect(out).toContain('=Продажи!D6*2');
    expect(out).toContain('1120 2240');
  });

  it('сохранение поверх исходного файла: всё непонятное — байт в байт, правка — только в своём листе', async () => {
    const JSZip = (await import('jszip')).default;
    const src = readFileSync('tests/fixtures/preserve.xlsx');
    const { wb, warnings, kept } = await importXlsx(src);
    expect(warnings).toEqual({});
    expect(Object.keys(kept)).toEqual(expect.arrayContaining(['условное форматирование', 'проверка данных', 'рисунки и диаграммы',
      'примечания к ячейкам', 'гиперссылки', '«умные» таблицы', 'сводные таблицы', 'именованные диапазоны']));
    // скрытые и сгруппированные строки/столбцы, стиль столбца
    expect(wb.sheets[0].rowInfo[14]).toEqual({ hidden: true });
    expect(wb.sheets[0].rowInfo[15]).toEqual({ hidden: true, level: 1 });
    expect(wb.sheets[0].rowInfo[18]).toEqual({ collapsed: true });
    expect(wb.sheets[0].colInfo[6]).toEqual({ hidden: true });
    expect(wb.sheets[0].colInfo[9]).toEqual({ level: 1 });
    expect(wb.sheets[0].colInfo[7].xf).toBeGreaterThan(0);
    const edited = setInput(wb, 0, 2, 1, '115'); // B3
    const out = await exportXlsx(edited, { source: src });
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/preserve.xlsx', out);
    const a = await JSZip.loadAsync(src); const b = await JSZip.loadAsync(out);
    const changed = [];
    for (const p of Object.keys(a.files)) {
      if (a.files[p].dir) continue;
      if (!b.file(p)) { changed.push(`-${p}`); continue; }
      if (await a.file(p).async('string') !== await b.file(p).async('string')) changed.push(p);
    }
    expect(changed.sort()).toEqual(['-xl/calcChain.xml', '[Content_Types].xml', 'docProps/core.xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    const s1 = await b.file('xl/worksheets/sheet1.xml').async('string');
    for (const t of ['<pane ', '<conditionalFormatting', '<dataValidations', '<hyperlinks>', '<drawing ', '<legacyDrawing ', '<pageSetup ', 'hidden="1" outlineLevel="1"'])
      expect(s1).toContain(t);
    expect(s1).toContain('<c r="F3" s="1" t="s"><v>18</v></c>'); // неизменённые ячейки — исходной разметкой
    expect(await b.file('xl/workbook.xml').async('string')).toContain('fullCalcOnLoad="1"');
    const back = (await importXlsx(out)).wb;
    expect(cell(back, 'B3').v).toBe(115);
    expect(cell(back, 'D3')).toMatchObject({ f: 'C3-B3', v: 6 });
    expect(back.sheets[0].rowInfo).toEqual(wb.sheets[0].rowInfo);
    expect(back.sheets[0].colInfo).toEqual(wb.sheets[0].colInfo);
    expect(back.sheets[0].colW).toEqual(wb.sheets[0].colW);
  });

  it('вставка строк, переименование, новый и удалённый лист поверх исходного файла', async () => {
    const JSZip = (await import('jszip')).default;
    const src = readFileSync('tests/fixtures/preserve.xlsx');
    let { wb } = await importXlsx(src);
    wb = shiftCells(wb, 0, 'row', 5, 2);
    wb = setInput(wb, 0, 0, 4, 'Новый текст');
    wb = setStyle(wb, 0, { r1: 0, c1: 4, r2: 0, c2: 4 }, { fill: '#ffcc00', numFmt: '0.000' });
    wb = renameSheet(wb, 0, 'Данные 2026');
    wb = addSheet(wb);
    wb = setInput(wb, 3, 0, 0, "='Данные 2026'!B3*2");
    wb = deleteSheet(wb, 1);
    wb = { ...wb, active: 0 };
    const out = await exportXlsx(wb, { source: src });
    writeFileSync('tests/out/preserve_ops.xlsx', out);
    const z = await JSZip.loadAsync(out);
    const wbXml = await z.file('xl/workbook.xml').async('string');
    expect(wbXml).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">'Данные 2026'!$A$1:$D$16</definedName>`);
    expect(wbXml).toContain(`<definedName name="План">'Данные 2026'!$B$2:$B$15</definedName>`);
    expect(wbXml).toMatch(/<sheet name="Свод" sheetId="3" r:id="rId3"\/><sheet name="Лист4" sheetId="4" r:id="rIdPeroS1"\/><\/sheets>/);
    expect(z.file('xl/worksheets/sheet2.xml')).toBeNull();
    const chart = await z.file('xl/charts/chart1.xml').async('string');
    expect(chart).toContain(`<c:f>'Данные 2026'!$C$2:$C$15</c:f>`);
    expect(await z.file('xl/drawings/drawing1.xml').async('string')).toContain('<xdr:row>23</xdr:row>');
    const s1 = await z.file('xl/worksheets/sheet1.xml').async('string');
    expect(s1).toContain('<conditionalFormatting sqref="D2:D15">');
    expect(s1).toContain('sqref="B2:B15"');
    const back = (await importXlsx(out)).wb;
    expect(back.sheets.map((s) => s.name)).toEqual(['Данные 2026', 'Свод', 'Лист4']);
    expect(cell(back, 'E1').s).toMatchObject({ fill: '#ffcc00', numFmt: '0.000' });
    expect(cell(back, 'A1', 2).v).toBe(220);
    expect(back.sheets[0].rowInfo[16]).toEqual({ hidden: true });
    const py = execFileSync(PYTHON, ['-c', `
import openpyxl
wb = openpyxl.load_workbook('tests/out/preserve_ops.xlsx')
s = wb['Данные 2026']
print(wb.sheetnames, s['E1'].value, s.freeze_panes, s.row_dimensions[17].hidden, len(s.conditional_formatting), s.data_validations.dataValidation[0].sqref)
`], { env: { ...process.env, PYTHONIOENCODING: 'utf-8' } }).toString('utf8');
    expect(py).toContain("['Данные 2026', 'Свод', 'Лист4'] Новый текст B2 True 2 B2:B15");
  });

  it('закрепление областей, скрытие и группировка: модель и файл', async () => {
    const m = await import('../src/sheet/model.js');
    let wb = put(newWorkbook(), { A1: 'шапка', A2: '1', A3: '2', A4: '3', A5: '=СУММ(A2:A4)' });
    wb = m.setFreeze(wb, 0, { rows: 1, cols: 1 });
    wb = m.groupLines(wb, 0, 'row', 1, 3, 1);
    wb = m.groupLines(wb, 0, 'row', 2, 2, 1);
    expect(m.outlineRuns(wb.sheets[0], 'row', 1)).toEqual([{ a: 1, b: 3 }]);
    wb = m.collapseGroup(wb, 0, 'row', 2, 2);          // вложенная — свёрнута
    wb = m.collapseGroup(wb, 0, 'row', 1, 3);
    expect(wb.sheets[0].rowInfo[4]).toEqual({ collapsed: true });
    wb = m.expandGroup(wb, 0, 'row', 1, 3, 1);           // внешняя развёрнута, вложенная осталась свёрнутой
    expect([1, 2, 3].map((r) => m.isHidden(wb.sheets[0], 'row', r))).toEqual([false, true, false]);
    wb = m.setHidden(wb, 0, 'col', 3, 4, true);
    const out = await exportXlsx(wb);
    const back = (await importXlsx(out)).wb.sheets[0];
    expect(back.freeze).toEqual({ rows: 1, cols: 1 });
    expect(back.rowInfo).toEqual(wb.sheets[0].rowInfo);
    expect(back.colInfo).toEqual({ 3: { hidden: true }, 4: { hidden: true } });
    // поверх исходного файла: закрепление меняется, остальное в <sheetView> — как было
    const src = readFileSync('tests/fixtures/preserve.xlsx');
    let p = (await importXlsx(src)).wb;
    expect(p.sheets[0].freeze).toEqual({ rows: 1, cols: 1 });
    p = m.setFreeze(p, 0, { rows: 2, cols: 0 });
    p = m.groupLines(p, 0, 'col', 1, 2, 1);
    const JSZip = (await import('jszip')).default;
    const s1 = await (await JSZip.loadAsync(await exportXlsx(p, { source: src }))).file('xl/worksheets/sheet1.xml').async('string');
    expect(s1).toContain('<sheetView tabSelected="1" workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView>');
    expect(s1).toMatch(/<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" outlineLevelRow="1" outlineLevelCol="1"\/>/);
  });

  it('новая книга из модели: русские функции → английские, кэшированные значения', async () => {
    let wb = put(newWorkbook(), { A1: '2', A2: '3', A3: '=СУММ(A1:A2)*10', B1: 'текст', B2: '=ЕСЛИ(A3>40;"да";"нет")' });
    wb = setStyle(wb, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, { bold: true, numFmt: '0.00' });
    const back = (await importXlsx(await exportXlsx(wb))).wb;
    expect(cell(back, 'A3')).toMatchObject({ f: 'SUM(A1:A2)*10', v: 50 });
    expect(cell(back, 'B2')).toMatchObject({ f: 'IF(A3>40,"да","нет")', v: 'да' });
    expect(cell(back, 'A1').s).toEqual({ bold: true, numFmt: '0.00' });
  });
});
