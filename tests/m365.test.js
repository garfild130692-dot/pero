// Функции Microsoft 365, которых нет в Excel 2021 (сверка с Excel через COM невозможна) — по примерам справки Microsoft
import { describe, it, expect } from 'vitest';
import { newWorkbook, setInput, setNames, valueAt } from '../src/sheet/model.js';
import { fileFormula, fromUser, localize } from '../src/sheet/formula.js';

// книга с данными: rows — массив строк (A1…), формула в F1 (или указанной ячейке); результат — разлив из неё
function calc(formula, rows = [], names = null) {
  let wb = newWorkbook();
  rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined) wb = setInput(wb, 0, i, j, String(v)); }));
  if (names) wb = setNames(wb, names);
  const R0 = 30;
  wb = setInput(wb, 0, R0, 0, formula);
  const sh = wb.sheets[0];
  const out = [];
  for (let i = 0; i < 20; i++) {
    const row = [];
    for (let j = 0; j < 12; j++) { const v = valueAt(sh, R0 + i, j); row.push(v && typeof v === 'object' ? v.error : v); }
    while (row.length && (row[row.length - 1] === null || row[row.length - 1] === undefined)) row.pop();
    if (!row.length) break;
    out.push(row);
  }
  return out.length === 1 && out[0].length === 1 ? out[0][0] : out;
}
const GRID = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];

describe('LAMBDA', () => {
  it('вызов, необязательные параметры, функция без вызова', () => {
    expect(calc('=LAMBDA(x;x+1)(2)')).toBe(3);
    expect(calc('=LAMBDA(x;y;ЕСЛИ(ISOMITTED(y);x;x*y))(5)')).toBe(5);
    expect(calc('=LAMBDA(x;y;ЕСЛИ(ISOMITTED(y);x;x*y))(5;3)')).toBe(15);
    expect(calc('=LAMBDA(x;x)')).toBe('#CALC!');
    expect(calc('=LAMBDA(x;x)(1;2)')).toBe('#VALUE!');
    expect(calc('=LET(f;LAMBDA(x;x*2);f(21))')).toBe(42);
  });
  it('именованные лямбды и рекурсия', () => {
    const names = [{ name: 'Квадрат', text: 'LAMBDA(x,x*x)', sheet: null }, { name: 'Факт', text: 'LAMBDA(n,IF(n<=1,1,n*Факт(n-1)))', sheet: null }];
    expect(calc('=Квадрат(7)', [], names)).toBe(49);
    expect(calc('=Факт(6)', [], names)).toBe(720);
  });
  it('запись в файле: _xlpm. у параметров', () => {
    expect(fileFormula(fromUser('LAMBDA(x;y;x+y)(1;2)'))).toBe('_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x+_xlpm.y)(1,2)');
    expect(localize('LAMBDA(x,x+1)(2)')).toBe('LAMBDA(x;x+1)(2)');
  });
});

describe('MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL', () => {
  it('примеры справки', () => {
    expect(calc('=MAP(A1:C3;LAMBDA(a;ЕСЛИ(a>4;a*a;a)))', GRID)).toEqual([[1, 2, 3], [4, 25, 36], [49, 64, 81]]);
    expect(calc('=MAP(A1:A3;B1:B3;LAMBDA(x;y;x+y))', GRID)).toEqual([[3], [9], [15]]);
    expect(calc('=REDUCE(0;A1:C3;LAMBDA(a;b;ЕСЛИ(b>4;a+b;a)))', GRID)).toBe(35);
    expect(calc('=SCAN(1;A1:C1;LAMBDA(a;b;a*b))', GRID)).toEqual([[1, 2, 6]]);
    expect(calc('=SCAN("";{"a";"b";"c"};LAMBDA(t;i;t&i))')).toEqual([['a', 'ab', 'abc']]);
    expect(calc('=MAKEARRAY(2;3;LAMBDA(r;c;r*c))')).toEqual([[1, 2, 3], [2, 4, 6]]);
    expect(calc('=BYROW(A1:C3;LAMBDA(r;МАКС(r)))', GRID)).toEqual([[3], [6], [9]]);
    expect(calc('=BYCOL(A1:C3;СУММ)', GRID)).toEqual([[12, 15, 18]]);
    // лямбда вернула массив там, где нужно одно значение
    expect(calc('=BYROW(A1:C3;LAMBDA(r;r))', GRID)).toEqual([['#CALC!'], ['#CALC!'], ['#CALC!']]);
  });
});

describe('форма массивов', () => {
  it('VSTACK/HSTACK/TOCOL/TOROW/WRAPROWS/WRAPCOLS', () => {
    expect(calc('=ВСТОЛБИК(A1:B1;C2:C3)', GRID)).toEqual([[1, 2], [6, '#N/A'], [9, '#N/A']]);
    expect(calc('=ГСТОЛБИК(A1:A2;B1:C1)', GRID)).toEqual([[1, 2, 3], [4, '#N/A', '#N/A']]);
    expect(calc('=ПОСТОЛБЦ(A1:B2)', GRID)).toEqual([[1], [2], [4], [5]]);
    expect(calc('=ПОСТРОК(A1:B2;0;ИСТИНА)', GRID)).toEqual([[1, 4, 2, 5]]);
    expect(calc('=ПОСТОЛБЦ(A1:B2;1)', [[1, null], [null, 2]])).toEqual([[1], [2]]);
    expect(calc('=СВЕРНСТРОК(A1:C1;2)', GRID)).toEqual([[1, 2], [3, '#N/A']]);
    expect(calc('=СВЕРНСТОЛБ(A1:C1;2;0)', GRID)).toEqual([[1, 3], [2, 0]]);
  });
  it('TAKE/DROP/CHOOSECOLS/CHOOSEROWS/EXPAND/TRIMRANGE', () => {
    expect(calc('=ВЗЯТЬ(A1:C3;2)', GRID)).toEqual([[1, 2, 3], [4, 5, 6]]);
    expect(calc('=ВЗЯТЬ(A1:C3;-1)', GRID)).toEqual([[7, 8, 9]]);
    expect(calc('=ВЗЯТЬ(A1:C3;;2)', GRID)).toEqual([[1, 2], [4, 5], [7, 8]]);
    expect(calc('=СБРОСИТЬ(A1:C3;1;-1)', GRID)).toEqual([[4, 5], [7, 8]]);
    expect(calc('=СБРОСИТЬ(A1:C3;3)', GRID)).toBe('#CALC!');
    expect(calc('=ВЫБОРСТОЛБЦ(A1:C3;-1;1)', GRID)).toEqual([[3, 1], [6, 4], [9, 7]]);
    expect(calc('=ВЫБОРСТОЛБЦ(A1:C3;4)', GRID)).toBe('#VALUE!');
    expect(calc('=ВЫБОРСТРОК(A1:C3;{1;3})', GRID)).toEqual([[1, 2, 3], [7, 8, 9]]);
    expect(calc('=РАЗВЕРНУТЬ(A1:B1;2;3;"-")', GRID)).toEqual([[1, 2, '-'], ['-', '-', '-']]);
    expect(calc('=РАЗВЕРНУТЬ(A1:B2;1)', GRID)).toBe('#VALUE!');
    expect(calc('=УРЕЗДИАПАЗОН(A1:E6)', [[null, null, null], [null, 1, 2], [null, 3, null]])).toEqual([[1, 2], [3, 0]]);
  });
});

describe('регулярные выражения', () => {
  const B = '\\';
  it('REGEXTEST/REGEXEXTRACT/REGEXREPLACE', () => {
    expect(calc(`=РЕГТЕСТ("abc123";"${B}d+")`)).toBe(true);
    expect(calc('=РЕГТЕСТ("ABC";"abc")')).toBe(false);
    expect(calc('=РЕГТЕСТ("ABC";"abc";1)')).toBe(true);
    expect(calc('=РЕГИЗВЛЕЧЬ("Иван Петров";"[А-Я][а-я]+")')).toBe('Иван');
    expect(calc('=РЕГИЗВЛЕЧЬ("Иван Петров";"[А-Я][а-я]+";1)')).toEqual([['Иван'], ['Петров']]);
    expect(calc(`=РЕГИЗВЛЕЧЬ("ab-12";"([a-z]+)-(${B}d+)";2)`)).toEqual([['ab', '12']]);
    expect(calc(`=РЕГИЗВЛЕЧЬ("abc";"${B}d")`)).toBe('#N/A');
    expect(calc(`=РЕГЗАМЕНИТЬ("a1b2c3";"${B}d";"#")`)).toBe('a#b#c#');
    expect(calc(`=РЕГЗАМЕНИТЬ("a1b2c3";"${B}d";"#";2)`)).toBe('a1b#c3');
    expect(calc(`=РЕГЗАМЕНИТЬ("a1b2c3";"${B}d";"#";-1)`)).toBe('a1b2c#');
    expect(calc(`=РЕГЗАМЕНИТЬ("Иван Петров";"(${B}S+) (${B}S+)";"$2 $1")`)).toBe('Петров Иван');
    expect(calc('=РЕГТЕСТ("a";"(")')).toBe('#VALUE!');
  });
});

describe('сводки', () => {
  const SALES = [['Регион', 'Товар', 'Сумма'], ['Север', 'A', 10], ['Юг', 'B', 20], ['Север', 'B', 5], ['Юг', 'A', 7], ['Север', 'A', 3]];
  it('PERCENTOF', () => {
    expect(calc('=ПРОЦЕНТ(A1:A2;A1:C3)', GRID)).toBeCloseTo(5 / 45, 12);
  });
  it('GROUPBY', () => {
    expect(calc('=ГРУПППО(A2:A6;C2:C6;СУММ)', SALES)).toEqual([['Север', 18], ['Юг', 27], ['Итог', 45]]);
    expect(calc('=ГРУПППО(A1:A6;C1:C6;СУММ;3)', SALES)).toEqual([['Регион', 'Сумма'], ['Север', 18], ['Юг', 27], ['Итог', 45]]);
    expect(calc('=ГРУПППО(A2:A6;C2:C6;СУММ;0;0;-2)', SALES)).toEqual([['Юг', 27], ['Север', 18]]);
    expect(calc('=ГРУПППО(A2:A6;C2:C6;СЧЁТ;0;0)', SALES)).toEqual([['Север', 3], ['Юг', 2]]);
    expect(calc('=ГРУПППО(A2:A6;C2:C6;ПРОЦЕНТ;0;0)', SALES)).toEqual([['Север', 0.4], ['Юг', 0.6]]);
    expect(calc('=ГРУПППО(A2:A6;C2:C6;LAMBDA(x;МАКС(x));0;0;;C2:C6>5)', SALES)).toEqual([['Север', 10], ['Юг', 20]]);
    expect(calc('=ГРУПППО(A2:B6;C2:C6;СУММ;0;2)', SALES)).toEqual([
      ['Север', 'A', 13], ['Север', 'B', 5], ['Север', '', 18], ['Юг', 'A', 7], ['Юг', 'B', 20], ['Юг', '', 27], ['Итог', '', 45]]);
  });
  it('PIVOTBY', () => {
    expect(calc('=СВОДПО(A2:A6;B2:B6;C2:C6;СУММ)', SALES)).toEqual([
      ['', 'A', 'B', 'Итог'], ['Север', 13, 5, 18], ['Юг', 7, 20, 27], ['Итог', 20, 25, 45]]);
    expect(calc('=СВОДПО(A2:A6;B2:B6;C2:C6;СУММ;0;0;;0)', SALES)).toEqual([['', 'A', 'B'], ['Север', 13, 5], ['Юг', 7, 20]]);
  });
});

describe('прочее', () => {
  it('EUROCONVERT (примеры справки)', () => {
    expect(calc('=EUROCONVERT(1,2;"DEM";"EUR")')).toBe(0.61);
    expect(calc('=EUROCONVERT(1;"FRF";"DEM";ИСТИНА;3)')).toBeCloseTo(0.29728616, 8);
    expect(calc('=EUROCONVERT(1;"FRF";"DEM";ЛОЖЬ;3)')).toBe(0.3);
    expect(calc('=EUROCONVERT(1;"FRF";"XXX")')).toBe('#VALUE!');
  });
  it('JIS, IMAGE, CALL', () => {
    expect(calc('=JIS("abc")')).toBe('abc');
    expect(calc('=IMAGE("https://example.com/a.png")')).toBe('#VALUE!');
  });
});
