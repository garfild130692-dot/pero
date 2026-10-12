// Пересчёт по графу зависимостей: после любой цепочки правок значения — как при полном пересчёте;
// пересчитываются только затронутые формулы; режим «вручную»; летучие функции.
import { describe, it, expect } from 'vitest';
import * as m from '../src/sheet/model.js';

const { newWorkbook, setInput, recalc, valueAt } = m;
// снимок значений книги (с «разливом»)
function snapshot(wb) {
  const out = [];
  for (const sh of wb.sheets) {
    const keys = new Set([...sh.cells.keys(), ...(sh.spill ? sh.spill.keys() : [])]);
    for (const k of [...keys].sort()) {
      const [r, c] = k.split(',').map(Number);
      const v = valueAt(sh, r, c);
      out.push(`${sh.name}!${k}=${v && typeof v === 'object' ? v.error : typeof v === 'number' ? +v.toPrecision(12) : v}`);
    }
  }
  return out.join('\n');
}
// детерминированный генератор
function rng(seed) { let s = seed; return () => { s = (s * 1103515245 + 12345) % 2147483648; return s / 2147483648; }; }

describe('пересчёт по графу зависимостей', () => {
  it('случайные правки: частичный пересчёт = полному', () => {
    let unstable = 0;
    for (let seed = 1; seed <= 25; seed++) {
      const rnd = rng(seed);
      const pick = (a) => a[Math.floor(rnd() * a.length)];
      const cell = () => `${pick('ABCDE')}${1 + Math.floor(rnd() * 8)}`;
      const rangeTxt = () => { const a = pick('ABCDE'); const b = pick('ABCDE'); const r1 = 1 + Math.floor(rnd() * 8); return `${a < b ? a : b}${r1}:${a < b ? b : a}${r1 + Math.floor(rnd() * 4)}`; };
      const formulas = [
        () => `=${cell()}+${cell()}`, () => `=СУММ(${rangeTxt()})`, () => `=СУММ(${pick('ABCDE')}:${pick('ABCDE')})`,
        () => `=ЕСЛИ(${cell()}>3;${cell()};${cell()}*2)`, () => `=Лист2!${cell()}+1`, () => `=СЧЁТ(${rangeTxt()})`,
        () => `=ПОСЛЕД(${1 + Math.floor(rnd() * 3)})`, () => `=СМЕЩ(A1;${Math.floor(rnd() * 3)};${Math.floor(rnd() * 3)})`,
        () => `=ДВССЫЛ("${cell()}")`, () => `=МАКС(${rangeTxt()})*2`, () => `=ВПР(${cell()};A1:C8;2;ЛОЖЬ)`,
        () => `=СУММ(Лист1:Лист2!${cell()})`, () => `=${cell()}&"x"`, () => `=Ставка*${cell()}`,
        () => `=ПОСЛЕД(ЕСЛИ(${cell()}>5;3;1))`, () => `=СУММ(${cell()}#)`, () => `=LET(x;${cell()};x*2)`,
        () => `=MAP(${rangeTxt()};LAMBDA(v;v+1))`, () => `=ТРАНСП(${rangeTxt()})`,
      ];
      let wb = newWorkbook();
      wb = m.addSheet(wb);
      wb = m.setNames(wb, [{ name: 'Ставка', text: 'Лист2!$A$1', sheet: null }]);
      for (let step = 0; step < 60; step++) {
        const si = rnd() < 0.75 ? 0 : 1;
        const ref = cell();
        const r = +ref.slice(1) - 1; const c = ref.charCodeAt(0) - 65;
        const kind = rnd();
        if (kind < 0.45) wb = setInput(wb, si, r, c, String(Math.floor(rnd() * 10)));
        else if (kind < 0.85) wb = setInput(wb, si, r, c, pick(formulas)());
        else if (kind < 0.88) wb = m.clearRange(wb, si, { r1: r, c1: c, r2: r + 1, c2: c + 1 });
        else if (kind < 0.91) wb = m.fill(wb, si, { r1: r, c1: c, r2: r + 2, c2: c }, 'down');
        else if (kind < 0.93) wb = m.sortRange(wb, si, { r1: 0, c1: 0, r2: 7, c2: 2 }, 0, rnd() < 0.5);
        else if (kind < 0.95) wb = m.pasteBlock(wb, si, r, c, m.copyBlock(wb.sheets[si], { r1: 0, c1: 0, r2: 1, c2: 1 }));
        else if (kind < 0.97) wb = m.shiftCells(wb, si, rnd() < 0.5 ? 'row' : 'col', Math.floor(rnd() * 6), rnd() < 0.5 ? 1 : -1);
        else wb = setInput(wb, si, r, c, '');
        const stable = m.recalcInfo.converged;
        const full = recalc(wb, { full: true });
        // «разлив», питающий сам себя через цепочку формул, не сходится (как циклическая ссылка в Excel — у книги
        // несколько устойчивых состояний, значения зависят от истории): такие состояния не сравниваются
        if (!stable || !m.recalcInfo.converged) { wb = full; unstable++; continue; }
        expect(snapshot(wb), `seed ${seed}, шаг ${step}`).toBe(snapshot(full));
      }
    }
    expect(unstable).toBeLessThan(25 * 60 * 0.15);
  });

  it('ветвление (отмена и новая правка): граф, общий для цепочки книг, даёт те же значения, что полный пересчёт', () => {
    // граф после частичного пересчёта правится на месте (этап 17); книга, к которой вернулись отменой, сверяется с ним
    for (let seed = 1; seed <= 15; seed++) {
      const rnd = rng(1000 + seed);
      const pick = (a) => a[Math.floor(rnd() * a.length)];
      const cell = () => `${pick('ABCD')}${1 + Math.floor(rnd() * 6)}`;
      const formulas = [() => `=${cell()}+${cell()}`, () => `=СУММ(A1:${pick('ABCD')}6)`, () => `=СУММ(${pick('ABCD')}:${pick('ABCD')})`,
        () => `=ВПР(${cell()};A1:C6;2;ЛОЖЬ)`, () => `=ЕСЛИ(${cell()}>4;${cell()};1)`, () => `=ПОСЛЕД(${1 + Math.floor(rnd() * 3)})`];
      let wb = newWorkbook();
      const history = [wb];
      for (let step = 0; step < 50; step++) {
        if (rnd() < 0.15) wb = pick(history); // «отмена» к одной из прежних книг
        const ref = cell();
        const r = +ref.slice(1) - 1; const c = ref.charCodeAt(0) - 65;
        wb = setInput(wb, 0, r, c, rnd() < 0.5 ? String(Math.floor(rnd() * 10)) : pick(formulas)());
        history.push(wb);
        const stable = m.recalcInfo.converged;
        const full = recalc(wb, { full: true });
        if (!stable || !m.recalcInfo.converged) continue;
        expect(snapshot(wb), `seed ${seed}, шаг ${step}`).toBe(snapshot(full));
      }
    }
  });

  it('открытие файла: граф по ссылкам формул (без пересчёта), дальше правки — как при полном пересчёте', async () => {
    const { exportXlsx, importXlsx } = await import('../src/sheet/xlsx.js');
    for (let seed = 1; seed <= 6; seed++) {
      const rnd = rng(500 + seed);
      const pick = (a) => a[Math.floor(rnd() * a.length)];
      const cell = () => `${pick('ABCDE')}${1 + Math.floor(rnd() * 8)}`;
      const formulas = [() => `=${cell()}*2+${cell()}`, () => `=СУММ(A1:${pick('ABCDE')}8)`, () => `=СУММ(${pick('ABCDE')}:${pick('ABCDE')})`,
        () => `=ВПР(${cell()};A1:C8;2;ЛОЖЬ)`, () => `=ЕСЛИ(${cell()}>4;${cell()};${cell()})`, () => `=ПОСЛЕД(${1 + Math.floor(rnd() * 3)})`,
        () => `=Ставка*${cell()}`, () => `=СУММ(Лист1:Лист2!${cell()})`, () => `=LET(x;${cell()};x+1)`, () => `=СУММ(${cell()}#)`,
        () => `=ИНДЕКС(A1:E8;${1 + Math.floor(rnd() * 8)};2)`, () => `=Лист2!${cell()}-${cell()}`];
      let wb = m.setNames(m.addSheet(newWorkbook()), [{ name: 'Ставка', text: 'Лист2!$B$2', sheet: null }]);
      for (let i = 0; i < 70; i++) {
        const ref = cell(); const si = rnd() < 0.8 ? 0 : 1;
        wb = setInput(wb, si, +ref.slice(1) - 1, ref.charCodeAt(0) - 65, rnd() < 0.5 ? String(Math.floor(rnd() * 10)) : pick(formulas)());
      }
      wb = (await importXlsx(await exportXlsx(wb))).wb;
      const opened = recalc(wb, { full: true });
      if (m.recalcInfo.converged) expect(snapshot(wb), `seed ${seed}: открытие`).toBe(snapshot(opened));
      for (let step = 0; step < 30; step++) {
        const ref = cell(); const si = rnd() < 0.8 ? 0 : 1;
        wb = setInput(wb, si, +ref.slice(1) - 1, ref.charCodeAt(0) - 65, rnd() < 0.6 ? String(Math.floor(rnd() * 10)) : pick(formulas)());
        const stable = m.recalcInfo.converged;
        const full = recalc(wb, { full: true });
        if (!stable || !m.recalcInfo.converged) continue;
        expect(snapshot(wb), `seed ${seed}, шаг ${step}`).toBe(snapshot(full));
      }
    }
  });

  it('считаются только затронутые формулы', () => {
    let wb = newWorkbook();
    for (let i = 0; i < 300; i++) { wb = setInput(wb, 0, i, 0, String(i)); }
    for (let i = 0; i < 300; i++) wb = setInput(wb, 0, i, 1, `=A${i + 1}*2`);
    wb = setInput(wb, 0, 0, 2, '=СУММ(B1:B300)');
    // правка одной ячейки: новыми становятся только её формула и сумма
    const before = wb.sheets[0].cells;
    wb = setInput(wb, 0, 5, 0, '100');
    const after = wb.sheets[0].cells;
    const changed = [...after.keys()].filter((k) => after.get(k) !== before.get(k));
    expect(changed.sort()).toEqual(['0,2', '5,0', '5,1']);
    expect(valueAt(wb.sheets[0], 0, 2)).toBe(2 * (299 * 300 / 2 - 5 + 100));
  });

  it('режим «вручную»: правки не пересчитывают зависимые до F9', () => {
    let wb = newWorkbook();
    wb = setInput(wb, 0, 0, 0, '1');
    wb = setInput(wb, 0, 0, 1, '=A1*10');
    wb = m.setCalcMode(wb, 'manual');
    wb = setInput(wb, 0, 0, 0, '2');
    expect(valueAt(wb.sheets[0], 0, 1)).toBe(10);
    expect(wb.calcPending).toBe(true);
    wb = setInput(wb, 0, 1, 1, '=A1+1'); // новая формула считается сразу (как в Excel)
    expect(valueAt(wb.sheets[0], 1, 1)).toBe(3);
    const f9 = recalc(wb, { full: true });
    expect(valueAt(f9.sheets[0], 0, 1)).toBe(20);
    expect(f9.calcPending).toBeUndefined();
    // переход на «авто» досчитывает отложенное
    const auto = m.setCalcMode(wb, 'auto');
    expect(valueAt(auto.sheets[0], 0, 1)).toBe(20);
    expect(auto.calcMode).toBeUndefined();
  });

  it('режим «вручную» сохраняется в файле (calcPr calcMode), при открытии — полный пересчёт', async () => {
    const { exportXlsx, importXlsx } = await import('../src/sheet/xlsx.js');
    const JSZip = (await import('jszip')).default;
    let wb = setInput(newWorkbook(), 0, 0, 0, '2');
    wb = setInput(wb, 0, 0, 1, '=A1*3');
    wb = m.setCalcMode(wb, 'manual');
    wb = setInput(wb, 0, 0, 0, '5'); // зависимая не пересчитана
    const bytes = await exportXlsx(wb);
    const xml = await (await JSZip.loadAsync(bytes)).file('xl/workbook.xml').async('string');
    expect(xml).toMatch(/<calcPr[^>]*calcMode="manual"/);
    const back = (await importXlsx(bytes)).wb;
    expect(back.calcMode).toBe('manual');
    expect(valueAt(back.sheets[0], 0, 1)).toBe(15);
    const auto = (await importXlsx(await exportXlsx(m.setCalcMode(back, 'auto')))).wb;
    expect(auto.calcMode).toBeUndefined();
  });

  it('летучие функции пересчитываются при любой правке', () => {
    let wb = newWorkbook();
    wb = setInput(wb, 0, 0, 0, '=СЛЧИС()');
    wb = setInput(wb, 0, 0, 1, '=A1*1');
    const seen = new Set();
    for (let i = 0; i < 5; i++) { wb = setInput(wb, 0, 5, 5, String(i)); seen.add(valueAt(wb.sheets[0], 0, 0)); expect(valueAt(wb.sheets[0], 0, 1)).toBe(valueAt(wb.sheets[0], 0, 0)); }
    expect(seen.size).toBeGreaterThan(1);
  });
});
