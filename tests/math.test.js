// Формулы Word (OMML) → MathML для показа (src/editor/omml.js); файл — tests/fixtures/math.docx из Word (make_math_fixture.ps1)
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { ommlToMathml } from '../src/editor/omml.js';

const mathXml = (doc) => {
  const out = [];
  const walk = (n) => { if (n.type === 'opaque' && n.attrs.kind === 'math') out.push(n.attrs.xml); (n.content || []).forEach(walk); };
  walk(doc);
  return out;
};

describe('формулы Word', () => {
  it('все формулы из Word разбираются в MathML нужного строения', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/math.docx'));
    const xs = mathXml(doc);
    expect(xs).toHaveLength(9);
    const ml = xs.map(ommlToMathml);
    expect(ml.every(Boolean)).toBe(true);
    // корень квадратного уравнения: дробь, квадратный корень, степень, минус вместо дефиса
    expect(ml[0]).toMatch(/^<math display="block">/);
    expect(ml[0]).toContain('<mfrac>');
    expect(ml[0]).toContain('<msqrt>');
    expect(ml[0]).toContain('<msup><mi>b</mi><mn>2</mn></msup>');
    expect(ml[0]).toContain('<mo>−</mo>');
    // сумма — пределы над и под знаком, интеграл — справа
    expect(ml[1]).toContain('<munderover><mo largeop="true" movablelimits="false">∑</mo>');
    expect(ml[2]).toContain('<msubsup><mo largeop="true" movablelimits="false">∫</mo><mn>0</mn><mn>1</mn></msubsup>');
    // предел: «lim» прямым шрифтом, под ним n→∞, применение функции
    expect(ml[3]).toContain('<munder><mi>lim</mi>');
    expect(ml[3]).toContain('<mo>⁡</mo>');
    expect(ml[4]).toContain('<msup><mi>sin</mi><mn>2</mn></msup>');
    // матрица в квадратных скобках, система с левой фигурной скобкой без правой
    expect(ml[5]).toMatch(/<mo fence="true" stretchy="true" lspace="0" rspace="0">\[<\/mo><mtable><mtr><mtd><mn>1<\/mn><\/mtd>/);
    expect(ml[7]).toContain('<mo fence="true" stretchy="true" lspace="0" rspace="0">{</mo><mtable>');
    expect(ml[7]).not.toMatch(/<\/mtable><mo/);
    // корень 3-й степени, двойной индекс, акценты: «крышка» по умолчанию и черта
    expect(ml[6]).toContain('<mroot><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mn>3</mn></mroot>');
    expect(ml[6]).toContain('<msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup>');
    expect(ml[6]).toContain('<mover accent="true"><mi>x</mi><mo stretchy="true">^</mo></mover>');
    expect(ml[6]).toContain('<mrow style="border-top:0.06em solid currentColor;padding-top:0.08em"><mi>y</mi></mrow>');
    // формула внутри строки — не «выключная»
    expect(ml[8]).toMatch(/^<math><mi>S<\/mi><mo>=<\/mo><mi>π<\/mi><msup>/);
  });

  it('сохранение не меняет формулы', async () => {
    const src = readFileSync('tests/fixtures/math.docx');
    const { doc } = await importDocx(src);
    const back = (await importDocx(await exportDocx(doc, { source: src }))).doc;
    expect(mathXml(back)).toEqual(mathXml(doc));
  });

  it('непонятное — по содержимому, мусор — null', () => {
    const ns = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    expect(ommlToMathml(`<m:oMath ${ns}><m:unknown><m:r><m:t>x</m:t></m:r></m:unknown></m:oMath>`)).toBe('<math><mi>x</mi></math>');
    expect(ommlToMathml('<w:r/>')).toBe(null);
  });
});
