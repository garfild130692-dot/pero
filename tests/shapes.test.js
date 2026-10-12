// Фигуры и надписи Word (shapes.js): разбор wps:wsp из файла Word, цвета темы, вид (SVG + текст), сохранность XML.
// Эталон — tests/fixtures/shapes.docx (tests/make_shapes_fixture.ps1, сам Word); раскладка строк вокруг фигур сверяется
// с Word отдельно — tests/shapes_check.mjs (83 из 83 строк).
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { parseShape, setThemeColors, shapeHtml, themeColorsFromXml } from '../src/editor/shapes.js';

const file = readFileSync('tests/fixtures/shapes.docx');
const drawings = (node, out = []) => {
  if (node.type === 'opaque' && node.attrs.kind === 'drawing') out.push(node.attrs.xml);
  (node.content || []).forEach((c) => drawings(c, out));
  return out;
};

describe('фигуры и надписи', async () => {
  const { doc } = await importDocx(file);
  setThemeColors(doc.attrs.themeColors);
  const shapes = drawings(doc).map(parseShape);

  it('все фигуры файла разобраны: геометрия, размер, обтекание', () => {
    expect(shapes.map((s) => s.geom)).toEqual(['rect', 'rect', 'ellipse', 'rightArrow', 'line', 'roundRect', 'rect']);
    expect(shapes.map((s) => s.anchor.wrap)).toEqual(['square', 'none', 'none', 'none', 'none', 'square', 'topBottom']);
    expect(shapes[0].w).toBeCloseTo(1800000 / 9525, 3); // 5 см
    expect(shapes[0].h).toBeCloseTo(1080000 / 9525, 3); // 3 см
    expect(shapes[1].anchor.behind).toBe(true);
    expect(shapes[0].effect.b).toBeCloseTo(25400 / 9525, 3); // запас wp:effectExtent — Word учитывает при обтекании
  });

  it('заливка и линия: из spPr, иначе из стиля фигуры с цветами темы (accent1, shade)', () => {
    expect(shapes[0].fill).toBe(null); // надпись: noFill
    expect(shapes[0].line.color).toBe('#000000');
    expect(shapes[3].fill).toBe('#4472c4'); // стрелка: fillRef accent1
    expect(shapes[3].line.color).toBe('#223962'); // lnRef accent1 shade 50000
    expect(shapes[4].line.tail).toBe('triangle');
    expect(shapes[4].line.width).toBeCloseTo(2 / 0.75, 2); // 2 pt
  });

  it('текст надписи: абзацы, цвет из fontRef, вертикальное выравнивание', () => {
    const text = (s) => s.paras.map((p) => p.runs.map((r) => r.text).join('')).join('|');
    expect(text(shapes[0])).toBe('Надпись: текст внутри рамки');
    expect(shapes[0].paras[0].runs[0].bold).toBe(true);
    expect(text(shapes[5])).toBe('Скруглённый');
    expect(shapes[5].paras[0].runs[0].color).toBe('#ffffff'); // fontRef lt1
    expect(shapes[5].vAnchor).toBe('ctr');
    expect(shapes[5].paras[0].align).toBe('center');
  });

  it('вид: SVG геометрии и текст поверх', () => {
    const html = shapeHtml(shapes[5]);
    expect(html).toMatch(/<svg[^>]*class="pero-shape-svg"/);
    expect(html).toMatch(/<path d="M[^"]*A/); // скругления — дуги
    expect(html).toMatch(/fill="#4472c4"/);
    expect(html).toContain('Скруглённый');
    expect(shapeHtml(shapes[4])).toMatch(/<line [^>]*marker-end=/);
  });

  it('цвета темы документа берутся из theme1.xml', async () => {
    const theme = await (await JSZip.loadAsync(file)).file('word/theme/theme1.xml').async('string');
    const c = themeColorsFromXml(theme);
    expect(c.accent1).toBe('4472C4');
    expect(c.dk1).toBe('000000'); // sysClr windowText — lastClr
    expect(doc.attrs.themeColors.accent2).toBe(c.accent2);
    // своя тема: акцент другим цветом — фигура перекрашивается
    setThemeColors({ ...c, accent1: '1F6F5C' });
    expect(parseShape(drawings(doc)[3]).fill).toBe('#1f6f5c');
    setThemeColors(doc.attrs.themeColors);
  });

  it('не фигура (рисунок, группа) — null: остаётся значком', () => {
    expect(parseShape('<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData><pic:pic/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')).toBe(null);
  });

  it('при сохранении фигуры пишутся исходным XML', async () => {
    const out = await exportDocx(doc, { source: file });
    const again = await importDocx(out);
    expect(drawings(again.doc)).toEqual(drawings(doc));
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml').async('string');
    expect((xml.match(/<wps:wsp\b/g) || []).length).toBe(7);
  });
});
