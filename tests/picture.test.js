// Обрезка и поворот рисунков: a:srcRect и a:xfrm ↔ модель (src/editor/picFormat.js), импорт/экспорт.
// Эталон — tests/make_crop_fixture.ps1 (Word COM: обрезка в строке, поворот 30° с отражением, поворот 90°).
import { describe, expect, it } from 'vitest';
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import JSZip from 'jszip';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { readPicFormat, rotatedPad, setPicFormatXml } from '../src/editor/picFormat.js';

const images = (doc) => { const out = []; const w = (n) => { if (n.type === 'image') out.push(n); (n.content || []).forEach(w); }; w(doc); return out; };

describe('обрезка и поворот рисунков', () => {
  it('файл Word: обрезка, поворот, отражение читаются', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/crop.docx'));
    const [a, b, c] = images(doc).map((n) => n.attrs);
    expect(a.crop).toEqual({ l: 0.09449, t: 0.06299, r: 0.02835, b: 0.0252 });
    expect([b.rotate, b.flipH, b.flipV]).toEqual([330, true, false]);
    expect(c.rotate).toBe(90);
  });

  it('неизменённое — как в файле; изменённое — srcRect по схеме, поворот и запас под габарит', async () => {
    const src = readFileSync('tests/fixtures/crop.docx');
    const { doc } = await importDocx(src);
    const xml = async (b) => (await JSZip.loadAsync(b)).file('word/document.xml').async('string');
    const pics = (x) => x.match(/<pic:pic\b[\s\S]*?<\/pic:pic>/g);
    expect(pics(await xml(await exportDocx(doc, { source: src })))).toEqual(pics(await xml(src)));
    const [a, b] = images(doc);
    a.attrs.crop = { l: 0.2, t: 0, r: 0, b: 0.1 };
    b.attrs.rotate = 45; b.attrs.flipH = false;
    const out = await exportDocx(doc, { source: src });
    const x = await xml(out);
    expect(x).toMatch(/<a:blip\b[^>]*(?:\/>|>[\s\S]*?<\/a:blip>)<a:srcRect l="20000" b="10000"\/><a:stretch>/);
    expect(x).toContain('<a:xfrm rot="2700000">');
    const back = images((await importDocx(out)).doc).map((n) => n.attrs);
    expect(back[0].crop).toEqual({ l: 0.2, t: 0, r: 0, b: 0.1 });
    expect([back[1].rotate, back[1].flipH]).toEqual([45, false]);
    mkdirSync('tests/out', { recursive: true });
    writeFileSync('tests/out/crop_changed.docx', out); // для проверки в Word (tests/crop_check.ps1)
  });

  it('запас под повёрнутую картинку — габарит минус рамка, пополам', () => {
    expect(rotatedPad(100, 50, 0)).toEqual({ x: 0, y: 0 });
    const p = rotatedPad(100, 50, 90);
    expect(p.x).toBe(0); // повёрнутая на 90° картинка 100×50 уже рамки по ширине — запаса нет
    expect(p.y).toBeCloseTo(25, 5);
    const q = rotatedPad(100, 100, 45);
    expect(q.x).toBeCloseTo((Math.SQRT2 * 100 - 100) / 2, 5);
    // без srcRect у рисунка — новый элемент встаёт после a:blip
    const plain = '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/></a:xfrm></pic:spPr></pic:pic>';
    const y = setPicFormatXml(plain, { crop: { l: 0.1, t: 0, r: 0, b: 0 }, rotate: 0 });
    expect(y).toContain('<a:blip r:embed="rId1"/><a:srcRect l="10000"/><a:stretch>');
    expect(readPicFormat(y).crop).toEqual({ l: 0.1, t: 0, r: 0, b: 0 });
  });
});
