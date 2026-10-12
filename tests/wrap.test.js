// Плавающие рисунки: wp:anchor ↔ модель обтекания и положения (src/editor/anchor.js), импорт/экспорт.
// Раскладка на экране сверяется с Word отдельно: tests/wrap_check.mjs (эталон — tests/make_wrap_fixture.ps1).
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'fs';
import JSZip from 'jszip';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { makeAnchor, readAnchor, setAnchorXml, wrapKind } from '../src/editor/anchor.js';

const images = (doc) => { const out = []; const w = (n) => { if (n.type === 'image') out.push(n); (n.content || []).forEach(w); }; w(doc); return out; };
const docXml = async (bytes) => (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');

describe('обтекание рисунков', () => {
  it('файл Word: виды обтекания и положение читаются', async () => {
    const { doc } = await importDocx(readFileSync('tests/fixtures/wrap.docx'));
    const a = images(doc).map((n) => n.attrs.anchor);
    expect(a.map(wrapKind)).toEqual(['square', 'square', 'tight', 'topBottom', 'behind', 'front']);
    expect(a[0]).toMatchObject({ h: { rel: 'column', align: null, off: 0 }, v: { rel: 'paragraph', off: 0 }, dist: { l: 12, r: 12 } });
    expect(a[1].h).toEqual({ rel: 'margin', align: 'right', off: 0 });
    expect(a[1].v.off).toBeCloseTo(18.87, 1); // 0,5 см
    expect(a[2].h.align).toBe('center');
  });

  it('неизменённые рисунки сохраняются как в файле; разметка ↔ модель без потерь', async () => {
    const src = readFileSync('tests/fixtures/wrap.docx');
    const { doc } = await importDocx(src);
    const anchors = (x) => x.match(/<wp:anchor[\s\S]*?<\/wp:anchor>/g);
    expect(anchors(await docXml(await exportDocx(doc, { source: src })))).toEqual(anchors(await docXml(src)));
    for (const n of images(doc)) expect(readAnchor(setAnchorXml(n.attrs.xml, n.attrs.anchor))).toEqual(n.attrs.anchor);
  });

  it('смена обтекания, в текст и обратно — рисунок (обрезка, эффекты) не теряется', async () => {
    const src = readFileSync('tests/fixtures/wrap.docx');
    const { doc } = await importDocx(src);
    const [a, b, c] = images(doc);
    a.attrs.anchor = makeAnchor('behind', a.attrs.anchor);
    b.attrs.anchor = null;
    c.attrs.anchor = makeAnchor('topBottom', c.attrs.anchor);
    const out = await exportDocx(doc, { source: src });
    const back = images((await importDocx(out)).doc).map((n) => n.attrs.anchor);
    expect(back.map(wrapKind)).toEqual(['behind', 'inline', 'topBottom', 'topBottom', 'behind', 'front']);
    expect(back[0].h).toEqual(a.attrs.anchor.h); // положение то же
    const x = await docXml(out);
    expect((x.match(/<wp:inline\b/g) || []).length).toBe(1);
    expect((x.match(/<a:blip\b/g) || []).length).toBe(6);
  });

  it('новый рисунок с обтеканием — wp:anchor по схеме', () => {
    const inline = '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Рисунок 1"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="x"><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r>';
    const x = setAnchorXml(inline, makeAnchor('square', null));
    // порядок по схеме CT_Anchor: simplePos, positionH, positionV, extent, effectExtent, wrap*, docPr, cNvGraphicFramePr, graphic
    expect(x).toMatch(/<wp:simplePos[^>]*\/><wp:positionH relativeFrom="column"><wp:posOffset>0<\/wp:posOffset><\/wp:positionH><wp:positionV relativeFrom="paragraph">.*<wp:extent .*<wp:effectExtent .*<wp:wrapSquare wrapText="bothSides"\/><wp:docPr .*<wp:cNvGraphicFramePr\/><a:graphic/);
    expect(readAnchor(x)).toMatchObject({ wrap: 'square', behind: false, dist: { t: 0, b: 0 } });
  });
});
