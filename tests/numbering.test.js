// Нумерация заголовков и формат нумерованных списков (src/editor/numbering.js, import/export)
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'fs';
import { importDocx } from '../src/docx/import.js';
import { exportDocx } from '../src/docx/export.js';
import { headingNumberCss, levelLabel, makeHeadingNumbering, markerKind } from '../src/editor/numbering.js';

const SRC = 'tests/fixtures/numbering.docx'; // tests/make_numbering_fixture.ps1 (Word COM)
const docXml = async (bytes) => (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
const numXml = async (bytes) => (await JSZip.loadAsync(bytes)).file('word/numbering.xml').async('string');
describe('нумерация заголовков', () => {
  it('схемы и подписи уровней', () => {
    const hn = makeHeadingNumbering('article');
    expect(levelLabel(hn.levels, 0, [3])).toBe('Статья 3.');
    expect(levelLabel(hn.levels, 2, [2, 1, 4])).toBe('2.1.4.');
    const sec = makeHeadingNumbering('section');
    expect(levelLabel(sec.levels, 0, [4])).toBe('Раздел IV');
    expect(levelLabel(sec.levels, 1, [4, 2])).toBe('2.');
    const css = headingNumberCss(makeHeadingNumbering('dots'));
    expect(css).toContain('.page-content h2:not([data-nonum])::before { content: counter(pero-h1, decimal) "." counter(pero-h2, decimal) ".";');
    expect(css).toContain('counter-set: pero-h2 0');
    expect(headingNumberCss({ ...makeHeadingNumbering('dots'), off: true })).toBe('');
  });

  it('из файла Word: нумерация стилей заголовков, «без номера», форматы списков', async () => {
    const { doc } = await importDocx(readFileSync(SRC));
    const hn = doc.attrs.headingNumbering;
    expect(hn.source).toEqual({ numId: '1', viaStyle: true });
    expect(hn.levels.slice(0, 3).map((l) => [l.text, l.on])).toEqual([['%1', true], ['%1.%2', true], ['%1.%2.%3', true]]);
    expect(hn.levels[3].on).toBe(false);
    const heads = doc.content.filter((n) => n.type === 'heading');
    expect(heads.map((h) => !!h.attrs.noNumber)).toEqual([false, false, false, false, false, true]);
    const lists = doc.content.filter((n) => n.type === 'orderedList').map((n) => n.attrs);
    expect(lists).toEqual([{ start: 1, numFmt: 'lowerLetter', lvlText: '%1)' }, { start: 1, numFmt: 'upperRoman' }]);
  });

  it('сохранение: как в файле / своя нумерация / отключена; списки a) и I. сохраняются', async () => {
    const src = readFileSync(SRC);
    const { doc } = await importDocx(src);
    // как в файле: нумерация остаётся в стилях, у заголовка «без номера» — numId 0
    const same = await docXml(await exportDocx(doc, { source: src }));
    expect((same.match(/<w:numId w:val="0"\/>/g) || []).length).toBe(1);
    // своя схема Пера: прямой numPr у заголовков на новый многоуровневый список
    const changed = JSON.parse(JSON.stringify(doc));
    changed.attrs.headingNumbering = makeHeadingNumbering('dots', doc.attrs.headingNumbering);
    const b2 = await exportDocx(changed, { source: src });
    const x2 = await docXml(b2);
    const ids = [...x2.matchAll(/<w:numPr><w:ilvl w:val="(\d)"\/><w:numId w:val="(\d+)"\/><\/w:numPr>/g)].map((m) => `${m[2]}:${m[1]}`);
    const hid = ids[0].split(':')[0];
    expect(ids.slice(0, 5)).toEqual([`${hid}:0`, `${hid}:1`, `${hid}:1`, `${hid}:2`, `${hid}:0`]);
    expect(ids[5]).toBe('0:0');
    const n2 = await numXml(b2);
    expect(n2).toMatch(/<w:multiLevelType w:val="multilevel"\/><w:lvl w:ilvl="0"><w:start w:val="1"\/><w:numFmt w:val="decimal"\/><w:lvlText w:val="%1."\/>/);
    expect(n2).toContain('<w:lvlText w:val="%1.%2."/>');
    const back = (await importDocx(b2)).doc;
    expect(back.attrs.headingNumbering.levels.slice(0, 2).map((l) => l.text)).toEqual(['%1.', '%1.%2.']);
    expect(back.attrs.headingNumbering.source.viaStyle).toBe(false);
    expect(back.content.filter((n) => n.type === 'heading').map((h) => !!h.attrs.noNumber)).toEqual([false, false, false, false, false, true]);
    // отключена: у всех заголовков numId 0
    const off = JSON.parse(JSON.stringify(doc));
    off.attrs.headingNumbering = { ...doc.attrs.headingNumbering, off: true };
    const x3 = await docXml(await exportDocx(off, { source: src }));
    expect((x3.match(/<w:numId w:val="0"\/>/g) || []).length).toBe(6);
    // форматы списков переживают сохранение (и в новый файл)
    for (const bytes of [b2, await exportDocx(changed)]) {
      const lists = (await importDocx(bytes)).doc.content.filter((n) => n.type === 'orderedList').map((n) => n.attrs);
      expect(lists).toEqual([{ start: 1, numFmt: 'lowerLetter', lvlText: '%1)' }, { start: 1, numFmt: 'upperRoman' }]);
    }
  });

  it('вид маркера списка по шаблону', () => {
    expect(markerKind('%1.')).toBe(null);
    expect(markerKind('%1)')).toBe('paren');
    expect(markerKind('(%1)')).toBe('both');
    expect(markerKind('%1.%2.')).toBe('path-dot');
  });
});
