// Документы Word 97–2003 (.doc, этап 19.4): свой разбор [MS-DOC] → пакет .docx → обычный импорт. Эталоны —
// tests/make_doc_fixtures.ps1: сам Word пересохраняет эталонные .docx в .doc, а .doc — снова в .docx (<имя>.doc.docx).
// Модель Перо из нашего разбора .doc должна совпасть с моделью из .doc.docx — так, как сам Word понял свой .doc.
import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { docToDocx, isDoc } from '../src/docx/doc.js';
import { importDocx } from '../src/docx/import.js';

// расхождения двух моделей; служебное, что Word пишет по-своему, не сравнивается:
// id стилей (у Word «aff8» — сверяются имена), rsid, номера закладок и связей в sectPr;
// рисунки, фигуры и формулы — у Word после пересохранения .doc это VML (w:pict), у нас — настоящие объекты
function diff(a, b) {
  const out = [];
  let vml = 0;
  const rs = (v, k) => (typeof v !== 'string' ? v : v.replace(/ w:rsid\w*="[^"]*"/g, '').replace(/(<w:bookmark(?:Start|End)[^>]*) w:id="\d+"/g, '$1')
    .replace(/w:rStyle w:val="[^"]*"/g, '').replace(/sectPr/.test(k) ? / r:id="[^"]*"/g : /$^/, ''));
  const cmp = (x0, y0, path) => {
    const x = rs(x0, path); const y = rs(y0, path);
    if (JSON.stringify(x) === JSON.stringify(y)) return;
    if (/fileStyle\.id$|[sS]tyles\.\d+\.id$|\.styleId$|marks\.\d+\.attrs\.id$/.test(path)) return;
    if (x && y && (x.type === 'image' || x.type === 'opaque') && y.type === 'opaque' && /<w:pict>/.test(y.attrs.xml || '') && !/<w:pict>/.test((x.attrs && x.attrs.xml) || '')) { vml++; return; }
    if (x && y && typeof x === 'object' && typeof y === 'object' && Array.isArray(x) === Array.isArray(y) && (!Array.isArray(x) || x.length === y.length)) {
      for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) cmp(x[k], y[k], `${path}.${k}`);
      return;
    }
    out.push(`${path}: ${JSON.stringify(x)?.slice(0, 160)} ≠ ${JSON.stringify(y)?.slice(0, 160)}`);
  };
  cmp(a, b, 'doc');
  return { out, vml };
}
async function compare(name) {
  const { docx, warnings } = await docToDocx(readFileSync(`tests/fixtures/${name}.doc`));
  const a = (await importDocx(docx)).doc;
  const b = (await importDocx(readFileSync(`tests/fixtures/${name}.doc.docx`))).doc;
  return { ...diff(a, b), warnings, a };
}
// объекты документа (рисунки и фигуры) — для сверки с исходным .docx
const objects = (d) => {
  const out = [];
  const walk = (n) => {
    if (n.type === 'image') out.push({ w: Math.round(n.attrs.width), h: Math.round(n.attrs.height), anchor: n.attrs.anchor, rotate: n.attrs.rotate, flipH: n.attrs.flipH });
    if (n.type === 'opaque' && /wps:wsp/.test(n.attrs.xml || '')) {
      const ext = (/<wp:extent cx="(\d+)" cy="(\d+)"/.exec(n.attrs.xml) || []).slice(1).map((v) => Math.round(v / 12700));
      out.push({ shape: (/prst="(\w+)"/.exec(n.attrs.xml) || [])[1], ext, wrap: (/<wp:(wrap\w+)/.exec(n.attrs.xml) || [])[1], text: /txbx/.test(n.attrs.xml) });
    }
    (n.content || []).forEach(walk);
  };
  walk(d);
  return out;
};

describe('.doc: чтение (свой разбор [MS-DOC])', () => {
  it('распознаёт документ Word 97–2003 по потоку WordDocument', () => {
    expect(isDoc(readFileSync('tests/fixtures/sample.doc'))).toBe(true);
    expect(isDoc(readFileSync('tests/fixtures/book.xls'))).toBe(false);
  });
  const SAME = [
    ['sample', 'текст, заголовки, Ж/К/Ч, выравнивание, списки стилями, таблица'],
    ['fonts', 'шрифты стиля «Обычный», прямое форматирование, заливка текста, символьные стили'],
    ['para', 'интервалы, отступы, межстрочный, пользовательский стиль абзаца'],
    ['tables', 'ширины, границы (точные цвета 0xD61A–D), заливка ячеек, высота строк, шапка, выравнивание'],
    ['layout', 'положение на странице, рамки и заливка абзаца, табуляция стиля (sprmPChgTabs)'],
    ['lists', 'маркеры (Wingdings), продолжение нумерации'],
    ['numbering', 'нумерация заголовков, списки a) и I.'],
    ['hf', 'колонтитулы первой страницы и чётных/нечётных (PlcfHdd)'],
    ['sections', 'разделы с разной ориентацией (PlcfSed + SEPX)'],
    ['footnotes', 'сноски (PlcffndRef/Txt)'],
    ['preserve', 'сноска, примечание (автор, дата AtrdExtra, диапазон), закладка, поле, формула, символ, разрывы, свойства документа'],
    ['toc', 'оглавление: поле TOC, гиперссылки-якоря (поле HYPERLINK → w:hyperlink), PAGEREF, закладки _Toc'],
    ['columns', 'колонки: равные, с разделителем, разной ширины'],
    ['track', 'исправления: вставка и удаление с автором и датой'],
    ['design', 'цвет страницы (фоновая фигура), границы страниц, подложка (WordArt → VML)'],
    ['linenum', 'нумерация строк'],
    ['wrap', 'плавающие рисунки (OfficeArt: FSPA, BStore, OPT) — у Word VML'],
    ['crop', 'обрезка (PICF, srcRect), поворот с отражением'],
    ['shapes', 'фигуры и надписи (wps:wsp, текст надписей) — у Word VML'],
  ];
  for (const [name, what] of SAME) {
    it(`${name}: ${what}`, async () => {
      const { out } = await compare(name);
      expect(out).toEqual([]);
    });
  }
  it('math: формулы (OMML из свойств картинки, поле QUOTE в строке) — как в исходном .docx', async () => {
    const pick = (d) => { const o = []; const w = (n) => { if (n.type === 'opaque' && n.attrs.kind === 'math') o.push(n.attrs.xml.replace(/\s+xmlns:\w+="[^"]*"/g, '').replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '').replace(/<m:ctrlPr>[\s\S]*?<\/m:ctrlPr>|<m:ctrlPr\/>/g, '')); (n.content || []).forEach(w); }; w(d); return o; };
    const a = (await importDocx((await docToDocx(readFileSync('tests/fixtures/math.doc'))).docx)).doc;
    const b = (await importDocx(readFileSync('tests/fixtures/math.docx'))).doc;
    expect(pick(a)).toHaveLength(9);
    expect(pick(a)).toEqual(pick(b));
  });
  it('плавающие рисунки и фигуры — те же размеры, обтекание, положение, что в исходном .docx (до твипа)', async () => {
    for (const name of ['wrap', 'shapes', 'crop']) {
      const a = objects((await importDocx((await docToDocx(readFileSync(`tests/fixtures/${name}.doc`))).docx)).doc);
      const b = objects((await importDocx(readFileSync(`tests/fixtures/${name}.docx`))).doc);
      // поворот на 90° Word сам теряет при сохранении в .doc (в его .doc.docx рисунок не повёрнут) — последний рисунок crop
      const n = name === 'crop' ? 2 : b.length;
      expect(a.slice(0, n).map((x) => JSON.stringify(x).replace(/("off":-?\d+\.\d)\d*/g, '$1')), name).toEqual(b.slice(0, n).map((x) => JSON.stringify(x).replace(/("off":-?\d+\.\d)\d*/g, '$1')));
    }
  });
  it.runIf(existsSync('tests/fixtures/big.doc'))('big.doc (234 стр., 10 рисунков; создаётся make_doc_fixtures.ps1 big, в репозиторий не кладётся)', async () => {
    const t0 = Date.now();
    const { out } = await compare('big');
    expect(out).toEqual([]);
    expect(Date.now() - t0).toBeLessThan(15000);
  });
});
