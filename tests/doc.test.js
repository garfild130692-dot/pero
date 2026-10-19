// Документы Word 97–2003 (.doc, этап 19.4): свой разбор [MS-DOC] → пакет .docx → обычный импорт. Эталоны —
// tests/make_doc_fixtures.ps1: сам Word пересохраняет эталонные .docx в .doc, а .doc — снова в .docx (<имя>.doc.docx).
// Модель Перо из нашего разбора .doc должна совпасть с моделью из .doc.docx — так, как сам Word понял свой .doc.
import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { readCfb, writeCfb } from '../src/formats/cfb.js';
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

// ---------- после ревью: связи частей, испорченные файлы, внешние ссылки ----------
describe('.doc: сторожа после ревью (связи, порча, внешние ссылки)', () => {
  const NAMES = readdirSync('tests/fixtures').filter((f) => /\.doc$/.test(f)).map((f) => f.slice(0, -4)).filter((n) => n !== 'big' || existsSync('tests/fixtures/big.doc'));
  const load = async (name) => JSZip.loadAsync((await docToDocx(readFileSync(`tests/fixtures/${name}.doc`))).docx);
  const norm = (base, target) => {
    const parts = (target.startsWith('/') ? target.slice(1) : `${base}/${target}`).split('/'); const out = [];
    for (const s of parts) { if (s === '..') out.pop(); else if (s && s !== '.') out.push(s); }
    return out.join('/');
  };

  it('связи у каждой части: r:id / r:embed есть в её собственных .rels, картинки лежат в пакете; все XML-части разбираются', async () => {
    expect(NAMES.length).toBeGreaterThan(15);
    for (const name of NAMES) {
      const zip = await load(name);
      for (const path of Object.keys(zip.files)) {
        if (!/\.(xml|rels)$/.test(path) || zip.files[path].dir) continue;
        const xml = await zip.file(path).async('string');
        const errs = [];
        new DOMParser({ onError: (lvl, msg) => { if (lvl !== 'warning') errs.push(`${lvl}: ${msg}`); } }).parseFromString(xml, 'text/xml');
        expect(errs, `${name}: ${path}`).toEqual([]);
        if (!/^word\/[^/]+\.xml$/.test(path)) continue;
        const ids = new Set([...xml.matchAll(/\br:(?:id|embed)="([^"]*)"/g)].map((m) => m[1]));
        if (!ids.size) continue;
        const relsFile = zip.file(`word/_rels/${path.slice(5)}.rels`);
        expect(relsFile, `${name}: у ${path} есть ссылки ${[...ids]}, но нет .rels`).toBeTruthy();
        const rels = new Map([...(await relsFile.async('string')).matchAll(/<Relationship\b[^>]*>/g)].map((m) => [/\bId="([^"]*)"/.exec(m[0])[1], m[0]]));
        for (const id of ids) {
          expect(rels.has(id), `${name}: ${path} → ${id} нет в своих связях`).toBe(true);
          const tag = rels.get(id);
          if (/relationships\/image"/.test(tag) && !/TargetMode="External"/.test(tag)) {
            const t = /\bTarget="([^"]*)"/.exec(tag)[1];
            expect(zip.file(norm('word', t)), `${name}: ${path} → картинка ${t}`).toBeTruthy();
          }
        }
      }
    }
  });

  // --- порча ---
  const rng = (seed) => () => { seed = (seed + 0x6D2B79F5) >>> 0; let t = seed; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  const rebuild = (cfb, over) => writeCfb(cfb.entries.filter((e) => e.type === 'stream').map((e) => ({ path: e.path, data: over[e.path] ? over[e.path] : e.data })));
  const dv = (a) => new DataView(a.buffer, a.byteOffset, a.byteLength);
  const tableName = (w) => (dv(w).getUint16(10, true) & 0x0200 ? '1Table' : '0Table');
  // смещение Pcdt (0x02) в Clx
  const clxPos = (w, T) => {
    const fc = dv(w).getUint32(154 + 33 * 8, true); const lcb = dv(w).getUint32(158 + 33 * 8, true);
    let p = fc;
    while (p < fc + lcb && T[p] === 0x01) p += 3 + dv(T).getInt16(p + 1, true);
    return { fc, lcb, pcdt: p };
  };

  it('испорченный файл не вешает разбор: ≥ 60 детерминированных порч — результат или Error, каждая < 3 с', async () => {
    const cases = [];
    for (const base of ['sample', 'preserve']) {
      const src = readFileSync(`tests/fixtures/${base}.doc`);
      const cfb = readCfb(src);
      const w0 = cfb.stream('WordDocument'); const tn = tableName(w0);
      const T0 = cfb.stream(tn); const D0 = cfb.stream('Data');
      const R = rng(base === 'sample' ? 12345 : 67890);
      const mk = (label, fn) => cases.push([`${base}: ${label}`, fn]);
      // огромные lcb в парах FIB
      for (let i = 0; i <= 90; i++) {
        mk(`FIB lcb[${i}] огромный`, () => { const w = w0.slice(); dv(w).setUint32(154 + i * 8 + 4, i % 2 ? 0xFFFFFFFF : 0x7FFFFFF0, true); return rebuild(cfb, { WordDocument: w }); });
      }
      for (let i = 0; i <= 90; i += 6) {
        mk(`FIB fc[${i}] огромный`, () => { const w = w0.slice(); dv(w).setUint32(154 + i * 8, 0xFFFFFFF0, true); return rebuild(cfb, { WordDocument: w }); });
      }
      mk('Clx: отрицательный cb у Prc', () => {
        const T = T0.slice(); const { fc } = clxPos(w0, T0);
        T[fc] = 0x01; dv(T).setInt16(fc + 1, -2, true);
        return rebuild(cfb, { [tn]: T });
      });
      mk('Clx: cb Prc = 0x7FFF', () => { const T = T0.slice(); const { fc } = clxPos(w0, T0); T[fc] = 0x01; dv(T).setInt16(fc + 1, 0x7FFF, true); return rebuild(cfb, { [tn]: T }); });
      for (const v of [0, 3, 0x7FFFFFFF, 0xFFFFFFFF, 0xFFFFFFF8]) {
        mk(`Pcdt lcb = ${v.toString(16)}`, () => { const T = T0.slice(); const { pcdt } = clxPos(w0, T0); dv(T).setUint32(pcdt + 1, v, true); return rebuild(cfb, { [tn]: T }); });
      }
      mk('Pcdt: CP кусков вразнобой', () => { const T = T0.slice(); const { pcdt } = clxPos(w0, T0); for (let k = 0; k < 12; k++) T[pcdt + 5 + Math.floor(R() * 24)] = Math.floor(R() * 256); return rebuild(cfb, { [tn]: T }); });
      // мусор в Data
      for (let k = 0; k < 3; k++) {
        mk(`Data: мусор в первых 512 байтах #${k}`, () => { const D = (D0 || new Uint8Array(1024)).slice(); for (let j = 0; j < Math.min(512, D.length); j++) D[j] = Math.floor(R() * 256); return rebuild(cfb, { Data: D }); });
      }
      mk('Data: 0xFF в первых 512', () => { const D = (D0 || new Uint8Array(1024)).slice(); D.fill(0xFF, 0, Math.min(512, D.length)); return rebuild(cfb, { Data: D }); });
      // обрезанные файлы
      for (let k = 0; k < 10; k++) { const n = Math.floor(R() * src.length); mk(`обрезан до ${n}`, () => src.slice(0, n)); }
      mk('обрезан до 512', () => src.slice(0, 512)); mk('обрезан до 513', () => src.slice(0, 513)); mk('обрезан на 1 байт', () => src.slice(0, src.length - 1));
      // заголовок CFB
      for (const v of [0, 1, 8, 9, 12, 13, 16, 31]) mk(`CFB shift = ${v}`, () => { const b = src.slice(); dv(b).setUint16(30, v, true); return b; });
      for (const v of [0, 1, 0xFFFF, 0x7FFFFFFF, 0xFFFFFFFF]) mk(`CFB nFat = ${v.toString(16)}`, () => { const b = src.slice(); dv(b).setUint32(44, v, true); return b; });
      for (const [o, v] of [[48, 0xFFFFFFFE], [48, 0], [48, 0x7FFFFFFF], [60, 0], [64, 0xFFFFFFFF], [68, 0], [68, 0xFFFFFFFC], [72, 0xFFFFFFFF], [72, 1000000], [56, 0], [32, 0], [32, 20]]) {
        mk(`CFB заголовок [${o}] = ${v.toString(16)}`, () => { const b = src.slice(); dv(b).setUint32(o, v, true); return b; });
      }
      for (let k = 0; k < 8; k++) mk(`CFB случайные байты заголовка #${k}`, () => { const b = src.slice(); for (let j = 0; j < 6; j++) b[8 + Math.floor(R() * 504)] = Math.floor(R() * 256); return b; });
      mk('CFB: DIFAT зацикленный', () => { const b = src.slice(); dv(b).setUint32(72, 5, true); dv(b).setUint32(68, 0, true); return b; });
    }
    expect(cases.length).toBeGreaterThanOrEqual(60);
    const slow = []; let ok = 0; let errs = 0;
    for (const [label, make] of cases) {
      const t0 = Date.now();
      try { await docToDocx(make()); ok++; } catch (e) {
        errs++;
        expect(e instanceof Error, `${label}: брошено не Error: ${e}`).toBe(true);
      }
      const dt = Date.now() - t0;
      if (dt >= 3000) slow.push(`${label}: ${dt} мс`);
    }
    expect(slow).toEqual([]);
    expect(ok + errs).toBe(cases.length);
  }, 60000);

  // --- точечная порча структур (второй пакет сторожей) ---
  const open = (name) => {
    const src = readFileSync(`tests/fixtures/${name}.doc`); const cfb = readCfb(src);
    const w = cfb.stream('WordDocument'); const tn = tableName(w);
    return { cfb, w, tn, T: cfb.stream(tn).slice(), pair: (i) => [dv(w).getUint32(154 + i * 8, true), dv(w).getUint32(158 + i * 8, true)] };
  };
  // результат или Error (не другое), не дольше 3 с; вернёт { docx, ms } или { err }
  const run = async (label, bytes) => {
    const t0 = Date.now(); let res;
    try { res = { docx: (await docToDocx(bytes)).docx }; } catch (e) {
      expect(e instanceof Error, `${label}: брошено не Error: ${e}`).toBe(true); res = { err: e };
    }
    res.ms = Date.now() - t0;
    expect(res.ms, `${label}: ${res.ms} мс`).toBeLessThan(3000);
    return res;
  };
  const docXmlSize = async (docx) => (await (await JSZip.loadAsync(docx)).file('word/document.xml').async('string')).length;

  it('истории за пределами текста (PlcffndTxt, PlcfandTxt): CP огромные и отрицательные — быстро, основной текст не дублируется', async () => {
    const base = await docToDocx(readFileSync('tests/fixtures/preserve.doc'));
    const size0 = await docXmlSize(base.docx);
    for (const i of [3, 5]) {
      for (const v of [0x7FFFFFF0, -1, -100000]) {
        for (const which of ['последний', 'все']) {
          const { cfb, w, tn, T, pair } = open('preserve');
          const [fc, lcb] = pair(i);
          expect(lcb, `в preserve.doc нет пары ${i}`).toBeGreaterThanOrEqual(12);
          const count = Math.floor(lcb / 4); // PlcfTxt без данных: n + 1 значений CP
          const idx = which === 'все' ? [...Array(count).keys()] : [count - 1];
          for (const k of idx) dv(T).setInt32(fc + k * 4, v, true);
          const label = `пара ${i}, CP=${v}, ${which}`;
          const r = await run(label, rebuild(cfb, { [tn]: T }));
          if (r.docx) expect(await docXmlSize(r.docx), `${label}: document.xml разросся`).toBeLessThanOrEqual(size0 * 2);
        }
      }
    }
  }, 60000);

  it('повтор страниц FKP: все pn в PlcfBteChpx/Papx одинаковые — быстро', async () => {
    for (const name of ['sample', 'preserve', 'tables']) {
      for (const i of [12, 13]) {
        const { cfb, tn, T, pair } = open(name);
        const [fc, lcb] = pair(i);
        const n = Math.floor((lcb - 4) / 8);
        expect(n, `${name}: нет страниц FKP (пара ${i})`).toBeGreaterThan(0);
        const pn = dv(T).getUint32(fc + (n + 1) * 4, true);
        for (let k = 0; k < n; k++) dv(T).setUint32(fc + (n + 1) * 4 + k * 4, pn, true);
        await run(`${name}: пара ${i}, pn=${pn}`, rebuild(cfb, { [tn]: T }));
      }
    }
  }, 60000);

  it('STSH: cstd = 4000 при короткой таблице стилей — быстро', async () => {
    for (const name of ['sample', 'preserve']) {
      const { cfb, tn, T, pair } = open(name);
      const [fc] = pair(1);
      dv(T).setUint16(fc + 2, 4000, true);
      await run(`${name}: cstd=4000`, rebuild(cfb, { [tn]: T }));
    }
  });

  it('надпись внутри себя: диапазон FTXBXS = основной текст — без переполнения стека, быстро', async () => {
    const { cfb, w, tn, T, pair } = open('shapes');
    const [fc, lcb] = pair(56);
    const n = Math.floor((lcb - 4) / 26);
    expect(n, 'в shapes.doc нет надписей (FTXBXS)').toBeGreaterThan(0);
    const ccp = (o) => dv(w).getUint32(o, true);
    const sum = [76, 80, 84, 92, 96].reduce((a, o) => a + ccp(o), 0); // начало истории надписей
    const ccpText = ccp(76);
    for (const [label, a, b] of [['[0,ccpText) через смещение', -sum, -sum + ccpText], ['всё от 0 до конца', -sum, -sum + sum + ccp(100)]]) {
      const T2 = T.slice();
      for (let k = 0; k <= n; k++) dv(T2).setInt32(fc + k * 4, k === 0 ? a : k === 1 ? b : b, true);
      const r = await run(`shapes: ${label}`, rebuild(cfb, { [tn]: T2 }));
      if (r.err) expect(String(r.err.message)).not.toMatch(/call stack/i);
    }
  });

  it('выравнивание фигуры: posh/posv = 1000 — нет «<wp:align>undefined»', async () => {
    let patched = 0;
    for (const name of ['shapes', 'wrap']) {
      const { cfb, tn, T, pair } = open(name);
      const [fc, lcb] = pair(50);
      // свойства 0x038F / 0x0391 в OPT (F00B) DggInfo: обход записей как artRecords
      const walk = (x, y) => {
        while (x + 8 <= y) {
          const vi = dv(T).getUint16(x, true); const type = dv(T).getUint16(x + 2, true); const len = dv(T).getUint32(x + 4, true);
          if (type < 0xF000 || type > 0xF200 || x + 8 + len > y) { x++; continue; }
          if ((vi & 0xF) === 0xF) walk(x + 8, x + 8 + len);
          else if (type === 0xF00B) {
            for (let i = 0; i < (vi >> 4); i++) {
              const id = dv(T).getUint16(x + 8 + i * 6, true) & 0x3FFF;
              const at = x + 8 + i * 6;
              // у эталонов Word posh/posv в OPT нет (положение — в FSPA): подменяем нигде не читаемые простые свойства
              // (0x7F, 0x80, 0x13F, 0xBF) на 0x038F / 0x0391 со значением 1000
              if (id === 0x038F || id === 0x0391) { dv(T).setUint32(at + 2, 1000, true); patched++; }
              else if (!(dv(T).getUint16(at, true) & 0x8000) && [0x7F, 0x80, 0x13F, 0xBF].includes(id)) {
                dv(T).setUint16(at, patched % 2 ? 0x0391 : 0x038F, true); dv(T).setUint32(at + 2, 1000, true); patched++;
              }
            }
          }
          x += 8 + len;
        }
      };
      walk(fc, fc + lcb);
      const r = await run(`${name}: posh=1000`, rebuild(cfb, { [tn]: T }));
      if (r.docx) {
        const xml = await (await JSZip.loadAsync(r.docx)).file('word/document.xml').async('string');
        expect(xml, `${name}`).not.toMatch(/<wp:align>undefined/);
        expect(xml).not.toMatch(/<wp:align>(?!left|center|right|inside|outside|top|bottom)/);
      }
    }
    expect(patched, 'в shapes.doc/wrap.doc нет свойств posh/posv в OPT — случай не проверяет ничего').toBeGreaterThan(0);
  });

  it('внешние связи document.xml.rels — только http(s):// и mailto:', async () => {
    for (const name of NAMES) {
      const zip = await load(name);
      const f = zip.file('word/_rels/document.xml.rels');
      if (!f) continue;
      for (const m of (await f.async('string')).matchAll(/<Relationship\b[^>]*TargetMode="External"[^>]*>/g)) {
        const t = /\bTarget="([^"]*)"/.exec(m[0])[1];
        expect(t, `${name}: ${m[0]}`).toMatch(/^(https?:\/\/|mailto:)/i);
      }
    }
  });
});

describe('.doc: эталон extra (после ревью: колонтитул с рисунком, ссылки, сноска со ссылкой, широкая таблица)', () => {
  it('как у Word, кроме ссылок на файл и UNC — у нас полем HYPERLINK, как в .doc; связи колонтитула и сносок — в их .rels; sprmPHugePapx', async () => {
    const { out } = await compare('extra');
    // единственное отличие — первый абзац: «файл» и «сеть» — поле HYPERLINK (как в исходном .doc), а не w:hyperlink
    // с внешней связью: ссылка не теряется, а в Пере поле по щелчку не открывается
    expect(out.every((x) => x.startsWith('doc.content.0.content'))).toBe(true);
    const { docx, warnings } = await docToDocx(readFileSync('tests/fixtures/extra.doc'));
    expect(Object.keys(warnings)).toEqual([]);
    const z = await JSZip.loadAsync(docx);
    const relsOf = async (part) => { const f = z.file(`word/_rels/${part}.rels`); return f ? f.async('string') : ''; };
    const headers = Object.keys(z.files).filter((n) => /^word\/header_doc\d+\.xml$/.test(n));
    const withPic = [];
    for (const h of headers) if (/r:embed=/.test(await z.file(h).async('string'))) withPic.push(h.slice(5));
    expect(withPic.length).toBeGreaterThan(0);
    for (const h of withPic) expect(await relsOf(h)).toMatch(/Type="[^"]*\/image" Target="media\//);
    expect(await relsOf('footnotes.xml')).toMatch(/https:\/\/example\.org\//);
    expect(await relsOf('document.xml')).not.toMatch(/file:|\\server/);
    const body = await z.file('word/document.xml').async('string');
    expect(body).toMatch(/<w:instrText[^>]*>[^<]*HYPERLINK[^<]*file:\/\/\/C:\/Windows\/win\.ini/);
    expect(body).toMatch(/<w:instrText[^>]*>[^<]*HYPERLINK[^<]*server[^<]*share/); // UNC Word записал как file:///\\server…
  });
});
