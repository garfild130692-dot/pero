// Книги Excel 97–2003 (.xls, этап 19): контейнер [MS-CFB] и разбор BIFF8. Эталоны — tests/make_xls_fixtures.ps1:
// сам Excel пересохраняет book/preserve/formulas.xlsx в .xls; модели из обоих файлов должны совпасть.
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { readCfb, writeCfb } from '../src/formats/cfb.js';
import { importXls, decompile } from '../src/sheet/xls.js';
import { exportXls } from '../src/sheet/xlsWrite.js';
import { importXlsx } from '../src/sheet/xlsx.js';

const sortK = (x) => (Array.isArray(x) ? x.map(sortK) : x && typeof x === 'object' ? Object.fromEntries(Object.keys(x).sort().map((k) => [k, sortK(x[k])])) : x);
const J = (x) => JSON.stringify(sortK(x ?? null));
const noSpace = (f) => (f == null ? f : String(f).replace(/("[^"]*")|\s+/g, (m, q) => q || ''));
const noXf = (o) => o && Object.fromEntries(Object.entries(o).map(([k, v]) => { const { xf, ...rest } = v; return [k, rest]; }).filter(([, v]) => Object.keys(v).length));
const oneCell = (ref) => !!ref && ref.split(':')[0] === ref.split(':')[1];

// расхождения моделей книги (allow(лист, ключ, что) — известные потери самого Excel при сохранении в .xls)
async function diff(name, allow = () => false) {
  const a = (await importXls(readFileSync(`tests/fixtures/${name}.xls`))).wb;
  const b = (await importXlsx(readFileSync(`tests/fixtures/${name}.xlsx`))).wb;
  const out = [];
  b.sheets.forEach((sb, i) => {
    const sa = a.sheets[i];
    if (!sa || sa.name !== sb.name) { out.push(`лист ${sb.name}`); return; }
    for (const k of new Set([...sa.cells.keys(), ...sb.cells.keys()])) {
      const x = sa.cells.get(k) || {}; const y = sb.cells.get(k) || {};
      const bad = (what) => { if (!allow(sb.name, k, what)) out.push(`${sb.name}!${k} ${what}: ${J(what === 'формула' ? x.f : what === 'значение' ? x.v : x.s)} ≠ ${J(what === 'формула' ? y.f : what === 'значение' ? y.v : y.s)}`); };
      if (noSpace(x.f) !== noSpace(y.f)) bad('формула');
      else if (J(x.v) !== J(y.v) && !(typeof x.v === 'number' && typeof y.v === 'number' && Math.abs(x.v - y.v) <= 1e-9 * Math.max(1, Math.abs(y.v)))) bad('значение');
      if (J(x.s) !== J(y.s)) bad('стиль');
      if (J(x.cse) !== J(y.cse) && !(oneCell(x.cse) && !y.cse)) bad('массив');
    }
    for (const p of ['colW', 'rowH', 'merges', 'freeze', 'page', 'printArea', 'printTitles', 'protection']) if (J(sa[p]) !== J(sb[p])) out.push(`${sb.name} ${p}: ${J(sa[p])} ≠ ${J(sb[p])}`);
    for (const p of ['colInfo', 'rowInfo']) if (J(noXf(sa[p])) !== J(noXf(sb[p]))) out.push(`${sb.name} ${p}`);
  });
  if (J(a.names) !== J(b.names)) out.push(`имена: ${J(a.names)} ≠ ${J(b.names)}`);
  if (a.active !== b.active) out.push('активный лист');
  return out;
}

describe('.xls: составной файл [MS-CFB]', () => {
  it('читает книги Excel и записывает обратно байт в байт (мини-поток, вложенные хранилища, DIFAT)', () => {
    const c = readCfb(readFileSync('tests/fixtures/preserve.xls'));
    expect(c.entries.map((e) => e.path)).toContain('Workbook');
    const files = c.entries.filter((e) => e.type === 'stream').map((e) => ({ path: e.path, data: e.data }));
    files.push({ path: 'Большой', data: new Uint8Array(8e6).map((_, i) => i % 251) }); // > 109 секторов FAT — нужен DIFAT
    const back = readCfb(writeCfb(files));
    for (const f of files) expect(Buffer.compare(Buffer.from(back.stream(f.path)), Buffer.from(f.data)), f.path).toBe(0);
  });
});

describe('.xls: чтение книги (BIFF8)', () => {
  it('book: значения, общие формулы, стили (цвета XFEXT и темы), объединение, ширины — как из .xlsx', async () => {
    expect(await diff('book')).toEqual([]);
  });
  it('preserve: имена, закрепление, скрытие и группировка, параметры страницы, область печати, защита', async () => {
    expect(await diff('preserve')).toEqual([]);
  });
  it('formulas: 870 формул всех групп функций (Ptg → текст), функции новее Excel 2003 и пакета анализа', async () => {
    // LET Excel в .xls не пишет (формула становится константой #VALUE!); СЛЧИС/СЛУЧМЕЖДУ — случайные
    const random = new Set(['358,2', '359,2']);
    const out = await diff('formulas', (sh, k, what) => (what === 'формула' && ['354,2', '491,2'].includes(k)) || (what === 'значение' && random.has(k)));
    expect(out).toEqual([]);
  });
  it('байт-код формулы: операторы, скобки, функции с постоянным и переменным числом аргументов, ссылки', () => {
    // =ROUND((A1+$B$2)*-C3%,2) & "x"
    const rgce = Uint8Array.from([
      0x24, 0, 0, 0, 0xC0, // A1 (относительная)
      0x24, 1, 0, 1, 0x00, // $B$2
      0x03, 0x15, // + ( )
      0x24, 2, 0, 2, 0xC0, 0x14, 0x13, // C3 % унарный минус
      0x05, // *
      0x1E, 2, 0, // 2
      0x21, 27, 0, // ROUND (2 аргумента)
      0x17, 1, 0, 0x78, // "x"
      0x08, // &
    ]);
    expect(decompile(rgce, new Uint8Array(0), { name: () => '', sheetRef: () => null, externName: () => '' })).toBe('ROUND((A1+$B$2)*-C3%,2)&"x"');
  });
});

// запись: модель из .xlsx → exportXls → importXls — та же модель (что Excel открывает наш файл без восстановления и
// считает те же значения, проверяется вручную: tests/xls_excel_dump.ps1 — book 11/11, preserve 15/15, formulas 862/871
// формул, остальные — LET, функции новее Excel 2021, ROWS(A:A) = 65536 строк формата и случайные)
async function roundTrip(name, allow = () => false) {
  const b = (await importXlsx(readFileSync(`tests/fixtures/${name}.xlsx`))).wb;
  const { bytes, lost } = exportXls(b);
  const a = (await importXls(bytes)).wb;
  const out = [];
  b.sheets.forEach((sb, i) => {
    const sa = a.sheets[i];
    if (!sa || sa.name !== sb.name) { out.push(`лист ${sb.name}`); return; }
    for (const k of new Set([...sa.cells.keys(), ...sb.cells.keys()])) {
      const x = sa.cells.get(k) || {}; const y = sb.cells.get(k) || {};
      if (noSpace(x.f) !== noSpace(y.f)) { if (!allow(k, 'формула')) out.push(`${sb.name}!${k} формула: ${J(x.f)} ≠ ${J(y.f)}`); }
      else if (J(x.v) !== J(y.v) && !(typeof x.v === 'number' && typeof y.v === 'number' && Math.abs(x.v - y.v) <= 1e-9 * Math.max(1, Math.abs(y.v)))) { if (!allow(k, 'значение')) out.push(`${sb.name}!${k} значение: ${J(x.v)} ≠ ${J(y.v)}`); }
      if (J(x.s) !== J(y.s)) out.push(`${sb.name}!${k} стиль: ${J(x.s)} ≠ ${J(y.s)}`);
    }
    for (const p of ['colW', 'rowH', 'merges', 'freeze', 'page', 'printArea', 'printTitles', 'protection']) if (J(sa[p]) !== J(sb[p])) out.push(`${sb.name} ${p}: ${J(sa[p])} ≠ ${J(sb[p])}`);
    for (const p of ['colInfo', 'rowInfo']) if (J(noXf(sa[p])) !== J(noXf(sb[p]))) out.push(`${sb.name} ${p}`);
  });
  if (J(a.names) !== J(b.names)) out.push(`имена: ${J(a.names)} ≠ ${J(b.names)}`);
  if (a.active !== b.active) out.push('активный лист');
  return { out, lost, bytes };
}

describe('.xls: запись книги (BIFF8)', () => {
  it('book: значения, формулы, стили, объединение, ширины — туда и обратно без расхождений', async () => {
    const { out, lost } = await roundTrip('book');
    expect(out).toEqual([]);
    expect(lost).toEqual({});
  });
  it('preserve: имена, закрепление, скрытие и группировка, параметры страницы, область печати, защита', async () => {
    const { out, lost } = await roundTrip('preserve');
    expect(out).toEqual([]);
    expect(Object.keys(lost).sort()).toEqual(['проверка данных', 'условное форматирование']);
  });
  it('formulas: 870 формул — классы операндов, функции новее Excel 2003 (_xlfn), формулы массива', async () => {
    // LET в .xls нет — записывается значением; СЛЧИС/СЛУЧМЕЖДУ — случайные
    const { out, lost, bytes } = await roundTrip('formulas', (k, what) => (what === 'формула' && ['354,2', '491,2'].includes(k)) || (what === 'значение' && ['358,2', '359,2'].includes(k)));
    expect(out).toEqual([]);
    expect(lost).toEqual({ 'формулы, которых нет в Excel 97–2003 (записаны значением)': 2 });
    // динамические формулы — ARRAY на одну ячейку с токенами класса A (как пишет Excel): =СУММ(ФИЛЬТР(…))
    const s = readCfb(bytes).stream('Workbook');
    let arrays = 0;
    for (let p = 0; p + 4 <= s.length;) { const t = s[p] | (s[p + 1] << 8); const l = s[p + 2] | (s[p + 3] << 8); if (t === 0x0221) arrays++; p += 4 + l; }
    expect(arrays).toBeGreaterThan(10);
  });
});
