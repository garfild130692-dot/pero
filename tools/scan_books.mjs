// Сканер книг компании (этап 12.4): какие функции встречаются в рабочих .xlsx и насколько Перо считает их как Excel.
//
//   node tools/scan_books.mjs <папка или файл> [ещё…] [--out отчёт.md] [--max-examples 30]
//   npm run scan -- "\\сервер\Общие\Отчёты" --out scan.md
//
// Для каждой книги: функции формул и имён (сколько раз, в скольких книгах), функции, которых нет в Перо,
// и сверка: значения формул, сохранённые самим Excel в файле, против полного пересчёта Перо.
// Книги только читаются; отчёт — Markdown (по умолчанию в консоль). .xls (старый двоичный формат) не читается —
// такие файлы только перечисляются.
import { readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { basename, extname, join, relative } from 'node:path';
import { importXlsx } from '../src/sheet/xlsx.js';
import { recalc, VOLATILE } from '../src/sheet/model.js';
import { FUNCTIONS, tokenize, localize, ruName, errText, isErr } from '../src/sheet/formula.js';

const args = process.argv.slice(2);
const opt = (name, def) => { const i = args.indexOf(name); if (i < 0) return def; const v = args[i + 1]; args.splice(i, 2); return v; };
const outFile = opt('--out', null);
const maxExamples = +opt('--max-examples', 30);
if (!args.length) {
  console.error('Использование: node tools/scan_books.mjs <папка или файл> [ещё…] [--out отчёт.md] [--max-examples 30]');
  process.exit(2);
}

// вычисляемые не функцией из FN, а самим вычислителем
const KNOWN = new Set([...FUNCTIONS, 'IF', 'IFERROR', 'IFS', 'SWITCH', 'IFNA', 'LET']);
const files = []; const oldXls = []; const skipped = [];
const walk = (p) => {
  let st;
  try { st = statSync(p); } catch { skipped.push([p, 'нет доступа']); return; }
  if (st.isDirectory()) {
    let list = [];
    try { list = readdirSync(p); } catch { skipped.push([p, 'нет доступа к папке']); return; }
    for (const n of list) walk(join(p, n));
    return;
  }
  const ext = extname(p).toLowerCase();
  if (basename(p).startsWith('~$')) return; // временные файлы открытых книг
  if (ext === '.xlsx' || ext === '.xlsm') files.push(p);
  else if (ext === '.xls') oldXls.push(p);
};
for (const a of args) walk(a);

const fnUse = new Map(); // функция → { count, books: Set }
const unknown = new Map(); // имя → { count, books: Set }
const mismatchByFn = new Map(); // функция → число расхождений
const examples = []; // расхождения
const keptAll = new Map(); const warnAll = new Map();
let formulas = 0; let compared = 0; let same = 0; let volatile = 0; let nameErrors = 0; let failed = 0;
const roots = args.map((a) => (statSync(a).isDirectory() ? a : join(a, '..')));
const rel = (p) => { for (const r of roots) { const x = relative(r, p); if (!x.startsWith('..')) return x; } return p; };

// функции формулы (имя → сколько раз)
const fnsOf = (f) => {
  const out = [];
  for (const t of tokenize(f)) if (t.type === 'func') out.push(t.name);
  return out;
};
const sameValue = (a, b) => {
  if (a == null || a === '') a = 0;
  if (b == null || b === '') b = 0;
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.error === b.error;
  if (typeof a === 'number' && typeof b === 'number') return a === b || Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  return a === b;
};
const show = (v) => (v == null ? '(пусто)' : isErr(v) ? errText(v.error) : typeof v === 'boolean' ? (v ? 'ИСТИНА' : 'ЛОЖЬ') : typeof v === 'string' ? `«${v}»` : String(+v.toPrecision(15)).replace('.', ','));

for (const path of files) {
  let res;
  try { res = await importXlsx(readFileSync(path), { recalc: false }); } catch (e) { failed++; skipped.push([path, `не открылась: ${e.message}`]); continue; }
  const { wb } = res;
  for (const k of Object.keys(res.kept || {})) keptAll.set(k, (keptAll.get(k) || 0) + 1);
  for (const k of Object.keys(res.warnings || {})) warnAll.set(k, (warnAll.get(k) || 0) + 1);
  const bookNames = new Set((wb.names || []).map((n) => n.name.toUpperCase()));
  const cached = new Map();
  const note = (f) => {
    for (const fn of fnsOf(f)) {
      if (KNOWN.has(fn)) {
        if (!fnUse.has(fn)) fnUse.set(fn, { count: 0, books: new Set() });
        const u = fnUse.get(fn); u.count++; u.books.add(path);
      } else if (!bookNames.has(fn)) { // имя книги, вызванное как функция, — LAMBDA из диспетчера имён
        if (!unknown.has(fn)) unknown.set(fn, { count: 0, books: new Set() });
        const u = unknown.get(fn); u.count++; u.books.add(path);
      }
    }
  };
  for (const n of wb.names || []) note(n.text);
  for (const sh of wb.sheets) {
    for (const [k, c] of sh.cells) {
      if (c.f == null) continue;
      formulas++;
      note(c.f);
      cached.set(`${sh.name}!${k}`, c.v);
    }
  }
  // полный пересчёт Перо против значений, сохранённых Excel
  let calc;
  try { calc = recalc(wb, { full: true }); } catch (e) { failed++; skipped.push([path, `ошибка пересчёта: ${e.message}`]); continue; }
  for (const sh of calc.sheets) {
    for (const [k, c] of sh.cells) {
      if (c.f == null) continue;
      const fns = fnsOf(c.f);
      if (fns.some((fn) => VOLATILE.has(fn))) { volatile++; continue; } // СЕГОДНЯ, СЛЧИС… — значение другое по определению
      const want = cached.get(`${sh.name}!${k}`);
      if (want === undefined) continue;
      compared++;
      if (isErr(c.v) && c.v.error === '#NAME?' && !(isErr(want) && want.error === '#NAME?')) nameErrors++;
      if (sameValue(c.v, want)) { same++; continue; }
      const top = fns[0] || '(без функций)';
      mismatchByFn.set(top, (mismatchByFn.get(top) || 0) + 1);
      if (examples.length < maxExamples) {
        const [r, cc] = k.split(',').map(Number);
        const colName = (n) => { let s = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
        examples.push({ book: rel(path), cell: `${sh.name}!${colName(cc)}${r + 1}`, f: localize(c.f), excel: show(want), pero: show(c.v) });
      }
    }
  }
}

// ---------- отчёт ----------
const L = [];
const pct = (a, b) => (b ? `${((100 * a) / b).toFixed(1).replace('.', ',')} %` : '—');
L.push('# Сканирование книг', '');
L.push(`Книг .xlsx/.xlsm: ${files.length}${failed ? ` (не прочитано: ${failed})` : ''}; старых .xls (не читаются): ${oldXls.length}.`);
L.push(`Формул: ${formulas}. Сверено с сохранёнными Excel значениями: ${compared}, совпало ${same} (${pct(same, compared)}); летучих (не сверяются): ${volatile}.`);
L.push(`#ИМЯ? у Перо там, где у Excel его нет: ${nameErrors}.`, '');
L.push('## Функции, которых нет в Перо', '');
if (!unknown.size) L.push('Нет — все функции книг поддерживаются.', '');
else {
  L.push('| Функция | Раз | Книг |', '|---|---:|---:|');
  for (const [fn, u] of [...unknown].sort((a, b) => b[1].count - a[1].count)) L.push(`| ${fn} | ${u.count} | ${u.books.size} |`);
  L.push('');
}
L.push('## Частота функций', '', '| Функция | По-русски | Раз | Книг |', '|---|---|---:|---:|');
for (const [fn, u] of [...fnUse].sort((a, b) => b[1].count - a[1].count)) L.push(`| ${fn} | ${ruName(fn)} | ${u.count} | ${u.books.size} |`);
L.push('');
if (mismatchByFn.size) {
  L.push('## Расхождения с Excel по функциям (внешняя функция формулы)', '', '| Функция | Расхождений |', '|---|---:|');
  for (const [fn, n] of [...mismatchByFn].sort((a, b) => b[1] - a[1])) L.push(`| ${ruName(fn)} | ${n} |`);
  L.push('', `### Примеры (первые ${examples.length})`, '', '| Книга | Ячейка | Формула | Excel | Перо |', '|---|---|---|---|---|');
  const esc = (s) => String(s).replace(/\|/g, '\\|');
  for (const e of examples) L.push(`| ${esc(e.book)} | ${esc(e.cell)} | \`=${esc(e.f)}\` | ${esc(e.excel)} | ${esc(e.pero)} |`);
  L.push('');
}
if (warnAll.size || keptAll.size) {
  L.push('## Возможности книг', '');
  if (warnAll.size) L.push(`Не поддерживается (теряется при сохранении): ${[...warnAll].map(([k, n]) => `${k} — ${n}`).join('; ')}.`);
  if (keptAll.size) L.push(`Сохраняется без изменений, но не редактируется: ${[...keptAll].map(([k, n]) => `${k} — ${n}`).join('; ')}.`);
  L.push('');
}
if (oldXls.length || skipped.length) {
  L.push('## Не прочитано', '');
  for (const p of oldXls.slice(0, 50)) L.push(`- ${rel(p)} — формат .xls (пересохранить в .xlsx)`);
  if (oldXls.length > 50) L.push(`- … ещё ${oldXls.length - 50} .xls`);
  for (const [p, why] of skipped) L.push(`- ${rel(p)} — ${why}`);
  L.push('');
}
const text = L.join('\n');
if (outFile) { writeFileSync(outFile, text); console.log(`Отчёт: ${outFile}`); } else console.log(text);
