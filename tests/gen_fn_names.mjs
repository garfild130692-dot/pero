// Генерирует src/sheet/fnnames.js — русские имена функций по ответу самого Excel (tests/fixtures/excel_functions.json,
// см. tests/make_function_list.ps1) + функции Microsoft 365, которых нет в Excel 2021 (по документации Microsoft).
// Запуск: node tests/gen_fn_names.mjs
import { readFileSync, writeFileSync } from 'node:fs';
const excel = JSON.parse(readFileSync('tests/fixtures/excel_functions.json', 'utf8'));
// Microsoft 365 (в Excel 2021 нет) — имена по русской справке Microsoft, не сверены с Excel
const M365 = {
  LAMBDA: 'LAMBDA', MAP: 'MAP', REDUCE: 'REDUCE', SCAN: 'SCAN', MAKEARRAY: 'MAKEARRAY', BYROW: 'BYROW', BYCOL: 'BYCOL', ISOMITTED: 'ISOMITTED',
  VSTACK: 'ВСТОЛБИК', HSTACK: 'ГСТОЛБИК', TOCOL: 'ПОСТОЛБЦ', TOROW: 'ПОСТРОК', WRAPROWS: 'СВЕРНСТРОК', WRAPCOLS: 'СВЕРНСТОЛБ',
  TAKE: 'ВЗЯТЬ', DROP: 'СБРОСИТЬ', CHOOSECOLS: 'ВЫБОРСТОЛБЦ', CHOOSEROWS: 'ВЫБОРСТРОК', EXPAND: 'РАЗВЕРНУТЬ',
  TEXTBEFORE: 'ТЕКСТДО', TEXTAFTER: 'ТЕКСТПОСЛЕ', TEXTSPLIT: 'ТЕКСТРАЗД', IMAGE: 'IMAGE',
  REGEXTEST: 'РЕГТЕСТ', REGEXEXTRACT: 'РЕГИЗВЛЕЧЬ', REGEXREPLACE: 'РЕГЗАМЕНИТЬ', GROUPBY: 'ГРУПППО', PIVOTBY: 'СВОДПО',
  PERCENTOF: 'ПРОЦЕНТ', TRIMRANGE: 'УРЕЗДИАПАЗОН',
};
const all = { ...M365 };
for (const [en, ru] of Object.entries(excel)) if (ru && ru !== en.replace(/^_xlfn\./, '')) all[en] = ru;
const lines = Object.keys(all).sort().map((en) => `  '${en}': '${all[en]}',`);
writeFileSync('src/sheet/fnnames.js', `// Русские имена функций Excel (английское → русское). Сгенерировано tests/gen_fn_names.mjs из ответа Excel
// (tests/fixtures/excel_functions.json); функции Microsoft 365 — по справке Microsoft. Не править вручную.
export const EXCEL_RU = {\n${lines.join('\n')}\n};\n`);
console.log('функций с русским именем:', lines.length);
