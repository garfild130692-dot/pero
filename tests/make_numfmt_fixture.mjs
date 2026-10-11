// tests/fixtures/numfmt_excel.json — как Excel показывает значения в числовых форматах (tests/numfmt.test.js).
// 1) книга с кодами из tests/numfmt_cases.json (коды в файле — английская запись, как их хранит Excel);
// 2) Excel (COM) открывает её и возвращает текст каждой ячейки. Запуск: node tests/make_numfmt_fixture.mjs
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import * as m from '../src/sheet/model.js';
import { exportXlsx } from '../src/sheet/xlsx.js';
const { codes, values } = JSON.parse(readFileSync('tests/numfmt_cases.json', 'utf8'));
let wb = m.newWorkbook();
codes.forEach((code, r) => values.forEach((v, c) => {
  wb = m.setInput(wb, 0, r, c, typeof v === 'number' ? String(v).replace('.', ',') : v);
  wb = m.setStyle(wb, 0, { r1: r, c1: c, r2: r, c2: c }, { numFmt: code });
}));
mkdirSync('tests/out', { recursive: true });
writeFileSync('tests/out/numfmt.xlsx', await exportXlsx(wb));
const ps = `$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try { $wb = $xl.Workbooks.Open([IO.Path]::GetFullPath('tests/out/numfmt.xlsx')); $ws = $wb.Worksheets.Item(1); $ws.Range('A:Z').ColumnWidth = 80
  $rows = @(); for ($r = 1; $r -le ${codes.length}; $r++) { $t = @(); for ($c = 1; $c -le ${values.length}; $c++) { $t += $ws.Cells.Item($r, $c).Text }; $rows += ,$t }
  $rows | ConvertTo-Json -Depth 3 -Compress | Out-File -Encoding utf8 tests/out/numfmt_texts.json; $wb.Close($false) } finally { $xl.Quit() }`;
execFileSync('powershell.exe', ['-NoProfile', '-Command', ps], { stdio: 'inherit' });
const texts = JSON.parse(readFileSync('tests/out/numfmt_texts.json', 'utf8').replace(/^﻿/, ''));
writeFileSync('tests/fixtures/numfmt_excel.json', JSON.stringify(codes.map((code, i) => ({ code, texts: texts[i].value || texts[i] })), null, 1));
console.log('tests/fixtures/numfmt_excel.json:', codes.length, 'кодов');
