// Сканер книг компании (tools/scan_books.mjs): функции, сверка со значениями Excel из файла
import { describe, it, expect } from 'vitest';
import { execFileSync } from 'node:child_process';

describe('сканер книг', () => {
  it('книга из Excel: функции посчитаны, все поддерживаются, значения совпадают', () => {
    const out = execFileSync(process.execPath, ['tools/scan_books.mjs', 'tests/fixtures/book.xlsx'], { encoding: 'utf8' });
    expect(out).toContain('Книг .xlsx/.xlsm: 1');
    expect(out).toContain('Нет — все функции книг поддерживаются.');
    expect(out).toMatch(/\| SUM \| СУММ \| \d+ \| 1 \|/);
    expect(out).toMatch(/#ИМЯ\? у Перо там, где у Excel его нет: 0/);
    expect(out).not.toContain('## Расхождения');
  });
});
