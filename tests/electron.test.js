// Модули главного процесса без Electron: журнал, последние файлы, автосохранение/восстановление
import { describe, it, expect, beforeEach } from 'vitest';
import { createRequire } from 'module';
import fs from 'fs';
import os from 'os';
import path from 'path';

const require = createRequire(import.meta.url);
const { createLog } = require('../electron/log.cjs');
const { RecentFiles } = require('../electron/recent.cjs');
const { createAutosave } = require('../electron/autosave.cjs');

let dir;
beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pero-test-')); });

describe('журнал', () => {
  it('пишет строки с уровнем, многострочные — с отступом, хвост читается', () => {
    const log = createLog(dir);
    log.info('Запуск');
    log.error(new Error('сломалось'));
    const t = log.tail(1000);
    expect(t[0]).toMatch(/^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d \[INFO\] Запуск$/);
    expect(t[1]).toMatch(/\[ERROR\] Error: сломалось/);
    expect(t.slice(2).every((l) => l.startsWith('    '))).toBe(true); // строки стека
  });
  it('переносит большой журнал в pero.1.log и читает хвост из обоих', () => {
    const log = createLog(dir, { maxBytes: 200 });
    for (let i = 0; i < 20; i++) log.info(`строка ${i}`);
    expect(fs.existsSync(path.join(dir, 'pero.1.log'))).toBe(true);
    const t = log.tail(5);
    expect(t[t.length - 1]).toMatch(/строка 19$/);
    expect(t).toHaveLength(5);
  });
});

describe('последние файлы', () => {
  it('новые сверху, без повторов (регистр пути не важен), не больше max, сохраняются между запусками', () => {
    const file = path.join(dir, 'recent.json');
    const r = new RecentFiles(file, 3);
    r.add('C:\\Док\\a.docx');
    r.add('C:\\Док\\b.xlsx');
    r.add('c:\\док\\A.DOCX');
    expect(r.list().map((x) => x.name)).toEqual(['A.DOCX', 'b.xlsx']);
    r.add('C:\\c.docx'); r.add('C:\\d.docx');
    expect(r.list()).toHaveLength(3);
    const again = new RecentFiles(file, 3);
    expect(again.list().map((x) => x.name)).toEqual(['d.docx', 'c.docx', 'A.DOCX']);
    again.remove('C:\\C.docx');
    expect(again.list().map((x) => x.name)).toEqual(['d.docx', 'A.DOCX']);
    again.clear();
    expect(new RecentFiles(file).list()).toEqual([]);
  });
  it('испорченный файл списка — пустой список, без ошибки', () => {
    const file = path.join(dir, 'recent.json');
    fs.writeFileSync(file, '{не json');
    expect(new RecentFiles(file).list()).toEqual([]);
  });
});

describe('автосохранение', () => {
  it('копии живого и текущего процесса не предлагаются, завершившегося — предлагаются', () => {
    const a = createAutosave(dir);
    a.write('100-1', { name: 'Отчёт.docx', path: 'C:\\Отчёт.docx', kind: 'doc', pid: 100 }, new Uint8Array([1, 2, 3]));
    a.write('200-1', { name: 'Книга1.xlsx', path: null, kind: 'sheet', pid: 200 }, new Uint8Array([4]));
    a.write('300-1', { name: 'Свой.docx', path: null, kind: 'doc', pid: 300 }, new Uint8Array([5]));
    const found = a.scan({ currentPid: 300, isAlive: (pid) => pid === 200 });
    expect(found).toHaveLength(1);
    expect(found[0]).toMatchObject({ id: '100-1', name: 'Отчёт.docx', path: 'C:\\Отчёт.docx', kind: 'doc' });
  });
  it('своя копия окна читается; adopt переносит копию под ключ другого окна; clear удаляет', () => {
    const a = createAutosave(dir);
    a.write('100-1', { name: 'a.docx', kind: 'doc', pid: 100 }, new Uint8Array([7, 8]));
    expect([...a.readOwn('100-1').data]).toEqual([7, 8]);
    const got = a.adopt('100-1', `${process.pid}-5`);
    expect(got.meta.name).toBe('a.docx');
    expect(a.readOwn('100-1')).toBeNull();
    expect([...a.readOwn(`${process.pid}-5`).data]).toEqual([7, 8]);
    a.clear(`${process.pid}-5`);
    expect(fs.readdirSync(dir)).toEqual([]);
  });
  it('мусор (недописанные .tmp, сведения без данных) при просмотре убирается', () => {
    const a = createAutosave(dir);
    fs.writeFileSync(path.join(dir, '1-1.bin.tmp'), 'x');
    fs.writeFileSync(path.join(dir, '2-1.json'), JSON.stringify({ pid: 2, name: 'x' }));
    expect(a.scan({ isAlive: () => false })).toEqual([]);
    expect(fs.readdirSync(dir)).toEqual([]);
  });
});
