// CSV: открытие (кодировка UTF-8 / UTF-16 / Windows-1251, разделитель «;», «,» или табуляция — определяются сами)
// и сохранение текущего листа (как «CSV (разделители — точка с запятой)» русского Excel: «;», значения — как на экране).
import { key, unkey } from './address.js';
import { fromUser, isErr } from './formula.js';
import { formatValue, parseInput } from './format.js';
import { newSheet, recalc } from './model.js';

// Байты → текст: BOM, иначе UTF-8, если он корректен, иначе Windows-1251 (CSV из 1С и старого Excel)
export function decodeText(data) {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(b.subarray(3)), encoding: 'utf-8' };
  if (b[0] === 0xff && b[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(b.subarray(2)), encoding: 'utf-16le' };
  if (b[0] === 0xfe && b[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(b.subarray(2)), encoding: 'utf-16be' };
  try {
    return { text: new TextDecoder('utf-8', { fatal: true }).decode(b), encoding: 'utf-8' };
  } catch {
    return { text: new TextDecoder('windows-1251').decode(b), encoding: 'windows-1251' };
  }
}

// Разделитель: тот из «;», «,», табуляции, что встречается в первых строках одинаковое (и ненулевое) число раз.
// При равенстве — «;» (русский Excel), потом табуляция, потом «,».
export function detectDelimiter(text) {
  const lines = [];
  let cur = ''; let q = false;
  for (let i = 0; i < text.length && lines.length < 30; i++) {
    const ch = text[i];
    if (ch === '"') q = !q;
    if (!q && (ch === '\n' || ch === '\r')) { if (cur) lines.push(cur); cur = ''; continue; }
    cur += q && ch !== '"' ? '' : ch; // содержимое кавычек не считаем
  }
  if (cur) lines.push(cur);
  let best = ';'; let bestScore = -1;
  for (const d of [';', '\t', ',']) {
    const counts = lines.map((l) => l.split(d).length - 1);
    if (!counts.length || !counts[0]) continue;
    const same = counts.filter((n) => n === counts[0]).length;
    const score = same * 1000 + counts[0];
    if (score > bestScore) { best = d; bestScore = score; }
  }
  return best;
}

// Разбор CSV (RFC 4180): кавычки, удвоенные кавычки, переводы строк внутри кавычек
export function parseCsv(text, delimiter = ';') {
  const rows = [];
  let row = []; let field = ''; let q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += ch;
    } else if (ch === '"' && field === '') q = true;
    else if (ch === delimiter) { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

// Ошибки в CSV русского Excel пишутся по-русски
const ERRORS = {
  '#ДЕЛ/0!': '#DIV/0!', '#Н/Д': '#N/A', '#ЗНАЧ!': '#VALUE!', '#ССЫЛКА!': '#REF!', '#ИМЯ?': '#NAME?', '#ЧИСЛО!': '#NUM!', '#ПУСТО!': '#NULL!',
  '#DIV/0!': '#DIV/0!', '#N/A': '#N/A', '#VALUE!': '#VALUE!', '#REF!': '#REF!', '#NAME?': '#NAME?', '#NUM!': '#NUM!', '#NULL!': '#NULL!',
};

// при сохранении — как русский Excel
const RU_ERROR = Object.fromEntries(Object.entries(ERRORS).filter(([k]) => /[А-Я]/.test(k)).map(([k, v]) => [v, k]));

// Книга из CSV: один лист с именем файла; числа, даты, проценты — как при вводе с клавиатуры, «=…» — формула
export function importCsv(data, fileName = 'Лист1.csv') {
  const { text, encoding } = decodeText(data);
  const delimiter = detectDelimiter(text);
  const name = String(fileName).replace(/\.[^.]*$/, '').replace(/[\\/?*[\]:]/g, '').slice(0, 31) || 'Лист1';
  const sheet = newSheet(name);
  parseCsv(text, delimiter).forEach((row, r) => row.forEach((raw, c) => {
    if (raw === '') return;
    if (raw.startsWith('=') && raw.length > 1) {
      try { sheet.cells.set(key(r, c), { v: null, f: fromUser(raw.slice(1)) }); return; } catch { /* не формула — как текст */ }
    }
    if (ERRORS[raw.trim().toUpperCase()]) { sheet.cells.set(key(r, c), { v: { error: ERRORS[raw.trim().toUpperCase()] } }); return; }
    const p = parseInput(raw);
    if (p.value == null) return;
    const cell = { v: p.value };
    if (p.format) cell.s = { numFmt: p.format };
    sheet.cells.set(key(r, c), cell);
  }));
  return { wb: recalc({ sheets: [sheet], active: 0 }), encoding, delimiter };
}

// CSV текущего листа: значения как на экране (с форматом), «;», UTF-8 с BOM (Excel открывает без вопросов о кодировке)
export function exportCsv(sheet, { delimiter = ';', bom = true } = {}) {
  let r2 = -1; let c2 = -1;
  for (const k of sheet.cells.keys()) { const [r, c] = unkey(k); if (r > r2) r2 = r; if (c > c2) c2 = c; }
  const quote = (s) => (s.includes(delimiter) || /["\r\n]/.test(s) || s !== s.trim() ? `"${s.replace(/"/g, '""')}"` : s);
  const lines = [];
  for (let r = 0; r <= r2; r++) {
    const out = [];
    for (let c = 0; c <= c2; c++) {
      const cell = sheet.cells.get(key(r, c));
      const v = cell ? cell.v : null;
      const text = v == null ? '' : isErr(v) ? RU_ERROR[v.error] || v.error : typeof v === 'boolean' ? (v ? 'ИСТИНА' : 'ЛОЖЬ') : formatValue(v, cell.s && cell.s.numFmt);
      out.push(quote(String(text).replace(/ /g, ' ')));
    }
    lines.push(out.join(delimiter));
  }
  const body = new TextEncoder().encode(`${lines.join('\r\n')}${lines.length ? '\r\n' : ''}`);
  if (!bom) return body;
  const outBytes = new Uint8Array(body.length + 3);
  outBytes.set([0xef, 0xbb, 0xbf]);
  outBytes.set(body, 3);
  return outBytes;
}
