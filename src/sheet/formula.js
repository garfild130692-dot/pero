// Формулы: токенизатор, разбор (приоритеты Excel), вычисление, функции, переписывание ссылок.
//
// Формула хранится без «=» в синтаксисе файла xlsx: английские имена функций, разделитель аргументов «,»,
// десятичная точка. При вводе принимаются и русские имена (СУММ, ЕСЛИ…) и «;» — они приводятся к хранимому виду.
import { colIndex, colName } from './address.js';

// ---------- Ошибки и значения ----------
export const ERR = {
  DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', NA: '#N/A', NUM: '#NUM!', CIRC: '#CIRC!',
};
export const isErr = (v) => v != null && typeof v === 'object' && 'error' in v;
export const err = (code) => ({ error: code });

// ---------- Русские имена функций ----------
const RU = {
  СУММ: 'SUM', СРЗНАЧ: 'AVERAGE', МИН: 'MIN', МАКС: 'MAX', СЧЁТ: 'COUNT', СЧЕТ: 'COUNT', СЧЁТЗ: 'COUNTA', СЧЕТЗ: 'COUNTA',
  СЧИТАТЬПУСТОТЫ: 'COUNTBLANK', ЕСЛИ: 'IF', И: 'AND', ИЛИ: 'OR', НЕ: 'NOT', ОКРУГЛ: 'ROUND', ОКРУГЛВВЕРХ: 'ROUNDUP',
  ОКРУГЛВНИЗ: 'ROUNDDOWN', ЦЕЛОЕ: 'INT', ОСТАТ: 'MOD', СТЕПЕНЬ: 'POWER', КОРЕНЬ: 'SQRT', СЦЕПИТЬ: 'CONCATENATE',
  СЦЕП: 'CONCAT', ДЛСТР: 'LEN', ПРОПИСН: 'UPPER', СТРОЧН: 'LOWER', СЖПРОБЕЛЫ: 'TRIM', ЛЕВСИМВ: 'LEFT', ПРАВСИМВ: 'RIGHT',
  ПСТР: 'MID', СЕГОДНЯ: 'TODAY', ТДАТА: 'NOW', СУММЕСЛИ: 'SUMIF', СЧЁТЕСЛИ: 'COUNTIF', СЧЕТЕСЛИ: 'COUNTIF',
  СРЗНАЧЕСЛИ: 'AVERAGEIF', ЕСЛИОШИБКА: 'IFERROR', ВПР: 'VLOOKUP', ПРОИЗВЕД: 'PRODUCT', ГОД: 'YEAR', МЕСЯЦ: 'MONTH',
  ДЕНЬ: 'DAY', ДАТА: 'DATE', ЕПУСТО: 'ISBLANK', ЕЧИСЛО: 'ISNUMBER', ЕТЕКСТ: 'ISTEXT', ЕОШИБКА: 'ISERROR', ПИ: 'PI',
  ЗНАК: 'SIGN', МЕДИАНА: 'MEDIAN', НАЙТИ: 'FIND', ПОДСТАВИТЬ: 'SUBSTITUTE', ПОВТОР: 'REPT', ЗНАЧЕН: 'VALUE',
  ИСТИНА: 'TRUE', ЛОЖЬ: 'FALSE', ИНДЕКС: 'INDEX', ПОИСКПОЗ: 'MATCH', СУММПРОИЗВ: 'SUMPRODUCT', РАНГ: 'RANK',
};

// ---------- Токенизатор ----------
// Токены: { type, text, ... }; text — исходный текст (для переписывания формулы без потери пробелов)
const SHEET_RE = /^(?:'((?:[^']|'')+)'|([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*))!/;
export function tokenize(src) {
  const out = [];
  let i = 0;
  const s = String(src);
  while (i < s.length) {
    const rest = s.slice(i);
    let m;
    if ((m = /^\s+/.exec(rest))) { out.push({ type: 'ws', text: m[0] }); i += m[0].length; continue; }
    if ((m = /^"((?:[^"]|"")*)"/.exec(rest))) { out.push({ type: 'str', text: m[0], value: m[1].replace(/""/g, '"') }); i += m[0].length; continue; }
    if ((m = /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!|CIRC!)/.exec(rest))) { out.push({ type: 'err', text: m[0] }); i += m[0].length; continue; }
    // ссылка (возможно с листом): Лист1!A1, 'Мой лист'!$B$2
    let sheet = null;
    let sheetText = '';
    if ((m = SHEET_RE.exec(rest))) {
      const after = rest.slice(m[0].length);
      if (/^\$?[A-Za-z]{1,3}\$?\d/.test(after)) { sheet = m[1] ? m[1].replace(/''/g, "'") : m[2]; sheetText = m[0]; }
    }
    const refPart = rest.slice(sheetText.length);
    if ((m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/.exec(refPart))) {
      out.push({ type: 'ref', text: sheetText + m[0], sheet, absC: !!m[1], c: colIndex(m[2]), absR: !!m[3], r: +m[4] - 1 });
      i += sheetText.length + m[0].length;
      continue;
    }
    if ((m = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest)) || (m = /^\.\d+(?:[eE][+-]?\d+)?/.exec(rest))) {
      out.push({ type: 'num', text: m[0], value: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    if ((m = /^([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)\s*\(/.exec(rest))) {
      const name = m[1].toUpperCase();
      out.push({ type: 'func', text: m[1], name: RU[name] || name });
      i += m[1].length;
      continue;
    }
    if ((m = /^(TRUE|FALSE|ИСТИНА|ЛОЖЬ)(?![\wА-Яа-яЁё])/i.exec(rest))) {
      const v = /^(TRUE|ИСТИНА)$/i.test(m[1]);
      out.push({ type: 'bool', text: m[1], value: v }); i += m[1].length; continue;
    }
    if ((m = /^(<=|>=|<>)/.exec(rest))) { out.push({ type: 'op', text: m[0] }); i += 2; continue; }
    const ch = s[i];
    if ('+-*/^&=<>%'.includes(ch)) { out.push({ type: 'op', text: ch }); i++; continue; }
    if (ch === '(' || ch === ')') { out.push({ type: ch, text: ch }); i++; continue; }
    if (ch === ',' || ch === ';') { out.push({ type: 'sep', text: ch }); i++; continue; }
    if (ch === ':') { out.push({ type: ':', text: ch }); i++; continue; }
    if ((m = /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*/.exec(rest))) { out.push({ type: 'name', text: m[0] }); i += m[0].length; continue; }
    out.push({ type: 'bad', text: ch });
    i++;
  }
  return out;
}

// ---------- Разбор ----------
export function parse(src) {
  const toks = tokenize(src).filter((t) => t.type !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type) => { const t = next(); if (!t || t.type !== type) throw new Error('syntax'); return t; };
  const isOp = (t, ops) => t && t.type === 'op' && ops.includes(t.text);

  function primary() {
    const t = next();
    if (!t) throw new Error('syntax');
    switch (t.type) {
      case 'num': return { t: 'num', v: t.value };
      case 'str': return { t: 'str', v: t.value };
      case 'bool': return { t: 'bool', v: t.value };
      case 'err': return { t: 'err', v: t.text };
      case 'ref': {
        if (peek() && peek().type === ':' && toks[p + 1] && toks[p + 1].type === 'ref') {
          next();
          const b = next();
          return { t: 'range', sheet: t.sheet, r1: Math.min(t.r, b.r), c1: Math.min(t.c, b.c), r2: Math.max(t.r, b.r), c2: Math.max(t.c, b.c) };
        }
        return { t: 'ref', sheet: t.sheet, r: t.r, c: t.c };
      }
      case 'func': {
        expect('(');
        const args = [];
        if (peek() && peek().type === ')') { next(); return { t: 'fn', name: t.name, args }; }
        for (;;) {
          if (peek() && (peek().type === 'sep' || peek().type === ')')) args.push({ t: 'empty' });
          else args.push(expr());
          const n = next();
          if (!n) throw new Error('syntax');
          if (n.type === ')') break;
          if (n.type !== 'sep') throw new Error('syntax');
        }
        return { t: 'fn', name: t.name, args };
      }
      case '(': { const e = expr(); expect(')'); return e; }
      case 'op':
        if (t.text === '-' || t.text === '+') return { t: 'un', op: t.text, a: unary() };
        throw new Error('syntax');
      case 'name': return { t: 'name', v: t.text };
      default: throw new Error('syntax');
    }
  }
  function postfix() {
    let a = primary();
    while (isOp(peek(), ['%'])) { next(); a = { t: 'pct', a }; }
    return a;
  }
  function unary() {
    if (isOp(peek(), ['-', '+'])) { const op = next().text; return { t: 'un', op, a: unary() }; }
    return postfix();
  }
  function power() {
    let a = unary();
    while (isOp(peek(), ['^'])) { next(); a = { t: 'bin', op: '^', a, b: unary() }; }
    return a;
  }
  function mul() {
    let a = power();
    while (isOp(peek(), ['*', '/'])) { const op = next().text; a = { t: 'bin', op, a, b: power() }; }
    return a;
  }
  function add() {
    let a = mul();
    while (isOp(peek(), ['+', '-'])) { const op = next().text; a = { t: 'bin', op, a, b: mul() }; }
    return a;
  }
  function concat() {
    let a = add();
    while (isOp(peek(), ['&'])) { next(); a = { t: 'bin', op: '&', a, b: add() }; }
    return a;
  }
  function expr() {
    let a = concat();
    while (isOp(peek(), ['=', '<>', '<', '>', '<=', '>='])) { const op = next().text; a = { t: 'bin', op, a, b: concat() }; }
    return a;
  }
  const ast = expr();
  if (p < toks.length) throw new Error('syntax');
  return ast;
}

// ---------- Приведение типов ----------
const parseNumStr = (s) => {
  const t = String(s).trim().replace(/\s/g, '').replace(',', '.');
  return t !== '' && isFinite(+t) ? +t : null;
};
// Текст → число, как русский Excel в арифметике и ЗНАЧЕН: «1 234,5», «12%», «100 ₽», «15.01.2026», «14:30»
export function textToNum(text) {
  const t = String(text).trim();
  let n = parseNumStr(t);
  if (n !== null && t !== '') return n;
  let m;
  const flat = t.replace(/[\s\u00a0]/g, '');
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)%$/.exec(flat))) return parseFloat(m[1].replace(',', '.')) / 100;
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)(₽|р\.?|руб\.?)$/i.exec(flat))) return parseFloat(m[1].replace(',', '.'));
  if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    const y = m[3].length === 2 ? (+m[3] < 30 ? 2000 : 1900) + +m[3] : +m[3];
    if (+m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1 && +m[1] <= 31) {
      n = dateToSerial(y, +m[2], +m[1]);
      return m[4] ? n + (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400 : n;
    }
  }
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) return (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
  return null;
}
// Пустая ячейка — 0; пустая строка (например, результат ="") — #ЗНАЧ!, как в Excel
export function toNum(v) {
  if (isErr(v)) return v;
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = textToNum(v);
  return n === null ? err(ERR.VALUE) : n;
}
// Число как текст в формате «Общий» русского Excel: до 15 значащих цифр, десятичная запятая
export function numText(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let t = String(+n.toPrecision(15));
  if (/e/.test(t)) t = t.replace(/e([+-])(\d)$/, 'E$10$2').replace('e', 'E');
  return t.replace('.', ',');
}
export function toText(v) {
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'number') return numText(v);
  return String(v);
}
function toBool(v) {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  if (v == null) return false;
  if (v === '') return err(ERR.VALUE);
  if (typeof v === 'number') return v !== 0;
  const s = String(v).toUpperCase();
  if (s === 'TRUE' || s === 'ИСТИНА') return true;
  if (s === 'FALSE' || s === 'ЛОЖЬ') return false;
  return err(ERR.VALUE);
}
const flat = (v) => (Array.isArray(v) ? v.flat() : [v]);

// Сравнение как в Excel: числа < текст < логические; текст — без учёта регистра
function compare(a, b) {
  const rank = (x) => (typeof x === 'number' || x == null ? 0 : typeof x === 'string' ? 1 : 2);
  const ra = rank(a);
  const rb = rank(b);
  if (a == null && typeof b === 'string') return b === '' ? 0 : -1;
  if (b == null && typeof a === 'string') return a === '' ? 0 : 1;
  if (ra !== rb) return ra - rb;
  if (ra === 1) { const x = a.toLowerCase(); const y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  const x = ra === 2 ? +a : a || 0;
  const y = rb === 2 ? +b : b || 0;
  if (x === y || (ra === 0 && eq15(x, y))) return 0;
  return x < y ? -1 : 1;
}
// Excel сравнивает числа с точностью 15 значащих цифр (0,1+0,2=0,3 — ИСТИНА)
const eq15 = (x, y) => x === y || (Number.isFinite(x) && Number.isFinite(y) && x.toPrecision(15) === y.toPrecision(15));

// Подстановочные знаки Excel (* ? и ~ для экранирования) → RegExp; null — их нет
function wildcard(text) {
  if (!/[*?]/.test(text)) return null;
  let re = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '~' && i + 1 < text.length) { re += text[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    re += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}
// Точное совпадение для ВПР/ПОИСКПОЗ/ГПР: текст — без регистра и с подстановочными знаками
function exactMatcher(val) {
  const re = typeof val === 'string' ? wildcard(val) : null;
  return re ? (v) => typeof v === 'string' && re.test(v) : (v) => compare(v, val) === 0;
}

// Условие СУММЕСЛИ/СЧЁТЕСЛИ: ">5", "<>x", "=текст", "те*ст", число
function criterion(c) {
  if (typeof c === 'boolean') return (v) => v === c;
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? eq15(v, c) : typeof v === 'string' && parseNumStr(v) !== null && eq15(parseNumStr(v), c));
  const s = typeof c === 'string' ? c : toText(c);
  if (s === '<>') return (v) => v != null;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s);
  const op = m[1] || '=';
  const raw = m[2];
  const n = parseNumStr(raw);
  const target = n !== null ? n : raw;
  const re = n === null && (op === '=' || op === '<>') ? wildcard(raw) : null;
  if (re) return (v) => (op === '=') === (typeof v === 'string' && re.test(v));
  return (v) => {
    if (n !== null && typeof v !== 'number') { const vn = typeof v === 'string' ? parseNumStr(v) : null; if (vn === null) return op === '<>'; v = vn; }
    const r = compare(v == null ? (n !== null ? 0 : '') : v, target);
    return { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[op];
  };
}

// ---------- Даты (серийные номера, система 1900) ----------
const EPOCH = Date.UTC(1899, 11, 30);
// Как Excel: годы 0–1899 — плюс 1900; 1900 считается високосным (серийный номер 60 — 29.02.1900, «ошибка 1900 года»)
export const dateToSerial = (y, m, d) => {
  const yy = y < 1900 ? y + 1900 : y;
  let first = (Date.UTC(yy, m - 1, 1) - EPOCH) / 86400000; // первое число месяца
  if (first < 61) first -= 1; // до 1 марта 1900 Excel отстаёт на день
  return first + d - 1;
};
export function serialToDate(n) {
  if (Math.floor(n) === 60) return { y: 1900, m: 2, d: 29, hh: 0, mm: 0, ss: 0 };
  const d = new Date(EPOCH + Math.round((n < 60 ? n + 1 : n) * 86400000));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds() };
}
const nowSerial = () => { const t = new Date(); return (Date.UTC(t.getFullYear(), t.getMonth(), t.getDate(), t.getHours(), t.getMinutes(), t.getSeconds()) - EPOCH) / 86400000; };

// ---------- Функции ----------
const round = (x, n, mode) => {
  const k = 10 ** n;
  const v = x * k;
  const r = mode === 'up' ? Math.sign(v) * Math.ceil(Math.abs(v) - 1e-9) : mode === 'down' ? Math.sign(v) * Math.floor(Math.abs(v) + 1e-9) : Math.sign(v) * Math.round(Math.abs(v) + 1e-9);
  return r / k;
};
const nums = (args, { strict = false } = {}) => {
  const out = [];
  for (const a of args) {
    if (Array.isArray(a)) { for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); } }
    else { if (isErr(a)) return a; if (a == null || a === '') { if (strict) out.push(0); continue; } const n = toNum(a); if (isErr(n)) return n; out.push(n); }
  }
  return out;
};
const firstErr = (...xs) => xs.find(isErr);
function logicals(args) {
  const out = [];
  for (const a of args) {
    for (const v of flat(a)) {
      if (v == null || v === '') continue;
      if (isErr(v) || typeof v === 'boolean') { out.push(v); continue; }
      if (typeof v === 'number') { out.push(v !== 0); continue; }
      const b = toBool(v);
      if (!isErr(b)) out.push(b);
    }
  }
  return out;
}

// аргументы приходят уже вычисленными (диапазон — двумерный массив), кроме «ленивых» IF/IFERROR (см. evaluate)
const FN = {
  SUM: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s + x, 0); },
  PRODUCT: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s * x, 1); },
  AVERAGE: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? n.reduce((s, x) => s + x, 0) / n.length : err(ERR.DIV0); },
  MIN: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; },
  MEDIAN: (a) => {
    const n = nums(a);
    if (isErr(n)) return n;
    if (!n.length) return err(ERR.NUM);
    const s = [...n].sort((x, y) => x - y);
    const h = s.length >> 1;
    return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
  },
  COUNT: (a) => a.reduce((s, x) => s + flat(x).filter((v) => typeof v === 'number').length, 0),
  COUNTA: (a) => a.reduce((s, x) => s + flat(x).filter((v) => v != null).length, 0),
  COUNTBLANK: (a) => a.reduce((s, x) => s + flat(x).filter((v) => v == null || v === '').length, 0),
  ABS: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.abs(n); },
  SIGN: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.sign(n); },
  INT: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.floor(n); },
  SQRT: ([x]) => { const n = toNum(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.sqrt(n); },
  PI: () => Math.PI,
  POWER: ([x, y]) => { const a = toNum(x); const b = toNum(y); return firstErr(a, b) || a ** b; },
  MOD: ([x, y]) => {
    const a = toNum(x); const b = toNum(y);
    if (firstErr(a, b)) return firstErr(a, b);
    if (b === 0) return err(ERR.DIV0);
    return a - b * Math.floor(a / b);
  },
  ROUND: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d)); },
  ROUNDUP: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d), 'up'); },
  ROUNDDOWN: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d), 'down'); },
  AND: (a) => { let r = true; for (const v of logicals(a)) { if (isErr(v)) return v; r = r && v; } return r; },
  OR: (a) => { let r = false; for (const v of logicals(a)) { if (isErr(v)) return v; r = r || v; } return r; },
  NOT: ([x]) => { const b = toBool(x); return isErr(b) ? b : !b; },
  TRUE: () => true,
  FALSE: () => false,
  CONCATENATE: (a) => { const e = a.flatMap(flat).find(isErr); return e || a.flatMap(flat).map(toText).join(''); },
  CONCAT: (a) => FN.CONCATENATE(a),
  LEN: ([x]) => (isErr(x) ? x : toText(x).length),
  UPPER: ([x]) => (isErr(x) ? x : toText(x).toUpperCase()),
  LOWER: ([x]) => (isErr(x) ? x : toText(x).toLowerCase()),
  TRIM: ([x]) => (isErr(x) ? x : toText(x).trim().replace(/\s+/g, ' ')),
  LEFT: ([x, n]) => { const k = toNum(n ?? 1); return firstErr(x, k) || toText(x).slice(0, Math.max(0, k)); },
  RIGHT: ([x, n]) => { const k = toNum(n ?? 1); const s = toText(x); return firstErr(x, k) || (k > 0 ? s.slice(-k) : ''); },
  MID: ([x, st, n]) => { const a = toNum(st); const k = toNum(n); return firstErr(x, a, k) || toText(x).substr(Math.max(0, a - 1), Math.max(0, k)); },
  FIND: ([what, where, st]) => {
    const s = toNum(st ?? 1);
    if (firstErr(what, where, s)) return firstErr(what, where, s);
    const i = toText(where).indexOf(toText(what), s - 1);
    return i < 0 ? err(ERR.VALUE) : i + 1;
  },
  SUBSTITUTE: ([t, o, n, k]) => {
    const e = firstErr(t, o, n, k);
    if (e) return e;
    const text = toText(t); const old = toText(o);
    if (!old) return text;
    if (k == null) return text.split(old).join(toText(n));
    const which = toNum(k);
    if (isErr(which) || which < 1) return err(ERR.VALUE);
    let at = -1;
    for (let i = 0; i < which; i++) { at = text.indexOf(old, at + 1); if (at < 0) return text; }
    return text.slice(0, at) + toText(n) + text.slice(at + old.length);
  },
  REPT: ([t, n]) => { const k = toNum(n); return firstErr(t, k) || toText(t).repeat(Math.max(0, Math.floor(k))); },
  VALUE: ([x]) => toNum(x),
  ISBLANK: ([x]) => x == null,
  ISNUMBER: ([x]) => typeof x === 'number',
  ISTEXT: ([x]) => typeof x === 'string',
  ISERROR: ([x]) => isErr(x),
  TODAY: () => Math.floor(nowSerial()),
  NOW: () => nowSerial(),
  DATE: ([y, m, d]) => {
    const a = toNum(y); const b = toNum(m); const c = toNum(d);
    if (firstErr(a, b, c)) return firstErr(a, b, c);
    const yy = Math.trunc(a);
    if (yy < 0 || yy > 9999) return err(ERR.NUM);
    const v = dateToSerial(yy, Math.trunc(b), Math.trunc(c));
    return v < 0 ? err(ERR.NUM) : v;
  },
  YEAR: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).y; },
  MONTH: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).m; },
  DAY: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).d; },
  SUMIF: ([range, crit, sumRange]) => {
    const test = criterion(crit);
    const r = Array.isArray(range) ? range : [[range]];
    const s = Array.isArray(sumRange) ? sumRange : r;
    let total = 0;
    r.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') total += x; } }));
    return total;
  },
  COUNTIF: ([range, crit]) => { const test = criterion(crit); return flat(range).filter((v) => test(v)).length; },
  AVERAGEIF: ([range, crit, avgRange]) => {
    const test = criterion(crit);
    const r = Array.isArray(range) ? range : [[range]];
    const s = Array.isArray(avgRange) ? avgRange : r;
    const pick = [];
    r.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') pick.push(x); } }));
    return pick.length ? pick.reduce((a, b) => a + b, 0) / pick.length : err(ERR.DIV0);
  },
  SUMPRODUCT: (a) => {
    const arrs = a.map((x) => flat(x));
    const n = arrs[0] ? arrs[0].length : 0;
    if (arrs.some((x) => x.length !== n)) return err(ERR.VALUE);
    let s = 0;
    for (let i = 0; i < n; i++) s += arrs.reduce((p, x) => p * (typeof x[i] === 'number' ? x[i] : 0), 1);
    return s;
  },
  VLOOKUP: ([val, table, col, approx]) => {
    if (isErr(val)) return val;
    if (!Array.isArray(table)) return err(ERR.NA);
    const c = toNum(col);
    if (isErr(c) || c < 1 || c > table[0].length) return err(ERR.REF);
    const exact = approx === false || approx === 0;
    if (exact) {
      const hit = exactMatcher(val);
      const row = table.find((r) => hit(r[0]));
      return row ? (row[c - 1] ?? 0) : err(ERR.NA);
    }
    let found = null;
    for (const r of table) { if (r[0] == null) continue; if (compare(r[0], val) <= 0) found = r; else break; }
    return found ? (found[c - 1] ?? 0) : err(ERR.NA);
  },
  MATCH: ([val, range, type]) => {
    const list = flat(range);
    const t = type == null ? 1 : toNum(type);
    if (isErr(val)) return val;
    if (t === 0) { const hit = exactMatcher(val); const i = list.findIndex(hit); return i < 0 ? err(ERR.NA) : i + 1; }
    let found = -1;
    list.forEach((v, i) => { if (v != null && (t > 0 ? compare(v, val) <= 0 : compare(v, val) >= 0)) found = i; });
    return found < 0 ? err(ERR.NA) : found + 1;
  },
  INDEX: ([range, r, c]) => {
    if (isErr(range)) return range;
    const t = Array.isArray(range) ? range : [[range]];
    const i = toNum(r ?? 1); const j = toNum(c ?? 1);
    if (firstErr(i, j)) return firstErr(i, j);
    const row = t[(i || 1) - 1];
    return row && (j || 1) <= row.length ? (row[(j || 1) - 1] ?? 0) : err(ERR.REF);
  },
  RANK: ([val, range, order]) => {
    const n = toNum(val);
    if (isErr(n)) return n;
    const list = flat(range).filter((x) => typeof x === 'number');
    const desc = !order;
    return 1 + list.filter((x) => (desc ? x > n : x < n)).length;
  },
};
export const FUNCTIONS = Object.keys(FN);

// ---------- Вычисление ----------
// ctx: { get(sheet, r, c) → значение ячейки (с вычислением формул), sheet — имя текущего листа }
export function evaluate(node, ctx) {
  switch (node.t) {
    case 'num': case 'str': case 'bool': return node.v;
    case 'err': return err(node.v);
    case 'empty': return null;
    case 'name': return err(ERR.NAME);
    case 'ref': return ctx.get(node.sheet || ctx.sheet, node.r, node.c);
    case 'range': {
      if ((node.r2 - node.r1 + 1) * (node.c2 - node.c1 + 1) > 2_000_000) return err(ERR.NUM);
      const out = [];
      for (let r = node.r1; r <= node.r2; r++) {
        const row = [];
        for (let c = node.c1; c <= node.c2; c++) row.push(ctx.get(node.sheet || ctx.sheet, r, c));
        out.push(row);
      }
      return out;
    }
    case 'pct': { const n = toNum(scalar(evaluate(node.a, ctx))); return isErr(n) ? n : n / 100; }
    case 'un': { const n = toNum(scalar(evaluate(node.a, ctx))); return isErr(n) ? n : node.op === '-' ? -n : n; }
    case 'bin': {
      const a = scalar(evaluate(node.a, ctx));
      const b = scalar(evaluate(node.b, ctx));
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      if (node.op === '&') return toText(a) + toText(b);
      if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
        const r = compare(a, b);
        return { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[node.op];
      }
      const x = toNum(a);
      const y = toNum(b);
      if (isErr(x)) return x;
      if (isErr(y)) return y;
      switch (node.op) {
        case '+': return x + y;
        case '-': return x - y;
        case '*': return x * y;
        case '/': return y === 0 ? err(ERR.DIV0) : x / y;
        case '^': { const v = x ** y; return isFinite(v) ? v : err(ERR.NUM); }
        default: return err(ERR.VALUE);
      }
    }
    case 'fn': {
      // «ленивые» функции: вычисляют только нужную ветку
      if (node.name === 'IF') {
        const cond = toBool(scalar(evaluate(node.args[0] || { t: 'empty' }, ctx)));
        if (isErr(cond)) return cond;
        const branch = cond ? node.args[1] : node.args[2];
        if (!branch) return cond ? true : false;
        return scalar(evaluate(branch, ctx)) ?? 0;
      }
      if (node.name === 'IFERROR') {
        const v = scalar(evaluate(node.args[0], ctx));
        return isErr(v) ? scalar(evaluate(node.args[1] || { t: 'str', v: '' }, ctx)) : v;
      }
      const f = FN[node.name];
      if (!f) return err(ERR.NAME);
      const args = node.args.map((a) => evaluate(a, ctx));
      const v = f(args);
      return typeof v === 'number' && !isFinite(v) ? err(ERR.NUM) : v;
    }
    default: return err(ERR.VALUE);
  }
}
// значение диапазона в скалярном контексте — левая верхняя ячейка (упрощение «неявного пересечения»)
const scalar = (v) => (Array.isArray(v) ? (v[0] ? v[0][0] : null) : v);

// ---------- Переписывание формул ----------
// Привести к хранимому виду: английские имена функций, «,» между аргументами
export function canonical(src) {
  return tokenize(src).map((t) => (t.type === 'func' ? t.name : t.type === 'sep' ? ',' : t.type === 'bool' ? (t.value ? 'TRUE' : 'FALSE') : t.text)).join('');
}

// Формула, введённая пользователем (русская запись): «1,5» между цифрами — десятичная запятая,
// разделитель аргументов — «;» (запятая между не-числами тоже принимается). Результат — каноническая запись.
export function fromUser(src) {
  const t = tokenize(src);
  let out = '';
  for (let i = 0; i < t.length; i++) {
    if (t[i].type === 'num' && /^\d+$/.test(t[i].text) && t[i + 1] && t[i + 1].text === ',' && t[i + 2] && t[i + 2].type === 'num' && /^\d/.test(t[i + 2].text)) {
      out += `${t[i].text}.${t[i + 2].text}`;
      i += 2;
    } else out += t[i].text;
  }
  return canonical(out);
}

// Каноническая формула → русская запись для строки формул: ЕСЛИ(A1>1,5;"да";"нет")
const EN_RU = {};
for (const [ru, en] of Object.entries(RU)) if (!EN_RU[en] && en !== 'TRUE' && en !== 'FALSE') EN_RU[en] = ru;
export function localize(src) {
  return tokenize(src).map((t) => {
    if (t.type === 'func') return EN_RU[t.name] || t.text;
    if (t.type === 'sep') return ';';
    if (t.type === 'num') return t.text.replace('.', ',');
    if (t.type === 'bool') return t.value ? 'ИСТИНА' : 'ЛОЖЬ';
    return t.text;
  }).join('');
}

const refText = (t, r, c) => {
  let sheet = '';
  if (t.sheet) sheet = /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*$/.test(t.sheet) ? `${t.sheet}!` : `'${t.sheet.replace(/'/g, "''")}'!`;
  return `${sheet}${t.absC ? '$' : ''}${colName(c)}${t.absR ? '$' : ''}${r + 1}`;
};

// Сдвиг ссылок: fn({ r, c, absR, absC, sheet }) → { r, c } | null (null → #REF!)
// Диапазоны «A1:B3» переписываются по концам; если удалён конец — диапазон сжимается (endFn), полностью — #REF!
export function rewriteRefs(src, fn) {
  const toks = tokenize(src);
  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'ref') { out += t.text; continue; }
    // диапазон?
    const isRange = toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref';
    if (isRange) {
      const u = toks[i + 2];
      const a = fn(t, 'start');
      const b = fn({ ...u, sheet: u.sheet || t.sheet }, 'end');
      const ta = a && 'sheet' in a ? { ...t, sheet: a.sheet } : t; // fn может сменить имя листа (переименование)
      out += a && b && (a.r <= b.r && a.c <= b.c) ? `${refText(ta, a.r, a.c)}:${refText({ ...u, sheet: null }, b.r, b.c)}` : '#REF!';
      i += 2;
      continue;
    }
    const p = fn(t, 'single');
    out += p ? refText(p && 'sheet' in p ? { ...t, sheet: p.sheet } : t, p.r, p.c) : '#REF!';
  }
  return out;
}

// Копирование формулы со смещением (dr, dc): относительные ссылки сдвигаются, абсолютные ($) — нет
export function offsetFormula(src, dr, dc) {
  return rewriteRefs(src, (t) => {
    const r = t.absR ? t.r : t.r + dr;
    const c = t.absC ? t.c : t.c + dc;
    return r < 0 || c < 0 ? null : { r, c };
  });
}

// Вставка/удаление строк или столбцов на листе sheet: axis 'row'|'col', at, count (>0 — вставка, <0 — удаление)
export function shiftFormula(src, { sheet, current, axis, at, count }) {
  return rewriteRefs(src, (t, part) => {
    if ((t.sheet || current) !== sheet) return { r: t.r, c: t.c };
    const k = axis === 'row' ? 'r' : 'c';
    let v = t[k];
    if (count > 0) { if (v >= at) v += count; }
    else {
      const n = -count;
      if (v >= at + n) v -= n;
      else if (v >= at) {
        if (part === 'single') return null;
        v = part === 'start' ? at : at - 1; // конец диапазона внутри удалённого — сжать
        if (v < 0) return null;
      }
    }
    return { ...t, [k]: v };
  });
}

// Ссылки формулы (для подсветки/зависимостей): [{ sheet, r1, c1, r2, c2 }]
export function formulaRefs(src) {
  const toks = tokenize(src);
  const out = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'ref') continue;
    if (toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref') {
      const u = toks[i + 2];
      out.push({ sheet: t.sheet, r1: Math.min(t.r, u.r), c1: Math.min(t.c, u.c), r2: Math.max(t.r, u.r), c2: Math.max(t.c, u.c) });
      i += 2;
    } else out.push({ sheet: t.sheet, r1: t.r, c1: t.c, r2: t.r, c2: t.c });
  }
  return out;
}
