// Библиотека функций Excel (этап 12, часть 4): математика, логические, условные итоги, поиск и ссылки, текст,
// дата и время, статистика, финансы, информация. Подключается в formula.js: Object.assign(FN, functions(h)),
// h — помощники движка (приведение типов, сравнение, условия, даты). Функция получает (args, ctx, nodes):
// args — вычисленные аргументы (диапазон — двумерный массив), nodes — их узлы разбора (нужны СТРОКА, СМЕЩ, ЕФОРМУЛА…).
// Значения сверяются с Excel — tests/formulas.test.js.
import { formatValue } from './format.js';

export function functions(h) {
  const { ERR, err, isErr, toNum, toText, toBool, flat, nums, firstErr, criterion, compare, exactMatcher, wildcard,
    dateToSerial, serialToDate, toArr, numText, approxIndex } = h;
  const num = (x, def) => (x == null && def !== undefined ? def : toNum(x));
  const int = (x, def) => { const n = num(x, def); return isErr(n) ? n : Math.trunc(n); };
  const all = (args) => args.flatMap(flat);
  const fix = (x) => Math.round(x * 1e12) / 1e12; // убрать хвосты двоичной арифметики перед округлением
  const sortedNums = (a) => { const n = nums(a); return isErr(n) ? n : [...n].sort((x, y) => x - y); };
  const mean = (n) => n.reduce((s, x) => s + x, 0) / n.length;
  const devsq = (n) => { const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0); };
  const isNA = (v) => isErr(v) && v.error === ERR.NA;
  const bool = (x, def) => { if (x == null) return def; const b = toBool(x); return isErr(b) ? b : b; };

  // ---------- математика ----------
  const ceilTo = (n, s) => Math.ceil(fix(n / s)) * s;
  const floorTo = (n, s) => Math.floor(fix(n / s)) * s;
  const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };
  const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  const math = {
    MROUND: ([x, m]) => {
      const n = num(x); const k = num(m);
      if (firstErr(n, k)) return firstErr(n, k);
      if (k === 0) return 0;
      if (n * k < 0) return err(ERR.NUM);
      return Math.round(fix(n / k)) * k; // n/k > 0: половина — от нуля
    },
    CEILING: ([x, s]) => {
      const n = num(x); const k = num(s);
      if (firstErr(n, k)) return firstErr(n, k);
      if (k === 0) return 0;
      if (n > 0 && k < 0) return err(ERR.NUM);
      return ceilTo(n, k);
    },
    'CEILING.MATH': ([x, s, mode]) => {
      const n = num(x); let k = num(s, 1); const md = num(mode, 0);
      if (firstErr(n, k, md)) return firstErr(n, k, md);
      k = Math.abs(k);
      if (k === 0) return 0;
      return n < 0 && md ? -ceilTo(-n, k) : ceilTo(n, k);
    },
    'CEILING.PRECISE': ([x, s]) => math['CEILING.MATH']([x, s]),
    FLOOR: ([x, s]) => {
      const n = num(x); const k = num(s);
      if (firstErr(n, k)) return firstErr(n, k);
      if (k === 0) return n === 0 ? 0 : err(ERR.DIV0);
      if (n > 0 && k < 0) return err(ERR.NUM);
      return floorTo(n, k);
    },
    'FLOOR.MATH': ([x, s, mode]) => {
      const n = num(x); let k = num(s, 1); const md = num(mode, 0);
      if (firstErr(n, k, md)) return firstErr(n, k, md);
      k = Math.abs(k);
      if (k === 0) return 0;
      return n < 0 && md ? -floorTo(-n, k) : floorTo(n, k);
    },
    'FLOOR.PRECISE': ([x, s]) => math['FLOOR.MATH']([x, s]),
    TRUNC: ([x, d]) => {
      const n = num(x); const k = int(d, 0);
      if (firstErr(n, k)) return firstErr(n, k);
      const p = 10 ** k;
      return Math.trunc(fix(n * p)) / p;
    },
    EVEN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.sign(n) * Math.ceil(fix(Math.abs(n) / 2)) * 2; },
    ODD: ([x]) => {
      const n = num(x);
      if (isErr(n)) return n;
      const a = Math.ceil(fix(Math.abs(n)));
      const o = a % 2 ? a : a + 1;
      return n < 0 ? -o : o;
    },
    EXP: ([x]) => { const n = num(x); return isErr(n) ? n : Math.exp(n); },
    LN: ([x]) => { const n = num(x); return isErr(n) ? n : n <= 0 ? err(ERR.NUM) : Math.log(n); },
    LOG: ([x, b]) => {
      const n = num(x); const base = num(b, 10);
      if (firstErr(n, base)) return firstErr(n, base);
      if (n <= 0 || base <= 0) return err(ERR.NUM);
      if (base === 1) return err(ERR.DIV0);
      return base === 10 ? Math.log10(n) : Math.log(n) / Math.log(base);
    },
    LOG10: ([x]) => { const n = num(x); return isErr(n) ? n : n <= 0 ? err(ERR.NUM) : Math.log10(n); },
    FACT: ([x]) => {
      const n = int(x);
      if (isErr(n)) return n;
      if (n < 0) return err(ERR.NUM);
      let r = 1; for (let i = 2; i <= n; i++) r *= i;
      return r;
    },
    GCD: (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      if (n.some((x) => x < 0)) return err(ERR.NUM);
      return n.map(Math.trunc).reduce(gcd2, 0);
    },
    LCM: (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      if (n.some((x) => x < 0)) return err(ERR.NUM);
      return n.map(Math.trunc).reduce((x, y) => (x && y ? (x / gcd2(x, y)) * y : 0), 1);
    },
    QUOTIENT: ([a, b]) => { const x = num(a); const y = num(b); return firstErr(x, y) || (y === 0 ? err(ERR.DIV0) : Math.trunc(x / y)); },
    SIN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.sin(n); },
    COS: ([x]) => { const n = num(x); return isErr(n) ? n : Math.cos(n); },
    TAN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.tan(n); },
    ASIN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.abs(n) > 1 ? err(ERR.NUM) : Math.asin(n); },
    ACOS: ([x]) => { const n = num(x); return isErr(n) ? n : Math.abs(n) > 1 ? err(ERR.NUM) : Math.acos(n); },
    ATAN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.atan(n); },
    ATAN2: ([a, b]) => { const x = num(a); const y = num(b); return firstErr(x, y) || (x === 0 && y === 0 ? err(ERR.DIV0) : Math.atan2(y, x)); },
    SINH: ([x]) => { const n = num(x); return isErr(n) ? n : Math.sinh(n); },
    COSH: ([x]) => { const n = num(x); return isErr(n) ? n : Math.cosh(n); },
    TANH: ([x]) => { const n = num(x); return isErr(n) ? n : Math.tanh(n); },
    DEGREES: ([x]) => { const n = num(x); return isErr(n) ? n : (n * 180) / Math.PI; },
    RADIANS: ([x]) => { const n = num(x); return isErr(n) ? n : (n * Math.PI) / 180; },
    ROMAN: ([x]) => {
      let n = int(x);
      if (isErr(n)) return n;
      if (n < 0 || n > 3999) return err(ERR.VALUE);
      let out = '';
      for (const [v, s] of ROMAN) while (n >= v) { out += s; n -= v; }
      return out;
    },
    ARABIC: ([x]) => {
      if (isErr(x)) return x;
      const s = toText(x).trim().toUpperCase();
      const neg = s.startsWith('-');
      const t = neg ? s.slice(1) : s;
      if (!/^[MDCLXVI]*$/.test(t)) return err(ERR.VALUE);
      const val = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
      let total = 0;
      for (let i = 0; i < t.length; i++) { const v = val[t[i]]; total += i + 1 < t.length && val[t[i + 1]] > v ? -v : v; }
      return neg ? -total : total;
    },
    SUMSQ: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s + x * x, 0); },
    COMBIN: ([a, b]) => {
      const n = int(a); const k = int(b);
      if (firstErr(n, k)) return firstErr(n, k);
      if (n < 0 || k < 0 || k > n) return err(ERR.NUM);
      let r = 1; for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
      return Math.round(r);
    },
    PERMUT: ([a, b]) => {
      const n = int(a); const k = int(b);
      if (firstErr(n, k)) return firstErr(n, k);
      if (n < 0 || k < 0 || k > n) return err(ERR.NUM);
      let r = 1; for (let i = 0; i < k; i++) r *= n - i;
      return r;
    },
    BASE: ([x, radix, len]) => {
      const n = int(x); const r = int(radix); const l = int(len, 0);
      if (firstErr(n, r, l)) return firstErr(n, r, l);
      if (n < 0 || r < 2 || r > 36) return err(ERR.NUM);
      return n.toString(r).toUpperCase().padStart(l, '0');
    },
    DECIMAL: ([x, radix]) => {
      const r = int(radix);
      if (firstErr(x, r)) return firstErr(x, r);
      const s = toText(x).trim().toLowerCase();
      if (r < 2 || r > 36 || !s || [...s].some((ch) => !(parseInt(ch, 36) < r))) return err(ERR.NUM);
      return parseInt(s, r);
    },
    SQRTPI: ([x]) => { const n = num(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.sqrt(n * Math.PI); },
    SUMX2MY2: ([a, b]) => pairSum(a, b, (x, y) => x * x - y * y),
    SUMX2PY2: ([a, b]) => pairSum(a, b, (x, y) => x * x + y * y),
    SUMXMY2: ([a, b]) => pairSum(a, b, (x, y) => (x - y) ** 2),
  };
  function pairSum(a, b, f) {
    const x = flat(a); const y = flat(b);
    if (x.length !== y.length) return err(ERR.NA);
    let s = 0;
    for (let i = 0; i < x.length; i++) {
      if (isErr(x[i])) return x[i];
      if (isErr(y[i])) return y[i];
      if (typeof x[i] === 'number' && typeof y[i] === 'number') s += f(x[i], y[i]);
    }
    return s;
  }

  // ---------- логические и информационные ----------
  const ERROR_TYPES = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 };
  const isRefNode = (n) => n && (n.t === 'ref' || n.t === 'range' || n.t === 'spill');
  const logic = {
    XOR: (a) => {
      let k = 0;
      for (const v of all(a)) {
        if (isErr(v)) return v;
        if (v == null || v === '') continue;
        const b = typeof v === 'string' ? toBool(v) : typeof v === 'number' ? v !== 0 : v;
        if (isErr(b)) continue;
        if (b) k++;
      }
      return k % 2 === 1;
    },
    NA: () => err(ERR.NA),
    ISNA: ([x]) => isNA(x),
    ISERR: ([x]) => isErr(x) && !isNA(x),
    ISLOGICAL: ([x]) => typeof x === 'boolean',
    ISNONTEXT: ([x]) => typeof x !== 'string',
    ISEVEN: ([x]) => { const n = num(x); return isErr(n) ? n : Math.trunc(n) % 2 === 0; },
    ISODD: ([x]) => { const n = num(x); return isErr(n) ? n : Math.abs(Math.trunc(n) % 2) === 1; },
    ISREF: (_a, _ctx, nodes) => isRefNode(nodes[0]),
    ISFORMULA: (_a, ctx, nodes) => {
      const n = nodes[0];
      if (!isRefNode(n)) return err(ERR.VALUE);
      const f = ctx.formulaAt && ctx.formulaAt(n.sheet || ctx.sheet, n.r ?? n.r1, n.c ?? n.c1);
      return f != null;
    },
    FORMULATEXT: (_a, ctx, nodes) => {
      const n = nodes[0];
      if (!isRefNode(n)) return err(ERR.VALUE);
      const f = ctx.formulaAt && ctx.formulaAt(n.sheet || ctx.sheet, n.r ?? n.r1, n.c ?? n.c1);
      return f == null ? err(ERR.NA) : `=${h.localize(f)}`;
    },
    N: ([x]) => (isErr(x) ? x : typeof x === 'number' ? x : typeof x === 'boolean' ? (x ? 1 : 0) : 0),
    TYPE: ([x]) => (Array.isArray(x) && (x.length > 1 || (x[0] || []).length > 1) ? 64 : (() => {
      const v = Array.isArray(x) ? x[0][0] : x;
      return isErr(v) ? 16 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 1;
    })()),
    'ERROR.TYPE': ([x]) => (isErr(x) && ERROR_TYPES[x.error] ? ERROR_TYPES[x.error] : err(ERR.NA)),
    SHEET: (_a, ctx, nodes) => {
      const n = nodes[0];
      const name = n ? (isRefNode(n) ? n.sheet || ctx.sheet : toText(_a[0])) : ctx.sheet;
      const i = ctx.sheetIndex ? ctx.sheetIndex(name) : -1;
      return i < 0 ? err(ERR.NA) : i + 1;
    },
    SHEETS: (_a, ctx) => (ctx.sheetCount ? ctx.sheetCount() : 1),
    CELL: ([info], ctx, nodes) => {
      const n = nodes[1];
      const what = toText(info).toLowerCase();
      const r = n && isRefNode(n) ? n.r ?? n.r1 : ctx.cur ? ctx.cur.r : 0;
      const c = n && isRefNode(n) ? n.c ?? n.c1 : ctx.cur ? ctx.cur.c : 0;
      const sh = (n && n.sheet) || ctx.sheet;
      const v = ctx.get(sh, r, c);
      if (what === 'row') return r + 1;
      if (what === 'col') return c + 1;
      if (what === 'address') return `${sh.toLowerCase() !== String(ctx.sheet).toLowerCase() ? h.sheetPrefix(sh) : ''}$${h.colName(c)}$${r + 1}`;
      if (what === 'contents') return v ?? 0;
      if (what === 'type') return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v';
      return err(ERR.VALUE);
    },
  };

  // ---------- условные итоги ----------
  // пары (диапазон, условие) → маска подходящих позиций; размеры диапазонов должны совпадать
  function ifsMask(pairs, size) {
    const mask = new Array(size).fill(true);
    for (let i = 0; i + 1 < pairs.length; i += 2) {
      const range = flat(pairs[i]);
      if (range.length !== size) return err(ERR.VALUE);
      const test = criterion(pairs[i + 1]);
      for (let j = 0; j < size; j++) if (mask[j] && !test(range[j])) mask[j] = false;
    }
    return mask;
  }
  const pickIfs = (target, pairs) => {
    const vals = flat(target);
    const mask = ifsMask(pairs, vals.length);
    if (isErr(mask)) return mask;
    return vals.filter((v, j) => mask[j]);
  };
  const cond = {
    SUMIFS: ([target, ...pairs]) => { const p = pickIfs(target, pairs); return isErr(p) ? p : p.reduce((s, v) => s + (typeof v === 'number' ? v : 0), 0); },
    COUNTIFS: (pairs) => { const m = ifsMask(pairs, flat(pairs[0]).length); return isErr(m) ? m : m.filter(Boolean).length; },
    AVERAGEIFS: ([target, ...pairs]) => {
      const p = pickIfs(target, pairs);
      if (isErr(p)) return p;
      const n = p.filter((v) => typeof v === 'number');
      return n.length ? mean(n) : err(ERR.DIV0);
    },
    MAXIFS: ([target, ...pairs]) => { const p = pickIfs(target, pairs); if (isErr(p)) return p; const n = p.filter((v) => typeof v === 'number'); return n.length ? Math.max(...n) : 0; },
    MINIFS: ([target, ...pairs]) => { const p = pickIfs(target, pairs); if (isErr(p)) return p; const n = p.filter((v) => typeof v === 'number'); return n.length ? Math.min(...n) : 0; },
  };
  // ПРОМЕЖУТОЧНЫЕ.ИТОГИ и АГРЕГАТ: значения диапазонов без вложенных итогов (и скрытых строк, если надо)
  function rangeValues(arg, node, ctx, { skipHidden, skipNested = true }) {
    if (!node || (node.t !== 'range' && node.t !== 'ref')) return flat(arg);
    const sh = node.sheet || ctx.sheet;
    const r1 = node.t === 'ref' ? node.r : node.r1; const r2 = node.t === 'ref' ? node.r : node.r2;
    const c1 = node.t === 'ref' ? node.c : node.c1; const c2 = node.t === 'ref' ? node.c : node.c2;
    const rows = toArr(arg);
    const out = [];
    rows.forEach((row, i) => {
      const r = r1 + i;
      if (r > r2) return;
      if (skipHidden && ctx.rowHidden && ctx.rowHidden(sh, r)) return;
      row.forEach((v, j) => {
        if (skipNested && ctx.formulaAt) { const f = ctx.formulaAt(sh, r, c1 + j); if (f && /^\s*(SUBTOTAL|AGGREGATE)\(/i.test(f)) return; }
        if (c1 + j <= c2) out.push(v);
      });
    });
    return out;
  }
  const SUB = {
    1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP',
    12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC',
  };
  const subtotal = {
    SUBTOTAL: ([fn, ...refs], ctx, nodes) => {
      const k = int(fn);
      if (isErr(k)) return k;
      const name = SUB[k > 100 ? k - 100 : k];
      if (!name || k % 100 > 11) return err(ERR.VALUE);
      const vals = refs.flatMap((a, i) => rangeValues(a, nodes[i + 1], ctx, { skipHidden: k > 100 }));
      return h.FN[name]([[vals]]);
    },
    AGGREGATE: ([fn, opt, ...rest], ctx, nodes) => {
      const k = int(fn); const o = int(opt, 0);
      if (firstErr(k, o)) return firstErr(k, o);
      const name = SUB[k];
      if (!name) return err(ERR.VALUE);
      const skipHidden = [1, 3, 5, 7].includes(o);
      const skipErrors = [2, 3, 6, 7].includes(o);
      const skipNested = o <= 3;
      const refs = k >= 14 ? rest.slice(0, 1) : rest;
      let vals = refs.flatMap((a, i) => rangeValues(a, nodes[i + 2], ctx, { skipHidden, skipNested }));
      if (skipErrors) vals = vals.filter((v) => !isErr(v));
      return k >= 14 ? h.FN[name]([[vals], rest[1]]) : h.FN[name]([[vals]]);
    },
  };

  // ---------- поиск и ссылки ----------
  function xmatchIndex(val, list, mode = 0, search = 1) {
    const order = search < 0 ? [...list.keys()].reverse() : [...list.keys()];
    if (mode === 2 || mode === 0) {
      const hit = mode === 2 ? exactMatcher(val) : (v) => compare(v, val) === 0 && !(v == null && val !== '' && val != null);
      for (const i of order) if (hit(list[i])) return i;
      return -1;
    }
    let best = -1;
    for (const i of order) {
      const v = list[i];
      if (v == null || typeof v !== typeof val) continue;
      const r = compare(v, val);
      if (r === 0) return i;
      if (mode === -1 && r < 0 && (best < 0 || compare(v, list[best]) > 0)) best = i;
      if (mode === 1 && r > 0 && (best < 0 || compare(v, list[best]) < 0)) best = i;
    }
    return best;
  }
  const refNodeOf = (node, ctx) => {
    if (!node) return null;
    if (node.t === 'ref') return { sheet: node.sheet || ctx.sheet, r1: node.r, c1: node.c, r2: node.r, c2: node.c };
    if (node.t === 'range' && !node.sheet2) return { sheet: node.sheet || ctx.sheet, r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 };
    return null;
  };
  const evalRange = (ctx, rg) => {
    if (rg.r1 < 0 || rg.c1 < 0 || rg.r2 >= 1048576 || rg.c2 >= 16384) return err(ERR.REF);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return ctx.get(rg.sheet, rg.r1, rg.c1);
    return h.evaluate({ t: 'range', sheet: rg.sheet, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, ctx);
  };
  const lookup = {
    HLOOKUP: ([val, table, row, approx]) => {
      if (isErr(val)) return val;
      if (!Array.isArray(table)) return err(ERR.NA);
      const r = int(row);
      if (isErr(r) || r < 1 || r > table.length) return err(ERR.REF);
      const head = table[0];
      const exact = approx === false || approx === 0;
      const j = exact ? head.findIndex(exactMatcher(val)) : approxIndex(head, val);
      return j < 0 ? err(ERR.NA) : (table[r - 1][j] ?? 0);
    },
    LOOKUP: ([val, vec, res]) => {
      if (isErr(val)) return val;
      const a = toArr(vec);
      let keys; let out;
      if (res !== undefined) { keys = flat(a); out = flat(res); }
      else if (a.length >= (a[0] || []).length) { keys = a.map((r) => r[0]); out = a.map((r) => r[r.length - 1]); }
      else { keys = a[0]; out = a[a.length - 1]; }
      const i = approxIndex(keys, val);
      return i < 0 ? err(ERR.NA) : (out[i] ?? 0);
    },
    XLOOKUP: ([val, look, ret, notFound, mode, search]) => {
      if (isErr(val)) return val;
      const L = toArr(look); const R = toArr(ret);
      const vertical = (L[0] || []).length === 1;
      const list = vertical ? L.map((r) => r[0]) : L[0];
      const m = int(mode, 0); const sm = int(search, 1);
      if (firstErr(m, sm)) return firstErr(m, sm);
      const i = xmatchIndex(val, list, m, sm);
      if (i < 0) return notFound !== undefined && notFound !== null ? notFound : err(ERR.NA);
      if (vertical) { if (R.length !== list.length) return err(ERR.VALUE); return R[i].length === 1 ? (R[i][0] ?? 0) : [R[i]]; }
      if ((R[0] || []).length !== list.length) return err(ERR.VALUE);
      return R.length === 1 ? (R[0][i] ?? 0) : R.map((row) => [row[i]]);
    },
    XMATCH: ([val, look, mode, search]) => {
      if (isErr(val)) return val;
      const m = int(mode, 0); const sm = int(search, 1);
      if (firstErr(m, sm)) return firstErr(m, sm);
      const i = xmatchIndex(val, flat(look), m, sm);
      return i < 0 ? err(ERR.NA) : i + 1;
    },
    CHOOSE: ([i, ...vals]) => { const k = int(i); if (isErr(k)) return k; return k < 1 || k > vals.length ? err(ERR.VALUE) : (vals[k - 1] ?? 0); },
    ROW: (_a, ctx, nodes) => {
      const rg = refNodeOf(nodes[0], ctx);
      if (!nodes[0]) return ctx.cur ? ctx.cur.r + 1 : err(ERR.VALUE);
      if (!rg) return err(ERR.VALUE);
      return rg.r1 === rg.r2 ? rg.r1 + 1 : Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => [rg.r1 + i + 1]);
    },
    COLUMN: (_a, ctx, nodes) => {
      const rg = refNodeOf(nodes[0], ctx);
      if (!nodes[0]) return ctx.cur ? ctx.cur.c + 1 : err(ERR.VALUE);
      if (!rg) return err(ERR.VALUE);
      return rg.c1 === rg.c2 ? rg.c1 + 1 : [Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, j) => rg.c1 + j + 1)];
    },
    ADDRESS: ([row, col, abs, a1, sheet]) => {
      const r = int(row); const c = int(col); const k = int(abs, 1);
      if (firstErr(r, c, k)) return firstErr(r, c, k);
      if (r < 1 || c < 1) return err(ERR.VALUE);
      const style = bool(a1, true);
      const absR = k === 1 || k === 2; const absC = k === 1 || k === 3;
      const ref = style === false
        ? `R${absR ? r : `[${r}]`}C${absC ? c : `[${c}]`}`
        : `${absC ? '$' : ''}${h.colName(c - 1)}${absR ? '$' : ''}${r}`;
      return sheet != null && sheet !== '' ? `${h.sheetPrefix(toText(sheet))}${ref}` : ref;
    },
    OFFSET: ([, rows, cols, height, width], ctx, nodes) => {
      const base = refNodeOf(nodes[0], ctx);
      if (!base) return err(ERR.VALUE);
      const dr = int(rows, 0); const dc = int(cols, 0);
      const hh = int(height, base.r2 - base.r1 + 1); const ww = int(width, base.c2 - base.c1 + 1);
      if (firstErr(dr, dc, hh, ww)) return firstErr(dr, dc, hh, ww);
      if (hh < 1 || ww < 1) return err(ERR.REF);
      return evalRange(ctx, { sheet: base.sheet, r1: base.r1 + dr, c1: base.c1 + dc, r2: base.r1 + dr + hh - 1, c2: base.c1 + dc + ww - 1 });
    },
    INDIRECT: ([text, a1], ctx) => {
      if (isErr(text)) return text;
      if (bool(a1, true) === false) return err(ERR.REF); // стиль R1C1 — пока не поддерживается
      let node;
      try { node = h.parse(toText(text).trim()); } catch { return err(ERR.REF); }
      if (node.t === 'name') {
        const def = ctx.nameDef && ctx.nameDef(node.v);
        let target = null;
        try { target = def != null ? h.parse(def) : null; } catch { target = null; }
        return target && (target.t === 'ref' || target.t === 'range') ? h.evaluate(target, ctx) : err(ERR.REF);
      }
      if (node.t !== 'ref' && node.t !== 'range') return err(ERR.REF);
      return h.evaluate(node, ctx);
    },
    'RANK.EQ': (args) => h.FN.RANK(args),
    'RANK.AVG': ([val, range, order]) => {
      const n = num(val);
      if (isErr(n)) return n;
      const list = flat(range).filter((x) => typeof x === 'number');
      if (!list.includes(n)) return err(ERR.NA);
      const desc = !order;
      const before = list.filter((x) => (desc ? x > n : x < n)).length;
      const same = list.filter((x) => x === n).length;
      return before + (same + 1) / 2;
    },
    HYPERLINK: ([link, name]) => (name !== undefined ? name : link),
  };

  // ---------- текст ----------
  const cp1251 = typeof TextDecoder !== 'undefined' ? new TextDecoder('windows-1251') : null;
  const CP = cp1251 ? Array.from({ length: 256 }, (_, i) => cp1251.decode(new Uint8Array([i]))) : [];
  const text = {
    PROPER: ([x]) => (isErr(x) ? x : toText(x).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a, b) => a + b.toUpperCase())),
    SEARCH: ([what, where, st]) => {
      const s = int(st, 1);
      if (firstErr(what, where, s)) return firstErr(what, where, s);
      const t = toText(where); const w = toText(what);
      if (s < 1 || s > t.length + 1) return err(ERR.VALUE);
      const re = wildcard(w);
      if (!re) { const i = t.toLowerCase().indexOf(w.toLowerCase(), s - 1); return i < 0 ? err(ERR.VALUE) : i + 1; }
      const part = new RegExp(re.source.slice(1, -1), 'i');
      for (let i = s - 1; i <= t.length; i++) { const m = part.exec(t.slice(i)); if (m && m.index === 0) return i + 1; }
      return err(ERR.VALUE);
    },
    REPLACE: ([old, st, n, rep]) => {
      const a = int(st); const k = int(n);
      if (firstErr(old, a, k, rep)) return firstErr(old, a, k, rep);
      if (a < 1 || k < 0) return err(ERR.VALUE);
      const t = toText(old);
      return t.slice(0, a - 1) + toText(rep) + t.slice(a - 1 + k);
    },
    TEXTJOIN: ([delim, ignore, ...rest]) => {
      const e = firstErr(delim, ignore, ...all(rest));
      if (e) return e;
      const d = flat(delim).map(toText);
      const skip = bool(ignore, true);
      const parts = all(rest).map(toText).filter((s) => !(skip && s === ''));
      let out = '';
      parts.forEach((p, i) => { if (i) out += d[(i - 1) % d.length]; out += p; });
      return out.length > 32767 ? err(ERR.CALC) : out;
    },
    EXACT: ([a, b]) => firstErr(a, b) || toText(a) === toText(b),
    CHAR: ([x]) => { const n = int(x); if (isErr(n)) return n; return n < 1 || n > 255 ? err(ERR.VALUE) : CP[n] || String.fromCharCode(n); },
    CODE: ([x]) => {
      if (isErr(x)) return x;
      const ch = toText(x)[0];
      if (!ch) return err(ERR.VALUE);
      const i = CP.indexOf(ch);
      return i > 0 ? i : ch.charCodeAt(0) < 128 ? ch.charCodeAt(0) : 63;
    },
    UNICHAR: ([x]) => { const n = int(x); if (isErr(n)) return n; return n < 1 || n > 0x10ffff ? err(ERR.VALUE) : String.fromCodePoint(n); },
    UNICODE: ([x]) => { if (isErr(x)) return x; const t = toText(x); return t ? t.codePointAt(0) : err(ERR.VALUE); },
    CLEAN: ([x]) => (isErr(x) ? x : toText(x).replace(/[\u0000-\u001f]/g, '')),
    NUMBERVALUE: ([x, dec, grp]) => {
      if (firstErr(x, dec, grp)) return firstErr(x, dec, grp);
      const d = dec == null ? ',' : toText(dec)[0]; const g = grp == null ? ' ' : toText(grp)[0];
      let t = toText(x).replace(/\s/g, (ch) => (g === ' ' ? '' : ch));
      let pct = 0;
      while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
      if (g && g !== ' ') t = t.split(g).join('');
      if (d) t = t.split(d).join('.');
      if (t === '') return 0;
      return /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t) ? +t / 100 ** pct : err(ERR.VALUE);
    },
    T: ([x]) => (isErr(x) ? x : typeof x === 'string' ? x : ''),
    TEXT: ([x, fmt]) => {
      if (firstErr(x, fmt)) return firstErr(x, fmt);
      const code = ruFormat(toText(fmt));
      if (isErr(code)) return code;
      const n = typeof x === 'number' ? x : typeof x === 'string' ? (h.textToNum(x) ?? x) : x;
      if (typeof n !== 'number') return toText(x);
      // код без заполнителей (например, латинское «dd.mm.yyyy» — для русского Excel просто текст) — выводится как есть
      const bare = code.replace(/"[^"]*"/g, '').replace(/\\./g, '');
      if (!/[0#?dmyhs@]/i.test(bare)) return code.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
      return formatValue(n, code).replace(/\u00a0/g, ' '); // ТЕКСТ отделяет разряды обычным пробелом
    },
    FIXED: ([x, d, noCommas]) => {
      const n = num(x); const k = int(d, 2);
      if (firstErr(n, k)) return firstErr(n, k);
      const v = h.round(n, k);
      const s = Math.abs(v).toFixed(Math.max(0, k));
      const [i, f] = s.split('.');
      const body = bool(noCommas, false) ? i : i.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
      return `${v < 0 ? '-' : ''}${body}${f ? `,${f}` : ''}`;
    },
    DOLLAR: ([x, d]) => { const t = text.FIXED([x, d ?? 2]); return isErr(t) ? t : `${t} ₽`; },
    TEXTBEFORE: ([x, delim, inst]) => {
      if (firstErr(x, delim)) return firstErr(x, delim);
      const t = toText(x); const d = toText(delim); const k = int(inst, 1);
      if (isErr(k)) return k;
      let at = -1;
      if (k > 0) { for (let i = 0; i < k; i++) { at = t.indexOf(d, at + 1); if (at < 0) return err(ERR.NA); } }
      else { at = t.length; for (let i = 0; i < -k; i++) { at = t.lastIndexOf(d, at - 1); if (at < 0) return err(ERR.NA); } }
      return t.slice(0, at);
    },
    TEXTAFTER: ([x, delim, inst]) => {
      if (firstErr(x, delim)) return firstErr(x, delim);
      const t = toText(x); const d = toText(delim); const k = int(inst, 1);
      if (isErr(k)) return k;
      let at = -1;
      if (k > 0) { for (let i = 0; i < k; i++) { at = t.indexOf(d, at + 1); if (at < 0) return err(ERR.NA); } }
      else { at = t.length; for (let i = 0; i < -k; i++) { at = t.lastIndexOf(d, at - 1); if (at < 0) return err(ERR.NA); } }
      return t.slice(at + d.length);
    },
    TEXTSPLIT: ([x, colDelim, rowDelim]) => {
      if (isErr(x)) return x;
      const t = toText(x);
      const rows = rowDelim != null ? t.split(toText(rowDelim)) : [t];
      const split = rows.map((r) => (colDelim != null ? r.split(toText(colDelim)) : [r]));
      const w = Math.max(...split.map((r) => r.length));
      return split.map((r) => Array.from({ length: w }, (_, j) => (j < r.length ? r[j] : err(ERR.NA))));
    },
  };
  // Код формата ТЕКСТ по-русски (как в русском Excel): Д/М/Г/ч/м/с — дата и время, «,» — десятичная, пробел —
  // разряды; латинские d m y h s — просто буквы. «0.00» русский Excel не принимает (#ЗНАЧ!).
  function ruFormat(code) {
    let out = ''; let q = false;
    const isDate = /[ДМГчсДдГг]/.test(code.replace(/"[^"]*"/g, ''));
    for (let i = 0; i < code.length; i++) {
      const ch = code[i];
      if (ch === '"') { q = !q; out += ch; continue; }
      if (q) { out += ch; continue; }
      if (ch === '\\') { out += ch + (code[i + 1] || ''); i++; continue; }
      if (/[Дд]/.test(ch)) out += 'd';
      else if (/[Гг]/.test(ch)) out += 'y';
      else if (ch === 'М') out += 'm';
      else if (ch === 'м') out += 'm';
      else if (ch === 'ч' || ch === 'Ч') out += 'h';
      else if (ch === 'с' || ch === 'С') out += 's';
      else if (/[dmyhsDMYHS]/.test(ch)) out += `\\${ch}`;
      else if (!isDate && ch === ',') out += '.';
      else if (!isDate && ch === '.' && /[0#]/.test(code[i - 1] || '') && /[0#]/.test(code[i + 1] || '')) return err(ERR.VALUE);
      else if (!isDate && ch === ' ' && /[#0]/.test(code[i - 1] || '') && /[#0]/.test(code[i + 1] || '')) out += ',';
      else out += ch;
    }
    return out;
  }

  // ---------- дата и время ----------
  const dateParts = (x) => { const n = num(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : serialToDate(n); };
  const WEEKEND = { 1: '0000011', 2: '1000001', 3: '1100000', 4: '0110000', 5: '0011000', 6: '0001100', 7: '0000110', 11: '0000001', 12: '1000000', 13: '0100000', 14: '0010000', 15: '0001000', 16: '0000100', 17: '0000010' };
  const weekendMask = (w) => {
    if (w == null) return '0000011';
    if (typeof w === 'string') return /^[01]{7}$/.test(w) && w !== '1111111' ? w : null;
    return WEEKEND[Math.trunc(w)] || null;
  };
  // день недели: 0 — понедельник … 6 — воскресенье (серийный номер 1900: 1 = воскресенье 01.01.1900)
  const dow = (n) => (((Math.floor(n) - 2) % 7) + 7) % 7;
  const holidaysOf = (hol) => new Set(flat(hol ?? []).filter((v) => typeof v === 'number').map(Math.floor));
  function workday(start, days, mask, hol) {
    let d = Math.floor(start); let left = Math.trunc(days);
    const step = left < 0 ? -1 : 1;
    while (left !== 0) { d += step; if (mask[dow(d)] === '0' && !hol.has(d)) left -= step; }
    return d;
  }
  function networkdays(a, b, mask, hol) {
    let s = Math.floor(a); let e = Math.floor(b); let sign = 1;
    if (e < s) { [s, e] = [e, s]; sign = -1; }
    let k = 0;
    for (let d = s; d <= e; d++) if (mask[dow(d)] === '0' && !hol.has(d)) k++;
    return k * sign;
  }
  const lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const addMonths = (n, k) => {
    const d = serialToDate(n);
    const total = d.y * 12 + (d.m - 1) + k;
    const y = Math.floor(total / 12); const m = (total % 12) + 1;
    return { y, m, d: Math.min(d.d, lastDay(y, m)) };
  };
  const isLeap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const date = {
    WEEKDAY: ([x, type]) => {
      const n = num(x); const t = int(type, 1);
      if (firstErr(n, t)) return firstErr(n, t);
      const w = dow(n); // 0 = пн
      if (t === 1 || t === 17) return ((w + 1) % 7) + 1;
      if (t === 2 || t === 11) return w + 1;
      if (t === 3) return w;
      if (t >= 12 && t <= 16) return ((w - (t - 11) + 7) % 7) + 1;
      return err(ERR.NUM);
    },
    WEEKNUM: ([x, type]) => {
      const n = num(x); const t = int(type, 1);
      if (firstErr(n, t)) return firstErr(n, t);
      if (t === 21) return date.ISOWEEKNUM([n]);
      const startDow = { 1: 6, 17: 6, 2: 0, 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5 }[t];
      if (startDow === undefined) return err(ERR.NUM);
      const { y } = serialToDate(n);
      const jan1 = dateToSerial(y, 1, 1);
      const offset = (dow(jan1) - startDow + 7) % 7;
      return Math.floor((Math.floor(n) - jan1 + offset) / 7) + 1;
    },
    ISOWEEKNUM: ([x]) => {
      const n = num(x);
      if (isErr(n)) return n;
      const d = Math.floor(n);
      const thursday = d - dow(d) + 3;
      const { y } = serialToDate(thursday);
      return Math.floor((thursday - dateToSerial(y, 1, 1)) / 7) + 1;
    },
    EDATE: ([x, k]) => { const n = num(x); const m = int(k); if (firstErr(n, m)) return firstErr(n, m); const d = addMonths(n, m); return dateToSerial(d.y, d.m, d.d); },
    EOMONTH: ([x, k]) => { const n = num(x); const m = int(k); if (firstErr(n, m)) return firstErr(n, m); const d = addMonths(n, m); return dateToSerial(d.y, d.m, lastDay(d.y, d.m)); },
    WORKDAY: ([s, d, hol]) => { const a = num(s); const k = num(d); return firstErr(a, k) || workday(a, k, '0000011', holidaysOf(hol)); },
    'WORKDAY.INTL': ([s, d, w, hol]) => {
      const a = num(s); const k = num(d);
      if (firstErr(a, k)) return firstErr(a, k);
      const mask = weekendMask(w);
      return mask ? workday(a, k, mask, holidaysOf(hol)) : err(ERR.NUM);
    },
    NETWORKDAYS: ([s, e, hol]) => { const a = num(s); const b = num(e); return firstErr(a, b) || networkdays(a, b, '0000011', holidaysOf(hol)); },
    'NETWORKDAYS.INTL': ([s, e, w, hol]) => {
      const a = num(s); const b = num(e);
      if (firstErr(a, b)) return firstErr(a, b);
      const mask = weekendMask(w);
      return mask ? networkdays(a, b, mask, holidaysOf(hol)) : err(ERR.NUM);
    },
    DATEDIF: ([s, e, unit]) => {
      const a = num(s); const b = num(e);
      if (firstErr(a, b, unit)) return firstErr(a, b, unit);
      if (b < a) return err(ERR.NUM);
      const x = serialToDate(a); const y = serialToDate(b);
      const u = toText(unit).toLowerCase();
      let months = (y.y - x.y) * 12 + (y.m - x.m);
      if (y.d < x.d) months--;
      if (u === 'd') return Math.floor(b) - Math.floor(a);
      if (u === 'm') return months;
      if (u === 'y') return Math.floor(months / 12);
      if (u === 'ym') return months % 12;
      if (u === 'md') {
        if (y.d >= x.d) return y.d - x.d;
        const pm = y.m === 1 ? 12 : y.m - 1; const py = y.m === 1 ? y.y - 1 : y.y;
        return lastDay(py, pm) - x.d + y.d;
      }
      if (u === 'yd') {
        let start = dateToSerial(y.y, x.m, Math.min(x.d, lastDay(y.y, x.m)));
        if (start > Math.floor(b)) start = dateToSerial(y.y - 1, x.m, Math.min(x.d, lastDay(y.y - 1, x.m)));
        return Math.floor(b) - start;
      }
      return err(ERR.NUM);
    },
    DAYS: ([e, s]) => { const b = num(e); const a = num(s); return firstErr(a, b) || Math.floor(b) - Math.floor(a); },
    DAYS360: ([s, e, method]) => {
      const a = num(s); const b = num(e);
      if (firstErr(a, b)) return firstErr(a, b);
      const x = serialToDate(a); const y = serialToDate(b);
      let d1 = x.d; let d2 = y.d;
      if (bool(method, false)) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
      else {
        if (d1 === 31 || (x.m === 2 && d1 === lastDay(x.y, 2))) d1 = 30;
        if (d2 === 31 && d1 >= 30) d2 = 30;
      }
      return (y.y - x.y) * 360 + (y.m - x.m) * 30 + (d2 - d1);
    },
    YEARFRAC: ([s, e, basis]) => {
      let a = num(s); let b = num(e); const k = int(basis, 0);
      if (firstErr(a, b, k)) return firstErr(a, b, k);
      a = Math.floor(a); b = Math.floor(b);
      if (a > b) [a, b] = [b, a];
      const x = serialToDate(a); const y = serialToDate(b);
      if (k === 0) {
        let d1 = x.d; let d2 = y.d;
        const lastFebX = x.m === 2 && d1 === lastDay(x.y, 2); const lastFebY = y.m === 2 && d2 === lastDay(y.y, 2);
        if (lastFebX && lastFebY) d2 = 30;
        if (lastFebX) d1 = 30;
        if (d2 === 31 && d1 >= 30) d2 = 30;
        if (d1 === 31) d1 = 30;
        return ((y.y - x.y) * 360 + (y.m - x.m) * 30 + (d2 - d1)) / 360;
      }
      if (k === 1) {
        const days = b - a;
        if (x.y === y.y || (y.y === x.y + 1 && (y.m < x.m || (y.m === x.m && y.d <= x.d)))) {
          let year = 365;
          if (x.y === y.y) year = isLeap(x.y) ? 366 : 365;
          else if ((isLeap(x.y) && (x.m < 3)) || (isLeap(y.y) && (y.m > 2 || (y.m === 2 && y.d === 29)))) year = 366;
          return days / year;
        }
        let total = 0;
        for (let yy = x.y; yy <= y.y; yy++) total += isLeap(yy) ? 366 : 365;
        return days / (total / (y.y - x.y + 1));
      }
      if (k === 2) return (b - a) / 360;
      if (k === 3) return (b - a) / 365;
      if (k === 4) {
        const d1 = Math.min(x.d, 30); const d2 = Math.min(y.d, 30);
        return ((y.y - x.y) * 360 + (y.m - x.m) * 30 + (d2 - d1)) / 360;
      }
      return err(ERR.NUM);
    },
    TIME: ([hh, mm, ss]) => {
      const a = int(hh); const b = int(mm); const c = int(ss);
      if (firstErr(a, b, c)) return firstErr(a, b, c);
      const t = a * 3600 + b * 60 + c;
      if (t < 0) return err(ERR.NUM);
      return (t % 86400) / 86400;
    },
    HOUR: ([x]) => { const n = num(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.floor(Math.round((n % 1) * 86400) / 3600) % 24; },
    MINUTE: ([x]) => { const n = num(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.floor(Math.round((n % 1) * 86400) / 60) % 60; },
    SECOND: ([x]) => { const n = num(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.round((n % 1) * 86400) % 60; },
    DATEVALUE: ([x]) => {
      if (isErr(x)) return x;
      const n = h.textToNum(toText(x));
      return n == null || /^\s*[-+]?[\d\s]+([.,]\d+)?\s*$/.test(toText(x)) ? err(ERR.VALUE) : Math.floor(n);
    },
    TIMEVALUE: ([x]) => {
      if (isErr(x)) return x;
      const n = h.textToNum(toText(x));
      return n == null ? err(ERR.VALUE) : n - Math.floor(n);
    },
  };

  // ---------- статистика ----------
  const numsA = (a) => {
    const out = [];
    for (const v of all(a)) {
      if (isErr(v)) return v;
      if (v == null) continue;
      out.push(typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0);
    }
    return out;
  };
  const kth = (a, k, largest) => {
    const n = sortedNums([a]);
    const i = int(k);
    if (firstErr(n, i)) return firstErr(n, i);
    if (i < 1 || i > n.length) return err(ERR.NUM);
    return largest ? n[n.length - i] : n[i - 1];
  };
  const pctInc = (n, p) => {
    if (!n.length || p < 0 || p > 1) return err(ERR.NUM);
    const pos = p * (n.length - 1); const lo = Math.floor(pos);
    return n[lo] + (pos - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]);
  };
  const pctExc = (n, p) => {
    const pos = p * (n.length + 1);
    if (!n.length || pos < 1 || pos > n.length) return err(ERR.NUM);
    const lo = Math.floor(pos);
    return n[lo - 1] + (pos - lo) * ((n[lo] ?? n[lo - 1]) - n[lo - 1]);
  };
  const pairs = (ys, xs) => {
    const y = flat(ys); const x = flat(xs);
    if (y.length !== x.length) return err(ERR.NA);
    const py = []; const px = [];
    for (let i = 0; i < y.length; i++) {
      if (isErr(y[i])) return y[i];
      if (isErr(x[i])) return x[i];
      if (typeof y[i] === 'number' && typeof x[i] === 'number') { py.push(y[i]); px.push(x[i]); }
    }
    return [py, px];
  };
  const linreg = (ys, xs) => {
    const p = pairs(ys, xs);
    if (isErr(p)) return p;
    const [y, x] = p;
    if (x.length < 2) return err(ERR.DIV0);
    const mx = mean(x); const my = mean(y);
    let sxy = 0; let sxx = 0;
    for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; }
    if (sxx === 0) return err(ERR.DIV0);
    const slope = sxy / sxx;
    return { slope, intercept: my - slope * mx };
  };
  // нормальное распределение: erf по Коди (точность ~1e-16), обратная — Акклам + шаг Галлея
  function erfc(x) {
    const z = Math.abs(x);
    const t = 1 / (1 + 0.5 * z);
    const r = t * Math.exp(-z * z - 1.26551223 + t * (1.00002368 + t * (0.37409196 + t * (0.09678418 + t * (-0.18628806
      + t * (0.27886807 + t * (-1.13520398 + t * (1.48851587 + t * (-0.82215223 + t * 0.17087277)))))))));
    return x >= 0 ? r : 2 - r;
  }
  // Φ(z) — точный вариант через ряд/цепную дробь (Маршалья): для |z| ≤ 37
  function normCdf(z) {
    if (z < -37) return 0;
    if (z > 37) return 1;
    if (Math.abs(z) < 7) {
      // ряд Тейлора для Φ (Маршалья, 2004) — до 1e-15
      let s = z; let t = z; const q = z * z;
      for (let i = 3; i < 400; i += 2) { t *= q / i; const ns = s + t; if (ns === s) break; s = ns; }
      return 0.5 + s * Math.exp(-0.5 * q - 0.91893853320467274178);
    }
    return 0.5 * erfc(-z / Math.SQRT2);
  }
  const normPdf = (z) => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
  function normInv(p) {
    const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
    const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
    const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
    const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
    let x;
    if (p < 0.02425) { const q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
    else if (p > 1 - 0.02425) { const q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
    else { const q = p - 0.5; const r = q * q; x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
    for (let i = 0; i < 3; i++) { const e = normCdf(x) - p; const u = e / normPdf(x); x -= u / (1 + (x * u) / 2); }
    return x;
  }
  const stat = {
    AVERAGEA: (a) => { const n = numsA(a); return isErr(n) ? n : n.length ? mean(n) : err(ERR.DIV0); },
    MAXA: (a) => { const n = numsA(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; },
    MINA: (a) => { const n = numsA(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; },
    LARGE: ([a, k]) => kth(a, k, true),
    SMALL: ([a, k]) => kth(a, k, false),
    'MODE.SNGL': (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      const count = new Map(); let best = null; let bc = 1;
      for (const x of n) { const c = (count.get(x) || 0) + 1; count.set(x, c); if (c > bc) { bc = c; best = x; } }
      if (best === null) return err(ERR.NA);
      // при равных частотах — первое встретившееся
      return n.find((x) => count.get(x) === bc);
    },
    'STDEV.S': (a) => { const n = nums(a); return isErr(n) ? n : n.length < 2 ? err(ERR.DIV0) : Math.sqrt(devsq(n) / (n.length - 1)); },
    'STDEV.P': (a) => { const n = nums(a); return isErr(n) ? n : !n.length ? err(ERR.DIV0) : Math.sqrt(devsq(n) / n.length); },
    'VAR.S': (a) => { const n = nums(a); return isErr(n) ? n : n.length < 2 ? err(ERR.DIV0) : devsq(n) / (n.length - 1); },
    'VAR.P': (a) => { const n = nums(a); return isErr(n) ? n : !n.length ? err(ERR.DIV0) : devsq(n) / n.length; },
    'QUARTILE.INC': ([a, q]) => { const n = sortedNums([a]); const k = int(q); if (firstErr(n, k)) return firstErr(n, k); return k < 0 || k > 4 ? err(ERR.NUM) : pctInc(n, k / 4); },
    'QUARTILE.EXC': ([a, q]) => { const n = sortedNums([a]); const k = int(q); if (firstErr(n, k)) return firstErr(n, k); return k <= 0 || k >= 4 ? err(ERR.NUM) : pctExc(n, k / 4); },
    'PERCENTILE.INC': ([a, p]) => { const n = sortedNums([a]); const k = num(p); return firstErr(n, k) || pctInc(n, k); },
    'PERCENTILE.EXC': ([a, p]) => { const n = sortedNums([a]); const k = num(p); return firstErr(n, k) || pctExc(n, k); },
    'PERCENTRANK.INC': ([a, x, sig]) => {
      const n = sortedNums([a]); const v = num(x); const s = int(sig, 3);
      if (firstErr(n, v, s)) return firstErr(n, v, s);
      if (!n.length || v < n[0] || v > n[n.length - 1] || s < 1) return err(n.length ? ERR.NA : ERR.NUM);
      let rank;
      const exact = n.indexOf(v);
      if (exact >= 0) rank = n.filter((y) => y < v).length / (n.length - 1);
      else {
        const i = n.findIndex((y) => y > v) - 1;
        const lo = n.filter((y) => y < n[i + 1] && y <= n[i]).length - 1;
        rank = (lo + (v - n[i]) / (n[i + 1] - n[i])) / (n.length - 1);
      }
      const p = 10 ** s;
      return Math.floor(fix(rank * p)) / p;
    },
    CORREL: ([a, b]) => {
      const p = pairs(a, b);
      if (isErr(p)) return p;
      const [x, y] = p;
      const mx = mean(x); const my = mean(y);
      let sxy = 0; let sxx = 0; let syy = 0;
      for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; }
      return sxx && syy ? sxy / Math.sqrt(sxx * syy) : err(ERR.DIV0);
    },
    SLOPE: ([y, x]) => { const r = linreg(y, x); return isErr(r) ? r : r.slope; },
    INTERCEPT: ([y, x]) => { const r = linreg(y, x); return isErr(r) ? r : r.intercept; },
    'FORECAST.LINEAR': ([x, y, xs]) => { const v = num(x); const r = linreg(y, xs); return firstErr(v, r) || r.intercept + r.slope * v; },
    TREND: ([y, xs, nx]) => {
      const ys = flat(y);
      const xsv = xs !== undefined && xs !== null ? xs : ys.map((_, i) => [i + 1]);
      const r = linreg(y, xsv);
      if (isErr(r)) return r;
      const at = nx !== undefined ? toArr(nx) : toArr(xsv);
      return at.map((row) => row.map((v) => (typeof v === 'number' ? r.intercept + r.slope * v : err(ERR.VALUE))));
    },
    FREQUENCY: ([data, bins]) => {
      const n = nums([data]); const b = nums([bins]);
      if (firstErr(n, b)) return firstErr(n, b);
      const order = b.map((v, i) => [v, i]).sort((x, y) => x[0] - y[0]);
      const out = new Array(b.length + 1).fill(0);
      for (const v of n) { const k = order.findIndex(([edge]) => v <= edge); out[k < 0 ? b.length : order[k][1]]++; }
      return out.map((v) => [v]);
    },
    GEOMEAN: (a) => { const n = nums(a); if (isErr(n)) return n; if (!n.length || n.some((x) => x <= 0)) return err(ERR.NUM); return Math.exp(n.reduce((s, x) => s + Math.log(x), 0) / n.length); },
    HARMEAN: (a) => { const n = nums(a); if (isErr(n)) return n; if (!n.length || n.some((x) => x <= 0)) return err(ERR.NUM); return n.length / n.reduce((s, x) => s + 1 / x, 0); },
    AVEDEV: (a) => { const n = nums(a); if (isErr(n)) return n; if (!n.length) return err(ERR.NUM); const m = mean(n); return n.reduce((s, x) => s + Math.abs(x - m), 0) / n.length; },
    DEVSQ: (a) => { const n = nums(a); return isErr(n) ? n : devsq(n); },
    'NORM.DIST': ([x, m, sd, cum]) => {
      const v = num(x); const mu = num(m); const s = num(sd); const c = bool(cum);
      if (firstErr(v, mu, s, c)) return firstErr(v, mu, s, c);
      if (s <= 0) return err(ERR.NUM);
      const z = (v - mu) / s;
      return c ? normCdf(z) : normPdf(z) / s;
    },
    'NORM.S.DIST': ([z, cum]) => { const v = num(z); const c = bool(cum, true); return firstErr(v, c) || (c ? normCdf(v) : normPdf(v)); },
    'NORM.INV': ([p, m, sd]) => {
      const q = num(p); const mu = num(m); const s = num(sd);
      if (firstErr(q, mu, s)) return firstErr(q, mu, s);
      if (q <= 0 || q >= 1 || s <= 0) return err(ERR.NUM);
      return mu + s * normInv(q);
    },
    'NORM.S.INV': ([p]) => { const q = num(p); if (isErr(q)) return q; return q <= 0 || q >= 1 ? err(ERR.NUM) : normInv(q); },
  };
  // старые имена тех же функций
  Object.assign(stat, {
    MODE: stat['MODE.SNGL'], STDEV: stat['STDEV.S'], STDEVP: stat['STDEV.P'], VAR: stat['VAR.S'], VARP: stat['VAR.P'],
    QUARTILE: stat['QUARTILE.INC'], PERCENTILE: stat['PERCENTILE.INC'], PERCENTRANK: stat['PERCENTRANK.INC'], FORECAST: stat['FORECAST.LINEAR'],
    NORMDIST: stat['NORM.DIST'], NORMINV: stat['NORM.INV'], NORMSDIST: ([z]) => stat['NORM.S.DIST']([z, true]), NORMSINV: stat['NORM.S.INV'],
  });

  // ---------- финансы ----------
  const pmt = (r, n, pv, fv, t) => (r === 0 ? -(pv + fv) / n : -(r * (fv + pv * (1 + r) ** n)) / ((1 + r * t) * ((1 + r) ** n - 1)));
  const fvf = (r, n, p, pv, t) => (r === 0 ? -(pv + p * n) : -(pv * (1 + r) ** n + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r));
  const solve = (f, guess) => {
    let x = guess;
    for (let i = 0; i < 100; i++) {
      const y = f(x); const d = (f(x + 1e-7) - y) / 1e-7;
      if (!isFinite(y) || !isFinite(d) || d === 0) return err(ERR.NUM);
      const nx = x - y / d;
      if (Math.abs(nx - x) < 1e-12) return nx;
      x = nx;
    }
    return err(ERR.NUM);
  };
  // аргументы с умолчаниями: недостающие в формуле берут значение из defs
  const fnum = (args, defs) => Array.from({ length: Math.max(args.length, defs.length) }, (_, i) => num(args[i], defs[i]));
  const fin = {
    PMT: (a) => { const [r, n, pv, fv, t] = fnum(a, [undefined, undefined, undefined, 0, 0]); return firstErr(r, n, pv, fv, t) || pmt(r, n, pv, fv, t ? 1 : 0); },
    FV: (a) => { const [r, n, p, pv, t] = fnum(a, [undefined, undefined, undefined, 0, 0]); return firstErr(r, n, p, pv, t) || fvf(r, n, p, pv, t ? 1 : 0); },
    PV: (a) => {
      const [r, n, p, fv, t0] = fnum(a, [undefined, undefined, undefined, 0, 0]);
      if (firstErr(r, n, p, fv, t0)) return firstErr(r, n, p, fv, t0);
      const t = t0 ? 1 : 0;
      return r === 0 ? -(fv + p * n) : -(fv + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n;
    },
    NPER: (a) => {
      const [r, p, pv, fv, t0] = fnum(a, [undefined, undefined, undefined, 0, 0]);
      if (firstErr(r, p, pv, fv, t0)) return firstErr(r, p, pv, fv, t0);
      const t = t0 ? 1 : 0;
      if (r === 0) return p === 0 ? err(ERR.NUM) : -(pv + fv) / p;
      const v = (p * (1 + r * t) - fv * r) / (p * (1 + r * t) + pv * r);
      return v <= 0 ? err(ERR.NUM) : Math.log(v) / Math.log(1 + r);
    },
    IPMT: (a) => {
      const [r, per, n, pv, fv, t0] = fnum(a, [undefined, undefined, undefined, undefined, 0, 0]);
      if (firstErr(r, per, n, pv, fv, t0)) return firstErr(r, per, n, pv, fv, t0);
      const t = t0 ? 1 : 0;
      if (per < 1 || per > n) return err(ERR.NUM);
      const p = pmt(r, n, pv, fv, t);
      if (per === 1 && t === 1) return 0;
      const ip = fvf(r, per - 1, p, pv, t) * r;
      return t === 1 ? ip / (1 + r) : ip;
    },
    PPMT: (a) => {
      const [r, per, n, pv, fv, t0] = fnum(a, [undefined, undefined, undefined, undefined, 0, 0]);
      if (firstErr(r, per, n, pv, fv, t0)) return firstErr(r, per, n, pv, fv, t0);
      const ip = fin.IPMT(a);
      return isErr(ip) ? ip : pmt(r, n, pv, fv, t0 ? 1 : 0) - ip;
    },
    RATE: (a) => {
      const [n, p, pv, fv, t0, g] = fnum(a, [undefined, undefined, undefined, 0, 0, 0.1]);
      if (firstErr(n, p, pv, fv, t0, g)) return firstErr(n, p, pv, fv, t0, g);
      const t = t0 ? 1 : 0;
      return solve((r) => (Math.abs(r) < 1e-12 ? pv + p * n + fv : pv * (1 + r) ** n + (p * (1 + r * t) * ((1 + r) ** n - 1)) / r + fv), g);
    },
    NPV: ([rate, ...vals]) => {
      const r = num(rate);
      if (isErr(r)) return r;
      let s = 0; let i = 1;
      for (const a of vals) {
        for (const v of flat(a)) {
          if (isErr(v)) return v;
          if (typeof v === 'number') s += v / (1 + r) ** i++;
          else if (!Array.isArray(a) && v != null) { const x = toNum(v); if (isErr(x)) return x; s += x / (1 + r) ** i++; }
        }
      }
      return s;
    },
    IRR: ([vals, guess]) => {
      const v = flat(vals).filter((x) => typeof x === 'number');
      const g = num(guess, 0.1);
      if (isErr(g)) return g;
      if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) return err(ERR.NUM);
      return solve((r) => v.reduce((s, x, i) => s + x / (1 + r) ** i, 0), g);
    },
    XNPV: ([rate, vals, dates]) => {
      const r = num(rate); const v = flat(vals); const d = flat(dates);
      if (isErr(r)) return r;
      if (v.length !== d.length) return err(ERR.NUM);
      return v.reduce((s, x, i) => s + x / (1 + r) ** ((d[i] - d[0]) / 365), 0);
    },
    XIRR: ([vals, dates, guess]) => {
      const v = flat(vals); const d = flat(dates); const g = num(guess, 0.1);
      if (isErr(g)) return g;
      if (v.length !== d.length) return err(ERR.NUM);
      return solve((r) => v.reduce((s, x, i) => s + x / (1 + r) ** ((d[i] - d[0]) / 365), 0), g);
    },
    SLN: (a) => { const [c, s, l] = fnum(a, []); return firstErr(c, s, l) || (l === 0 ? err(ERR.DIV0) : (c - s) / l); },
    SYD: (a) => { const [c, s, l, p] = fnum(a, []); return firstErr(c, s, l, p) || ((c - s) * (l - p + 1) * 2) / (l * (l + 1)); },
    DB: (a) => {
      const [c, s, l, p, m0] = fnum(a, [undefined, undefined, undefined, undefined, 12]);
      if (firstErr(c, s, l, p, m0)) return firstErr(c, s, l, p, m0);
      if (c < 0 || s < 0 || l <= 0 || p <= 0 || m0 < 1 || m0 > 12 || p > l + 1) return err(ERR.NUM);
      const rate = Math.round((1 - (s / c) ** (1 / l)) * 1000) / 1000;
      let total = 0; let dep = 0;
      for (let i = 1; i <= p; i++) {
        if (i === 1) dep = (c * rate * m0) / 12;
        else if (i === l + 1) dep = ((c - total) * rate * (12 - m0)) / 12;
        else dep = (c - total) * rate;
        total += dep;
      }
      return dep;
    },
    DDB: (a) => {
      const [c, s, l, p, f] = fnum(a, [undefined, undefined, undefined, undefined, 2]);
      if (firstErr(c, s, l, p, f)) return firstErr(c, s, l, p, f);
      if (c < 0 || s < 0 || l <= 0 || p <= 0 || p > l) return err(ERR.NUM);
      let total = 0; let dep = 0;
      for (let i = 1; i <= p; i++) { dep = Math.max(0, Math.min((c - total) * (f / l), c - s - total)); total += dep; }
      return dep;
    },
    EFFECT: ([r, n]) => { const x = num(r); const k = int(n); if (firstErr(x, k)) return firstErr(x, k); return x <= 0 || k < 1 ? err(ERR.NUM) : (1 + x / k) ** k - 1; },
    NOMINAL: ([r, n]) => { const x = num(r); const k = int(n); if (firstErr(x, k)) return firstErr(x, k); return x <= 0 || k < 1 ? err(ERR.NUM) : k * ((1 + x) ** (1 / k) - 1); },
  };

  return { ...math, ...logic, ...cond, ...subtotal, ...lookup, ...text, ...date, ...stat, ...fin };
}

// «Ленивые» функции: аргументы вычисляются по необходимости (ev(узел, ctx) — вычислитель движка)
export function lazyFunctions(h) {
  const { ERR, err, isErr, toBool, compare } = h;
  return {
    IFS: (nodes, ctx, ev) => {
      for (let i = 0; i + 1 < nodes.length; i += 2) {
        const c = ev(nodes[i], ctx);
        const v = Array.isArray(c) ? c[0][0] : c;
        if (isErr(v)) return v;
        const b = toBool(v);
        if (isErr(b)) return b;
        if (b) return ev(nodes[i + 1], ctx) ?? 0;
      }
      return err(ERR.NA);
    },
    SWITCH: (nodes, ctx, ev) => {
      const x = ev(nodes[0], ctx);
      if (isErr(x)) return x;
      let i = 1;
      for (; i + 1 < nodes.length; i += 2) {
        const v = ev(nodes[i], ctx);
        if (isErr(v)) return v;
        if (compare(x, v) === 0 && typeof x === typeof v) return ev(nodes[i + 1], ctx) ?? 0;
      }
      return i < nodes.length ? ev(nodes[i], ctx) ?? 0 : err(ERR.NA);
    },
    IFNA: (nodes, ctx, ev) => {
      const v = ev(nodes[0], ctx);
      return isErr(v) && v.error === ERR.NA ? ev(nodes[1], ctx) : v;
    },
    // LET(имя1; значение1; …; выражение): имена видны дальше по списку и в выражении
    LET: (nodes, ctx, ev) => {
      if (nodes.length < 3 || nodes.length % 2 === 0) return err(ERR.VALUE);
      const vars = new Map();
      const scope = Object.create(ctx);
      scope.name = (n) => (vars.has(n.toLowerCase()) ? vars.get(n.toLowerCase()) : ctx.name ? ctx.name(n) : err(ERR.NAME));
      for (let i = 0; i + 1 < nodes.length; i += 2) {
        if (nodes[i].t !== 'name') return err(ERR.VALUE);
        vars.set(nodes[i].v.toLowerCase(), ev(nodes[i + 1], scope));
      }
      return ev(nodes[nodes.length - 1], scope);
    },
  };
}
