// Библиотека функций, пакет A (этап 12.5): математика и матрицы, системы счисления и биты, функции ошибок,
// базы данных (БДСУММ…), описательная статистика и регрессия (ЛИНЕЙН…), текст (байтовые, БАТТЕКСТ…), прочее и
// функции внешних сервисов (ошибка как у Excel без подключения). Подключается в formula.js, как fnlib.js.
import { DOMParser } from '@xmldom/xmldom';

export function functions2(h) {
  const { ERR, err, isErr, toNum, toText, toBool, flat, nums, firstErr, criterion, toArr, FN } = h;
  const num = (x, def) => (x == null && def !== undefined ? def : toNum(x));
  const int = (x, def) => { const n = num(x, def); return isErr(n) ? n : Math.trunc(n); };
  const bool = (x, def) => { if (x == null) return def; const b = toBool(x); return isErr(b) ? b : b; };
  const mean = (n) => n.reduce((s, x) => s + x, 0) / n.length;
  const all = (args) => args.flatMap(flat);
  const unary = (f, ok = () => true) => ([x]) => { const n = num(x); if (isErr(n)) return n; return ok(n) ? f(n) : err(ERR.NUM); };

  // ---------- математика ----------
  const fact = (n) => { let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; };
  const math = {
    ACOSH: unary(Math.acosh, (n) => n >= 1), ASINH: unary(Math.asinh), ATANH: unary(Math.atanh, (n) => Math.abs(n) < 1),
    ACOT: unary((n) => Math.PI / 2 - Math.atan(n)), ACOTH: unary((n) => 0.5 * Math.log((n + 1) / (n - 1)), (n) => Math.abs(n) > 1),
    COT: ([x]) => { const n = num(x); if (isErr(n)) return n; return n === 0 ? err(ERR.DIV0) : 1 / Math.tan(n); },
    COTH: ([x]) => { const n = num(x); if (isErr(n)) return n; return n === 0 ? err(ERR.DIV0) : 1 / Math.tanh(n); },
    CSC: ([x]) => { const n = num(x); if (isErr(n)) return n; return n === 0 ? err(ERR.DIV0) : 1 / Math.sin(n); },
    CSCH: ([x]) => { const n = num(x); if (isErr(n)) return n; return n === 0 ? err(ERR.DIV0) : 1 / Math.sinh(n); },
    SEC: unary((n) => 1 / Math.cos(n)), SECH: unary((n) => 1 / Math.cosh(n)),
    COMBINA: ([a, b]) => {
      const n = int(a); const k = int(b);
      if (firstErr(n, k)) return firstErr(n, k);
      if (n < 0 || k < 0 || (n === 0 && k > 0)) return err(ERR.NUM);
      return FN.COMBIN([n + k - 1, k]);
    },
    PERMUTATIONA: ([a, b]) => { const n = int(a); const k = int(b); if (firstErr(n, k)) return firstErr(n, k); return n < 0 || k < 0 ? err(ERR.NUM) : n ** k; },
    FACTDOUBLE: ([x]) => {
      const n = int(x);
      if (isErr(n)) return n;
      if (n < -1) return err(ERR.NUM);
      let r = 1; for (let i = n; i > 1; i -= 2) r *= i;
      return r;
    },
    MULTINOMIAL: (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      if (n.some((x) => x < 0)) return err(ERR.NUM);
      const t = n.map(Math.trunc);
      let r = fact(t.reduce((s, x) => s + x, 0));
      for (const x of t) r /= fact(x);
      return Math.round(r);
    },
    SERIESSUM: ([x, n0, m, coefs]) => {
      const v = num(x); const n = num(n0); const step = num(m);
      if (firstErr(v, n, step)) return firstErr(v, n, step);
      const c = flat(coefs);
      let s = 0;
      for (let i = 0; i < c.length; i++) { const k = toNum(c[i]); if (isErr(k)) return err(ERR.VALUE); s += k * v ** (n + i * step); }
      return s;
    },
    MUNIT: ([x]) => { const n = int(x); if (isErr(n)) return n; if (n < 1) return err(ERR.VALUE); return Array.from({ length: n }, (_, i) => Array.from({ length: n }, (__, j) => (i === j ? 1 : 0))); },
    MMULT: ([a, b]) => {
      const A = matrix(a); const B = matrix(b);
      if (isErr(A)) return A;
      if (isErr(B)) return B;
      if (A[0].length !== B.length) return err(ERR.VALUE);
      return A.map((row) => B[0].map((_, j) => row.reduce((s, v, k) => s + v * B[k][j], 0)));
    },
    MDETERM: ([a]) => {
      const A = matrix(a);
      if (isErr(A)) return A;
      if (A.length !== A[0].length) return err(ERR.VALUE);
      return lu(A).det;
    },
    MINVERSE: ([a]) => {
      const A = matrix(a);
      if (isErr(A)) return A;
      if (A.length !== A[0].length) return err(ERR.VALUE);
      const inv = invert(A);
      return inv || err(ERR.NUM);
    },
  };
  // матрица чисел из диапазона/массива (пустые и текст — #ЗНАЧ!)
  function matrix(x) {
    if (isErr(x)) return x;
    const a = toArr(x);
    for (const row of a) for (const v of row) { if (isErr(v)) return v; if (typeof v !== 'number') return err(ERR.VALUE); }
    return a;
  }
  function lu(A) {
    const n = A.length; const m = A.map((r) => [...r]); let det = 1;
    for (let i = 0; i < n; i++) {
      let p = i;
      for (let r = i + 1; r < n; r++) if (Math.abs(m[r][i]) > Math.abs(m[p][i])) p = r;
      if (m[p][i] === 0) return { det: 0 };
      if (p !== i) { [m[p], m[i]] = [m[i], m[p]]; det = -det; }
      det *= m[i][i];
      for (let r = i + 1; r < n; r++) { const f = m[r][i] / m[i][i]; for (let c = i; c < n; c++) m[r][c] -= f * m[i][c]; }
    }
    return { det };
  }
  function invert(A) {
    const n = A.length;
    const m = A.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
    for (let i = 0; i < n; i++) {
      let p = i;
      for (let r = i + 1; r < n; r++) if (Math.abs(m[r][i]) > Math.abs(m[p][i])) p = r;
      if (Math.abs(m[p][i]) < 1e-300) return null;
      [m[p], m[i]] = [m[i], m[p]];
      const d = m[i][i];
      for (let c = 0; c < 2 * n; c++) m[i][c] /= d;
      for (let r = 0; r < n; r++) if (r !== i) { const f = m[r][i]; for (let c = 0; c < 2 * n; c++) m[r][c] -= f * m[i][c]; }
    }
    return m.map((r) => r.slice(n));
  }
  // решение нормальных уравнений (МНК) для ЛИНЕЙН: X — строки наблюдений, y — значения
  function ols(X, y, withConst) {
    const rows = X.map((r) => (withConst ? [...r, 1] : [...r]));
    const k = rows[0].length;
    const XtX = Array.from({ length: k }, (_, i) => Array.from({ length: k }, (__, j) => rows.reduce((s, r) => s + r[i] * r[j], 0)));
    const Xty = Array.from({ length: k }, (_, i) => rows.reduce((s, r, t) => s + r[i] * y[t], 0));
    const inv = invert(XtX);
    if (!inv) return null;
    const beta = inv.map((row) => row.reduce((s, v, j) => s + v * Xty[j], 0));
    return { beta, inv, rows };
  }

  // ---------- системы счисления, биты, функции ошибок ----------
  const BASES = { BIN: [2, 10], OCT: [8, 10], DEC: [10, 0], HEX: [16, 10] };
  // из системы счисления (10 знаков; старший бит — знак, дополнительный код)
  const fromBase = (text, base) => {
    const t = toText(text).trim().toUpperCase();
    if (t.length > 10 || !t || [...t].some((c) => !(parseInt(c, base) < base))) return err(ERR.NUM);
    let v = parseInt(t, base);
    const bits = { 2: 10, 8: 30, 16: 40 }[base];
    if (t.length === 10 && v >= 2 ** (bits - 1)) v -= 2 ** bits;
    return v;
  };
  const toBase = (v, base, places) => {
    const bits = { 2: 10, 8: 30, 16: 40 }[base];
    const lim = 2 ** (bits - 1);
    if (v < -lim || v >= lim) return err(ERR.NUM);
    if (v < 0) return (v + 2 ** bits).toString(base).toUpperCase();
    const s = v.toString(base).toUpperCase();
    if (places == null) return s;
    const p = Math.trunc(toNum(places));
    if (isErr(p) || p < 1 || p > 10 || s.length > p) return err(ERR.NUM);
    return s.padStart(p, '0');
  };
  const conv = {};
  for (const [from, [fb]] of Object.entries(BASES)) {
    for (const [to, [tb]] of Object.entries(BASES)) {
      if (from === to) continue;
      conv[`${from}2${to}`] = ([x, places]) => {
        if (isErr(x)) return x;
        let v;
        if (fb === 10) { const n = num(x); if (isErr(n)) return n; v = Math.trunc(n); } else { v = fromBase(x, fb); if (isErr(v)) return v; }
        return tb === 10 ? v : toBase(v, tb, places);
      };
    }
  }
  const bitArg = (x) => { const n = num(x); if (isErr(n)) return n; return n < 0 || n >= 2 ** 48 || n !== Math.trunc(n) ? err(ERR.NUM) : BigInt(n); };
  const bitOp = (f) => ([a, b]) => { const x = bitArg(a); const y = bitArg(b); return firstErr(x, y) || Number(f(x, y)); };
  const shift = (left) => ([a, s]) => {
    const x = bitArg(a); const k = int(s);
    if (firstErr(x, k)) return firstErr(x, k);
    if (Math.abs(k) > 53) return err(ERR.NUM);
    const r = (left ? k : -k) >= 0 ? x << BigInt(left ? k : -k) : x >> BigInt(left ? -k : k);
    return r >= 2n ** 48n ? err(ERR.NUM) : Number(r);
  };
  function erf(x) {
    if (Math.abs(x) < 2.5) {
      let s = 0; let t = x; const q = x * x;
      for (let n = 0; n < 200; n++) { const term = t / (2 * n + 1); s += term; if (Math.abs(term) < 1e-17 * Math.abs(s)) break; t *= -q / (n + 1); }
      return (2 / Math.sqrt(Math.PI)) * s;
    }
    return x > 0 ? 1 - erfcCF(x) : erfcCF(-x) - 1;
  }
  // erfc(x), x > 0 — цепная дробь (метод Ленца)
  function erfcCF(x) {
    let f = x; let C = x; let D = 0;
    for (let i = 1; i < 300; i++) {
      const a = i / 2;
      D = x + a * D; D = 1 / D;
      C = x + a / C;
      const d = C * D; f *= d;
      if (Math.abs(d - 1) < 1e-16) break;
    }
    return Math.exp(-x * x) / Math.sqrt(Math.PI) / f;
  }
  const erfc = (x) => (x > 2.5 ? erfcCF(x) : 1 - erf(x));
  const eng = {
    ...conv,
    BITAND: bitOp((x, y) => x & y), BITOR: bitOp((x, y) => x | y), BITXOR: bitOp((x, y) => x ^ y),
    BITLSHIFT: shift(true), BITRSHIFT: shift(false),
    DELTA: ([a, b]) => { const x = num(a); const y = num(b, 0); return firstErr(x, y) || (x === y ? 1 : 0); },
    GESTEP: ([a, b]) => { const x = num(a); const y = num(b, 0); return firstErr(x, y) || (x >= y ? 1 : 0); },
    ERF: ([a, b]) => { const x = num(a); if (isErr(x)) return x; if (b == null) return erf(x); const y = num(b); return isErr(y) ? y : erf(y) - erf(x); },
    'ERF.PRECISE': ([a]) => { const x = num(a); return isErr(x) ? x : erf(x); },
    ERFC: ([a]) => { const x = num(a); return isErr(x) ? x : erfc(x); },
    'ERFC.PRECISE': ([a]) => { const x = num(a); return isErr(x) ? x : erfc(x); },
  };
  h.erf = erf; h.erfc = erfc; // для распределений (пакет B)

  // ---------- базы данных: БДСУММ(база; поле; критерии) ----------
  function dbSelect(database, field, crit) {
    const db = toArr(database); const cr = toArr(crit);
    if (db.length < 1 || cr.length < 1) return err(ERR.VALUE);
    const head = db[0].map((v) => toText(v).toLowerCase());
    let col = -1;
    if (field == null) col = -1;
    else if (typeof field === 'number') col = Math.trunc(field) - 1;
    else col = head.indexOf(toText(field).toLowerCase());
    if (field != null && (col < 0 || col >= head.length)) return err(ERR.VALUE);
    const ch = cr[0].map((v) => toText(v).toLowerCase());
    const conds = cr.slice(1).map((row) => row.map((v, j) => ({ col: head.indexOf(ch[j]), test: v == null || v === '' ? null : criterion(v) })));
    const match = (rec) => !conds.length || conds.some((row) => row.every((c) => !c.test || (c.col >= 0 && c.test(rec[c.col]))));
    const recs = db.slice(1).filter(match);
    return { values: col >= 0 ? recs.map((r) => r[col]) : recs.map(() => 1), col };
  }
  const dbNums = (sel) => sel.values.filter((v) => typeof v === 'number');
  const dbf = (f) => ([d, fld, c]) => { const sel = dbSelect(d, fld, c); return isErr(sel) ? sel : f(sel); };
  const varOf = (n, pop) => { if (n.length < (pop ? 1 : 2)) return err(ERR.DIV0); const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - (pop ? 0 : 1)); };
  const db = {
    DSUM: dbf((s) => dbNums(s).reduce((a, b) => a + b, 0)),
    DAVERAGE: dbf((s) => { const n = dbNums(s); return n.length ? mean(n) : err(ERR.DIV0); }),
    DCOUNT: dbf((s) => (s.col < 0 ? s.values.length : dbNums(s).length)),
    DCOUNTA: dbf((s) => s.values.filter((v) => v != null && v !== '').length),
    DMAX: dbf((s) => { const n = dbNums(s); return n.length ? Math.max(...n) : 0; }),
    DMIN: dbf((s) => { const n = dbNums(s); return n.length ? Math.min(...n) : 0; }),
    DPRODUCT: dbf((s) => dbNums(s).reduce((a, b) => a * b, 1)),
    DGET: dbf((s) => (s.values.length === 0 ? err(ERR.VALUE) : s.values.length > 1 ? err(ERR.NUM) : s.values[0] ?? 0)),
    DSTDEV: dbf((s) => { const v = varOf(dbNums(s), false); return isErr(v) ? v : Math.sqrt(v); }),
    DSTDEVP: dbf((s) => { const v = varOf(dbNums(s), true); return isErr(v) ? v : Math.sqrt(v); }),
    DVAR: dbf((s) => varOf(dbNums(s), false)),
    DVARP: dbf((s) => varOf(dbNums(s), true)),
  };

  // ---------- статистика: описательная, регрессия ----------
  const numsA = (a) => {
    const out = [];
    for (const x of a) {
      for (const v of flat(x)) {
        if (isErr(v)) return v;
        if (v == null) continue;
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') out.push(v ? 1 : 0);
        else if (Array.isArray(x)) out.push(0); // текст в диапазоне — 0
        else { const n = toNum(v); if (isErr(n)) return n; out.push(n); }
      }
    }
    return out;
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
  const cov = (a, b, pop) => {
    const p = pairs(a, b);
    if (isErr(p)) return p;
    const [x, y] = p;
    if (x.length < (pop ? 1 : 2)) return err(ERR.DIV0);
    const mx = mean(x); const my = mean(y);
    return x.reduce((s, v, i) => s + (v - mx) * (y[i] - my), 0) / (x.length - (pop ? 0 : 1));
  };
  // ЛИНЕЙН: известные_y; [известные_x] (столбцы/строки — переменные); конст; статистика
  function linest(ys, xs, cnst, stats) {
    const y = flat(ys);
    if (y.some((v) => typeof v !== 'number')) return err(ERR.VALUE);
    let X;
    if (xs == null) X = y.map((_, i) => [i + 1]);
    else {
      const a = toArr(xs);
      if (a.flat().some((v) => typeof v !== 'number')) return err(ERR.VALUE);
      const yVertical = toArr(ys).length === y.length;
      if (a.length === y.length && (yVertical || a[0].length !== y.length)) X = a.map((r) => [...r]); // переменные — столбцы
      else if (a[0].length === y.length) X = y.map((_, i) => a.map((row) => row[i])); // переменные — строки
      else return err(ERR.REF);
    }
    const withConst = cnst == null ? true : bool(cnst, true);
    const n = y.length; const k = X[0].length;
    const fit = ols(X, y, withConst);
    if (!fit) return err(ERR.NUM);
    const { beta, inv, rows } = fit;
    // порядок Excel: коэффициенты m_k … m_1, b
    const coefs = beta.slice(0, k).reverse().concat(withConst ? [beta[k]] : [0]);
    if (!bool(stats, false)) return [coefs];
    const yhat = rows.map((r) => r.reduce((s, v, i) => s + v * beta[i], 0));
    const my = withConst ? mean(y) : 0;
    const ssr = yhat.reduce((s, v) => s + (v - my) ** 2, 0);
    const sse = y.reduce((s, v, i) => s + (v - yhat[i]) ** 2, 0);
    const df = n - k - (withConst ? 1 : 0);
    const sey = Math.sqrt(sse / df);
    const se = beta.map((_, i) => Math.sqrt(inv[i][i] * sse / df));
    const seRow = se.slice(0, k).reverse().concat(withConst ? [se[k]] : [err(ERR.NA)]);
    const width = k + 1;
    const na = () => err(ERR.NA);
    const pad = (row) => [...row, ...Array.from({ length: width - row.length }, na)];
    return [coefs, seRow, pad([ssr / (ssr + sse), sey]), pad([(ssr / k) / (sse / df), df]), pad([ssr, sse])];
  }
  const stat = {
    COVAR: ([a, b]) => cov(a, b, true), 'COVARIANCE.P': ([a, b]) => cov(a, b, true), 'COVARIANCE.S': ([a, b]) => cov(a, b, false),
    PEARSON: ([a, b]) => FN.CORREL([a, b]),
    RSQ: ([y, x]) => { const r = FN.CORREL([y, x]); return isErr(r) ? r : r * r; },
    STEYX: ([ys, xs]) => {
      const p = pairs(ys, xs);
      if (isErr(p)) return p;
      const [y, x] = p;
      if (y.length < 3) return err(ERR.DIV0);
      const mx = mean(x); const my = mean(y);
      let sxy = 0; let sxx = 0; let syy = 0;
      for (let i = 0; i < x.length; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; }
      return Math.sqrt((syy - (sxy * sxy) / sxx) / (y.length - 2));
    },
    KURT: (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      const N = n.length;
      if (N < 4) return err(ERR.DIV0);
      const m = mean(n); const s = Math.sqrt(n.reduce((t, x) => t + (x - m) ** 2, 0) / (N - 1));
      if (s === 0) return err(ERR.DIV0);
      const s4 = n.reduce((t, x) => t + ((x - m) / s) ** 4, 0);
      return (N * (N + 1)) / ((N - 1) * (N - 2) * (N - 3)) * s4 - (3 * (N - 1) ** 2) / ((N - 2) * (N - 3));
    },
    SKEW: (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      const N = n.length;
      if (N < 3) return err(ERR.DIV0);
      const m = mean(n); const s = Math.sqrt(n.reduce((t, x) => t + (x - m) ** 2, 0) / (N - 1));
      if (s === 0) return err(ERR.DIV0);
      return (N / ((N - 1) * (N - 2))) * n.reduce((t, x) => t + ((x - m) / s) ** 3, 0);
    },
    'SKEW.P': (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      if (!n.length) return err(ERR.DIV0);
      const m = mean(n); const s = Math.sqrt(n.reduce((t, x) => t + (x - m) ** 2, 0) / n.length);
      return s === 0 ? err(ERR.DIV0) : n.reduce((t, x) => t + ((x - m) / s) ** 3, 0) / n.length;
    },
    STANDARDIZE: ([x, m, s]) => { const a = num(x); const b = num(m); const c = num(s); return firstErr(a, b, c) || (c <= 0 ? err(ERR.NUM) : (a - b) / c); },
    STDEVA: (a) => { const n = numsA(a); if (isErr(n)) return n; const v = varOf(n, false); return isErr(v) ? v : Math.sqrt(v); },
    STDEVPA: (a) => { const n = numsA(a); if (isErr(n)) return n; const v = varOf(n, true); return isErr(v) ? v : Math.sqrt(v); },
    VARA: (a) => { const n = numsA(a); return isErr(n) ? n : varOf(n, false); },
    VARPA: (a) => { const n = numsA(a); return isErr(n) ? n : varOf(n, true); },
    TRIMMEAN: ([a, p]) => {
      const n = nums([a]); const pc = num(p);
      if (firstErr(n, pc)) return firstErr(n, pc);
      if (pc < 0 || pc >= 1 || !n.length) return err(ERR.NUM);
      const k = Math.floor((n.length * pc) / 2);
      const s = [...n].sort((x, y) => x - y).slice(k, n.length - k);
      return mean(s);
    },
    FISHER: unary((x) => 0.5 * Math.log((1 + x) / (1 - x)), (x) => Math.abs(x) < 1),
    FISHERINV: unary((y) => (Math.exp(2 * y) - 1) / (Math.exp(2 * y) + 1)),
    GAUSS: ([z]) => { const v = num(z); return isErr(v) ? v : FN['NORM.S.DIST']([v, true]) - 0.5; },
    PHI: ([z]) => { const v = num(z); return isErr(v) ? v : Math.exp(-0.5 * v * v) / Math.sqrt(2 * Math.PI); },
    PROB: ([xr, pr, lo, hi]) => {
      const x = flat(xr); const p = flat(pr);
      if (x.length !== p.length) return err(ERR.NA);
      const a = num(lo); const b = hi == null ? a : num(hi);
      if (firstErr(a, b)) return firstErr(a, b);
      let total = 0; let s = 0;
      for (let i = 0; i < x.length; i++) {
        if (typeof p[i] !== 'number') continue;
        if (p[i] < 0 || p[i] > 1) return err(ERR.NUM);
        total += p[i];
        if (typeof x[i] === 'number' && x[i] >= a && x[i] <= b) s += p[i];
      }
      return Math.abs(total - 1) > 1e-7 ? err(ERR.NUM) : s;
    },
    'MODE.MULT': (a) => {
      const n = nums(a);
      if (isErr(n)) return n;
      const c = new Map();
      for (const x of n) c.set(x, (c.get(x) || 0) + 1);
      const best = Math.max(...c.values());
      if (best < 2) return err(ERR.NA);
      const out = []; const seen = new Set();
      for (const x of n) if (c.get(x) === best && !seen.has(x)) { seen.add(x); out.push([x]); }
      return out;
    },
    'PERCENTRANK.EXC': ([a, x, sig]) => {
      const n = nums([a]); const v = num(x); const s = int(sig, 3);
      if (firstErr(n, v, s)) return firstErr(n, v, s);
      const sorted = [...n].sort((p, q) => p - q);
      if (!sorted.length || v < sorted[0] || v > sorted[sorted.length - 1] || s < 1) return err(sorted.length ? ERR.NA : ERR.NUM);
      let pos;
      const i = sorted.indexOf(v);
      if (i >= 0) pos = i + 1;
      else { const j = sorted.findIndex((y) => y > v) - 1; pos = j + 1 + (v - sorted[j]) / (sorted[j + 1] - sorted[j]); }
      const p = 10 ** s;
      return Math.floor(Math.round((pos / (sorted.length + 1)) * p * 1e6) / 1e6) / p;
    },
    'CONFIDENCE.NORM': ([a, s, n]) => {
      const al = num(a); const sd = num(s); const k = int(n);
      if (firstErr(al, sd, k)) return firstErr(al, sd, k);
      if (al <= 0 || al >= 1 || sd <= 0 || k < 1) return err(ERR.NUM);
      return (FN['NORM.S.INV']([1 - al / 2]) * sd) / Math.sqrt(k);
    },
    LINEST: ([y, x, c, s]) => linest(y, x, c, s),
    LOGEST: ([y, x, c, s]) => {
      const ys = toArr(y).map((row) => row.map((v) => (typeof v === 'number' && v > 0 ? Math.log(v) : NaN)));
      if (ys.flat().some((v) => Number.isNaN(v))) return err(ERR.NUM);
      const r = linest(ys, x, c, s);
      if (isErr(r)) return r;
      r[0] = r[0].map((v) => Math.exp(v));
      return r;
    },
    GROWTH: ([y, x, nx, c]) => {
      const ys = toArr(y).map((row) => row.map((v) => (typeof v === 'number' && v > 0 ? Math.log(v) : NaN)));
      if (ys.flat().some((v) => Number.isNaN(v))) return err(ERR.NUM);
      const r = linest(ys, x, c, false);
      if (isErr(r)) return r;
      const coef = r[0];
      const k = coef.length - 1;
      const pts = nx != null ? toArr(nx) : x != null ? toArr(x) : flat(y).map((_, i) => [i + 1]);
      // точки: строки (или столбцы, если переменных — строки) с k значениями
      const rowsOf = pts[0].length === k ? pts : pts.length === k ? pts[0].map((_, j) => pts.map((row) => row[j])) : pts.map((row) => row);
      const val = (v) => Math.exp(coef[k] + v.reduce((s0, xv, i) => s0 + xv * coef[k - 1 - i], 0));
      if (k === 1) return pts.map((row) => row.map((v) => val([v])));
      return rowsOf.map((v) => [val(v)]);
    },
  };
  stat.CONFIDENCE = stat['CONFIDENCE.NORM'];

  // ---------- текст ----------
  const text = {
    // в русской Windows (однобайтовая кодировка) байтовые варианты совпадают с обычными
    LEFTB: (a) => FN.LEFT(a), RIGHTB: (a) => FN.RIGHT(a), MIDB: (a) => FN.MID(a), LENB: (a) => FN.LEN(a),
    FINDB: (a) => FN.FIND(a), SEARCHB: (a) => FN.SEARCH(a), REPLACEB: (a) => FN.REPLACE(a),
    // полуширинные/полноширинные знаки — только в восточноазиатских версиях; в русской текст не меняется (как Excel)
    ASC: ([x]) => (isErr(x) ? x : toText(x)),
    DBCS: ([x]) => (isErr(x) ? x : toText(x)),
    PHONETIC: ([x]) => (isErr(x) ? x : toText(x)),
    USDOLLAR: (a) => FN.DOLLAR(a), // русский Excel пишет в валюте системы — как РУБЛЬ
    VALUETOTEXT: ([v, fmt]) => {
      const strict = int(fmt, 0);
      if (isErr(strict)) return strict;
      if (isErr(v)) return h.errText(v.error);
      if (typeof v === 'string') return strict === 1 ? `"${v.replace(/"/g, '""')}"` : v;
      return toText(v);
    },
    ARRAYTOTEXT: ([arr, fmt]) => {
      const strict = int(fmt, 0);
      if (isErr(strict)) return strict;
      const a = toArr(arr);
      const one = (v) => text.VALUETOTEXT([v ?? '', strict]);
      if (strict !== 1) return a.flat().map(one).join('; ');
      return `{${a.map((row) => row.map(one).join(';')).join(':')}}`;
    },
    BAHTTEXT: ([x]) => { const n = num(x); return isErr(n) ? n : bahtText(n); },
    ENCODEURL: ([x]) => (isErr(x) ? x : encodeURIComponent(toText(x)).replace(/[!'()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`)),
    FILTERXML: ([xml, path]) => {
      if (firstErr(xml, path)) return firstErr(xml, path);
      let doc;
      try { doc = new DOMParser({ onError: () => { throw new Error('xml'); } }).parseFromString(toText(xml), 'text/xml'); } catch { return err(ERR.VALUE); }
      const res = xpath(doc, toText(path));
      if (!res) return err(ERR.VALUE);
      if (!res.length) return err(ERR.VALUE);
      const vals = res.map((t) => { const n = h.textToNum(t); return n != null && /^\s*-?[\d\s,.]+\s*$/.test(t) ? n : t; });
      return vals.length === 1 ? vals[0] : vals.map((v) => [v]);
    },
  };
  // XPath для ФИЛЬТР.XML — частый набор: /a/b, //b, //b/@attr, //b[@x='1'], //b[2], text()
  function xpath(doc, path) {
    const steps = []; const re = /(\/\/?)([^/[]+)(\[[^\]]+\])?/g;
    let m; let pos = 0;
    while ((m = re.exec(path))) { if (m.index !== pos) return null; steps.push({ deep: m[1] === '//', name: m[2], pred: m[3] ? m[3].slice(1, -1) : null }); pos = re.lastIndex; }
    if (pos !== path.length || !steps.length) return null;
    let nodes = [doc];
    for (const st of steps) {
      const out = [];
      for (const n of nodes) {
        if (st.name.startsWith('@')) { if (n.getAttribute && n.hasAttribute(st.name.slice(1))) out.push({ text: n.getAttribute(st.name.slice(1)) }); continue; }
        if (st.name === 'text()') { out.push({ text: n.textContent }); continue; }
        const cand = [];
        const walk = (p, deep) => { for (let c = p.firstChild; c; c = c.nextSibling) { if (c.nodeType === 1) { if (st.name === '*' || c.nodeName === st.name) cand.push(c); if (deep) walk(c, true); } } };
        walk(n, st.deep);
        let sel = cand;
        if (st.pred) {
          const idx = /^\d+$/.test(st.pred) ? +st.pred : null;
          const at = /^@([\w:-]+)\s*=\s*['"]([^'"]*)['"]$/.exec(st.pred);
          if (idx) sel = cand[idx - 1] ? [cand[idx - 1]] : [];
          else if (at) sel = cand.filter((c) => c.getAttribute(at[1]) === at[2]);
          else return null;
        }
        out.push(...sel);
      }
      nodes = out;
    }
    return nodes.map((n) => (n.text !== undefined ? n.text : n.textContent));
  }
  // БАТТЕКСТ — число тайскими словами с «бат/сатанг» (как Excel)
  function bahtText(n) {
    const D = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const P = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    const words = (v) => {
      if (v === 0) return '';
      let s = '';
      if (v >= 1e6) { s = words(Math.floor(v / 1e6)) + 'ล้าน'; v %= 1e6; }
      const ds = String(v).padStart(6, '0').split('').map(Number);
      ds.forEach((d, i) => {
        const p = 5 - i;
        if (!d) return;
        if (p === 1 && d === 1) s += 'สิบ';
        else if (p === 1 && d === 2) s += 'ยี่สิบ';
        else if (p === 0 && d === 1 && (s || Math.floor(v / 10) > 0)) s += 'เอ็ด';
        else s += D[d] + P[p];
      });
      return s;
    };
    const neg = n < 0;
    const v = Math.round(Math.abs(n) * 100);
    const baht = Math.floor(v / 100); const sat = v % 100;
    let s = '';
    if (baht) s += `${words(baht)}บาท`;
    if (sat) s += `${words(sat)}สตางค์`; else s += baht ? 'ถ้วน' : 'ศูนย์บาทถ้วน';
    return (neg ? 'ลบ' : '') + s;
  }

  // ---------- прочее и функции внешних сервисов ----------
  const other = {
    INFO: ([t], ctx) => {
      const k = toText(t).toLowerCase();
      if (k === 'system') return 'pcdos';
      if (k === 'osversion') return 'Windows (64-bit) NT 10.00';
      if (k === 'release') return '16.0';
      if (k === 'recalc') return 'Автоматически';
      if (k === 'numfile') return 1;
      if (k === 'origin') return '$A:$A$1';
      if (k === 'directory') return '';
      return err(ERR.VALUE);
    },
    AREAS: (_a, _ctx, nodes) => (nodes[0] && (nodes[0].t === 'ref' || nodes[0].t === 'range' || nodes[0].t === 'name') ? 1 : err(ERR.VALUE)),
    // нужны подключение к данным/службе — без него Excel возвращает те же ошибки
    GETPIVOTDATA: () => err(ERR.REF),
    CUBEKPIMEMBER: () => err(ERR.NA), CUBEMEMBER: () => err(ERR.NA), CUBEMEMBERPROPERTY: () => err(ERR.NA),
    CUBERANKEDMEMBER: () => err(ERR.NA), CUBESET: () => err(ERR.NA), CUBESETCOUNT: () => err(ERR.NA), CUBEVALUE: () => err(ERR.NA),
    RTD: () => err(ERR.NA), WEBSERVICE: () => err(ERR.VALUE), STOCKHISTORY: () => err(ERR.VALUE),
  };

  return { ...math, ...eng, ...db, ...stat, ...text, ...other };
}
