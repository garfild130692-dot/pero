// Функции Microsoft 365: LAMBDA-помощники (MAP, REDUCE, SCAN, MAKEARRAY, BYROW, BYCOL), форма массивов
// (VSTACK, HSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, EXPAND, TRIMRANGE),
// регулярные выражения (REGEXTEST, REGEXEXTRACT, REGEXREPLACE), сводки (GROUPBY, PIVOTBY, PERCENTOF) и прочее
// (EUROCONVERT, JIS, IMAGE, CALL, REGISTER.ID). Сами LAMBDA и ISOMITTED — в вычислителе (formula.js).
export function functions6(h) {
  const { ERR, err, isErr, toNum, toText, toBool, flat, toArr, isFn, callFn, broadcast } = h;
  const NA = () => err(ERR.NA);
  const VALUE = () => err(ERR.VALUE);
  const CALC = () => err(ERR.CALC);
  const dims = (a) => [a.length, (a[0] || []).length];
  // результат лямбды для одной ячейки: массив 1×1 — значение, больший массив — #ВЫЧИСЛ! (вложенные массивы)
  const scalar = (v) => {
    if (Array.isArray(v)) return v.length === 1 && v[0].length === 1 ? v[0][0] ?? 0 : CALC();
    if (isFn(v)) return CALC();
    return v ?? 0;
  };
  const num = (v, d) => (v == null ? d : toNum(v));
  const bool = (v, d) => (v == null ? d : toBool(v));
  const transpose = (a) => (a.length ? a[0].map((_, j) => a.map((row) => row[j])) : a);
  const pad = (v) => (v === undefined ? NA() : v);

  // ---------- LAMBDA-помощники ----------
  const lambdaHelpers = {
    MAP: (args, ctx) => {
      const f = args[args.length - 1];
      const arrs = args.slice(0, -1);
      if (!isFn(f) || !arrs.length) return VALUE();
      for (const a of arrs) if (isErr(a)) return a;
      const as = arrs.map(toArr);
      const R = Math.max(...as.map((a) => a.length)); const C = Math.max(...as.map((a) => (a[0] || []).length));
      const pick = (a, i, j) => {
        const [r, c] = dims(a);
        const ii = r === 1 ? 0 : i; const jj = c === 1 ? 0 : j;
        return ii < r && jj < c ? a[ii][jj] : NA();
      };
      return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (__, j) => scalar(callFn(f, as.map((a) => pick(a, i, j)), ctx))));
    },
    REDUCE: ([init, arr, f], ctx) => {
      if (!isFn(f)) return VALUE();
      if (isErr(arr)) return arr;
      let acc = init ?? 0;
      for (const v of toArr(arr).flat()) acc = callFn(f, [acc, v], ctx);
      return Array.isArray(acc) && acc.length === 1 && acc[0].length === 1 ? acc[0][0] : acc;
    },
    SCAN: ([init, arr, f], ctx) => {
      if (!isFn(f)) return VALUE();
      if (isErr(arr)) return arr;
      let acc = init ?? 0;
      return toArr(arr).map((row) => row.map((v) => { acc = scalar(callFn(f, [acc, v], ctx)); return acc; }));
    },
    MAKEARRAY: ([rows, cols, f], ctx) => {
      const R = toNum(rows); const C = toNum(cols);
      if (isErr(R)) return R; if (isErr(C)) return C;
      if (!isFn(f)) return VALUE();
      if (Math.trunc(R) < 1 || Math.trunc(C) < 1) return VALUE();
      if (R * C > 1_048_576) return err(ERR.NUM);
      return Array.from({ length: Math.trunc(R) }, (_, i) => Array.from({ length: Math.trunc(C) }, (__, j) => scalar(callFn(f, [i + 1, j + 1], ctx))));
    },
    BYROW: ([arr, f], ctx) => {
      if (isErr(arr)) return arr;
      if (!isFn(f)) return VALUE();
      return toArr(arr).map((row) => [scalar(callFn(f, [[row]], ctx))]);
    },
    BYCOL: ([arr, f], ctx) => {
      if (isErr(arr)) return arr;
      if (!isFn(f)) return VALUE();
      return [transpose(toArr(arr)).map((col) => scalar(callFn(f, [col.map((v) => [v])], ctx)))];
    },
  };

  // ---------- форма массивов ----------
  function stack(args, vertical) {
    const as = [];
    for (const a of args) { if (a === undefined) continue; if (isErr(a) && !Array.isArray(a)) as.push([[a]]); else as.push(toArr(a)); }
    if (!as.length) return VALUE();
    if (vertical) {
      const C = Math.max(...as.map((a) => (a[0] || []).length));
      return as.flatMap((a) => a.map((row) => Array.from({ length: C }, (_, j) => pad(row[j]))));
    }
    const R = Math.max(...as.map((a) => a.length));
    return Array.from({ length: R }, (_, i) => as.flatMap((a) => { const C = (a[0] || []).length; return Array.from({ length: C }, (__, j) => (a[i] ? pad(a[i][j]) : NA())); }));
  }
  // элементы массива по строкам или по столбцам, с пропуском пустых (1), ошибок (2) или того и другого (3)
  function items(arr, ignore, byCol) {
    const a = byCol ? transpose(toArr(arr)) : toArr(arr);
    return a.flat().filter((v) => !((ignore & 1) && (v == null)) && !((ignore & 2) && isErr(v))).map((v) => v ?? 0);
  }
  function toLine(arr, ign, scan) {
    if (isErr(arr) && !Array.isArray(arr)) return arr;
    const ignore = num(ign, 0); const byCol = bool(scan, false);
    if (isErr(ignore)) return ignore; if (isErr(byCol)) return byCol;
    if (![0, 1, 2, 3].includes(Math.trunc(ignore))) return VALUE();
    const xs = items(arr, Math.trunc(ignore), byCol);
    return xs.length ? xs : CALC();
  }
  function wrap(vec, count, padWith, byCols) {
    if (isErr(vec) && !Array.isArray(vec)) return vec;
    const a = toArr(vec); const [r, c] = dims(a);
    if (r > 1 && c > 1) return VALUE();
    const n = Math.trunc(toNum(count));
    if (isErr(n)) return n;
    if (n < 1) return err(ERR.NUM);
    const xs = a.flat().map((v) => v ?? 0);
    const fill = padWith === undefined || padWith === null ? NA() : padWith;
    const out = [];
    for (let i = 0; i < xs.length; i += n) out.push(Array.from({ length: n }, (_, j) => (i + j < xs.length ? xs[i + j] : fill)));
    return byCols ? transpose(out) : out;
  }
  // номера, отсчитанные с конца для отрицательных
  function pickIndex(list, n) {
    const k = Math.trunc(toNum(n));
    if (isErr(k)) return k;
    if (k === 0 || Math.abs(k) > list) return VALUE();
    return k > 0 ? k - 1 : list + k;
  }
  function takeDrop(arr, rows, cols, take) {
    if (isErr(arr) && !Array.isArray(arr)) return arr;
    let a = toArr(arr); const [R, C] = dims(a);
    const cut = (len, n) => {
      if (n == null) return [0, len];
      const k = Math.trunc(toNum(n));
      if (isErr(k)) return k;
      if (take) return k >= 0 ? [0, Math.min(k, len)] : [Math.max(0, len + k), len];
      return k >= 0 ? [Math.min(k, len), len] : [0, Math.max(0, len + k)];
    };
    const rr = cut(R, rows); const cc = cut(C, cols);
    if (isErr(rr)) return rr; if (isErr(cc)) return cc;
    a = a.slice(rr[0], rr[1]).map((row) => row.slice(cc[0], cc[1]));
    return a.length && a[0].length ? a.map((row) => row.map((v) => v ?? 0)) : CALC();
  }
  const isBlank = (v) => v == null;
  const shape = {
    VSTACK: (args) => stack(args, true),
    HSTACK: (args) => stack(args, false),
    TOCOL: ([a, ign, scan]) => { const xs = toLine(a, ign, scan); return Array.isArray(xs) ? xs.map((v) => [v]) : xs; },
    TOROW: ([a, ign, scan]) => { const xs = toLine(a, ign, scan); return Array.isArray(xs) ? [xs] : xs; },
    WRAPROWS: ([v, n, p]) => wrap(v, n, p, false),
    WRAPCOLS: ([v, n, p]) => wrap(v, n, p, true),
    TAKE: ([a, r, c]) => takeDrop(a, r, c, true),
    DROP: ([a, r, c]) => takeDrop(a, r, c, false),
    CHOOSECOLS: ([arr, ...ns]) => {
      if (isErr(arr) && !Array.isArray(arr)) return arr;
      const a = toArr(arr); const C = (a[0] || []).length;
      if (!ns.length) return VALUE();
      const idx = [];
      for (const n of ns) for (const x of toArr(n).flat()) { const k = pickIndex(C, x); if (isErr(k)) return k; idx.push(k); }
      return a.map((row) => idx.map((j) => row[j] ?? 0));
    },
    CHOOSEROWS: ([arr, ...ns]) => {
      if (isErr(arr) && !Array.isArray(arr)) return arr;
      const a = toArr(arr);
      if (!ns.length) return VALUE();
      const idx = [];
      for (const n of ns) for (const x of toArr(n).flat()) { const k = pickIndex(a.length, x); if (isErr(k)) return k; idx.push(k); }
      return idx.map((i) => a[i].map((v) => v ?? 0));
    },
    EXPAND: ([arr, rows, cols, padWith]) => {
      if (isErr(arr) && !Array.isArray(arr)) return arr;
      const a = toArr(arr); const [R0, C0] = dims(a);
      const R = rows == null ? R0 : Math.trunc(toNum(rows)); const C = cols == null ? C0 : Math.trunc(toNum(cols));
      if (isErr(R)) return R; if (isErr(C)) return C;
      if (R < R0 || C < C0) return VALUE();
      if (R * C > 1_048_576) return err(ERR.NUM);
      const fill = padWith == null ? NA() : padWith;
      return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (__, j) => (i < R0 && j < C0 ? a[i][j] ?? 0 : fill)));
    },
    TRIMRANGE: ([arr, tr, tc]) => {
      if (isErr(arr) && !Array.isArray(arr)) return arr;
      const a = toArr(arr);
      const mr = Math.trunc(num(tr, 3)); const mc = Math.trunc(num(tc, 3));
      if (isErr(mr)) return mr; if (isErr(mc)) return mc;
      if (![0, 1, 2, 3].includes(mr) || ![0, 1, 2, 3].includes(mc)) return VALUE();
      const [R, C] = dims(a);
      let r0 = 0; let r1 = R; let c0 = 0; let c1 = C;
      const rowEmpty = (i) => a[i].every(isBlank); const colEmpty = (j) => a.every((row) => isBlank(row[j]));
      if (mr & 1) while (r0 < r1 && rowEmpty(r0)) r0++;
      if (mr & 2) while (r1 > r0 && rowEmpty(r1 - 1)) r1--;
      if (mc & 1) while (c0 < c1 && colEmpty(c0)) c0++;
      if (mc & 2) while (c1 > c0 && colEmpty(c1 - 1)) c1--;
      if (r0 >= r1 || c0 >= c1) return [[0]];
      return a.slice(r0, r1).map((row) => row.slice(c0, c1).map((v) => v ?? 0));
    },
  };

  // ---------- регулярные выражения ----------
  // Excel — PCRE2; движок JavaScript совпадает с ним в обычных случаях (классы, группы, квантификаторы, \d \w \s, $1)
  function regex(pattern, flags, cs) {
    const ci = num(cs, 0);
    if (isErr(ci)) return ci;
    if (ci !== 0 && ci !== 1) return VALUE();
    try { return new RegExp(toText(pattern), flags + (ci === 1 ? 'i' : '') + 'u'); } catch { return VALUE(); }
  }
  const lift = (args, f) => (args.some(Array.isArray) ? broadcast(args, f) : f(args));
  const regexFns = {
    REGEXTEST: (args) => lift(args.slice(0, 3), ([t, p, cs]) => {
      const e = [t, p, cs].find(isErr); if (e) return e;
      const re = regex(p, '', cs); if (isErr(re)) return re;
      return re.test(toText(t ?? ''));
    }),
    REGEXEXTRACT: ([text, pattern, mode, cs]) => {
      const one = ([t, p, m]) => {
        const e = [t, p, m, cs].find(isErr); if (e) return e;
        const md = Math.trunc(num(m, 0)); if (isErr(md)) return md;
        if (![0, 1, 2].includes(md)) return VALUE();
        const re = regex(p, md === 1 ? 'g' : '', cs); if (isErr(re)) return re;
        const s = toText(t ?? '');
        if (md === 1) { const all = [...s.matchAll(re)].map((x) => [x[0]]); return all.length ? all : NA(); }
        const m0 = re.exec(s);
        if (!m0) return NA();
        if (md === 0) return m0[0];
        return m0.length > 1 ? [m0.slice(1).map((g) => g ?? '')] : NA();
      };
      if (Array.isArray(text) || Array.isArray(pattern)) return broadcast([text, pattern, mode], (xs) => scalar(one(xs)));
      return one([text, pattern, mode]);
    },
    REGEXREPLACE: (args) => lift(args.slice(0, 5), ([t, p, rep, occ, cs]) => {
      const e = [t, p, rep, occ, cs].find(isErr); if (e) return e;
      const n = Math.trunc(num(occ, 0)); if (isErr(n)) return n;
      const re = regex(p, 'g', cs); if (isErr(re)) return re;
      const s = toText(t ?? ''); const r = toText(rep ?? '');
      if (n === 0) return s.replace(re, r);
      const ms = [...s.matchAll(re)];
      const k = n > 0 ? n - 1 : ms.length + n;
      const m = ms[k];
      if (!m) return s;
      const one = new RegExp(re.source, re.flags.replace('g', ''));
      return s.slice(0, m.index) + m[0].replace(one, r) + s.slice(m.index + m[0].length);
    }),
  };

  // ---------- сводки ----------
  // Сводка группы: функция-значение от столбца значений; ПРОЦЕНТ получает (группа, все значения)
  const isPercentOf = (f) => f && f.fnref === 'PERCENTOF';
  function aggregate(f, vals, all, ctx) {
    const col = vals.map((v) => [v]);
    const res = isPercentOf(f) ? callFn(f, [col, all.map((v) => [v])], ctx) : callFn(f, [col], ctx);
    return scalar(res);
  }
  const TOTAL = 'Итог';
  const keyOf = (vs) => JSON.stringify(vs.map((v) => (typeof v === 'string' ? v.toLowerCase() : v ?? null)));
  const cmpKeys = (x, y) => {
    for (let i = 0; i < x.length; i++) {
      const a = x[i]; const b = y[i];
      const ra = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2; const rb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
      if (ra !== rb) return ra - rb;
      const r = ra === 0 ? a - b : ra === 1 ? a.localeCompare(b, 'ru', { sensitivity: 'base' }) : 0;
      if (r) return r;
    }
    return 0;
  };
  // заголовки: 0 — нет, 1 — есть, не показывать, 2 — нет, создать, 3 — есть, показать; по умолчанию — угадать
  function headersMode(hdr, fields, values) {
    if (hdr != null) return Math.trunc(toNum(hdr));
    const f0 = fields[0] || []; const v0 = values[0] || []; const v1 = values[1] || [];
    const looksHeader = v0.length && v0.every((v) => typeof v === 'string') && v1.some((v) => typeof v === 'number') && f0.every((v) => typeof v === 'string');
    return looksHeader ? 1 : 0;
  }
  function prepare(fields, values, hdr, filter) {
    if (isErr(fields) && !Array.isArray(fields)) return fields;
    if (isErr(values) && !Array.isArray(values)) return values;
    let F = toArr(fields); let V = toArr(values);
    if (F.length !== V.length) return VALUE();
    const mode = headersMode(hdr, F, V);
    if (isErr(mode)) return mode;
    let fh = null; let vh = null;
    if (mode === 1 || mode === 3) { fh = F[0]; vh = V[0]; F = F.slice(1); V = V.slice(1); }
    if (mode === 2) { fh = F[0].map((_, j) => `Поле ${j + 1}`); vh = V[0].map((_, j) => `Значение ${j + 1}`); }
    let keep = F.map(() => true);
    if (filter != null) {
      let m = toArr(filter).map((row) => row[0]);
      if (mode === 1 || mode === 3) m = m.slice(m.length - F.length);
      if (m.length !== F.length) return VALUE();
      keep = m.map((v) => { const b = toBool(v); return !isErr(b) && b; });
    }
    const rows = [];
    for (let i = 0; i < F.length; i++) if (keep[i]) rows.push({ f: F[i].map((v) => v ?? ''), v: V[i] });
    return { rows, fh: mode >= 2 ? fh : null, vh: mode >= 2 ? vh : null };
  }
  function groups(rows, depth, getKey) {
    const map = new Map();
    for (const r of rows) {
      const k = getKey(r).slice(0, depth);
      const s = keyOf(k);
      if (!map.has(s)) map.set(s, { key: k, rows: [] });
      map.get(s).rows.push(r);
    }
    return [...map.values()];
  }
  // GROUPBY(поля_строк; значения; функция; [заголовки]; [итоги]; [сортировка]; [фильтр]; [связь полей])
  function groupby([fields, values, f, hdr, totalDepth, sortOrder, filter], ctx) {
    if (!isFn(f)) return VALUE();
    const p = prepare(fields, values, hdr, filter);
    if (isErr(p) && !Array.isArray(p)) return p;
    const { rows, fh, vh } = p;
    if (!rows.length) return CALC();
    const nf = rows[0].f.length; const nv = rows[0].v.length;
    const td = Math.trunc(num(totalDepth, nf > 1 ? 2 : 1));
    if (isErr(td)) return td;
    const atTop = td < 0; const depth = Math.abs(td);
    const all = (j) => rows.map((r) => r.v[j]);
    const aggRow = (rs) => Array.from({ length: nv }, (_, j) => aggregate(f, rs.map((r) => r.v[j]), all(j), ctx));
    const leaves = groups(rows, nf, (r) => r.f).map((g) => ({ key: g.key, vals: aggRow(g.rows) }));
    // сортировка: номера столбцов результата (1..nf — поля, далее — значения), минус — по убыванию
    const orders = sortOrder == null ? [] : toArr(sortOrder).flat().map((x) => Math.trunc(toNum(x)));
    if (orders.some(isErr)) return orders.find(isErr);
    const cmpLeaf = (a, b) => {
      for (const o of orders) {
        const k = Math.abs(o) - 1; const sign = o < 0 ? -1 : 1;
        const x = k < nf ? a.key[k] : a.vals[k - nf]; const y = k < nf ? b.key[k] : b.vals[k - nf];
        const r = cmpKeys([x], [y]); if (r) return sign * r;
      }
      return cmpKeys(a.key, b.key);
    };
    // вывод с промежуточными итогами по уровням 1..depth-1 (уровень 1 — общий итог)
    const out = [];
    const emit = (list, level) => {
      if (level >= nf || level >= depth - 1) { list.sort(cmpLeaf).forEach((l) => out.push([...l.key, ...l.vals])); return; }
      const byLevel = new Map();
      for (const l of list) { const s = keyOf(l.key.slice(0, level + 1)); if (!byLevel.has(s)) byLevel.set(s, []); byLevel.get(s).push(l); }
      const gs = [...byLevel.values()].sort((a, b) => cmpLeaf(a[0], b[0]));
      for (const g of gs) {
        const k = g[0].key.slice(0, level + 1);
        const sub = rows.filter((r) => keyOf(r.f.slice(0, level + 1)) === keyOf(k));
        const subRow = [...k, ...Array(nf - level - 1).fill(''), ...aggRow(sub)];
        if (atTop) out.push(subRow);
        emit(g, level + 1);
        if (!atTop) out.push(subRow);
      }
    };
    emit(leaves, 0);
    if (depth >= 1) {
      const total = [TOTAL, ...Array(nf - 1).fill(''), ...aggRow(rows)];
      if (atTop) out.unshift(total); else out.push(total);
    }
    if (fh) out.unshift([...fh, ...vh]);
    return out;
  }
  // PIVOTBY(поля_строк; поля_столбцов; значения; функция; [заголовки]; [итоги_строк]; [сорт_строк]; [итоги_столбцов]; [сорт_столбцов]; [фильтр])
  function pivotby([rf, cf, values, f, hdr, rowTotal, rowSort, colTotal, colSort, filter], ctx) {
    if (!isFn(f)) return VALUE();
    if (isErr(cf) && !Array.isArray(cf)) return cf;
    const C = toArr(cf);
    const p = prepare(rf, values, hdr, filter);
    if (isErr(p) && !Array.isArray(p)) return p;
    const V = toArr(values);
    // поля столбцов — те же строки, что остались после заголовка и фильтра
    const F = toArr(rf);
    const mode = headersMode(hdr, F, V);
    const hdrRows = mode === 1 || mode === 3 ? 1 : 0;
    let keep = F.slice(hdrRows).map(() => true);
    if (filter != null) keep = toArr(filter).map((row) => row[0]).slice(-keep.length).map((v) => { const b = toBool(v); return !isErr(b) && b; });
    const cols = C.slice(hdrRows).filter((_, i) => keep[i]).map((row) => row.map((v) => v ?? ''));
    const rows = p.rows.map((r, i) => ({ ...r, c: cols[i] }));
    if (!rows.length) return CALC();
    const nv = rows[0].v.length;
    const all = (j) => rows.map((r) => r.v[j]);
    const agg = (rs) => Array.from({ length: nv }, (_, j) => (rs.length ? aggregate(f, rs.map((r) => r.v[j]), all(j), ctx) : ''));
    const sortList = (list, so) => {
      const o = so == null ? 1 : Math.trunc(toNum(toArr(so)[0][0]));
      return list.sort((a, b) => (o < 0 ? -1 : 1) * cmpKeys(a.key, b.key));
    };
    const rg = sortList(groups(rows, rows[0].f.length, (r) => r.f), rowSort);
    const cg = sortList(groups(rows, rows[0].c.length, (r) => r.c), colSort);
    const rt = Math.trunc(num(rowTotal, 1)); const ct = Math.trunc(num(colTotal, 1));
    const nf = rows[0].f.length; const nc = rows[0].c.length;
    const out = [];
    // шапка: значения полей столбцов (по уровню в строке), затем — «Итог»
    for (let lv = 0; lv < nc; lv++) {
      const row = Array(nf).fill('');
      for (const g of cg) for (let j = 0; j < nv; j++) row.push(j === 0 ? g.key[lv] : '');
      if (ct !== 0) for (let j = 0; j < nv; j++) row.push(lv === 0 && j === 0 ? TOTAL : '');
      out.push(row);
    }
    if (p.vh && nv > 1) out.push([...Array(nf).fill(''), ...[...cg, ...(ct !== 0 ? [0] : [])].flatMap(() => p.vh)]);
    const line = (key, rs) => {
      const row = [...key];
      for (const g of cg) row.push(...agg(rs.filter((r) => keyOf(r.c) === keyOf(g.key))));
      if (ct !== 0) row.push(...agg(rs));
      return row;
    };
    const body = rg.map((g) => line(g.key, g.rows));
    const total = rt !== 0 ? line([TOTAL, ...Array(nf - 1).fill('')], rows) : null;
    if (total && rt < 0) out.push(total);
    out.push(...body);
    if (total && rt > 0) out.push(total);
    if (p.fh) out[0].splice(0, nf, ...p.fh);
    return out;
  }
  const summary = {
    GROUPBY: groupby,
    PIVOTBY: pivotby,
    PERCENTOF: ([part, all]) => {
      const sum = (x) => { let s = 0; for (const v of flat(x)) { if (isErr(v)) return v; if (typeof v === 'number') s += v; } return s; };
      const a = sum(part); const b = sum(all);
      if (isErr(a)) return a; if (isErr(b)) return b;
      return b === 0 ? err(ERR.DIV0) : a / b;
    },
  };

  // ---------- прочее ----------
  // Курсы и точность округления валют еврозоны (фиксированные курсы ЕС)
  const EURO = {
    EUR: [1, 2], BEF: [40.3399, 0], LUF: [40.3399, 0], DEM: [1.95583, 2], ESP: [166.386, 0], FRF: [6.55957, 2], IEP: [0.787564, 2],
    ITL: [1936.27, 0], NLG: [2.20371, 2], ATS: [13.7603, 2], PTE: [200.482, 0], FIM: [5.94573, 2], GRD: [340.75, 0], SIT: [239.64, 2],
    CYP: [0.585274, 2], MTL: [0.4293, 2], SKK: [30.126, 2], EEK: [15.6466, 2], LVL: [0.702804, 2], LTL: [3.4528, 2],
  };
  const r10 = (x, d) => { const m = 10 ** d; return Math.round(x * m + Math.sign(x) * 1e-9) / m; };
  const other = {
    EUROCONVERT: ([n, src, dst, full, tri]) => {
      const x = toNum(n); if (isErr(x)) return x;
      const s = EURO[toText(src ?? '').toUpperCase()]; const d = EURO[toText(dst ?? '').toUpperCase()];
      if (!s || !d) return VALUE();
      const fp = bool(full, false); if (isErr(fp)) return fp;
      let eur = x / s[0];
      if (tri != null) { const t = Math.trunc(toNum(tri)); if (isErr(t)) return t; if (t < 3) return VALUE(); if (s[0] !== 1 && d[0] !== 1) eur = r10(eur, t); }
      const res = eur * d[0];
      return fp ? res : r10(res, d[1]);
    },
    // в русской версии Excel JIS текст не меняет (как ASC/DBCS)
    JIS: ([t]) => (isErr(t) ? t : toText(t ?? '')),
    // картинка из Интернета — внешний сервис, в Перо не загружается
    IMAGE: ([src]) => (isErr(src) ? src : VALUE()),
    // вызов функций DLL Excel отключён и сам; в Перо недоступен
    CALL: () => VALUE(),
    'REGISTER.ID': () => VALUE(),
  };
  return { ...lambdaHelpers, ...shape, ...regexFns, ...summary, ...other };
}
