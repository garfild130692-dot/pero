// Библиотека функций, пакет D (этап 12.5): финансовые — ценные бумаги (купоны, облигации, векселя, нерегулярные
// периоды), накопленные суммы и прочее. Базисы подсчёта дней: 0 — 30/360 (США), 1 — факт/факт, 2 — факт/360,
// 3 — факт/365, 4 — 30/360 (Европа). Формулы — по спецификации OOXML; сверка с Excel — tests/formulas.test.js.
export function functions5(h) {
  const { ERR, err, isErr, toNum, toBool, flat, firstErr, dateToSerial, serialToDate, FN } = h;
  const num = (x, def) => (x == null && def !== undefined ? def : toNum(x));
  const NUM = () => err(ERR.NUM);
  const day = (x) => Math.floor(x);
  const lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const isLeap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const yearfrac = (a, b, basis) => FN.YEARFRAC([a, b, basis]);

  // аргументы: d — дата (целая часть), i — целое, n — число, b — логическое; f(…) — результат
  const fn = (defs, f) => (args) => {
    const v = defs.map((d, i) => {
      const [kind, def] = Array.isArray(d) ? d : [d];
      if (args[i] == null && def !== undefined) return def;
      if (kind === 'b') { const b = toBool(args[i]); return b; }
      const n = toNum(args[i]);
      if (isErr(n)) return n;
      return kind === 'd' || kind === 'i' ? Math.trunc(n) : n;
    });
    const e = firstErr(...v);
    if (e) return e;
    const r = f(...v);
    return typeof r === 'number' && !isFinite(r) ? NUM() : r;
  };
  const okFreq = (f) => f === 1 || f === 2 || f === 4;
  const okBasis = (b) => b >= 0 && b <= 4;

  // ---------- купоны ----------
  // сдвиг даты на k месяцев с правилом «конец месяца» (если дата погашения — последний день месяца)
  const addMonths = (s, k, eom) => {
    const d = serialToDate(s);
    const total = d.y * 12 + d.m - 1 + k;
    const y = Math.floor(total / 12); const m = (total % 12) + 1;
    const dd = eom ? lastDay(y, m) : Math.min(d.d, lastDay(y, m));
    return dateToSerial(y, m, dd);
  };
  // предыдущая и следующая даты купона относительно расчёта (от погашения назад)
  function coupons(settle, mat, freq) {
    const step = 12 / freq;
    const md = serialToDate(mat);
    const eom = md.d === lastDay(md.y, md.m);
    let n = 0; let pcd = mat;
    while (pcd > settle) { n++; pcd = addMonths(mat, -step * n, eom); }
    const ncd = addMonths(mat, -step * (n - 1), eom);
    return { pcd, ncd, num: n };
  }
  // дни между датами по базису (для 30/360 — правила Excel для купонов)
  function days360(a, b, european) {
    const x = serialToDate(a); const y = serialToDate(b);
    let d1 = x.d; let d2 = y.d;
    if (european) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
    else {
      const febEnd = (d) => d.m === 2 && d.d === lastDay(d.y, 2);
      if (febEnd(x) && febEnd(y)) d2 = 30;
      if (febEnd(x)) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;
      if (d1 === 31) d1 = 30;
    }
    return (y.y - x.y) * 360 + (y.m - x.m) * 30 + (d2 - d1);
  }
  const daysBasis = (a, b, basis) => (basis === 0 ? days360(a, b, false) : basis === 4 ? days360(a, b, true) : b - a);
  const coupDays = (s, m, f, basis) => {
    if (basis === 1) { const c = coupons(s, m, f); return c.ncd - c.pcd; }
    return (basis === 3 ? 365 : 360) / f;
  };
  const coupDayBs = (s, m, f, basis) => daysBasis(coupons(s, m, f).pcd, s, basis);
  const coupDaysNc = (s, m, f, basis) => {
    if (basis === 0) return coupDays(s, m, f, basis) - coupDayBs(s, m, f, basis);
    if (basis === 4) return days360(s, coupons(s, m, f).ncd, true);
    return coupons(s, m, f).ncd - s;
  };
  // дни до следующего купона внутри формул ЦЕНА/ДЛИТ: при постоянной длине периода — «период минус прошедшее» (как Excel)
  const dscOf = (s, m, f, basis) => (basis === 1 ? coupons(s, m, f).ncd - s : coupDays(s, m, f, basis) - coupDayBs(s, m, f, basis));
  const coupArgs = ['d', 'd', 'i', ['i', 0]];
  const coupCheck = (s, m, f, b) => s < m && okFreq(f) && okBasis(b);
  const cp = (f) => fn(coupArgs, (s, m, fr, b) => (coupCheck(s, m, fr, b) ? f(s, m, fr, b) : NUM()));

  // ---------- облигации ----------
  function price(s, m, rate, yld, red, f, basis) {
    const N = coupons(s, m, f).num; const E = coupDays(s, m, f, basis);
    const DSC = dscOf(s, m, f, basis); const A = coupDayBs(s, m, f, basis);
    const c = (100 * rate) / f;
    if (N === 1) return (red + c) / (1 + (DSC / E) * (yld / f)) - c * (A / E);
    let p = red / (1 + yld / f) ** (N - 1 + DSC / E);
    for (let k = 1; k <= N; k++) p += c / (1 + yld / f) ** (k - 1 + DSC / E);
    return p - c * (A / E);
  }
  function solve(f, guess, lo = -0.999, hi = 100) {
    let x = guess;
    for (let i = 0; i < 100; i++) {
      const y = f(x); const d = (f(x + 1e-8) - y) / 1e-8;
      if (!isFinite(y) || !isFinite(d) || d === 0) break;
      const nx = x - y / d;
      if (Math.abs(nx - x) < 1e-13) return nx;
      x = nx;
    }
    // запасной путь — деление отрезка
    let a = lo; let b = hi; let fa = f(a);
    if (!isFinite(fa) || fa * f(b) > 0) return NaN;
    for (let i = 0; i < 300; i++) { const c = (a + b) / 2; const fc = f(c); if (fa * fc <= 0) b = c; else { a = c; fa = fc; } if (b - a < 1e-15) break; }
    return (a + b) / 2;
  }
  function duration(s, m, cpn, yld, f, basis) {
    const N = coupons(s, m, f).num;
    const frac = dscOf(s, m, f, basis) / coupDays(s, m, f, basis); // доля периода до ближайшего купона
    const c = (cpn * 100) / f; const y = yld / f + 1;
    let dur = 0; let p = 0;
    for (let k = 1; k <= N; k++) {
      const t = k - 1 + frac; const cf = k === N ? c + 100 : c;
      dur += (t * cf) / y ** t; p += cf / y ** t;
    }
    return dur / p / f;
  }
  const bond = {
    COUPPCD: cp((s, m, f) => coupons(s, m, f).pcd),
    COUPNCD: cp((s, m, f) => coupons(s, m, f).ncd),
    COUPNUM: cp((s, m, f) => coupons(s, m, f).num),
    COUPDAYS: cp(coupDays), COUPDAYBS: cp(coupDayBs), COUPDAYSNC: cp(coupDaysNc),
    PRICE: fn(['d', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, r, y, red, f, b) => (s >= m || r < 0 || y < 0 || red <= 0 || !okFreq(f) || !okBasis(b) ? NUM() : price(s, m, r, y, red, f, b))),
    YIELD: fn(['d', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, r, pr, red, f, b) => {
      if (s >= m || r < 0 || pr <= 0 || red <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      const N = coupons(s, m, f).num;
      if (N === 1) {
        const A = coupDayBs(s, m, f, b); const E = coupDays(s, m, f, b); const DSR = E - A;
        return ((red / 100 + r / f) - (pr / 100 + (A / E) * (r / f))) / (pr / 100 + (A / E) * (r / f)) * ((f * E) / DSR);
      }
      return solve((y) => price(s, m, r, y, red, f, b) - pr, r || 0.05);
    }),
    DURATION: fn(['d', 'd', 'n', 'n', 'i', ['i', 0]], (s, m, c, y, f, b) => (s >= m || c < 0 || y < 0 || !okFreq(f) || !okBasis(b) ? NUM() : duration(s, m, c, y, f, b))),
    MDURATION: fn(['d', 'd', 'n', 'n', 'i', ['i', 0]], (s, m, c, y, f, b) => (s >= m || c < 0 || y < 0 || !okFreq(f) || !okBasis(b) ? NUM() : duration(s, m, c, y, f, b) / (1 + y / f))),
    ACCRINT: fn(['d', 'd', 'd', 'n', ['n', 1000], 'i', ['i', 0], ['b', true]], (iss, first, s, r, par, f, b, calc) => {
      if (iss >= s || r <= 0 || par <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      // способ_расчёта ЛОЖЬ и расчёт после первой выплаты — проценты от первой выплаты, иначе от выпуска
      const from = !calc && s > first ? first : iss;
      return par * r * yearfrac(from, s, b);
    }),
    ACCRINTM: fn(['d', 'd', 'n', ['n', 1000], ['i', 0]], (iss, s, r, par, b) => (iss >= s || r <= 0 || par <= 0 || !okBasis(b) ? NUM() : par * r * yearfrac(iss, s, b))),
    DISC: fn(['d', 'd', 'n', 'n', ['i', 0]], (s, m, pr, red, b) => (s >= m || pr <= 0 || red <= 0 || !okBasis(b) ? NUM() : (red - pr) / red / yearfrac(s, m, b))),
    INTRATE: fn(['d', 'd', 'n', 'n', ['i', 0]], (s, m, inv, red, b) => (s >= m || inv <= 0 || red <= 0 || !okBasis(b) ? NUM() : (red - inv) / inv / yearfrac(s, m, b))),
    RECEIVED: fn(['d', 'd', 'n', 'n', ['i', 0]], (s, m, inv, disc, b) => {
      if (s >= m || inv <= 0 || disc <= 0 || !okBasis(b)) return NUM();
      const v = 1 - disc * yearfrac(s, m, b);
      return v <= 0 ? NUM() : inv / v;
    }),
    PRICEDISC: fn(['d', 'd', 'n', 'n', ['i', 0]], (s, m, disc, red, b) => (s >= m || disc <= 0 || red <= 0 || !okBasis(b) ? NUM() : red - disc * red * yearfrac(s, m, b))),
    YIELDDISC: fn(['d', 'd', 'n', 'n', ['i', 0]], (s, m, pr, red, b) => (s >= m || pr <= 0 || red <= 0 || !okBasis(b) ? NUM() : (red - pr) / pr / yearfrac(s, m, b))),
    PRICEMAT: fn(['d', 'd', 'd', 'n', 'n', ['i', 0]], (s, m, iss, r, y, b) => {
      if (s >= m || r < 0 || y < 0 || !okBasis(b)) return NUM();
      const B = b === 1 ? (isLeap(serialToDate(s).y) ? 366 : 365) : b === 3 ? 365 : 360;
      const dsm = daysBasis(s, m, b); const dim = daysBasis(iss, m, b); const a = daysBasis(iss, s, b);
      return (100 + (dim / B) * r * 100) / (1 + (dsm / B) * y) - (a / B) * r * 100;
    }),
    YIELDMAT: fn(['d', 'd', 'd', 'n', 'n', ['i', 0]], (s, m, iss, r, pr, b) => {
      if (s >= m || r < 0 || pr <= 0 || !okBasis(b)) return NUM();
      const B = b === 1 ? (isLeap(serialToDate(s).y) ? 366 : 365) : b === 3 ? 365 : 360;
      const dsm = daysBasis(s, m, b); const dim = daysBasis(iss, m, b); const a = daysBasis(iss, s, b);
      return ((1 + (dim / B) * r) / (pr / 100 + (a / B) * r) - 1) * (B / dsm);
    }),
    TBILLPRICE: fn(['d', 'd', 'n'], (s, m, d) => (s >= m || m - s > 366 || d <= 0 ? NUM() : 100 * (1 - (d * (m - s)) / 360))),
    TBILLYIELD: fn(['d', 'd', 'n'], (s, m, pr) => (s >= m || m - s > 366 || pr <= 0 ? NUM() : ((100 - pr) / pr) * (360 / (m - s)))),
    TBILLEQ: fn(['d', 'd', 'n'], (s, m, d) => {
      const dsm = m - s;
      if (s >= m || dsm > 366 || d <= 0) return NUM();
      if (dsm <= 182) return (365 * d) / (360 - d * dsm);
      // больше полугода — эквивалент купонной доходности (квадратное уравнение, как Excel)
      const pr = 100 * (1 - (d * dsm) / 360);
      const a = dsm / 365 / 2 - 0.25; const bq = dsm / 365; const c = (pr - 100) / pr;
      return (-bq + Math.sqrt(bq * bq - 4 * a * c)) / (2 * a);
    }),
    // нерегулярный последний период
    ODDLPRICE: fn(['d', 'd', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, last, r, y, red, f, b) => {
      if (!(last < s && s < m) || r < 0 || y < 0 || red <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      const E = (b === 3 ? 365 : 360) / f; // длина квазипериода
      const DCi = daysBasis(last, m, b); const Ai = daysBasis(last, s, b); const DSCi = daysBasis(s, m, b);
      const nc = Math.ceil(DCi / E - 1e-9) || 1;
      const x = DCi / E; const aRatio = Ai / E; const dsc = DSCi / E;
      if (nc === 1 || true) return (red + x * 100 * (r / f)) / (1 + dsc * (y / f)) - aRatio * 100 * (r / f);
    }),
    ODDLYIELD: fn(['d', 'd', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, last, r, pr, red, f, b) => {
      if (!(last < s && s < m) || r < 0 || pr <= 0 || red <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      const E = (b === 3 ? 365 : 360) / f;
      const x = daysBasis(last, m, b) / E; const aRatio = daysBasis(last, s, b) / E; const dsc = daysBasis(s, m, b) / E;
      return ((red + x * 100 * (r / f)) - (pr + aRatio * 100 * (r / f))) / (pr + aRatio * 100 * (r / f)) * (f / dsc);
    }),
    // нерегулярный первый период
    ODDFPRICE: fn(['d', 'd', 'd', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, iss, first, r, y, red, f, b) => {
      if (!(iss < s && s < first && first < m) || r < 0 || y < 0 || red <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      return oddfPrice(s, m, iss, first, r, y, red, f, b);
    }),
    ODDFYIELD: fn(['d', 'd', 'd', 'd', 'n', 'n', 'n', 'i', ['i', 0]], (s, m, iss, first, r, pr, red, f, b) => {
      if (!(iss < s && s < first && first < m) || r < 0 || pr <= 0 || red <= 0 || !okFreq(f) || !okBasis(b)) return NUM();
      return solve((y) => oddfPrice(s, m, iss, first, r, y, red, f, b) - pr, r || 0.05);
    }),
  };
  // ЦЕНАПЕРВНЕРЕГ: формула OOXML (короткий и длинный первый период)
  function oddfPrice(s, m, iss, first, r, y, red, f, b) {
    const step = 12 / f;
    const md = serialToDate(m); const eom = md.d === lastDay(md.y, md.m);
    // длина квазипериода, который заканчивается первой выплатой (факт/факт — по календарю)
    const E = b === 1 ? first - addMonths(first, -step, eom) : (b === 3 ? 365 : 360) / f;
    const N = coupons(first, m, f).num + 1; // купоны после расчёта, включая первый
    const DSC = daysBasis(s, first, b);
    // первый период: короткий (один квазипериод) или длинный (несколько)
    let pcd = first; let nq = 0;
    while (pcd > iss) { nq++; pcd = addMonths(first, -step * nq, eom); }
    const c = (100 * r) / f; const yf = 1 + y / f;
    if (nq <= 1) {
      const DFC = daysBasis(iss, first, b); const A = daysBasis(iss, s, b);
      let p = red / yf ** (N - 1 + DSC / E) + (c * DFC / E) / yf ** (DSC / E);
      for (let k = 2; k <= N; k++) p += c / yf ** (k - 1 + DSC / E);
      return p - c * (A / E);
    }
    // длинный первый период: суммы по квазипериодам
    let sumA = 0; let sumDC = 0; let nl = 0;
    for (let i = nq; i >= 1; i--) {
      const qs = addMonths(first, -step * i, eom); const qe = addMonths(first, -step * (i - 1), eom);
      const ql = b === 1 ? qe - qs : E;
      const from = Math.max(qs, iss);
      const dc = daysBasis(from, qe, b) / ql;
      sumDC += Math.min(dc, 1);
      if (s > qs) { sumA += daysBasis(from, Math.min(s, qe), b) / ql; }
      if (s >= qe) nl++;
    }
    const Nq = nq; const DSCq = daysBasis(s, coupons(s, first, f).ncd, b) / E;
    let p = red / yf ** (N + Nq - nl - 1 - 1 + DSCq + 1) ;
    p = red / yf ** (N - 1 + Nq - nl - 1 + DSCq);
    p += (c * sumDC) / yf ** (Nq - nl - 1 + DSCq);
    for (let k = 1; k < N; k++) p += c / yf ** (k + Nq - nl - 1 + DSCq);
    return p - c * sumA;
  }

  // ---------- прочие финансовые ----------
  // ПУО — как в LibreOffice/Excel: уменьшение остатка с переходом на линейную, дробные начало и конец
  function ddbOf(cost, salvage, life, period, factor) {
    let rate = factor / life; let old;
    if (rate >= 1) { rate = 1; old = period === 1 ? cost : 0; } else old = cost * (1 - rate) ** (period - 1);
    const now = cost * (1 - rate) ** period;
    const d = now < salvage ? old - salvage : old - now;
    return d < 0 ? 0 : d;
  }
  function interVdb(cost, salvage, life, life1, period, factor) {
    let v = 0; const intEnd = Math.ceil(period - 1e-12); let rest = cost - salvage; let sln = 0; let nowSln = false;
    for (let i = 1; i <= intEnd; i++) {
      let term;
      if (!nowSln) {
        const ddb = ddbOf(cost, salvage, life, i, factor);
        sln = rest / (life1 - (i - 1));
        if (sln > ddb) { term = sln; nowSln = true; } else { term = ddb; rest -= ddb; }
      } else term = sln;
      if (i === intEnd) term *= period + 1 - intEnd;
      v += term;
    }
    return v;
  }
  function vdb(cost, salvage, life, start, end, factor, noSwitch) {
    const intStart = Math.floor(start + 1e-12); const intEnd = Math.ceil(end - 1e-12);
    if (noSwitch) {
      let total = 0;
      for (let i = intStart + 1; i <= intEnd; i++) {
        let term = ddbOf(cost, salvage, life, i, factor);
        if (i === intStart + 1) term *= Math.min(end, intStart + 1) - start;
        else if (i === intEnd) term *= end + 1 - intEnd;
        total += term;
      }
      return total;
    }
    let part = 0;
    if (Math.abs(start - intStart) > 1e-12) {
      const tmp = cost - interVdb(cost, salvage, life, life, intStart, factor);
      part += (start - intStart) * interVdb(tmp, salvage, life, life - intStart, 1, factor);
    }
    if (Math.abs(end - intEnd) > 1e-12) {
      const ts = intEnd - 1;
      const tmp = cost - interVdb(cost, salvage, life, life, ts, factor);
      part += (intEnd - end) * interVdb(tmp, salvage, life, life - ts, intEnd - ts, factor);
    }
    const c = cost - interVdb(cost, salvage, life, life, intStart, factor);
    return interVdb(c, salvage, life, life - intStart, intEnd - intStart, factor) - part;
  }
  const other = {
    CUMIPMT: fn(['n', 'i', 'n', 'i', 'i', 'i'], (r, n, pv, st, en, t) => {
      if (r <= 0 || n <= 0 || pv <= 0 || st < 1 || en < st || en > n || (t !== 0 && t !== 1)) return NUM();
      let s = 0; for (let p = st; p <= en; p++) s += FN.IPMT([r, p, n, pv, 0, t]);
      return s;
    }),
    CUMPRINC: fn(['n', 'i', 'n', 'i', 'i', 'i'], (r, n, pv, st, en, t) => {
      if (r <= 0 || n <= 0 || pv <= 0 || st < 1 || en < st || en > n || (t !== 0 && t !== 1)) return NUM();
      let s = 0; for (let p = st; p <= en; p++) s += FN.PPMT([r, p, n, pv, 0, t]);
      return s;
    }),
    DOLLARDE: fn(['n', 'i'], (x, fr) => {
      if (fr < 0) return NUM();
      if (fr === 0) return err(ERR.DIV0);
      const ip = Math.trunc(x); const dec = x - ip;
      return ip + (dec * 10 ** Math.ceil(Math.log10(fr))) / fr;
    }),
    DOLLARFR: fn(['n', 'i'], (x, fr) => {
      if (fr < 0) return NUM();
      if (fr === 0) return err(ERR.DIV0);
      const ip = Math.trunc(x); const dec = x - ip;
      return ip + (dec * fr) / 10 ** Math.ceil(Math.log10(fr));
    }),
    FVSCHEDULE: ([pr, sched]) => {
      let v = num(pr);
      if (isErr(v)) return v;
      for (const r of flat(sched)) { if (r == null) continue; const x = toNum(r); if (isErr(x)) return err(ERR.VALUE); v *= 1 + x; }
      return v;
    },
    ISPMT: fn(['n', 'n', 'n', 'n'], (r, per, n, pv) => (n === 0 ? err(ERR.DIV0) : pv * r * (per / n - 1))),
    MIRR: ([vals, fr, rr]) => {
      const f = num(fr); const rv = num(rr);
      if (firstErr(f, rv)) return firstErr(f, rv);
      const v = flat(vals).filter((x) => typeof x === 'number');
      const n = v.length;
      let pos = 0; let neg = 0;
      v.forEach((x, i) => { if (x > 0) pos += x / (1 + rv) ** i; else neg += x / (1 + f) ** i; });
      if (pos === 0 || neg === 0) return err(ERR.DIV0);
      return ((-pos * (1 + rv) ** n) / (neg * (1 + rv))) ** (1 / (n - 1)) - 1;
    },
    PDURATION: fn(['n', 'n', 'n'], (r, pv, fv) => (r <= 0 || pv <= 0 || fv <= 0 ? NUM() : (Math.log(fv) - Math.log(pv)) / Math.log(1 + r))),
    RRI: fn(['n', 'n', 'n'], (n, pv, fv) => (n <= 0 || pv === 0 ? NUM() : (fv / pv) ** (1 / n) - 1)),
    VDB: fn(['n', 'n', 'n', 'n', 'n', ['n', 2], ['b', false]], (c, s, l, st, en, fac, ns) => {
      if (c < 0 || s < 0 || l <= 0 || st < 0 || en < st || en > l || fac <= 0) return NUM();
      return vdb(c, s, l, st, en, fac, ns);
    }),
    AMORLINC: fn(['n', 'd', 'd', 'n', 'i', 'n', ['i', 0]], (cost, dp, fp, sal, per, rate, b) => {
      if (rate <= 0 || cost <= 0 || sal < 0 || per < 0 || !okBasis(b) || b === 2) return NUM();
      const one = cost * rate; const delta = cost - sal;
      const r0 = yearfrac(dp, fp, b) * rate * cost;
      const full = Math.floor((cost - sal - r0) / one);
      if (per === 0) return r0;
      if (per <= full) return one;
      if (per === full + 1) return delta - one * full - r0;
      return 0;
    }),
    AMORDEGRC: fn(['n', 'd', 'd', 'n', 'i', 'n', ['i', 0]], (cost, dp, fp, sal, per, rate, b) => {
      if (rate <= 0 || cost <= 0 || sal < 0 || per < 0 || !okBasis(b) || b === 2) return NUM();
      const use = 1 / rate;
      if (use >= 3 && use < 5) rate *= 1.5; else if (use >= 5 && use <= 6) rate *= 2; else if (use > 6) rate *= 2.5;
      const rnd = (x) => Math.round(x);
      let nr = rnd(yearfrac(dp, fp, b) * rate * cost);
      cost -= nr;
      let rest = cost - sal;
      for (let n = 0; n < per; n++) {
        nr = rnd(rate * cost);
        rest -= nr;
        if (rest < 0) return per - n === 0 || per - n === 1 ? rnd(cost * 0.5) : 0;
        cost -= nr;
      }
      return nr;
    }),
  };

  return { ...bond, ...other };
}
