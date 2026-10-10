// Библиотека функций, пакет B (этап 12.5): статистические распределения и тесты.
// Спецфункции: ln Γ (Ланцош), регуляризованные неполные гамма P(a,x)/Q(a,x) и бета I_x(a,b) (ряды и цепные
// дроби, точность ~1e-15), обратные — подбором (деление отрезка + Ньютон). Сверка с Excel — tests/formulas.test.js.
export function functions3(h) {
  const { ERR, err, isErr, toNum, toBool, flat, nums, firstErr, toArr, FN } = h;
  const num = (x, def) => (x == null && def !== undefined ? def : toNum(x));
  const bool = (x, def) => { if (x == null) return def; const b = toBool(x); return isErr(b) ? b : b; };
  const mean = (n) => n.reduce((s, x) => s + x, 0) / n.length;
  const NUM = () => err(ERR.NUM);

  // ---------- спецфункции ----------
  const LG = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  function gammaln(x) {
    if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - gammaln(1 - x);
    const z = x - 1;
    let a = LG[0];
    const t = z + 7.5;
    for (let i = 1; i < 9; i++) a += LG[i] / (z + i);
    return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
  }
  function gamma(x) {
    if (x === Math.floor(x)) { if (x <= 0) return NaN; if (x < 171) { let r = 1; for (let i = 2; i < x; i++) r *= i; return r; } }
    if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
    return Math.exp(gammaln(x));
  }
  // P(a, x) — регуляризованная нижняя неполная гамма
  function gammaP(a, x) {
    if (x <= 0) return 0;
    if (x < a + 1) {
      let sum = 1 / a; let term = sum;
      for (let n = 1; n < 1000; n++) { term *= x / (a + n); sum += term; if (Math.abs(term) < Math.abs(sum) * 1e-16) break; }
      return sum * Math.exp(-x + a * Math.log(x) - gammaln(a));
    }
    return 1 - gammaQcf(a, x);
  }
  function gammaQ(a, x) { return x < a + 1 ? 1 - gammaP(a, x) : gammaQcf(a, x); }
  function gammaQcf(a, x) {
    const FP = 1e-300;
    let b = x + 1 - a; let c = 1 / FP; let d = 1 / b; let hh = d;
    for (let i = 1; i < 1000; i++) {
      const an = -i * (i - a); b += 2;
      d = an * d + b; if (Math.abs(d) < FP) d = FP;
      c = b + an / c; if (Math.abs(c) < FP) c = FP;
      d = 1 / d; const del = d * c; hh *= del;
      if (Math.abs(del - 1) < 1e-16) break;
    }
    return Math.exp(-x + a * Math.log(x) - gammaln(a)) * hh;
  }
  // I_x(a, b) — регуляризованная неполная бета
  function betaI(x, a, b) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    const bt = Math.exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(1 - x));
    return x < (a + 1) / (a + b + 2) ? (bt * betacf(x, a, b)) / a : 1 - (bt * betacf(1 - x, b, a)) / b;
  }
  function betacf(x, a, b) {
    const FP = 1e-300;
    const qab = a + b; const qap = a + 1; const qam = a - 1;
    let c = 1; let d = 1 - (qab * x) / qap; if (Math.abs(d) < FP) d = FP; d = 1 / d;
    let hh = d;
    for (let m = 1; m < 1000; m++) {
      const m2 = 2 * m;
      let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < FP) d = FP; c = 1 + aa / c; if (Math.abs(c) < FP) c = FP; d = 1 / d; hh *= d * c;
      aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < FP) d = FP; c = 1 + aa / c; if (Math.abs(c) < FP) c = FP; d = 1 / d;
      const del = d * c; hh *= del;
      if (Math.abs(del - 1) < 1e-16) break;
    }
    return hh;
  }
  // обратная к возрастающей cdf на (lo, hi): деление отрезка, затем Ньютон (если есть плотность)
  function invert(cdf, p, lo, hi, pdf) {
    // расширить верхнюю границу, пока cdf(hi) < p
    let guard = 0;
    while (cdf(hi) < p && guard++ < 200) { lo = hi; hi *= 2; }
    let x = (lo + hi) / 2;
    for (let i = 0; i < 300; i++) {
      x = (lo + hi) / 2;
      const v = cdf(x);
      if (v < p) lo = x; else hi = x;
      if (hi - lo <= 1e-15 * Math.max(1, Math.abs(x))) break;
    }
    if (pdf) {
      for (let i = 0; i < 5; i++) {
        const d = pdf(x);
        if (!(d > 0)) break;
        const nx = x - (cdf(x) - p) / d;
        if (!isFinite(nx) || nx <= lo - (hi - lo) * 10) break;
        if (Math.abs(nx - x) < 1e-16 * Math.max(1, Math.abs(x))) { x = nx; break; }
        x = nx;
      }
    }
    return x;
  }
  const lnC = (n, k) => gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1);

  // распределения (x — уже числа)
  const tCdf = (t, df) => { const x = df / (df + t * t); const tail = 0.5 * betaI(x, df / 2, 0.5); return t >= 0 ? 1 - tail : tail; };
  const tPdf = (t, df) => Math.exp(gammaln((df + 1) / 2) - gammaln(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log(1 + (t * t) / df));
  const tInv = (p, df) => (p === 0.5 ? 0 : p < 0.5 ? -tInv(1 - p, df) : invert((t) => tCdf(t, df), p, 0, 10, (t) => tPdf(t, df)));
  const chiCdf = (x, k) => gammaP(k / 2, x / 2);
  const chiPdf = (x, k) => (x <= 0 ? 0 : Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.log(2) - gammaln(k / 2)));
  const fCdf = (x, a, b) => (x <= 0 ? 0 : betaI((a * x) / (a * x + b), a / 2, b / 2));
  const fPdf = (x, a, b) => (x <= 0 ? 0 : Math.exp(0.5 * (a * Math.log(a * x) + b * Math.log(b) - (a + b) * Math.log(a * x + b)) - Math.log(x) - (gammaln(a / 2) + gammaln(b / 2) - gammaln((a + b) / 2))));
  const gammaPdf = (x, a, b) => (x < 0 ? 0 : x === 0 ? (a === 1 ? 1 / b : a < 1 ? Infinity : 0) : Math.exp((a - 1) * Math.log(x) - x / b - gammaln(a) - a * Math.log(b)));
  const betaPdf = (x, a, b) => (x <= 0 || x >= 1 ? 0 : Math.exp((a - 1) * Math.log(x) + (b - 1) * Math.log(1 - x) - (gammaln(a) + gammaln(b) - gammaln(a + b))));
  const normCdf = (z) => FN['NORM.S.DIST']([z, true]);
  const normInv = (p) => FN['NORM.S.INV']([p]);
  const binomPmf = (k, n, p) => (k < 0 || k > n ? 0 : Math.exp(lnC(n, k) + (k ? k * Math.log(p) : 0) + (n - k ? (n - k) * Math.log(1 - p) : 0)));
  const binomCdf = (k, n, p) => (k < 0 ? 0 : k >= n ? 1 : betaI(1 - p, n - k, k + 1));
  const okP = (p) => p >= 0 && p <= 1;

  // аргументы-числа с проверкой; f(…числа) — значение или ошибка
  const fn = (defs, f) => (args) => {
    // 'b' — логический, 'i' — целое, 'o' — необязательное число (нет — null), иначе число (с умолчанием)
    const v = defs.map((d, i) => (d === 'b' ? bool(args[i], true) : d === 'i' ? (() => { const n = num(args[i]); return isErr(n) ? n : Math.trunc(n); })()
      : d === 'o' ? (args[i] == null ? null : num(args[i])) : num(args[i], d === undefined ? undefined : d)));
    const e = firstErr(...v);
    if (e) return e;
    const r = f(...v);
    return typeof r === 'number' && !isFinite(r) ? NUM() : r;
  };
  const sampleNums = (a) => { const n = nums([a]); return isErr(n) ? n : n; };

  const dist = {
    'BETA.DIST': fn([undefined, undefined, undefined, 'b', 0, 1], (x, a, b, cum, A, B) => {
      if (a <= 0 || b <= 0 || x < A || x > B || A === B) return NUM();
      const z = (x - A) / (B - A);
      return cum ? betaI(z, a, b) : betaPdf(z, a, b) / (B - A);
    }),
    BETADIST: fn([undefined, undefined, undefined, 0, 1], (x, a, b, A, B) => {
      if (a <= 0 || b <= 0 || x < A || x > B || A === B) return NUM();
      return betaI((x - A) / (B - A), a, b);
    }),
    'BETA.INV': fn([undefined, undefined, undefined, 0, 1], (p, a, b, A, B) => {
      if (p <= 0 || p > 1 || a <= 0 || b <= 0 || A >= B) return NUM();
      return A + (B - A) * invert((z) => betaI(z, a, b), p, 0, 1, (z) => betaPdf(z, a, b));
    }),
    'BINOM.DIST': fn(['i', 'i', undefined, 'b'], (k, n, p, cum) => {
      if (k < 0 || k > n || !okP(p)) return NUM();
      return cum ? binomCdf(k, n, p) : binomPmf(k, n, p);
    }),
    'BINOM.DIST.RANGE': fn(['i', undefined, 'i', 'o'], (n, p, s, s2) => {
      const e = s2 === undefined || s2 === null ? s : Math.trunc(s2);
      if (!okP(p) || s < 0 || s > n || e < s || e > n) return NUM();
      let r = 0; for (let k = s; k <= e; k++) r += binomPmf(k, n, p);
      return r;
    }),
    'BINOM.INV': fn(['i', undefined, undefined], (n, p, a) => {
      if (n < 0 || !okP(p) || a <= 0 || a >= 1) return NUM();
      let c = 0;
      for (let k = 0; k <= n; k++) { c += binomPmf(k, n, p); if (c >= a - 1e-14) return k; }
      return n;
    }),
    'CHISQ.DIST': fn([undefined, 'i', 'b'], (x, k, cum) => {
      if (x < 0 || k < 1 || k > 1e10) return NUM();
      return cum ? chiCdf(x, k) : chiPdf(x, k);
    }),
    'CHISQ.DIST.RT': fn([undefined, 'i'], (x, k) => (x < 0 || k < 1 ? NUM() : gammaQ(k / 2, x / 2))),
    'CHISQ.INV': fn([undefined, 'i'], (p, k) => (p < 0 || p >= 1 || k < 1 ? NUM() : p === 0 ? 0 : invert((x) => chiCdf(x, k), p, 0, Math.max(1, k), (x) => chiPdf(x, k)))),
    'CHISQ.INV.RT': fn([undefined, 'i'], (p, k) => (p <= 0 || p > 1 || k < 1 ? NUM() : p === 1 ? 0 : invert((x) => -gammaQ(k / 2, x / 2), -p, 0, Math.max(1, k), (x) => chiPdf(x, k)))),
    'CHISQ.TEST': ([act, exp]) => {
      const A = toArr(act); const E = toArr(exp);
      if (A.length !== E.length || A[0].length !== E[0].length) return err(ERR.NA);
      let x = 0;
      for (let i = 0; i < A.length; i++) {
        for (let j = 0; j < A[0].length; j++) {
          const a = A[i][j]; const e = E[i][j];
          if (typeof a !== 'number' || typeof e !== 'number') continue;
          if (e === 0) return err(ERR.DIV0);
          x += (a - e) ** 2 / e;
        }
      }
      const r = A.length; const c = A[0].length;
      const df = r > 1 && c > 1 ? (r - 1) * (c - 1) : Math.max(r, c) - 1;
      return df < 1 ? err(ERR.NA) : gammaQ(df / 2, x / 2);
    },
    'EXPON.DIST': fn([undefined, undefined, 'b'], (x, l, cum) => (x < 0 || l <= 0 ? NUM() : cum ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x))),
    'F.DIST': fn([undefined, 'i', 'i', 'b'], (x, a, b, cum) => (x < 0 || a < 1 || b < 1 ? NUM() : cum ? fCdf(x, a, b) : fPdf(x, a, b))),
    'F.DIST.RT': fn([undefined, 'i', 'i'], (x, a, b) => (x < 0 || a < 1 || b < 1 ? NUM() : x === 0 ? 1 : betaI(b / (b + a * x), b / 2, a / 2))),
    'F.INV': fn([undefined, 'i', 'i'], (p, a, b) => (p < 0 || p > 1 || a < 1 || b < 1 ? NUM() : p === 0 ? 0 : invert((x) => fCdf(x, a, b), p, 0, 10, (x) => fPdf(x, a, b)))),
    'F.INV.RT': fn([undefined, 'i', 'i'], (p, a, b) => (p <= 0 || p > 1 || a < 1 || b < 1 ? NUM() : p === 1 ? 0 : invert((x) => fCdf(x, a, b), 1 - p, 0, 10, (x) => fPdf(x, a, b)))),
    'F.TEST': ([a1, a2]) => {
      const x = sampleNums(a1); const y = sampleNums(a2);
      if (firstErr(x, y)) return firstErr(x, y);
      if (x.length < 2 || y.length < 2) return err(ERR.DIV0);
      const v1 = FN['VAR.S']([[x]]); const v2 = FN['VAR.S']([[y]]);
      if (!v1 || !v2) return err(ERR.DIV0);
      const p = fCdf(v1 / v2, x.length - 1, y.length - 1);
      return 2 * Math.min(p, 1 - p);
    },
    GAMMA: fn([undefined], (x) => (x <= 0 && x === Math.floor(x) ? NUM() : gamma(x))),
    'GAMMA.DIST': fn([undefined, undefined, undefined, 'b'], (x, a, b, cum) => (x < 0 || a <= 0 || b <= 0 ? NUM() : cum ? gammaP(a, x / b) : gammaPdf(x, a, b))),
    'GAMMA.INV': fn([undefined, undefined, undefined], (p, a, b) => (p < 0 || p > 1 || a <= 0 || b <= 0 ? NUM() : p === 0 ? 0 : invert((x) => gammaP(a, x / b), p, 0, a * b * 2 + 1, (x) => gammaPdf(x, a, b)))),
    GAMMALN: fn([undefined], (x) => (x <= 0 ? NUM() : gammaln(x))),
    'HYPGEOM.DIST': fn(['i', 'i', 'i', 'i', 'b'], (s, n, M, N, cum) => {
      if (s < 0 || n <= 0 || M <= 0 || N <= 0 || n > N || M > N || s > n || s > M || n - s > N - M) return NUM();
      const pmf = (k) => Math.exp(lnC(M, k) + lnC(N - M, n - k) - lnC(N, n));
      if (!cum) return pmf(s);
      let r = 0; for (let k = Math.max(0, n - (N - M)); k <= s; k++) r += pmf(k);
      return r;
    }),
    'LOGNORM.DIST': fn([undefined, undefined, undefined, 'b'], (x, m, s, cum) => {
      if (x <= 0 || s <= 0) return NUM();
      const z = (Math.log(x) - m) / s;
      return cum ? normCdf(z) : Math.exp(-0.5 * z * z) / (x * s * Math.sqrt(2 * Math.PI));
    }),
    'LOGNORM.INV': fn([undefined, undefined, undefined], (p, m, s) => (p <= 0 || p >= 1 || s <= 0 ? NUM() : Math.exp(m + s * normInv(p)))),
    'NEGBINOM.DIST': fn(['i', 'i', undefined, 'b'], (f, s, p, cum) => {
      if (f < 0 || s < 1 || p < 0 || p > 1) return NUM();
      return cum ? betaI(p, s, f + 1) : Math.exp(lnC(f + s - 1, s - 1) + s * Math.log(p) + (f ? f * Math.log(1 - p) : 0));
    }),
    'POISSON.DIST': fn(['i', undefined, 'b'], (x, l, cum) => {
      if (x < 0 || l < 0) return NUM();
      if (cum) return l === 0 ? 1 : gammaQ(x + 1, l);
      return l === 0 ? (x === 0 ? 1 : 0) : Math.exp(x * Math.log(l) - l - gammaln(x + 1));
    }),
    'T.DIST': fn([undefined, 'i', 'b'], (x, df, cum) => (df < 1 ? NUM() : cum ? tCdf(x, df) : tPdf(x, df))),
    'T.DIST.2T': fn([undefined, 'i'], (x, df) => (x < 0 || df < 1 ? NUM() : betaI(df / (df + x * x), df / 2, 0.5))),
    'T.DIST.RT': fn([undefined, 'i'], (x, df) => (df < 1 ? NUM() : 1 - tCdf(x, df))),
    'T.INV': fn([undefined, 'i'], (p, df) => (p <= 0 || p >= 1 || df < 1 ? NUM() : tInv(p, df))),
    'T.INV.2T': fn([undefined, 'i'], (p, df) => (p <= 0 || p > 1 || df < 1 ? NUM() : tInv(1 - p / 2, df))),
    'T.TEST': ([a1, a2, tails, type]) => {
      const t0 = num(tails); const ty = num(type);
      if (firstErr(t0, ty)) return firstErr(t0, ty);
      const tl = Math.trunc(t0); const tp = Math.trunc(ty);
      if (![1, 2].includes(tl) || ![1, 2, 3].includes(tp)) return NUM();
      let t; let df;
      if (tp === 1) {
        const x = flat(a1); const y = flat(a2);
        if (x.length !== y.length) return err(ERR.NA);
        const d = [];
        for (let i = 0; i < x.length; i++) if (typeof x[i] === 'number' && typeof y[i] === 'number') d.push(x[i] - y[i]);
        if (d.length < 2) return err(ERR.DIV0);
        const m = mean(d); const s = Math.sqrt(d.reduce((q, v) => q + (v - m) ** 2, 0) / (d.length - 1));
        t = m / (s / Math.sqrt(d.length)); df = d.length - 1;
      } else {
        const x = sampleNums(a1); const y = sampleNums(a2);
        if (firstErr(x, y)) return firstErr(x, y);
        if (x.length < 2 || y.length < 2) return err(ERR.DIV0);
        const v1 = FN['VAR.S']([[x]]); const v2 = FN['VAR.S']([[y]]);
        const n1 = x.length; const n2 = y.length;
        if (tp === 2) {
          const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
          t = (mean(x) - mean(y)) / Math.sqrt(sp * (1 / n1 + 1 / n2)); df = n1 + n2 - 2;
        } else {
          const q1 = v1 / n1; const q2 = v2 / n2;
          t = (mean(x) - mean(y)) / Math.sqrt(q1 + q2);
          df = (q1 + q2) ** 2 / (q1 * q1 / (n1 - 1) + q2 * q2 / (n2 - 1));
        }
      }
      const p2 = betaI(df / (df + t * t), df / 2, 0.5);
      return tl === 2 ? p2 : p2 / 2;
    },
    'WEIBULL.DIST': fn([undefined, undefined, undefined, 'b'], (x, a, b, cum) => {
      if (x < 0 || a <= 0 || b <= 0) return NUM();
      return cum ? 1 - Math.exp(-((x / b) ** a)) : (a / b ** a) * x ** (a - 1) * Math.exp(-((x / b) ** a));
    }),
    'Z.TEST': ([arr, x, sigma]) => {
      const n = sampleNums(arr); const v = num(x);
      if (firstErr(n, v)) return firstErr(n, v);
      if (!n.length) return err(ERR.NA);
      const s = sigma != null ? num(sigma) : Math.sqrt(FN['VAR.S']([[n]]));
      if (isErr(s)) return s;
      return 1 - normCdf((mean(n) - v) / (s / Math.sqrt(n.length)));
    },
    'CONFIDENCE.T': fn([undefined, undefined, 'i'], (a, s, n) => (a <= 0 || a >= 1 || s <= 0 || n < 2 ? (n === 1 ? err(ERR.DIV0) : NUM()) : (tInv(1 - a / 2, n - 1) * s) / Math.sqrt(n))),
  };
  // старые имена (совместимость)
  Object.assign(dist, {
    BINOMDIST: dist['BINOM.DIST'], CRITBINOM: dist['BINOM.INV'], CHIDIST: dist['CHISQ.DIST.RT'], CHIINV: dist['CHISQ.INV.RT'],
    CHITEST: dist['CHISQ.TEST'], EXPONDIST: dist['EXPON.DIST'], FDIST: dist['F.DIST.RT'], FINV: dist['F.INV.RT'], FTEST: dist['F.TEST'],
    GAMMADIST: dist['GAMMA.DIST'], GAMMAINV: dist['GAMMA.INV'], 'GAMMALN.PRECISE': dist.GAMMALN, BETAINV: dist['BETA.INV'],
    LOGINV: dist['LOGNORM.INV'], LOGNORMDIST: fn([undefined, undefined, undefined], (x, m, s) => (x <= 0 || s <= 0 ? NUM() : normCdf((Math.log(x) - m) / s))),
    HYPGEOMDIST: (a) => dist['HYPGEOM.DIST']([...a.slice(0, 4), false]),
    NEGBINOMDIST: (a) => dist['NEGBINOM.DIST']([...a.slice(0, 3), false]),
    POISSON: dist['POISSON.DIST'], WEIBULL: dist['WEIBULL.DIST'], ZTEST: dist['Z.TEST'], TTEST: dist['T.TEST'], TINV: dist['T.INV.2T'],
    TDIST: fn([undefined, 'i', 'i'], (x, df, tails) => (x < 0 || df < 1 || ![1, 2].includes(tails) ? NUM() : (tails === 2 ? 1 : 0.5) * betaI(df / (df + x * x), df / 2, 0.5))),
  });
  h.special = { gammaln, gammaP, gammaQ, betaI };
  return dist;
}
