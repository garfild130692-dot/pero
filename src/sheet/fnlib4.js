// Библиотека функций, пакет C (этап 12.5): инженерные — комплексные числа (КОМПЛЕКСН, МНИМ.*), функции Бесселя,
// ПРЕОБР (единицы измерения). Комплексное число — текст «a+bi» (или «a+bj»), как в Excel.
export function functions4(h) {
  const { ERR, err, isErr, toNum, toText, flat, firstErr } = h;
  const num = (x, def) => (x == null && def !== undefined ? def : toNum(x));
  const NUM = () => err(ERR.NUM);

  // ---------- комплексные числа ----------
  // текст → { re, im, s } (s — суффикс i/j); число — вещественное
  function parse(x) {
    if (isErr(x)) return x;
    if (typeof x === 'number') return { re: x, im: 0, s: 'i' };
    if (typeof x === 'boolean') return err(ERR.VALUE);
    let t = toText(x).trim();
    if (t === '') return { re: 0, im: 0, s: 'i' };
    t = t.replace(/,/g, '.');
    const n = '(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][-+]?\\d+)?';
    let m = new RegExp(`^([-+]?${n})?(?:([-+])(${n})?([ij]))?$`).exec(t);
    if (m && (m[1] !== undefined || m[2] !== undefined)) {
      if (m[1] !== undefined && m[2] === undefined) return { re: +m[1], im: 0, s: 'i' };
      return { re: m[1] !== undefined ? +m[1] : 0, im: (m[2] === '-' ? -1 : 1) * (m[3] !== undefined ? +m[3] : 1), s: m[4] };
    }
    m = new RegExp(`^([-+]?)(${n})?([ij])$`).exec(t);
    if (m) return { re: 0, im: (m[1] === '-' ? -1 : 1) * (m[2] !== undefined ? +m[2] : 1), s: m[3] };
    return NUM();
  }
  // число → текст как в Excel: до 15 значащих цифр, десятичная запятая
  const part = (v) => { let s = String(+v.toPrecision(15)); if (/e/.test(s)) s = s.replace('e', 'E'); return s.replace('.', ','); };
  function fmt(c, s = 'i') {
    let { re, im } = c;
    if (Math.abs(re) < 1e-300) re = 0;
    if (Math.abs(im) < 1e-300) im = 0;
    if (!isFinite(re) || !isFinite(im)) return NUM();
    if (im === 0) return part(re);
    const imStr = im === 1 ? '' : im === -1 ? '-' : part(im);
    if (re === 0) return `${imStr}${s}`;
    return `${part(re)}${im > 0 ? '+' : ''}${imStr}${s}`;
  }
  const C = {
    add: (a, b) => ({ re: a.re + b.re, im: a.im + b.im }),
    sub: (a, b) => ({ re: a.re - b.re, im: a.im - b.im }),
    mul: (a, b) => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re }),
    div: (a, b) => { const d = b.re * b.re + b.im * b.im; return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d }; },
    abs: (a) => Math.hypot(a.re, a.im),
    arg: (a) => Math.atan2(a.im, a.re),
    exp: (a) => ({ re: Math.exp(a.re) * Math.cos(a.im), im: Math.exp(a.re) * Math.sin(a.im) }),
    ln: (a) => ({ re: Math.log(Math.hypot(a.re, a.im)), im: Math.atan2(a.im, a.re) }),
    sin: (a) => ({ re: Math.sin(a.re) * Math.cosh(a.im), im: Math.cos(a.re) * Math.sinh(a.im) }),
    cos: (a) => ({ re: Math.cos(a.re) * Math.cosh(a.im), im: -Math.sin(a.re) * Math.sinh(a.im) }),
    sinh: (a) => ({ re: Math.sinh(a.re) * Math.cos(a.im), im: Math.cosh(a.re) * Math.sin(a.im) }),
    cosh: (a) => ({ re: Math.cosh(a.re) * Math.cos(a.im), im: Math.sinh(a.re) * Math.sin(a.im) }),
  };
  const one = { re: 1, im: 0 };
  // функция одного комплексного аргумента
  const im1 = (f, check) => ([x]) => {
    const a = parse(x);
    if (isErr(a)) return a;
    if (check && !check(a)) return NUM();
    const r = f(a);
    return typeof r === 'number' ? r : fmt(r, a.s);
  };
  const nonZero = (a) => a.re !== 0 || a.im !== 0;
  // суффикс результата нескольких аргументов: смешение i и j — #ЗНАЧ!
  const suffixOf = (list) => { const s = new Set(list.filter((c) => c.im !== 0 || c.explicit).map((c) => c.s)); return s.size > 1 ? null : [...s][0] || 'i'; };
  const imMany = (f, init) => (args) => {
    const list = [];
    for (const v of args.flatMap(flat)) { if (v == null) continue; const c = parse(v); if (isErr(c)) return c; list.push(c); }
    const s = suffixOf(list);
    if (!s) return err(ERR.VALUE);
    return fmt(list.reduce(f, init), s);
  };
  const complex = {
    COMPLEX: ([r, i, suf]) => {
      const re = num(r); const imv = num(i);
      if (firstErr(re, imv)) return firstErr(re, imv);
      const s = suf == null || suf === '' ? 'i' : toText(suf);
      if (s !== 'i' && s !== 'j') return err(ERR.VALUE);
      return fmt({ re, im: imv }, s);
    },
    IMREAL: im1((a) => a.re), IMAGINARY: im1((a) => a.im), IMABS: im1((a) => C.abs(a)),
    IMARGUMENT: im1((a) => C.arg(a), nonZero),
    IMCONJUGATE: im1((a) => ({ re: a.re, im: -a.im })),
    IMEXP: im1(C.exp), IMLN: im1(C.ln, nonZero),
    IMLOG10: im1((a) => { const l = C.ln(a); return { re: l.re / Math.LN10, im: l.im / Math.LN10 }; }, nonZero),
    IMLOG2: im1((a) => { const l = C.ln(a); return { re: l.re / Math.LN2, im: l.im / Math.LN2 }; }, nonZero),
    IMSQRT: im1((a) => { const r = Math.sqrt(C.abs(a)); const t = C.arg(a) / 2; return { re: r * Math.cos(t), im: r * Math.sin(t) }; }),
    IMSIN: im1(C.sin), IMCOS: im1(C.cos), IMSINH: im1(C.sinh), IMCOSH: im1(C.cosh),
    IMTAN: im1((a) => C.div(C.sin(a), C.cos(a))),
    IMCOT: im1((a) => C.div(C.cos(a), C.sin(a)), nonZero),
    IMSEC: im1((a) => C.div(one, C.cos(a))),
    IMCSC: im1((a) => C.div(one, C.sin(a)), nonZero),
    IMSECH: im1((a) => C.div(one, C.cosh(a))),
    IMCSCH: im1((a) => C.div(one, C.sinh(a)), nonZero),
    IMSUM: imMany(C.add, { re: 0, im: 0 }),
    IMPRODUCT: imMany(C.mul, one),
    IMSUB: ([x, y]) => { const a = parse(x); const b = parse(y); if (firstErr(a, b)) return firstErr(a, b); const s = suffixOf([a, b]); return s ? fmt(C.sub(a, b), s) : err(ERR.VALUE); },
    IMDIV: ([x, y]) => {
      const a = parse(x); const b = parse(y);
      if (firstErr(a, b)) return firstErr(a, b);
      if (!nonZero(b)) return NUM();
      const s = suffixOf([a, b]);
      return s ? fmt(C.div(a, b), s) : err(ERR.VALUE);
    },
    IMPOWER: ([x, p]) => {
      const a = parse(x); const n = num(p);
      if (firstErr(a, n)) return firstErr(a, n);
      if (!nonZero(a)) return n > 0 ? '0' : NUM();
      const r = C.abs(a) ** n; const t = C.arg(a) * n;
      return fmt({ re: r * Math.cos(t), im: r * Math.sin(t) }, a.s);
    },
  };

  // ---------- функции Бесселя (полиномиальные приближения, как в классических библиотеках) ----------
  function j0(x) {
    const ax = Math.abs(x);
    if (ax < 8) {
      const y = x * x;
      const a = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))));
      const b = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y))));
      return a / b;
    }
    const z = 8 / ax; const y = z * z; const xx = ax - 0.785398164;
    const a = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
    const b = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)));
    return Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a - z * Math.sin(xx) * b);
  }
  function j1(x) {
    const ax = Math.abs(x);
    if (ax < 8) {
      const y = x * x;
      const a = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.48260 + y * -30.16036606)))));
      const b = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y))));
      return a / b;
    }
    const z = 8 / ax; const y = z * z; const xx = ax - 2.356194491;
    const a = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)));
    const b = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    const r = Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a - z * Math.sin(xx) * b);
    return x < 0 ? -r : r;
  }
  function jn(n, x) {
    if (n === 0) return j0(x);
    if (n === 1) return j1(x);
    const ax = Math.abs(x);
    if (ax === 0) return 0;
    let ans;
    if (ax > n) {
      const tox = 2 / ax; let bjm = j0(ax); let bj = j1(ax);
      for (let j = 1; j < n; j++) { const bjp = j * tox * bj - bjm; bjm = bj; bj = bjp; }
      ans = bj;
    } else {
      const tox = 2 / ax; const m = 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2);
      let jsum = false; let bjp = 0; let sum = 0; let bj = 1; ans = 0;
      for (let j = m; j > 0; j--) {
        const bjm = j * tox * bj - bjp; bjp = bj; bj = bjm;
        if (Math.abs(bj) > 1e10) { bj *= 1e-10; bjp *= 1e-10; ans *= 1e-10; sum *= 1e-10; }
        if (jsum) sum += bj;
        jsum = !jsum;
        if (j === n) ans = bjp;
      }
      sum = 2 * sum - bj;
      ans /= sum;
    }
    return x < 0 && n & 1 ? -ans : ans;
  }
  function y0(x) {
    if (x < 8) {
      const y = x * x;
      const a = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
      const b = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.26470 + y * (226.1030244 + y))));
      return a / b + 0.636619772 * j0(x) * Math.log(x);
    }
    const z = 8 / x; const y = z * z; const xx = x - 0.785398164;
    const a = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
    const b = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934945152e-7)));
    return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a + z * Math.cos(xx) * b);
  }
  function y1(x) {
    if (x < 8) {
      const y = x * x;
      const a = x * (-0.4900604943e13 + y * (0.1275274390e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
      const b = 0.2499580570e14 + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.1020426050e6 + y * (0.3549632885e3 + y)))));
      return a / b + 0.636619772 * (j1(x) * Math.log(x) - 1 / x);
    }
    const z = 8 / x; const y = z * z; const xx = x - 2.356194491;
    const a = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)));
    const b = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
    return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a + z * Math.cos(xx) * b);
  }
  function yn(n, x) {
    if (n === 0) return y0(x);
    if (n === 1) return y1(x);
    const tox = 2 / x; let by = y1(x); let bym = y0(x);
    for (let j = 1; j < n; j++) { const byp = j * tox * by - bym; bym = by; by = byp; }
    return by;
  }
  function i0(x) {
    const ax = Math.abs(x);
    if (ax < 3.75) { const y = (x / 3.75) ** 2; return 1 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2))))); }
    const y = 3.75 / ax;
    return (Math.exp(ax) / Math.sqrt(ax)) * (0.39894228 + y * (0.1328592e-1 + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2
      + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))));
  }
  function i1(x) {
    const ax = Math.abs(x); let ans;
    if (ax < 3.75) { const y = (x / 3.75) ** 2; ans = ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3)))))); }
    else {
      const y = 3.75 / ax;
      ans = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2));
      ans = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * ans))));
      ans *= Math.exp(ax) / Math.sqrt(ax);
    }
    return x < 0 ? -ans : ans;
  }
  function inn(n, x) {
    if (n === 0) return i0(x);
    if (n === 1) return i1(x);
    if (x === 0) return 0;
    const tox = 2 / Math.abs(x); let bip = 0; let bi = 1; let ans = 0;
    for (let j = 2 * (n + Math.floor(Math.sqrt(40 * n))); j > 0; j--) {
      const bim = bip + j * tox * bi; bip = bi; bi = bim;
      if (Math.abs(bi) > 1e10) { ans *= 1e-10; bi *= 1e-10; bip *= 1e-10; }
      if (j === n) ans = bip;
    }
    ans *= i0(x) / bi;
    return x < 0 && n & 1 ? -ans : ans;
  }
  function k0(x) {
    if (x <= 2) { const y = (x * x) / 4; return -Math.log(x / 2) * i0(x) + (-0.57721566 + y * (0.42278420 + y * (0.23069756 + y * (0.3488590e-1 + y * (0.262698e-2 + y * (0.10750e-3 + y * 0.74e-5)))))); }
    const y = 2 / x;
    return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.251540e-2 + y * 0.53208e-3))))));
  }
  function k1(x) {
    if (x <= 2) { const y = (x * x) / 4; return Math.log(x / 2) * i1(x) + (1 / x) * (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4)))))); }
    const y = 2 / x;
    return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))));
  }
  function kn(n, x) {
    if (n === 0) return k0(x);
    if (n === 1) return k1(x);
    const tox = 2 / x; let bkm = k0(x); let bk = k1(x);
    for (let j = 1; j < n; j++) { const bkp = bkm + j * tox * bk; bkm = bk; bk = bkp; }
    return bk;
  }
  const bessel = (f, pos) => ([x, n]) => {
    const v = num(x); const k = num(n);
    if (firstErr(v, k)) return firstErr(v, k);
    const o = Math.trunc(k);
    if (o < 0 || (pos && v <= 0)) return NUM();
    const r = f(o, v);
    return isFinite(r) ? r : NUM();
  };
  const besselFns = { BESSELJ: bessel(jn), BESSELY: bessel(yn, true), BESSELI: bessel(inn), BESSELK: bessel(kn, true) };

  // ---------- ПРЕОБР: единицы измерения ----------
  // [категория, множитель к базовой единице категории, допускаются ли приставки]
  const U = {};
  const add = (cat, list) => { for (const [names, f, pre] of list) for (const n of names.split(' ')) U[n] = { cat, f, pre: pre !== false }; };
  add('mass', [['g', 1], ['sg', 14593.902937206364, false], ['lbm', 453.59237, false], ['u', 1.66053906660e-24], ['ozm', 28.349523125, false],
    ['grain', 0.06479891, false], ['cwt shweight', 45359.237, false], ['uk_cwt lcwt hweight', 50802.34544, false], ['stone', 6350.29318, false],
    ['ton', 907184.74, false], ['uk_ton LTON brton', 1016046.9088, false]]);
  add('dist', [['m', 1], ['mi', 1609.344, false], ['Nmi', 1852, false], ['in', 0.0254, false], ['ft', 0.3048, false], ['yd', 0.9144, false],
    ['ang', 1e-10], ['ell', 1.143, false], ['ly', 9460730472580800], ['parsec pc', 30856775814671916], ['Pica', 0.00035277777777777776, false],
    ['pica', 0.004233333333333333, false], ['survey_mi', 1609.3472186944373, false]]);
  add('time', [['yr', 31557600], ['day d', 86400], ['hr', 3600], ['mn min', 60], ['sec s', 1]]);
  add('press', [['Pa p', 1], ['atm at', 101325], ['mmHg', 133.322], ['psi', 6894.757293168361, false], ['Torr', 133.32236842105263, false]]);
  add('force', [['N', 1], ['dyn dy', 1e-5], ['lbf', 4.4482216152605, false], ['pond', 0.00980665]]);
  add('energy', [['J', 1], ['e', 1e-7], ['c', 4.184], ['cal', 4.1868], ['eV ev', 1.602176634e-19], ['HPh hh', 2684519.537696173, false],
    ['Wh wh', 3600], ['flb', 0.04214011009380476, false], ['BTU btu', 1055.05585262, false]]);
  add('power', [['HP h', 745.6998715822702, false], ['PS', 735.49875, false], ['W w', 1]]);
  add('magnet', [['T', 1], ['ga', 1e-4]]);
  add('volume', [['tsp', 4.92892159375e-6, false], ['tspm', 5e-6, false], ['tbs', 1.478676478125e-5, false], ['oz', 2.95735295625e-5, false],
    ['cup', 2.365882365e-4, false], ['pt us_pt', 4.73176473e-4, false], ['uk_pt', 5.6826125e-4, false], ['qt', 9.46352946e-4, false],
    ['uk_qt', 1.1365225e-3, false], ['gal', 3.785411784e-3, false], ['uk_gal', 4.54609e-3, false], ['l L lt', 1e-3], ['ang3 ang^3', 1e-30],
    ['barrel', 0.158987294928, false], ['bushel', 0.03523907016688, false], ['ft3 ft^3', 0.028316846592, false], ['in3 in^3', 1.6387064e-5, false],
    ['ly3 ly^3', 8.467866646237125e47, false], ['m3 m^3', 1], ['mi3 mi^3', 4168181825.440579584, false], ['yd3 yd^3', 0.764554857984, false],
    ['Nmi3 Nmi^3', 6352182208, false], ['Pica3 Picapt3 Pica^3 Picapt^3', 4.390300714449e-11, false], ['GRT regton', 2.8316846592, false],
    ['MTON', 1.13267386368, false]]);
  add('area', [['uk_acre', 4046.8564224, false], ['us_acre', 4046.872609874252, false], ['ang2 ang^2', 1e-20], ['ar', 100], ['ft2 ft^2', 0.09290304, false],
    ['ha', 10000, false], ['in2 in^2', 0.00064516, false], ['ly2 ly^2', 8.95054210748189e31, false], ['m2 m^2', 1], ['Morgen', 2500, false],
    ['mi2 mi^2', 2589988.110336, false], ['Nmi2 Nmi^2', 3429904, false], ['Pica2 Pica^2', 1.244521605e-7, false], ['yd2 yd^2', 0.83612736, false]]);
  add('info', [['bit', 1], ['byte', 8]]);
  add('speed', [['admkn', 0.514773333333333, false], ['kn', 0.514444444444444, false], ['m/h m/hr', 1 / 3600], ['m/s m/sec', 1], ['mph', 0.44704, false]]);
  const PREFIX = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, e: 10, d: 0.1, c: 0.01, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 };
  const BIN = { ki: 2 ** 10, Mi: 2 ** 20, Gi: 2 ** 30, Ti: 2 ** 40, Pi: 2 ** 50, Ei: 2 ** 60, Zi: 2 ** 70, Yi: 2 ** 80 };
  const TEMP = { C: 'C', cel: 'C', F: 'F', fah: 'F', K: 'K', kel: 'K', Rank: 'R', Reau: 'Re' };
  function unit(name) {
    if (U[name]) return { ...U[name], f: U[name].f };
    if (TEMP[name]) return { cat: 'temp', t: TEMP[name] };
    // приставка + единица: квадратные/кубические единицы — приставка в степени
    for (const [p, v] of [...Object.entries(BIN), ...Object.entries(PREFIX)].sort((a, b) => b[0].length - a[0].length)) {
      if (!name.startsWith(p)) continue;
      const base = U[name.slice(p.length)];
      if (!base || !base.pre) continue;
      if (BIN[p] && base.cat !== 'info') continue;
      const pow = /\^?3$/.test(name.slice(p.length)) && base.cat === 'volume' ? 3 : /\^?2$/.test(name.slice(p.length)) && base.cat === 'area' ? 2 : 1;
      return { cat: base.cat, f: base.f * v ** pow };
    }
    if (name === 'K' || name === 'kel') return { cat: 'temp', t: 'K' };
    return null;
  }
  const toK = (v, t) => ({ C: v + 273.15, F: (v - 32) * 5 / 9 + 273.15, K: v, R: v * 5 / 9, Re: v * 1.25 + 273.15 })[t];
  const fromK = (k, t) => ({ C: k - 273.15, F: (k - 273.15) * 9 / 5 + 32, K: k, R: k * 9 / 5, Re: (k - 273.15) * 0.8 })[t];
  const convert = {
    CONVERT: ([x, from, to]) => {
      const v = num(x);
      if (firstErr(v, from, to)) return firstErr(v, from, to);
      const a = unit(toText(from)); const b = unit(toText(to));
      if (!a || !b || a.cat !== b.cat) return err(ERR.NA);
      if (a.cat === 'temp') return fromK(toK(v, a.t), b.t);
      return (v * a.f) / b.f;
    },
  };

  return { ...complex, ...besselFns, ...convert };
}
