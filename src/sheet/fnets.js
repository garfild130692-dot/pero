// Прогноз экспоненциальным сглаживанием (FORECAST.ETS, .CONFINT, .SEASONALITY, .STAT) — модель Хольта–Винтерса
// с аддитивными трендом и сезонностью (AAA, как у Excel). Своя реализация: начальные значения — по первым двум сезонам,
// параметры α, β, γ ∈ [0,001; 0,999] — минимум суммы квадратов ошибок прогноза на шаг (сетка + уточнение до 0,001),
// сезонность — пик автокорреляции ряда без линейного тренда. Подбор параметров у Excel свой (не документирован),
// поэтому прогноз совпадает с Excel приближённо (см. tests/formulas.test.js, ETS).
export function functionsEts(h) {
  const { ERR, err, isErr, toNum, toArr, FN } = h;
  const NUM = () => err(ERR.NUM);
  const VALUE = () => err(ERR.VALUE);
  const MAX_SEASON = 8760;

  // ---------- данные: сортировка, одинаковые даты, шаг, пропуски ----------
  const AGG = {
    1: (v) => v.reduce((a, b) => a + b, 0) / v.length, 2: (v) => v.length, 3: (v) => v.length,
    4: (v) => Math.max(...v), 5: (v) => { const s = [...v].sort((a, b) => a - b); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; },
    6: (v) => Math.min(...v), 7: (v) => v.reduce((a, b) => a + b, 0),
  };
  function prepare(values, timeline, completion, aggregation) {
    if (isErr(values) && !Array.isArray(values)) return values;
    if (isErr(timeline) && !Array.isArray(timeline)) return timeline;
    const ys = toArr(values).flat(); const ts = toArr(timeline).flat();
    if (ys.length !== ts.length) return err(ERR.NA);
    const fill = completion == null ? 1 : toNum(completion);
    const agg = aggregation == null ? 1 : Math.trunc(toNum(aggregation));
    if (isErr(fill)) return fill; if (isErr(agg)) return agg;
    if ((fill !== 0 && fill !== 1) || !AGG[agg]) return NUM();
    const pts = new Map();
    for (let i = 0; i < ts.length; i++) {
      const t = ts[i]; const y = ys[i];
      if (isErr(t)) return t;
      if (typeof t !== 'number') return VALUE();
      if (y == null || y === '') continue; // пустое значение — пропуск
      if (isErr(y)) return y;
      if (typeof y !== 'number') return VALUE();
      if (!pts.has(t)) pts.set(t, []);
      pts.get(t).push(y);
    }
    const xs = [...pts.keys()].sort((a, b) => a - b);
    if (xs.length < 3) return NUM();
    let step = Infinity;
    for (let i = 1; i < xs.length; i++) step = Math.min(step, xs[i] - xs[i - 1]);
    if (!(step > 0)) return NUM();
    const y = []; const x0 = xs[0]; let missing = 0;
    for (let i = 0; i < xs.length; i++) {
      const v = AGG[agg](pts.get(xs[i]));
      if (i > 0) {
        const gap = (xs[i] - xs[i - 1]) / step;
        if (Math.abs(gap - Math.round(gap)) > 1e-9) return NUM(); // шаг не постоянный
        const prev = y[y.length - 1];
        for (let k = 1; k < Math.round(gap); k++) { y.push(fill ? prev + ((v - prev) * k) / Math.round(gap) : 0); missing++; }
      }
      y.push(v);
    }
    if (missing > 0.3 * xs.length) return NUM(); // пропусков больше 30 %
    return { y, x0, step, xEnd: x0 + (y.length - 1) * step };
  }

  // ---------- сезонность ----------
  function detectSeason(y) {
    const n = y.length; const mx = (n - 1) / 2; const my = y.reduce((a, b) => a + b, 0) / n;
    let sxy = 0; let sxx = 0;
    for (let i = 0; i < n; i++) { sxy += (i - mx) * (y[i] - my); sxx += (i - mx) ** 2; }
    const b = sxy / sxx; const z = y.map((v, i) => v - my - b * (i - mx));
    const den = z.reduce((a, v) => a + v * v, 0);
    if (den === 0) return 1;
    const maxLag = Math.min(Math.floor(n / 2), MAX_SEASON);
    const r = [1];
    for (let L = 1; L <= maxLag + 1; L++) { let s = 0; for (let i = L; i < n; i++) s += z[i] * z[i - L]; r.push(s / den); }
    let best = 1; let bestR = 2 / Math.sqrt(n); // порог значимости
    for (let L = 2; L <= maxLag; L++) if (r[L] > r[L - 1] && r[L] >= r[L + 1] && r[L] > bestR) { best = L; bestR = r[L]; }
    return best;
  }

  // ---------- модель ----------
  // m — длина сезона (1 — без сезонности); возвращает сумму квадратов ошибок и конечное состояние
  function run(y, m, a, b, g, keep) {
    const n = y.length;
    let level; let trend; const season = new Array(m).fill(0);
    let start;
    if (m > 1) {
      const avg = (k) => { let s = 0; for (let j = 0; j < m; j++) s += y[k * m + j]; return s / m; };
      const periods = Math.floor(n / m);
      trend = (avg(1) - avg(0)) / m;
      for (let j = 0; j < m; j++) {
        let s = 0;
        for (let k = 0; k < periods; k++) s += y[k * m + j] - (avg(k) + (j - (m - 1) / 2) * trend);
        season[j] = s / periods;
      }
      level = avg(0) + ((m - 1) / 2) * trend; // уровень в конце первого сезона
      start = m;
    } else {
      level = y[0]; trend = y[1] - y[0]; start = 1;
    }
    let sse = 0; const fitted = keep ? new Array(n).fill(null) : null; let cnt = 0;
    for (let t = start; t < n; t++) {
      const s = m > 1 ? season[t % m] : 0;
      const f = level + trend + s;
      const e = y[t] - f;
      sse += e * e; cnt++;
      if (keep) fitted[t] = f;
      const l0 = level;
      level = a * (y[t] - s) + (1 - a) * (level + trend);
      trend = b * (level - l0) + (1 - b) * trend;
      if (m > 1) season[t % m] = g * (y[t] - level) + (1 - g) * s;
    }
    return { sse, cnt, level, trend, season, fitted, start };
  }
  const clamp = (v) => Math.min(0.999, Math.max(0.001, Math.round(v * 1000) / 1000));
  function fit(y, m) {
    const seasonal = m > 1;
    let best = { sse: Infinity, a: 0.5, b: 0.1, g: seasonal ? 0.1 : 0 };
    const grid = [0.001, 0.05, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 0.999];
    for (const a of grid) for (const b of grid) for (const g of seasonal ? grid : [0]) {
      const r = run(y, m, a, b, g);
      if (r.sse < best.sse) best = { sse: r.sse, a, b, g };
    }
    // уточнение по координатам с уменьшением шага до 0,001
    for (let step = 0.05; step >= 0.001 - 1e-12; step /= 2) {
      let improved = true;
      while (improved) {
        improved = false;
        for (const k of seasonal ? ['a', 'b', 'g'] : ['a', 'b']) {
          for (const d of [-step, step]) {
            const p = { ...best, [k]: clamp(best[k] + d) };
            if (p[k] === best[k]) continue;
            const r = run(y, m, p.a, p.b, p.g);
            if (r.sse < best.sse - 1e-12) { best = { ...p, sse: r.sse }; improved = true; }
          }
        }
      }
    }
    return { ...best, state: run(y, m, best.a, best.b, best.g, true) };
  }

  function model(values, timeline, seasonality, completion, aggregation) {
    const d = prepare(values, timeline, completion, aggregation);
    if (isErr(d)) return d;
    let m = seasonality == null ? 1 : toNum(seasonality);
    if (isErr(m)) return m;
    m = Math.trunc(m);
    if (m < 0 || m > MAX_SEASON) return NUM();
    if (m === 1) m = detectSeason(d.y); else if (m === 0) m = 1;
    if (m > 1 && d.y.length < 2 * m) return NUM();
    return { ...d, m, ...fit(d.y, m) };
  }
  // прогноз на h шагов после конца ряда (дробный h — интерполяция между соседними шагами)
  function forecastAt(md, target) {
    const { y, x0, step, xEnd, m, state } = md;
    const point = (k) => {
      if (k <= 0) { const i = y.length - 1 + k; return i >= state.start && state.fitted[i] != null ? state.fitted[i] : y[i]; }
      const s = m > 1 ? state.season[(y.length - 1 + k) % m] : 0;
      return state.level + k * state.trend + s;
    };
    if (target < x0) return NUM();
    const k = (target - xEnd) / step;
    const k0 = Math.floor(k); const fr = k - k0;
    const v0 = k <= 0 && fr === 0 ? (k0 === 0 ? y[y.length - 1] : point(k0)) : point(k0);
    return fr < 1e-9 ? v0 : v0 + fr * (point(k0 + 1) - v0);
  }
  const stats = (md) => {
    const { y, state } = md;
    let ae = 0; let se = 0; let pe = 0; let n = 0;
    for (let i = state.start; i < y.length; i++) {
      const f = state.fitted[i]; const e = y[i] - f;
      ae += Math.abs(e); se += e * e; pe += Math.abs(e) / ((Math.abs(f) + Math.abs(y[i])) / 2 || 1); n++;
    }
    let naive = 0;
    for (let i = 1; i < y.length; i++) naive += Math.abs(y[i] - y[i - 1]);
    naive /= y.length - 1;
    return { mae: ae / n, rmse: Math.sqrt(se / n), smape: pe / n, mase: naive ? ae / n / naive : 0, mse: se / n };
  };
  const lift = (target, f) => (Array.isArray(target) ? target.map((row) => row.map((t) => f(t))) : f(target));
  const tnum = (t) => { const v = toNum(t); return isErr(v) ? v : v; };

  return {
    'FORECAST.ETS': ([target, values, timeline, seas, compl, agg]) => {
      const md = model(values, timeline, seas, compl, agg);
      if (isErr(md)) return md;
      return lift(target, (t) => { const v = tnum(t); return isErr(v) ? v : forecastAt(md, v); });
    },
    'FORECAST.ETS.CONFINT': ([target, values, timeline, conf, seas, compl, agg]) => {
      const c = conf == null ? 0.95 : toNum(conf);
      if (isErr(c)) return c;
      if (c <= 0 || c >= 1) return NUM();
      const md = model(values, timeline, seas, compl, agg);
      if (isErr(md)) return md;
      const { mse } = stats(md);
      const z = FN['NORM.S.INV']([(1 + c) / 2]);
      // дисперсия прогноза на h шагов для модели AAA (Hyndman et al.): σ²·(1 + Σ c_j²), c_j = α(1 + jβ) + γ(1−α)·[j кратно m]
      return lift(target, (t) => {
        const v = tnum(t); if (isErr(v)) return v;
        if (v < md.xEnd) return NUM();
        const hSteps = Math.max(1, Math.ceil((v - md.xEnd) / md.step - 1e-9));
        let s = 1;
        for (let j = 1; j < hSteps; j++) {
          const cj = md.a * (1 + j * md.b) + (md.m > 1 && j % md.m === 0 ? md.g * (1 - md.a) : 0);
          s += cj * cj;
        }
        return z * Math.sqrt(mse * s);
      });
    },
    'FORECAST.ETS.SEASONALITY': ([values, timeline, compl, agg]) => {
      const md = model(values, timeline, 1, compl, agg);
      if (isErr(md)) return md;
      return md.m > 1 ? md.m : 0;
    },
    'FORECAST.ETS.STAT': ([values, timeline, type, seas, compl, agg]) => {
      const md = model(values, timeline, seas, compl, agg);
      if (isErr(md)) return md;
      const st = stats(md);
      return lift(type, (k) => {
        const n = Math.trunc(toNum(k));
        if (isErr(n)) return n;
        switch (n) {
          case 1: return md.a; case 2: return md.b; case 3: return md.m > 1 ? md.g : 0;
          case 4: return st.mase; case 5: return st.smape; case 6: return st.mae; case 7: return st.rmse; case 8: return md.step;
          default: return NUM();
        }
      });
    },
  };
}
