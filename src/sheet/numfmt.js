// Числовые форматы Excel (коды как в файле: 0.00, #,##0, dd.mm.yyyy…) → текст в русском виде (запятая, пробел между
// разрядами), как показывает русский Excel. Сверено с Excel: tests/numfmt.test.js (Excel открывает книгу с кодами и
// возвращает текст ячеек). formatCode(v, code) → { text, color }.
//
// Секции «;»: одна — для всех; две — ≥0 и <0 (минус не ставится); три — >0, <0, 0; четвёртая — текст. Условия [>=100].
// Цвет [Красный]/[Red]/[Color5]; [$₽-419] — символ валюты; [$-419] — язык; [h] [mm] [ss] — прошедшее время;
// "текст" и \x — как есть; _x — пробел; *x — заполнение (не выводится); , — разряды или ÷1000 (после цифр); % — ×100;
// E+00 — экспонента; ?/? и ?/8 — дроби; ? — цифра или пробел; # — цифра без незначащих нулей; 0 — цифра.

const DEC = ','; const GROUP = '\u00a0'; const DATE_SEP = '.'; // разряды — неразрывный пробел, как в русском Excel
const MON = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
const MONTH = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
const DAY = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
const DAY_SHORT = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
const COLORS = {
  black: '#000000', черный: '#000000', чёрный: '#000000', blue: '#0000ff', синий: '#0000ff', cyan: '#00ffff', голубой: '#00ffff',
  green: '#00ff00', зеленый: '#00ff00', зелёный: '#00ff00', magenta: '#ff00ff', фиолетовый: '#ff00ff', red: '#ff0000', красный: '#ff0000',
  white: '#ffffff', белый: '#ffffff', yellow: '#ffff00', желтый: '#ffff00', жёлтый: '#ffff00',
};
const PALETTE = ['#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff', '#800000', '#008000', '#000080', '#808000',
  '#800080', '#008080', '#c0c0c0', '#808080', '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff'];

// ---------- общий формат ----------
// «Общий»: до 11 знаков (как в ячейке стандартной ширины), без лишних нулей
export function general(v) {
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const a = Math.abs(v);
  if (a >= 1e11 || (a < 1e-9 && a > 0)) {
    let s = v.toExponential(5).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    s = `${m}E${e[0] === '-' ? '-' : '+'}${e.replace(/^[+-]/, '').padStart(2, '0')}`;
    return s.replace('.', DEC);
  }
  // 10 знаков после округления до 15 значащих и не длиннее 11 символов
  const intLen = Math.max(1, Math.floor(Math.log10(a || 1)) + 1);
  const dec = Math.max(0, 10 - intLen - (v < 0 ? 1 : 0));
  let s = String(Number(roundDec(v, Math.min(dec, 10))));
  if (/e/.test(s)) s = v.toFixed(Math.min(dec, 10)).replace(/\.?0+$/, '');
  return s.replace('.', DEC);
}

// Округление «как в Excel»: по 15 значащим цифрам, половина — от нуля
function roundDec(x, d) {
  const s = Number(Math.abs(x).toPrecision(15));
  const r = Math.round(Number((s * 10 ** d).toPrecision(15))) / 10 ** d;
  return x < 0 ? -r : r;
}

// ---------- разбор кода ----------
const cache = new Map();
function splitSections(code) {
  const out = []; let cur = ''; let q = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (q) { cur += ch; if (ch === '"') q = false; continue; }
    if (ch === '"') { q = true; cur += ch; continue; }
    if (ch === '\\' || ch === '_' || ch === '*') { cur += ch + (code[i + 1] || ''); i++; continue; }
    if (ch === '[') { const j = code.indexOf(']', i); cur += code.slice(i, j < 0 ? code.length : j + 1); i = j < 0 ? code.length : j; continue; }
    if (ch === ';') { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}
// раздел → { type: 'number'|'date'|'text'|'general', tokens, color, cond }
function parseSection(src) {
  const tokens = []; let color = null; let cond = null; let hasDate = false; let hasAt = false; let general = false;
  const lit = (s) => { const last = tokens[tokens.length - 1]; if (last && last.t === 'lit') last.s += s; else tokens.push({ t: 'lit', s }); };
  let i = 0;
  while (i < src.length) {
    const ch = src[i]; const low = src.slice(i).toLowerCase();
    if (ch === '"') { const j = src.indexOf('"', i + 1); lit(src.slice(i + 1, j < 0 ? src.length : j)); i = j < 0 ? src.length : j + 1; continue; }
    if (ch === '\\') { lit(src[i + 1] || ''); i += 2; continue; }
    if (ch === '_') { tokens.push({ t: 'space' }); i += 2; continue; }
    if (ch === '*') { tokens.push({ t: 'fill', ch: src[i + 1] || ' ' }); i += 2; continue; }
    if (ch === '[') {
      const j = src.indexOf(']', i); const inner = src.slice(i + 1, j < 0 ? src.length : j); i = j < 0 ? src.length : j + 1;
      const il = inner.toLowerCase();
      if (/^(h+|m+|s+)$/.test(il)) { tokens.push({ t: 'elapsed', unit: il[0], len: il.length }); hasDate = true; continue; }
      const cm = /^(<=|>=|<>|<|>|=)\s*(-?[\d.,]+)$/.exec(inner);
      if (cm) { cond = { op: cm[1], val: +cm[2].replace(',', '.') }; continue; }
      if (inner.startsWith('$')) { const sym = inner.slice(1).split('-')[0]; if (sym) lit(sym); continue; }
      const cn = /^(?:color|цвет)\s*(\d+)$/i.exec(inner);
      if (cn) { color = PALETTE[+cn[1] - 1] || null; continue; }
      if (COLORS[il]) { color = COLORS[il]; continue; }
      continue; // прочие служебные ([DBNum1] и т. п.) — без вывода
    }
    if (low.startsWith('general') || low.startsWith('основной')) { tokens.push({ t: 'general' }); general = true; i += low.startsWith('general') ? 7 : 8; continue; }
    if (low.startsWith('am/pm') || low.startsWith('a/p')) { const n = low.startsWith('am/pm') ? 5 : 3; tokens.push({ t: 'ampm', s: src.slice(i, i + n) }); hasDate = true; i += n; continue; }
    if (/[yдгmмdhчsс]/i.test(ch) && /[ymdhs]/i.test(ch)) {
      let j = i; while (j < src.length && src[j].toLowerCase() === ch.toLowerCase()) j++;
      tokens.push({ t: 'date', c: ch.toLowerCase(), len: j - i }); hasDate = true; i = j; continue;
    }
    if (ch === 'e' || ch === 'E') {
      if (src[i + 1] === '+' || src[i + 1] === '-') { tokens.push({ t: 'exp', sign: src[i + 1] }); i += 2; continue; }
    }
    if (ch === '0' || ch === '#' || ch === '?') { tokens.push({ t: 'digit', ch }); i++; continue; }
    if (ch === '.') { tokens.push({ t: 'dot' }); i++; continue; }
    if (ch === ',') { tokens.push({ t: 'comma' }); i++; continue; }
    if (ch === '%') { tokens.push({ t: 'pct' }); i++; continue; }
    if (ch === '/') { tokens.push({ t: 'slash' }); i++; continue; }
    if (ch === '@') { tokens.push({ t: 'at' }); hasAt = true; i++; continue; }
    if (/[1-9]/.test(ch)) { tokens.push({ t: 'num', s: ch }); i++; continue; } // знаменатель дроби вида ?/8
    lit(ch); i++;
  }
  const type = hasDate ? 'date' : general ? 'general' : hasAt && !tokens.some((t) => t.t === 'digit') ? 'text' : 'number';
  return { type, tokens, color, cond };
}
function parseCode(code) {
  let p = cache.get(code);
  if (!p) {
    p = splitSections(code).map(parseSection);
    if (cache.size > 2000) cache.clear();
    cache.set(code, p);
  }
  return p;
}
export function isDateCode(code) {
  const p = parseCode(String(code || 'General'));
  return p[0].type === 'date';
}

// ---------- числа ----------
function bestFraction(x, maxDen) {
  let best = [Math.round(x), 1]; let err = Math.abs(x - best[0]);
  for (let d = 2; d <= maxDen; d++) {
    const n = Math.round(x * d);
    const e = Math.abs(x - n / d);
    if (e < err - 1e-12) { best = [n, d]; err = e; if (!e) break; }
  }
  return best;
}
// заполнить шаблон целых цифр (справа налево; лишние цифры — в самый левый знак)
function fillInt(intTokens, digits, grouping) {
  const ph = intTokens.filter((t) => t.t === 'digit');
  let ds = digits === '0' && ph.length && ph.every((t) => t.ch !== '0') ? '' : digits;
  if (grouping) {
    // группировка по всей целой части (с учётом обязательных нулей)
    const zeros = ph.filter((t) => t.ch === '0').length;
    ds = ds.padStart(zeros, '0');
    const grouped = ds.replace(/\B(?=(\d{3})+(?!\d))/g, GROUP);
    const pre = intTokens.slice(0, intTokens.findIndex((t) => t.t === 'digit')).map((t) => (t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join('');
    return pre + grouped;
  }
  let out = ''; let k = ds.length - 1; let phLeft = ph.length;
  for (let i = intTokens.length - 1; i >= 0; i--) {
    const t = intTokens[i];
    if (t.t !== 'digit') { out = (t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '') + out; continue; }
    phLeft--;
    if (phLeft === 0) { out = (k >= 0 ? ds.slice(0, k + 1) : t.ch === '0' ? '0' : t.ch === '?' ? ' ' : '') + out; k = -1; }
    else if (k >= 0) { out = ds[k] + out; k--; }
    else out = (t.ch === '0' ? '0' : t.ch === '?' ? ' ' : '') + out;
  }
  return out;
}
function fillDec(decTokens, digits) {
  let out = ''; let k = 0;
  // незначащие нули в конце у # не выводятся, у ? — пробел
  const ph = decTokens.filter((t) => t.t === 'digit');
  let lastSig = -1;
  for (let i = 0; i < ph.length; i++) if (digits[i] && digits[i] !== '0') lastSig = i;
  for (const t of decTokens) {
    if (t.t !== 'digit') { out += t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : ''; continue; }
    const d = digits[k] || '0';
    if (k > lastSig && t.ch === '#') { /* пропуск */ } else if (k > lastSig && t.ch === '?') out += ' ';
    else out += d;
    k++;
  }
  return out;
}
function formatNumberSection(x, sec, autoMinus) {
  const toks = sec.tokens;
  const out = [];
  const neg = x < 0 && autoMinus;
  let v = Math.abs(x);
  const pct = toks.filter((t) => t.t === 'pct').length;
  v *= 100 ** pct;
  const expAt = toks.findIndex((t) => t.t === 'exp');
  const slashAt = toks.findIndex((t) => t.t === 'slash');
  // ÷1000: запятые сразу после последнего знака целой части (перед точкой или концом числа)
  const numEnd = expAt >= 0 ? expAt : toks.length;
  let lastDigitBeforeDot = -1;
  const dotAt = toks.findIndex((t, i) => t.t === 'dot' && i < numEnd);
  for (let i = 0; i < (dotAt >= 0 ? dotAt : numEnd); i++) if (toks[i].t === 'digit') lastDigitBeforeDot = i;
  let scale = 0;
  for (let i = lastDigitBeforeDot + 1; i < (dotAt >= 0 ? dotAt : numEnd) && toks[i] && toks[i].t === 'comma'; i++) scale++;
  v /= 1000 ** scale;
  const firstDigit = toks.findIndex((t) => t.t === 'digit');
  const grouping = toks.some((t, i) => t.t === 'comma' && i > firstDigit && i < lastDigitBeforeDot);
  const sign = neg ? '-' : '';
  const intEnd = dotAt >= 0 ? dotAt : numEnd;
  void intEnd;
  if (slashAt >= 0) return sign + formatFraction(v, toks, slashAt);
  // токены: до первой цифры / целая часть / дробная часть / после последней цифры (запятые-разряды и ÷1000 — пропуск)
  const lastNum = toks.reduce((a, t, i) => (t.t === 'digit' && i < numEnd ? i : a), -1);
  const intToks = []; const decToks = []; const before = []; const after = [];
  toks.forEach((t, i) => {
    if (i >= numEnd) return;
    if (firstDigit < 0 || i < firstDigit) { before.push(t); return; }
    if (i > lastNum) { if (t.t !== 'comma' && t.t !== 'dot') after.push(t); return; }
    if (t.t === 'comma') return;
    if (dotAt < 0 || i < dotAt) intToks.push(t); else if (i > dotAt) decToks.push(t);
  });
  const nDec = decToks.filter((t) => t.t === 'digit').length;
  // заполнение *x выводится одним пробелом (в ячейке Excel растягивает его на ширину столбца)
  const piece = (arr) => arr.map((t) => (t.t === 'lit' ? t.s : t.t === 'space' || t.t === 'fill' ? ' ' : t.t === 'pct' ? '%' : '')).join('');
  if (expAt >= 0) {
    const nInt = intToks.filter((t) => t.t === 'digit').length || 1;
    let e = v === 0 ? 0 : Math.floor(Math.log10(v));
    if (nInt > 1) e = Math.floor(e / nInt) * nInt; else e -= nInt - 1;
    let m = v / 10 ** e;
    m = roundDec(m, nDec);
    if (m >= 10 ** nInt) { m /= 10 ** (nInt > 1 ? nInt : 1); e += nInt > 1 ? nInt : 1; m = roundDec(m, nDec); }
    // ноль в инженерном формате Excel показывает всеми знаками целой части как нулями (##0.0E+0 → 000,0E+0)
    const intZero = v === 0 && nInt > 1 ? intToks.map((t) => (t.t === 'digit' ? { ...t, ch: '0' } : t)) : intToks;
    const [ip, dp = ''] = m.toFixed(nDec).split('.');
    const expToks = toks.slice(expAt + 1).filter((t) => t.t === 'digit');
    const es = String(Math.abs(e)).padStart(expToks.filter((t) => t.ch === '0').length, '0');
    const esign = e < 0 ? '-' : toks[expAt].sign === '+' ? '+' : '';
    const tail = toks.slice(expAt + 1).filter((t) => t.t !== 'digit');
    return sign + piece(before) + fillInt(intZero, ip, false) + (dotAt >= 0 ? DEC + fillDec(decToks, dp) : '') + piece(after) + `E${esign}${es}` + piece(tail);
  }
  const r = roundDec(v, nDec);
  const [ip, dp = ''] = r.toFixed(nDec).split('.');
  const sg = r === 0 ? '' : sign; // −0 после округления — без минуса (как Excel)
  return sg + piece(before) + fillInt(intToks, ip, grouping) + (dotAt >= 0 ? DEC + fillDec(decToks, dp) : '') + piece(after) + piece(toks.slice(numEnd));
}
function formatFraction(v, toks, slashAt) {
  // [целая] [пробел] числитель / знаменатель
  const denToks = toks.slice(slashAt + 1);
  const fixed = denToks.filter((t) => t.t === 'num').map((t) => t.s).join('');
  const denPh = denToks.filter((t) => t.t === 'digit').length;
  const left = toks.slice(0, slashAt);
  // числитель — последняя группа цифр перед «/»; целая часть — группа до неё (через литерал)
  let k = left.length - 1;
  while (k >= 0 && left[k].t !== 'digit') k--;
  let numStart = k;
  while (numStart > 0 && left[numStart - 1].t === 'digit') numStart--;
  const numToks = left.slice(numStart, k + 1);
  const hasInt = left.slice(0, numStart).some((t) => t.t === 'digit');
  let whole = 0; let frac = v;
  if (hasInt) { whole = Math.floor(v); frac = v - whole; }
  let [n, d] = fixed ? [Math.round(frac * +fixed), +fixed] : bestFraction(frac, 10 ** Math.max(1, denPh) - 1);
  if (hasInt && n === d && d) { whole += 1; n = 0; }
  const intToks = left.slice(0, numStart);
  const pad = (s, t, right = false) => { const w = t.length; return right ? s.padEnd(w, ' ') : s.padStart(w, ' '); };
  const between = left.slice(k + 1).map((t) => (t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join('');
  const tail = denToks.filter((t) => t.t !== 'digit' && t.t !== 'num').map((t) => (t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join('');
  const den = fixed || pad(String(d), denToks.filter((t) => t.t === 'digit'), true);
  const intPart = hasInt ? fillInt(intToks, String(whole), false) : '';
  if (hasInt && n === 0) {
    // целое без дроби: дробь заменяется пробелами той же ширины
    const fracStr = `${pad('0', numToks)}${between}/${den}`;
    return (intPart.trim() ? intPart : fillInt(intToks.map((t) => (t.t === 'digit' ? { ...t, ch: '0' } : t)), '0', false)) + ' '.repeat(fracStr.length) + tail;
  }
  return `${intPart}${pad(String(n), numToks)}${between}/${den}${tail}`;
}

// ---------- даты и время ----------
// Как Excel: 0 — «00.01.1900», 60 — 29.02.1900 (ошибка 1900 года)
function excelDate(n) {
  const day = Math.floor(n);
  if (day === 0) return { y: 1900, m: 1, d: 0 };
  if (day === 60) return { y: 1900, m: 2, d: 29 };
  const t = new Date(Date.UTC(1899, 11, 30) + (day < 60 ? day + 1 : day) * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}
function formatDateSection(v, sec) {
  if (v < 0) return '#'.repeat(10);
  const toks = sec.tokens;
  // точность: доли секунды, если показаны (s.0), иначе — до секунды
  let fracDigits = 0;
  toks.forEach((t, i) => { if (t.t === 'date' && t.c === 's' && toks[i + 1] && toks[i + 1].t === 'dot') { let j = i + 2; while (toks[j] && toks[j].t === 'digit' && toks[j].ch === '0') { fracDigits++; j++; } } });
  const per = 10 ** fracDigits; const unit = 86400 * per;
  const units = Math.round(Number((v * unit).toPrecision(15)));
  const day = Math.floor(units / unit);
  const rest = units - day * unit; // доли секунды в сутках
  const secsInt = Math.floor(rest / per);
  const total = units / unit;
  const { y, m, d } = excelDate(day);
  const hh = Math.floor(secsInt / 3600); const mi = Math.floor((secsInt % 3600) / 60); const ss = secsInt % 60;
  const frac = (rest % per) / per;
  const ampm = toks.some((t) => t.t === 'ampm');
  const wd = ((day - 1) % 7 + 7) % 7;
  const p = (n, w) => String(n).padStart(w, '0');
  // «m» — минуты после часов или перед секундами
  const isMinute = (i) => {
    for (let j = i - 1; j >= 0; j--) { const t = toks[j]; if (t.t === 'date' || t.t === 'elapsed') return t.c === 'h' || t.unit === 'h'; }
    for (let j = i + 1; j < toks.length; j++) { const t = toks[j]; if (t.t === 'date' || t.t === 'elapsed') return t.c === 's' || t.unit === 's'; }
    return false;
  };
  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.t) {
      case 'date': {
        const n = t.len;
        if (t.c === 'y') out += n <= 2 ? p(y % 100, 2) : p(y, 4);
        else if (t.c === 'm' && n <= 2 && isMinute(i)) out += p(mi, n);
        else if (t.c === 'm') out += n === 1 ? m : n === 2 ? p(m, 2) : n === 3 ? MON[m - 1] : n === 4 ? MONTH[m - 1] : MONTH[m - 1][0];
        else if (t.c === 'd') out += n === 1 ? d : n === 2 ? p(d, 2) : n === 3 ? DAY_SHORT[wd] : DAY[wd];
        else if (t.c === 'h') { const h = ampm ? (hh % 12 || 12) : hh; out += n >= 2 ? p(h, 2) : h; }
        else if (t.c === 's') {
          out += n >= 2 ? p(ss, 2) : ss;
          if (toks[i + 1] && toks[i + 1].t === 'dot' && fracDigits) { out += DEC + String(Math.round(frac * 10 ** fracDigits)).padStart(fracDigits, '0').slice(0, fracDigits); i += 1 + fracDigits; }
        }
        break;
      }
      case 'elapsed': {
        const all = Math.floor(units / per);
        const val = t.unit === 'h' ? Math.floor(all / 3600) : t.unit === 'm' ? Math.floor(all / 60) : all;
        out += p(val, t.len);
        break;
      }
      case 'ampm': { const pm = hh >= 12; out += t.s.length === 5 ? (pm ? 'PM' : 'AM') : (pm ? 'P' : 'A'); break; }
      case 'lit': out += t.s.replace(/\//g, DATE_SEP); break;
      case 'slash': out += DATE_SEP; break;
      case 'dot': out += '.'; break;
      case 'comma': out += ','; break;
      case 'space': out += ' '; break;
      case 'digit': out += t.ch; break;
      case 'num': out += t.s; break;
      default: break;
    }
  }
  return out;
}

// ---------- главное ----------
const condOk = (c, v) => {
  switch (c.op) {
    case '<': return v < c.val; case '<=': return v <= c.val; case '>': return v > c.val; case '>=': return v >= c.val;
    case '=': return v === c.val; case '<>': return v !== c.val; default: return true;
  }
};
export function formatCode(v, code) {
  const secs = parseCode(String(code || 'General'));
  if (typeof v === 'string') {
    const tsec = secs.length >= 4 ? secs[3] : secs.find((s) => s.type === 'text') || (secs.length === 1 && secs[0].tokens.some((t) => t.t === 'at') ? secs[0] : null);
    if (!tsec) return { text: v, color: null };
    return { text: tsec.tokens.map((t) => (t.t === 'at' ? v : t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join(''), color: tsec.color };
  }
  let sec; let autoMinus = true;
  if (secs.some((s) => s.cond)) {
    sec = secs.find((s) => s.cond && condOk(s.cond, v));
    if (!sec) sec = secs.find((s) => !s.cond && s.type !== 'text') || secs[secs.length - 1];
    autoMinus = !(sec.cond && (sec.cond.op === '<' || sec.cond.op === '<=') && sec.cond.val <= 0);
  } else if (secs.length === 1 || (secs.length === 2 && secs[1].type === 'text')) sec = secs[0];
  else if (v > 0 || (v === 0 && secs.length < 3)) sec = secs[0];
  else if (v < 0) { sec = secs[1]; autoMinus = false; }
  else sec = secs[2];
  if (!sec.tokens.length) return { text: '', color: sec.color };
  let text;
  if (sec.type === 'text') return { text: general(v), color: null }; // раздел только для текста — число «Общим»
  if (sec.type === 'general') text = sec.tokens.map((t) => (t.t === 'general' ? general(autoMinus ? v : Math.abs(v)) : t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join('');
  else if (sec.type === 'date') text = formatDateSection(autoMinus ? v : Math.abs(v), sec);
  else if (sec.type === 'text') text = sec.tokens.map((t) => (t.t === 'lit' ? t.s : t.t === 'space' ? ' ' : '')).join('');
  else text = formatNumberSection(v, sec, autoMinus);
  return { text, color: sec.color };
}
