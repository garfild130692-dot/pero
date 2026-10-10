// Числовые форматы и разбор ввода — с русскими разделителями (пробел между разрядами, запятая в дробной части).
import { dateToSerial, errText, isErr, serialToDate } from './formula.js';

// Встроенные форматы Excel (numFmtId → код), которые встречаются в файлах
export const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  14: 'dd.mm.yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'hh:mm', 21: 'hh:mm:ss',
  22: 'dd.mm.yyyy hh:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 49: '@',
};

// Форматы в списке на ленте
export const FORMAT_PRESETS = [
  ['General', 'Общий'],
  ['0.00', 'Числовой'],
  ['#,##0.00 "₽"', 'Денежный'],
  ['0.00%', 'Процентный'],
  ['dd.mm.yyyy', 'Дата'],
  ['hh:mm:ss', 'Время'],
  ['@', 'Текстовый'],
];

export const isDateFormat = (code) => {
  const c = String(code || '').replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  return /[dmyhs]/i.test(c) && !/^[#0.,%\s]*$/.test(c);
};

const group = (intStr) => intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');

function formatNumberCode(v, code) {
  // код: [префикс]основа[суффикс]; основа — 0, 0.00, #,##0.00, 0%, 0.00E+00
  const lit = (s) => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const m = /^(.*?)([#0,]*0(?:\.0+)?|#,##0(?:\.0+)?)(%?)(E\+00)?(.*)$/.exec(code);
  if (!m) return general(v);
  const [, pre, base, pct, sci, post] = m;
  let x = pct ? v * 100 : v;
  const dec = (base.split('.')[1] || '').length;
  let body;
  if (sci) body = x.toExponential(dec).replace('.', ',').replace('e', 'E');
  else {
    const neg = x < 0;
    const fixed = Math.abs(x).toFixed(dec);
    const [i, f] = fixed.split('.');
    body = (neg ? '-' : '') + (base.includes(',') ? group(i) : i) + (f ? `,${f}` : '');
  }
  return lit(pre) + body + (pct ? '%' : '') + lit(post);
}

// «Общий»: до 11 значащих символов, без лишних нулей
export function general(v) {
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  let s = String(Math.round(v * 1e10) / 1e10);
  if (s.length > 11 && Math.abs(v) >= 1e11) s = v.toExponential(5);
  return s.replace('.', ',');
}

function formatDate(v, code) {
  const { y, m, d, hh, mm, ss } = serialToDate(v);
  const p2 = (n) => String(n).padStart(2, '0');
  const MON = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
  const MONTH = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const DAY = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  const DAY_SHORT = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
  const wd = ((Math.floor(v) - 1) % 7 + 7) % 7; // 0 — воскресенье (серийный 1 = 01.01.1900, воскресенье)
  const c = String(code).replace(/\[\$-[^\]]*\]/g, '');
  if (/^\[h\]/i.test(c)) {
    const total = Math.round(v * 86400);
    return `${Math.floor(total / 3600)}:${p2(Math.floor(total / 60) % 60)}:${p2(total % 60)}`;
  }
  // минуты «mm» после h или перед s — иначе месяц
  // текст в кавычках и символ после «\» — как есть
  const tokens = c.match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|./gi) || [];
  let out = '';
  tokens.forEach((t, i) => {
    const low = t.toLowerCase();
    const prev = tokens.slice(0, i).reverse().find((x) => /^[hs]/i.test(x) || /^[dmy]/i.test(x));
    const next = tokens.slice(i + 1).find((x) => /^[hs]/i.test(x) || /^[dmy]/i.test(x));
    const isMin = (low === 'mm' || low === 'm') && ((prev && /^h/i.test(prev)) || (next && /^s/i.test(next)));
    switch (low) {
      case 'yyyy': out += y; break;
      case 'yy': out += p2(y % 100); break;
      case 'mmmmm': out += MONTH[m - 1][0]; break;
      case 'mmmm': out += MONTH[m - 1]; break;
      case 'mmm': out += MON[m - 1]; break;
      case 'dddd': out += DAY[wd]; break;
      case 'ddd': out += DAY_SHORT[wd]; break;
      case 'mm': out += isMin ? p2(mm) : p2(m); break;
      case 'm': out += isMin ? mm : m; break;
      case 'dd': out += p2(d); break;
      case 'd': out += d; break;
      case 'hh': out += p2(hh); break;
      case 'h': out += hh; break;
      case 'ss': out += p2(ss); break;
      case 's': out += ss; break;
      case 'am/pm': break;
      default: out += t.startsWith('"') ? t.slice(1, -1) : t.startsWith('\\') ? t.slice(1) : t;
    }
  });
  return out.trim();
}

// Значение ячейки → текст для показа
export function formatValue(v, code = 'General') {
  if (v == null) return '';
  if (isErr(v)) return errText(v.error); // как в русском Excel: #ДЕЛ/0!, #Н/Д…
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'string') return v;
  if (!isFinite(v)) return '#NUM!';
  const c = !code || code === 'General' ? null : code.split(';')[v < 0 && code.includes(';') ? 1 : 0];
  if (!c) return general(v);
  if (c === '@') return general(v);
  if (isDateFormat(c)) return formatDate(v, c);
  const text = formatNumberCode(v < 0 && code.includes(';') ? -v : v, c);
  return text;
}

// Разбор ввода пользователя: { value, format? } — число/дата/проценты/логическое/текст; формулы обрабатывает модель
export function parseInput(raw) {
  const s = String(raw);
  if (s === '') return { value: null };
  if (s.startsWith("'")) return { value: s.slice(1), format: null, text: true }; // «как текст»
  const t = s.trim();
  // число: пробелы между разрядами, запятая или точка
  const num = t.replace(/[\s ]/g, '');
  if (/^[-+]?(\d+([.,]\d*)?|[.,]\d+)([eE][-+]?\d+)?$/.test(num)) return { value: parseFloat(num.replace(',', '.')) };
  let m;
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)\s*%$/.exec(num))) return { value: parseFloat(m[1].replace(',', '.')) / 100, format: /[.,]/.test(m[1]) ? '0.00%' : '0%' }; // 12,5% → 12,50%, как в Excel
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)\s*(₽|р\.?|руб\.?)$/i.exec(t.replace(/[\s ](?=\d)/g, '')))) return { value: parseFloat(m[1].replace(',', '.')), format: '#,##0.00 "₽"' };
  if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    const mo = +m[2];
    const d = +m[1];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      const time = m[4] ? (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400 : 0;
      return { value: dateToSerial(y, mo, d) + time, format: m[4] ? 'dd.mm.yyyy hh:mm' : 'dd.mm.yyyy' };
    }
  }
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) return { value: (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, format: m[3] ? 'hh:mm:ss' : 'hh:mm' };
  if (/^(ИСТИНА|TRUE)$/i.test(t)) return { value: true };
  if (/^(ЛОЖЬ|FALSE)$/i.test(t)) return { value: false };
  return { value: s };
}

// Изменить число знаков после запятой в коде формата (кнопки «Увеличить/уменьшить разрядность»)
export function changeDecimals(code, delta, sample) {
  let c = !code || code === 'General' ? null : code;
  if (!c) {
    const dec = typeof sample === 'number' && !Number.isInteger(sample) ? Math.min(10, (String(sample).split('.')[1] || '').length) : 0;
    c = dec ? `0.${'0'.repeat(dec)}` : '0';
  }
  if (isDateFormat(c)) return code;
  return c.replace(/(0)(?:\.(0+))?(?=[%"\sE₽]|$)/, (all, z, dec = '') => {
    const n = Math.max(0, dec.length + delta);
    return n ? `0.${'0'.repeat(n)}` : '0';
  });
}
