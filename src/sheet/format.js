// Числовые форматы и разбор ввода — с русскими разделителями (пробел между разрядами, запятая в дробной части).
import { formatCode } from './numfmt.js';
import { dateToSerial, errText, isErr } from './formula.js';

// Встроенные форматы Excel (numFmtId → код), которые встречаются в файлах
export const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  14: 'dd.mm.yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
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

// «Общий»: до 11 значащих символов, без лишних нулей
export function general(v) {
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  let s = String(Math.round(v * 1e10) / 1e10);
  if (s.length > 11 && Math.abs(v) >= 1e11) s = v.toExponential(5);
  return s.replace('.', ',');
}

// Значение ячейки → текст для показа
export function formatValue(v, code = 'General') {
  if (v == null) return '';
  if (isErr(v)) return errText(v.error); // как в русском Excel: #ДЕЛ/0!, #Н/Д…
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'string') return code && code.includes('@') ? formatCode(v, code).text : v;
  if (!isFinite(v)) return '#NUM!';
  if (!code || code === 'General' || code === '@') return general(v);
  return formatCode(v, code).text; // разделы, условия, цвета, дроби, даты — numfmt.js
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
