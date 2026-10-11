// Маркер заполнения (как в Excel): продолжение рядов при протягивании выделения.
// extendLine получает ячейки-источники одной строки/столбца в порядке заполнения и сколько ячеек заполнить;
// возвращает для каждой новой ячейки { v } — новое значение (стиль — от источника from) или { copy: from } —
// копия источника from (формула сдвигается вызывающим кодом).
import { isDateFormat } from './format.js';
import { dateToSerial, serialToDate } from './formula.js';

// Списки автозаполнения (как встроенные списки русского Excel — английских в нём нет): месяцы и дни недели
const LISTS = [
  ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'],
  ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
  ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'],
  ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'],
];
// Кварталы («Кв 1», «1 квартал», «кв.1»): после 4 — снова 1
const QUARTER = /^(кв\.?\s*|квартал\s*)[1-4]$|^[1-4](\s*кв\.?|\s+квартал)$/i;
// «май» есть и в полных, и в коротких названиях — при одном источнике берём полные (январь…)
const listItem = (s) => {
  const t = s.trim().toLowerCase();
  for (let li = 0; li < LISTS.length; li++) { const i = LISTS[li].indexOf(t); if (i >= 0) return { li, i }; }
  return null;
};
const caseOf = (s) => (s === s.toUpperCase() && s !== s.toLowerCase() ? 'upper' : s[0] === s[0].toUpperCase() && s[0] !== s[0].toLowerCase() ? 'cap' : 'lower');
const applyCase = (s, c) => (c === 'upper' ? s.toUpperCase() : c === 'cap' ? s[0].toUpperCase() + s.slice(1) : s);
// текст с числом: последнее число в конце («Кв 1», «Item10») или в начале («1 квартал»)
const textNum = (s) => {
  let m = /^(.*?)(\d+)$/.exec(s);
  if (m && (m[1] === '' || /\D$/.test(m[1]))) return { pre: m[1], n: +m[2], width: m[2].length, post: '' };
  m = /^(\d+)(\D.*)$/.exec(s);
  if (m) return { pre: '', n: +m[1], width: m[1].length, post: m[2] };
  return null;
};
const showTextNum = (t, n) => {
  const digits = String(Math.abs(n));
  return `${t.pre}${n < 0 ? '-' : ''}${t.width > digits.length ? digits.padStart(t.width, '0') : digits}${t.post}`;
};
const clean = (x) => parseFloat(x.toPrecision(15)); // 0,1 + 0,2 и т. п. — без хвостов
const isDateCell = (c) => !!(c && c.s && c.s.numFmt && isDateFormat(c.s.numFmt));

// Шаг дат: месяцы (одно и то же число месяца или все — последние дни месяцев), годы, иначе дни
const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
function dateStep(vals) {
  const ds = vals.map((v) => serialToDate(Math.floor(v)));
  const whole = vals.every((v) => v === Math.floor(v));
  const monthEnds = whole && ds.every((d) => d.d === daysIn(d.y, d.m));
  const sameDay = whole && ds.every((d) => d.d === ds[0].d);
  if ((sameDay || monthEnds) && vals.length > 1) {
    const months = ds.map((d) => d.y * 12 + d.m - 1);
    const step = months[1] - months[0];
    if (step !== 0 && months.every((m, i) => i === 0 || m - months[i - 1] === step)) {
      const last = ds[ds.length - 1]; const lastM = months[months.length - 1];
      return (k) => {
        const mm = lastM + step * k; const y = Math.floor(mm / 12); const m = mm - y * 12 + 1;
        const dim = daysIn(y, m);
        return dateToSerial(y, m, monthEnds && !sameDay ? dim : Math.min(last.d, dim));
      };
    }
  }
  return null;
}
// Линейный тренд по методу наименьших квадратов (как Excel для двух и более чисел)
function trend(vals) {
  const n = vals.length;
  const mx = (n - 1) / 2; const my = vals.reduce((a, b) => a + b, 0) / n;
  let sxy = 0; let sxx = 0;
  for (let i = 0; i < n; i++) { sxy += (i - mx) * (vals[i] - my); sxx += (i - mx) ** 2; }
  const b = sxx ? sxy / sxx : 0;
  return (k) => clean(my + b * (n - 1 + k - mx));
}

// src — массив ячеек (null — пустая), count — сколько заполнить, ctrl — Ctrl при протягивании (инвертирует
// поведение для одного числа: копия ↔ ряд), как в Excel
export function extendLine(src, count, { ctrl = false } = {}) {
  const n = src.length;
  const out = [];
  const vals = src.map((c) => (c && c.f == null ? c.v : undefined));
  const hasFormula = src.some((c) => c && c.f != null);
  const allNum = !hasFormula && n > 0 && vals.every((v) => typeof v === 'number');
  if (allNum) {
    const dates = src.every(isDateCell);
    let next = null;
    if (n === 1) {
      // одно число копируется (Ctrl — ряд с шагом 1), одна дата — ряд по дням (Ctrl — копия)
      if (dates ? !ctrl : ctrl) next = (k) => clean(vals[0] + k);
    } else if (!ctrl) next = (dates && dateStep(vals)) || trend(vals);
    if (next) { for (let k = 1; k <= count; k++) out.push({ v: next(k), from: n - 1 }); return out; }
  }
  if (!hasFormula && n > 0 && vals.every((v) => typeof v === 'string') && !ctrl) {
    // список (месяцы, дни недели): все из одного списка — шаг по списку
    const items = vals.map(listItem);
    if (items.every(Boolean)) {
      let li = items[0].li;
      if (n === 1 && li === 1 && items[0].i === 4) li = 0; // «май» — полные названия
      const L = LISTS[li]; const idx = vals.map((v) => L.indexOf(v.trim().toLowerCase()));
      if (idx.every((x) => x >= 0)) {
        const step = n === 1 ? 1 : ((idx[1] - idx[0]) % L.length + L.length) % L.length;
        if (idx.every((x, i) => i === 0 || ((x - idx[i - 1]) % L.length + L.length) % L.length === step)) {
          const cs = caseOf(vals[n - 1].trim());
          for (let k = 1; k <= count; k++) out.push({ v: applyCase(L[(idx[n - 1] + step * k) % L.length], cs), from: n - 1 });
          return out;
        }
      }
    }
    // кварталы: шаг 1 по кругу 1–4, иначе — копия (как Excel)
    const tn = vals.map(textNum);
    if (tn.every(Boolean) && vals.every((v) => QUARTER.test(v.trim())) && tn.every((t) => t.pre === tn[0].pre && t.post === tn[0].post)) {
      if (tn.every((t, i) => i === 0 || (t.n - tn[i - 1].n + 4) % 4 === 1)) {
        for (let k = 1; k <= count; k++) out.push({ v: showTextNum(tn[n - 1], ((tn[n - 1].n - 1 + k) % 4) + 1), from: n - 1 });
        return out;
      }
      for (let k = 0; k < count; k++) out.push({ copy: k % n });
      return out;
    }
    // текст с числом: одинаковые приставка и окончание — ряд (один источник — шаг 1)
    if (tn.every(Boolean) && tn.every((t) => t.pre === tn[0].pre && t.post === tn[0].post)) {
      const next = n === 1 ? (k) => tn[0].n + k : (() => { const f = trend(tn.map((t) => t.n)); return (k) => Math.round(f(k)); })();
      for (let k = 1; k <= count; k++) out.push({ v: showTextNum(tn[n - 1], next(k)), from: n - 1 });
      return out;
    }
  }
  // остальное — по кругу; текст с числом в смешанном выделении растёт на 1 с каждым кругом
  for (let k = 0; k < count; k++) {
    const j = k % n; const round = Math.floor(k / n) + 1;
    const c = src[j];
    const t = !ctrl && c && c.f == null && typeof c.v === 'string' ? textNum(c.v) : null;
    out.push(t ? { v: showTextNum(t, t.n + round), from: j } : { copy: j });
  }
  return out;
}
