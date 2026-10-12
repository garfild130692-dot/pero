// Формулы: токенизатор, разбор (приоритеты Excel), вычисление, функции, переписывание ссылок.
//
// Формула хранится без «=» в синтаксисе файла xlsx: английские имена функций, разделитель аргументов «,»,
// десятичная точка. При вводе принимаются и русские имена (СУММ, ЕСЛИ…) и «;» — они приводятся к хранимому виду.
import { colIndex, colName, MAX_COLS, MAX_ROWS } from './address.js';
import { functions, lazyFunctions } from './fnlib.js';
import { functions2 } from './fnlib2.js';
import { functions3 } from './fnlib3.js';
import { functions4 } from './fnlib4.js';
import { functions5 } from './fnlib5.js';
import { functions6 } from './fnlib6.js';
import { functionsEts } from './fnets.js';
import { EXCEL_RU } from './fnnames.js';

// ---------- Ошибки и значения ----------
export const ERR = {
  DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', NA: '#N/A', NUM: '#NUM!', CIRC: '#CIRC!',
  SPILL: '#SPILL!', CALC: '#CALC!', NULL: '#NULL!',
};
export const isErr = (v) => v != null && typeof v === 'object' && 'error' in v;
// Ошибки по-русски (как в интерфейсе русского Excel) ↔ хранимые английские
export const ERR_RU = {
  '#DIV/0!': '#ДЕЛ/0!', '#VALUE!': '#ЗНАЧ!', '#REF!': '#ССЫЛКА!', '#NAME?': '#ИМЯ?', '#N/A': '#Н/Д', '#NUM!': '#ЧИСЛО!',
  '#NULL!': '#ПУСТО!', '#SPILL!': '#ПЕРЕНОС!', '#CALC!': '#ВЫЧИСЛ!', '#CIRC!': '#ЦИКЛ!',
};
const ERR_EN = Object.fromEntries(Object.entries(ERR_RU).map(([en, ru]) => [ru, en]));
export const errText = (code) => ERR_RU[code] || code;

// Функции, появившиеся после Excel 2007: в файле пишутся с префиксом _xlfn. (ФИЛЬТР и СОРТ — _xlfn._xlws.),
// иначе Excel покажет #ИМЯ?. В модели формулы хранятся без префиксов.
export const FUTURE = new Set(`ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND
  BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST
  CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP
  ECMA.CEILING ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTER FILTERXML
  FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT
  GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN
  ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT
  NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC
  PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REDUCE
  RRI SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORT SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST
  TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE
  WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR`.split(/\s+/));
export const XLWS = new Set(['FILTER', 'SORT']);
export const err = (code) => ({ error: code });

// ---------- Русские имена функций ----------
const RU = {
  СУММ: 'SUM', СРЗНАЧ: 'AVERAGE', МИН: 'MIN', МАКС: 'MAX', СЧЁТ: 'COUNT', СЧЕТ: 'COUNT', СЧЁТЗ: 'COUNTA', СЧЕТЗ: 'COUNTA',
  СЧИТАТЬПУСТОТЫ: 'COUNTBLANK', ЕСЛИ: 'IF', И: 'AND', ИЛИ: 'OR', НЕ: 'NOT', ОКРУГЛ: 'ROUND', ОКРУГЛВВЕРХ: 'ROUNDUP',
  ОКРУГЛВНИЗ: 'ROUNDDOWN', ЦЕЛОЕ: 'INT', ОСТАТ: 'MOD', СТЕПЕНЬ: 'POWER', КОРЕНЬ: 'SQRT', СЦЕПИТЬ: 'CONCATENATE',
  СЦЕП: 'CONCAT', ДЛСТР: 'LEN', ПРОПИСН: 'UPPER', СТРОЧН: 'LOWER', СЖПРОБЕЛЫ: 'TRIM', ЛЕВСИМВ: 'LEFT', ПРАВСИМВ: 'RIGHT',
  ПСТР: 'MID', СЕГОДНЯ: 'TODAY', ТДАТА: 'NOW', СУММЕСЛИ: 'SUMIF', СЧЁТЕСЛИ: 'COUNTIF', СЧЕТЕСЛИ: 'COUNTIF',
  СРЗНАЧЕСЛИ: 'AVERAGEIF', ЕСЛИОШИБКА: 'IFERROR', ВПР: 'VLOOKUP', ПРОИЗВЕД: 'PRODUCT', ГОД: 'YEAR', МЕСЯЦ: 'MONTH',
  ДЕНЬ: 'DAY', ДАТА: 'DATE', ЕПУСТО: 'ISBLANK', ЕЧИСЛО: 'ISNUMBER', ЕТЕКСТ: 'ISTEXT', ЕОШИБКА: 'ISERROR', ПИ: 'PI',
  ЗНАК: 'SIGN', МЕДИАНА: 'MEDIAN', НАЙТИ: 'FIND', ПОДСТАВИТЬ: 'SUBSTITUTE', ПОВТОР: 'REPT', ЗНАЧЕН: 'VALUE',
  ИСТИНА: 'TRUE', ЛОЖЬ: 'FALSE', ПОСЛЕД: 'SEQUENCE', ТРАНСП: 'TRANSPOSE', ФИЛЬТР: 'FILTER', СОРТ: 'SORT', СОРТПО: 'SORTBY',
  УНИК: 'UNIQUE', СЛЧИС: 'RAND', СЛУЧМЕЖДУ: 'RANDBETWEEN', СЛУЧМАССИВ: 'RANDARRAY', ЧСТРОК: 'ROWS', ЧИСЛСТОЛБ: 'COLUMNS',
  // этап 12, часть 4
  ОКРУГЛТ: 'MROUND', ОКРВВЕРХ: 'CEILING', 'ОКРВВЕРХ.МАТ': 'CEILING.MATH', 'ОКРВВЕРХ.ТОЧН': 'CEILING.PRECISE', ОКРВНИЗ: 'FLOOR',
  'ОКРВНИЗ.МАТ': 'FLOOR.MATH', 'ОКРВНИЗ.ТОЧН': 'FLOOR.PRECISE', ОТБР: 'TRUNC', ЧЁТН: 'EVEN', ЧЕТН: 'EVEN', НЕЧЁТ: 'ODD', НЕЧЕТ: 'ODD',
  ФАКТР: 'FACT', НОД: 'GCD', НОК: 'LCM', ЧАСТНОЕ: 'QUOTIENT', ГРАДУСЫ: 'DEGREES', РАДИАНЫ: 'RADIANS', РИМСКОЕ: 'ROMAN',
  АРАБСКОЕ: 'ARABIC', СУММКВ: 'SUMSQ', ЧИСЛКОМБ: 'COMBIN', ПЕРЕСТ: 'PERMUT', ОСНОВАНИЕ: 'BASE', ДЕС: 'DECIMAL', КОРЕНЬПИ: 'SQRTPI',
  СУММРАЗНКВ: 'SUMX2MY2', СУММСУММКВ: 'SUMX2PY2', СУММКВРАЗН: 'SUMXMY2',
  ИСКЛИЛИ: 'XOR', ЕСЛИМН: 'IFS', ПЕРЕКЛЮЧ: 'SWITCH', ЕСНД: 'IFNA', НД: 'NA', ЕНД: 'ISNA', ЕОШ: 'ISERR', ЕЛОГИЧ: 'ISLOGICAL',
  ЕНЕТЕКСТ: 'ISNONTEXT', ЕЧЁТН: 'ISEVEN', ЕЧЕТН: 'ISEVEN', ЕНЕЧЁТ: 'ISODD', ЕНЕЧЕТ: 'ISODD', ЕССЫЛКА: 'ISREF', ЕФОРМУЛА: 'ISFORMULA',
  'Ф.ТЕКСТ': 'FORMULATEXT', Ч: 'N', ТИП: 'TYPE', 'ТИП.ОШИБКИ': 'ERROR.TYPE', ЛИСТ: 'SHEET', ЛИСТЫ: 'SHEETS', ЯЧЕЙКА: 'CELL',
  СУММЕСЛИМН: 'SUMIFS', СЧЁТЕСЛИМН: 'COUNTIFS', СЧЕТЕСЛИМН: 'COUNTIFS', СРЗНАЧЕСЛИМН: 'AVERAGEIFS', МАКСЕСЛИ: 'MAXIFS',
  МИНЕСЛИ: 'MINIFS', 'ПРОМЕЖУТОЧНЫЕ.ИТОГИ': 'SUBTOTAL', АГРЕГАТ: 'AGGREGATE',
  ГПР: 'HLOOKUP', ПРОСМОТР: 'LOOKUP', ПРОСМОТРX: 'XLOOKUP', ПРОСМОТРХ: 'XLOOKUP', ПОИСКПОЗX: 'XMATCH', ПОИСКПОЗХ: 'XMATCH',
  ВЫБОР: 'CHOOSE', СТРОКА: 'ROW', СТОЛБЕЦ: 'COLUMN', АДРЕС: 'ADDRESS', СМЕЩ: 'OFFSET', ДВССЫЛ: 'INDIRECT', 'РАНГ.РВ': 'RANK.EQ',
  'РАНГ.СР': 'RANK.AVG', ГИПЕРССЫЛКА: 'HYPERLINK',
  ПРОПНАЧ: 'PROPER', ПОИСК: 'SEARCH', ЗАМЕНИТЬ: 'REPLACE', ОБЪЕДИНИТЬ: 'TEXTJOIN', СОВПАД: 'EXACT', СИМВОЛ: 'CHAR', КОДСИМВ: 'CODE',
  ЮНИСИМВ: 'UNICHAR', ПЕЧСИМВ: 'CLEAN', ЧЗНАЧ: 'NUMBERVALUE', Т: 'T', ТЕКСТ: 'TEXT', ФИКСИРОВАННЫЙ: 'FIXED', РУБЛЬ: 'DOLLAR',
  ТЕКСТДО: 'TEXTBEFORE', ТЕКСТПОСЛЕ: 'TEXTAFTER', ТЕКСТРАЗД: 'TEXTSPLIT',
  ДЕНЬНЕД: 'WEEKDAY', НОМНЕДЕЛИ: 'WEEKNUM', 'НОМНЕДЕЛИ.ISO': 'ISOWEEKNUM', ДАТАМЕС: 'EDATE', КОНМЕСЯЦА: 'EOMONTH', РАБДЕНЬ: 'WORKDAY',
  'РАБДЕНЬ.МЕЖД': 'WORKDAY.INTL', ЧИСТРАБДНИ: 'NETWORKDAYS', 'ЧИСТРАБДНИ.МЕЖД': 'NETWORKDAYS.INTL', РАЗНДАТ: 'DATEDIF', ДНИ: 'DAYS',
  ДОЛЯГОДА: 'YEARFRAC', ДНЕЙ360: 'DAYS360', ВРЕМЯ: 'TIME', ЧАС: 'HOUR', МИНУТЫ: 'MINUTE', СЕКУНДЫ: 'SECOND', ДАТАЗНАЧ: 'DATEVALUE',
  ВРЕМЗНАЧ: 'TIMEVALUE',
  СРЗНАЧА: 'AVERAGEA', МАКСА: 'MAXA', МИНА: 'MINA', НАИБОЛЬШИЙ: 'LARGE', НАИМЕНЬШИЙ: 'SMALL', МОДА: 'MODE', 'МОДА.ОДН': 'MODE.SNGL',
  СТАНДОТКЛОН: 'STDEV', 'СТАНДОТКЛОН.В': 'STDEV.S', 'СТАНДОТКЛОН.Г': 'STDEV.P', СТАНДОТКЛОНП: 'STDEVP', ДИСП: 'VAR', 'ДИСП.В': 'VAR.S',
  'ДИСП.Г': 'VAR.P', ДИСПР: 'VARP', КВАРТИЛЬ: 'QUARTILE', 'КВАРТИЛЬ.ВКЛ': 'QUARTILE.INC', 'КВАРТИЛЬ.ИСКЛ': 'QUARTILE.EXC',
  ПЕРСЕНТИЛЬ: 'PERCENTILE', 'ПРОЦЕНТИЛЬ.ВКЛ': 'PERCENTILE.INC', 'ПРОЦЕНТИЛЬ.ИСКЛ': 'PERCENTILE.EXC',
  'ПЕРСЕНТИЛЬ.ВКЛ': 'PERCENTILE.INC', 'ПЕРСЕНТИЛЬ.ИСКЛ': 'PERCENTILE.EXC', ПРОЦЕНТРАНГ: 'PERCENTRANK',
  'ПРОЦЕНТРАНГ.ВКЛ': 'PERCENTRANK.INC', КОРРЕЛ: 'CORREL', НАКЛОН: 'SLOPE', ОТРЕЗОК: 'INTERCEPT', ПРЕДСКАЗ: 'FORECAST',
  'ПРЕДСКАЗ.ЛИНЕЙН': 'FORECAST.LINEAR', ТЕНДЕНЦИЯ: 'TREND', ЧАСТОТА: 'FREQUENCY', СРГЕОМ: 'GEOMEAN', СРГАРМ: 'HARMEAN',
  СРОТКЛ: 'AVEDEV', КВАДРОТКЛ: 'DEVSQ', 'НОРМ.РАСП': 'NORM.DIST', 'НОРМ.ОБР': 'NORM.INV', 'НОРМ.СТ.РАСП': 'NORM.S.DIST',
  'НОРМ.СТ.ОБР': 'NORM.S.INV', НОРМРАСП: 'NORMDIST', НОРМОБР: 'NORMINV', НОРМСТРАСП: 'NORMSDIST', НОРМСТОБР: 'NORMSINV',
  ПЛТ: 'PMT', ПРПЛТ: 'IPMT', ОСПЛТ: 'PPMT', БС: 'FV', ПС: 'PV', КПЕР: 'NPER', СТАВКА: 'RATE', ЧПС: 'NPV', ВСД: 'IRR', ЧИСТНЗ: 'XNPV',
  ЧИСТВНДОХ: 'XIRR', АПЛ: 'SLN', ФУО: 'DB', ДДОБ: 'DDB', АСЧ: 'SYD', ЭФФЕКТ: 'EFFECT', НОМИНАЛ: 'NOMINAL', ИНДЕКС: 'INDEX', ПОИСКПОЗ: 'MATCH', СУММПРОИЗВ: 'SUMPRODUCT', РАНГ: 'RANK',
};

for (const [en, ru] of Object.entries(EXCEL_RU)) if (!RU[ru.toUpperCase()]) RU[ru.toUpperCase()] = en;

// ---------- Токенизатор ----------
// Токены: { type, text, ... }; text — исходный текст (для переписывания формулы без потери пробелов)
// Префикс листа: Лист1!, 'Мой лист'!, трёхмерный Лист1:Лист3! или 'Лист 1:Лист 3'!
const SHEET_RE = /^(?:'((?:[^']|'')+)'|([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*(?::[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)?))!/;
const splitSheets = (name) => {
  const i = name.indexOf(':');
  return i < 0 ? [name, null] : [name.slice(0, i), name.slice(i + 1)];
};
export function tokenize(src) {
  const out = [];
  let i = 0;
  const s = String(src);
  while (i < s.length) {
    const rest = s.slice(i);
    let m;
    if ((m = /^\s+/.exec(rest))) { out.push({ type: 'ws', text: m[0] }); i += m[0].length; continue; }
    if ((m = /^"((?:[^"]|"")*)"/.exec(rest))) { out.push({ type: 'str', text: m[0], value: m[1].replace(/""/g, '"') }); i += m[0].length; continue; }
    if ((m = /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!|CIRC!|SPILL!|CALC!|ДЕЛ\/0!|ЗНАЧ!|ССЫЛКА!|ИМЯ\?|Н\/Д|ЧИСЛО!|ПУСТО!|ПЕРЕНОС!|ВЫЧИСЛ!|ЦИКЛ!)/.exec(rest))) {
      out.push({ type: 'err', text: m[0], code: ERR_EN[m[0]] || m[0] }); i += m[0].length; continue;
    }
    // ссылка (возможно с листом): Лист1!A1, 'Мой лист'!$B$2, Лист1:Лист3!A1; целые столбцы A:C и строки 1:5
    let sheet = null; let sheet2 = null;
    let sheetText = '';
    if ((m = SHEET_RE.exec(rest))) {
      const after = rest.slice(m[0].length);
      if (/^(\$?[A-Za-z]{1,3}\$?\d|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])|\$?\d{1,7}:\$?\d)/.test(after)) {
        [sheet, sheet2] = splitSheets(m[1] ? m[1].replace(/''/g, "'") : m[2]);
        sheetText = m[0];
      }
    }
    const refPart = rest.slice(sheetText.length);
    if ((m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(!.])/.exec(refPart))) {
      const c1 = colIndex(m[2]); const c2 = colIndex(m[4]);
      out.push({ type: 'cols', text: sheetText + m[0], sheet, sheet2, abs1: !!m[1], c1: Math.min(c1, c2), abs2: !!m[3], c2: Math.max(c1, c2) });
      i += sheetText.length + m[0].length;
      continue;
    }
    if ((m = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\d.])/.exec(refPart))) {
      const r1 = +m[2] - 1; const r2 = +m[4] - 1;
      out.push({ type: 'rows', text: sheetText + m[0], sheet, sheet2, abs1: !!m[1], r1: Math.min(r1, r2), abs2: !!m[3], r2: Math.max(r1, r2) });
      i += sheetText.length + m[0].length;
      continue;
    }
    if ((m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(#?)(?![\w(])/.exec(refPart))) {
      out.push({ type: 'ref', text: sheetText + m[0], sheet, sheet2, absC: !!m[1], c: colIndex(m[2]), absR: !!m[3], r: +m[4] - 1, spill: !!m[5] });
      i += sheetText.length + m[0].length;
      continue;
    }
    if ((m = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest)) || (m = /^\.\d+(?:[eE][+-]?\d+)?/.exec(rest))) {
      out.push({ type: 'num', text: m[0], value: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    if ((m = /^([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)\s*\(/.exec(rest))) {
      const name = m[1].toUpperCase().replace(/^_XLFN\.(_XLWS\.)?/, '');
      out.push({ type: 'func', text: m[1], name: RU[name] || name });
      i += m[1].length;
      continue;
    }
    if ((m = /^(TRUE|FALSE|ИСТИНА|ЛОЖЬ)(?![\wА-Яа-яЁё])/i.exec(rest))) {
      const v = /^(TRUE|ИСТИНА)$/i.test(m[1]);
      out.push({ type: 'bool', text: m[1], value: v }); i += m[1].length; continue;
    }
    if ((m = /^(<=|>=|<>)/.exec(rest))) { out.push({ type: 'op', text: m[0] }); i += 2; continue; }
    const ch = s[i];
    if ('+-*/^&=<>%'.includes(ch)) { out.push({ type: 'op', text: ch }); i++; continue; }
    if (ch === '(' || ch === ')' || ch === '{' || ch === '}') { out.push({ type: ch, text: ch }); i++; continue; }
    if (ch === ',' || ch === ';') { out.push({ type: 'sep', text: ch }); i++; continue; }
    if (ch === ':') { out.push({ type: ':', text: ch }); i++; continue; }
    // имя (именованный диапазон, параметр LET); _xlpm. — префикс параметров в файле
    if ((m = /^[A-Za-zА-Яа-яЁё_\\][\wА-Яа-яЁё.]*/.exec(rest))) { out.push({ type: 'name', text: m[0], name: m[0].replace(/^_xlpm\./i, '') }); i += m[0].length; continue; }
    out.push({ type: 'bad', text: ch });
    i++;
  }
  return out;
}

// ---------- Разбор ----------
export function parse(src) {
  const toks = tokenize(src).filter((t) => t.type !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (type) => { const t = next(); if (!t || t.type !== type) throw new Error('syntax'); return t; };
  const isOp = (t, ops) => t && t.type === 'op' && ops.includes(t.text);

  function primary() {
    const t = next();
    if (!t) throw new Error('syntax');
    switch (t.type) {
      case 'num': return { t: 'num', v: t.value };
      case 'str': return { t: 'str', v: t.value };
      case 'bool': return { t: 'bool', v: t.value };
      case 'err': return { t: 'err', v: t.code };
      case '{': {
        // константа массива: столбцы через «,», строки через «;» (запись файла)
        const rows = [[]];
        for (;;) {
          let neg = false;
          let t2 = next();
          if (t2 && t2.type === 'op' && (t2.text === '-' || t2.text === '+')) { neg = t2.text === '-'; t2 = next(); }
          if (!t2) throw new Error('syntax');
          let v;
          if (t2.type === 'num') v = neg ? -t2.value : t2.value;
          else if (t2.type === 'str' && !neg) v = t2.value;
          else if (t2.type === 'bool' && !neg) v = t2.value;
          else if (t2.type === 'err' && !neg) v = err(t2.code);
          else throw new Error('syntax');
          rows[rows.length - 1].push(v);
          const sep = next();
          if (!sep) throw new Error('syntax');
          if (sep.type === '}') break;
          if (sep.type !== 'sep') throw new Error('syntax');
          if (sep.text === ';') rows.push([]);
        }
        if (rows.some((r) => r.length !== rows[0].length)) throw new Error('syntax');
        return { t: 'array', v: rows };
      }
      case 'ref': {
        if (t.spill) return { t: 'spill', sheet: t.sheet, r: t.r, c: t.c };
        if (peek() && peek().type === ':' && toks[p + 1] && toks[p + 1].type === 'ref') {
          next();
          const b = next();
          // abs — признаки «$» концов диапазона (нужны записи формулы в .xls): [строка 1, столбец 1, строка 2, столбец 2]
          const [ra, rb] = t.r <= b.r ? [t, b] : [b, t]; const [ca, cb] = t.c <= b.c ? [t, b] : [b, t];
          return { t: 'range', sheet: t.sheet, sheet2: t.sheet2, r1: Math.min(t.r, b.r), c1: Math.min(t.c, b.c), r2: Math.max(t.r, b.r), c2: Math.max(t.c, b.c), abs: [!!ra.absR, !!ca.absC, !!rb.absR, !!cb.absC] };
        }
        return t.sheet2 ? { t: 'range', sheet: t.sheet, sheet2: t.sheet2, r1: t.r, c1: t.c, r2: t.r, c2: t.c, abs: [!!t.absR, !!t.absC, !!t.absR, !!t.absC] } : { t: 'ref', sheet: t.sheet, r: t.r, c: t.c, abs: [!!t.absR, !!t.absC] };
      }
      case 'cols': return { t: 'range', sheet: t.sheet, sheet2: t.sheet2, r1: 0, r2: MAX_ROWS - 1, c1: t.c1, c2: t.c2, whole: 'col', abs: [true, !!t.abs1, true, !!t.abs2] };
      case 'rows': return { t: 'range', sheet: t.sheet, sheet2: t.sheet2, r1: t.r1, r2: t.r2, c1: 0, c2: MAX_COLS - 1, whole: 'row', abs: [!!t.abs1, true, !!t.abs2, true] };
      case 'func': {
        expect('(');
        // ANCHORARRAY(A1) — так в файле записана ссылка на «разлив» A1#
        if (t.name === 'ANCHORARRAY' && peek() && peek().type === 'ref' && toks[p + 1] && toks[p + 1].type === ')') {
          const r0 = next(); next();
          return { t: 'spill', sheet: r0.sheet, r: r0.r, c: r0.c };
        }
        const args = [];
        if (peek() && peek().type === ')') { next(); return { t: 'fn', name: t.name, text: t.text, args }; }
        for (;;) {
          if (peek() && (peek().type === 'sep' || peek().type === ')')) args.push({ t: 'empty' });
          else args.push(expr());
          const n = next();
          if (!n) throw new Error('syntax');
          if (n.type === ')') break;
          if (n.type !== 'sep') throw new Error('syntax');
        }
        return { t: 'fn', name: t.name, text: t.text, args };
      }
      case '(': { const e = expr(); expect(')'); return e; }
      case 'op':
        if (t.text === '-' || t.text === '+') return { t: 'un', op: t.text, a: unary() };
        throw new Error('syntax');
      case 'name': return { t: 'name', v: t.name };
      default: throw new Error('syntax');
    }
  }
  function postfix() {
    let a = primary();
    for (;;) {
      if (isOp(peek(), ['%'])) { next(); a = { t: 'pct', a }; continue; }
      // вызов результата выражения: LAMBDA(x;x+1)(2), МояФункция(1)(2)
      if (peek() && peek().type === '(' && (a.t === 'fn' || a.t === 'call' || a.t === 'name')) {
        next();
        const args = [];
        if (peek() && peek().type === ')') next();
        else {
          for (;;) {
            if (peek() && (peek().type === 'sep' || peek().type === ')')) args.push({ t: 'empty' });
            else args.push(expr());
            const n = next();
            if (!n) throw new Error('syntax');
            if (n.type === ')') break;
            if (n.type !== 'sep') throw new Error('syntax');
          }
        }
        a = { t: 'call', fn: a, args };
        continue;
      }
      return a;
    }
  }
  function unary() {
    if (isOp(peek(), ['-', '+'])) { const op = next().text; return { t: 'un', op, a: unary() }; }
    return postfix();
  }
  function power() {
    let a = unary();
    while (isOp(peek(), ['^'])) { next(); a = { t: 'bin', op: '^', a, b: unary() }; }
    return a;
  }
  function mul() {
    let a = power();
    while (isOp(peek(), ['*', '/'])) { const op = next().text; a = { t: 'bin', op, a, b: power() }; }
    return a;
  }
  function add() {
    let a = mul();
    while (isOp(peek(), ['+', '-'])) { const op = next().text; a = { t: 'bin', op, a, b: mul() }; }
    return a;
  }
  function concat() {
    let a = add();
    while (isOp(peek(), ['&'])) { next(); a = { t: 'bin', op: '&', a, b: add() }; }
    return a;
  }
  function expr() {
    let a = concat();
    while (isOp(peek(), ['=', '<>', '<', '>', '<=', '>='])) { const op = next().text; a = { t: 'bin', op, a, b: concat() }; }
    return a;
  }
  const ast = expr();
  if (p < toks.length) throw new Error('syntax');
  return ast;
}

// ---------- Приведение типов ----------
const parseNumStr = (s) => {
  const t = String(s).trim().replace(/\s/g, '').replace(',', '.');
  return t !== '' && isFinite(+t) ? +t : null;
};
// Текст → число, как русский Excel в арифметике и ЗНАЧЕН: «1 234,5», «12%», «100 ₽», «15.01.2026», «14:30»
export function textToNum(text) {
  const t = String(text).trim();
  let n = parseNumStr(t);
  if (n !== null && t !== '') return n;
  let m;
  const flat = t.replace(/[\s\u00a0]/g, '');
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)%$/.exec(flat))) return parseFloat(m[1].replace(',', '.')) / 100;
  if ((m = /^([-+]?\d+(?:[.,]\d+)?)(₽|р\.?|руб\.?)$/i.exec(flat))) return parseFloat(m[1].replace(',', '.'));
  if ((m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t))) {
    const y = m[3].length === 2 ? (+m[3] < 30 ? 2000 : 1900) + +m[3] : +m[3];
    if (+m[2] >= 1 && +m[2] <= 12 && +m[1] >= 1 && +m[1] <= 31) {
      n = dateToSerial(y, +m[2], +m[1]);
      return m[4] ? n + (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400 : n;
    }
  }
  if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t))) return (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400;
  return null;
}
// Пустая ячейка — 0; пустая строка (например, результат ="") — #ЗНАЧ!, как в Excel
export function toNum(v) {
  if (isErr(v)) return v;
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = textToNum(v);
  return n === null ? err(ERR.VALUE) : n;
}
// Число как текст в формате «Общий» русского Excel: до 15 значащих цифр, десятичная запятая
export function numText(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let t = String(+n.toPrecision(15));
  if (/e/.test(t)) t = t.replace(/e([+-])(\d)$/, 'E$10$2').replace('e', 'E');
  return t.replace('.', ',');
}
export function toText(v) {
  if (v == null) return '';
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v === 'number') return numText(v);
  return String(v);
}
function toBool(v) {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  if (v == null) return false;
  if (v === '') return err(ERR.VALUE);
  if (typeof v === 'number') return v !== 0;
  const s = String(v).toUpperCase();
  if (s === 'TRUE' || s === 'ИСТИНА') return true;
  if (s === 'FALSE' || s === 'ЛОЖЬ') return false;
  return err(ERR.VALUE);
}
const flat = (v) => (Array.isArray(v) ? v.flat() : [v]);

// Сравнение как в Excel: числа < текст < логические; текст — без учёта регистра
function compare(a, b) {
  const rank = (x) => (typeof x === 'number' || x == null ? 0 : typeof x === 'string' ? 1 : 2);
  const ra = rank(a);
  const rb = rank(b);
  if (a == null && typeof b === 'string') return b === '' ? 0 : -1;
  if (b == null && typeof a === 'string') return a === '' ? 0 : 1;
  if (ra !== rb) return ra - rb;
  if (ra === 1) { const x = a.toLowerCase(); const y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  const x = ra === 2 ? +a : a || 0;
  const y = rb === 2 ? +b : b || 0;
  if (x === y || (ra === 0 && eq15(x, y))) return 0;
  return x < y ? -1 : 1;
}
// Excel сравнивает числа с точностью 15 значащих цифр (0,1+0,2=0,3 — ИСТИНА)
const eq15 = (x, y) => x === y || (Number.isFinite(x) && Number.isFinite(y) && x.toPrecision(15) === y.toPrecision(15));

// Подстановочные знаки Excel (* ? и ~ для экранирования) → RegExp; null — их нет
export function wildcard(text) {
  if (!/[*?]/.test(text)) return null;
  let re = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '~' && i + 1 < text.length) { re += text[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    re += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'i');
}
// Приблизительный поиск как в Excel (ВПР/ГПР/ПОИСКПОЗ/ПРОСМОТР): двоичный — последняя позиция со значением ≤ искомого
// (по убыванию — ≥). На неотсортированных данных Excel из-за этого находит «не то» — и Перо так же.
export function approxIndex(list, val, desc = false) {
  let lo = 0; let hi = list.length - 1; let res = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    // пустые и значения другого типа пропускаются (ищем ближайшее подходящее вниз)
    while (m >= lo && (list[m] == null || (typeof list[m] !== typeof val && !(typeof val === 'number' && typeof list[m] === 'number')))) m--;
    if (m < lo) { lo = mid + 1; continue; }
    const r = compare(list[m], val);
    if (desc ? r >= 0 : r <= 0) { res = m; lo = mid + 1; } else hi = m - 1;
  }
  return res;
}
// Точный поиск по запомненному диапазону (rangeEntry): индекс «значение → первая строка» по столбцу, с тем же сравнением,
// что compare (текст без регистра, числа до 15 значащих цифр, пустая = 0 и ""). null — индексом не найти (подстановочные
// знаки, пустое искомое, нет запомненного диапазона) — тогда обычный перебор.
const exactKey = (v) => (typeof v === 'number' ? (Number.isFinite(v) ? `n${v.toPrecision(15)}` : null) : typeof v === 'string' ? `s${v.toLowerCase()}` : typeof v === 'boolean' ? `b${+v}` : null);
function indexedFind(table, col, val, byRow = false) {
  const entry = table && table.rangeEntry;
  if (!entry || val == null || (typeof val === 'string' && /[*?]/.test(val))) return null;
  const k = exactKey(val);
  if (k == null) return null;
  const ik = `${byRow ? 'r' : 'c'}${col}`;
  let idx = entry.idx.get(ik);
  if (!idx) {
    idx = new Map();
    const a = entry.arr;
    const n = byRow ? (a[0] ? a[0].length : 0) : a.length;
    for (let i = 0; i < n; i++) {
      const v = byRow ? a[col][i] : a[i][col];
      const keys = v == null ? ['n0.00000000000000', 's'] : [exactKey(v)];
      for (const x of keys) if (x != null && !idx.has(x)) idx.set(x, i);
    }
    entry.idx.set(ik, idx);
  }
  return idx.has(k) ? idx.get(k) : -1;
}
// Точное совпадение для ВПР/ПОИСКПОЗ/ГПР: текст — без регистра и с подстановочными знаками
function exactMatcher(val) {
  const re = typeof val === 'string' ? wildcard(val) : null;
  return re ? (v) => typeof v === 'string' && re.test(v) : (v) => compare(v, val) === 0;
}

// Условие СУММЕСЛИ/СЧЁТЕСЛИ: ">5", "<>x", "=текст", "те*ст", число
function criterion(c) {
  if (typeof c === 'boolean') return (v) => v === c;
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? eq15(v, c) : typeof v === 'string' && parseNumStr(v) !== null && eq15(parseNumStr(v), c));
  const s = typeof c === 'string' ? c : toText(c);
  if (s === '<>') return (v) => v != null;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s);
  const op = m[1] || '=';
  const raw = m[2];
  const n = parseNumStr(raw);
  const target = n !== null ? n : raw;
  const re = n === null && (op === '=' || op === '<>') ? wildcard(raw) : null;
  if (re) return (v) => (op === '=') === (typeof v === 'string' && re.test(v));
  return (v) => {
    if (n !== null && typeof v !== 'number') { const vn = typeof v === 'string' ? parseNumStr(v) : null; if (vn === null) return op === '<>'; v = vn; }
    const r = compare(v == null ? (n !== null ? 0 : '') : v, target);
    return { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[op];
  };
}

// ---------- Даты (серийные номера, система 1900) ----------
const EPOCH = Date.UTC(1899, 11, 30);
// Как Excel: годы 0–1899 — плюс 1900; 1900 считается високосным (серийный номер 60 — 29.02.1900, «ошибка 1900 года»)
export const dateToSerial = (y, m, d) => {
  const yy = y < 1900 ? y + 1900 : y;
  let first = (Date.UTC(yy, m - 1, 1) - EPOCH) / 86400000; // первое число месяца
  if (first < 61) first -= 1; // до 1 марта 1900 Excel отстаёт на день
  return first + d - 1;
};
export function serialToDate(n) {
  if (Math.floor(n) === 60) return { y: 1900, m: 2, d: 29, hh: 0, mm: 0, ss: 0 };
  const d = new Date(EPOCH + Math.round((n < 60 ? n + 1 : n) * 86400000));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds() };
}
const nowSerial = () => { const t = new Date(); return (Date.UTC(t.getFullYear(), t.getMonth(), t.getDate(), t.getHours(), t.getMinutes(), t.getSeconds()) - EPOCH) / 86400000; };

// ---------- Функции ----------
const round = (x, n, mode) => {
  const k = 10 ** n;
  const v = x * k;
  const r = mode === 'up' ? Math.sign(v) * Math.ceil(Math.abs(v) - 1e-9) : mode === 'down' ? Math.sign(v) * Math.floor(Math.abs(v) + 1e-9) : Math.sign(v) * Math.round(Math.abs(v) + 1e-9);
  return r / k;
};
const nums = (args, { strict = false } = {}) => {
  const out = [];
  for (const a of args) {
    if (Array.isArray(a)) { for (const v of a.flat()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); } }
    else { if (isErr(a)) return a; if (a == null || a === '') { if (strict) out.push(0); continue; } const n = toNum(a); if (isErr(n)) return n; out.push(n); }
  }
  return out;
};
const firstErr = (...xs) => xs.find(isErr);
const toArr = (v) => (Array.isArray(v) ? v : [[v]]);
// Порядок сортировки Excel: числа < текст < логические < ошибки; пустые — в конце
function sortCompare(a, b) {
  const rank = (x) => (x == null ? 4 : isErr(x) ? 3 : typeof x === 'boolean' ? 2 : typeof x === 'string' ? 1 : 0);
  const ra = rank(a); const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b;
  if (ra === 1) return a.localeCompare(b, 'ru', { sensitivity: 'base' });
  if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return 0;
}
const stableSort = (list, cmp) => list.map((x, i) => [x, i]).sort((a, b) => cmp(a[0], b[0]) || a[1] - b[1]).map((x) => x[0]);
function logicals(args) {
  const out = [];
  for (const a of args) {
    for (const v of flat(a)) {
      if (v == null || v === '') continue;
      if (isErr(v) || typeof v === 'boolean') { out.push(v); continue; }
      if (typeof v === 'number') { out.push(v !== 0); continue; }
      const b = toBool(v);
      if (!isErr(b)) out.push(b);
    }
  }
  return out;
}

// аргументы приходят уже вычисленными (диапазон — двумерный массив), кроме «ленивых» IF/IFERROR (см. evaluate)
const FN = {
  SUM: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s + x, 0); },
  PRODUCT: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s * x, 1); },
  AVERAGE: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? n.reduce((s, x) => s + x, 0) / n.length : err(ERR.DIV0); },
  MIN: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; },
  MEDIAN: (a) => {
    const n = nums(a);
    if (isErr(n)) return n;
    if (!n.length) return err(ERR.NUM);
    const s = [...n].sort((x, y) => x - y);
    const h = s.length >> 1;
    return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
  },
  COUNT: (a) => a.reduce((s, x) => s + flat(x).filter((v) => typeof v === 'number').length, 0),
  COUNTA: (a) => a.reduce((s, x) => s + flat(x).filter((v) => v != null).length, 0),
  COUNTBLANK: (a) => a.reduce((s, x) => {
    const filled = flat(x).filter((v) => v != null && v !== '').length;
    const size = Array.isArray(x) && x.full ? x.full.rows * x.full.cols : flat(x).length;
    return s + size - filled;
  }, 0),
  ABS: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.abs(n); },
  SIGN: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.sign(n); },
  INT: ([x]) => { const n = toNum(x); return isErr(n) ? n : Math.floor(n); },
  SQRT: ([x]) => { const n = toNum(x); return isErr(n) ? n : n < 0 ? err(ERR.NUM) : Math.sqrt(n); },
  PI: () => Math.PI,
  POWER: ([x, y]) => { const a = toNum(x); const b = toNum(y); return firstErr(a, b) || a ** b; },
  MOD: ([x, y]) => {
    const a = toNum(x); const b = toNum(y);
    if (firstErr(a, b)) return firstErr(a, b);
    if (b === 0) return err(ERR.DIV0);
    return a - b * Math.floor(a / b);
  },
  ROUND: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d)); },
  ROUNDUP: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d), 'up'); },
  ROUNDDOWN: ([x, n]) => { const a = toNum(x); const d = toNum(n ?? 0); return firstErr(a, d) || round(a, Math.trunc(d), 'down'); },
  AND: (a) => { let r = true; for (const v of logicals(a)) { if (isErr(v)) return v; r = r && v; } return r; },
  OR: (a) => { let r = false; for (const v of logicals(a)) { if (isErr(v)) return v; r = r || v; } return r; },
  NOT: ([x]) => { const b = toBool(x); return isErr(b) ? b : !b; },
  TRUE: () => true,
  FALSE: () => false,
  CONCATENATE: (a) => { const e = a.flatMap(flat).find(isErr); return e || a.flatMap(flat).map(toText).join(''); },
  CONCAT: (a) => FN.CONCATENATE(a),
  LEN: ([x]) => (isErr(x) ? x : toText(x).length),
  UPPER: ([x]) => (isErr(x) ? x : toText(x).toUpperCase()),
  LOWER: ([x]) => (isErr(x) ? x : toText(x).toLowerCase()),
  TRIM: ([x]) => (isErr(x) ? x : toText(x).trim().replace(/\s+/g, ' ')),
  LEFT: ([x, n]) => { const k = toNum(n ?? 1); return firstErr(x, k) || toText(x).slice(0, Math.max(0, k)); },
  RIGHT: ([x, n]) => { const k = toNum(n ?? 1); const s = toText(x); return firstErr(x, k) || (k > 0 ? s.slice(-k) : ''); },
  MID: ([x, st, n]) => { const a = toNum(st); const k = toNum(n); return firstErr(x, a, k) || toText(x).substr(Math.max(0, a - 1), Math.max(0, k)); },
  FIND: ([what, where, st]) => {
    const s = toNum(st ?? 1);
    if (firstErr(what, where, s)) return firstErr(what, where, s);
    const i = toText(where).indexOf(toText(what), s - 1);
    return i < 0 ? err(ERR.VALUE) : i + 1;
  },
  SUBSTITUTE: ([t, o, n, k]) => {
    const e = firstErr(t, o, n, k);
    if (e) return e;
    const text = toText(t); const old = toText(o);
    if (!old) return text;
    if (k == null) return text.split(old).join(toText(n));
    const which = toNum(k);
    if (isErr(which) || which < 1) return err(ERR.VALUE);
    let at = -1;
    for (let i = 0; i < which; i++) { at = text.indexOf(old, at + 1); if (at < 0) return text; }
    return text.slice(0, at) + toText(n) + text.slice(at + old.length);
  },
  REPT: ([t, n]) => { const k = toNum(n); return firstErr(t, k) || toText(t).repeat(Math.max(0, Math.floor(k))); },
  VALUE: ([x]) => toNum(x),
  ISBLANK: ([x]) => x == null,
  ISNUMBER: ([x]) => typeof x === 'number',
  ISTEXT: ([x]) => typeof x === 'string',
  ISERROR: ([x]) => isErr(x),
  TODAY: () => Math.floor(nowSerial()),
  NOW: () => nowSerial(),
  DATE: ([y, m, d]) => {
    const a = toNum(y); const b = toNum(m); const c = toNum(d);
    if (firstErr(a, b, c)) return firstErr(a, b, c);
    const yy = Math.trunc(a);
    if (yy < 0 || yy > 9999) return err(ERR.NUM);
    const v = dateToSerial(yy, Math.trunc(b), Math.trunc(c));
    return v < 0 ? err(ERR.NUM) : v;
  },
  YEAR: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).y; },
  MONTH: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).m; },
  DAY: ([x]) => { const n = toNum(x); return isErr(n) ? n : serialToDate(n).d; },
  SUMIF: ([range, crit, sumRange]) => {
    const test = criterion(crit);
    const r = Array.isArray(range) ? range : [[range]];
    const s = Array.isArray(sumRange) ? sumRange : r;
    let total = 0;
    r.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') total += x; } }));
    return total;
  },
  COUNTIF: ([range, crit]) => { const test = criterion(crit); return flat(range).filter((v) => test(v)).length; },
  AVERAGEIF: ([range, crit, avgRange]) => {
    const test = criterion(crit);
    const r = Array.isArray(range) ? range : [[range]];
    const s = Array.isArray(avgRange) ? avgRange : r;
    const pick = [];
    r.forEach((row, i) => row.forEach((v, j) => { if (test(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') pick.push(x); } }));
    return pick.length ? pick.reduce((a, b) => a + b, 0) / pick.length : err(ERR.DIV0);
  },
  SUMPRODUCT: (a) => {
    const arrs = a.map((x) => flat(x));
    const n = arrs[0] ? arrs[0].length : 0;
    if (arrs.some((x) => x.length !== n)) return err(ERR.VALUE);
    let s = 0;
    for (let i = 0; i < n; i++) s += arrs.reduce((p, x) => p * (typeof x[i] === 'number' ? x[i] : 0), 1);
    return s;
  },
  VLOOKUP: ([val, table, col, approx]) => {
    if (isErr(val)) return val;
    if (!Array.isArray(table)) return err(ERR.NA);
    const c = toNum(col);
    if (isErr(c) || c < 1 || c > table[0].length) return err(ERR.REF);
    const exact = approx === false || approx === 0;
    if (exact) {
      const at = indexedFind(table, 0, val);
      if (at != null) return at >= 0 ? (table[at][c - 1] ?? 0) : err(ERR.NA);
      const hit = exactMatcher(val);
      const row = table.find((r) => hit(r[0]));
      return row ? (row[c - 1] ?? 0) : err(ERR.NA);
    }
    const i = approxIndex(table.map((r) => r[0]), val);
    return i >= 0 ? (table[i][c - 1] ?? 0) : err(ERR.NA);
  },
  MATCH: ([val, range, type]) => {
    const list = flat(range);
    const t = type == null ? 1 : toNum(type);
    if (isErr(val)) return val;
    if (t === 0) {
      // столбец или строка запомненного диапазона — по индексу
      const one = Array.isArray(range) && range.rangeEntry && (range[0].length === 1 ? indexedFind(range, 0, val) : range.length === 1 ? indexedFind(range, 0, val, true) : null);
      if (one != null) return one < 0 ? err(ERR.NA) : one + 1;
      const hit = exactMatcher(val); const i = list.findIndex(hit); return i < 0 ? err(ERR.NA) : i + 1;
    }
    const found = approxIndex(list, val, t < 0);
    return found < 0 ? err(ERR.NA) : found + 1;
  },
  INDEX: ([range, r, c]) => {
    if (isErr(range)) return range;
    const t = Array.isArray(range) ? range : [[range]];
    if (t.length === 1 && c == null && r != null) { c = r; r = 1; } // одна строка — номер столбца
    const i = toNum(r ?? 1); const j = toNum(c ?? 1);
    if (firstErr(i, j)) return firstErr(i, j);
    const row = t[(i || 1) - 1];
    return row && (j || 1) <= row.length ? (row[(j || 1) - 1] ?? 0) : err(ERR.REF);
  },
  // --- массивы ---
  ROWS: ([x]) => (isErr(x) ? x : Array.isArray(x) ? (x.full ? x.full.rows : x.length) : 1),
  COLUMNS: ([x]) => (isErr(x) ? x : Array.isArray(x) ? (x.full ? x.full.cols : (x[0] || []).length) : 1),
  SEQUENCE: ([rows, cols, start, step]) => {
    const r = toNum(rows ?? 1); const c = toNum(cols ?? 1); const s0 = toNum(start ?? 1); const st = toNum(step ?? 1);
    if (firstErr(r, c, s0, st)) return firstErr(r, c, s0, st);
    const R = Math.trunc(r); const C = Math.trunc(c);
    if (R < 1 || C < 1) return err(ERR.CALC);
    if (R * C > 1_048_576) return err(ERR.NUM);
    return Array.from({ length: R }, (_, i) => Array.from({ length: C }, (__, j) => s0 + (i * C + j) * st));
  },
  TRANSPOSE: ([x]) => (isErr(x) ? x : !Array.isArray(x) ? x : x[0].map((_, j) => x.map((row) => row[j]))),
  FILTER: ([arr, inc, empty]) => {
    if (isErr(arr)) return arr;
    const a = toArr(arr); const m = toArr(inc);
    const byRow = m.length === a.length && (m[0] || []).length === 1;
    const byCol = m.length === 1 && (m[0] || []).length === (a[0] || []).length && !byRow;
    if (!byRow && !byCol) return err(ERR.VALUE);
    const keep = (v) => { if (isErr(v)) return v; const b = toBool(v); return isErr(b) ? b : b; };
    let out;
    if (byRow) {
      out = [];
      for (let i = 0; i < a.length; i++) { const k = keep(m[i][0]); if (isErr(k)) return k; if (k) out.push(a[i]); }
    } else {
      const cols = [];
      for (let j = 0; j < a[0].length; j++) { const k = keep(m[0][j]); if (isErr(k)) return k; if (k) cols.push(j); }
      out = cols.length ? a.map((row) => cols.map((j) => row[j])) : [];
    }
    if (!out.length) return empty !== undefined ? empty : err(ERR.CALC);
    return out;
  },
  SORT: ([arr, idx, order, byCol]) => {
    if (isErr(arr)) return arr;
    const a = toArr(arr);
    const k = toNum(idx ?? 1) - 1; const desc = toNum(order ?? 1) === -1;
    if (isErr(k)) return k;
    const cols = !!(byCol && toBool(byCol) === true);
    const data = cols ? FN.TRANSPOSE([a]) : a;
    if (k < 0 || k >= (data[0] || []).length) return err(ERR.VALUE);
    const sorted = stableSort(data, (x, y) => (desc ? -1 : 1) * sortCompare(x[k], y[k]));
    return cols ? FN.TRANSPOSE([sorted]) : sorted;
  },
  SORTBY: (args) => {
    const [arr, ...rest] = args;
    if (isErr(arr)) return arr;
    const a = toArr(arr);
    const keys = [];
    for (let i = 0; i < rest.length; i += 2) {
      const by = toArr(rest[i]).map((row) => row[0]);
      if (by.length !== a.length) return err(ERR.VALUE);
      keys.push({ by, desc: toNum(rest[i + 1] ?? 1) === -1 });
    }
    const idx = stableSort(a.map((_, i) => i), (x, y) => {
      for (const k of keys) { const r = sortCompare(k.by[x], k.by[y]); if (r) return k.desc ? -r : r; }
      return 0;
    });
    return idx.map((i) => a[i]);
  },
  UNIQUE: ([arr, byCol, once]) => {
    if (isErr(arr)) return arr;
    const cols = byCol != null && toBool(byCol) === true;
    const a = cols ? FN.TRANSPOSE([toArr(arr)]) : toArr(arr);
    const keyOf = (row) => JSON.stringify(row.map((v) => (typeof v === 'string' ? v.toLowerCase() : v)));
    const count = new Map();
    for (const row of a) count.set(keyOf(row), (count.get(keyOf(row)) || 0) + 1);
    const seen = new Set();
    const out = [];
    for (const row of a) {
      const k = keyOf(row);
      if (seen.has(k)) continue;
      seen.add(k);
      if (once != null && toBool(once) === true && count.get(k) > 1) continue;
      out.push(row);
    }
    if (!out.length) return err(ERR.CALC);
    return cols ? FN.TRANSPOSE([out]) : out;
  },
  RAND: () => Math.random(),
  RANDBETWEEN: ([lo, hi]) => {
    const a = toNum(lo); const b = toNum(hi);
    if (firstErr(a, b)) return firstErr(a, b);
    const x = Math.ceil(a); const y = Math.floor(b);
    return y < x ? err(ERR.NUM) : x + Math.floor(Math.random() * (y - x + 1));
  },
  RANDARRAY: ([rows, cols, min, max, whole]) => {
    const r = Math.trunc(toNum(rows ?? 1)); const c = Math.trunc(toNum(cols ?? 1));
    const lo = toNum(min ?? 0); const hi = toNum(max ?? 1);
    if (firstErr(r, c, lo, hi)) return firstErr(r, c, lo, hi);
    if (r < 1 || c < 1 || hi < lo) return err(ERR.VALUE);
    const int = whole != null && toBool(whole) === true;
    return Array.from({ length: r }, () => Array.from({ length: c }, () => (int ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo))));
  },
  RANK: ([val, range, order]) => {
    const n = toNum(val);
    if (isErr(n)) return n;
    const list = flat(range).filter((x) => typeof x === 'number');
    const desc = !order;
    return 1 + list.filter((x) => (desc ? x > n : x < n)).length;
  },
};
// Библиотека функций (fnlib.js) — помощники движка передаются ей явно
const HELPERS = {
  ERR, err, isErr, toNum, toText, toBool, flat, nums, firstErr, criterion, compare, exactMatcher, wildcard, dateToSerial, serialToDate,
  toArr, numText, textToNum, round, colName, FN, approxIndex, errText, isFn: (v) => isFn(v), callFn: (...a) => callFn(...a), broadcast: (...a) => broadcast(...a),
  evaluate: (...a) => evaluate(...a), parse: (...a) => parse(...a), localize: (...a) => localize(...a), sheetPrefix: (...a) => sheetPrefix(...a),
};
Object.assign(FN, functions(HELPERS));
Object.assign(FN, functions2(HELPERS));
Object.assign(FN, functions3(HELPERS));
Object.assign(FN, functions4(HELPERS));
Object.assign(FN, functions5(HELPERS));
Object.assign(FN, functions6(HELPERS));
Object.assign(FN, functionsEts(HELPERS));
const LAZY = lazyFunctions(HELPERS);
export const FUNCTIONS = [...Object.keys(FN), 'LAMBDA', 'ISOMITTED'];
const SCALAR_FN = new Set(['ABS', 'SIGN', 'INT', 'SQRT', 'POWER', 'MOD', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'NOT', 'LEN', 'UPPER', 'LOWER',
  'TRIM', 'LEFT', 'RIGHT', 'MID', 'FIND', 'SUBSTITUTE', 'REPT', 'VALUE', 'ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'DATE', 'YEAR',
  'MONTH', 'DAY']);

// ---------- Вычисление ----------
// ctx: { get(sheet, r, c) → значение ячейки (с вычислением формул), sheet — имя текущего листа }
export function evaluate(node, ctx) {
  switch (node.t) {
    case 'num': case 'str': case 'bool': return node.v;
    case 'err': return err(node.v);
    case 'empty': return null;
    case 'name': {
      const v = ctx.name ? ctx.name(node.v) : err(ERR.NAME);
      // имя встроенной функции как значение (BYROW(A1:C3;СУММ)) — «функция без аргументов»
      if (isErr(v) && v.error === ERR.NAME) {
        const up = node.v.toUpperCase();
        const en = RU[up] || up;
        if (FN[en] || LAZY[en]) return { fnref: en };
      }
      return v;
    }
    case 'call': return applyFn(evaluate(node.fn, ctx), node.args, ctx);
    case 'ref': return ctx.get(node.sheet || ctx.sheet, node.r, node.c);
    case 'range': {
      // трёхмерная ссылка — те же ячейки на листах от первого до последнего (строки складываются)
      const sheets = node.sheet2 ? (ctx.sheetsBetween ? ctx.sheetsBetween(node.sheet, node.sheet2) : null) : [node.sheet || ctx.sheet];
      if (!sheets) return err(ERR.REF);
      const out = [];
      for (const sh of sheets) {
        // целые столбцы/строки — только занятая часть листа (дальше пусто); полный размер — в out.full
        let { r2, c2 } = node;
        if (node.whole && ctx.extent) { const e = ctx.extent(sh); r2 = Math.min(r2, e.r); c2 = Math.min(c2, e.c); }
        if ((r2 - node.r1 + 1) * (c2 - node.c1 + 1) > 2_000_000) return err(ERR.NUM);
        // зависимость — весь диапазон (у целых столбцов — до конца листа), ячейки по отдельности не записываются
        if (ctx.depRange) ctx.depRange(sh, node.r1, node.c1, node.r2, node.c2);
        // уже собранный в этом проходе диапазон — копия строк (функции не испортят запомненное)
        const ck = sheets.length === 1 && ctx.rangeCache && (r2 - node.r1 + 1) * (c2 - node.c1 + 1) >= 16 ? `${String(sh).toLowerCase()}|${node.r1}|${node.c1}|${r2}|${c2}` : null;
        const hit = ck && ctx.rangeCache.get(ck);
        if (hit) {
          if (ctx.depDone) ctx.depDone();
          const copy = hit.arr.map((row) => row.slice());
          copy.rangeEntry = hit;
          if (node.whole) copy.full = { rows: node.r2 - node.r1 + 1, cols: node.c2 - node.c1 + 1 };
          return copy;
        }
        const taint0 = ck ? ctx.rangeTaint() : 0;
        try {
          for (let r = node.r1; r <= r2; r++) {
            const row = [];
            for (let c = node.c1; c <= c2; c++) row.push(ctx.get(sh, r, c));
            out.push(row);
          }
        } finally { if (ctx.depDone) ctx.depDone(); }
        if (ck && ctx.rangeTaint() === taint0) {
          const entry = { arr: out.map((row) => row.slice()), idx: new Map() };
          ctx.rangeCache.set(ck, entry);
          out.rangeEntry = entry;
        }
      }
      if (node.whole) out.full = { rows: (node.r2 - node.r1 + 1) * sheets.length, cols: node.c2 - node.c1 + 1 };
      return out;
    }
    case 'array': return node.v.map((row) => [...row]);
    case 'value': return node.v; // уже вычисленная часть (пошаговое вычисление)
    case 'spill': return ctx.spill ? ctx.spill(node.sheet || ctx.sheet, node.r, node.c) : err(ERR.REF);
    case 'pct': return lift1(evaluate(node.a, ctx), ctx, (x) => { const n = toNum(x); return isErr(n) ? n : n / 100; });
    case 'un': return lift1(evaluate(node.a, ctx), ctx, (x) => { const n = toNum(x); return isErr(n) ? n : node.op === '-' ? -n : n; });
    case 'bin': {
      const av = evaluate(node.a, ctx);
      const bv = evaluate(node.b, ctx);
      // поэлементно, если хотя бы один операнд — массив/диапазон (как в Excel с динамическими массивами)
      if (Array.isArray(av) || Array.isArray(bv)) { markArray(ctx); return broadcast([av, bv], ([x, y]) => binop(node.op, x, y)); }
      return binop(node.op, av, bv);
    }
    case 'fn': {
      // «ленивые» функции: вычисляют только нужную ветку (с массивом-условием — обе, поэлементно)
      if (node.name === 'IF') {
        const cv = evaluate(node.args[0] || { t: 'empty' }, ctx);
        const pick = (c, yes, no) => {
          const b = toBool(c);
          if (isErr(b)) return b;
          if (b) return node.args[1] ? (yes ?? 0) : true;
          return node.args[2] ? (no ?? 0) : false;
        };
        if (Array.isArray(cv)) {
          markArray(ctx);
          const yes = node.args[1] ? evaluate(node.args[1], ctx) : true;
          const no = node.args[2] ? evaluate(node.args[2], ctx) : false;
          return broadcast([cv, yes, no], ([c, y, n]) => pick(c, y, n));
        }
        const b = toBool(cv);
        if (isErr(b)) return b;
        const branch = b ? node.args[1] : node.args[2];
        if (!branch) return b;
        return evaluate(branch, ctx) ?? 0;
      }
      if (node.name === 'LAMBDA') return makeLambda(node.args, ctx);
      if (node.name === 'ISOMITTED') { const a = node.args[0]; return !!(a && a.t === 'name' && ctx.omitted && ctx.omitted(a.v)); }
      if (LAZY[node.name]) return fin(LAZY[node.name](node.args, ctx, evaluate));
      // имя, определённое как LAMBDA (диспетчер имён), вызывается как функция
      if (!FN[node.name] && ctx.name) {
        const lv = ctx.name(node.text || node.name);
        if (lv && (lv.lambda || lv.fnref)) return applyFn(lv, node.args, ctx);
      }
      if (node.name === 'IFERROR') {
        const v = evaluate(node.args[0], ctx);
        const alt = () => evaluate(node.args[1] || { t: 'str', v: '' }, ctx);
        if (Array.isArray(v)) { if (!v.flat().some(isErr)) return v; markArray(ctx); return broadcast([v, alt()], ([x, y]) => (isErr(x) ? y : x)); }
        return isErr(v) ? alt() : v;
      }
      const f = FN[node.name];
      if (!f) return err(ERR.NAME);
      const args = node.args.map((a) => evaluate(a, ctx));
      // функции от отдельных значений с массивом в аргументе — поэлементно
      if (SCALAR_FN.has(node.name) && args.some(Array.isArray)) {
        markArray(ctx);
        return broadcast(args, (xs) => fin(f(xs)));
      }
      return fin(f(args, ctx, node.args));
    }
    default: return err(ERR.VALUE);
  }
}
const fin = (v) => (typeof v === 'number' && !isFinite(v) ? err(ERR.NUM) : v);

// ---------- LAMBDA ----------
// Значение-функция: { lambda: { params, body, ctx } } (контекст захватывается) или { fnref: 'SUM' } — встроенная функция
export const isFn = (v) => !!v && typeof v === 'object' && !Array.isArray(v) && (!!v.lambda || !!v.fnref);
function makeLambda(nodes, ctx) {
  if (!nodes.length) return err(ERR.VALUE);
  const params = nodes.slice(0, -1);
  if (params.some((p) => p.t !== 'name')) return err(ERR.VALUE);
  return { lambda: { params: params.map((p) => p.v.toLowerCase()), body: nodes[nodes.length - 1], ctx } };
}
let lambdaDepth = 0;
// Вызвать функцию-значение: args — узлы (вычисляются в ctx) или готовые значения ({ t: 'value', v })
export function applyFn(f, args, ctx) {
  if (isErr(f)) return f;
  if (!isFn(f)) return err(ERR.VALUE);
  const vals = args.map((a) => (a.t === 'empty' ? undefined : evaluate(a, ctx)));
  if (f.fnref) {
    const lz = LAZY[f.fnref];
    if (lz) return fin(lz(vals.map((v) => ({ t: 'value', v })), ctx, evaluate));
    return fin(FN[f.fnref](vals.map((v) => (v === undefined ? null : v)), ctx, []));
  }
  const { params, body, ctx: home } = f.lambda;
  if (vals.length > params.length) return err(ERR.VALUE);
  const bound = new Map(); const omitted = new Set();
  params.forEach((p, i) => { if (vals[i] === undefined) { omitted.add(p); bound.set(p, null); } else bound.set(p, vals[i]); });
  const scope = Object.create(home);
  scope.name = (n) => { const k = String(n).toLowerCase(); return bound.has(k) ? bound.get(k) : home.name ? home.name(n) : err(ERR.NAME); };
  scope.omitted = (n) => omitted.has(String(n).toLowerCase());
  if (lambdaDepth > 1000) return err(ERR.NUM);
  lambdaDepth++;
  try { return evaluate(body, scope); } catch { return err(ERR.NUM); } finally { lambdaDepth--; }
}
// вызвать с готовыми значениями
export const callFn = (f, values, ctx) => applyFn(f, values.map((v) => ({ t: 'value', v })), ctx);
const markArray = (ctx) => { if (ctx.markArray) ctx.markArray(); };

// Бинарная операция над двумя значениями
function binop(op, a, b) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (op === '&') return toText(a) + toText(b);
  if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
    const r = compare(a, b);
    return { '=': r === 0, '<>': r !== 0, '<': r < 0, '>': r > 0, '<=': r <= 0, '>=': r >= 0 }[op];
  }
  const x = toNum(a);
  const y = toNum(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  switch (op) {
    case '+': return x + y;
    case '-': return x - y;
    case '*': return x * y;
    case '/': return y === 0 ? err(ERR.DIV0) : x / y;
    case '^': { const v = x ** y; return isFinite(v) ? v : err(ERR.NUM); }
    default: return err(ERR.VALUE);
  }
}
// Поэлементное вычисление как в Excel: размер — наибольший; строка/столбец из одного элемента растягивается,
// за пределами меньшего массива — #Н/Д
export function broadcast(vals, f) {
  const dims = vals.map((v) => (Array.isArray(v) ? [v.length, v[0] ? v[0].length : 0] : [1, 1]));
  const rows = Math.max(...dims.map((d) => d[0]));
  const cols = Math.max(...dims.map((d) => d[1]));
  const at = (v, d, i, j) => {
    if (!Array.isArray(v)) return v;
    const ii = d[0] === 1 ? 0 : i; const jj = d[1] === 1 ? 0 : j;
    return ii < d[0] && jj < d[1] ? v[ii][jj] : err(ERR.NA);
  };
  const out = [];
  for (let i = 0; i < rows; i++) {
    const row = [];
    for (let j = 0; j < cols; j++) row.push(f(vals.map((v, k) => at(v, dims[k], i, j))));
    out.push(row);
  }
  return out;
}
function lift1(v, ctx, f) {
  if (!Array.isArray(v)) return f(v);
  markArray(ctx);
  return v.map((row) => row.map(f));
}

// ---------- Переписывание формул ----------
// Привести к хранимому виду: английские имена функций, «,» между аргументами
export function canonical(src) {
  let depth = 0; // внутри {…} разделители — столбцы «,» и строки «;» — сохраняются
  const out = tokenize(src).map((t) => {
    if (t.type === '{') depth++;
    if (t.type === '}') depth = Math.max(0, depth - 1);
    return t.type === 'func' ? t.name : t.type === 'sep' ? (depth ? t.text : ',') : t.type === 'bool' ? (t.value ? 'TRUE' : 'FALSE')
      : t.type === 'err' ? t.code : t.type === 'name' ? t.name : t.text;
  }).join('');
  // ANCHORARRAY(A1) из файла — ссылка на «разлив» A1#
  return out.replace(/ANCHORARRAY\(((?:(?:'(?:[^']|'')+'|[^()'!,]+)!)?\$?[A-Za-z]{1,3}\$?\d+)\)/g, '$1#');
}
// Хранимая формула → запись в файл: префиксы _xlfn. у функций новее Excel 2007
// Быстрая проверка: в формуле нет ни «разлива» (#), ни вызова функции из FUTURE / LET / LAMBDA — она пишется как есть
// (разбор на лексемы каждой из 400 тыс. формул книги — секунды при сохранении)
const CALL = /([A-Za-z_][\w.]*)\s*\(/g;
function plainForFile(src) {
  if (src.includes('#')) return false;
  CALL.lastIndex = 0;
  let m;
  while ((m = CALL.exec(src))) { const n = m[1].toUpperCase(); if (FUTURE.has(n) || n === 'LET' || n === 'LAMBDA') return false; }
  return true;
}
export function fileFormula(src) {
  if (plainForFile(src)) return src;
  const toks = tokenize(src);
  // имена параметров LET (нечётные аргументы, кроме последнего) — в файле с префиксом _xlpm.
  const params = new Set();
  toks.forEach((t, i) => {
    if (t.type !== 'func' || (t.name !== 'LET' && t.name !== 'LAMBDA')) return;
    const isLambda = t.name === 'LAMBDA';
    let depth = 0; let arg = 0;
    for (let j = i + 1; j < toks.length; j++) {
      const u = toks[j];
      if (u.type === '(') { depth++; continue; }
      if (u.type === ')') { if (--depth === 0) break; continue; }
      if (depth === 1 && u.type === 'sep') { arg++; continue; }
      // LET: имена — чётные аргументы; LAMBDA: все аргументы-имена (кроме тела — оно не просто имя)
      if (depth === 1 && u.type === 'name' && (isLambda || arg % 2 === 0)) params.add(u.name.toLowerCase());
    }
  });
  return toks.map((t) => (t.type === 'func' && FUTURE.has(t.name) ? `_xlfn.${XLWS.has(t.name) ? '_xlws.' : ''}${t.name}`
    : t.type === 'ref' && t.spill ? `_xlfn.ANCHORARRAY(${t.text.slice(0, -1)})`
      : t.type === 'name' && params.has(t.name.toLowerCase()) ? `_xlpm.${t.name}` : t.text)).join('');
}

// Формула, введённая пользователем (русская запись): «1,5» между цифрами — десятичная запятая,
// разделитель аргументов — «;» (запятая между не-числами тоже принимается). Результат — каноническая запись.
// Константа массива по-русски: {1,5;2:3;4} → {1.5,2;3,4} — в скобках «,» десятичная, «;» — столбцы, «:» — строки.
// Переводится до разбора на токены (иначе «2:3» прочиталось бы как ссылка на строки).
function arraysFromRu(src) {
  let out = ''; let q = false; let depth = 0;
  for (const ch of String(src)) {
    if (ch === '"') q = !q;
    if (!q && ch === '{') depth++;
    if (!q && ch === '}') depth = Math.max(0, depth - 1);
    out += !q && depth ? ({ ',': '.', ';': ',', ':': ';' }[ch] || ch) : ch;
  }
  return out;
}
export function fromUser(src) {
  const t = tokenize(arraysFromRu(src));
  let out = '';
  let depth = 0;
  for (let i = 0; i < t.length; i++) {
    if (t[i].type === '{') depth++;
    if (t[i].type === '}') depth = Math.max(0, depth - 1);
    if (!depth && t[i].type === 'num' && /^\d+$/.test(t[i].text) && t[i + 1] && t[i + 1].text === ',' && t[i + 2] && t[i + 2].type === 'num' && /^\d/.test(t[i + 2].text)) {
      out += `${t[i].text}.${t[i + 2].text}`;
      i += 2;
    } else out += t[i].text;
  }
  return canonical(out);
}

// Каноническая формула → русская запись для строки формул: ЕСЛИ(A1>1,5;"да";"нет")
// Русские имена: главный источник — ответ Excel (fnnames.js), ручной словарь RU — варианты ввода (ЧЕТН, СЧЕТ…)
const EN_RU = { ...EXCEL_RU };
for (const [ru, en] of Object.entries(RU)) if (!EN_RU[en] && en !== 'TRUE' && en !== 'FALSE') EN_RU[en] = ru;
export const ruName = (en) => EN_RU[en] || en;
// все имена функций для подсказок: русское → английское (русские и английские варианты ввода)
export const functionNames = () => ({ ...Object.fromEntries(Object.keys(FN).map((en) => [en, en])), ...RU });

// Токены с позициями в тексте: [{ ...токен, start, end }]
export function tokenizeWithPos(src) {
  let at = 0;
  return tokenize(src).map((t) => { const out = { ...t, start: at, end: at + t.text.length }; at += t.text.length; return out; });
}
// F4: ссылка под курсором — $A$1 → A$1 → $A1 → A1 → $A$1 (у диапазона — оба конца). { text, caret } или null
export function cycleRef(text, caret) {
  const toks = tokenizeWithPos(text);
  const i = toks.findIndex((t) => (t.type === 'ref' || t.type === 'cols' || t.type === 'rows') && caret >= t.start && caret <= t.end);
  if (i < 0) return null;
  let a = i; let b = i;
  if (toks[i].type === 'ref') {
    if (toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref') b = i + 2;
    else if (toks[i - 1] && toks[i - 1].type === ':' && toks[i - 2] && toks[i - 2].type === 'ref') a = i - 2;
  }
  const next = (absR, absC) => (absR && absC ? [true, false] : absR ? [false, true] : absC ? [false, false] : [true, true]);
  const first = toks[a];
  const [nr, nc] = first.type === 'ref' ? next(first.absR, first.absC) : next(first.abs1, first.abs1);
  const fmt = (t) => {
    const pre = t.text.slice(0, t.text.lastIndexOf('!') + 1);
    if (t.type === 'ref') return `${pre}${nc ? '$' : ''}${colName(t.c)}${nr ? '$' : ''}${t.r + 1}`;
    const abs = t.type === 'cols' ? nc : nr;
    const parts = t.text.slice(pre.length).split(':').map((p) => p.replace(/\$/g, ''));
    return `${pre}${abs ? '$' : ''}${parts[0]}:${abs ? '$' : ''}${parts[1]}`;
  };
  let out = '';
  for (let j = a; j <= b; j++) out += toks[j].type === 'ref' || toks[j].type === 'cols' || toks[j].type === 'rows' ? fmt(toks[j]) : toks[j].text;
  const start = toks[a].start; const end = toks[b].end;
  return { text: text.slice(0, start) + out + text.slice(end), caret: start + out.length };
}
export function localize(src) {
  let depth = 0;
  return tokenize(src).map((t) => {
    if (t.type === '{') depth++;
    if (t.type === '}') depth = Math.max(0, depth - 1);
    if (t.type === 'func') return EN_RU[t.name] || t.text;
    if (t.type === 'sep') return depth && t.text === ';' ? ':' : ';';
    if (t.type === 'num') return t.text.replace('.', ',');
    if (t.type === 'bool') return t.value ? 'ИСТИНА' : 'ЛОЖЬ';
    if (t.type === 'err') return errText(t.code);
    return t.text;
  }).join('');
}

// Префикс листа для записи: без кавычек, если имя простое (у трёхмерной ссылки — оба имени)
const PLAIN_SHEET = /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*$/;
export function sheetPrefix(sheet, sheet2 = null) {
  if (!sheet) return '';
  const both = sheet2 ? `${sheet}:${sheet2}` : sheet;
  return PLAIN_SHEET.test(sheet) && (!sheet2 || PLAIN_SHEET.test(sheet2)) ? `${both}!` : `'${both.replace(/'/g, "''")}'!`;
}
const refText = (t, r, c) => `${sheetPrefix(t.sheet, t.sheet2)}${t.absC ? '$' : ''}${colName(c)}${t.absR ? '$' : ''}${r + 1}`;

// Сдвиг ссылок: fn({ r, c, absR, absC, sheet }) → { r, c } | null (null → #REF!)
// Диапазоны «A1:B3» переписываются по концам; если удалён конец — диапазон сжимается (endFn), полностью — #REF!
export function rewriteRefs(src, fn) {
  const toks = tokenize(src);
  let out = '';
  // имя второго листа трёхмерной ссылки — через fn (переименование листа)
  const sheet2Of = (t) => (t.sheet2 ? ((fn({ ...t, sheet: t.sheet2, r: 0, c: 0 }, 'sheet2') || {}).sheet ?? t.sheet2) : null);
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'cols' || t.type === 'rows') {
      // целые столбцы/строки: меняется только своя ось
      const col = t.type === 'cols';
      const a = fn({ ...t, whole: col ? 'col' : 'row', r: col ? 0 : t.r1, c: col ? t.c1 : 0, absR: col ? true : t.abs1, absC: col ? t.abs1 : true }, 'start');
      const b = fn({ ...t, whole: col ? 'col' : 'row', r: col ? 0 : t.r2, c: col ? t.c2 : 0, absR: col ? true : t.abs2, absC: col ? t.abs2 : true }, 'end');
      if (!a || !b || (col ? b.c < a.c : b.r < a.r)) { out += '#REF!'; continue; }
      const pre = sheetPrefix(a && 'sheet' in a ? a.sheet : t.sheet, sheet2Of(t));
      out += col ? `${pre}${t.abs1 ? '$' : ''}${colName(a.c)}:${t.abs2 ? '$' : ''}${colName(b.c)}`
        : `${pre}${t.abs1 ? '$' : ''}${a.r + 1}:${t.abs2 ? '$' : ''}${b.r + 1}`;
      continue;
    }
    if (t.type !== 'ref') { out += t.text; continue; }
    if (t.sheet2) { // трёхмерная ссылка: меняется только имя листа
      const p = fn(t, 'single');
      out += refText({ ...t, sheet: p && 'sheet' in p ? p.sheet : t.sheet, sheet2: sheet2Of(t) }, t.r, t.c);
      if (toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref') { out += `:${refText({ ...toks[i + 2], sheet: null }, toks[i + 2].r, toks[i + 2].c)}`; i += 2; }
      continue;
    }
    // диапазон?
    const isRange = toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref';
    if (isRange) {
      const u = toks[i + 2];
      const a = fn(t, 'start');
      const b = fn({ ...u, sheet: u.sheet || t.sheet }, 'end');
      const ta = a && 'sheet' in a ? { ...t, sheet: a.sheet } : t; // fn может сменить имя листа (переименование)
      out += a && b && (a.r <= b.r && a.c <= b.c) ? `${refText(ta, a.r, a.c)}:${refText({ ...u, sheet: null }, b.r, b.c)}` : '#REF!';
      i += 2;
      continue;
    }
    const p = fn(t, 'single');
    out += p ? refText(p && 'sheet' in p ? { ...t, sheet: p.sheet } : t, p.r, p.c) : '#REF!';
  }
  return out;
}

// Копирование формулы со смещением (dr, dc): относительные ссылки сдвигаются, абсолютные ($) — нет
export function offsetFormula(src, dr, dc) {
  return rewriteRefs(src, (t) => {
    const r = t.absR || t.whole === 'col' ? t.r : t.r + dr;
    const c = t.absC || t.whole === 'row' ? t.c : t.c + dc;
    return r < 0 || c < 0 ? null : { r, c };
  });
}

// Вставка/удаление строк или столбцов на листе sheet: axis 'row'|'col', at, count (>0 — вставка, <0 — удаление)
export function shiftFormula(src, { sheet, current, axis, at, count }) {
  return rewriteRefs(src, (t, part) => {
    if ((t.sheet || current) !== sheet || part === 'sheet2' || t.sheet2) return { r: t.r, c: t.c };
    if ((t.whole === 'col' && axis === 'row') || (t.whole === 'row' && axis === 'col')) return { r: t.r, c: t.c };
    const k = axis === 'row' ? 'r' : 'c';
    let v = t[k];
    if (count > 0) { if (v >= at) v += count; }
    else {
      const n = -count;
      if (v >= at + n) v -= n;
      else if (v >= at) {
        if (part === 'single') return null;
        v = part === 'start' ? at : at - 1; // конец диапазона внутри удалённого — сжать
        if (v < 0) return null;
      }
    }
    return { ...t, [k]: v };
  });
}

// Ссылки формулы (для подсветки/зависимостей): [{ sheet, r1, c1, r2, c2 }]
export function formulaRefs(src) {
  const toks = tokenize(src);
  const out = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'ref') continue;
    if (toks[i + 1] && toks[i + 1].type === ':' && toks[i + 2] && toks[i + 2].type === 'ref') {
      const u = toks[i + 2];
      out.push({ sheet: t.sheet, r1: Math.min(t.r, u.r), c1: Math.min(t.c, u.c), r2: Math.max(t.r, u.r), c2: Math.max(t.c, u.c) });
      i += 2;
    } else out.push({ sheet: t.sheet, r1: t.r, c1: t.c, r2: t.r, c2: t.c });
  }
  return out;
}
