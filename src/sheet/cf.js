// Условное форматирование (как в Excel). Лист: cf = [{ ranges: [{ r1, c1, r2, c2 }], rules: [правило] }].
// Правило: { type, priority, stopIfTrue, operator, formulas: [канонические], text, rank, percent, bottom,
//   aboveAverage, equalAverage, stdDev, timePeriod, dxf: { color, fill, bold, italic, underline, strike, numFmt, border },
//   scale: { cfvo: [{ type, val }], colors: ['#rrggbb'] } (цветовая шкала), bar: { cfvo, color } (гистограмма),
//   raw — исходный XML правила, которое Перо не вычисляет (наборы значков…) или не меняло }.
// cfView(wb, si) → Map("r,c" → { style, bar: { pct, color }, }) — что показать поверх стиля ячейки.
import { key } from './address.js';
import { canonical, evaluate, fileFormula, isErr, offsetFormula, parse, serialToDate } from './formula.js';
import { viewContext } from './model.js';

const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const kids = (el, name) => [...(el ? el.childNodes : [])].filter((n) => n.nodeType === 1 && (!name || (n.localName || n.nodeName) === name));
const kid = (el, name) => kids(el, name)[0] || null;
const on = (el, a, def = false) => { const v = el.getAttribute(a); return v == null || v === '' ? def : ['1', 'true'].includes(v); };

// ---------- стили правил (dxf) ----------
export function readDxf(el, readColor) {
  const out = {};
  const font = kid(el, 'font');
  if (font) {
    const v = (n) => kid(font, n);
    const flag = (n) => { const e = v(n); return !!e && !['0', 'false'].includes(e.getAttribute('val')); };
    if (flag('b')) out.bold = true;
    if (flag('i')) out.italic = true;
    if (flag('strike')) out.strike = true;
    if (v('u') && v('u').getAttribute('val') !== 'none') out.underline = true;
    const c = readColor(v('color'));
    if (c) out.color = c;
  }
  const fill = kid(el, 'fill');
  const pf = fill && kid(fill, 'patternFill');
  if (pf && pf.getAttribute('patternType') !== 'none') {
    // в dxf сплошная заливка — bgColor (так пишет Excel), иногда fgColor
    const c = readColor(kid(pf, 'bgColor')) || readColor(kid(pf, 'fgColor'));
    if (c) out.fill = c;
  }
  const nf = kid(el, 'numFmt');
  if (nf) out.numFmt = nf.getAttribute('formatCode');
  const b = kid(el, 'border');
  if (b) {
    const border = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const e = kid(b, side);
      const style = e && e.getAttribute('style');
      if (style) border[side] = { style, color: readColor(kid(e, 'color')) };
    }
    if (Object.keys(border).length) out.border = border;
  }
  return out;
}
const rgb = (c) => `rgb="FF${c.slice(1).toUpperCase()}"`;
export function dxfXml(d) {
  const font = (d.bold || d.italic || d.strike || d.underline || d.color)
    ? `<font>${d.bold ? '<b/>' : ''}${d.italic ? '<i/>' : ''}${d.strike ? '<strike/>' : ''}${d.underline ? '<u/>' : ''}${d.color ? `<color ${rgb(d.color)}/>` : ''}</font>` : '';
  const numFmt = d.numFmt ? `<numFmt numFmtId="${d.numFmtId || 164}" formatCode="${xmlEsc(d.numFmt)}"/>` : '';
  const fill = d.fill ? `<fill><patternFill><bgColor ${rgb(d.fill)}/></patternFill></fill>` : '';
  const border = d.border
    ? `<border>${['left', 'right', 'top', 'bottom'].map((s) => (d.border[s] ? `<${s} style="${d.border[s].style}">${d.border[s].color ? `<color ${rgb(d.border[s].color)}/>` : '<color auto="1"/>'}</${s}>` : '')).join('')}</border>` : '';
  return `<dxf>${font}${numFmt}${fill}${border}</dxf>`;
}

// ---------- разбор и запись ----------
const parseSqref = (s, parseRange) => String(s || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean);
// root — <worksheet>; dxfs — разобранные стили правил; readColor — цвет OOXML → '#rrggbb'
export function readCf(root, { dxfs, readColor, parseRange, serialize }) {
  const out = [];
  for (const cfEl of kids(root, 'conditionalFormatting')) {
    const ranges = parseSqref(cfEl.getAttribute('sqref'), parseRange);
    if (!ranges.length) continue;
    const rules = [];
    for (const r of kids(cfEl, 'cfRule')) {
      const type = r.getAttribute('type');
      const rule = { type, priority: +(r.getAttribute('priority') || 0) };
      if (on(r, 'stopIfTrue')) rule.stopIfTrue = true;
      for (const a of ['operator', 'text', 'timePeriod']) if (r.getAttribute(a)) rule[a] = r.getAttribute(a);
      if (r.getAttribute('rank')) rule.rank = +r.getAttribute('rank');
      if (on(r, 'percent')) rule.percent = true;
      if (on(r, 'bottom')) rule.bottom = true;
      if (type === 'aboveAverage') { rule.aboveAverage = on(r, 'aboveAverage', true); if (on(r, 'equalAverage')) rule.equalAverage = true; if (r.getAttribute('stdDev')) rule.stdDev = +r.getAttribute('stdDev'); }
      const fs = kids(r, 'formula').map((f) => canonical(f.textContent));
      if (fs.length) rule.formulas = fs;
      if (r.getAttribute('dxfId') != null && dxfs[+r.getAttribute('dxfId')]) rule.dxf = dxfs[+r.getAttribute('dxfId')];
      const cs = kid(r, 'colorScale');
      const db = kid(r, 'dataBar');
      const cfvo = (el) => kids(el, 'cfvo').map((c) => ({ type: c.getAttribute('type'), ...(c.getAttribute('val') != null ? { val: c.getAttribute('val') } : {}) }));
      if (cs) rule.scale = { cfvo: cfvo(cs), colors: kids(cs, 'color').map(readColor) };
      if (db) rule.bar = { cfvo: cfvo(db), color: readColor(kid(db, 'color')) || '#638ec6' };
      const known = ['cellIs', 'expression', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks',
        'containsErrors', 'notContainsErrors', 'duplicateValues', 'uniqueValues', 'top10', 'aboveAverage', 'timePeriod', 'colorScale', 'dataBar'];
      // исходная разметка — чтобы неизменённое правило записалось как было (с расширениями Excel)
      rule.raw = serialize(r);
      if (!known.includes(type)) rule.unknown = true;
      rules.push(rule);
    }
    if (rules.length) out.push({ ranges, rules });
  }
  return out;
}
// правило → XML (dxfIdOf — номер стиля правила в таблице dxfs)
function ruleXml(rule, dxfIdOf) {
  // неизменённое или непонятное Перо правило — исходной разметкой (наборы значков и т. п. — всегда)
  if (rule.raw && (!rule.edited || rule.unknown)) return rule.raw.replace(/\spriority="\d+"/, ` priority="${rule.priority}"`);
  const a = [`type="${rule.type}"`];
  if (rule.dxf && Object.keys(rule.dxf).length) a.push(`dxfId="${dxfIdOf(rule.dxf)}"`);
  a.push(`priority="${rule.priority}"`);
  if (rule.stopIfTrue) a.push('stopIfTrue="1"');
  if (rule.type === 'aboveAverage') { if (rule.aboveAverage === false) a.push('aboveAverage="0"'); if (rule.equalAverage) a.push('equalAverage="1"'); if (rule.stdDev) a.push(`stdDev="${rule.stdDev}"`); }
  if (rule.type === 'top10') { if (rule.percent) a.push('percent="1"'); if (rule.bottom) a.push('bottom="1"'); a.push(`rank="${rule.rank || 10}"`); }
  if (rule.operator) a.push(`operator="${rule.operator}"`);
  if (rule.text != null) a.push(`text="${xmlEsc(rule.text)}"`);
  if (rule.timePeriod) a.push(`timePeriod="${rule.timePeriod}"`);
  const fs = (rule.formulas || []).map((f) => `<formula>${xmlEsc(fileFormula(f))}</formula>`).join('');
  const cfvo = (list) => list.map((c) => `<cfvo type="${c.type}"${c.val != null ? ` val="${xmlEsc(c.val)}"` : ''}/>`).join('');
  let inner = fs;
  if (rule.scale) inner = `<colorScale>${cfvo(rule.scale.cfvo)}${rule.scale.colors.map((c) => `<color ${rgb(c)}/>`).join('')}</colorScale>`;
  if (rule.bar) inner = `<dataBar>${cfvo(rule.bar.cfvo)}<color ${rgb(rule.bar.color)}/></dataBar>`;
  return `<cfRule ${a.join(' ')}${inner ? `>${inner}</cfRule>` : '/>'}`;
}
export function cfXml(cf, dxfIdOf, rangeName) {
  return (cf || []).map((b) => `<conditionalFormatting sqref="${b.ranges.map(rangeName).join(' ')}">${b.rules.map((r) => ruleXml(r, dxfIdOf)).join('')}</conditionalFormatting>`).join('');
}
// для сравнения «менялось ли» — без исходной разметки
export const cfSignature = (cf) => JSON.stringify((cf || []).map((b) => ({ ranges: b.ranges, rules: b.rules.map(({ raw, ...r }) => r) })));

// ---------- вычисление ----------
const lc = (v) => (typeof v === 'string' ? v.toLowerCase() : v);
function cmp(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, 'ru', { sensitivity: 'base' });
  if (a == null || a === '') return cmp(typeof b === 'number' ? 0 : '', b);
  return typeof a === 'number' ? -1 : 1; // числа меньше текста
}
const interp = (c1, c2, t) => {
  const p = (c) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16));
  const a = p(c1); const b = p(c2);
  return `#${a.map((x, i) => Math.round(x + (b[i] - x) * t).toString(16).padStart(2, '0')).join('')}`;
};
const percentile = (sorted, p) => {
  if (!sorted.length) return 0;
  const x = (sorted.length - 1) * p; const lo = Math.floor(x);
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (x - lo);
};
const todaySerial = () => { const t = new Date(); return (Date.UTC(t.getFullYear(), t.getMonth(), t.getDate()) - Date.UTC(1899, 11, 30)) / 86400000; };

export function cfView(wb, si) {
  const sheet = wb.sheets[si];
  const out = new Map();
  if (!sheet.cf || !sheet.cf.length) return out;
  // ячейки, которые нужно перебрать: занятая часть листа в пределах диапазонов
  let maxR = 0; let maxC = 0;
  for (const k of sheet.cells.keys()) { const i = k.indexOf(','); maxR = Math.max(maxR, +k.slice(0, i)); maxC = Math.max(maxC, +k.slice(i + 1)); }
  const valueAt = (r, c) => { const x = sheet.cells.get(key(r, c)); if (x && (x.v != null || x.f != null)) return x.v; return sheet.spill ? (sheet.spill.get(key(r, c)) ?? null) : null; };
  const astCache = new Map();
  const evalAt = (f, r0, c0, r, c) => {
    const text = offsetFormula(f, r - r0, c - c0);
    let ast = astCache.get(text);
    if (ast === undefined) { try { ast = parse(text); } catch { ast = null; } astCache.set(text, ast); }
    if (!ast) return null;
    let v = evaluate(ast, viewContext(wb, si, r, c));
    if (Array.isArray(v)) v = v[0] && v[0][0];
    return v;
  };
  const rules = [];
  for (const b of sheet.cf) for (const rule of b.rules) rules.push({ rule, ranges: b.ranges });
  rules.sort((a, b) => a.rule.priority - b.rule.priority);
  const done = new Set(); // ячейки, где сработал stopIfTrue
  for (const { rule, ranges } of rules) {
    if (rule.unknown) continue;
    const cells = [];
    for (const rg of ranges) {
      for (let r = rg.r1; r <= Math.min(rg.r2, maxR); r++) for (let c = rg.c1; c <= Math.min(rg.c2, maxC); c++) cells.push([r, c]);
    }
    if (!cells.length) continue;
    const r0 = ranges[0].r1; const c0 = ranges[0].c1;
    const nums = () => cells.map(([r, c]) => valueAt(r, c)).filter((v) => typeof v === 'number');
    let test = null;
    switch (rule.type) {
      case 'cellIs': {
        const op = rule.operator || 'equal';
        test = (v, r, c) => {
          if (v == null || v === '') v = null;
          const a = evalAt(rule.formulas[0], r0, c0, r, c);
          const d = cmp(lc(v ?? (typeof a === 'number' ? 0 : '')), lc(a));
          switch (op) {
            case 'equal': return d === 0; case 'notEqual': return d !== 0;
            case 'greaterThan': return d > 0; case 'greaterThanOrEqual': return d >= 0;
            case 'lessThan': return d < 0; case 'lessThanOrEqual': return d <= 0;
            case 'between': case 'notBetween': {
              const b2 = evalAt(rule.formulas[1], r0, c0, r, c);
              const lo = cmp(a, b2) <= 0 ? a : b2; const hi = lo === a ? b2 : a;
              const inside = cmp(lc(v ?? 0), lc(lo)) >= 0 && cmp(lc(v ?? 0), lc(hi)) <= 0;
              return op === 'between' ? inside : !inside;
            }
            default: return false;
          }
        };
        break;
      }
      case 'expression': test = (v, r, c) => { const x = evalAt(rule.formulas[0], r0, c0, r, c); return x === true || (typeof x === 'number' && x !== 0); }; break;
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const t = String(rule.text || '').toLowerCase();
        test = (v) => {
          if (isErr(v)) return false;
          const s = v == null ? '' : String(typeof v === 'number' ? v : v).toLowerCase();
          if (rule.type === 'containsText') return s.includes(t);
          if (rule.type === 'notContainsText') return !s.includes(t);
          if (rule.type === 'beginsWith') return s.startsWith(t);
          return s.endsWith(t);
        };
        break;
      }
      case 'containsBlanks': test = (v) => v == null || (typeof v === 'string' && !v.trim()); break;
      case 'notContainsBlanks': test = (v) => !(v == null || (typeof v === 'string' && !v.trim())); break;
      case 'containsErrors': test = (v) => isErr(v); break;
      case 'notContainsErrors': test = (v) => !isErr(v); break;
      case 'duplicateValues': case 'uniqueValues': {
        const count = new Map();
        for (const [r, c] of cells) { const v = valueAt(r, c); if (v == null || v === '') continue; const k2 = JSON.stringify(lc(v)); count.set(k2, (count.get(k2) || 0) + 1); }
        test = (v) => { if (v == null || v === '') return false; const n = count.get(JSON.stringify(lc(v))); return rule.type === 'duplicateValues' ? n > 1 : n === 1; };
        break;
      }
      case 'top10': {
        const list = nums().sort((a, b) => b - a);
        let n = rule.rank || 10;
        if (rule.percent) n = Math.max(1, Math.floor((list.length * n) / 100));
        const sorted = rule.bottom ? [...list].reverse() : list;
        const edge = sorted[Math.min(n, sorted.length) - 1];
        test = (v) => typeof v === 'number' && edge !== undefined && (rule.bottom ? v <= edge : v >= edge);
        break;
      }
      case 'aboveAverage': {
        const list = nums();
        const avg = list.reduce((a, b) => a + b, 0) / (list.length || 1);
        const sd = Math.sqrt(list.reduce((a, b) => a + (b - avg) ** 2, 0) / Math.max(1, list.length - 1));
        const above = rule.aboveAverage !== false;
        const lim = rule.stdDev ? avg + (above ? 1 : -1) * rule.stdDev * sd : avg;
        test = (v) => typeof v === 'number' && (above ? (rule.equalAverage ? v >= lim : v > lim) : (rule.equalAverage ? v <= lim : v < lim));
        break;
      }
      case 'timePeriod': {
        const t0 = todaySerial(); const td = serialToDate(t0);
        const dow = (new Date(Date.UTC(td.y, td.m - 1, td.d)).getUTCDay()); // 0 — воскресенье (неделя Excel — с воскресенья)
        const monthOf = (n) => { const d = serialToDate(n); return d.y * 12 + d.m; };
        const tm = td.y * 12 + td.m;
        test = (v) => {
          if (typeof v !== 'number') return false;
          const d = Math.floor(v);
          switch (rule.timePeriod) {
            case 'today': return d === t0; case 'yesterday': return d === t0 - 1; case 'tomorrow': return d === t0 + 1;
            case 'last7Days': return d <= t0 && d > t0 - 7;
            case 'thisWeek': return d >= t0 - dow && d < t0 - dow + 7;
            case 'lastWeek': return d >= t0 - dow - 7 && d < t0 - dow;
            case 'nextWeek': return d >= t0 - dow + 7 && d < t0 - dow + 14;
            case 'thisMonth': return monthOf(d) === tm; case 'lastMonth': return monthOf(d) === tm - 1; case 'nextMonth': return monthOf(d) === tm + 1;
            default: return false;
          }
        };
        break;
      }
      case 'colorScale': case 'dataBar': {
        const spec = rule.scale || rule.bar;
        const list = nums().sort((a, b) => a - b);
        if (!list.length) break;
        const point = (c) => {
          const val = c.val != null ? +String(c.val).replace(',', '.') : NaN;
          switch (c.type) {
            case 'min': return list[0]; case 'max': return list[list.length - 1];
            case 'num': return isFinite(val) ? val : evalAt(c.val, r0, c0, r0, c0);
            case 'percent': return list[0] + ((list[list.length - 1] - list[0]) * val) / 100;
            case 'percentile': return percentile(list, val / 100);
            case 'formula': { const x = evalAt(canonical(c.val), r0, c0, r0, c0); return typeof x === 'number' ? x : list[0]; }
            default: return list[0];
          }
        };
        const pts = spec.cfvo.map(point);
        for (const [r, c] of cells) {
          const k = key(r, c);
          if (done.has(k)) continue;
          const v = valueAt(r, c);
          if (typeof v !== 'number') continue;
          const cur = out.get(k) || { style: {} };
          if (rule.scale) {
            const cols = rule.scale.colors;
            let color;
            if (pts.length === 3) {
              color = v <= pts[1] ? interp(cols[0], cols[1], pts[1] === pts[0] ? 1 : Math.max(0, (v - pts[0]) / (pts[1] - pts[0]))) : interp(cols[1], cols[2], pts[2] === pts[1] ? 1 : Math.min(1, (v - pts[1]) / (pts[2] - pts[1])));
            } else color = interp(cols[0], cols[1], pts[1] === pts[0] ? 1 : Math.min(1, Math.max(0, (v - pts[0]) / (pts[1] - pts[0]))));
            if (!cur.style.fill) cur.style = { ...cur.style, fill: color };
          } else if (!cur.bar) {
            const lo = pts[0]; const hi = pts[1];
            cur.bar = { pct: hi === lo ? 1 : Math.min(1, Math.max(0, (v - lo) / (hi - lo))) * 0.9 + 0.1, color: rule.bar.color };
          }
          out.set(k, cur);
        }
        break;
      }
      default: break;
    }
    if (!test) continue;
    for (const [r, c] of cells) {
      const k = key(r, c);
      if (done.has(k)) continue;
      const v = valueAt(r, c);
      let hit = false;
      try { hit = test(v, r, c); } catch { hit = false; }
      if (!hit) continue;
      const cur = out.get(k) || { style: {} };
      const d = rule.dxf || {};
      const st = { ...cur.style };
      for (const [p, x] of Object.entries(d)) if (st[p] === undefined) st[p] = x; // у правила с меньшим priority — первенство
      out.set(k, { ...cur, style: st });
      if (rule.stopIfTrue) done.add(k);
    }
  }
  return out;
}
