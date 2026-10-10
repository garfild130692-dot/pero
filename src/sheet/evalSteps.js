// «Вычислить формулу» (как в Excel): формула по шагам — следующее вычисляемое выражение подчёркивается,
// «Вычислить» заменяет его значением. Порядок — слева направо, изнутри наружу; ЕСЛИ/ЕСЛИОШИБКА — сначала условие,
// потом только нужная ветка. Значения — по уже посчитанной книге (viewContext).
import { errText, evaluate, isErr, numText, parse, ruName, sheetPrefix } from './formula.js';
import { colName } from './address.js';

const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
const LITERAL = new Set(['num', 'str', 'bool', 'err', 'value', 'empty', 'array']);

// Значение для показа в формуле (как Excel: текст в кавычках, массив — {…;…:…})
export function showValue(v) {
  if (v == null) return '0';
  if (Array.isArray(v)) {
    const rows = v.slice(0, 20).map((row) => row.slice(0, 20).map(showValue).join(';'));
    return `{${rows.join(':')}${v.length > 20 ? ':…' : ''}}`;
  }
  if (isErr(v)) return errText(v.error);
  if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  return numText(v);
}
const refText = (n) => `${sheetPrefix(n.sheet || null, n.sheet2 || null)}${colName(n.c)}${n.r + 1}`;
function rangeText(n) {
  const pre = sheetPrefix(n.sheet || null, n.sheet2 || null);
  if (n.whole === 'col') return `${pre}${colName(n.c1)}:${colName(n.c2)}`;
  if (n.whole === 'row') return `${pre}${n.r1 + 1}:${n.r2 + 1}`;
  return `${pre}${colName(n.c1)}${n.r1 + 1}:${colName(n.c2)}${n.r2 + 1}`;
}

// Текст выражения по-русски; mark — узел, который подчеркнуть: он выводится между \u0001 и \u0002
export function exprText(node, mark = null, parentPrec = 0) {
  const wrap = (s) => (node === mark ? `\u0001${s}\u0002` : s);
  switch (node.t) {
    case 'num': return wrap(numText(node.v));
    case 'str': return wrap(`"${node.v.replace(/"/g, '""')}"`);
    case 'bool': return wrap(node.v ? 'ИСТИНА' : 'ЛОЖЬ');
    case 'err': return wrap(errText(node.v));
    case 'empty': return '';
    case 'value': return wrap(showValue(node.v));
    case 'array': return wrap(showValue(node.v));
    case 'name': return wrap(node.v);
    case 'ref': return wrap(refText(node));
    case 'range': return wrap(rangeText(node));
    case 'spill': return wrap(`${refText(node)}#`);
    case 'pct': return wrap(`${exprText(node.a, mark, 6)}%`);
    case 'un': return wrap(`${node.op}${exprText(node.a, mark, 6)}`);
    case 'bin': {
      const p = PREC[node.op];
      const s = `${exprText(node.a, mark, p)}${node.op}${exprText(node.b, mark, p + (node.op === '^' ? 0 : 0.5))}`;
      return wrap(p < parentPrec ? `(${s})` : s);
    }
    case 'fn': return wrap(`${ruName(node.name)}(${node.args.map((a) => exprText(a, mark, 0)).join(';')})`);
    case 'call': return wrap(`${exprText(node.fn, mark, 7)}(${node.args.map((a) => exprText(a, mark, 0)).join(';')})`);
    default: return '?';
  }
}

// Следующий шаг: узел, который вычисляется сейчас (или null — формула вычислена)
function nextNode(node) {
  if (LITERAL.has(node.t) || node.t === 'name') return node.t === 'name' ? node : null;
  // LET/LAMBDA и вызов лямбды — одним шагом (внутри имена параметров, вне их области не вычисляются)
  if (node.t === 'call' || (node.t === 'fn' && (node.name === 'LET' || node.name === 'LAMBDA'))) return node;
  if (node.t === 'fn' && (node.name === 'IF' || node.name === 'IFERROR')) {
    const first = node.args[0];
    const inner = first ? nextNode(first) : null;
    return inner || node; // условие посчитано — следующий шаг: выбор ветки
  }
  const kids = node.t === 'fn' ? node.args : node.t === 'bin' ? [node.a, node.b] : node.t === 'un' || node.t === 'pct' ? [node.a] : [];
  for (const k of kids) { const n = nextNode(k); if (n) return n; }
  return node;
}
// Выполнить шаг: вернуть новое дерево (узел заменён значением или выбранной веткой ЕСЛИ)
function replace(node, target, by) {
  if (node === target) return by;
  if (node.t === 'fn') return { ...node, args: node.args.map((a) => replace(a, target, by)) };
  if (node.t === 'bin') return { ...node, a: replace(node.a, target, by), b: replace(node.b, target, by) };
  if (node.t === 'un' || node.t === 'pct') return { ...node, a: replace(node.a, target, by) };
  return node;
}
function stepValue(node, ctx) {
  if (node.t === 'fn' && node.name === 'IF') {
    const cond = evaluate(node.args[0], ctx);
    const c = Array.isArray(cond) ? cond[0][0] : cond;
    if (isErr(c)) return { t: 'value', v: c };
    const yes = typeof c === 'boolean' ? c : typeof c === 'number' ? c !== 0 : null;
    if (yes === null) return { t: 'value', v: { error: '#VALUE!' } };
    const branch = yes ? node.args[1] : node.args[2];
    return branch && branch.t !== 'empty' ? branch : { t: 'value', v: yes };
  }
  if (node.t === 'fn' && node.name === 'IFERROR') {
    const v = evaluate(node.args[0], ctx);
    return isErr(v) ? (node.args[1] || { t: 'str', v: '' }) : { t: 'value', v };
  }
  let v = evaluate(node, ctx);
  if (node.t === 'range' && Array.isArray(v) && v.length === 1 && v[0].length === 1) v = v[0][0];
  return { t: 'value', v };
}

// Сессия вычисления: { ast, done } — exprText(ast, nextNode(ast)) для показа; step() — следующий шаг
export function startEval(formula) {
  let ast;
  try { ast = parse(formula); } catch { return null; }
  return { ast, mark: nextNode(ast) };
}
export function stepEval(state, ctx) {
  if (!state.mark) return state;
  const ast = replace(state.ast, state.mark, stepValue(state.mark, ctx));
  return { ast, mark: nextNode(ast) };
}
