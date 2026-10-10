// Подсказки при вводе формулы (как в Excel): список функций и имён по набранному началу, подпись функции
// с выделенным текущим аргументом. Используется редактором ячейки и строкой формул.
import React, { useMemo, useState } from 'react';
import { tokenizeWithPos } from './formula.js';
import { FN_INFO, FN_LIST, argIndex } from './fninfo.js';

// Что под курсором: { prefix, start } — набираемое имя; call — { info, arg } — функция, внутри которой курсор
export function formulaContext(text, caret) {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const toks = tokenizeWithPos(before);
  // незакрытая строка в кавычках — подсказок нет
  if ((before.match(/"/g) || []).length % 2) return { prefix: null, call: null };
  let prefix = null;
  const last = toks[toks.length - 1];
  if (last && (last.type === 'name' || last.type === 'bool' || (last.type === 'ref' && /^[A-Za-z]+\d*$/.test(last.text))) && last.end === caret) {
    const prev = toks[toks.length - 2];
    if (!prev || prev.type === 'op' || prev.type === '(' || prev.type === 'sep' || prev.type === 'ws' || prev.type === '{') prefix = { text: last.text, start: last.start };
  }
  // стек вызовов функций до курсора
  const stack = [];
  let braces = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'func') { stack.push({ name: t.name, arg: 0, pending: true }); continue; }
    if (t.type === '(') { if (stack.length && stack[stack.length - 1].pending) stack[stack.length - 1].pending = false; else stack.push({ name: null, arg: 0 }); continue; }
    if (t.type === ')') { stack.pop(); continue; }
    if (t.type === '{') braces++;
    if (t.type === '}') braces = Math.max(0, braces - 1);
    if (t.type === 'sep' && !braces && stack.length) stack[stack.length - 1].arg++;
  }
  const top = [...stack].reverse().find((s) => s.name && !s.pending);
  const info = top && FN_INFO[top.name];
  return { prefix, call: info ? { info, arg: top.arg } : null };
}

export function useFormulaAssist(text, names = []) {
  const [caret, setCaret] = useState(text.length);
  const [sel, setSel] = useState(0);
  const [closed, setClosed] = useState(null); // текст, при котором список закрыт Esc
  const ctx = useMemo(() => formulaContext(text, Math.min(caret, text.length)), [text, caret]);
  const items = useMemo(() => {
    if (!ctx || !ctx.prefix || closed === text) return [];
    const p = ctx.prefix.text.toLowerCase();
    const fns = FN_LIST.filter((f) => f.ru.toLowerCase().startsWith(p) || f.en.toLowerCase().startsWith(p))
      .map((f) => ({ kind: 'fn', label: f.ru, info: f }));
    const nm = names.filter((n) => n.name.toLowerCase().startsWith(p)).map((n) => ({ kind: 'name', label: n.name }));
    return [...nm, ...fns].slice(0, 14);
  }, [ctx, names, closed, text]);
  const cur = Math.min(sel, Math.max(0, items.length - 1));
  // вставить выбранное: имя функции с «(» вместо набранного начала
  const accept = (i = cur) => {
    const it = items[i];
    if (!it || !ctx || !ctx.prefix) return null;
    const ins = it.kind === 'fn' ? `${it.label}(` : it.label;
    const nextText = text.slice(0, ctx.prefix.start) + ins + text.slice(caret);
    return { text: nextText, caret: ctx.prefix.start + ins.length };
  };
  // клавиши списка: ↑/↓ — выбор, Tab — вставить, Esc — закрыть список. true — клавиша обработана
  const onKey = (e, apply) => {
    if (!items.length) return false;
    if (e.key === 'ArrowDown') { e.preventDefault(); setSel((cur + 1) % items.length); return true; }
    if (e.key === 'ArrowUp') { e.preventDefault(); setSel((cur - 1 + items.length) % items.length); return true; }
    if (e.key === 'Tab') { e.preventDefault(); const r = accept(); if (r) apply(r); setSel(0); return true; }
    if (e.key === 'Escape') { e.preventDefault(); setClosed(text); return true; }
    return false;
  };
  return { caret, setCaret: (c) => { setCaret(c); setSel(0); }, items, sel: cur, setSel, accept, onKey, call: ctx && ctx.call };
}

// Всплывающие подсказки под полем ввода (style — положение)
export function AssistPopup({ assist, style, onPick }) {
  const { items, sel, call } = assist;
  if (!items.length && !call) return null;
  const sig = call && (() => {
    const { info, arg } = call;
    const k = argIndex(info, arg);
    const parts = info.args.map((a, i) => {
      const t = a.optional ? `[${a.name}]` : a.name;
      return <span key={i} className={i === k ? 'on' : ''}>{t}</span>;
    });
    if (info.more) parts.push(<span key="more" className={k === info.args.length - 1 && arg >= info.args.length ? 'on' : ''}>…</span>);
    return <div className="fa-sig" title={info.description}>{info.ru}({parts.reduce((acc, p, i) => (i ? [...acc, '; ', p] : [p]), [])})</div>;
  })();
  return (
    <div className="fa-pop" style={style} onMouseDown={(e) => e.preventDefault()}>
      {items.length > 0 ? (
        <div className="fa-list" role="listbox">
          {items.map((it, i) => (
            <div key={`${it.kind}${it.label}`} role="option" aria-selected={i === sel} className={`fa-item${i === sel ? ' on' : ''}`}
              onMouseDown={(e) => { e.preventDefault(); onPick(i); }}>
              <span className={`fa-ico ${it.kind}`}>{it.kind === 'fn' ? 'ƒ' : '▭'}</span>{it.label}
            </div>
          ))}
          {items[sel] && items[sel].info && items[sel].info.description && <div className="fa-desc">{items[sel].info.description}</div>}
        </div>
      ) : sig}
    </div>
  );
}
