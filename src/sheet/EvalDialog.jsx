// Диалог «Вычисление формулы» (как в Excel): следующее выражение подчёркнуто, «Вычислить» заменяет его значением
import React, { useState } from 'react';
import { exprText, showValue, startEval, stepEval } from './evalSteps.js';

export default function EvalDialog({ cellName, formula, ctx, onClose }) {
  const [state, setState] = useState(() => startEval(formula));
  if (!state) {
    return (
      <div className="modal-backdrop" onMouseDown={onClose}>
        <div className="modal" role="dialog" aria-label="Вычисление формулы" onMouseDown={(e) => e.stopPropagation()}>
          <div className="modal-title">Вычисление формулы</div>
          <div className="modal-body">Формулу не удалось разобрать.</div>
          <div className="modal-actions"><button type="button" className="btn" onClick={onClose}>Закрыть</button></div>
        </div>
      </div>
    );
  }
  const text = exprText(state.ast, state.mark);
  const [pre, rest = ''] = text.split('\u0001');
  const [mid, post = ''] = rest.split('\u0002');
  const done = !state.mark;
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal eval-dlg" role="dialog" aria-label="Вычисление формулы" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && !done) setState(stepEval(state, ctx)); }}>
        <div className="modal-title">Вычисление формулы — {cellName}</div>
        <div className="modal-body">
          <div className="eval-label">{done ? 'Результат:' : 'Вычисление (подчёркнуто следующее действие):'}</div>
          <div className="eval-expr" aria-live="polite">={pre}{mid && <u>{mid}</u>}{post}</div>
          {done && state.ast.t === 'value' && <div className="eval-label">Значение: {showValue(state.ast.v)}</div>}
        </div>
        <div className="modal-actions">
          <button type="button" className="btn primary" disabled={done} autoFocus onClick={() => setState(stepEval(state, ctx))}>Вычислить</button>
          <button type="button" className="btn" onClick={() => setState(startEval(formula))}>Заново</button>
          <button type="button" className="btn" onClick={onClose}>Закрыть</button>
        </div>
      </div>
    </div>
  );
}
