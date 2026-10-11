// Диалог «Параметры страницы» (уголок группы «Параметры страницы», «Поля → Настраиваемые поля…», «Размер → Другие размеры
// бумаги…»), как в Word: вкладки «Поля» (поля, переплёт, ориентация, зеркальные поля), «Размер бумаги», «Источник бумаги»
// (начало раздела, колонтитулы — различать, от края). «Применить»: к текущему разделу / до конца документа / ко всему.
import React, { useState } from 'react';
import { EMPTY_HF, PAPER, TWIPS_PER_CM, paperOf, parseSectPr, patchSectPr, sectionAt, sectionList } from './editor/page.js';
import { SECTION_TYPES, sectionType, withSectionType } from './editor/breaks.js';
import { fmtNum } from './editor/paragraph.js';
import { TextSelection } from '@tiptap/pm/state';

const cm = (tw) => fmtNum(tw / TWIPS_PER_CM, 2);
const num = (s) => {
  const v = parseFloat(String(s).replace(',', '.').replace(/[^\d.-]/g, ''));
  return isFinite(v) ? v : null;
};
const LAST_KEY = 'pero.lastMargins';
export function lastMargins() {
  try { const v = JSON.parse(localStorage.getItem(LAST_KEY) || 'null'); return v && typeof v.top === 'number' ? v : null; } catch { return null; }
}
const saveLast = (m) => { try { localStorage.setItem(LAST_KEY, JSON.stringify(m)); } catch { /* нет хранилища */ } };

// Применить параметры к разделам одной транзакцией (одна отмена)
export function applyPageSetup(editor, { patch, scope = 'section', sectType = null, mirror = null, hf = null }) {
  editor.chain().focus().command(({ tr, state }) => {
    let sel = tr.selection;
    if (scope === 'forward') {
      // «до конца документа»: разрыв раздела (со следующей страницы) перед текущим абзацем верхнего уровня
      const $p = sel.$from;
      const idx = $p.index(0);
      const prev = idx > 0 ? tr.doc.child(idx - 1) : null;
      if ($p.depth >= 1 && idx > 0 && !(prev.attrs && prev.attrs.sectPr)) {
        const start = $p.before(1);
        const cur = sectionAt({ selection: sel, doc: tr.doc });
        const at = start - prev.nodeSize; // предыдущий абзац получает sectPr текущего раздела — он закрывает старый раздел
        if (prev.type.name === 'paragraph' || prev.type.name === 'heading') {
          tr.setNodeMarkup(at, undefined, { ...prev.attrs, sectPr: cur.sectPr || withSectionType(null, 'nextPage') });
        } else {
          // перед таблицей и т.п. — пустой абзац с разрывом раздела
          tr.insert(start, state.schema.nodes.paragraph.create({ sectPr: cur.sectPr || withSectionType(null, 'nextPage') }));
        }
        sel = TextSelection.near(tr.doc.resolve(tr.mapping.map(sel.from)));
        tr.setSelection(sel);
      }
    }
    const setSect = (pos, xml) => {
      const next = patchSectPr(xml, patch);
      const typed = sectType ? withSectionType(next, sectType) : next;
      if (pos == null) tr.setDocAttribute('sectPr', typed);
      else tr.setNodeMarkup(pos, undefined, { ...tr.doc.nodeAt(pos).attrs, sectPr: typed });
    };
    if (scope === 'all') {
      tr.doc.forEach((n, pos) => { if (n.attrs && n.attrs.sectPr) setSect(pos, n.attrs.sectPr); });
      setSect(null, tr.doc.attrs.sectPr);
    } else {
      const sec = sectionAt({ selection: sel, doc: tr.doc });
      setSect(sec.pos, sec.sectPr);
    }
    if (mirror != null && mirror !== !!tr.doc.attrs.mirrorMargins) {
      tr.setDocAttribute('mirrorMargins', mirror);
      tr.setDocAttribute('mirrorChanged', true);
    }
    if (hf) tr.setDocAttribute('hf', hf);
    return true;
  }).run();
}

export default function PageSetupDialog({ editor, initialTab = 'margins', onClose }) {
  const state = editor.state;
  const sec = sectionAt(state);
  const g = parseSectPr(sec.sectPr);
  const many = sectionList(state.doc).length > 1;
  const hf0 = { ...EMPTY_HF, ...(state.doc.attrs.hf || {}) };
  const own = { ...EMPTY_HF.own, ...(hf0.own || {}) };
  const initTitlePg = own.titlePg ? !!hf0.titlePg : parseSectPr(sectionList(state.doc)[0]).titlePg;
  const [tab, setTab] = useState(initialTab);
  // исходные значения (твипы): неизменённое поле пишется ими, а не округлённым до 0,01 см (иначе 850 → 851)
  const [orig] = useState(() => ({ top: g.top, bottom: g.bottom, left: g.marLeft, right: g.right, gutter: g.gutter, header: g.header, footer: g.footer, w: g.w, h: g.h }));
  const [f, setF] = useState(() => Object.fromEntries(Object.entries(orig).map(([k, v]) => [k, cm(v)])));
  const [exact, setExact] = useState(null); // формат из списка: точные твипы { w, h }
  const [orient, setOrient] = useState(g.orient === 'landscape' ? 'landscape' : 'portrait');
  const [mirror, setMirror] = useState(!!state.doc.attrs.mirrorMargins);
  const [sectType, setSectType] = useState(sectionType(sec.sectPr));
  const [titlePg, setTitlePg] = useState(initTitlePg);
  const [evenOdd, setEvenOdd] = useState(!!hf0.evenOdd);
  const [scope, setScope] = useState(many ? 'section' : 'all');
  const [err, setErr] = useState('');
  const set = (k, v) => { setErr(''); if (k === 'w' || k === 'h') setExact(null); setF((x) => ({ ...x, [k]: v })); };
  // значение поля в твипах: не трогали — исходное
  const twOf = (k, v) => {
    if (exact && (k === 'w' || k === 'h')) return exact[k];
    if (f[k] === cm(orig[k])) return orig[k];
    if ((k === 'w' || k === 'h') && orient !== (g.orient === 'landscape' ? 'landscape' : 'portrait') && f[k] === cm(orig[k === 'w' ? 'h' : 'w'])) return orig[k === 'w' ? 'h' : 'w'];
    return Math.round(v * TWIPS_PER_CM);
  };

  // ориентация меняет местами ширину и высоту (как в Word)
  const turn = (o) => {
    if (o === orient) return;
    setOrient(o);
    setF((x) => ({ ...x, w: x.h, h: x.w }));
    setExact((e) => (e ? { w: e.h, h: e.w } : e));
  };
  const paperW = num(f.w); const paperH = num(f.h);
  const paper = paperW && paperH ? paperOf(Math.round(paperW * TWIPS_PER_CM), Math.round(paperH * TWIPS_PER_CM)) : null;
  const pickPaper = (name) => {
    const p = PAPER.find((x) => x[0] === name);
    if (!p) return;
    const [, w, h] = p;
    const land = orient === 'landscape';
    setF((x) => ({ ...x, w: cm(land ? h : w), h: cm(land ? w : h) }));
    setExact({ w: land ? h : w, h: land ? w : h });
  };

  const apply = () => {
    const t = {};
    for (const k of ['top', 'bottom', 'left', 'right', 'gutter', 'header', 'footer']) {
      const v = num(f[k]);
      if (v == null || v < 0 || v > 55) { setErr('Поля и расстояния — числа от 0 до 55 см'); return; }
      t[k] = twOf(k, v);
    }
    for (const k of ['w', 'h']) {
      const v = num(f[k]);
      if (v == null || v < 0.26 || v > 55.87) { setErr('Размер бумаги — от 0,26 до 55,87 см'); return; }
      t[k] = twOf(k, v);
    }
    if (t.left + t.gutter + t.right > t.w - 567 || t.top + t.bottom > t.h - 567) { setErr('Поля слишком большие для этого размера бумаги'); return; }
    const p = paperOf(t.w, t.h);
    const patch = {
      top: t.top, bottom: t.bottom, right: t.right, left: t.left + t.gutter, gutter: t.gutter,
      header: t.header, footer: t.footer, w: t.w, h: t.h, orient, ...(p ? { code: p[3] } : {}),
    };
    let hf = null;
    if (titlePg !== initTitlePg || evenOdd !== !!hf0.evenOdd) {
      hf = { ...hf0, own: { ...own } };
      if (titlePg !== initTitlePg) { hf.titlePg = titlePg; hf.own.titlePg = true; }
      if (evenOdd !== !!hf0.evenOdd) { hf.evenOdd = evenOdd; hf.own.evenOdd = true; }
    }
    const changedType = sectType !== sectionType(sec.sectPr);
    applyPageSetup(editor, { patch, scope, sectType: changedType || scope === 'forward' ? sectType : null, mirror, hf });
    saveLast({ top: t.top, bottom: t.bottom, left: t.left + t.gutter, right: t.right, mirror });
    onClose();
  };

  const field = (k, label, unit = 'см') => (
    <label className="pd-field"><span>{label}</span>
      <span className="pd-input"><input value={f[k]} onChange={(e) => set(k, e.target.value)} /><em>{unit}</em></span>
    </label>
  );
  // образец: лист и область текста в масштабе
  const preview = () => {
    const W = num(f.w) || 21; const H = num(f.h) || 29.7;
    const k = 120 / Math.max(W, H);
    const L = (num(f.left) || 0) + (num(f.gutter) || 0); const R = num(f.right) || 0;
    const T = num(f.top) || 0; const B = num(f.bottom) || 0;
    const page = (x, mirrored) => {
      const l = mirrored ? R : L; const r = mirrored ? L : R;
      return (
        <g key={x} transform={`translate(${x} 0)`}>
          <rect x="0" y="0" width={W * k} height={H * k} fill="#fff" stroke="#888" />
          <rect x={l * k} y={T * k} width={Math.max(1, (W - l - r) * k)} height={Math.max(1, (H - T - B) * k)} fill="#e8f1ee" stroke="#1f6f5c" strokeDasharray="3 2" />
        </g>
      );
    };
    const two = mirror;
    return (
      <svg className="ps-preview" width={two ? W * k * 2 + 4 : W * k} height={H * k} aria-hidden="true">
        {two ? [page(0, true), page(W * k + 4, false)] : page(0, false)}
      </svg>
    );
  };

  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd ps-dlg" role="dialog" aria-label="Параметры страницы" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Параметры страницы</div>
        <div className="modal-body">
          <div className="pd-tabs" role="tablist">
            {[['margins', 'Поля'], ['paper', 'Размер бумаги'], ['layout', 'Источник бумаги']].map(([k, label]) => (
              <button key={k} type="button" role="tab" className={tab === k ? 'active' : ''} onClick={() => setTab(k)}>{label}</button>
            ))}
          </div>
          {tab === 'margins' && (
            <>
              <fieldset><legend>Поля</legend>
                <div className="pd-grid">
                  {field('top', 'Верхнее')}{field('bottom', 'Нижнее')}
                  {field('left', mirror ? 'Внутри' : 'Левое')}{field('right', mirror ? 'Снаружи' : 'Правое')}
                  {field('gutter', 'Переплёт')}
                  <label className="pd-field"><span>Положение переплёта</span><select value="left" disabled><option value="left">Слева</option></select></label>
                </div>
              </fieldset>
              <fieldset><legend>Ориентация</legend>
                <div className="pd-checks ps-row">
                  <label className="pd-check"><input type="radio" checked={orient === 'portrait'} onChange={() => turn('portrait')} /> Книжная</label>
                  <label className="pd-check"><input type="radio" checked={orient === 'landscape'} onChange={() => turn('landscape')} /> Альбомная</label>
                </div>
              </fieldset>
              <fieldset><legend>Страницы</legend>
                <label className="pd-field"><span>несколько страниц</span>
                  <select value={mirror ? 'mirror' : 'normal'} onChange={(e) => setMirror(e.target.value === 'mirror')}>
                    <option value="normal">Обычный</option><option value="mirror">Зеркальные поля</option>
                  </select>
                </label>
                {mirror && <div className="hf-hint">На экране все страницы показываются как нечётные; при печати и в PDF у чётных страниц поля зеркальные.</div>}
              </fieldset>
            </>
          )}
          {tab === 'paper' && (
            <fieldset><legend>Размер бумаги</legend>
              <label className="pd-field"><span>Формат</span>
                <select value={paper ? paper[0] : ''} onChange={(e) => pickPaper(e.target.value)}>
                  {PAPER.map(([name, w, h]) => <option key={name} value={name}>{name} ({cm(w)} × {cm(h)} см)</option>)}
                  <option value="" disabled>Другой</option>
                </select>
              </label>
              <div className="pd-grid">{field('w', 'Ширина')}{field('h', 'Высота')}</div>
            </fieldset>
          )}
          {tab === 'layout' && (
            <>
              <fieldset><legend>Раздел</legend>
                <label className="pd-field"><span>Начать раздел</span>
                  <select value={sectType} onChange={(e) => setSectType(e.target.value)}>
                    {SECTION_TYPES.map(([k, label]) => <option key={k} value={k}>{label.toLowerCase()}</option>)}
                  </select>
                </label>
              </fieldset>
              <fieldset><legend>Колонтитулы</legend>
                <div className="pd-checks">
                  <label className="pd-check"><input type="checkbox" checked={evenOdd} onChange={(e) => setEvenOdd(e.target.checked)} /> различать колонтитулы чётных и нечётных страниц</label>
                  <label className="pd-check"><input type="checkbox" checked={titlePg} onChange={(e) => setTitlePg(e.target.checked)} /> различать колонтитулы первой страницы</label>
                </div>
                <div className="pd-grid">{field('header', 'От края до верхнего')}{field('footer', 'От края до нижнего')}</div>
              </fieldset>
            </>
          )}
          <div className="ps-bottom">
            <fieldset className="ps-sample"><legend>Образец</legend>{preview()}</fieldset>
            <label className="pd-field"><span>Применить</span>
              <select value={scope} onChange={(e) => setScope(e.target.value)}>
                {many && <option value="section">к текущему разделу</option>}
                <option value="forward">до конца документа</option>
                <option value="all">ко всему документу</option>
              </select>
            </label>
          </div>
          {err && <div className="pd-error" role="alert">{err}</div>}
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}
