// Страница: вкладка «Макет» (поля, ориентация, размер), колонтитулы и номера страниц, CSS для печати/PDF.
import React, { useState } from 'react';
import { usePopover } from './FontControls.jsx';
import { LayoutParaGroup } from './ParaSpinners.jsx';
import { EMPTY_HF, MARGIN_PRESETS, PAPER, TWIPS_PER_CM, hfFor, hfText, pagePx, paperOf, parseSectPr, patchSectPr, sectionAt, sectionList } from './editor/page.js';
import { applyPageSetup, lastMargins } from './PageSetupDialog.jsx';
import { ColumnsButton } from './ColumnsControls.jsx';
import { LN_AUTO_DISTANCE, parseLnNum } from './editor/lineNumbers.js';
import { fmtNum } from './editor/paragraph.js';
import { SECTION_TYPES } from './editor/breaks.js';

// Изменить атрибут документа одной транзакцией (с отменой по Ctrl+Z)
const setDocAttr = (editor, name, value) => editor.chain().focus().command(({ tr }) => { tr.setDocAttribute(name, value); return true; }).run();
// Параметры страницы — раздела под курсором (как в Word): абзац с разрывом раздела или последний раздел документа
export const setPage = (editor, patch) => {
  const sec = sectionAt(editor.state);
  if (sec.pos == null) { setDocAttr(editor, 'sectPr', patchSectPr(sec.sectPr, patch)); return; }
  editor.chain().focus().command(({ tr }) => {
    const node = tr.doc.nodeAt(sec.pos);
    tr.setNodeMarkup(sec.pos, undefined, { ...node.attrs, sectPr: patchSectPr(sec.sectPr, patch) });
    return true;
  }).run();
};
const cm = (tw) => fmtNum(tw / TWIPS_PER_CM, 2);
const tw = (s) => {
  const v = parseFloat(String(s).replace(',', '.'));
  return isFinite(v) && v >= 0 && v <= 55 ? Math.round(v * TWIPS_PER_CM) : null;
};

function MenuButton({ icon, label, title, children }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title={title || label}
        onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
        {icon}<span className="rb-label">{label} ▾</span>
      </button>
      {pop.open && <div className="menu-pop" style={pop.style} role="menu">{children(() => pop.setOpen(false))}</div>}
    </div>
  );
}
const Item = ({ onClick, checked, children }) => (
  <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
    <span className="check">{checked ? '✓' : ''}</span>{children}
  </button>
);

const S = ({ children }) => (
  <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinejoin="round" aria-hidden="true">{children}</svg>
);
const MarginsIcon = () => <S><rect x="5" y="3" width="14" height="18" /><rect x="8" y="6" width="8" height="12" strokeDasharray="2 1.5" /></S>;
const OrientIcon = () => <S><rect x="3" y="7" width="13" height="10" /><rect x="11" y="3" width="9" height="13" opacity=".5" /></S>;
const SizeIcon = () => <S><rect x="5" y="3" width="14" height="18" /><path d="M8 8h8M8 12h8M8 16h5" /></S>;
const HfIcon = () => <S><rect x="5" y="3" width="14" height="18" /><path d="M5 7h14M5 17h14" /><path d="M8 5h5M8 19h3" strokeWidth="1.2" /></S>;
const BreaksIcon = () => <S><path d="M5 3v6h14V3M5 21v-6h14v6" /><path d="M3 12h3M9 12h2M14 12h2M19 12h2" /></S>;
const HyphIcon = () => <S><path d="M4 7h9M4 12h6M4 17h12" /><path d="M13 12h4" strokeWidth="2.2" /></S>;
const LnIcon = () => <S><path d="M10 7h10M10 12h10M10 17h10" /><text x="5" y="9" fontSize="5" textAnchor="middle" fill="currentColor" stroke="none">1</text><text x="5" y="14" fontSize="5" textAnchor="middle" fill="currentColor" stroke="none">2</text><text x="5" y="19" fontSize="5" textAnchor="middle" fill="currentColor" stroke="none">3</text></S>;
const NumIcon = () => <S><rect x="5" y="3" width="14" height="18" /><text x="12" y="19" fontSize="6" textAnchor="middle" fill="currentColor" stroke="none">1</text></S>;

// ---------- Колонтитулы и номер страницы (вкладки «Вставка» и «Макет») ----------
export function HfGroup({ editor, Group, onHf, extra = null }) {
  const hf = editor.state.doc.attrs.hf || EMPTY_HF;
  const pageNum = (where) => {
    const para = [{ text: '{PAGE}', align: where.endsWith('right') ? 'right' : 'center' }];
    const next = { ...EMPTY_HF, ...hf, own: { ...EMPTY_HF.own, ...(hf.own || {}) } };
    if (where.startsWith('top')) { next.header = para; next.own = { ...next.own, header: true }; }
    else { next.footer = para; next.own = { ...next.own, footer: true }; }
    setDocAttr(editor, 'hf', next);
  };
  return (
      <Group title="Колонтитулы">
        <button type="button" className="rb-btn big" title="Верхний и нижний колонтитулы" onMouseDown={(e) => e.preventDefault()} onClick={onHf}>
          <HfIcon /><span className="rb-label">Колонтитулы</span>
        </button>
        <MenuButton icon={<NumIcon />} label="Номер страницы" title="Номера страниц">
          {(close) => (
            <>
              <Item onClick={() => { close(); pageNum('bottom-center'); }}>Внизу по центру</Item>
              <Item onClick={() => { close(); pageNum('bottom-right'); }}>Внизу справа</Item>
              <Item onClick={() => { close(); pageNum('top-right'); }}>Вверху справа</Item>
            </>
          )}
        </MenuButton>
        {extra}
      </Group>
  );
}

// ---------- «Разрывы ▾»: страницы, колонки, разделов ----------
export function BreaksButton({ editor }) {
  const c = () => editor.chain().focus();
  return (
    <MenuButton icon={<BreaksIcon />} label="Разрывы" title="Разрывы страниц и разделов">
      {(close) => (
        <>
          <div className="pop-title">Разрывы страниц</div>
          <Item onClick={() => { close(); c().insertBreak('page').run(); }}>Страница (Ctrl+Enter)</Item>
          <Item onClick={() => { close(); c().insertBreak('column').run(); }}>Колонка (Ctrl+Shift+Enter)</Item>
          <div className="pop-title">Разрывы разделов</div>
          {SECTION_TYPES.map(([k, l]) => <Item key={k} onClick={() => { close(); c().insertSectionBreak(k).run(); }}>{l}</Item>)}
        </>
      )}
    </MenuButton>
  );
}

// ---------- Вкладка «Макет» ----------
export function LayoutPanel({ editor, Group, onMargins, onHf, onParagraph }) {
  const g = parseSectPr(sectionAt(editor.state).sectPr);
  const mirror = !!editor.state.doc.attrs.mirrorMargins;
  const same = (p, mir = false) => mir === mirror && Math.abs(p.top - g.top) < 15 && Math.abs(p.right - g.right) < 15 && Math.abs(p.bottom - g.bottom) < 15 && Math.abs(p.left - g.left) < 15;
  const paper = paperOf(g.w, g.h);
  // заготовка полей: только поля (переплёт сохраняется), «Зеркальные» включают зеркальные поля, остальные — выключают
  const preset = (m, mir) => applyPageSetup(editor, { patch: { top: m.top, right: m.right, bottom: m.bottom, left: m.left + g.gutter }, mirror: !!mir });
  const last = lastMargins();
  return (
    <>
      <Group title="Параметры страницы" launcher={{ title: 'Параметры страницы…', onClick: () => onMargins('margins') }}>
        <MenuButton icon={<MarginsIcon />} label="Поля" title="Поля страницы">
          {(close) => (
            <>
              {last && (
                <Item checked={same({ ...last, left: last.left }, !!last.mirror)} onClick={() => { close(); preset({ ...last, left: last.left - g.gutter }, last.mirror); }}>
                  Последнее пользовательское: {cm(last.top)} / {cm(last.right)} / {cm(last.bottom)} / {cm(last.left)} см
                </Item>
              )}
              {MARGIN_PRESETS.map(([key, label, m, mir]) => (
                <Item key={key} checked={same({ ...m, left: m.left + g.gutter }, !!mir)} onClick={() => { close(); preset(m, mir); }}>
                  {label}: {mir ? `внутри ${cm(m.left)} / снаружи ${cm(m.right)}` : `${cm(m.top)} / ${cm(m.right)} / ${cm(m.bottom)} / ${cm(m.left)}`} см
                </Item>
              ))}
              <div className="menu-sep" />
              <Item onClick={() => { close(); onMargins('margins'); }}>Настраиваемые поля…</Item>
            </>
          )}
        </MenuButton>
        <MenuButton icon={<OrientIcon />} label="Ориентация" title="Ориентация страницы">
          {(close) => (
            <>
              <Item checked={g.orient !== 'landscape'} onClick={() => { close(); if (g.orient === 'landscape') setPage(editor, { w: g.h, h: g.w, orient: 'portrait' }); }}>Книжная</Item>
              <Item checked={g.orient === 'landscape'} onClick={() => { close(); if (g.orient !== 'landscape') setPage(editor, { w: g.h, h: g.w, orient: 'landscape' }); }}>Альбомная</Item>
            </>
          )}
        </MenuButton>
        <MenuButton icon={<SizeIcon />} label="Размер" title="Размер бумаги">
          {(close) => (
            <>
              {PAPER.map(([name, w, h, code]) => (
                <Item key={name} checked={paper && paper[0] === name} onClick={() => {
                  close();
                  const land = g.orient === 'landscape';
                  setPage(editor, { w: land ? h : w, h: land ? w : h, code });
                }}>{name}: {cm(w)} × {cm(h)} см</Item>
              ))}
              <div className="menu-sep" />
              <Item onClick={() => { close(); onMargins('paper'); }}>Другие размеры бумаги…</Item>
            </>
          )}
        </MenuButton>
        <ColumnsButton editor={editor} />
        <BreaksButton editor={editor} />
        <LineNumbersButton editor={editor} sectPr={sectionAt(editor.state).sectPr} />
        <HyphenButton editor={editor} />
      </Group>
      <LayoutParaGroup editor={editor} Group={Group} onParagraph={onParagraph} />
      <HfGroup editor={editor} Group={Group} onHf={onHf} />
    </>
  );
}

// ---------- Номера строк ----------
const RESTARTS = [['continuous', 'Непрерывно'], ['newPage', 'Заново на каждой странице'], ['newSection', 'Заново в каждом разделе']];
const RESTART_DLG = { continuous: 'непрерывная', newPage: 'на каждой странице', newSection: 'в каждом разделе' };
function LineNumbersButton({ editor, sectPr }) {
  const [dlg, setDlg] = useState(false);
  const ln = parseLnNum(sectPr);
  const set = (restart) => editor.chain().focus().setLineNumbering(restart ? { countBy: 1, start: 1, distance: null, ...(ln || {}), restart } : null).run();
  const para = editor.state.selection.$from.parent;
  return (
    <>
      <MenuButton icon={<LnIcon />} label="Номера строк" title="Номера строк на полях рядом с текстом">
        {(close) => (
          <>
            <Item checked={!ln} onClick={() => { close(); set(null); }}>Нет</Item>
            {RESTARTS.map(([k, label]) => <Item key={k} checked={!!ln && ln.restart === k} onClick={() => { close(); set(k); }}>{label}</Item>)}
            <div className="menu-sep" />
            <Item checked={!!para.attrs.noLineNumbers} onClick={() => { close(); editor.chain().focus().toggleNoLineNumbers().run(); }}>Запретить в текущем абзаце</Item>
            <div className="menu-sep" />
            <Item onClick={() => { close(); setDlg(true); }}>Варианты нумерации строк…</Item>
          </>
        )}
      </MenuButton>
      {dlg && <LineNumbersDialog editor={editor} ln={ln} onClose={() => setDlg(false)} />}
    </>
  );
}

function LineNumbersDialog({ editor, ln, onClose }) {
  const [on, setOn] = useState(!!ln);
  const [start, setStart] = useState(String(ln ? ln.start : 1));
  const [dist, setDist] = useState(ln && ln.distance != null ? cm(ln.distance) : 'Авто');
  const [step, setStep] = useState(String(ln ? ln.countBy : 1));
  const [restart, setRestart] = useState(ln ? ln.restart : 'newPage');
  const [err, setErr] = useState('');
  const apply = () => {
    if (!on) { editor.chain().focus().setLineNumbering(null).run(); onClose(); return; }
    const s = parseInt(start, 10); const c = parseInt(step, 10);
    const d = /^\s*(авто)?\s*$/i.test(dist) ? null : parseFloat(String(dist).replace(',', '.'));
    if (!(s >= 1 && s <= 32767) || !(c >= 1 && c <= 100) || (d != null && !(d >= 0 && d <= 55))) { setErr('Начать с — 1…32767, шаг — 1…100, от текста — «Авто» или 0…55 см'); return; }
    editor.chain().focus().setLineNumbering({ start: s, countBy: c, distance: d == null ? null : Math.round(d * TWIPS_PER_CM), restart }).run();
    onClose();
  };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Нумерация строк" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter' && e.target.tagName !== 'SELECT') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Нумерация строк</div>
        <div className="modal-body">
          <label className="pd-check"><input type="checkbox" checked={on} onChange={(e) => setOn(e.target.checked)} /> Добавить нумерацию строк</label>
          <div className="pd-grid">
            <label className="pd-field"><span>Начать с</span><span className="pd-input"><input disabled={!on} value={start} onChange={(e) => { setErr(''); setStart(e.target.value); }} /></span></label>
            <label className="pd-field"><span>От текста</span><span className="pd-input"><input disabled={!on} value={dist} onChange={(e) => { setErr(''); setDist(e.target.value); }} /><em>см</em></span></label>
            <label className="pd-field"><span>Шаг</span><span className="pd-input"><input disabled={!on} value={step} onChange={(e) => { setErr(''); setStep(e.target.value); }} /></span></label>
          </div>
          <fieldset><legend>Нумерация</legend>
            <div className="pd-checks">
              {RESTARTS.map(([k, label]) => (
                <label key={k} className="pd-check"><input type="radio" disabled={!on} checked={restart === k} onChange={() => setRestart(k)} /> {RESTART_DLG[k]}</label>
              ))}
            </div>
          </fieldset>
          <div className="hf-hint">Применяется к текущему разделу. «Авто» — {cm(LN_AUTO_DISTANCE)} см.</div>
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

// ---------- Расстановка переносов ----------
function HyphenButton({ editor }) {
  const [dlg, setDlg] = useState(false);
  const auto = !!editor.state.doc.attrs.autoHyphenation;
  return (
    <>
      <MenuButton icon={<HyphIcon />} label="Расстановка переносов" title="Перенос слов по слогам в конце строк">
        {(close) => (
          <>
            <Item checked={!auto} onClick={() => { close(); editor.chain().focus().setHyphenation({ auto: false }).run(); }}>Нет</Item>
            <Item checked={auto} onClick={() => { close(); editor.chain().focus().setHyphenation({ auto: true }).run(); }}>Авто</Item>
            <div className="menu-sep" />
            <Item onClick={() => { close(); setDlg(true); }}>Параметры расстановки переносов…</Item>
          </>
        )}
      </MenuButton>
      {dlg && <HyphenDialog editor={editor} onClose={() => setDlg(false)} />}
    </>
  );
}

function HyphenDialog({ editor, onClose }) {
  const a = editor.state.doc.attrs;
  const [auto, setAuto] = useState(!!a.autoHyphenation);
  const [caps, setCaps] = useState(a.hyphenateCaps !== false);
  const apply = () => { editor.chain().focus().setHyphenation({ auto, caps }).run(); onClose(); };
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Расстановка переносов" onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') { e.preventDefault(); apply(); } }}>
        <div className="modal-title">Расстановка переносов</div>
        <div className="modal-body">
          <div className="pd-checks">
            <label className="pd-check"><input type="checkbox" checked={auto} onChange={(e) => setAuto(e.target.checked)} /> Автоматическая расстановка переносов</label>
            <label className="pd-check"><input type="checkbox" checked={caps} onChange={(e) => setCaps(e.target.checked)} /> Переносы в словах из ПРОПИСНЫХ БУКВ</label>
          </div>
          <div className="hf-hint">Переносятся русские слова по правилам слогоделения; места переноса могут отличаться от Word.
            Запретить переносы в отдельном абзаце — «Абзац → Положение на странице».</div>
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Диалог «Колонтитулы» ----------
const toText = (paras) => (paras || []).map((p) => p.text).join('\n');
const toParas = (text, align) => {
  const lines = String(text).split('\n');
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  return lines.length ? lines.map((t) => ({ text: t, align })) : null;
};

// Диалог «Колонтитулы»: обычные (все/нечётные страницы), особые для первой страницы, для чётных страниц
export function HeaderFooterDialog({ editor, onClose }) {
  const hf = { ...EMPTY_HF, ...(editor.state.doc.attrs.hf || {}) };
  const own = { ...EMPTY_HF.own, ...(hf.own || {}) };
  const sectTitlePg = parseSectPr(sectionList(editor.state.doc)[0]).titlePg;
  const initTitlePg = own.titlePg ? !!hf.titlePg : sectTitlePg;
  // поле формы: текст и выравнивание варианта/вида колонтитула
  const init = (paras) => ({ text: toText(paras), align: (paras && paras[0] && paras[0].align) || 'center' });
  const [f, setF] = useState(() => ({
    header: init(hf.header), footer: init(hf.footer),
    firstHeader: init(hf.first && hf.first.header), firstFooter: init(hf.first && hf.first.footer),
    evenHeader: init(hf.even && hf.even.header), evenFooter: init(hf.even && hf.even.footer),
  }));
  const [titlePg, setTitlePg] = useState(initTitlePg);
  const [evenOdd, setEvenOdd] = useState(!!hf.evenOdd);
  const [tab, setTab] = useState('main');
  const set = (k, patch) => setF((x) => ({ ...x, [k]: { ...x[k], ...patch } }));
  const changed = (k, paras) => f[k].text !== toText(paras) || (paras && paras[0] && f[k].align !== paras[0].align);
  const fromFile = (!own.header && hf.header) || (!own.footer && hf.footer) || (!own.first && hf.first) || (!own.even && hf.even);

  const apply = () => {
    const next = { ...hf, own: { ...own } };
    if (changed('header', hf.header)) { next.header = toParas(f.header.text, f.header.align); next.own.header = true; }
    if (changed('footer', hf.footer)) { next.footer = toParas(f.footer.text, f.footer.align); next.own.footer = true; }
    const fh = hf.first || {};
    if (changed('firstHeader', fh.header) || changed('firstFooter', fh.footer)) {
      next.first = { header: toParas(f.firstHeader.text, f.firstHeader.align), footer: toParas(f.firstFooter.text, f.firstFooter.align) };
      next.own.first = true;
    }
    const eh = hf.even || {};
    if (changed('evenHeader', eh.header) || changed('evenFooter', eh.footer)) {
      next.even = { header: toParas(f.evenHeader.text, f.evenHeader.align), footer: toParas(f.evenFooter.text, f.evenFooter.align) };
      next.own.even = true;
    }
    if (titlePg !== initTitlePg) { next.titlePg = titlePg; next.own.titlePg = true; }
    if (evenOdd !== !!hf.evenOdd) { next.evenOdd = evenOdd; next.own.evenOdd = true; }
    setDocAttr(editor, 'hf', next);
    onClose();
  };
  const block = (label, k, id) => (
    <fieldset>
      <legend>{label}</legend>
      <textarea id={id} className="hf-text" rows={2} value={f[k].text} onChange={(e) => set(k, { text: e.target.value })} placeholder="Текст колонтитула (каждая строка — абзац)" />
      <div className="hf-tools">
        <select value={f[k].align} onChange={(e) => set(k, { align: e.target.value })} aria-label={`${label}: выравнивание`}>
          <option value="left">По левому краю</option><option value="center">По центру</option><option value="right">По правому краю</option>
        </select>
        <button type="button" className="btn small" onClick={() => set(k, { text: `${f[k].text}{PAGE}` })}>+ номер страницы</button>
        <button type="button" className="btn small" onClick={() => set(k, { text: `${f[k].text}{NUMPAGES}` })}>+ всего страниц</button>
      </div>
    </fieldset>
  );
  const tabs = [['main', evenOdd ? 'Нечётные страницы' : 'Все страницы'], titlePg && ['first', 'Первая страница'], evenOdd && ['even', 'Чётные страницы']].filter(Boolean);
  const cur = tabs.some(([k]) => k === tab) ? tab : 'main';
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal pd" role="dialog" aria-label="Колонтитулы" onMouseDown={(e) => e.stopPropagation()} onKeyDown={(e) => { if (e.key === 'Escape') onClose(); }}>
        <div className="modal-title">Колонтитулы</div>
        <div className="modal-body">
          <div className="pd-checks hf-variants">
            <label className="pd-check"><input type="checkbox" checked={titlePg} onChange={(e) => setTitlePg(e.target.checked)} /> Особый колонтитул для первой страницы</label>
            <label className="pd-check"><input type="checkbox" checked={evenOdd} onChange={(e) => setEvenOdd(e.target.checked)} /> Разные колонтитулы для чётных и нечётных страниц</label>
          </div>
          {tabs.length > 1 && (
            <div className="pd-tabs" role="tablist">
              {tabs.map(([k, label]) => <button key={k} type="button" role="tab" className={cur === k ? 'active' : ''} onClick={() => setTab(k)}>{label}</button>)}
            </div>
          )}
          {cur === 'main' && <>{block('Верхний колонтитул', 'header', 'hf-head')}{block('Нижний колонтитул', 'footer', 'hf-foot')}</>}
          {cur === 'first' && <>{block('Верхний колонтитул первой страницы', 'firstHeader', 'hf-first-head')}{block('Нижний колонтитул первой страницы', 'firstFooter', 'hf-first-foot')}</>}
          {cur === 'even' && <>{block('Верхний колонтитул чётных страниц', 'evenHeader', 'hf-even-head')}{block('Нижний колонтитул чётных страниц', 'evenFooter', 'hf-even-foot')}</>}
          <div className="hf-hint">{'{PAGE}'} — номер страницы, {'{NUMPAGES}'} — всего страниц.</div>
          {fromFile && <div className="hf-hint warn">Колонтитул из файла при изменении заменится простым текстом: рисунки и оформление в нём не сохранятся. Неизменённый колонтитул сохраняется как был.</div>}
        </div>
        <div className="modal-actions">
          <button className="btn primary" onClick={apply}>ОК</button>
          <button className="btn" onClick={onClose}>Отмена</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Колонтитулы первой и последней страницы (остальные рисует слой полос разбиения) ----------
// chrome — из разбиения: какой вариант колонтитула на первом и последнем листе (первая страница раздела, чётная)
// footGeo — геометрия последнего листа (разделы с разной геометрией), иначе как geo
export function PageChrome({ geo, footGeo = geo, hf, chrome, total }) {
  const head = chrome ? chrome.firstHeader : hf.header;
  const foot = chrome ? chrome.lastFooter : hf.footer;
  const line = (p, i, n) => <div key={i} className={`hf-line hf-${p.align || 'left'}`}>{hfText(p.text, n, total) || ' '}</div>;
  return (
    <>
      <div className="pero-hf pero-header page-chrome-top" style={{ height: geo.top, paddingTop: geo.header, paddingLeft: geo.left, paddingRight: geo.right }}>
        {(head || []).map((p, i) => line(p, i, 1))}
      </div>
      <div className="pero-hf pero-footer page-chrome-bottom" style={{ height: footGeo.bottom, paddingBottom: footGeo.footer, paddingLeft: footGeo.left, paddingRight: footGeo.right }}>
        {(foot || []).map((p, i) => line(p, i, total))}
      </div>
    </>
  );
}

// ---------- CSS печати/PDF: размер листа, поля и колонтитулы в полях страницы (@page) ----------
const cssString = (s) => `"${String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
function marginBox(paras, page = 'counter(page)', pages = 'counter(pages)') {
  return paras.map((p) => String(p.text).split(/(\{PAGE\}|\{NUMPAGES\})/).filter(Boolean)
    .map((x) => (x === '{PAGE}' ? page : x === '{NUMPAGES}' ? pages : cssString(x))).join(' ')).join(' "\\A" ');
}
// firstSectPr — XML первого раздела (w:titlePg — «особый колонтитул первой страницы»)
export function printCss(sectPr, hf, firstSectPr = sectPr, mirror = false, lnPad = 0) {
  const g = parseSectPr(sectPr);
  const px = pagePx(g);
  const h = hf || EMPTY_HF;
  // поля полей страницы (@top-left…@bottom-right) для варианта колонтитулов; force — гасить пустые (перекрывают умолчание)
  const boxesOf = (pick, force) => {
    const out = [];
    for (const [kind, where] of [['header', 'top'], ['footer', 'bottom']]) {
      const paras = pick(kind) || [];
      for (const align of ['left', 'center', 'right']) {
        const ps = paras.filter((p) => (p.align || 'left') === align && p.text);
        if (!ps.length) { if (force) out.push(`@${where}-${align} { content: none; }`); continue; }
        const pad = kind === 'header' ? `padding-top: ${px.header}px; vertical-align: top;` : `padding-bottom: ${px.footer}px; vertical-align: bottom;`;
        out.push(`@${where}-${align} { content: ${marginBox(ps)}; white-space: pre; font: 10pt Calibri, Carlito, sans-serif; ${pad} }`);
      }
    }
    return out.join(' ');
  };
  const sectTitlePg = parseSectPr(firstSectPr).titlePg;
  // lnPad — колонка номеров строк внутри области печати (lnPrintPad)
  let css = `@page { size: ${px.width}px ${px.height}px; margin: ${px.top}px ${px.right}px ${px.bottom}px ${px.left - lnPad}px; ${boxesOf((k) => h[k], false)} }`;
  // зеркальные поля: у чётных (левых) страниц «внутри» — справа
  if (mirror) css += ` @page :left { margin-left: ${px.right - lnPad}px; margin-right: ${px.left}px; }`;
  if (lnPad) css += ` @media print { .page .page-content { padding-left: ${lnPad}px !important; } }`;
  // чётные страницы — левые (первая — правая), первая страница — :first (объявлена последней: перекрывает :left/:right)
  if (h.evenOdd) css += ` @page :left { ${boxesOf((k) => hfFor(h, k, { page: 2, firstOfSection: false, sectTitlePg }), true)} }`;
  const titlePg = h.own && h.own.titlePg ? !!h.titlePg : sectTitlePg;
  if (titlePg) css += ` @page :first { ${boxesOf((k) => hfFor(h, k, { page: 1, firstOfSection: true, sectTitlePg }), true)} }`;
  return `@media print { ${css} }`;
}
