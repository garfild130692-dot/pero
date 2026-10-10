import React, { useCallback, useEffect, useRef, useState } from 'react';
import { useEditor, EditorContent, useEditorState } from '@tiptap/react';
import { EditorState } from '@tiptap/pm/state';
import { extensions, loadDocument } from './editor/extensions.js';
import { importDocx } from './docx/import.js';
import { exportDocx } from './docx/export.js';
import { OPAQUE_SUMMARY } from './editor/opaque.js';
import Ribbon, { QuickAccess } from './Ribbon.jsx';
import Ruler from './Ruler.jsx';
import ParagraphDialog from './ParagraphDialog.jsx';
import { cssFontStack } from './editor/font.js';
import { NEW_DOC_PARA } from './editor/paragraph.js';
import { lineCss } from './editor/paraFormat.js';
import { insertImageFiles } from './insertImage.js';
import { EMPTY_HF, pagePx, parseSectPr, sectionFrame, sectionList } from './editor/page.js';
import { pageChrome, pageInfo } from './editor/pagination.js';
import { HeaderFooterDialog, MarginsDialog, PageChrome, printCss } from './PageControls.jsx';
import FindBar from './FindBar.jsx';
import StyleDialog from './StyleDialog.jsx';
import { styleCss } from './editor/styles.js';
import { sectionPrintCss } from './editor/sections.js';
import CommentsPanel from './CommentsPanel.jsx';
import { orderedComments, commentsKey } from './editor/comments.js';
import { trackInfo } from './editor/track.js';
import { useAutosave } from './autosave.js';

// «Одинарный» интервал Word = естественная высота строки шрифта (ascent + descent + lineGap).
// В CSS её даёт line-height: normal; меряем отношение к размеру шрифта и кешируем.
const lineRatioCache = {};
function lineRatio(stack) {
  if (!stack) return 1.15;
  if (lineRatioCache[stack]) return lineRatioCache[stack];
  const el = document.createElement('span');
  el.style.cssText = `position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;line-height:normal;font-family:${stack}`;
  el.textContent = 'АgЙ';
  document.body.appendChild(el);
  const r = Math.round((el.getBoundingClientRect().height / 100) * 1000) / 1000;
  el.remove();
  return (lineRatioCache[stack] = r > 0.8 && r < 2 ? r : 1.15);
}

// Масштаб страницы — transform, а не CSS zoom: раскладка текста всегда «как на 100%», поэтому переносы строк
// и разбиение на страницы не зависят от масштаба (как в Word). Обёртка занимает место масштабированной страницы.
function PageScaler({ scale, children }) {
  const inner = useRef(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  useEffect(() => {
    const el = inner.current;
    if (!el) return undefined;
    const ro = new ResizeObserver(() => setSize({ w: el.offsetWidth, h: el.offsetHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  return (
    <div className="page-scaler" style={{ width: size.w * scale, height: size.h * scale }}>
      <div ref={inner} className="page-scaled" style={{ transform: scale === 1 ? 'none' : `scale(${scale})` }}>{children}</div>
    </div>
  );
}

// Геометрия листа → размеры и поля .page (на экране листы разбивает плагин pagination)
function pageBox(g) {
  return {
    width: `${g.width}px`, minHeight: `${g.height}px`, padding: `${g.top}px ${g.right}px ${g.bottom}px ${g.left}px`,
    '--page-ml': `${g.left}px`, '--page-mr': `${g.right}px`,
  };
}

// Умолчания документа → CSS-переменные страницы
function pageVars(d) {
  if (!d) return {};
  const stack = cssFontStack(d.family);
  const p = { ...NEW_DOC_PARA, ...(d.para || {}) };
  return {
    '--doc-font': stack,
    '--doc-size': d.size ? `${d.size}pt` : undefined,
    '--line-ratio': lineRatio(stack),
    '--p-before': `${p.spaceBefore}pt`,
    '--p-after': `${p.spaceAfter}pt`,
    '--p-line': lineCss(p.lineRule, p.lineSpacing),
    '--p-left': `${p.indentLeft}pt`,
    '--p-right': `${p.indentRight}pt`,
    '--p-indent': `${p.indentFirst}pt`,
  };
}
import * as fileApi from './fileApi.js';

const EMPTY = { type: 'doc', content: [{ type: 'paragraph' }] };
const DEFAULT_NAME = 'Документ1.docx';

// Сброс истории отмен после загрузки файла: Ctrl+Z не должен «отменять» открытие документа
function resetHistory(editor) {
  const { state, view } = editor;
  view.updateState(EditorState.create({ doc: state.doc, plugins: state.plugins, selection: state.selection }));
}

function Modal({ title, children, actions, onClose }) {
  useEffect(() => {
    const h = (e) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [onClose]);
  return (
    <div className="modal-backdrop" onMouseDown={onClose}>
      <div className="modal" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-label={title}>
        <div className="modal-title">{title}</div>
        <div className="modal-body">{children}</div>
        <div className="modal-actions">{actions}</div>
      </div>
    </div>
  );
}

function LinkDialog({ editor, onClose }) {
  const prev = editor.getAttributes('link').href || '';
  const hasSelection = !editor.state.selection.empty;
  const [href, setHref] = useState(prev);
  const [text, setText] = useState('');
  const apply = () => {
    let url = href.trim();
    if (!url) { editor.chain().focus().extendMarkRange('link').unsetLink().run(); return onClose(); }
    if (!/^(https?:|mailto:|#)/i.test(url)) url = 'https://' + url;
    const ch = editor.chain().focus().extendMarkRange('link');
    if (!hasSelection && !prev) ch.insertContent({ type: 'text', text: text || url, marks: [{ type: 'link', attrs: { href: url } }] });
    else ch.setLink({ href: url });
    ch.run();
    onClose();
  };
  return (
    <Modal
      title="Гиперссылка"
      onClose={onClose}
      actions={<><button className="btn primary" onClick={apply}>ОК</button><button className="btn" onClick={onClose}>Отмена</button></>}
    >
      {!hasSelection && !prev && (
        <label className="field">Текст<input value={text} onChange={(e) => setText(e.target.value)} placeholder="Отображаемый текст" /></label>
      )}
      <label className="field">Адрес
        <input autoFocus value={href} onChange={(e) => setHref(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && apply()} placeholder="https://…" />
      </label>
    </Modal>
  );
}

function StatusBar({ editor, warnings, preserved, zoom, setZoom }) {
  const { words, chars, page, total, track } = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const t = e.state.doc.textBetween(0, e.state.doc.content.size, ' ', ' ');
      return { words: (t.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) || []).length, chars: t.replace(/\s/g, '').length, ...pageInfo(e.state), track: trackInfo(e.state) };
    },
  });
  const warnList = Object.entries(warnings);
  // что Перо показывает значком и сохраняет без изменений (без служебных меток вроде закладок)
  const keptList = [...new Set(Object.keys(preserved).map((k) => OPAQUE_SUMMARY[k]).filter(Boolean))];
  return (
    <div className="statusbar">
      <span>Страница {page} из {total}</span>
      <span>Слов: {words.toLocaleString('ru-RU')}</span>
      <span>Знаков: {chars.toLocaleString('ru-RU')}</span>
      <span>Русский</span>
      {track.enabled && <span className="track-on" title="Правки записываются как исправления">Исправления: вкл.</span>}
      {!track.enabled && track.count > 0 && <span title="В документе есть исправления — см. вкладку «Рецензирование»">Исправлений: {track.count}</span>}
      {warnList.length > 0 && (
        <span className="warn" title={'Пока не поддерживается и будет потеряно при сохранении:\n' + warnList.map(([k, n]) => `• ${k}: ${n}`).join('\n')}>
          ⚠ Будет потеряно при сохранении: {warnList.map(([k]) => k).join(', ')}
        </span>
      )}
      {keptList.length > 0 && (
        <span className="kept" title={'Перо пока не редактирует эти элементы, но сохраняет их без изменений:\n' + keptList.map((k) => `• ${k}`).join('\n')}>
          Сохраняются без изменений: {keptList.join(', ')}
        </span>
      )}
      <span className="spacer" />
      <button className="zoom-btn" onClick={() => setZoom((z) => Math.max(50, z - 10))}>−</button>
      <input type="range" min="50" max="200" step="10" value={zoom} onChange={(e) => setZoom(+e.target.value)} />
      <button className="zoom-btn" onClick={() => setZoom((z) => Math.min(200, z + 10))}>+</button>
      <span className="zoom-val">{zoom}%</span>
    </div>
  );
}

export default function App({ initialFile, onOpenOther, onNewSheet, autosaveSec }) {
  const [doc, setDoc] = useState({ path: null, name: DEFAULT_NAME });
  const [dirty, setDirtyState] = useState(false);
  const [warnings, setWarnings] = useState({});   // что теряется при сохранении
  const [preserved, setPreserved] = useState({}); // что сохраняется как есть
  // Исходный .docx: при сохранении всё, что Перо не меняет, переносится из него без изменений
  const source = useRef(null);
  const [zoom, setZoom] = useState(100);
  const [modal, setModal] = useState(null);
  const [findMode, setFindMode] = useState(null); // null | 'find' | 'replace'
  const [showComments, setShowComments] = useState(false);
  const [userName, setUserName] = useState('Пользователь');
  const revision = useRef(0); // счётчик правок — для автосохранения
  useEffect(() => { fileApi.userName().then((n) => n && setUserName(n)).catch(() => {}); }, []);

  const editor = useEditor({
    extensions,
    content: EMPTY,
    autofocus: true,
    editorProps: {
      attributes: { class: 'page-content', spellcheck: 'true', lang: 'ru' },
      // картинки из буфера обмена и перетаскиванием — как рисунки
      handlePaste: (view, event) => {
        const files = [...(event.clipboardData?.files || [])].filter((f) => f.type.startsWith('image/'));
        if (!files.length) return false;
        insertImageFiles(view.dom.editor, files);
        return true;
      },
      handleDrop: (view, event) => {
        const files = [...(event.dataTransfer?.files || [])].filter((f) => f.type.startsWith('image/'));
        if (!files.length) return false;
        event.preventDefault();
        const at = view.posAtCoords({ left: event.clientX, top: event.clientY });
        insertImageFiles(view.dom.editor, files, at ? at.pos : null);
        return true;
      },
    },
    onUpdate: ({ transaction }) => { if (transaction.docChanged) { revision.current++; setDirtyState(true); } },
  });

  useEffect(() => { if (editor) editor.commands.setTrackAuthor(userName); }, [editor, userName]);

  useEffect(() => {
    fileApi.setDirty(dirty);
    fileApi.setTitle(`${doc.name}${dirty ? ' •' : ''} — Перо`);
  }, [dirty, doc.name]);
  useEffect(() => { fileApi.setDoc({ path: doc.path, kind: 'doc' }); }, [doc.path]);

  useAutosave({
    intervalSec: autosaveSec, dirty, revision: () => revision.current, kind: 'doc', name: doc.name, path: doc.path,
    getBytes: () => exportDocx(editor.getJSON(), { source: source.current }),
  });

  const load = useCallback(async (file) => {
    try {
      const { doc: json, warnings: w, preserved: kept } = await importDocx(file.data);
      loadDocument(editor, json);
      resetHistory(editor);
      editor.commands.focus('start');
      setDoc({ path: file.path, name: file.name });
      setWarnings(w);
      setPreserved(kept);
      setShowComments((json.attrs.comments || []).length > 0);
      source.current = file.data;
      setDirtyState(!!file.recovered); // восстановленный после сбоя — ещё не сохранён
    } catch (e) {
      fileApi.showError('Не удалось открыть файл', `${file.name}\n\n${e.message || e}`);
    }
  }, [editor]);

  // файл, с которым открыли окно редактора (переход из книги)
  useEffect(() => { if (initialFile && editor) load(initialFile); }, [initialFile, editor, load]);

  const newDoc = useCallback(async () => {
    if (fileApi.newWindow('doc')) return; // в программе — новое окно, как Ctrl+N в Word
    if (dirty && !(await fileApi.confirmDiscard())) return;
    loadDocument(editor, EMPTY);
    resetHistory(editor);
    editor.commands.focus();
    setDoc({ path: null, name: DEFAULT_NAME });
    setWarnings({});
    setPreserved({});
    source.current = null;
    setDirtyState(false);
  }, [dirty, editor]);

  // Файл, который нужно открыть в этом окне (в программе — только если окно было пустым, иначе файл уже открыт в новом)
  const openHere = useCallback(async (f) => {
    if (!f) return;
    if (/\.(xlsx|csv)$/i.test(f.name)) { fileApi.setDirty(false); onOpenOther(f); } else await load(f);
  }, [load, onOpenOther]);

  const open = useCallback(async () => {
    if (!fileApi.isDesktop && dirty && !(await fileApi.confirmDiscard())) return;
    openHere(await fileApi.openFile());
  }, [dirty, openHere]);

  const newSheet = useCallback(async () => {
    if (fileApi.newWindow('sheet')) return;
    if (dirty && !(await fileApi.confirmDiscard())) return;
    fileApi.setDirty(false);
    onNewSheet();
  }, [dirty, onNewSheet]);

  const write = useCallback(async (asNew) => {
    try {
      const bytes = await exportDocx(editor.getJSON(), { source: source.current });
      const r = asNew || !doc.path ? await fileApi.saveFileAs(doc.name, bytes) : await fileApi.saveFile(doc.path, doc.name, bytes);
      if (!r) return false;
      setDoc({ path: r.path, name: r.name });
      setDirtyState(false);
      if (r.path !== doc.path) setWarnings({});
      return true;
    } catch (e) {
      fileApi.showError('Не удалось сохранить файл', String(e.message || e));
      return false;
    }
  }, [editor, doc]);

  // Сохранение поверх файла, в котором были неподдерживаемые элементы, — только с подтверждением
  const save = useCallback((then) => new Promise((resolve) => {
    const done = (ok) => { resolve(ok); if (ok && typeof then === 'function') then(); };
    if (doc.path && Object.keys(warnings).length) {
      setModal({
        title: 'Сохранение может удалить часть содержимого',
        body: `В исходном файле есть элементы, которые Перо пока не умеет сохранять (${Object.keys(warnings).join(', ')}). ` +
          'Если перезаписать файл, они пропадут. Рекомендуется сохранить копию.',
        actions: [
          ['Сохранить копию', true, async () => { setModal(null); done(await write(true)); }],
          ['Перезаписать', false, async () => { setModal(null); done(await write(false)); }],
          ['Отмена', false, () => { setModal(null); done(false); }],
        ],
      });
    } else write(false).then(done);
  }), [doc.path, warnings, write]);

  const saveAs = useCallback(() => write(true), [write]);

  const doPrint = useCallback(() => fileApi.print(), []);
  const doPdf = useCallback(async () => {
    try {
      await fileApi.exportPdf(doc.name.replace(/\.docx$/i, '') + '.pdf');
    } catch (e) {
      fileApi.showError('Не удалось сохранить PDF', String(e.message || e));
    }
  }, [doc.name]);

  // Команды меню Electron. Держим актуальные обработчики в ref, подписку — одну.
  const handlers = useRef({});
  handlers.current = { newDoc, newSheet, open, save, saveAs, load, doPrint, doPdf, onOpenOther };
  useEffect(() => fileApi.onMenu((cmd, payload) => {
    const h = handlers.current;
    if (cmd === 'new') h.newDoc();
    else if (cmd === 'open') h.open();
    else if (cmd === 'save') h.save();
    else if (cmd === 'saveAs') h.saveAs();
    else if (cmd === 'newSheet') h.newSheet();
    else if (cmd === 'openData') { if (/\.(xlsx|csv)$/i.test(payload.name)) h.onOpenOther(payload); else h.load(payload); }
    else if (cmd === 'saveAndClose') h.save(() => fileApi.closeNow());
    else if (cmd === 'print') h.doPrint();
    else if (cmd === 'find' || cmd === 'replace') setFindMode(cmd);
    else if (cmd === 'pdf') h.doPdf();
    else if (cmd === 'zoom') setZoom((z) => (payload === 0 ? 100 : Math.min(200, Math.max(50, z + payload * 10))));
  }), []);

  // Горячие клавиши, которые в Electron обслуживает меню, а в браузере (режим отладки) — мы сами; плюс Ctrl+K
  useEffect(() => {
    const onKey = (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = e.code;
      if (k === 'KeyK') { e.preventDefault(); setModal('link'); return; }
      if (k === 'KeyF') { e.preventDefault(); setFindMode('find'); return; }
      if (k === 'KeyH') { e.preventDefault(); setFindMode('replace'); return; }
      if (fileApi.isDesktop) return;
      const h = handlers.current;
      if (k === 'KeyS') { e.preventDefault(); e.shiftKey ? h.saveAs() : h.save(); }
      else if (k === 'KeyO') { e.preventDefault(); h.open(); }
      else if (k === 'KeyP') { e.preventDefault(); h.doPrint(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  // Перетаскивание .docx в окно
  const onDrop = async (e) => {
    const f = [...e.dataTransfer.files].find((x) => /\.(docx|xlsx|csv)$/i.test(x.name));
    if (!f) return;
    e.preventDefault();
    const p = fileApi.pathForFile(f);
    if (p) { openHere(await fileApi.openPath(p)); return; } // программа сама решит: это окно или новое
    if (dirty && !(await fileApi.confirmDiscard())) return;
    openHere({ path: null, name: f.name, data: await f.arrayBuffer() });
  };

  // Шрифт и размер документа по умолчанию (атрибуты документа) → CSS-переменные страницы
  const docFont = useEditorState({
    editor,
    selector: ({ editor: e }) => (e ? {
      family: e.state.doc.attrs.fontFamily, size: e.state.doc.attrs.fontSize, para: e.state.doc.attrs.para,
      sectPr: e.state.doc.attrs.sectPr, hf: e.state.doc.attrs.hf, total: pageInfo(e.state).total, styles: e.state.doc.attrs.styles, chrome: pageChrome(e.state), firstSect: sectionList(e.state.doc)[0], sects: sectionList(e.state.doc), sectCss: sectionPrintCss(e.state.doc),
    } : null),
  });

  if (!editor) return null;
  const geo = pagePx(parseSectPr(docFont && docFont.sectPr));
  // разделы с разной геометрией: лист шириной с самый широкий раздел, поля колонтитулов — по первому/последнему разделу
  const frame = sectionFrame((docFont && docFont.sects) || [docFont && docFont.sectPr]);
  const firstSec = frame.sections[0];
  const lastSec = frame.sections[frame.sections.length - 1];
  const pageGeo = frame.mixed
    ? { ...geo, width: frame.width, height: firstSec.height, top: firstSec.top, bottom: lastSec.bottom, left: frame.padLeft, right: frame.padRight }
    : geo;
  const chromeGeo = frame.mixed
    ? { head: { ...firstSec, left: firstSec.offset + firstSec.left, right: frame.width - firstSec.offset - firstSec.width + firstSec.right },
      foot: { ...lastSec, left: lastSec.offset + lastSec.left, right: frame.width - lastSec.offset - lastSec.width + lastSec.right } }
    : { head: geo, foot: geo };

  return (
    <div className="app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <div className="titlebar">
        <QuickAccess editor={editor} onSave={() => save()} />
        <div className="doc-title">{doc.name}{dirty ? ' •' : ''}</div>
      </div>
      <Ribbon editor={editor} file={{ newDoc, newSheet, open, openHere, save: () => save(), saveAs, print: doPrint, pdf: doPdf }} onLink={() => setModal('link')}
        onFind={(m) => setFindMode(m)}
        onStyle={(m) => (m && m.remove ? editor.chain().focus().deleteStyle(m.remove).run() : setModal({ style: m }))}
        review={{
          showComments,
          toggleComments: () => setShowComments((v) => !v),
          addComment: () => {
            const initials = userName.split(/\s+/).map((x) => x[0] || '').join('').slice(0, 3).toUpperCase();
            editor.chain().focus().addComment({ author: userName, initials }).run();
            setShowComments(true);
          },
          stepComment: (dir) => {
            const list = orderedComments(editor.state);
            if (!list.length) return;
            const cur = list.findIndex((c) => c.id === (commentsKey.getState(editor.state) || {}).active);
            const next = list[(cur + dir + list.length) % list.length] || list[0];
            editor.commands.goToComment(next.id);
            setShowComments(true);
          },
        }}
        onParagraph={() => setModal('para')} onMargins={() => setModal('margins')} onHf={() => setModal('hf')} />
      <div className="main-row">
      <div
        className="workspace"
        onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}
        onWheel={(e) => { if (e.ctrlKey) setZoom((z) => Math.min(200, Math.max(50, z + (e.deltaY < 0 ? 10 : -10)))); }} // Ctrl+колесо — масштаб
      >
        {findMode && <FindBar editor={editor} mode={findMode} setMode={setFindMode} />}
        <Ruler editor={editor} zoom={zoom} />
        <PageScaler scale={zoom / 100}>
          <div className={`page${frame.mixed ? ' pero-multi' : ''}`} style={{ ...pageVars(docFont), ...pageBox(pageGeo) }}>
            <PageChrome geo={chromeGeo.head} footGeo={chromeGeo.foot} hf={(docFont && docFont.hf) || EMPTY_HF} chrome={docFont && docFont.chrome} total={(docFont && docFont.total) || 1} />
            <EditorContent editor={editor} />
          </div>
        </PageScaler>
      </div>
      {showComments && <CommentsPanel editor={editor} onClose={() => setShowComments(false)} />}
      </div>
      <StatusBar editor={editor} warnings={warnings} preserved={preserved} zoom={zoom} setZoom={setZoom} />

      {modal === 'link' && <LinkDialog editor={editor} onClose={() => setModal(null)} />}
      {modal === 'para' && <ParagraphDialog editor={editor} onClose={() => setModal(null)} />}
      {modal === 'margins' && <MarginsDialog editor={editor} onClose={() => setModal(null)} />}
      {modal === 'hf' && <HeaderFooterDialog editor={editor} onClose={() => setModal(null)} />}
      {modal && modal.style && <StyleDialog editor={editor} mode={modal.style} onClose={() => setModal(null)} />}
      {/* пользовательские стили абзацев — правила по data-pstyle */}
      <style>{styleCss(docFont && docFont.styles)}</style>
      {/* печать/PDF: размер листа, поля и колонтитулы — правилами @page */}
      <style>{printCss(docFont && docFont.sectPr, (docFont && docFont.hf) || EMPTY_HF, docFont && docFont.firstSect)}{docFont && docFont.sectCss ? `@media print { ${docFont.sectCss} }` : ''}</style>
      {modal && typeof modal === 'object' && !modal.style && (
        <Modal
          title={modal.title}
          onClose={() => modal.actions[modal.actions.length - 1][2]()}
          actions={modal.actions.map(([label, primary, fn]) => (
            <button key={label} className={`btn${primary ? ' primary' : ''}`} onClick={fn}>{label}</button>
          ))}
        >
          {modal.body}
        </Modal>
      )}
    </div>
  );
}
