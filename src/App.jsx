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
  const { words, chars } = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const t = e.state.doc.textBetween(0, e.state.doc.content.size, ' ', ' ');
      return { words: (t.match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) || []).length, chars: t.replace(/\s/g, '').length };
    },
  });
  const warnList = Object.entries(warnings);
  // что Перо показывает значком и сохраняет без изменений (без служебных меток вроде закладок)
  const keptList = [...new Set(Object.keys(preserved).map((k) => OPAQUE_SUMMARY[k]).filter(Boolean))];
  return (
    <div className="statusbar">
      <span>Слов: {words.toLocaleString('ru-RU')}</span>
      <span>Знаков: {chars.toLocaleString('ru-RU')}</span>
      <span>Русский</span>
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

export default function App() {
  const [doc, setDoc] = useState({ path: null, name: DEFAULT_NAME });
  const [dirty, setDirtyState] = useState(false);
  const [warnings, setWarnings] = useState({});   // что теряется при сохранении
  const [preserved, setPreserved] = useState({}); // что сохраняется как есть
  // Исходный .docx: при сохранении всё, что Перо не меняет, переносится из него без изменений
  const source = useRef(null);
  const [zoom, setZoom] = useState(100);
  const [modal, setModal] = useState(null);

  const editor = useEditor({
    extensions,
    content: EMPTY,
    autofocus: true,
    editorProps: { attributes: { class: 'page-content', spellcheck: 'true', lang: 'ru' } },
    onUpdate: ({ transaction }) => { if (transaction.docChanged) setDirtyState(true); },
  });

  useEffect(() => {
    fileApi.setDirty(dirty);
    fileApi.setTitle(`${doc.name}${dirty ? ' •' : ''} — Перо`);
  }, [dirty, doc.name]);

  const load = useCallback(async (file) => {
    try {
      const { doc: json, warnings: w, preserved: kept } = await importDocx(file.data);
      loadDocument(editor, json);
      resetHistory(editor);
      editor.commands.focus('start');
      setDoc({ path: file.path, name: file.name });
      setWarnings(w);
      setPreserved(kept);
      source.current = file.data;
      setDirtyState(false);
    } catch (e) {
      fileApi.showError('Не удалось открыть файл', `${file.name}\n\n${e.message || e}`);
    }
  }, [editor]);

  const newDoc = useCallback(async () => {
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

  const open = useCallback(async () => {
    if (dirty && !(await fileApi.confirmDiscard())) return;
    const f = await fileApi.openFile();
    if (f) await load(f);
  }, [dirty, load]);

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

  // Команды меню Electron. Держим актуальные обработчики в ref, подписку — одну.
  const handlers = useRef({});
  handlers.current = { newDoc, open, save, saveAs, load };
  useEffect(() => fileApi.onMenu((cmd, payload) => {
    const h = handlers.current;
    if (cmd === 'new') h.newDoc();
    else if (cmd === 'open') h.open();
    else if (cmd === 'save') h.save();
    else if (cmd === 'saveAs') h.saveAs();
    else if (cmd === 'openData') h.load(payload);
    else if (cmd === 'saveAndClose') h.save(() => fileApi.closeNow());
    else if (cmd === 'zoom') setZoom((z) => (payload === 0 ? 100 : Math.min(200, Math.max(50, z + payload * 10))));
  }), []);

  // Горячие клавиши, которые в Electron обслуживает меню, а в браузере (режим отладки) — мы сами; плюс Ctrl+K
  useEffect(() => {
    const onKey = (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = e.code;
      if (k === 'KeyK') { e.preventDefault(); setModal('link'); return; }
      if (fileApi.isDesktop) return;
      const h = handlers.current;
      if (k === 'KeyS') { e.preventDefault(); e.shiftKey ? h.saveAs() : h.save(); }
      else if (k === 'KeyO') { e.preventDefault(); h.open(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  // Перетаскивание .docx в окно
  const onDrop = async (e) => {
    const f = [...e.dataTransfer.files].find((x) => /\.docx$/i.test(x.name));
    if (!f) return;
    e.preventDefault();
    if (dirty && !(await fileApi.confirmDiscard())) return;
    await load({ path: fileApi.pathForFile(f), name: f.name, data: await f.arrayBuffer() });
  };

  // Шрифт и размер документа по умолчанию (атрибуты документа) → CSS-переменные страницы
  const docFont = useEditorState({
    editor,
    selector: ({ editor: e }) => (e ? { family: e.state.doc.attrs.fontFamily, size: e.state.doc.attrs.fontSize, para: e.state.doc.attrs.para } : null),
  });

  if (!editor) return null;

  return (
    <div className="app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <div className="titlebar">
        <QuickAccess editor={editor} onSave={() => save()} />
        <div className="doc-title">{doc.name}{dirty ? ' •' : ''}</div>
      </div>
      <Ribbon editor={editor} file={{ newDoc, open, save: () => save(), saveAs }} onLink={() => setModal('link')} onParagraph={() => setModal('para')} />
      <div
        className="workspace"
        onMouseDown={(e) => { if (e.target === e.currentTarget) editor.commands.focus('end'); }}
        onWheel={(e) => { if (e.ctrlKey) setZoom((z) => Math.min(200, Math.max(50, z + (e.deltaY < 0 ? 10 : -10)))); }} // Ctrl+колесо — масштаб
      >
        <Ruler editor={editor} zoom={zoom} />
        <div
          className="page"
          style={{ zoom: zoom / 100, ...pageVars(docFont) }}
        >
          <EditorContent editor={editor} />
        </div>
      </div>
      <StatusBar editor={editor} warnings={warnings} preserved={preserved} zoom={zoom} setZoom={setZoom} />

      {modal === 'link' && <LinkDialog editor={editor} onClose={() => setModal(null)} />}
      {modal === 'para' && <ParagraphDialog editor={editor} onClose={() => setModal(null)} />}
      {modal && typeof modal === 'object' && (
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
