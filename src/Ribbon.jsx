// Лента в стиле классического офисного редактора: вкладки → группы → кнопки.
import React, { useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { Icon } from './icons.jsx';
import { about, clipboard, newWindow, openImage, openLogs, reportProblem } from './fileApi.js';
import { QuickButton } from './TitleBar.jsx';
import RecentButton from './RecentButton.jsx';
import { insertImageFiles } from './insertImage.js';
import { PicturePanel } from './PictureControls.jsx';
import { HfGroup, LayoutPanel } from './PageControls.jsx';
import { TABLE_STYLES } from './editor/tableFormat.js';
import { trackInfo } from './editor/track.js';
import { sectionAt } from './editor/page.js';
import { FontFamilyBox, FontSizeBox, HighlightButton, TextColorButton, fontState, usePopover } from './FontControls.jsx';
import { effectivePara } from './editor/paraFormat.js';
import { LINE_PRESETS, fmtNum } from './editor/paragraph.js';
import { ParaBordersButton, ParaShadingButton } from './ParaLayoutControls.jsx';
import { UNDERLINES } from './editor/extensions.js';
import { CASES, painterKey } from './editor/runFormat.js';
import { marksKey } from './editor/marks.js';
import { SymbolButton } from './SymbolDialog.jsx';
import { QuickPartsButton } from './FieldDialogs.jsx';
import { MultilevelButton } from './NumberingDialogs.jsx';
import { BulletButton, NumberButton } from './ListButtons.jsx';
import { sampleOf, styleEntries } from './StylesPane.jsx';
import { ViewPanel } from './ViewControls.jsx';
import { DesignPanel } from './DesignControls.jsx';
import { PasteButton } from './PasteControls.jsx';
import { BordersButton, RowHeightField, ShadingButton, TableIcons, tableState } from './TableControls.jsx';

function Btn({ icon, label, title, onClick, onDoubleClick, active, disabled, big, children }) {
  const I = icon && Icon[icon];
  return (
    <button
      type="button"
      className={`rb-btn${big ? ' big' : ''}${active ? ' active' : ''}`}
      title={title || label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} // не забираем фокус у текста
      onClick={onClick}
      onDoubleClick={onDoubleClick}
    >
      {I && <I size={big ? 28 : 18} />}
      {children}
      {big && label && <span className="rb-label">{label}</span>}
    </button>
  );
}

// launcher — кнопка в углу заголовка группы, открывает полный диалог (как в Word)
const Group = ({ title, children, launcher }) => (
  <div className="rb-group">
    <div className="rb-group-body">{children}</div>
    <div className="rb-group-title">
      {title}
      {launcher && (
        <button type="button" className="rb-launcher" title={launcher.title} onMouseDown={(e) => e.preventDefault()} onClick={launcher.onClick}>
          <Icon.launcher size={11} />
        </button>
      )}
    </div>
  </div>
);

// Состояние абзаца под курсором для ленты
function paraState(e) {
  const { $from } = e.state.selection;
  let inList = false;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.name === 'listItem') inList = true;
  const node = $from.parent;
  if (!['paragraph', 'heading'].includes(node.type.name)) return { inList, lineRule: null, lineSpacing: null, before: 0, after: 0, left: 0 };
  const p = effectivePara(node, e.state.doc.attrs, inList);
  return { inList, lineRule: p.lineRule, lineSpacing: p.lineSpacing, before: p.spaceBefore, after: p.spaceAfter, left: p.indentLeft };
}

// Кнопка «Интервал» с меню: множители, «Другие варианты…», интервалы перед/после
// Подчёркнутый: кнопка + стрелка со списком видов (как в Word), «Другие подчёркивания…» — диалог «Шрифт»
function UnderlineButton({ editor, kind }) {
  const pop = usePopover();
  const c = () => editor.chain().focus();
  return (
    <div className="split rb-popover-anchor" ref={pop.ref}>
      <Btn title="Подчёркнутый (Ctrl+U)" active={!!kind} onClick={() => c().toggleUnderline().run()}><u className="glyph">Ч</u></Btn>
      <button type="button" className="rb-btn split-arrow" title="Вид подчёркивания" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {UNDERLINES.filter(([k]) => k !== 'words').map(([k, label]) => (
            <button key={k} type="button" className="menu-item" title={label} onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); c().setUnderlineKind(k).run(); }}>
              <span className="check">{kind === k ? '✓' : ''}</span>
              <span className="u-sample" data-u={k}>{label}</span>
            </button>
          ))}
          <div className="menu-sep" />
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); c().setUnderlineKind('words').run(); }}>
            <span className="check">{kind === 'words' ? '✓' : ''}</span>Только слова (Ctrl+Shift+W)
          </button>
          <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); c().unsetMark('underline').run(); }}>
            <span className="check" />Без подчёркивания
          </button>
        </div>
      )}
    </div>
  );
}

// Оглавление: после изменения номера страниц сдвигаются — второе обновление по новому разбиению (как в Word)
function refreshToc(editor) {
  editor.chain().focus().updateToc().run();
  setTimeout(() => editor.commands.updateToc(), 400);
}

// «Оглавление ▾»: автособираемое (уровни 1–3 или 1–2), обновить, удалить
function TocButton({ editor, Btn }) {
  const pop = usePopover();
  const item = (label, fn) => (
    <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}><span className="check" />{label}</button>
  );
  const insert = (levels) => {
    editor.chain().focus().insertToc(`TOC \\o "${levels}" \\h \\z \\u`).run();
    setTimeout(() => editor.commands.updateToc(), 400);
  };
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn big icon="toc" label="Оглавление" title="Оглавление по заголовкам" active={pop.open} onClick={() => pop.setOpen(!pop.open)} />
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {item('Автособираемое оглавление (уровни 1–3)', () => insert('1-3'))}
          {item('Автособираемое оглавление (уровни 1–2)', () => insert('1-2'))}
          <div className="menu-sep" />
          {item('Обновить оглавление', () => refreshToc(editor))}
          {item('Удалить оглавление', () => editor.chain().focus().removeToc().run())}
        </div>
      )}
    </div>
  );
}

// «Регистр» (Аа▾): как в предложениях, все строчные, ВСЕ ПРОПИСНЫЕ, Начинать С Прописных, иЗМЕНИТЬ РЕГИСТР; Shift+F3 — по кругу
function CaseButton({ editor }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn title="Регистр (Shift+F3 — по кругу)" active={pop.open} onClick={() => pop.setOpen(!pop.open)}><span className="glyph">Аа<small>▾</small></span></Btn>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {CASES.map(([k, label]) => (
            <button key={k} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); editor.chain().focus().changeCase(k).run(); }}>
              <span className="check" />{label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function LineSpacingButton({ editor, para, onMore }) {
  const pop = usePopover();
  const c = () => editor.chain().focus();
  const item = (label, onClick, checked) => (
    <button type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); onClick(); }}>
      <span className="check">{checked ? '✓' : ''}</span>{label}
    </button>
  );
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn icon="lineSpacing" title="Интервал между строками и абзацами" active={pop.open} onClick={() => pop.setOpen(!pop.open)} />
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {LINE_PRESETS.map((m) => item(fmtNum(m, 2).includes(',') ? fmtNum(m, 2) : `${m},0`, () => c().setLineSpacing(m).run(),
            para.lineRule === 'auto' && Math.abs(para.lineSpacing - m) < 1e-6))}
          <div className="menu-sep" />
          {item('Другие варианты междустрочных интервалов…', onMore)}
          <div className="menu-sep" />
          {item(para.before > 0 ? 'Удалить интервал перед абзацем' : 'Добавить интервал перед абзацем', () => c().toggleSpaceBefore().run())}
          {item(para.after > 0 ? 'Удалить интервал после абзаца' : 'Добавить интервал после абзаца', () => c().toggleSpaceAfter().run())}
        </div>
      )}
    </div>
  );
}


function TablePicker({ onPick, style, items = [] }) {
  const [hover, setHover] = useState([0, 0]);
  const N = 8;
  return (
    <div className="table-picker" style={style} onMouseLeave={() => setHover([0, 0])}>
      <div className="tp-title">{hover[0] ? `Таблица ${hover[1]}×${hover[0]}` : 'Вставка таблицы'}</div>
      <div className="tp-grid">
        {Array.from({ length: N * N }, (_, i) => {
          const r = Math.floor(i / N) + 1;
          const c = (i % N) + 1;
          const on = r <= hover[0] && c <= hover[1];
          return (
            <div
              key={i}
              className={`tp-cell${on ? ' on' : ''}`}
              onMouseEnter={() => setHover([r, c])}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => onPick(r, c)}
            />
          );
        })}
      </div>
      {items.length > 0 && <div className="menu-sep" />}
      {items.map(([label, fn, disabled]) => (
        <button key={label} type="button" className="menu-item" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={fn}><span className="check" />{label}</button>
      ))}
    </div>
  );
}

// Кнопка с меню (строчная, как «Найти ▾» в Word): щелчок по стрелке — меню
function MenuSplit({ Btn, icon, label, title, onClick, items }) {
  const pop = usePopover();
  return (
    <div className="split rb-popover-anchor" ref={pop.ref}>
      <Btn icon={icon} title={title} onClick={onClick || (() => pop.setOpen(!pop.open))}><span className="rb-inline">{label}</span></Btn>
      <button type="button" className="rb-btn split-arrow" title={title} onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>▾</button>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {items.map(([l, fn]) => (
            <button key={l} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}><span className="check" />{l}</button>
          ))}
        </div>
      )}
    </div>
  );
}

// Формат текста под курсором — для «Выделить → Текст с похожим форматированием» (поиск по формату)
function similarFormat(editor) {
  const { state } = editor;
  const marks = state.storedMarks || state.selection.$from.marks();
  const f = {};
  for (const k of ['bold', 'italic', 'underline', 'strike', 'highlight', 'superscript', 'subscript']) f[k] = marks.some((m) => m.type.name === k);
  const par = state.selection.$from.parent;
  f.styleId = par.type.name === 'heading' ? `h${par.attrs.level}` : par.attrs.styleId || null;
  return f;
}

// Меню «Стили ▾»: создать, изменить, удалить
function StylesMenu({ s, onStyle }) {
  const pop = usePopover();
  const item = (label, fn, disabled) => (
    <button type="button" className="menu-item" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); fn(); }}>
      <span className="check" />{label}
    </button>
  );
  const current = s.styleId ? (s.styles.find((x) => x.id === s.styleId) || {}).name : s.heading ? null : 'Обычный';
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn title="Стили: создать, изменить, удалить" onClick={() => pop.setOpen(!pop.open)}>▾</Btn>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {item('Создать стиль из текущего абзаца…', () => onStyle('new'))}
          {item(current ? `Изменить стиль «${current}»…` : 'Изменить стиль…', () => onStyle(s.styleId || 'normal'), !current)}
          {s.styleId && item(`Удалить стиль «${current}»`, () => onStyle({ remove: s.styleId }))}
        </div>
      )}
    </div>
  );
}

// Стили таблиц: готовые оформления (границы, шапка, чередование строк)
function TableStylesButton({ editor }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <button type="button" className={`rb-btn big${pop.open ? ' active' : ''}`} title="Стили таблиц" onMouseDown={(e) => e.preventDefault()} onClick={() => pop.setOpen(!pop.open)}>
        <Icon.table size={28} /><span className="rb-label">Стили ▾</span>
      </button>
      {pop.open && (
        <div className="menu-pop table-styles" style={pop.style} role="menu">
          {TABLE_STYLES.map(([key, label, preview]) => (
            <button key={key} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()}
              onClick={() => { pop.setOpen(false); editor.chain().focus().applyTableStyle(key).run(); }}>
              <span className={`ts-preview ts-${key}`}>{preview}</span>{label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

export default function Ribbon({ editor, file, docFile, onInsertFile, view, onStylesPane, onLink, onFont, onSymbol, onRefDialog, onParagraph, onMargins, onHf, onFind, onStyle, review }) {
  const [tab, setTab] = useState('home');
  const tablePicker = usePopover();

  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      bold: e.isActive('bold'),
      italic: e.isActive('italic'),
      underline: e.isActive('underline'),
      underlineKind: e.isActive('underline') ? e.getAttributes('underline').kind || 'single' : null,
      painter: painterKey.getState(e.state).mode,
      showMarks: marksKey.getState(e.state).on,
      numbering: {
        headingNumbering: e.state.doc.attrs.headingNumbering,
        inHeading: e.state.selection.$from.parent.type.name === 'heading',
        headingNoNumber: !!e.state.selection.$from.parent.attrs.noNumber,
        inOrdered: e.isActive('orderedList'),
      },
      strike: e.isActive('strike'),
      bullet: e.isActive('bulletList'),
      ordered: e.isActive('orderedList'),
      link: e.isActive('link'),
      heading: [1, 2, 3, 4, 5, 6].find((l) => e.isActive('heading', { level: l })) || 0,
      align: ['center', 'right', 'justify'].find((a) => e.isActive({ textAlign: a })) || 'left',
      inTable: e.isActive('table'),
      canUndo: e.can().undo(),
      canRedo: e.can().redo(),
      canSink: e.can().sinkListItem('listItem'),
      canLift: e.can().liftListItem('listItem'),
      canMerge: e.can().mergeCells(),
      canSplit: e.can().splitCell(),
      font: fontState(e),
      para: paraState(e),
      table: e.isActive('table') ? tableState(e) : null,
      image: e.state.selection.node && e.state.selection.node.type.name === 'image' ? e.state.selection.node.attrs : null,
      page: { sectPr: sectionAt(e.state).sectPr, hf: e.state.doc.attrs.hf }, // для вкладки «Макет» — раздел под курсором
      styles: e.state.doc.attrs.styles || [],
      track: trackInfo(e.state),
      styleId: e.state.selection.$from.parent.attrs.styleId || null,
      charStyles: e.state.doc.attrs.charStyles || [],
      charStyle: ((e.state.storedMarks || e.state.selection.$from.marks()).find((m) => m.type.name === 'charStyle') || { attrs: {} }).attrs.id || null,
    }),
    // в состоянии есть вложенный объект — сравниваем по значению, чтобы лента не перерисовывалась зря
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  });

  const c = () => editor.chain().focus();
  const activeTab = (tab === 'table' && !s.inTable) || (tab === 'picture' && !s.image) ? 'home' : tab;

  const tabs = [
    ['file', 'Файл'],
    ['home', 'Главная'],
    ['insert', 'Вставка'],
    ['design', 'Конструктор'],
    ['layout', 'Макет'],
    ['refs', 'Ссылки'],
    ['review', 'Рецензирование'],
    ['view', 'Вид'],
  ];

  return (
    <div className="ribbon">
      <div className="rb-tabs">
        {tabs.map(([k, label]) => (
          <button key={k} className={`rb-tab${k === 'file' ? ' file' : ''}${activeTab === k ? ' active' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setTab(k)}>
            {label}
          </button>
        ))}
        {s.image && (
          <button className={`rb-tab contextual${activeTab === 'picture' ? ' active' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setTab('picture')}>
            Рисунок
          </button>
        )}
        {s.inTable && (
          <button className={`rb-tab contextual${activeTab === 'table' ? ' active' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setTab('table')}>
            Таблица
          </button>
        )}
      </div>

      <div className="rb-panel">
        {activeTab === 'file' && (
          <>
            <Group title="Документ">
              <Btn big icon="file" label="Создать" onClick={file.newDoc} title="Создать (Ctrl+N)" />
              <Btn big icon="table" label="Книга" onClick={file.newSheet} title="Создать электронную таблицу (.xlsx)" />
              <Btn big icon="open" label="Открыть" onClick={file.open} title="Открыть (Ctrl+O)" />
              <Btn big icon="save" label="Сохранить" onClick={file.save} title="Сохранить (Ctrl+S)" />
              <Btn big icon="saveAs" label="Сохранить как" onClick={file.saveAs} title="Сохранить как (Ctrl+Shift+S)" />
            </Group>
            <Group title="Вывод">
              <Btn big icon="print" label="Печать" onClick={file.print} title="Печать (Ctrl+P)" />
              <Btn big icon="pdf" label="Экспорт в PDF" onClick={file.pdf} title="Сохранить как PDF" />
            </Group>
            <Group title="Последние">
              <RecentButton onOpen={file.openHere} />
              <Btn big icon="home" label="Начальный экран" onClick={() => newWindow('start')} title="Начальный экран в новом окне: создать, последние файлы, восстановление" />
            </Group>
            <Group title="Справка">
              <Btn big icon="report" label="Сообщить о проблеме" onClick={() => reportProblem()} title="Собрать сведения о проблеме для администратора" />
              <Btn big icon="open" label="Папка журналов" onClick={() => openLogs()} title="Открыть папку с журналом ошибок" />
              <Btn big icon="info" label="О программе" onClick={() => about()} title="Версия Перо" />
              <Btn big icon="props" label="Свойства" onClick={() => onRefDialog('docProps')} title="Свойства документа: название, тема, автор, ключевые слова" />
            </Group>
          </>
        )}

        {activeTab === 'home' && (
          <>
            <Group title="Буфер обмена">
              <PasteButton editor={editor} Btn={Btn} nativePaste={() => clipboard('paste')} onSpecial={() => onRefDialog('pasteSpecial')} />
              <div className="rb-col">
                <Btn icon="cut" title="Вырезать (Ctrl+X)" onClick={() => { editor.commands.focus(); clipboard('cut'); }} />
                <Btn icon="copy" title="Копировать (Ctrl+C)" onClick={() => { editor.commands.focus(); clipboard('copy'); }} />
                <Btn icon="painter" active={!!s.painter}
                  title="Формат по образцу: щелчок — один раз, двойной щелчок — несколько раз (Esc — закончить); Ctrl+Shift+C / Ctrl+Shift+V"
                  onClick={() => c().setPainter(s.painter ? 0 : 1).run()} onDoubleClick={() => c().setPainter(2).run()} />
              </div>
            </Group>

            <Group title="Шрифт" launcher={{ title: 'Шрифт… (Ctrl+D)', onClick: onFont }}>
              <div className="rb-rows">
                <div className="rb-row">
                  <FontFamilyBox editor={editor} value={s.font.fontFamily} />
                  <FontSizeBox editor={editor} value={s.font.fontSize} />
                  <Btn title="Увеличить размер (Ctrl+Shift+>)" onClick={() => c().changeFontSize(1).run()}><span className="glyph-grow">A<small>▲</small></span></Btn>
                  <Btn title="Уменьшить размер (Ctrl+Shift+<)" onClick={() => c().changeFontSize(-1).run()}><span className="glyph-grow small">A<small>▼</small></span></Btn>
                  <CaseButton editor={editor} />
                  <Btn icon="clear" title="Очистить форматирование (Ctrl+Пробел)" onClick={() => c().clearFormatting().run()} />
                </div>
                <div className="rb-row">
                  <Btn title="Полужирный (Ctrl+B)" active={s.bold} onClick={() => c().toggleBold().run()}><b className="glyph">Ж</b></Btn>
                  <Btn title="Курсив (Ctrl+I)" active={s.italic} onClick={() => c().toggleItalic().run()}><i className="glyph serif">К</i></Btn>
                  <UnderlineButton editor={editor} kind={s.underlineKind} />
                  <Btn title="Зачёркнутый" active={s.strike} onClick={() => c().toggleStrike().run()}><s className="glyph">abc</s></Btn>
                  <Btn title="Подстрочный знак (Ctrl+=)" active={s.font.subscript} onClick={() => c().toggleSubscript().run()}><span className="glyph">x<sub>2</sub></span></Btn>
                  <Btn title="Надстрочный знак (Ctrl+Shift++)" active={s.font.superscript} onClick={() => c().toggleSuperscript().run()}><span className="glyph">x<sup>2</sup></span></Btn>
                  <span className="rb-sep" />
                  <HighlightButton editor={editor} />
                  <TextColorButton editor={editor} />
                </div>
              </div>
            </Group>

            <Group title="Абзац" launcher={{ title: 'Параметры абзаца…', onClick: onParagraph }}>
              <div className="rb-rows">
                <div className="rb-row">
                  <BulletButton editor={editor} Btn={Btn} active={s.bullet} onDialog={onRefDialog} />
                  <NumberButton editor={editor} Btn={Btn} active={s.ordered} onDialog={onRefDialog} />
                  <MultilevelButton editor={editor} Btn={Btn} state={s.numbering} onDialog={onRefDialog} />
                  {/* в списке — уровень списка, иначе левый отступ абзаца шагом 1,25 см */}
                  <Btn icon="outdent" title={s.para.inList ? 'Уменьшить уровень списка (Shift+Tab)' : 'Уменьшить отступ (Ctrl+Shift+M)'}
                    disabled={s.para.inList ? !s.canLift : s.para.left <= 0} onClick={() => c().changeIndent(-1).run()} />
                  <Btn icon="indent" title={s.para.inList ? 'Увеличить уровень списка (Tab)' : 'Увеличить отступ (Ctrl+M)'}
                    disabled={s.para.inList && !s.canSink} onClick={() => c().changeIndent(1).run()} />
                </div>
                <div className="rb-row">
                  <Btn icon="alignLeft" title="По левому краю (Ctrl+Shift+L)" active={s.align === 'left'} onClick={() => c().setTextAlign('left').run()} />
                  <Btn icon="alignCenter" title="По центру (Ctrl+Shift+E)" active={s.align === 'center'} onClick={() => c().setTextAlign('center').run()} />
                  <Btn icon="alignRight" title="По правому краю (Ctrl+Shift+R)" active={s.align === 'right'} onClick={() => c().setTextAlign('right').run()} />
                  <Btn icon="alignJustify" title="По ширине (Ctrl+Shift+J)" active={s.align === 'justify'} onClick={() => c().setTextAlign('justify').run()} />
                  <LineSpacingButton editor={editor} para={s.para} onMore={onParagraph} />
                  <Btn icon="sort" title="Сортировка: абзацы, пункты списка, строки таблицы" onClick={() => onRefDialog('sortPara')} />
                  <ParaShadingButton editor={editor} />
                  <Btn title="Отобразить все знаки (Ctrl+Shift+8)" active={s.showMarks} onClick={() => c().toggleFormattingMarks().run()}><span className="glyph">¶</span></Btn>
                  <ParaBordersButton editor={editor} />
                </div>
              </div>
            </Group>

            <Group title="Редактирование">
              <div className="rb-rows">
                <MenuSplit Btn={Btn} icon="search" label="Найти" title="Найти (Ctrl+F)" onClick={() => onFind('find')} items={[
                  ['Найти… (Ctrl+F)', () => onFind('find')],
                  ['Расширенный поиск…', () => onFind('find', { more: true })],
                  ['Перейти… (Ctrl+G)', () => onRefDialog('goto')],
                ]} />
                <Btn icon="replace" title="Заменить (Ctrl+H)" onClick={() => onFind('replace')}><span className="rb-inline">Заменить</span></Btn>
                <MenuSplit Btn={Btn} icon="select" label="Выделить" title="Выделить" items={[
                  ['Выделить всё (Ctrl+A)', () => editor.chain().focus().selectAll().run()],
                  ['Текст с похожим форматированием', () => onFind('find', { format: similarFormat(editor) })],
                ]} />
              </div>
            </Group>

            <Group title="Стили" launcher={{ title: 'Панель «Стили» (Ctrl+Alt+Shift+S)', onClick: onStylesPane }}>
              <div className="style-gallery">
                {styleEntries(s).map((st) => (
                  <button
                    key={st.key}
                    className={`style-card ${st.cls || ''}${st.active ? ' active' : ''}${st.kind === 'char' ? ' char' : ''}`}
                    title={st.kind === 'char' ? `${st.label} — символьный стиль` : st.label}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => st.apply(editor)}
                  >
                    <span className="style-sample" style={st.def ? sampleOf(st.def) : undefined}>АаБбВв</span>
                    <span className="style-name">{st.label}</span>
                  </button>
                ))}
              </div>
              <StylesMenu s={s} onStyle={onStyle} />
            </Group>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <Group title="Страницы">
              <Btn big icon="blankPage" label="Пустая страница" title="Вставить пустую страницу" onClick={() => c().insertBlankPage().run()} />
              <Btn big icon="pageBreak" label="Разрыв страницы" title="Следующий текст — с новой страницы (Ctrl+Enter)" onClick={() => c().insertBreak('page').run()} />
            </Group>
            <Group title="Таблицы">
              <div className="rb-popover-anchor" ref={tablePicker.ref}>
                <Btn big icon="table" label="Таблица" onClick={() => tablePicker.setOpen(!tablePicker.open)} />
                {tablePicker.open && (
                  <TablePicker
                    style={tablePicker.style}
                    items={[
                      ['Вставить таблицу…', () => { tablePicker.setOpen(false); onRefDialog('insertTable'); }],
                      ['Преобразовать в таблицу…', () => { tablePicker.setOpen(false); onRefDialog('textToTable'); }, s.inTable],
                      ['Преобразовать в текст…', () => { tablePicker.setOpen(false); onRefDialog('tableToText'); }, !s.inTable],
                    ]}
                    onPick={(rows, cols) => {
                      tablePicker.setOpen(false);
                      c().insertTable({ rows, cols, withHeaderRow: false }).run();
                    }}
                  />
                )}
              </div>
            </Group>
            <Group title="Ссылки">
              <Btn big icon="link" label="Ссылка" title="Вставить ссылку (Ctrl+K)" onClick={onLink} />
              <Btn big icon="unlink" label="Убрать ссылку" disabled={!s.link} onClick={() => c().extendMarkRange('link').unsetLink().run()} />
              <div className="rb-col">
                <Btn title="Закладка: имя для места в документе (для ссылок и перехода)" onClick={() => onRefDialog('bookmark')}><span className="rb-inline">Закладка…</span></Btn>
                <Btn title="Перекрёстная ссылка на заголовок или закладку" onClick={() => onRefDialog('xref')}><span className="rb-inline">Перекрёстная ссылка…</span></Btn>
              </div>
            </Group>
            <Group title="Иллюстрации">
              <Btn big icon="picture" label="Рисунок" title="Вставить рисунок из файла (можно и перетащить или вставить из буфера)"
                onClick={async () => { const f = await openImage(); if (f) insertImageFiles(editor, [f]); }} />
            </Group>
            <Group title="Текст">
              <Btn big icon="lineBreak" label="Разрыв строки" title="Разрыв строки (Shift+Enter)" onClick={() => c().setHardBreak().run()} />
              <QuickPartsButton editor={editor} Btn={Btn} file={docFile} onDialog={onRefDialog} onInsertFile={onInsertFile} />
              <Btn big icon="calendar" label="Дата и время" title="Вставить текущую дату и время" onClick={() => onRefDialog('date')} />
            </Group>
            <Group title="Примечания">
              <Btn big icon="comment" label="Примечание" title="Примечание к выделенному фрагменту" onClick={review.addComment} />
            </Group>
            <HfGroup editor={editor} Group={Group} onHf={onHf} />
            <Group title="Символы">
              <SymbolButton editor={editor} Btn={Btn} onMore={onSymbol} />
            </Group>
          </>
        )}

        {activeTab === 'refs' && (
          <>
            <Group title="Оглавление">
              <TocButton editor={editor} Btn={Btn} />
              <Btn title="Обновить оглавление: заголовки и номера страниц" onClick={() => refreshToc(editor)}><span className="rb-inline">Обновить таблицу</span></Btn>
            </Group>
            <Group title="Сноски">
              <Btn big icon="footnote" label="Вставить сноску" title="Сноска внизу страницы (Ctrl+Alt+F)" onClick={() => c().insertFootnote().run()} />
              <div className="rb-rows">
                <Btn title="Следующая сноска" onClick={() => c().goToFootnote(1).run()}><span className="rb-inline">Следующая ▶</span></Btn>
                <Btn title="Предыдущая сноска" onClick={() => c().goToFootnote(-1).run()}><span className="rb-inline">◀ Предыдущая</span></Btn>
                <Btn title="Показать панель сносок" active={review.showNotes} onClick={review.toggleNotes}><span className="rb-inline">Показать сноски</span></Btn>
              </div>
            </Group>
            <Group title="Ссылки">
              <Btn big icon="xref" label="Перекрёстная ссылка" title="Ссылка на заголовок или закладку (текст, номер страницы)" onClick={() => onRefDialog('xref')} />
              <Btn big icon="bookmark" label="Закладка" title="Закладка: имя для места в документе" onClick={() => onRefDialog('bookmark')} />
            </Group>
            <Group title="Поля">
              <Btn big label="Обновить поля" title="Обновить все поля документа: оглавление, номера страниц, даты, ссылки (F9 — в выделении)"
                onClick={() => { c().updateFields(true).run(); setTimeout(() => editor.commands.updateFields(true), 400); }} icon="refresh" />
            </Group>
          </>
        )}

        {activeTab === 'picture' && s.image && <PicturePanel editor={editor} image={s.image} Group={Group} />}

        {activeTab === 'layout' && <LayoutPanel editor={editor} Group={Group} onMargins={onMargins} onHf={onHf} />}

        {activeTab === 'design' && <DesignPanel Group={Group} Btn={Btn} editor={editor} onDialog={onRefDialog} />}

        {activeTab === 'view' && <ViewPanel Group={Group} Btn={Btn} editor={editor} view={view} showMarks={s.showMarks} />}

        {activeTab === 'review' && (
          <>
            <Group title="Примечания">
              <Btn big icon="comment" label="Создать примечание" title="Примечание к выделенному фрагменту" onClick={review.addComment} />
              <div className="rb-rows">
                <Btn title="Предыдущее примечание" onClick={() => review.stepComment(-1)}><span className="rb-inline">◀ Предыдущее</span></Btn>
                <Btn title="Следующее примечание" onClick={() => review.stepComment(1)}><span className="rb-inline">Следующее ▶</span></Btn>
              </div>
              <Btn big icon="comments" label="Показать примечания" active={review.showComments} onClick={review.toggleComments} />
            </Group>
            <Group title="Исправления">
              <Btn big icon="track" label="Исправления" active={s.track.enabled}
                title="Режим исправлений: правки записываются и показываются, их можно принять или отклонить"
                onClick={() => c().setTrackChanges(!s.track.enabled).run()} />
              <div className="rb-rows">
                <Btn title="Принять исправление под курсором или в выделении" disabled={!s.track.count} onClick={() => c().resolveChange(true).run()}>
                  <Icon.accept /> <span className="rb-inline">Принять</span>
                </Btn>
                <Btn title="Отклонить исправление под курсором или в выделении" disabled={!s.track.count} onClick={() => c().resolveChange(false).run()}>
                  <Icon.reject /> <span className="rb-inline">Отклонить</span>
                </Btn>
              </div>
              <div className="rb-rows">
                <Btn title="Принять все исправления" disabled={!s.track.count} onClick={() => c().resolveAllChanges(true).run()}><span className="rb-inline">Принять все</span></Btn>
                <Btn title="Отклонить все исправления" disabled={!s.track.count} onClick={() => c().resolveAllChanges(false).run()}><span className="rb-inline">Отклонить все</span></Btn>
              </div>
              <div className="rb-rows">
                <Btn title="Предыдущее исправление" disabled={!s.track.count} onClick={() => c().goToChange(-1).run()}><span className="rb-inline">◀ Предыдущее</span></Btn>
                <Btn title="Следующее исправление" disabled={!s.track.count} onClick={() => c().goToChange(1).run()}><span className="rb-inline">Следующее ▶</span></Btn>
              </div>
            </Group>
          </>
        )}

        {activeTab === 'table' && (
          <>
            <Group title="Строки и столбцы">
              <Btn big icon="rowAbove" label="Вставить сверху" onClick={() => c().insertRowFormatted(-1).run()} />
              <Btn big icon="rowBelow" label="Вставить снизу" onClick={() => c().insertRowFormatted(1).run()} />
              <Btn big icon="colLeft" label="Вставить слева" onClick={() => c().insertColumnFormatted(-1).run()} />
              <Btn big icon="colRight" label="Вставить справа" onClick={() => c().insertColumnFormatted(1).run()} />
            </Group>
            <Group title="Удаление">
              <Btn big icon="delRow" label="Удалить строку" onClick={() => c().deleteRow().run()} />
              <Btn big icon="delCol" label="Удалить столбец" onClick={() => c().deleteColumn().run()} />
              <Btn big icon="delTable" label="Удалить таблицу" onClick={() => c().deleteTable().run()} />
            </Group>
            <Group title="Стили таблиц">
              <TableStylesButton editor={editor} />
            </Group>
            <Group title="Объединение">
              <Btn big icon="merge" label="Объединить ячейки" disabled={!s.canMerge} onClick={() => c().mergeCells().run()} />
              <Btn big icon="split" label="Разделить ячейки" disabled={!s.canSplit} onClick={() => c().splitCell().run()} />
            </Group>
            {s.table && (
              <>
                <Group title="Оформление">
                  <div className="rb-rows">
                    <div className="rb-row"><BordersButton editor={editor} /></div>
                    <div className="rb-row"><ShadingButton editor={editor} /></div>
                  </div>
                </Group>
                <Group title="Выравнивание">
                  <div className="rb-rows">
                    <div className="rb-row">
                      {[['top', 'Сверху'], ['center', 'По центру'], ['bottom', 'Снизу']].map(([v, l]) => (
                        <Btn key={v} title={`Выравнивание в ячейке: ${l.toLowerCase()}`} active={s.table.verticalAlign === v}
                          onClick={() => c().setCellAttribute('verticalAlign', v === 'top' ? null : v).run()}><TableIcons.VAlignIcon v={v} /></Btn>
                      ))}
                    </div>
                    <div className="rb-row">
                      {[['left', 'слева'], ['center', 'по центру'], ['right', 'справа']].map(([v, l]) => (
                        <Btn key={v} title={`Таблица на странице: ${l}`} active={s.table.align === v}
                          onClick={() => c().setTableAlign(v).run()}><TableIcons.PosIcon a={v} /></Btn>
                      ))}
                    </div>
                  </div>
                </Group>
                <Group title="Размер">
                  <div className="rb-rows">
                    <RowHeightField editor={editor} t={s.table} />
                    <Btn title="Выровнять ширину столбцов" onClick={() => c().distributeColumns().run()}><TableIcons.DistributeIcon /> <span className="rb-inline">Выровнять столбцы</span></Btn>
                  </div>
                </Group>
                <Group title="Данные">
                  <Btn big label="Повторять заголовки" title="Повторять строку как заголовок на каждой странице" active={s.table.repeatHeader}
                    onClick={() => c().toggleRepeatHeader().run()}><TableIcons.HeaderIcon /></Btn>
                </Group>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

export function QuickAccess({ editor, onSave }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({ canUndo: e.can().undo(), canRedo: e.can().redo() }),
  });
  return (
    <>
      <QuickButton icon="save" title="Сохранить (Ctrl+S)" onClick={onSave} />
      <QuickButton icon="undo" title="Отменить (Ctrl+Z)" disabled={!s.canUndo} onClick={() => editor.chain().focus().undo().run()} />
      <QuickButton icon="redo" title="Повторить (Ctrl+Y)" disabled={!s.canRedo} onClick={() => editor.chain().focus().redo().run()} />
    </>
  );
}
