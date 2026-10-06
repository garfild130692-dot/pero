// Лента в стиле классического офисного редактора: вкладки → группы → кнопки.
import React, { useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { Icon } from './icons.jsx';
import { clipboard, openImage } from './fileApi.js';
import { insertImageFiles } from './insertImage.js';
import { PicturePanel } from './PictureControls.jsx';
import { LayoutPanel } from './PageControls.jsx';
import { FontFamilyBox, FontSizeBox, HighlightButton, TextColorButton, fontState, usePopover } from './FontControls.jsx';
import { effectivePara } from './editor/paraFormat.js';
import { LINE_PRESETS, fmtNum } from './editor/paragraph.js';
import { BordersButton, RowHeightField, ShadingButton, TableIcons, tableState } from './TableControls.jsx';

function Btn({ icon, label, title, onClick, active, disabled, big, children }) {
  const I = icon && Icon[icon];
  return (
    <button
      type="button"
      className={`rb-btn${big ? ' big' : ''}${active ? ' active' : ''}`}
      title={title || label}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} // не забираем фокус у текста
      onClick={onClick}
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

const STYLES = [
  { key: 'p', label: 'Обычный', cls: 'st-normal' },
  { key: 1, label: 'Заголовок 1', cls: 'st-h1' },
  { key: 2, label: 'Заголовок 2', cls: 'st-h2' },
  { key: 3, label: 'Заголовок 3', cls: 'st-h3' },
];

function TablePicker({ onPick, style }) {
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
    </div>
  );
}

export default function Ribbon({ editor, file, onLink, onParagraph, onMargins, onHf, onFind }) {
  const [tab, setTab] = useState('home');
  const tablePicker = usePopover();

  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => ({
      bold: e.isActive('bold'),
      italic: e.isActive('italic'),
      underline: e.isActive('underline'),
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
      page: { sectPr: e.state.doc.attrs.sectPr, hf: e.state.doc.attrs.hf }, // для вкладки «Макет»
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
    ['layout', 'Макет'],
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
              <Btn big icon="open" label="Открыть" onClick={file.open} title="Открыть (Ctrl+O)" />
              <Btn big icon="save" label="Сохранить" onClick={file.save} title="Сохранить (Ctrl+S)" />
              <Btn big icon="saveAs" label="Сохранить как" onClick={file.saveAs} title="Сохранить как (Ctrl+Shift+S)" />
            </Group>
            <Group title="Вывод">
              <Btn big icon="print" label="Печать" onClick={file.print} title="Печать (Ctrl+P)" />
              <Btn big icon="pdf" label="Экспорт в PDF" onClick={file.pdf} title="Сохранить как PDF" />
            </Group>
          </>
        )}

        {activeTab === 'home' && (
          <>
            <Group title="Буфер обмена">
              <Btn big icon="paste" label="Вставить" onClick={() => { editor.commands.focus(); clipboard('paste'); }} title="Вставить (Ctrl+V)" />
              <div className="rb-col">
                <Btn icon="cut" title="Вырезать (Ctrl+X)" onClick={() => { editor.commands.focus(); clipboard('cut'); }} />
                <Btn icon="copy" title="Копировать (Ctrl+C)" onClick={() => { editor.commands.focus(); clipboard('copy'); }} />
              </div>
            </Group>

            <Group title="Шрифт">
              <div className="rb-rows">
                <div className="rb-row">
                  <FontFamilyBox editor={editor} value={s.font.fontFamily} />
                  <FontSizeBox editor={editor} value={s.font.fontSize} />
                  <Btn title="Увеличить размер (Ctrl+Shift+>)" onClick={() => c().changeFontSize(1).run()}><span className="glyph-grow">A<small>▲</small></span></Btn>
                  <Btn title="Уменьшить размер (Ctrl+Shift+<)" onClick={() => c().changeFontSize(-1).run()}><span className="glyph-grow small">A<small>▼</small></span></Btn>
                  <Btn icon="clear" title="Очистить форматирование (Ctrl+Пробел)" onClick={() => c().clearFormatting().run()} />
                </div>
                <div className="rb-row">
                  <Btn title="Полужирный (Ctrl+B)" active={s.bold} onClick={() => c().toggleBold().run()}><b className="glyph">Ж</b></Btn>
                  <Btn title="Курсив (Ctrl+I)" active={s.italic} onClick={() => c().toggleItalic().run()}><i className="glyph serif">К</i></Btn>
                  <Btn title="Подчёркнутый (Ctrl+U)" active={s.underline} onClick={() => c().toggleUnderline().run()}><u className="glyph">Ч</u></Btn>
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
                  <Btn icon="bullets" title="Маркированный список" active={s.bullet} onClick={() => c().toggleBulletList().run()} />
                  <Btn icon="numbers" title="Нумерованный список" active={s.ordered} onClick={() => c().toggleOrderedList().run()} />
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
                </div>
              </div>
            </Group>

            <Group title="Редактирование">
              <div className="rb-rows">
                <Btn icon="search" title="Найти (Ctrl+F)" onClick={() => onFind('find')}><span className="rb-inline">Найти</span></Btn>
                <Btn icon="replace" title="Заменить (Ctrl+H)" onClick={() => onFind('replace')}><span className="rb-inline">Заменить</span></Btn>
              </div>
            </Group>

            <Group title="Стили">
              <div className="style-gallery">
                {STYLES.map((st) => {
                  const active = st.key === 'p' ? !s.heading : s.heading === st.key;
                  return (
                    <button
                      key={st.key}
                      className={`style-card ${st.cls}${active ? ' active' : ''}`}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() => (st.key === 'p' ? c().setParagraph().run() : c().setHeading({ level: st.key }).run())}
                    >
                      <span className="style-sample">АаБбВв</span>
                      <span className="style-name">{st.label}</span>
                    </button>
                  );
                })}
              </div>
            </Group>
          </>
        )}

        {activeTab === 'insert' && (
          <>
            <Group title="Таблицы">
              <div className="rb-popover-anchor" ref={tablePicker.ref}>
                <Btn big icon="table" label="Таблица" onClick={() => tablePicker.setOpen(!tablePicker.open)} />
                {tablePicker.open && (
                  <TablePicker
                    style={tablePicker.style}
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
            </Group>
            <Group title="Иллюстрации">
              <Btn big icon="picture" label="Рисунок" title="Вставить рисунок из файла (можно и перетащить или вставить из буфера)"
                onClick={async () => { const f = await openImage(); if (f) insertImageFiles(editor, [f]); }} />
            </Group>
            <Group title="Текст">
              <Btn big icon="lineBreak" label="Разрыв строки" title="Разрыв строки (Shift+Enter)" onClick={() => c().setHardBreak().run()} />
            </Group>
          </>
        )}

        {activeTab === 'picture' && s.image && <PicturePanel editor={editor} image={s.image} Group={Group} />}

        {activeTab === 'layout' && <LayoutPanel editor={editor} Group={Group} onMargins={onMargins} onHf={onHf} />}

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
    <div className="quick-access">
      <span className="app-mark">П</span>
      <Btn icon="save" title="Сохранить (Ctrl+S)" onClick={onSave} />
      <Btn icon="undo" title="Отменить (Ctrl+Z)" disabled={!s.canUndo} onClick={() => editor.chain().focus().undo().run()} />
      <Btn icon="redo" title="Повторить (Ctrl+Y)" disabled={!s.canRedo} onClick={() => editor.chain().focus().redo().run()} />
    </div>
  );
}
