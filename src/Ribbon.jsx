// Лента в стиле классического офисного редактора: вкладки → группы → кнопки.
import React, { useState } from 'react';
import { useEditorState } from '@tiptap/react';
import { Icon } from './icons.jsx';
import { clipboard } from './fileApi.js';

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

const Group = ({ title, children }) => (
  <div className="rb-group">
    <div className="rb-group-body">{children}</div>
    <div className="rb-group-title">{title}</div>
  </div>
);

const STYLES = [
  { key: 'p', label: 'Обычный', cls: 'st-normal' },
  { key: 1, label: 'Заголовок 1', cls: 'st-h1' },
  { key: 2, label: 'Заголовок 2', cls: 'st-h2' },
  { key: 3, label: 'Заголовок 3', cls: 'st-h3' },
];

function TablePicker({ onPick }) {
  const [hover, setHover] = useState([0, 0]);
  const N = 8;
  return (
    <div className="table-picker" onMouseLeave={() => setHover([0, 0])}>
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

export default function Ribbon({ editor, file, onLink }) {
  const [tab, setTab] = useState('home');
  const [tablePicker, setTablePicker] = useState(false);

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
    }),
  });

  const c = () => editor.chain().focus();
  const activeTab = tab === 'table' && !s.inTable ? 'home' : tab;

  const tabs = [
    ['file', 'Файл'],
    ['home', 'Главная'],
    ['insert', 'Вставка'],
  ];

  return (
    <div className="ribbon">
      <div className="rb-tabs">
        {tabs.map(([k, label]) => (
          <button key={k} className={`rb-tab${k === 'file' ? ' file' : ''}${activeTab === k ? ' active' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setTab(k)}>
            {label}
          </button>
        ))}
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
              <div className="rb-row">
                <Btn title="Полужирный (Ctrl+B)" active={s.bold} onClick={() => c().toggleBold().run()}><b className="glyph">Ж</b></Btn>
                <Btn title="Курсив (Ctrl+I)" active={s.italic} onClick={() => c().toggleItalic().run()}><i className="glyph serif">К</i></Btn>
                <Btn title="Подчёркнутый (Ctrl+U)" active={s.underline} onClick={() => c().toggleUnderline().run()}><u className="glyph">Ч</u></Btn>
                <Btn title="Зачёркнутый" active={s.strike} onClick={() => c().toggleStrike().run()}><s className="glyph">abc</s></Btn>
                <Btn icon="clear" title="Очистить форматирование" onClick={() => c().unsetAllMarks().run()} />
              </div>
            </Group>

            <Group title="Абзац">
              <div className="rb-rows">
                <div className="rb-row">
                  <Btn icon="bullets" title="Маркированный список" active={s.bullet} onClick={() => c().toggleBulletList().run()} />
                  <Btn icon="numbers" title="Нумерованный список" active={s.ordered} onClick={() => c().toggleOrderedList().run()} />
                  <Btn icon="outdent" title="Уменьшить уровень списка (Shift+Tab)" disabled={!s.canLift} onClick={() => c().liftListItem('listItem').run()} />
                  <Btn icon="indent" title="Увеличить уровень списка (Tab)" disabled={!s.canSink} onClick={() => c().sinkListItem('listItem').run()} />
                </div>
                <div className="rb-row">
                  <Btn icon="alignLeft" title="По левому краю (Ctrl+Shift+L)" active={s.align === 'left'} onClick={() => c().setTextAlign('left').run()} />
                  <Btn icon="alignCenter" title="По центру (Ctrl+Shift+E)" active={s.align === 'center'} onClick={() => c().setTextAlign('center').run()} />
                  <Btn icon="alignRight" title="По правому краю (Ctrl+Shift+R)" active={s.align === 'right'} onClick={() => c().setTextAlign('right').run()} />
                  <Btn icon="alignJustify" title="По ширине (Ctrl+Shift+J)" active={s.align === 'justify'} onClick={() => c().setTextAlign('justify').run()} />
                </div>
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
              <div className="rb-popover-anchor">
                <Btn big icon="table" label="Таблица" onClick={() => setTablePicker((v) => !v)} />
                {tablePicker && (
                  <TablePicker
                    onPick={(rows, cols) => {
                      setTablePicker(false);
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
            <Group title="Текст">
              <Btn big icon="lineBreak" label="Разрыв строки" title="Разрыв строки (Shift+Enter)" onClick={() => c().setHardBreak().run()} />
            </Group>
          </>
        )}

        {activeTab === 'table' && (
          <>
            <Group title="Строки и столбцы">
              <Btn big icon="rowAbove" label="Вставить сверху" onClick={() => c().addRowBefore().run()} />
              <Btn big icon="rowBelow" label="Вставить снизу" onClick={() => c().addRowAfter().run()} />
              <Btn big icon="colLeft" label="Вставить слева" onClick={() => c().addColumnBefore().run()} />
              <Btn big icon="colRight" label="Вставить справа" onClick={() => c().addColumnAfter().run()} />
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
