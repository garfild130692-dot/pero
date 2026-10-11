// Лента табличного редактора: Файл / Главная (буфер, шрифт, выравнивание, число, ячейки, редактирование) /
// Разметка страницы (параметры печати) / Формулы (имена) / Данные (структура) / Рецензирование (защита листа) / Вид (закрепление областей).
import CfMenu from './CfMenu.jsx';
import React, { useState } from 'react';
import { Icon } from '../icons.jsx';
import { about, clipboard, openLogs, reportProblem } from '../fileApi.js';
import RecentButton from '../RecentButton.jsx';
import { ColorPalette, SplitColor, usePopover } from '../FontControls.jsx';
import { FORMAT_PRESETS, changeDecimals } from './format.js';
import { MARGIN_PRESETS, PAPERS, pageOf, setPage, setPrintArea, setPrintTitles } from './pageSetup.js';
import { isProtected } from './protect.js';
import { CATEGORIES, FN_LIST } from './fninfo.js';
import { key } from './address.js';
import {
  clearRange, collapseGroup, expandGroup, groupAt, groupLines, mergeAt, mergeRange, setFreeze, setHidden, setStyle, shiftCells, sortRange, unmergeRange,
} from './model.js';

function Btn({ icon, label, title, onClick, active, disabled, big, children }) {
  const I = icon && Icon[icon];
  return (
    <button type="button" className={`rb-btn${big ? ' big' : ''}${active ? ' active' : ''}`} title={title || label} disabled={disabled}
      onMouseDown={(e) => e.preventDefault()} onClick={onClick}>
      {I && <I size={big ? 28 : 18} />}
      {children}
      {big && label && <span className="rb-label">{label}</span>}
    </button>
  );
}
const Group = ({ title, children }) => (
  <div className="rb-group"><div className="rb-group-body">{children}</div><div className="rb-group-title">{title}</div></div>
);
const T = ({ children, style }) => <span className="txt-icon" style={style}>{children}</span>;

// Выпадающее меню с пунктами [подпись, действие]
function MenuButton({ label, title, icon, items }) {
  const pop = usePopover();
  return (
    <div className="rb-popover-anchor" ref={pop.ref}>
      <Btn icon={icon} title={title} active={pop.open} onClick={() => pop.setOpen(!pop.open)}>{label}<span className="dd">▾</span></Btn>
      {pop.open && (
        <div className="menu-pop" style={pop.style} role="menu">
          {items.map((it, i) => (it === '-' ? <div key={i} className="menu-sep" /> : (
            <button key={i} type="button" className="menu-item" onMouseDown={(e) => e.preventDefault()} onClick={() => { pop.setOpen(false); it[1](); }}>
              <span className="check">{it[2] ? '✓' : ''}</span>{it[0]}
            </button>
          )))}
        </div>
      )}
    </div>
  );
}

function ColorButton({ title, glyph, auto, onApply, initial }) {
  const [last, setLast] = useState(initial);
  const apply = (c) => { setLast(c); onApply(c); };
  return (
    <SplitColor title={title} color={last} icon={glyph}>
      {{
        apply: () => apply(last),
        palette: (close) => (
          <ColorPalette autoLabel={auto} autoClass={auto === 'Нет заливки' ? 'none' : 'auto'} current={last}
            onAuto={() => { close(); onApply(null); }} onPick={(x) => { close(); apply(x); }} />
        ),
      }}
    </SplitColor>
  );
}

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Cambria', 'Verdana', 'Tahoma', 'Segoe UI', 'Courier New', 'Consolas'];

export default function SheetRibbon({ api }) {
  const [tab, setTab] = useState('home');
  const { sheet, si, rg, sel, change, file } = api;
  const cur = sheet.cells.get(key(sel.r, sel.c));
  const s = (cur && cur.s) || {};
  const style = (patch) => change((w) => setStyle(w, si, rg, patch));
  const toggle = (name) => style({ [name]: !s[name] });
  api.toggle = toggle;

  const merged = mergeAt(sheet, sel.r, sel.c);
  const multi = rg.r2 > rg.r1 || rg.c2 > rg.c1;
  const fmt = s.numFmt || 'General';
  const sample = cur && typeof cur.v === 'number' ? cur.v : 0;

  // границы: для каждой ячейки диапазона — свои стороны
  const borders = (mode) => change((w) => {
    let out = w;
    const line = { style: mode === 'thick' ? 'medium' : 'thin', color: null };
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const cell = out.sheets[si].cells.get(key(r, c));
        const b = { ...((cell && cell.s && cell.s.border) || {}) };
        if (mode === 'none') { for (const n of ['left', 'right', 'top', 'bottom']) delete b[n]; }
        else if (mode === 'all') { b.left = b.right = b.top = b.bottom = line; }
        else if (mode === 'bottom') { if (r === rg.r2) b.bottom = line; else continue; }
        else if (mode === 'outer' || mode === 'thick') {
          if (r === rg.r1) b.top = line; if (r === rg.r2) b.bottom = line;
          if (c === rg.c1) b.left = line; if (c === rg.c2) b.right = line;
          if (r !== rg.r1 && r !== rg.r2 && c !== rg.c1 && c !== rg.c2) continue;
        }
        out = setStyle(out, si, { r1: r, c1: c, r2: r, c2: c }, { border: Object.keys(b).length ? b : null });
      }
    }
    return out;
  });

  // выделены целые столбцы (щелчок по заголовку) — группировка и скрытие по столбцам
  const byCols = sel.whole === 'col';
  const lines = (axis) => (axis === 'row' ? [rg.r1, rg.r2] : [rg.c1, rg.c2]);
  const hide = (axis, hidden) => change((w) => setHidden(w, si, axis, ...lines(axis), hidden));
  const group = (axis, delta) => change((w) => groupLines(w, si, axis, ...lines(axis), delta));
  // скрыть/показать детали группы под курсором
  const detail = (show) => change((w) => {
    const axis = byCols ? 'col' : 'row';
    const sh = w.sheets[si];
    const at = axis === 'row' ? sel.r : sel.c;
    // курсор на итоговой строке группы — тоже её группа
    const g = groupAt(sh, axis, at) || groupAt(sh, axis, at - 1);
    if (!g) return w;
    return show ? expandGroup(w, si, axis, g.a, g.b, g.level) : collapseGroup(w, si, axis, g.a, g.b);
  });
  const frozen = sheet.freeze || null;

  // параметры страницы
  const page = pageOf(sheet);
  const setPg = (patch) => change((w) => setPage(w, si, patch));
  const fitMode = page.fitWidth == null && page.fitHeight == null ? 'none'
    : page.fitWidth === 1 && page.fitHeight === 1 ? 'page' : page.fitWidth === 1 && !page.fitHeight ? 'cols' : !page.fitWidth && page.fitHeight === 1 ? 'rows' : 'custom';
  const sameMargins = (m) => ['left', 'right', 'top', 'bottom'].every((k) => Math.abs(page.margins[k] - m[k]) < 0.01);
  const titles = sheet.printTitles || {};

  const prot = isProtected(sheet);
  const tabs = [['file', 'Файл'], ['home', 'Главная'], ['layout', 'Разметка страницы'], ['formulas', 'Формулы'], ['data', 'Данные'], ['review', 'Рецензирование'], ['view', 'Вид']];
  return (
    <div className="ribbon">
      <div className="rb-tabs">
        {tabs.map(([k, label]) => (
          <button key={k} className={`rb-tab${k === 'file' ? ' file' : ''}${tab === k ? ' active' : ''}`} onMouseDown={(e) => e.preventDefault()} onClick={() => setTab(k)}>{label}</button>
        ))}
      </div>
      <div className="rb-panel">
        {tab === 'file' && (
          <>
            <Group title="Создать">
              <Btn big icon="file" label="Документ" onClick={file.newDocument} title="Новый документ (Ctrl+N)" />
              <Btn big icon="table" label="Книга" onClick={file.newBook} title="Новая книга" />
            </Group>
            <Group title="Файл">
              <Btn big icon="open" label="Открыть" onClick={file.open} title="Открыть (Ctrl+O)" />
              <Btn big icon="save" label="Сохранить" onClick={file.save} title="Сохранить (Ctrl+S)" />
              <Btn big icon="saveAs" label="Сохранить как" onClick={file.saveAs} title="Сохранить как (Ctrl+Shift+S)" />
              <Btn big icon="file" label="CSV" onClick={file.saveCsv} title="Сохранить текущий лист в CSV (разделитель «;», UTF-8)" />
            </Group>
            <Group title="Вывод">
              <Btn big icon="print" label="Печать" onClick={file.print} title="Печать (Ctrl+P)" />
              <Btn big icon="pdf" label="Экспорт в PDF" onClick={file.pdf} title="Сохранить как PDF" />
            </Group>
            <Group title="Последние">
              <RecentButton onOpen={file.openHere} />
            </Group>
            <Group title="Справка">
              <Btn big icon="report" label="Сообщить о проблеме" onClick={() => reportProblem()} title="Собрать сведения о проблеме для администратора" />
              <Btn big icon="open" label="Папка журналов" onClick={() => openLogs()} title="Открыть папку с журналом ошибок" />
              <Btn big icon="info" label="О программе" onClick={() => about()} title="Версия Перо" />
            </Group>
          </>
        )}
        {tab === 'home' && (
          <>
            <Group title="Буфер обмена">
              <Btn big icon="paste" label="Вставить" title="Вставить (Ctrl+V)" onClick={() => clipboard('paste')} />
              <div className="rb-col">
                <Btn icon="cut" title="Вырезать (Ctrl+X)" onClick={() => clipboard('cut')} />
                <Btn icon="copy" title="Копировать (Ctrl+C)" onClick={() => clipboard('copy')} />
              </div>
            </Group>
            <Group title="Шрифт">
              <div className="rb-col">
                <div className="rb-row">
                  <select className="rb-select" title="Шрифт" style={{ width: 130 }} value={s.fontFamily || 'Calibri'}
                    onChange={(e) => style({ fontFamily: e.target.value === 'Calibri' ? null : e.target.value })}>
                    {[...new Set([s.fontFamily || 'Calibri', ...FONTS])].map((f) => <option key={f} value={f}>{f}</option>)}
                  </select>
                  <select className="rb-select" title="Размер шрифта" style={{ width: 52 }} value={s.fontSize || 11}
                    onChange={(e) => style({ fontSize: +e.target.value === 11 ? null : +e.target.value })}>
                    {[...new Set([s.fontSize || 11, ...SIZES])].sort((a, b) => a - b).map((n) => <option key={n} value={n}>{n}</option>)}
                  </select>
                </div>
                <div className="rb-row">
                  <Btn title="Полужирный (Ctrl+B)" active={!!s.bold} onClick={() => toggle('bold')}><T style={{ fontWeight: 700 }}>Ж</T></Btn>
                  <Btn title="Курсив (Ctrl+I)" active={!!s.italic} onClick={() => toggle('italic')}><T style={{ fontStyle: 'italic', fontFamily: 'Georgia, serif' }}>К</T></Btn>
                  <Btn title="Подчёркнутый (Ctrl+U)" active={!!s.underline} onClick={() => toggle('underline')}><T style={{ textDecoration: 'underline' }}>Ч</T></Btn>
                  <Btn title="Зачёркнутый" active={!!s.strike} onClick={() => toggle('strike')}><T style={{ textDecoration: 'line-through' }}>abc</T></Btn>
                  <MenuButton title="Границы" label={<T>▦</T>} items={[
                    ['Все границы', () => borders('all')], ['Внешние границы', () => borders('outer')],
                    ['Толстая внешняя граница', () => borders('thick')], ['Нижняя граница', () => borders('bottom')], '-',
                    ['Нет границы', () => borders('none')],
                  ]} />
                  <ColorButton title="Цвет заливки" auto="Нет заливки" initial="#ffff00" glyph={<T>◩</T>} onApply={(c) => style({ fill: c })} />
                  <ColorButton title="Цвет текста" auto="Авто" initial="#c00000" glyph={<span className="glyph-a">A</span>} onApply={(c) => style({ color: c })} />
                </div>
              </div>
            </Group>
            <Group title="Выравнивание">
              <div className="rb-col">
                <div className="rb-row">
                  <Btn title="По верхнему краю" active={s.valign === 'top'} onClick={() => style({ valign: s.valign === 'top' ? null : 'top' })}><T>⤒</T></Btn>
                  <Btn title="По середине" active={s.valign === 'center'} onClick={() => style({ valign: s.valign === 'center' ? null : 'center' })}><T>↕</T></Btn>
                  <Btn title="По нижнему краю" active={!s.valign || s.valign === 'bottom'} onClick={() => style({ valign: null })}><T>⤓</T></Btn>
                  <Btn title="Переносить текст" active={!!s.wrap} onClick={() => toggle('wrap')}><T>↵</T></Btn>
                </div>
                <div className="rb-row">
                  <Btn icon="alignLeft" title="По левому краю" active={s.align === 'left'} onClick={() => style({ align: s.align === 'left' ? null : 'left' })} />
                  <Btn icon="alignCenter" title="По центру" active={s.align === 'center'} onClick={() => style({ align: s.align === 'center' ? null : 'center' })} />
                  <Btn icon="alignRight" title="По правому краю" active={s.align === 'right'} onClick={() => style({ align: s.align === 'right' ? null : 'right' })} />
                  <Btn icon="merge" title={merged ? 'Отменить объединение ячеек' : 'Объединить и поместить в центре'} active={!!merged} disabled={!merged && !multi}
                    onClick={() => change((w) => (merged ? unmergeRange(w, si, rg) : setStyle(mergeRange(w, si, rg), si, { r1: rg.r1, c1: rg.c1, r2: rg.r1, c2: rg.c1 }, { align: 'center' })))} />
                </div>
              </div>
            </Group>
            <Group title="Число">
              <div className="rb-col">
                <select className="rb-select" title="Числовой формат" style={{ width: 140 }}
                  value={FORMAT_PRESETS.some(([c]) => c === fmt) ? fmt : '*'}
                  onChange={(e) => e.target.value !== '*' && style({ numFmt: e.target.value === 'General' ? null : e.target.value })}>
                  {FORMAT_PRESETS.map(([c, label]) => <option key={c} value={c}>{label}</option>)}
                  {!FORMAT_PRESETS.some(([c]) => c === fmt) && <option value="*">{fmt}</option>}
                </select>
                <div className="rb-row">
                  <Btn title="Денежный формат" onClick={() => style({ numFmt: '#,##0.00 "₽"' })}><T>₽</T></Btn>
                  <Btn title="Процентный формат" onClick={() => style({ numFmt: '0%' })}><T>%</T></Btn>
                  <Btn title="Формат с разделителями" onClick={() => style({ numFmt: '#,##0.00' })}><T>000</T></Btn>
                  <Btn title="Увеличить разрядность" onClick={() => style({ numFmt: changeDecimals(fmt, 1, sample) })}><T>,0→</T></Btn>
                  <Btn title="Уменьшить разрядность" onClick={() => style({ numFmt: changeDecimals(fmt, -1, sample) })}><T>←,0</T></Btn>
                </div>
              </div>
            </Group>
            <Group title="Стили">
              <CfMenu onDialog={api.cfDialog} onRule={api.addCf} onClear={api.clearCf} />
            </Group>
            <Group title="Ячейки">
              <div className="rb-col">
              <MenuButton icon="rowAbove" title="Вставить" label="Вставить" items={[
                ['Вставить строки на лист', () => change((w) => shiftCells(w, si, 'row', rg.r1, rg.r2 - rg.r1 + 1))],
                ['Вставить столбцы на лист', () => change((w) => shiftCells(w, si, 'col', rg.c1, rg.c2 - rg.c1 + 1))],
              ]} />
              <MenuButton icon="delRow" title="Удалить" label="Удалить" items={[
                ['Удалить строки с листа', () => change((w) => shiftCells(w, si, 'row', rg.r1, -(rg.r2 - rg.r1 + 1)))],
                ['Удалить столбцы с листа', () => change((w) => shiftCells(w, si, 'col', rg.c1, -(rg.c2 - rg.c1 + 1)))],
              ]} />
              </div>
              <MenuButton title="Формат: скрыть или отобразить" label="Формат" items={[
                ['Скрыть строки', () => hide('row', true)],
                ['Скрыть столбцы', () => hide('col', true)],
                '-',
                ['Отобразить строки', () => hide('row', false)],
                ['Отобразить столбцы', () => hide('col', false)],
                '-',
                ['Блокировать ячейку (действует при защите листа)', () => style({ unlocked: !s.unlocked }), !s.unlocked],
                [prot ? 'Снять защиту листа…' : 'Защитить лист…', api.protect],
              ]} />
            </Group>
            <Group title="Редактирование">
              <Btn title="Автосумма (Alt+=)" onClick={api.autoSum}><T>Σ</T> Автосумма</Btn>
              <MenuButton title="Сортировка" label="Сортировка" items={[
                ['Сортировка от А до Я / по возрастанию', () => change((w) => sortRange(w, si, rg, sel.c, true))],
                ['Сортировка от Я до А / по убыванию', () => change((w) => sortRange(w, si, rg, sel.c, false))],
                ['Настраиваемая сортировка…', api.sortDialog],
                '-',
                ['Фильтр', api.toggleFilter, api.filterOn],
              ]} />
              <MenuButton icon="clear" title="Очистить" label="" items={[
                ['Очистить всё', () => change((w) => clearRange(w, si, rg, 'all'))],
                ['Очистить форматы', () => change((w) => clearRange(w, si, rg, 'formats'))],
                ['Очистить содержимое', () => change((w) => clearRange(w, si, rg, 'contents'))],
              ]} />
            </Group>
          </>
        )}
        {tab === 'layout' && (
          <>
            <Group title="Параметры страницы">
              <MenuButton title="Поля" label="Поля" items={Object.values(MARGIN_PRESETS).map((m) => [m.name, () => setPg({ margins: { left: m.left, right: m.right, top: m.top, bottom: m.bottom, header: m.header, footer: m.footer } }), sameMargins(m)])} />
              <MenuButton title="Ориентация" label="Ориентация" items={[
                ['Книжная', () => setPg({ orientation: 'portrait' }), page.orientation !== 'landscape'],
                ['Альбомная', () => setPg({ orientation: 'landscape' }), page.orientation === 'landscape'],
              ]} />
              <MenuButton title="Размер бумаги" label="Размер" items={Object.entries(PAPERS).map(([code, p]) => [`${p.name} (${p.w} × ${p.h} мм)`, () => setPg({ paper: +code }), page.paper === +code])} />
              <MenuButton title="Область печати" label="Область печати" items={[
                ['Задать (выделенный диапазон)', () => change((w) => setPrintArea(w, si, rg))],
                ['Убрать', () => change((w) => setPrintArea(w, si, null)), false],
              ]} />
              <MenuButton title="Сквозные строки и столбцы: печатаются на каждой странице" label="Печатать заголовки" items={[
                ['Сквозные строки — выделенные строки', () => change((w) => setPrintTitles(w, si, { ...titles, rows: [rg.r1, rg.r2] })), !!titles.rows],
                ['Сквозные столбцы — выделенные столбцы', () => change((w) => setPrintTitles(w, si, { ...titles, cols: [rg.c1, rg.c2] })), !!titles.cols],
                '-',
                ['Убрать сквозные строки и столбцы', () => change((w) => setPrintTitles(w, si, null))],
              ]} />
            </Group>
            <Group title="Вписать">
              <MenuButton title="Масштаб при печати" label={fitMode === 'none' ? `Масштаб ${page.scale || 100}%` : 'Вписать'} items={[
                ['Текущий размер', () => setPg({ fitWidth: null, fitHeight: null, scale: 100 }), fitMode === 'none' && (page.scale || 100) === 100],
                ['Вписать лист на одну страницу', () => setPg({ fitWidth: 1, fitHeight: 1 }), fitMode === 'page'],
                ['Вписать все столбцы на одну страницу', () => setPg({ fitWidth: 1, fitHeight: 0 }), fitMode === 'cols'],
                ['Вписать все строки на одну страницу', () => setPg({ fitWidth: 0, fitHeight: 1 }), fitMode === 'rows'],
                '-',
                ...[90, 75, 50].map((k) => [`Масштаб ${k}%`, () => setPg({ fitWidth: null, fitHeight: null, scale: k }), fitMode === 'none' && page.scale === k]),
              ]} />
            </Group>
            <Group title="Параметры листа">
              <div className="rb-col">
                <Btn title="Печатать линии сетки" active={page.gridLines} onClick={() => setPg({ gridLines: !page.gridLines })}><T>▦</T> Печать сетки</Btn>
                <Btn title="Центрировать на странице по горизонтали" active={page.centerH} onClick={() => setPg({ centerH: !page.centerH })}><T>↔</T> По центру</Btn>
              </div>
            </Group>
            <Group title="Печать">
              <Btn big icon="print" label="Печать" onClick={file.print} title="Печать (Ctrl+P)" />
              <Btn big icon="pdf" label="PDF" onClick={file.pdf} title="Экспорт в PDF" />
            </Group>
          </>
        )}
        {tab === 'formulas' && (
          <>
            <Group title="Библиотека функций">
              <Btn big icon="fx" label="Вставить функцию" onClick={() => api.insertFunction()} title="Мастер функций (Shift+F3)" />
              <div className="rb-col">
                <Btn title="Автосумма (Alt+=)" onClick={api.autoSum}><T>Σ</T> Автосумма</Btn>
                <MenuButton title="Функции по категориям" label="Категории" items={CATEGORIES.map((c) => [c, () => api.insertFunction(FN_LIST.find((f) => f.category === c).en)])} />
              </div>
            </Group>
            <Group title="Зависимости формул">
              <div className="rb-col">
                <Btn title="Стрелки от ячеек, на которые ссылается формула" onClick={api.precedents}><T>→▣</T> Влияющие ячейки</Btn>
                <Btn title="Стрелки к формулам, которые ссылаются на ячейку" onClick={api.dependents}><T>▣→</T> Зависимые ячейки</Btn>
              </div>
              <div className="rb-col">
                <Btn title="Убрать стрелки" onClick={api.clearArrows}><T>✕</T> Убрать стрелки</Btn>
                <Btn title="Показать формулы (Ctrl+`)" active={api.showFormulas} onClick={api.toggleFormulas}><T>ƒx</T> Показать формулы</Btn>
              </div>
              <Btn big icon="evaluate" label="Вычислить формулу" onClick={api.evaluateFormula} title="Вычислить формулу по шагам" />
            </Group>
            <Group title="Вычисление">
              <MenuButton title="Когда пересчитывать формулы" label="Параметры вычислений" items={[
                ['Автоматически', () => api.setCalcMode('auto'), api.calcMode === 'auto'],
                ['Вручную', () => api.setCalcMode('manual'), api.calcMode === 'manual'],
              ]} />
              <Btn big icon="recalc" label="Пересчёт" onClick={api.recalc} title="Пересчитать книгу (F9)" />
            </Group>
            <Group title="Определённые имена">
              <Btn big icon="names" label="Диспетчер имён" onClick={api.names} title="Именованные диапазоны и константы книги" />
            </Group>
          </>
        )}
        {tab === 'data' && (
          <>
            <Group title="Сортировка и фильтр">
              <div className="rb-col">
                <Btn title="Сортировка по возрастанию" onClick={() => change((w) => sortRange(w, si, rg, sel.c, true))}><T>А→Я</T></Btn>
                <Btn title="Сортировка по убыванию" onClick={() => change((w) => sortRange(w, si, rg, sel.c, false))}><T>Я→А</T></Btn>
              </div>
              <Btn big icon="sort" label="Сортировка" onClick={api.sortDialog} title="Сортировка по нескольким столбцам" />
              <Btn big icon="filter" label="Фильтр" active={api.filterOn} onClick={api.toggleFilter} title="Автофильтр: кнопки в заголовках столбцов (Ctrl+Shift+L)" />
              <div className="rb-col">
                <Btn title="Снять условия фильтра во всех столбцах" onClick={api.clearFilter} disabled={!api.filterOn}><T>✕</T> Очистить</Btn>
                <Btn title="Применить фильтр заново (после правок данных)" onClick={api.reapplyFilter} disabled={!api.filterOn}><T>↻</T> Повторить</Btn>
              </div>
            </Group>
            <Group title="Структура">
              <MenuButton icon="group" title="Группировать выделенные строки или столбцы" label="Группировать" items={[
                ['Строки', () => group('row', 1)],
                ['Столбцы', () => group('col', 1)],
              ]} />
              <MenuButton icon="ungroup" title="Разгруппировать" label="Разгруппировать" items={[
                ['Строки', () => group('row', -1)],
                ['Столбцы', () => group('col', -1)],
              ]} />
              <div className="rb-col">
                <Btn title="Показать детали группы" onClick={() => detail(true)}><T>+</T> Показать детали</Btn>
                <Btn title="Скрыть детали группы" onClick={() => detail(false)}><T>−</T> Скрыть детали</Btn>
              </div>
            </Group>
          </>
        )}
        {tab === 'review' && (
          <>
            <Group title="Защита">
              <Btn big icon="lock" label={prot ? 'Снять защиту листа' : 'Защитить лист'} onClick={api.protect}
                title={prot ? 'Снять защиту листа (может потребоваться пароль)' : 'Запретить изменение заблокированных ячеек'} />
              <div className="rb-col">
                <Btn title="Блокировать выделенные ячейки (действует при защите листа)" active={!s.unlocked} onClick={() => style({ unlocked: !s.unlocked })}><T>▣</T> Блокировать ячейку</Btn>
                <Btn title="Скрывать формулы выделенных ячеек при защите листа" active={!!s.formulaHidden} onClick={() => style({ formulaHidden: !s.formulaHidden })}><T>ƒ</T> Скрыть формулы</Btn>
              </div>
            </Group>
          </>
        )}
        {tab === 'view' && (
          <>
            <Group title="Окно">
              <MenuButton icon="freeze" title="Закрепить области" label="Закрепить области" items={[
                [frozen ? 'Снять закрепление областей' : 'Закрепить области', () => change((w) => setFreeze(w, si, frozen ? null : { rows: sel.r, cols: sel.c }))],
                ['Закрепить верхнюю строку', () => change((w) => setFreeze(w, si, { rows: 1, cols: 0 }))],
                ['Закрепить первый столбец', () => change((w) => setFreeze(w, si, { rows: 0, cols: 1 }))],
              ]} />
            </Group>
          </>
        )}
      </div>
    </div>
  );
}
