# Генерирует tests/fixtures/preserve.xlsx средствами Excel (COM) — всё, что Перо не редактирует, но обязано сохранить:
# диаграмма, условное форматирование, проверка данных, именованные диапазоны, примечание, гиперссылка, форматированный
# текст в ячейке, «умная» таблица, сводная таблица, закрепление областей, скрытые и сгруппированные строки/столбцы,
# область печати и сквозные строки, параметры страницы, защита листа с незаблокированными ячейками.
# Автор — нейтральный «Перо Тест» (UserName возвращается после генерации). Файл в UTF-8 с BOM (PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_xlsx_preserve_fixture.ps1 tests/fixtures/preserve.xlsx
param([string]$Out = 'tests/fixtures/preserve.xlsx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$en = [Globalization.CultureInfo]::GetCultureInfo('en-US')
function SetProp($obj, $name, $value) {
  [void]$obj.GetType().InvokeMember($name, [Reflection.BindingFlags]::SetProperty, $null, $obj, @($value), $en)
}
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$oldName = $x.UserName
$x.UserName = 'Перо Тест'
try {
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $s = $wb.Worksheets.Item(1)
  $s.Name = 'Данные'
  $s.Range('A1').Value2 = 'Месяц'; $s.Range('B1').Value2 = 'План'; $s.Range('C1').Value2 = 'Факт'; $s.Range('D1').Value2 = 'Отклонение'
  $months = @('Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь')
  for ($i = 0; $i -lt 12; $i++) {
    $r = $i + 2
    $s.Range("A$r").Value2 = $months[$i]
    $s.Range("B$r").Value2 = [double](100 + 10 * $i)
    $s.Range("C$r").Value2 = [double](95 + 13 * (($i * 7) % 5))
  }
  $s.Range('D2').Formula = '=C2-B2'
  $s.Range('D2:D13').FillDown() | Out-Null
  $s.Range('A14').Value2 = 'Итого'
  $s.Range('B14').Formula = '=SUM(B2:B13)'
  $s.Range('C14').Formula = '=SUM(C2:C13)'
  $s.Range('D14').Formula = '=SUM(D2:D13)'
  $s.Range('A1:D1').Font.Bold = $true
  $s.Columns.Item(1).ColumnWidth = 14
  $s.Columns.Item(4).ColumnWidth = 12
  # условное форматирование: отрицательные отклонения — красным; шкала цветов на факте
  $fc = $s.Range('D2:D13').FormatConditions.Add(1, 6, '=0')   # xlCellValue, xlLess
  $fc.Font.Color = 0x0000C0
  [void]$s.Range('C2:C13').FormatConditions.AddColorScale(3)
  # проверка данных: план — целое от 0 до 1000
  $s.Range('B2:B13').Validation.Add(1, 1, 1, '0', '1000')     # xlValidateWholeNumber, xlValidAlertStop, xlBetween
  # именованные диапазоны
  [void]$wb.Names.Add('План', '=Данные!$B$2:$B$13')
  [void]$wb.Names.Add('Ставка', '=1/5')
  # примечание и гиперссылка
  [void]$s.Range('A1').AddComment('Месяцы года')
  [void]$s.Hyperlinks.Add($s.Range('F1'), 'https://example.com/', '', '', 'Справка')
  # форматированный текст в ячейке: часть жирным, часть красным
  $s.Range('F3').Value2 = 'Важно: проверить'
  $s.Range('F3').Characters(1, 6).Font.Bold = $true
  $s.Range('F4').Value2 = 'Красное слово'
  $s.Range('F4').Characters(1, 7).Font.Color = 0x0000FF
  # стиль столбца: заливка всего столбца H
  $s.Columns.Item(8).Interior.Color = 0xCCFFFF
  # скрытые столбец и строка, группировка строк 16–18 (свёрнута) и столбцов J–K
  $s.Range('G1').Value2 = 'скрыт'
  $s.Columns.Item(7).Hidden = $true
  $s.Range('A15').Value2 = 'скрыта'
  $s.Rows.Item(15).Hidden = $true
  for ($r = 16; $r -le 18; $r++) { $s.Range("A$r").Value2 = "деталь $r" }
  [void]$s.Rows('16:18').Group()
  $s.Rows('16:18').Hidden = $true
  $s.Range('J1').Value2 = 'гр1'; $s.Range('K1').Value2 = 'гр2'
  [void]$s.Columns('J:K').Group()
  # диаграмма по плану и факту
  $ch = $s.Shapes.AddChart2(201, 51, 400, 260, 360, 220)      # xlColumnClustered
  $ch.Chart.SetSourceData($s.Range('A1:C13'))
  # закрепление: строка заголовка и первый столбец
  $s.Activate()
  $x.ActiveWindow.FreezePanes = $false
  $x.ActiveWindow.SplitRow = 1
  $x.ActiveWindow.SplitColumn = 1
  $x.ActiveWindow.FreezePanes = $true
  # печать: область, сквозная строка, альбомная, в ширину одной страницы, поля
  $s.PageSetup.PrintArea = '$A$1:$D$14'
  $s.PageSetup.PrintTitleRows = '$1:$1'
  $s.PageSetup.Orientation = 2
  $s.PageSetup.Zoom = $false
  $s.PageSetup.FitToPagesWide = 1
  $s.PageSetup.FitToPagesTall = $false
  $s.PageSetup.LeftMargin = $x.CentimetersToPoints(1.5)
  $s.PageSetup.RightMargin = $x.CentimetersToPoints(1.5)

  # второй лист: «умная» таблица и защита листа (B2:B4 не заблокированы)
  $s2 = $wb.Worksheets.Add([Type]::Missing, $s)
  $s2.Name = 'Справочник'
  $s2.Range('A1').Value2 = 'Код'; $s2.Range('B1').Value2 = 'Название'
  $s2.Range('A2').Value2 = 1; $s2.Range('B2').Value2 = 'Альфа'
  $s2.Range('A3').Value2 = 2; $s2.Range('B3').Value2 = 'Бета'
  $s2.Range('A4').Value2 = 3; $s2.Range('B4').Value2 = 'Гамма'
  $lo = $s2.ListObjects.Add(1, $s2.Range('A1:B4'), $null, 1)  # xlSrcRange, xlYes
  $lo.Name = 'Коды'
  $s2.Range('B2:B4').Locked = $false
  $s2.Protect()

  # третий лист: сводная таблица по данным
  $s3 = $wb.Worksheets.Add([Type]::Missing, $s2)
  $s3.Name = 'Свод'
  $pc = $wb.PivotCaches().Create(1, 'Данные!R1C1:R13C3')        # xlDatabase
  $pt = $pc.CreatePivotTable('Свод!R3C1', 'Сводная1')
  $pt.PivotFields('Месяц').Orientation = 1                       # xlRowField
  [void]$pt.AddDataField($pt.PivotFields('Факт'))

  $s.Activate()
  $s.Range('A1').Select()
  $wb.SaveAs($path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($false)
  "ok: $path"
} finally {
  $x.UserName = $oldName
  $x.Quit()
}
