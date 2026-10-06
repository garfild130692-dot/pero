# Генерирует tests/fixtures/book.xlsx средствами Excel (COM): числа, даты, проценты, денежный формат, формулы
# (в т.ч. заполнение вниз — Excel пишет «общие» формулы), стили, объединение, ширины, второй лист с русским именем.
# Автор — нейтральный «Перо Тест» (UserName возвращается после генерации). Файл в UTF-8 с BOM (PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_xlsx_fixture.ps1 tests/fixtures/book.xlsx
param([string]$Out = 'tests/fixtures/book.xlsx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
# Excel через COM разбирает коды форматов по культуре вызывающего потока (в русской Windows «0.00» не принимается,
# а «dd.mm.yyyy» понимается как текст). Вызов через InvokeMember с культурой en-US — коды всегда английские.
function SetFmt($range, $code) {
  [void]$range.GetType().InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $range, @($code), [Globalization.CultureInfo]::GetCultureInfo('en-US'))
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
  $s.Name = 'Продажи'
  $s.Range('A1').Value2 = 'Товар'; $s.Range('B1').Value2 = 'Цена'; $s.Range('C1').Value2 = 'Кол-во'; $s.Range('D1').Value2 = 'Сумма'; $s.Range('E1').Value2 = 'Дата'
  $data = @(@('Ручка', 45.5, 10), @('Тетрадь', 120, 3), @('Карандаш', 15.25, 20), @('Линейка', 60, 0))
  for ($i = 0; $i -lt 4; $i++) {
    $r = $i + 2
    $s.Range("A$r").Value2 = [string]$data[$i][0]; $s.Range("B$r").Value2 = [double]$data[$i][1]; $s.Range("C$r").Value2 = [double]$data[$i][2]
  }
  $s.Range('D2').Formula = '=B2*C2'
  $s.Range('D2:D5').FillDown() | Out-Null           # общие формулы
  $s.Range('E2').Formula = '=DATE(2026,10,6)'
  $s.Range('E3').Value2 = 46302
  $s.Range('A6').Value2 = 'Итого'
  $s.Range('D6').Formula = '=SUM(D2:D5)'
  $s.Range('B7').Formula = '=AVERAGE(B2:B5)'
  $s.Range('C7').Formula = '=COUNTIF(C2:C5,">0")'
  $s.Range('D7').Formula = '=IF(D6>1000,"много","мало")'
  $s.Range('E7').Formula = '=D6/0'
  $s.Range('A8').Value2 = 0.125
  SetFmt $s.Range('A8') '0.0%'
  SetFmt $s.Range('B2:B6') '#,##0.00 "₽"'
  SetFmt $s.Range('D2:D6') '#,##0.00'
  SetFmt $s.Range('E2:E3') 'dd.mm.yyyy'
  $s.Range('A1:E1').Font.Bold = $true
  $s.Range('A1:E1').Interior.Color = 0xE9EFDC        # BGR → #DCEFE9
  $s.Range('A6:D6').Font.Italic = $true
  $s.Range('A1:E1').HorizontalAlignment = -4108      # по центру
  $s.Range('A6:C6').Merge() | Out-Null
  $s.Range('A1:E6').Borders.LineStyle = 1
  $s.Columns.Item(1).ColumnWidth = 20
  $s.Rows.Item(1).RowHeight = 24
  $s2 = $wb.Worksheets.Add([Type]::Missing, $s)
  $s2.Name = 'Свод'
  $s2.Range('A1').Value2 = 'Всего продаж'
  $s2.Range('B1').Formula = "='Продажи'!D6*2"
  $s.Activate()
  $wb.SaveAs($path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($false)
  "ok: $path"
} finally {
  $x.UserName = $oldName
  $x.Quit()
}
