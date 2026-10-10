# Генерирует tests/fixtures/formulas.xlsx средствами Excel (COM) — эталон для сверки формул Перо с Excel.
# Формулы — из tests/formula_cases.json (английская запись); лист «Формулы»: A — группа, B — текст формулы, C — формула,
# E — её русская запись (Formula2Local)
# (Formula2 — с динамическими массивами), значения в файле посчитаны самим Excel. Лист «Данные» и листы Кв1–Кв3 — исходные
# данные, имена Ставка и Суммы. Автор — нейтральный «Перо Тест». Файл в UTF-8 с BOM (PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_formulas_fixture.ps1
param([string]$Cases = 'tests/formula_cases.json', [string]$Out = 'tests/fixtures/formulas.xlsx')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($Out)
$json = Get-Content -Raw -Encoding UTF8 $Cases | ConvertFrom-Json
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$oldName = $x.UserName
$x.UserName = 'Перо Тест'
try {
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $d = $wb.Worksheets.Item(1)
  $d.Name = 'Данные'
  $head = @('Имя', 'Отдел', 'Сумма', 'Дата', 'Кол', 'Флаг', 'Смесь', 'Текст')
  for ($c = 0; $c -lt 8; $c++) { $d.Cells.Item(1, $c + 1).Value2 = $head[$c] }
  $names = @('Иван', 'Пётр', 'Мария', 'Анна', 'Олег', 'Иван', 'Мария', 'Пётр', 'Анна', 'Ольга')
  $deps = @('Продажи', 'Склад', 'Продажи', 'Бухгалтерия', 'Склад', 'Продажи', 'Бухгалтерия', 'Склад', 'Продажи', 'Склад')
  $sums = @(1500, 230.5, 980, 4200, 0, 15.25, 3100, -200, 760, 1200)
  $dates = @(@(2026, 1, 15), @(2026, 2, 28), @(2025, 12, 31), @(2026, 3, 1), @(2024, 2, 29), @(2026, 6, 30), @(2026, 10, 10), @(2023, 1, 1), @(2026, 12, 25), @(2026, 5, 9))
  $qty = @(3, 1, 7, 2, 5, 8, 1, 4, 6, 2)
  $texts = @('  Привет   мир ', 'abc', 'ABC-123', 'Иванов Иван Иванович', 'a,b,c', 'Москва', 'Сумма: 100', 'x', 'Test', 'тест')
  for ($i = 0; $i -lt 10; $i++) {
    $r = $i + 2
    $d.Cells.Item($r, 1).Value2 = $names[$i]
    $d.Cells.Item($r, 2).Value2 = $deps[$i]
    $d.Cells.Item($r, 3).Value2 = [double]$sums[$i]
    $d.Cells.Item($r, 4).Value2 = [double]([DateTime]::new([int]$dates[$i][0], [int]$dates[$i][1], [int]$dates[$i][2]).ToOADate())
    $d.Cells.Item($r, 5).Value2 = [double]$qty[$i]
    $d.Cells.Item($r, 6).Formula = $(if ($i % 2 -eq 0) { 'TRUE' } else { 'FALSE' })  # логическая константа
    $d.Cells.Item($r, 8).NumberFormat = '@'
    $d.Cells.Item($r, 8).Value2 = $texts[$i]
  }
  # код формата — в английской записи (иначе русский Excel поймёт «dd.mm.yyyy» как текст)
  $rng = $d.Range('D2:D11')
  [void]$rng.GetType().InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $rng, @('dd.mm.yyyy'), [Globalization.CultureInfo]::GetCultureInfo('en-US'))
  # «Смесь»: число, число-текст, пусто, логическое, #Н/Д, дробь, пустая строка (формула), отрицательное, текст, ноль
  $d.Range('G2').Value2 = [double]5
  $d.Range('G3').NumberFormat = '@'
  $d.Range('G3').Value2 = '12'
  $d.Range('G5').Formula = 'TRUE'
  $d.Range('G6').Formula = '=NA()'
  $d.Range('G7').Value2 = [double]3.5
  $d.Range('G8').Formula = '=""'
  $d.Range('G9').Value2 = [double]-2
  $d.Range('G10').Value2 = 'текст'
  $d.Range('G11').Value2 = [double]0
  # диапазоны критериев для функций баз данных: J1:K2 — «И», L1:L3 — «ИЛИ»
  $d.Range('J1').Value2 = 'Отдел'; $d.Range('J2').Value2 = 'Склад'; $d.Range('K1').Value2 = 'Сумма'; $d.Range('K2').Value2 = '>500'
  $d.Range('L1').Value2 = 'Отдел'; $d.Range('L2').Value2 = 'Продажи'; $d.Range('L3').Value2 = 'Склад'
  # трёхмерные ссылки
  $prev = $d
  foreach ($q in 1..3) {
    $s = $wb.Worksheets.Add([Type]::Missing, $prev)
    $s.Name = "Кв$q"
    $s.Range('A1').Value2 = [double](10 * $q)
    $s.Range('B2').Value2 = [double]$q
    $prev = $s
  }
  [void]$wb.Names.Add('Ставка', '=1/5')
  [void]$wb.Names.Add('Суммы', '=Данные!$C$2:$C$11')

  $f = $wb.Worksheets.Add([Type]::Missing, $prev)
  $f.Name = 'Формулы'
  $f.Range('A1').Value2 = 'Группа'; $f.Range('B1').Value2 = 'Формула'; $f.Range('C1').Value2 = 'Excel'
  $f.Columns.Item(2).NumberFormat = '@'
  $row = 2
  foreach ($g in $json.PSObject.Properties) {
    if ($g.Name -eq '_') { continue }
    foreach ($text in $g.Value) {
      $f.Cells.Item($row, 1).Value2 = $g.Name
      $f.Cells.Item($row, 2).Value2 = $text
      try {
        $f.Cells.Item($row, 3).Formula2 = "=$text"
        # русская запись той же формулы — эталон для русских имён функций и разделителей
        $f.Cells.Item($row, 5).NumberFormat = '@'
        $f.Cells.Item($row, 5).Value2 = $f.Cells.Item($row, 3).Formula2Local
      }
      catch { $f.Cells.Item($row, 4).Value2 = "Excel не принял формулу" }
      $row++
    }
  }
  $x.CalculateFull()
  $d.Activate()
  $wb.SaveAs($path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
  $wb.Close($false)
  "ok: $path, формул: $($row - 2)"
} finally {
  $x.UserName = $oldName
  $x.Quit()
}
