# Сводка книги глазами Excel (COM) — для сверки «до/после сохранения в Пере»: открывается ли без восстановления,
# листы, диаграммы, условное форматирование, проверка, имена, примечания, ссылки, таблицы, сводные, закрепление,
# скрытые/сгруппированные строки, печать, защита, значения выбранных ячеек.
# Запуск: powershell -ExecutionPolicy Bypass -File tests/xlsx_check.ps1 <файл.xlsx> [A1,B3,...]
param([string]$Path, [string]$Cells = '')
$ErrorActionPreference = 'Stop'
$full = [IO.Path]::GetFullPath($Path)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
  # восстановление повреждённого файла Excel записывает журнал error*.xml во временную папку
  $before = Get-Date
  $wb = $x.Workbooks.Open($full, 0, $true)
  $repair = @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $before })
  $out = [ordered]@{ name = $wb.Name; repaired = ($repair.Count -gt 0); sheets = @() }
  $names = @()
  foreach ($n in $wb.Names) { $names += "$($n.Name)=$($n.RefersTo)" }
  $out.names = $names
  foreach ($s in $wb.Worksheets) {
    $s.Activate()
    $w = $x.ActiveWindow
    $hidden = @()
    for ($r = 1; $r -le 30; $r++) { if ($s.Rows.Item($r).Hidden) { $hidden += $r } }
    $hcols = @()
    for ($c = 1; $c -le 20; $c++) { if ($s.Columns.Item($c).Hidden) { $hcols += $c } }
    $outline = @()
    for ($r = 1; $r -le 30; $r++) { $l = $s.Rows.Item($r).OutlineLevel; if ($l -gt 1) { $outline += "${r}:${l}" } }
    $vals = @()
    if ($Cells) { foreach ($a in $Cells.Split(',')) { try { $vals += "$a=$($s.Range($a).Text)" } catch { } } }
    $out.sheets += [ordered]@{
      name = $s.Name
      used = $s.UsedRange.Address(0, 0)
      charts = $s.ChartObjects().Count
      cf = $s.Cells.FormatConditions.Count
      validation = $(try { $s.Cells.SpecialCells(-4174).Address(0, 0) } catch { '' })
      comments = $s.Comments.Count
      links = $s.Hyperlinks.Count
      tables = $s.ListObjects.Count
      pivots = $s.PivotTables().Count
      freeze = "$($w.FreezePanes) $($w.SplitRow) $($w.SplitColumn)"
      hiddenRows = ($hidden -join ' ')
      hiddenCols = ($hcols -join ' ')
      outline = ($outline -join ' ')
      printArea = $s.PageSetup.PrintArea
      printTitles = $s.PageSetup.PrintTitleRows
      orientation = $s.PageSetup.Orientation
      fitWide = "$($s.PageSetup.FitToPagesWide)"
      protected = $s.ProtectContents
      values = ($vals -join ' | ')
    }
  }
  $wb.Close($false)
  $out | ConvertTo-Json -Depth 4
} finally {
  $x.Quit()
}
