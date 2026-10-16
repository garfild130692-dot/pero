# Объекты книги .xls глазами Excel (COM): восстанавливалась ли, и по листам — диаграммы, рисунки/фигуры, примечания,
# условное форматирование, гиперссылки, проверка данных, таблицы, автофильтр. Для сверки переноса записей (этап 19.3).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/xls_objects_check.ps1 <файл>
param([string]$Path)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$full = [IO.Path]::GetFullPath($Path)
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
  $tmp = [IO.Path]::GetTempPath()
  $before = @(Get-ChildItem $tmp -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
  $wb = $x.Workbooks.Open($full, 0, $true)
  $after = @(Get-ChildItem $tmp -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.FullName })
  $sheets = @()
  foreach ($ws in $wb.Worksheets) {
    $dv = 0
    try { $dv = $ws.Cells.SpecialCells(-4174).Count } catch { }
    $notes = @(); foreach ($c in $ws.Comments) { $notes += "$($c.Parent.Address(0,0))=$($c.Text())" }
    $links = @(); foreach ($h in $ws.Hyperlinks) { $links += "$($h.Range.Address(0,0))=$($h.Address)$($h.SubAddress)" }
    $charts = @(); foreach ($co in $ws.ChartObjects()) { $s = @(); foreach ($se in $co.Chart.SeriesCollection()) { $s += $se.Formula }; $charts += ($s -join ';') }
    $sheets += @{ name = $ws.Name; charts = $charts; shapes = $ws.Shapes.Count; notes = $notes; cf = $ws.Cells.FormatConditions.Count;
      links = $links; dv = $dv; tables = $ws.ListObjects.Count; filter = [bool]$ws.AutoFilterMode; header = $ws.PageSetup.CenterHeader }
  }
  $names = @(); foreach ($n in $wb.Names) { $names += "$($n.Name)=$($n.RefersTo)" }
  $wb.Close($false)
  @{ repaired = ($after.Count -gt 0); sheets = $sheets; names = $names } | ConvertTo-Json -Depth 5
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
