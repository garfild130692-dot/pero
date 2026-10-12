# Книга глазами Excel (COM) для сверки записи .xls (этап 19): открылась ли без восстановления, и по каждой непустой
# ячейке — формула (английская запись) и значение после полного пересчёта. Вывод — JSON.
# Запуск: powershell -ExecutionPolicy Bypass -File tests/xls_excel_dump.ps1 <файл> > out.json
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
  $x.CalculateFull()
  $sheets = @()
  foreach ($ws in $wb.Worksheets) {
    $cells = @{}
    $ur = $ws.UsedRange
    foreach ($c in $ur.Cells) {
      $f = $c.Formula
      $v = $c.Value2
      if ($null -eq $v -and -not $c.HasFormula) { continue }
      $key = "$($c.Row - 1),$($c.Column - 1)"
      $val = if ($v -is [int]) { [double]$v } else { $v }
      $cells[$key] = @{ f = $(if ($c.HasFormula) { $f } else { $null }); v = $val; t = $c.Text }
    }
    $sheets += @{ name = $ws.Name; cells = $cells }
  }
  $wb.Close($false)
  @{ repaired = ($after.Count -gt 0); sheets = $sheets } | ConvertTo-Json -Depth 6 -Compress
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
