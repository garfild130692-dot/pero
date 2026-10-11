# tests/fixtures/pages_*.xlsx и tests/fixtures/pages_excel.json — разбиение листа на страницы самим Excel (COM):
# для каждой настройки страницы (ориентация, поля, масштаб, «вписать», сквозные строки, ручной разрыв) книга
# сохраняется, а номера строк/столбцов, с которых начинаются страницы (HPageBreaks/VPageBreaks), записываются.
# Автор — нейтральный «Перо Тест» (UserName возвращается после генерации). Файл в UTF-8 с BOM (PowerShell 5.1).
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$oldName = $xl.UserName
$xl.UserName = 'Перо Тест'
$configs = @(
  @{ name = 'a4'; },
  @{ name = 'land'; orientation = 2 },
  @{ name = 'scale70'; zoom = 70 },
  @{ name = 'fitwide'; fitWide = 1 },
  @{ name = 'narrow'; margins = 0.25 },
  @{ name = 'titles'; titles = '$1:$2' },
  @{ name = 'manual'; manualRow = 30; manualCol = 5 }
)
$result = @{}
try {
  foreach ($cfg in $configs) {
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    for ($c = 1; $c -le 30; $c++) {
      $ws.Columns.Item($c).ColumnWidth = [double](6 + ($c * 7) % 17)
      $ws.Cells.Item(1, $c).Value2 = "Столбец $c"
    }
    for ($r = 2; $r -le 160; $r++) {
      if ($r % 7 -eq 0) { $ws.Rows.Item($r).RowHeight = [double](15 + ($r % 5) * 6) }
      $ws.Cells.Item($r, 1).Value2 = [double]$r
      $ws.Cells.Item($r, 30).Value2 = [double]($r * 2)
    }
    $ps = $ws.PageSetup
    $xl.PrintCommunication = $false
    if ($cfg.orientation) { $ps.Orientation = $cfg.orientation }
    if ($cfg.zoom) { $ps.Zoom = $cfg.zoom }
    if ($cfg.fitWide) { $ps.Zoom = $false; $ps.FitToPagesWide = $cfg.fitWide; $ps.FitToPagesTall = $false }
    if ($cfg.margins) { $in = $xl.InchesToPoints($cfg.margins); $ps.LeftMargin = $in; $ps.RightMargin = $in; $ps.TopMargin = $in; $ps.BottomMargin = $in }
    if ($cfg.titles) { $ps.PrintTitleRows = $cfg.titles }
    $xl.PrintCommunication = $true
    if ($cfg.manualRow) { [void]$ws.HPageBreaks.Add($ws.Cells.Item($cfg.manualRow, 1)); [void]$ws.VPageBreaks.Add($ws.Cells.Item(1, $cfg.manualCol)) }
    $ws.Activate(); $xl.ActiveWindow.View = 2  # страничный режим — разрывы пересчитываются
    $rows = @(); for ($i = 1; $i -le $ws.HPageBreaks.Count; $i++) { $rows += $ws.HPageBreaks.Item($i).Location.Row }
    $cols = @(); for ($i = 1; $i -le $ws.VPageBreaks.Count; $i++) { $cols += $ws.VPageBreaks.Item($i).Location.Column }
    $xl.ActiveWindow.View = 1
    $wb.SaveAs([IO.Path]::GetFullPath("tests/fixtures/pages_$($cfg.name).xlsx"), 51)
    $result[$cfg.name] = @{ rows = $rows; cols = $cols }
    $wb.Close($false)
  }
  $result | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 tests/fixtures/pages_excel.json
} finally {
  $xl.UserName = $oldName
  $xl.Quit()
}
