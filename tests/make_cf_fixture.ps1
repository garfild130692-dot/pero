# tests/fixtures/cf.xlsx — создаётся самим Excel (COM): условное форматирование разных видов,
# и tests/fixtures/cf_excel.json — что Excel показывает в каждой ячейке (DisplayFormat: заливка, цвет и жирность текста).
# Автор — нейтральный «Перо Тест» (UserName возвращается после генерации). Файл в UTF-8 с BOM (PowerShell 5.1).
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$oldName = $xl.UserName
$xl.UserName = 'Перо Тест'
function Bgr($rgb) { $r = [Convert]::ToInt32($rgb.Substring(0, 2), 16); $g = [Convert]::ToInt32($rgb.Substring(2, 2), 16); $b = [Convert]::ToInt32($rgb.Substring(4, 2), 16); return $r + 256 * $g + 65536 * $b }
function Hex($bgr) { $v = [int]$bgr; return ('#{0:x2}{1:x2}{2:x2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255)) }
try {
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  $a = @(5, 15, 55, 12, 80, 20, 50, 3, 99, 18, 51, 10)
  $b = @('касса', 'банк', 'склад', 'банк', 'офис', 'Касса', 'кадры', 'склад', 'цех', 'офис', 'банк', 'сайт')
  $c = @(7, 2, 9, 4, 9, 1, 6, 3, 8, 5, 10, 2)
  for ($i = 0; $i -lt 12; $i++) {
    $ws.Cells.Item($i + 1, 1).Value2 = [double]$a[$i]
    $ws.Cells.Item($i + 1, 2).Value2 = [string]$b[$i]
    $ws.Cells.Item($i + 1, 3).Value2 = [double]$c[$i]
    $ws.Cells.Item($i + 1, 4).Value2 = [double]($c[$i] * 10)
    $ws.Cells.Item($i + 1, 5).Value2 = [double]$i
  }
  # A: > 50 — красная заливка; между 10 и 20 — зелёная заливка, тёмно-зелёный текст
  $f = $ws.Range('A1:A12').FormatConditions.Add(1, 5, '=50'); $f.Interior.Color = (Bgr 'FFC7CE'); $f.Font.Color = (Bgr '9C0006')
  $f = $ws.Range('A1:A12').FormatConditions.Add(1, 1, '=10', '=20'); $f.Interior.Color = (Bgr 'C6EFCE'); $f.Font.Color = (Bgr '006100')
  # B: текст содержит «ка»; повторяющиеся — жёлтая заливка
  $f = $ws.Range('B1:B12').FormatConditions.Add(9, $null, $null, $null, 'ка', 0); $f.Font.Bold = $true
  $f = $ws.Range('B1:B12').FormatConditions.AddUniqueValues(); $f.DupeUnique = 1; $f.Interior.Color = (Bgr 'FFEB9C')
  # C: первые 3 — голубая заливка; выше среднего — красный текст
  $f = $ws.Range('C1:C12').FormatConditions.AddTop10(); $f.TopBottom = 1; $f.Rank = 3; $f.Interior.Color = (Bgr 'BDD7EE')
  $f = $ws.Range('C1:C12').FormatConditions.AddAboveAverage(); $f.AboveBelow = 0; $f.Font.Color = (Bgr 'FF0000')
  # D: трёхцветная шкала (красный — жёлтый — зелёный)
  [void]$ws.Range('D1:D12').FormatConditions.AddColorScale(3)
  # E: формула — строка, где A > C*10 (относительная ссылка на строку)
  $f = $ws.Range('E1:E12').FormatConditions.Add(2, $null, '=$A1>$C1*10'); $f.Interior.Color = (Bgr 'E2EFDA'); $f.Font.Bold = $true
  $wb.SaveAs([IO.Path]::GetFullPath('tests/fixtures/cf.xlsx'), 51)
  $out = @{}
  for ($r = 1; $r -le 12; $r++) {
    for ($cc = 1; $cc -le 5; $cc++) {
      $d = $ws.Cells.Item($r, $cc).DisplayFormat
      $fill = if ($d.Interior.ColorIndex -eq -4142) { $null } else { Hex $d.Interior.Color }
      $out["$($r - 1),$($cc - 1)"] = @{ fill = $fill; color = (Hex $d.Font.Color); bold = [bool]$d.Font.Bold }
    }
  }
  $out | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 tests/fixtures/cf_excel.json
  $wb.Close($false)
} finally {
  $xl.UserName = $oldName
  $xl.Quit()
}
