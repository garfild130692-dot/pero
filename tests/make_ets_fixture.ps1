# Ответы Excel (COM) по FORECAST.ETS* на четырёх рядах → tests/fixtures/ets_excel.json (для tests/ets.test.js).
# Ряды: тренд + сезон 4 + шум; тренд + шум; сезон 12 + шум; квадратичный рост. Книга не сохраняется.
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  # ряды: A — время 1..N, B..E — значения
  $N = 36
  for ($i = 1; $i -le $N; $i++) {
    $ws.Cells.Item($i, 1).Value2 = [double]$i
    $s4 = @(5, -3, 2, -4)[($i - 1) % 4]
    $s12 = [math]::Round(10 * [math]::Sin(2 * [math]::PI * $i / 12), 3)
    $noise = [math]::Round(3 * [math]::Sin($i * 1.7) + 2 * [math]::Cos($i * 0.9), 3)
    $ws.Cells.Item($i, 2).Value2 = [double](100 + 2 * $i + $s4 + $noise)
    $ws.Cells.Item($i, 3).Value2 = [double](50 + 0.5 * $i + $noise)
    $ws.Cells.Item($i, 4).Value2 = [double](200 + $s12 + $noise)
    $ws.Cells.Item($i, 5).Value2 = [double][math]::Round(20 + $i * $i * 0.1 + $noise, 3)
  }
  $out = @{}
  foreach ($col in 'B', 'C', 'D', 'E') {
    $r = "$($col)1:$($col)$N"
    $f = @()
    $f += "=FORECAST.ETS.SEASONALITY($r,A1:A$N)"
    foreach ($k in 1..8) { $f += "=FORECAST.ETS.STAT($r,A1:A$N,$k)" }
    foreach ($t in 37, 38, 40, 45, 36.5) { $f += "=FORECAST.ETS($t,$r,A1:A$N)" }
    $f += "=FORECAST.ETS.CONFINT(37,$r,A1:A$N)"
    $f += "=FORECAST.ETS.CONFINT(40,$r,A1:A$N,0.9)"
    $f += "=FORECAST.ETS(37,$r,A1:A$N,0)"
    foreach ($k in 1..3) { $f += "=FORECAST.ETS.STAT($r,A1:A$N,$k,0)" }
    foreach ($k in 1..3) { $f += "=FORECAST.ETS.STAT($r,A1:A$N,$k,4)" }
    $f += "=FORECAST.ETS(37,$r,A1:A$N,4)"
    $vals = @()
    foreach ($x in $f) { $ws.Range('H1').Formula2 = $x; $vals += ,@($x, $ws.Range('H1').Value2) }
    $out[$col] = $vals
  }
  $data = @()
  for ($i = 1; $i -le $N; $i++) { $data += ,@($ws.Cells.Item($i, 2).Value2, $ws.Cells.Item($i, 3).Value2, $ws.Cells.Item($i, 4).Value2, $ws.Cells.Item($i, 5).Value2) }
  $out['data'] = $data
  $out | ConvertTo-Json -Depth 5 | Out-File -Encoding utf8 tests/fixtures/ets_excel.json
  $wb.Close($false)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
