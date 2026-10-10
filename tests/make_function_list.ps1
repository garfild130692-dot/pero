# Список функций Excel с русскими именами — по самому Excel (COM): для каждого имени из tests/excel_functions.txt
# в ячейку пишется =ИМЯ() (Formula2), Excel отвечает русской записью (Formula2Local) — известна ли функция, видно по тому,
# перевёл ли он имя или оставил #ИМЯ?. Результат — tests/fixtures/excel_functions.json { EN: русское имя | null }.
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_function_list.ps1
param([string]$List = 'tests/excel_functions.txt', [string]$Out = 'tests/fixtures/excel_functions.json')
$ErrorActionPreference = 'Stop'
$names = (Get-Content -Encoding UTF8 $List | Where-Object { $_ -notmatch '^#' }) -join ' ' -split '\s+' | Where-Object { $_ }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$result = [ordered]@{}
try {
  $wb = $x.Workbooks.Add()
  $s = $wb.Worksheets.Item(1)
  $cell = $s.Range('A1')
  foreach ($n in $names) {
    $ru = $null
    # Excel проверяет число и вид аргументов при вводе — подбираем 0–8 чисел, затем 1–8 диапазонов (СУММЕСЛИ, СМЕЩ, БДСУММ…)
    $tries = @(0..8 | ForEach-Object { (@('1') * $_) -join ',' }) + @(1..8 | ForEach-Object { (@('A1:A2') * $_) -join ',' }) + @('x,1,x')
    foreach ($argsText in $tries) {
      try { $cell.Formula2 = "=$n($argsText)" } catch { continue }
      $loc = [string]$cell.Formula2Local
      # неизвестную функцию Excel хранит как _xludf.ИМЯ или оставляет английское имя с ошибкой #ИМЯ?
      $known = -not ($cell.Formula2 -match '_xludf') -and -not ($cell.Text -eq '#ИМЯ?' -and $loc -match "^=$([regex]::Escape($n))\(")
      if ($known) { $ru = ($loc -replace '^=', '') -replace '\(.*$', '' }
      break
    }
    $result[$n] = $ru
  }
  $wb.Close($false)
} finally { $x.Quit() }
$json = $result | ConvertTo-Json
[IO.File]::WriteAllText([IO.Path]::GetFullPath($Out), $json, (New-Object Text.UTF8Encoding $false))
"ok: известно Excel $(($result.Values | Where-Object { $_ }).Count) из $($names.Count)"
