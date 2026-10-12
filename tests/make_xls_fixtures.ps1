# Эталоны .xls (этап 19): сам Excel (COM) пересохраняет эталонные книги .xlsx в формат Excel 97–2003 (FileFormat 56).
# Перо читает обе версии одной книги — модели должны совпасть (tests/xls.test.js).
#   book.xlsx → book.xls; preserve.xlsx → preserve.xls; formulas.xlsx → formulas.xls
# Автор — нейтральный «Перо Тест» (UserName возвращается). Файл в UTF-8 с BOM (PowerShell 5.1).
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_xls_fixtures.ps1
$ErrorActionPreference = 'Stop'
$dir = [IO.Path]::GetFullPath('tests/fixtures')
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$oldName = $x.UserName
$x.UserName = 'Перо Тест'
try {
  foreach ($name in @('book', 'preserve', 'formulas')) {
    $src = Join-Path $dir "$name.xlsx"
    $dst = Join-Path $dir "$name.xls"
    if (Test-Path $dst) { Remove-Item $dst }
    $wb = $x.Workbooks.Open($src)
    $wb.SaveAs($dst, 56)
    $wb.Close($false)
    Write-Output "$name.xls"
  }
} finally {
  $x.UserName = $oldName
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
