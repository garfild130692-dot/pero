# Генерирует tests/fixtures/book.csv средствами Excel (COM): лист «Продажи» из book.xlsx, сохранённый как
# «CSV (разделители — запятые)» с региональными настройками (Local) — в русской Windows это «;» и кодировка Windows-1251.
# Запуск: powershell -ExecutionPolicy Bypass -File tests/make_csv_fixture.ps1
param([string]$Src = 'tests/fixtures/book.xlsx', [string]$Out = 'tests/fixtures/book.csv')
$ErrorActionPreference = 'Stop'
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Open([IO.Path]::GetFullPath($Src))
  $wb.Worksheets.Item(1).Activate()
  # SaveAs(Filename, FileFormat=6 xlCSV, ..., Local=True) — 12-й параметр
  $m = [Type]::Missing
  $wb.SaveAs([IO.Path]::GetFullPath($Out), 6, $m, $m, $m, $m, 1, $m, $m, $m, $m, $true)
  $wb.Close($false)
  "ok: $Out"
} finally {
  $x.Quit()
}
