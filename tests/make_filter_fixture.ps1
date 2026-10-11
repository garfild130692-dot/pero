# tests/fixtures/filter.xlsx — создаётся самим Excel (COM): автофильтр «Регион = Юг» и «100 <= Сумма < 300» (tests/filter.test.js).
# Автор — нейтральный «Перо Тест» (UserName возвращается после генерации). Файл в UTF-8 с BOM (PowerShell 5.1).
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$oldName = $xl.UserName
$xl.UserName = 'Перо Тест'
try {
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    $data = @(@('Регион','Товар','Сумма','Дата'), @('Север','Стол',100,'15.01.2026'), @('Юг','Стул',250,'20.02.2026'), @('Север','Шкаф',$null,'03.03.2026'), @('Запад','Стол',75,'15.01.2026'), @('Юг','Стол',300,'01.04.2026'))
    for ($r = 0; $r -lt $data.Count; $r++) { for ($c = 0; $c -lt 4; $c++) { $v = $data[$r][$c]; if ($v -ne $null) { $ws.Cells.Item($r + 1, $c + 1).FormulaLocal = [string]$v } } }
    $rg = $ws.Range('A1:D6')
    [void]$rg.AutoFilter(1, 'Юг')
    [void]$rg.AutoFilter(3, '>=100', 1, '<300')
    $wb.SaveAs([IO.Path]::GetFullPath('tests/fixtures/filter.xlsx'), 51)
    $wb.Close($false)
} finally {
  $xl.UserName = $oldName
  $xl.Quit()
}
