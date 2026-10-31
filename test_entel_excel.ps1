# Prueba del XLSX descargado con Microsoft Excel. No guarda los cambios de prueba.
param(
    [string]$WorkbookPath = 'data/browser-results/Entel-file.xlsx',
    [string]$ResultPath = 'data/format-tests/excel-verification.json'
)
$ErrorActionPreference = 'Stop'
$excelApp = $null
$book = $null
function Assert-True($condition, $message) {
    if (-not $condition) { throw $message }
}
function Assert-Amount($actual, $expected, $message) {
    Assert-True ([Math]::Abs([double]$actual - [double]$expected) -lt 0.00001) $message
}
try {
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AutomationSecurity = 3
    $path = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $book = $excelApp.Workbooks.Open($path, 0, $true)
    Assert-True (-not $book.RepairMode) 'Excel abrió el libro en modo reparación.'
    $sheet = $book.Worksheets.Item(1)
    Assert-True ($sheet.Name -eq 'Resultado Entel') 'La primera hoja es incorrecta.'
    Assert-True ($book.Windows.Item(1).SelectedSheets.Count -eq 1) 'El archivo abre pestañas agrupadas.'
    $sheet.Select()
    Assert-True ($sheet.Range('A1').Text -eq 'NUMERO') 'Falta NUMERO en A1.'
    Assert-True ($sheet.Range('B1').Text -eq 'MONTO') 'Falta MONTO en B1.'
    Assert-True ($sheet.Range('C1').Text -eq 'CC') 'Falta CC en C1.'
    Assert-True ($sheet.ListObjects.Count -eq 1) 'Falta la tabla principal con filtros.'
    Assert-True ($sheet.PivotTables().Count -eq 1) 'Falta la tabla dinámica nativa.'
    $source = $sheet.ListObjects.Item(1)
    $pivot = $sheet.PivotTables().Item(1)
    Assert-True ($pivot.RowFields(1).Name -eq 'CC') 'La agrupación no usa CC.'
    Assert-True ($pivot.DataFields(1).Function -eq -4157) 'El resumen no suma MONTO.'
    Assert-True ($pivot.TableRange1.Row -eq 4 -and $pivot.TableRange1.Column -eq 7) 'Ubicación de tabla dinámica incorrecta.'
    $lineCount = $source.ListRows.Count
    $initial = [double]$source.TotalsRowRange.Cells.Item(1, 2).Value2
    $initialPivot = [double]$pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2
    Assert-Amount $initialPivot $initial 'Los totales iniciales no coinciden.'
    $fieldItemCount = $pivot.PivotFields('CC').PivotItems().Count
    $refreshed = $pivot.RefreshTable()
    Assert-True $refreshed 'Falló Actualizar tabla dinámica.'
    $excelApp.CalculateFull()
    Assert-Amount $pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2 $initial 'Actualizar alteró el total.'
    # Cambiar un monto en memoria y actualizar la tabla dinámica real.
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = [double]$source.DataBodyRange.Cells.Item(1, 2).Value2 + 1.23
    [void]$pivot.PivotCache().Refresh()
    [void]$pivot.RefreshTable()
    $excelApp.CalculateFull()
    Assert-Amount $pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2 ($initial + 1.23) 'No se actualizó el monto modificado.'
    # Agregar una línea y un CC nuevo: el origen estructurado debe expandirse.
    $newLine = $source.ListRows.Add()
    $newLine.Range.Cells.Item(1, 1).Value2 = '954567890'
    $newLine.Range.Cells.Item(1, 2).Value2 = [double]7.89
    $newLine.Range.Cells.Item(1, 3).Value2 = 'CC-PRUEBA-NUEVO'
    [void]$pivot.PivotCache().Refresh()
    [void]$pivot.RefreshTable()
    $excelApp.CalculateFull()
    Assert-True ($source.ListRows.Count -eq $lineCount + 1) 'No se amplió la tabla principal.'
    Assert-True ($pivot.PivotFields('CC').PivotItems().Count -eq $fieldItemCount + 1) 'No apareció el CC nuevo en el filtro.'
    Assert-Amount $pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2 ($initial + 1.23 + 7.89) 'El resumen excluye la línea añadida.'
    Assert-Amount $source.TotalsRowRange.Cells.Item(1, 2).Value2 ($initial + 1.23 + 7.89) 'El total principal excluye la línea añadida.'
    $verification = [ordered]@{
        OpenedWithMicrosoftExcel = $true
        RepairMode = [bool]$book.RepairMode
        SheetCount = $book.Worksheets.Count
        FirstSheet = $sheet.Name
        Lines = $lineCount
        MainTotal = [Math]::Round($initial, 2)
        PivotTotal = [Math]::Round($initialPivot, 2)
        NativePivotTablesOnResult = $sheet.PivotTables().Count
        CostCenters = $fieldItemCount
        RefreshPassed = $true
        ChangedAmountPassed = $true
        NewLineAndCostCenterPassed = $true
        TestChangesSaved = $false
    }
    $json = $verification | ConvertTo-Json
    $json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
    Write-Output $json
} finally {
    if ($null -ne $book) { $book.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excelApp) { $excelApp.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
}
