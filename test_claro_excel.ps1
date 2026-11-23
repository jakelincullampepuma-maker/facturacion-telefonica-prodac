# Verifica la tabla dinámica con Microsoft Excel en una copia de prueba.
param(
    [string]$WorkbookPath = 'data/claro-results/restored-july.xlsx',
    [string]$ResultPath = 'data/claro-results/excel-verification.json',
    [ValidateSet('Hoja1', 'Hoja2')][string]$SheetName = 'Hoja1'
)
$ErrorActionPreference = 'Stop'
$excelApp = $null
$book = $null
$normalOpen = $null
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class ClaroRunningWorkbook {
    [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx context);
    public static object Find(string filename) {
        IRunningObjectTable table; GetRunningObjectTable(0, out table);
        IBindCtx context; CreateBindCtx(0, out context);
        IEnumMoniker enumerator; table.EnumRunning(out enumerator);
        IMoniker[] current = new IMoniker[1];
        try {
            while (enumerator.Next(1, current, IntPtr.Zero) == 0) {
                try {
                    string name; current[0].GetDisplayName(context, null, out name);
                    if (name.TrimStart('!').Equals(filename, StringComparison.OrdinalIgnoreCase)) {
                        object book; table.GetObject(current[0], out book); return book;
                    }
                } catch { } finally { Marshal.ReleaseComObject(current[0]); }
            }
        } finally { Marshal.ReleaseComObject(enumerator); Marshal.ReleaseComObject(context); Marshal.ReleaseComObject(table); }
        return null;
    }
}
'@
function Assert-True($condition, $message) { if (-not $condition) { throw $message } }
function Assert-Amount($actual, $expected, $message) {
    Assert-True ([Math]::Abs([double]$actual - [double]$expected) -lt 0.00001) $message
}
function Total($pivot) { return [double]$pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2 }
function CcTotal($pivot, [string]$cc) {
    for ($r = 2; $r -lt $pivot.TableRange1.Rows.Count; $r++) {
        if ([string]$pivot.TableRange1.Cells.Item($r, 1).Value2 -eq $cc) { return [double]$pivot.TableRange1.Cells.Item($r, 2).Value2 }
    }
    return [double]0
}
function Refresh($app, $workbook) { $workbook.RefreshAll(); $app.CalculateUntilAsyncQueriesDone(); $app.CalculateFull() }
try {
    $path = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AutomationSecurity = 3
    $book = $excelApp.Workbooks.Open($path, 0, $true)
    Assert-True (-not $book.RepairMode) 'Excel abrió el libro en modo reparación.'
    Assert-True ($book.Worksheets.Count -eq 3) 'Deben existir exactamente tres hojas.'
    Assert-True ($book.Worksheets.Item(1).Name -eq 'Factura Detallada' -and $book.Worksheets.Item(2).Name -eq 'Hoja2' -and $book.Worksheets.Item(3).Name -eq 'Hoja1') 'Orden de hojas incorrecto.'
    $sheet = $book.Worksheets.Item($SheetName)
    $tableName = if ($SheetName -eq 'Hoja1') { 'ClaroFinal' } else { 'ClaroTrabajo' }
    Assert-True ($sheet.ListObjects.Count -eq 1 -and $sheet.PivotTables().Count -eq 1) 'Falta tabla principal o tabla dinámica nativa.'
    $source = $sheet.ListObjects.Item(1)
    $pivot = $sheet.PivotTables().Item(1)
    Assert-True ($source.HeaderRowRange.Row -eq 1) 'Los encabezados deben estar en la fila 1.'
    Assert-True ($source.Name -eq $tableName) 'Origen estructurado incorrecto.'
    Assert-True (-not $source.ShowAutoFilter) 'A:C no debe mostrar filtros.'
    Assert-True ($pivot.TableRange1.Row -eq 7 -and $pivot.TableRange1.Column -eq 7) 'El pivote debe comenzar en G7.'
    Assert-True (-not $source.Range.Font.Bold -and $source.Range.Font.Color -eq 0) 'A:C debe tener texto negro sin negrita.'
    Assert-True ($source.Range.Interior.ColorIndex -eq -4142 -or $source.Range.Interior.Color -eq 16777215) 'A:C debe tener fondo blanco.'
    for ($border = 7; $border -le 12; $border++) {
        Assert-True ($source.Range.Borders.Item($border).LineStyle -eq -4142) 'A:C no debe tener bordes decorativos.'
    }
    Assert-True ($pivot.PivotCache().RefreshOnFileOpen) 'No está activada la actualización al abrir.'
    Assert-True ([string]$pivot.PivotCache().SourceData -eq $tableName) 'El origen del pivote debe ser la tabla.'
    if ($SheetName -eq 'Hoja1') {
        Assert-True ($null -eq $sheet.Range('I7').Value2 -and $null -eq $sheet.Range('J7').Value2) 'Hoja1 no debe contener columnas de prorrateo.'
        Assert-Amount (CcTotal $pivot 'Varios') 0 'Hoja1 no debe incluir Varios.'
    } else {
        Assert-True ($sheet.Range('I7').Value2 -ceq 'prorrateo' -and $sheet.Range('J7').Value2 -ceq 'total ajustado') 'Faltan las columnas de trabajo.'
        Assert-True (-not $sheet.Range('I7:J7').Font.Bold -and $sheet.Range('I7:J7').Font.Color -eq 0) 'Los titulos deben ser sencillos.'
        Assert-True ($sheet.Range('I7:J7').Interior.ColorIndex -eq -4142) 'Los titulos de prorrateo no deben tener relleno.'
        Assert-True ($sheet.Range('I8').NumberFormat -notmatch '\[Red\]' -and $sheet.Range('I8').Font.Color -eq 0) 'El prorrateo negativo debe verse negro.'
        $lastRow = $pivot.TableRange1.Row + $pivot.TableRange1.Rows.Count - 1
        Assert-Amount $sheet.Cells.Item($lastRow, 9).Value2 (CcTotal $pivot 'Varios') 'Los prorrateos no suman Varios.'
        Assert-Amount $sheet.Cells.Item($lastRow, 10).Value2 (Total $pivot) 'Los totales ajustados no coinciden.'
    }
    $initial = [double]0
    for ($r = 1; $r -le $source.ListRows.Count; $r++) { $initial += [double]$source.DataBodyRange.Cells.Item($r, 2).Value2 }
    Assert-Amount (Total $pivot) $initial 'El total inicial no coincide con la tabla principal.'
    Refresh $excelApp $book
    Assert-Amount (Total $pivot) $initial 'Actualizar todo cambió el total.'
    $originalAmount = [double]$source.DataBodyRange.Cells.Item(1, 2).Value2
    $originalCC = [string]$source.DataBodyRange.Cells.Item(1, 3).Value2
    $targetCC = 'CC-PRUEBA-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount + 1.23
    $source.DataBodyRange.Cells.Item(1, 3).Value2 = [string]$targetCC
    Refresh $excelApp $book
    Assert-Amount (Total $pivot) ($initial + 1.23) 'El pivote no refleja el cambio de monto.'
    Assert-Amount (CcTotal $pivot $targetCC) ($originalAmount + 1.23) 'El pivote no refleja el cambio de CC.'
    $rowsBefore = $source.ListRows.Count
    $newLine = $source.ListRows.Add()
    $newLine.Range.Cells.Item(1, 1).Value2 = 'AJUSTE GENERAL'
    $newLine.Range.Cells.Item(1, 2).Value2 = [double]-3.21
    $newLine.Range.Cells.Item(1, 3).Value2 = [string]$targetCC
    Refresh $excelApp $book
    Assert-Amount (Total $pivot) ($initial + 1.23 - 3.21) 'El pivote no incluye la nueva fila negativa.'
    Assert-True ($pivot.PivotCache().RecordCount -eq $rowsBefore + 1) 'La tabla dinámica no amplió su origen.'
    $source.ListRows.Item($rowsBefore + 1).Delete()
    Refresh $excelApp $book
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount + 2.34
    $reopenPath = Join-Path (Split-Path -Parent $path) ('claro-refresh-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $book.SaveCopyAs($reopenPath)
    $excelExecutable = Join-Path $excelApp.Path 'EXCEL.EXE'
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    $book = $null
    $excelApp.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
    $excelApp = $null
    # Workbooks.Open no ejecuta RefreshOnFileOpen: se usa una apertura normal.
    $normalOpen = Start-Process -FilePath $excelExecutable -ArgumentList @('/x', '/e', '/r', ('"' + $reopenPath + '"')) -WindowStyle Hidden -PassThru
    for ($attempt = 0; $attempt -lt 40 -and $null -eq $book; $attempt++) {
        Start-Sleep -Milliseconds 500
        $book = [ClaroRunningWorkbook]::Find($reopenPath)
    }
    Assert-True ($null -ne $book) 'No se pudo inspeccionar el libro abierto normalmente.'
    $excelApp = $book.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.CalculateUntilAsyncQueriesDone()
    Assert-True (-not $book.RepairMode) 'La copia requirió reparación.'
    $pivot = $book.Worksheets.Item($SheetName).PivotTables().Item(1)
    Assert-Amount (Total $pivot) ($initial + 2.34) 'La apertura normal no actualizó el importe.'
    Assert-Amount (CcTotal $pivot $targetCC) ($originalAmount + 2.34) 'La apertura normal no actualizó el CC.'
    $result = [ordered]@{ OpenedWithMicrosoftExcel = $true; RepairMode = [bool]$book.RepairMode;
        NativePivotTables = 1; Sheet = $SheetName; StructuredSource = $tableName; InitialTotal = $initial;
        SimpleFormattingPassed = $true; PivotStartsAtG7 = $true; SourceFiltersHidden = $true;
        ChangedAmountPassed = $true; ChangedCostCenterPassed = $true; AddedNegativeRowPassed = $true;
        RefreshAllPassed = $true; RefreshOnReopenPassed = $true; OriginalFileModified = $false }
    $json = $result | ConvertTo-Json
    $json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
    Write-Output $json
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excelApp) { $excelApp.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
    if ($null -ne $normalOpen -and -not $normalOpen.HasExited) {
        [void]$normalOpen.WaitForExit(2000)
        if (-not $normalOpen.HasExited) { $normalOpen.Kill() }
    }
}
