# Verifica la descarga real con Microsoft Excel, sin guardar cambios de prueba.
param(
    [string]$WorkbookPath = 'data/browser-results/Movistar_Conciliado_Noviembre_2026.xlsx',
    [string]$ResultPath = 'data/browser-results/movistar-excel-verification.json'
)
$ErrorActionPreference = 'Stop'
$excelApp = $null
$book = $null
$normalOpen = $null
# Consultar exclusivamente un libro ya abierto; no usar BindToMoniker, que puede
# abrirlo por automatización y omitir RefreshOnFileOpen antes de la apertura normal.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class MovistarRunningWorkbook {
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
function Assert-True($condition, $message) {
    if (-not $condition) { throw $message }
}
function Assert-Amount($actual, $expected, $message) {
    Assert-True ([Math]::Abs([double]$actual - [double]$expected) -lt 0.00001) $message
}
function Update-All($application, $workbook) {
    $workbook.RefreshAll()
    $application.CalculateUntilAsyncQueriesDone()
    $application.CalculateFull()
}
function Get-PivotTotal($pivotTable) {
    return [double]$pivotTable.TableRange1.Cells.Item($pivotTable.TableRange1.Rows.Count, 2).Value2
}
function Get-CcTotal($pivotTable, [string]$costCenter) {
    for ($index = 2; $index -lt $pivotTable.TableRange1.Rows.Count; $index++) {
        if ([string]$pivotTable.TableRange1.Cells.Item($index, 1).Value2 -eq $costCenter) {
            return [double]$pivotTable.TableRange1.Cells.Item($index, 2).Value2
        }
    }
    return [double]0
}
try {
    $path = (Resolve-Path -LiteralPath $WorkbookPath).Path
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.AutomationSecurity = 3
    $book = $excelApp.Workbooks.Open($path, 0, $true)
    Assert-True (-not $book.RepairMode) 'Excel abrió el libro en modo reparación.'
    Assert-True ($book.Worksheets.Count -eq 3) 'Se requieren exactamente tres hojas.'
    $names = @('CARGOS ADICIONALES', 'CARGO MENSUAL', 'Hoja1')
    for ($index = 1; $index -le 3; $index++) {
        Assert-True ($book.Worksheets.Item($index).Name -eq $names[$index - 1]) 'Orden de hojas incorrecto.'
    }
    $ca = $book.Worksheets.Item(1)
    $cm = $book.Worksheets.Item(2)
    $h1 = $book.Worksheets.Item(3)
    $headers = @(@('NUMERO', 'DESCRIPCION', 'MONTO'), @(('N' + [char]0x00B0), 'NUMERO', 'PLAN', 'MONTO', 'DESC', 'MONTO F'), @('NUMERO', 'MONTO', 'CC'))
    for ($index = 1; $index -le 3; $index++) {
        for ($col = 1; $col -le $headers[$index - 1].Count; $col++) {
            Assert-True ($book.Worksheets.Item($index).Cells.Item(1, $col).Value2 -eq $headers[$index - 1][$col - 1]) 'Encabezados incorrectos en la fila 1.'
        }
    }
    Assert-True (-not $book.HasVBProject -and $book.FileFormat -eq 51) 'El archivo debe ser XLSX sin macros.'
    Assert-True ($ca.Range('C94').Formula -eq '=SUM(C2:C93)') 'Fórmula de cuotas incorrecta.'
    for ($row = 2; $row -le 326; $row++) {
        Assert-True ($cm.Cells.Item($row, 6).Formula -eq "=D$row-E$row") "Fórmula MONTO F incorrecta en fila $row."
        Assert-True ($cm.Cells.Item($row, 2).Value2 -is [string]) 'Teléfono de CARGO MENSUAL no es texto.'
    }
    for ($row = 2; $row -le 326; $row++) {
        Assert-True ($h1.Cells.Item($row, 1).Value2 -is [string]) 'Teléfono de Hoja1 no es texto.'
    }
    Assert-True ($h1.ListObjects.Count -eq 1) 'Falta la tabla estructurada de origen.'
    $source = $h1.ListObjects.Item(1)
    Assert-True ($source.Name -eq 'MovistarDatos') 'Nombre de tabla de origen incorrecto.'
    Assert-True ($source.ListRows.Count -eq 325 -and $source.ListColumns.Count -eq 3) 'La tabla debe incluir las 325 líneas y los tres campos.'
    Assert-True ($source.AutoFilter.Range.Address() -eq '$A$1:$C$326') 'El filtro debe cubrir todos los datos.'
    Assert-True ($source.ShowTotals) 'Falta la fila de totales de la tabla estructurada.'
    Assert-True ($h1.PivotTables().Count -eq 1) 'Falta la tabla dinámica nativa.'
    $pivot = $h1.PivotTables().Item(1)
    Assert-True ($pivot.RowFields(1).Name -eq 'CC') 'La agrupación no utiliza CC.'
    Assert-True ($pivot.DataFields(1).Name -eq 'Suma de MONTO') 'El valor no tiene el nombre requerido.'
    Assert-True ($pivot.DataFields(1).Function -eq -4157) 'La operación no es suma.'
    Assert-True ($pivot.TableRange1.Row -eq 4 -and $pivot.TableRange1.Column -eq 7) 'Tabla dinámica fuera de G4.'
    Assert-True ($pivot.PivotCache().RecordCount -eq 325) 'La caché no contiene los 325 registros.'
    Assert-True ($pivot.PivotCache().SourceData -eq 'MovistarDatos') 'El pivote utiliza un rango fijo.'
    Assert-True ($pivot.PivotCache().RefreshOnFileOpen) 'Falta Actualizar al abrir el archivo.'
    $excelApp.CalculateFull()
    Assert-Amount $ca.Range('C94').Value2 25452.11 'Total de cargos incorrecto.'
    Assert-Amount $cm.Range('D327').Value2 12597.72 'MONTO incorrecto.'
    Assert-Amount $cm.Range('E327').Value2 1046.68 'DESC incorrecto.'
    Assert-Amount $cm.Range('F327').Value2 11551.04 'MONTO F incorrecto.'
    Assert-Amount $cm.Range('D328').Value2 11551.04 'Subtotal PDF incorrecto.'
    Assert-Amount $cm.Range('D329').Value2 0 'La diferencia no es cero.'
    Assert-Amount $h1.Range('B327').Value2 11551.04 'Total de Hoja1 incorrecto.'
    Assert-Amount $pivot.TableRange1.Cells.Item($pivot.TableRange1.Rows.Count, 2).Value2 11551.04 'Total dinámico incorrecto.'
    $initial = Get-PivotTotal $pivot
    $originalAmount = [double]$source.DataBodyRange.Cells.Item(1, 2).Value2
    $originalCC = [string]$source.DataBodyRange.Cells.Item(1, 3).Value2
    $originalCcTotal = Get-CcTotal $pivot $originalCC
    $targetCC = 'CC-PRUEBA-' + [guid]::NewGuid().ToString('N').Substring(0, 8)

    # Datos -> Actualizar todo después de cambiar MONTO.
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount + 1.23
    Update-All $excelApp $book
    Assert-Amount (Get-PivotTotal $pivot) ($initial + 1.23) 'Actualizar todo no refleja el monto cambiado.'
    Assert-Amount (Get-CcTotal $pivot $originalCC) ($originalCcTotal + 1.23) 'El CC no refleja el monto cambiado.'
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount
    Update-All $excelApp $book

    # Mover una línea a un CC nuevo; verificar ambos lados del movimiento.
    $source.DataBodyRange.Cells.Item(1, 3).Value2 = [string]$targetCC
    Update-All $excelApp $book
    Assert-Amount (Get-CcTotal $pivot $originalCC) ($originalCcTotal - $originalAmount) 'El CC anterior no resta la línea.'
    Assert-Amount (Get-CcTotal $pivot $targetCC) $originalAmount 'El CC nuevo no recibe la línea.'
    Assert-Amount (Get-PivotTotal $pivot) $initial 'Cambiar CC altera el total general.'
    $source.DataBodyRange.Cells.Item(1, 3).Value2 = [string]$originalCC
    Update-All $excelApp $book
    Assert-Amount (Get-CcTotal $pivot $originalCC) $originalCcTotal 'No se restauró el CC original.'

    # Insertar y eliminar una fila de la tabla; el origen debe ampliarse/reducirse.
    $newLine = $source.ListRows.Add()
    $newLine.Range.Cells.Item(1, 1).Value2 = '954567890'
    $newLine.Range.Cells.Item(1, 2).Value2 = [double]7.89
    $newLine.Range.Cells.Item(1, 3).Value2 = [string]$targetCC
    Update-All $excelApp $book
    Assert-True ($source.ListRows.Count -eq 326 -and $pivot.PivotCache().RecordCount -eq 326) 'El pivote excluye la fila nueva.'
    Assert-Amount (Get-PivotTotal $pivot) ($initial + 7.89) 'Total incorrecto al agregar una fila.'
    Assert-Amount (Get-CcTotal $pivot $targetCC) 7.89 'El CC nuevo excluye la fila añadida.'
    Assert-Amount $source.TotalsRowRange.Cells.Item(1, 2).Value2 ($initial + 7.89) 'La suma principal no incluye la fila nueva.'
    $source.ListRows.Item(326).Delete()
    Update-All $excelApp $book
    Assert-True ($source.ListRows.Count -eq 325 -and $pivot.PivotCache().RecordCount -eq 325) 'La fila eliminada permanece en el pivote.'
    Assert-Amount (Get-PivotTotal $pivot) 11551.04 'El total no volvió al valor original.'
    Assert-Amount $source.TotalsRowRange.Cells.Item(1, 2).Value2 11551.04 'La suma principal no se restauró.'

    # Guardar una copia con datos modificados y caché sin actualizar. La apertura
    # debe refrescarla automáticamente, sin llamar a Actualizar todo primero.
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount + 2.34
    $source.DataBodyRange.Cells.Item(1, 3).Value2 = [string]$targetCC
    $reopenPath = Join-Path (Split-Path -Parent $path) ('movistar-refresh-on-open-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $book.SaveCopyAs($reopenPath)
    $excelExecutable = Join-Path $excelApp.Path 'EXCEL.EXE'
    $book.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    $book = $null
    $excelApp.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp)
    $excelApp = $null
    # Workbooks.Open omite RefreshOnFileOpen, según Microsoft. Abrir por la vía
    # normal en una instancia propia y enlazar después el libro ya abierto.
    # https://learn.microsoft.com/office/vba/api/excel.pivotcache.refreshonfileopen
    $normalOpen = Start-Process -FilePath $excelExecutable -ArgumentList @('/x', '/e', '/r', ('"' + $reopenPath + '"')) -WindowStyle Hidden -PassThru
    for ($attempt = 0; $attempt -lt 30 -and $null -eq $book; $attempt++) {
        Start-Sleep -Milliseconds 500
        $book = [MovistarRunningWorkbook]::Find($reopenPath)
    }
    Assert-True ($null -ne $book) 'No se pudo enlazar la copia abierta normalmente con Excel.'
    $excelApp = $book.Application
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    Assert-True (-not $book.RepairMode) 'La copia necesitó reparación al reabrir.'
    $excelApp.CalculateUntilAsyncQueriesDone()
    $h1 = $book.Worksheets.Item('Hoja1')
    $source = $h1.ListObjects.Item(1)
    $pivot = $h1.PivotTables().Item(1)
    for ($attempt = 0; $attempt -lt 20 -and [Math]::Abs((Get-PivotTotal $pivot) - ($initial + 2.34)) -gt 0.00001; $attempt++) {
        Start-Sleep -Milliseconds 500
    }
    Assert-Amount (Get-PivotTotal $pivot) ($initial + 2.34) 'Actualizar al abrir no reflejó el monto.'
    Assert-Amount (Get-CcTotal $pivot $targetCC) ($originalAmount + 2.34) 'Actualizar al abrir no reflejó el CC nuevo.'
    Assert-Amount (Get-CcTotal $pivot $originalCC) ($originalCcTotal - $originalAmount) 'Actualizar al abrir no descontó el CC anterior.'
    $source.DataBodyRange.Cells.Item(1, 2).Value2 = $originalAmount
    $source.DataBodyRange.Cells.Item(1, 3).Value2 = [string]$originalCC
    Update-All $excelApp $book
    Assert-Amount (Get-PivotTotal $pivot) 11551.04 'El total final restaurado es incorrecto.'
    $verification = [ordered]@{
        OpenedWithMicrosoftExcel = $true
        RepairMode = [bool]$book.RepairMode
        SheetNames = $names
        Lines = 325
        FinancingRows = 92
        FormulaRowsVerified = 325
        Amount = 12597.72
        Discounts = 1046.68
        Subtotal = 11551.04
        UnaffectedCharges = 25452.11
        Difference = 0
        NativePivotTables = $h1.PivotTables().Count
        PivotTotal = 11551.04
        RefreshPassed = $true
        ChangedAmountPassed = $true
        HeadersOnFirstRow = $true
        StructuredSource = 'MovistarDatos'
        ChangedCostCenterPassed = $true
        AddedRowPassed = $true
        DeletedRowPassed = $true
        RefreshAllPassed = $true
        RefreshOnReopenPassed = $true
        ReopenMethod = 'Normal Excel file open; inspect the registered workbook without opening it through COM'
        RestoredTotal = Get-PivotTotal $pivot
        OriginalDownloadModified = $false
        TestChangesSaved = $false
    }
    $json = $verification | ConvertTo-Json
    $json | Set-Content -LiteralPath $ResultPath -Encoding UTF8
    Write-Output $json
} finally {
    if ($null -ne $book) { $book.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excelApp) { $excelApp.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) }
    if ($null -ne $normalOpen -and -not $normalOpen.HasExited) {
        [void]$normalOpen.WaitForExit(2000)
        if (-not $normalOpen.HasExited) { $normalOpen.Kill() }
    }
}
