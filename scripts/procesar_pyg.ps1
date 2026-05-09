# ── procesar_pyg.ps1 ─────────────────────────────────────────────────────────
# Reads the Excel P&G (Income Statement) files and extracts monthly metrics.
# Returns an array of hashtables, one per month.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\..\config\rutas.ps1"

$MONTH_ABBREV = @{
    "Enero"="Ene"; "Febrero"="Feb"; "Marzo"="Mar"; "Abril"="Abr"
    "Mayo"="May"; "Junio"="Jun"; "Julio"="Jul"; "Agosto"="Ago"
    "Septiembre"="Sep"; "Octubre"="Oct"; "Noviembre"="Nov"; "Diciembre"="Dic"
}

function ConvertMonthLabel($text) {
    $t = $text -replace "`r`n|`n|`r", " " -replace "\s+", " "
    foreach ($k in $MONTH_ABBREV.Keys) { $t = $t -replace $k, $MONTH_ABBREV[$k] }
    $t = $t -replace "20(\d\d)", '$1'
    return $t.Trim()
}

function GetRowsByLabel($ws, $lastRow) {
    $map = @{}
    for ($r = 1; $r -le $lastRow; $r++) {
        $txt = $ws.Cells.Item($r, 1).Text.Trim()
        if ($txt -eq "") { continue }
        if (-not $map.ContainsKey($txt)) {
            $map[$txt] = [System.Collections.Generic.List[int]]::new()
        }
        $map[$txt].Add($r)
    }
    return $map
}

function ReadCellPyG($ws, $row, $col) {
    if ($row -eq 0) { return 0 }
    $v = $ws.Cells.Item($row, $col).Value2
    if ($null -eq $v) { return 0 }
    $n = $v -as [double]
    if ($null -eq $n) { return 0 }
    return [math]::Round($n)
}

function ReadPyGFile($excel, $filePath, $year) {
    if (-not (Test-Path $filePath)) {
        Write-Warning "File not found: $filePath"
        return @()
    }
    Write-Host "  Reading $year..." -ForegroundColor Cyan

    $wb      = $excel.Workbooks.Open($filePath)
    $ws      = $wb.Sheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count

    # Scan row 8 for month headers; store as string keys to avoid
    # integer-indexing on OrderedDictionary (ArgumentOutOfRangeException bug)
    $monthColKeys = [System.Collections.Generic.List[string]]::new()
    $monthColMap  = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $header = $ws.Cells.Item(8, $col).Text.Trim()
        if ($header -eq "" -or $header -match "Total|Codigo|Nombre") { continue }
        $label = ConvertMonthLabel $header
        # Only accept labels that look like a real month abbreviation
        if ($label -match "^(Ene|Feb|Mar|Abr|May|Jun|Jul|Ago|Sep|Oct|Nov|Dic)") {
            $key = "$col"
            $monthColKeys.Add($key)
            $monthColMap[$key] = $label
        }
    }

    $rows = GetRowsByLabel $ws $lastRow

    $rowSales     = if ($rows["Total Operacionales"])           { $rows["Total Operacionales"][0] }           else { 0 }
    $rowRevenue   = if ($rows["Total Ingresos"])                { $rows["Total Ingresos"][0] }                else { 0 }
    $rowPersonal  = if ($rows["Total Gastos de personal"])      { $rows["Total Gastos de personal"][-1] }     else { 0 }
    $rowRent      = if ($rows["Total Arrendamientos"])          { $rows["Total Arrendamientos"][-1] }         else { 0 }
    $rowServices  = if ($rows["Total Servicios"])               { $rows["Total Servicios"][-1] }              else { 0 }
    $rowFees      = if ($rows["Total Honorarios"])              { $rows["Total Honorarios"][-1] }             else { 0 }
    $rowTaxes     = if ($rows["Total Impuestos"])               { $rows["Total Impuestos"][-1] }              else { 0 }
    $rowMisc      = if ($rows["Total Diversos"])                { $rows["Total Diversos"][-1] }               else { 0 }
    $rowSellExp   = if ($rows["Total Operacionales de ventas"]) { $rows["Total Operacionales de ventas"][0] } else { 0 }
    $rowTotalExp  = if ($rows["Total Gastos"])                  { $rows["Total Gastos"][0] }                  else { 0 }
    $rowCOGS      = if ($rows["Total Costos de ventas"])        { $rows["Total Costos de ventas"][0] }        else { 0 }
    $rowResult    = if ($rows["Resultado del Ejercicio"])       { $rows["Resultado del Ejercicio"][0] }       else { 0 }

    $data = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($key in $monthColKeys) {
        $col = [int]$key
        $mes = $monthColMap[$key]

        $sales      = ReadCellPyG $ws $rowSales    $col
        $revenue    = ReadCellPyG $ws $rowRevenue  $col
        $cogs       = ReadCellPyG $ws $rowCOGS     $col
        $personal   = ReadCellPyG $ws $rowPersonal $col
        $rent       = ReadCellPyG $ws $rowRent     $col
        $services   = ReadCellPyG $ws $rowServices $col
        $fees       = ReadCellPyG $ws $rowFees     $col
        $taxes      = ReadCellPyG $ws $rowTaxes    $col
        $misc       = ReadCellPyG $ws $rowMisc     $col
        $sellExp    = ReadCellPyG $ws $rowSellExp  $col
        $totalExp   = ReadCellPyG $ws $rowTotalExp $col
        $result     = ReadCellPyG $ws $rowResult   $col

        $grossMargin = if ($sales -ne 0) { [math]::Round((($sales - $cogs) / $sales) * 100, 2) } else { 0 }

        $data.Add(@{
            periodo      = $year
            mes          = $mes
            ventas       = $sales
            ingresos     = $revenue
            costo_ventas = $cogs
            margen_pesos = $sales - $cogs
            margen_pct   = $grossMargin
            personal     = $personal
            arriendo     = $rent
            servicios    = $services
            honorarios   = $fees
            impuestos    = $taxes
            diversos     = $misc
            gastos_ventas = $sellExp
            total_gastos = $totalExp
            resultado    = $result
        })
    }

    $wb.Close($false)
    Write-Host "    -> $($data.Count) months read" -ForegroundColor Green
    return $data.ToArray()
}

function ProcessAllPyG {
    Write-Host "`nReading P&G files..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    $allData = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($year in @("2024","2025","2026")) {
        $path   = $script:CONFIG.Archivos_PyG[$year]
        $months = ReadPyGFile $excel $path $year
        foreach ($m in $months) { $allData.Add($m) }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "`nTotal months processed: $($allData.Count)" -ForegroundColor Green
    return $allData.ToArray()
}
