# ── procesar_balance.ps1 ─────────────────────────────────────────────────────
# Reads Balance Sheet Excel files and extracts monthly liquidity indicators.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\..\config\rutas.ps1"

$MONTH_ABBREV_BAL = @{
    "Enero"="Ene"; "Febrero"="Feb"; "Marzo"="Mar"; "Abril"="Abr"
    "Mayo"="May"; "Junio"="Jun"; "Julio"="Jul"; "Agosto"="Ago"
    "Septiembre"="Sep"; "Octubre"="Oct"; "Noviembre"="Nov"; "Diciembre"="Dic"
}

function ConvertBalanceMonth($text) {
    $t = $text -replace "`r`n|`n|`r", " " -replace "\s+", " "
    foreach ($k in $MONTH_ABBREV_BAL.Keys) { $t = $t -replace $k, $MONTH_ABBREV_BAL[$k] }
    $t = $t -replace "20(\d\d)", '$1'
    return $t.Trim()
}

function FindRow($ws, $lastRow, $label) {
    for ($r = 9; $r -le $lastRow; $r++) {
        $txt = $ws.Cells.Item($r, 1).Text.Trim()
        if ($txt -eq $label) { return $r }
    }
    return 0
}

function ReadCell($ws, $row, $col) {
    if ($row -eq 0) { return 0.0 }
    $v = $ws.Cells.Item($row, $col).Value2
    if ($v -eq $null) { return 0.0 }
    return [double]$v
}

function ReadBalanceFile($excel, $filePath) {
    Write-Host "  [BALANCE] Reading: $filePath" -ForegroundColor Gray
    if (-not (Test-Path $filePath)) {
        Write-Host "  SKIPPED - file not found" -ForegroundColor Yellow
        return @()
    }

    $wb      = $excel.Workbooks.Open($filePath)
    $ws      = $wb.Worksheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count

    $rowCash      = FindRow $ws $lastRow "Total Disponible"
    $rowReceiv    = FindRow $ws $lastRow "Total Clientes"
    $rowDebtors   = FindRow $ws $lastRow "Total Deudores"
    $rowInventory = FindRow $ws $lastRow "Total Inventarios"
    $rowTotAsset  = FindRow $ws $lastRow "Total Activo"
    $rowBankDebt  = FindRow $ws $lastRow "Total Obligaciones financieras"
    $rowSuppliers = FindRow $ws $lastRow "Total Proveedores"
    $rowAPOther   = FindRow $ws $lastRow "Total Cuentas por pagar"
    $rowTaxes     = FindRow $ws $lastRow "Total Impuestos, gravamenes y tasas"
    $rowLabor     = FindRow $ws $lastRow "Total Obligaciones laborales"
    $rowOtherLiab = FindRow $ws $lastRow "Total Otros pasivos"
    $rowTotLiab   = FindRow $ws $lastRow "Total Pasivo"
    $rowEquity    = FindRow $ws $lastRow "Total Patrimonio"

    $results = [System.Collections.Generic.List[hashtable]]::new()

    for ($c = 3; $c -le $lastCol; $c++) {
        $header = $ws.Cells.Item(8, $c).Value2
        if (-not $header) { continue }
        $mes = ConvertBalanceMonth $header
        if (-not $mes) { continue }

        $cash      = ReadCell $ws $rowCash      $c
        $cxc       = ReadCell $ws $rowReceiv    $c
        $debtors   = ReadCell $ws $rowDebtors   $c
        $inventory = ReadCell $ws $rowInventory $c
        $totAsset  = ReadCell $ws $rowTotAsset  $c
        $bankDebt  = ReadCell $ws $rowBankDebt  $c
        $suppliers = ReadCell $ws $rowSuppliers $c
        $apOther   = ReadCell $ws $rowAPOther   $c
        $taxes     = ReadCell $ws $rowTaxes     $c
        $labor     = ReadCell $ws $rowLabor     $c
        $otherLiab = ReadCell $ws $rowOtherLiab $c
        $totLiab   = ReadCell $ws $rowTotLiab   $c
        $equity    = ReadCell $ws $rowEquity    $c

        $currentAssets = $cash + $debtors + $inventory
        $currentLiab   = $suppliers + $apOther + $taxes + $labor + $otherLiab

        $results.Add(@{
            mes          = $mes
            disponible   = [long]$cash
            cxc          = [long]$cxc
            deudores     = [long]$debtors
            inventarios  = [long]$inventory
            activo_cte   = [long]$currentAssets
            total_activo = [long]$totAsset
            oblig_fin    = [long]$bankDebt
            proveedores  = [long]$suppliers
            cxp_otros    = [long]$apOther
            pasivo_cte   = [long]$currentLiab
            total_pasivo = [long]$totLiab
            patrimonio   = [long]$equity
        })
    }

    $wb.Close($false)
    Write-Host "  [BALANCE] OK - $($results.Count) months read" -ForegroundColor Green
    return $results.ToArray()
}

function ProcessAllBalances {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    $all = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($key in ($script:CONFIG.Archivos_Balance.Keys | Sort-Object)) {
        $rows = ReadBalanceFile $excel $script:CONFIG.Archivos_Balance[$key]
        foreach ($r in $rows) { $all.Add($r) }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    return $all.ToArray()
}
