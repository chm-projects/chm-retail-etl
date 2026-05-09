# ── procesar_balance.ps1 ──────────────────────────────────────────────────────
# Lee los archivos Excel del Balance General Comparativo y extrae indicadores
# mensuales de liquidez, endeudamiento y estructura del balance.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\..\config\rutas.ps1"

$MESES_ABREV_BAL = @{
    "Enero"="Ene"; "Febrero"="Feb"; "Marzo"="Mar"; "Abril"="Abr"
    "Mayo"="May"; "Junio"="Jun"; "Julio"="Jul"; "Agosto"="Ago"
    "Septiembre"="Sep"; "Octubre"="Oct"; "Noviembre"="Nov"; "Diciembre"="Dic"
}

function ConvertirMesBalance($texto) {
    $t = $texto -replace "`r`n|`n|`r", " " -replace "\s+", " "
    foreach ($k in $MESES_ABREV_BAL.Keys) { $t = $t -replace $k, $MESES_ABREV_BAL[$k] }
    $t = $t -replace "20(\d\d)", '$1'
    return $t.Trim()
}

function BuscarFila($ws, $lastRow, $etiqueta) {
    for ($r = 9; $r -le $lastRow; $r++) {
        $txt = $ws.Cells.Item($r, 1).Text.Trim()
        if ($txt -eq $etiqueta) { return $r }
    }
    return 0
}

# Funcion separada (no anidada) para leer celda numerica de Excel
function LeerCelda($ws, $fila, $col) {
    if ($fila -eq 0) { return 0.0 }
    $v = $ws.Cells.Item($fila, $col).Value2
    if ($v -eq $null) { return 0.0 }
    return [double]$v
}

function LeerArchivoBalance($excel, $rutaArchivo) {
    Write-Host "  [BALANCE] Leyendo: $rutaArchivo" -ForegroundColor Gray
    if (-not (Test-Path $rutaArchivo)) {
        Write-Host "  OMITIDO - archivo no encontrado" -ForegroundColor Yellow
        return @()
    }

    $wb      = $excel.Workbooks.Open($rutaArchivo)
    $ws      = $wb.Worksheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count

    $fDisponible  = BuscarFila $ws $lastRow "Total Disponible"
    $fClientes    = BuscarFila $ws $lastRow "Total Clientes"
    $fDeudores    = BuscarFila $ws $lastRow "Total Deudores"
    $fInventarios = BuscarFila $ws $lastRow "Total Inventarios"
    $fTotalActivo = BuscarFila $ws $lastRow "Total Activo"
    $fObligFin    = BuscarFila $ws $lastRow "Total Obligaciones financieras"
    $fProveedores = BuscarFila $ws $lastRow "Total Proveedores"
    $fCxP         = BuscarFila $ws $lastRow "Total Cuentas por pagar"
    $fImpuestos   = BuscarFila $ws $lastRow "Total Impuestos, gravamenes y tasas"
    $fObligLab    = BuscarFila $ws $lastRow "Total Obligaciones laborales"
    $fOtrosPas    = BuscarFila $ws $lastRow "Total Otros pasivos"
    $fTotalPasivo = BuscarFila $ws $lastRow "Total Pasivo"
    $fPatrimonio  = BuscarFila $ws $lastRow "Total Patrimonio"

    # Si no encontro con nombre sin tilde, intentar con tilde
    if ($fImpuestos -eq 0) {
        $fImpuestos = BuscarFila $ws $lastRow "Total Impuestos, gravamenes y tasas"
    }

    $resultados = [System.Collections.Generic.List[hashtable]]::new()

    for ($c = 3; $c -le $lastCol; $c++) {
        $encabezado = $ws.Cells.Item(8, $c).Value2
        if (-not $encabezado) { continue }
        $mes = ConvertirMesBalance $encabezado
        if (-not $mes) { continue }

        $disponible   = LeerCelda $ws $fDisponible  $c
        $cxc          = LeerCelda $ws $fClientes    $c
        $deudores     = LeerCelda $ws $fDeudores    $c
        $inventarios  = LeerCelda $ws $fInventarios $c
        $total_activo = LeerCelda $ws $fTotalActivo $c
        $oblig_fin    = LeerCelda $ws $fObligFin    $c
        $proveedores  = LeerCelda $ws $fProveedores $c
        $cxp_otros    = LeerCelda $ws $fCxP         $c
        $impuestos    = LeerCelda $ws $fImpuestos   $c
        $oblig_lab    = LeerCelda $ws $fObligLab    $c
        $otros_pas    = LeerCelda $ws $fOtrosPas    $c
        $total_pasivo = LeerCelda $ws $fTotalPasivo $c
        $patrimonio   = LeerCelda $ws $fPatrimonio  $c

        $activo_cte = $disponible + $deudores + $inventarios
        $pasivo_cte = $proveedores + $cxp_otros + $impuestos + $oblig_lab + $otros_pas

        $resultados.Add(@{
            mes           = $mes
            disponible    = [long]$disponible
            cxc           = [long]$cxc
            deudores      = [long]$deudores
            inventarios   = [long]$inventarios
            activo_cte    = [long]$activo_cte
            total_activo  = [long]$total_activo
            oblig_fin     = [long]$oblig_fin
            proveedores   = [long]$proveedores
            cxp_otros     = [long]$cxp_otros
            pasivo_cte    = [long]$pasivo_cte
            total_pasivo  = [long]$total_pasivo
            patrimonio    = [long]$patrimonio
        })
    }

    $wb.Close($false)
    Write-Host "  [BALANCE] OK - $($resultados.Count) meses leidos" -ForegroundColor Green
    return $resultados.ToArray()
}

function ProcesarTodosLosBalances {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false

    $todos = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($key in ($script:CONFIG.Archivos_Balance.Keys | Sort-Object)) {
        $datos = LeerArchivoBalance $excel $script:CONFIG.Archivos_Balance[$key]
        foreach ($d in $datos) { $todos.Add($d) }
    }

    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    return $todos.ToArray()
}
