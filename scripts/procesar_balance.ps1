# ── procesar_balance.ps1 ──────────────────────────────────────────────────────
# Lee los archivos Excel del Balance General Comparativo y extrae indicadores
# mensuales de liquidez, endeudamiento y estructura del balance.
# Retorna una lista de hashtables, una por mes.
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
    # Busca en columna A la etiqueta exacta; retorna número de fila o 0
    for ($r = 9; $r -le $lastRow; $r++) {
        $txt = $ws.Cells.Item($r, 1).Text.Trim()
        if ($txt -eq $etiqueta) { return $r }
    }
    return 0
}

function LeerArchivoBalance($excel, $rutaArchivo) {
    Write-Host "  Balance: $rutaArchivo" -ForegroundColor Gray
    if (-not (Test-Path $rutaArchivo)) {
        Write-Host "  OMITIDO — archivo no encontrado" -ForegroundColor Yellow
        return @()
    }

    $wb      = $excel.Workbooks.Open($rutaArchivo)
    $ws      = $wb.Worksheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count

    # Mapear etiquetas clave → fila (una sola búsqueda por archivo)
    $filas = @{
        Disponible    = BuscarFila $ws $lastRow "Total Disponible"
        Clientes      = BuscarFila $ws $lastRow "Total Clientes"
        Deudores      = BuscarFila $ws $lastRow "Total Deudores"
        Inventarios   = BuscarFila $ws $lastRow "Total Inventarios"
        TotalActivo   = BuscarFila $ws $lastRow "Total Activo"
        ObligFin      = BuscarFila $ws $lastRow "Total Obligaciones financieras"
        Proveedores   = BuscarFila $ws $lastRow "Total Proveedores"
        CxP           = BuscarFila $ws $lastRow "Total Cuentas por pagar"
        Impuestos     = BuscarFila $ws $lastRow "Total Impuestos, gravámenes y tasas"
        ObligLab      = BuscarFila $ws $lastRow "Total Obligaciones laborales"
        OtrosPas      = BuscarFila $ws $lastRow "Total Otros pasivos"
        TotalPasivo   = BuscarFila $ws $lastRow "Total Pasivo"
        Patrimonio    = BuscarFila $ws $lastRow "Total Patrimonio"
    }

    function Valor($fila, $col) {
        if ($fila -eq 0) { return 0.0 }
        $v = $ws.Cells.Item($fila, $col).Value2
        if ($v -eq $null) { return 0.0 }
        return [double]$v
    }

    $resultados = [System.Collections.Generic.List[hashtable]]::new()

    # Columnas de datos: col 3 en adelante
    for ($c = 3; $c -le $lastCol; $c++) {
        $encabezado = $ws.Cells.Item(8, $c).Value2
        if (-not $encabezado) { continue }
        $mes = ConvertirMesBalance $encabezado
        if (-not $mes) { continue }

        $disponible    = Valor $filas.Disponible  $c
        $cxc           = Valor $filas.Clientes    $c
        $deudores      = Valor $filas.Deudores    $c
        $inventarios   = Valor $filas.Inventarios $c
        $total_activo  = Valor $filas.TotalActivo $c
        $oblig_fin     = Valor $filas.ObligFin    $c
        $proveedores   = Valor $filas.Proveedores $c
        $cxp_otros     = Valor $filas.CxP         $c
        $impuestos     = Valor $filas.Impuestos   $c
        $oblig_lab     = Valor $filas.ObligLab    $c
        $otros_pas     = Valor $filas.OtrosPas    $c
        $total_pasivo  = Valor $filas.TotalPasivo $c
        $patrimonio    = Valor $filas.Patrimonio  $c

        # Activo corriente = efectivo + deudores + inventarios
        $activo_cte = $disponible + $deudores + $inventarios

        # Pasivo corriente operativo (excluye deuda financiera bancaria)
        $pasivo_cte = $proveedores + $cxp_otros + $impuestos + $oblig_lab + $otros_pas

        $resultados.Add(@{
            mes              = $mes
            disponible       = [long]$disponible
            cxc              = [long]$cxc
            deudores         = [long]$deudores
            inventarios      = [long]$inventarios
            activo_cte       = [long]$activo_cte
            total_activo     = [long]$total_activo
            oblig_fin        = [long]$oblig_fin
            proveedores      = [long]$proveedores
            cxp_otros        = [long]$cxp_otros
            pasivo_cte       = [long]$pasivo_cte
            total_pasivo     = [long]$total_pasivo
            patrimonio       = [long]$patrimonio
        })
    }

    $wb.Close($false)
    return $resultados
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
    return $todos
}
