# ── procesar_pyg.ps1 ─────────────────────────────────────────────────────────
# Lee los 3 archivos Excel del Estado de Resultados y extrae métricas clave.
# Retorna un array de hashtables, una por mes, con todos los indicadores.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\..\config\rutas.ps1"

$MESES_ABREV = @{
    "Enero"="Ene"; "Febrero"="Feb"; "Marzo"="Mar"; "Abril"="Abr"
    "Mayo"="May"; "Junio"="Jun"; "Julio"="Jul"; "Agosto"="Ago"
    "Septiembre"="Sep"; "Octubre"="Oct"; "Noviembre"="Nov"; "Diciembre"="Dic"
}

function ConvertirMes($texto) {
    # "Junio`n2024" → "Jun 24"
    $t = $texto -replace "`r`n|`n|`r", " " -replace "\s+", " "
    foreach ($k in $MESES_ABREV.Keys) { $t = $t -replace $k, $MESES_ABREV[$k] }
    $t = $t -replace "20(\d\d)", '$1'   # "Jun 2024" → "Jun 24"
    return $t.Trim()
}

function ObtenerFilasPorNombre($ws, $lastRow) {
    # Recolecta TODAS las ocurrencias de cada etiqueta clave en col A
    $ocurrencias = @{}
    for ($r = 1; $r -le $lastRow; $r++) {
        $txt = $ws.Cells.Item($r, 1).Text.Trim()
        if ($txt -eq "") { continue }
        if (-not $ocurrencias.ContainsKey($txt)) {
            $ocurrencias[$txt] = [System.Collections.Generic.List[int]]::new()
        }
        $ocurrencias[$txt].Add($r)
    }
    return $ocurrencias
}

function LeerArchivoPyG($excel, $rutaArchivo, $periodo) {
    if (-not (Test-Path $rutaArchivo)) {
        Write-Warning "Archivo no encontrado: $rutaArchivo"
        return @()
    }
    Write-Host "  Leyendo $periodo..." -ForegroundColor Cyan
    $wb = $excel.Workbooks.Open($rutaArchivo)
    $ws = $wb.Sheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count

    # ── Mapeo de columnas (fila 8 = encabezados de mes) ──────────────────────
    $columnasMes = [ordered]@{}
    for ($col = 1; $col -le 30; $col++) {
        $header = $ws.Cells.Item(8, $col).Text.Trim()
        if ($header -eq "" -or $header -match "Total|Código|Nombre") { continue }
        $label = ConvertirMes $header
        if ($label -ne "") { $columnasMes[$col] = $label }
    }

    # ── Localizar filas clave por etiqueta ───────────────────────────────────
    $oc = ObtenerFilasPorNombre $ws $lastRow

    # Ingresos: primera ocurrencia de "Total Operacionales" = ventas retail
    $fVentas      = if ($oc["Total Operacionales"])           { $oc["Total Operacionales"][0] }           else { 0 }
    $fIngresos    = if ($oc["Total Ingresos"])                { $oc["Total Ingresos"][0] }                else { 0 }
    # Costos fijos: ÚLTIMA ocurrencia de cada etiqueta (sección 52-Ventas)
    $fPersonal    = if ($oc["Total Gastos de personal"])      { $oc["Total Gastos de personal"][-1] }     else { 0 }
    $fArriendo    = if ($oc["Total Arrendamientos"])          { $oc["Total Arrendamientos"][-1] }         else { 0 }
    $fServicios   = if ($oc["Total Servicios"])               { $oc["Total Servicios"][-1] }              else { 0 }
    $fHonorarios  = if ($oc["Total Honorarios"])              { $oc["Total Honorarios"][-1] }             else { 0 }
    $fImpuestos   = if ($oc["Total Impuestos"])               { $oc["Total Impuestos"][-1] }              else { 0 }
    $fDiversos    = if ($oc["Total Diversos"])                { $oc["Total Diversos"][-1] }               else { 0 }
    $fGastosVtas  = if ($oc["Total Operacionales de ventas"]) { $oc["Total Operacionales de ventas"][0] } else { 0 }
    $fTotalGastos = if ($oc["Total Gastos"])                  { $oc["Total Gastos"][0] }                  else { 0 }
    $fCVentas     = if ($oc["Total Costos de ventas"])        { $oc["Total Costos de ventas"][0] }        else { 0 }
    $fResultado   = if ($oc["Resultado del Ejercicio"])       { $oc["Resultado del Ejercicio"][0] }       else { 0 }

    # ── Extraer valores por mes ───────────────────────────────────────────────
    $datos = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($col in $columnasMes.Keys) {
        $mes = $columnasMes[$col]

        function Val($fila) {
            if ($fila -eq 0) { return 0 }
            $v = $ws.Cells.Item($fila, $col).Value2
            if ($null -eq $v) { return 0 }
            return [math]::Round($v)
        }

        $ventas      = Val $fVentas
        $ingresos    = Val $fIngresos
        $cVentas     = Val $fCVentas
        $personal    = Val $fPersonal
        $arriendo    = Val $fArriendo
        $servicios   = Val $fServicios
        $honorarios  = Val $fHonorarios
        $impuestos   = Val $fImpuestos
        $diversos    = Val $fDiversos
        $gastosVtas  = Val $fGastosVtas
        $totalGastos = Val $fTotalGastos
        $resultado   = Val $fResultado

        $margenBruto = if ($ventas -ne 0) { [math]::Round((($ventas - $cVentas) / $ventas) * 100, 2) } else { 0 }

        $datos.Add(@{
            periodo      = $periodo
            mes          = $mes
            ventas       = $ventas
            ingresos     = $ingresos
            costo_ventas = $cVentas
            margen_pesos = $ventas - $cVentas
            margen_pct   = $margenBruto
            personal     = $personal
            arriendo     = $arriendo
            servicios    = $servicios
            honorarios   = $honorarios
            impuestos    = $impuestos
            diversos     = $diversos
            gastos_ventas = $gastosVtas
            total_gastos = $totalGastos
            resultado    = $resultado
        })
    }

    $wb.Close($false)
    Write-Host "    → $($datos.Count) meses leídos" -ForegroundColor Green
    return $datos
}

function ProcesarTodosLosPyG {
    Write-Host "`nIniciando lectura de P&G..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $todosDatos = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($año in @("2024","2025","2026")) {
        $ruta = $script:CONFIG.Archivos_PyG[$año]
        $meses = LeerArchivoPyG $excel $ruta $año
        foreach ($m in $meses) { $todosDatos.Add($m) }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "`nTotal de meses procesados: $($todosDatos.Count)" -ForegroundColor Green
    return $todosDatos
}
