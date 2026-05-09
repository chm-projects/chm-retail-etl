# ── ejecutar.ps1 ─────────────────────────────────────────────────────────────
# Script principal. Doble clic o ejecutar desde el dashboard para actualizar
# los datos del P&G y el Balance General.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\scripts\procesar_pyg.ps1"
. "$PSScriptRoot\scripts\procesar_balance.ps1"

Write-Host "=== retail-etl Alternativas El Descuento ===" -ForegroundColor Cyan

# ── PASO 1: Leer P&G ─────────────────────────────────────────────────────────
Write-Host "[PASO 1/5] Procesando Estado de Resultados..." -ForegroundColor Cyan
$datos = ProcesarTodosLosPyG

if ($datos.Count -eq 0) {
    Write-Host "[ERROR] No se pudieron leer datos del P&G. Verifica config/rutas.ps1" -ForegroundColor Red
    exit 1
}

Write-Host "[PASO 2/5] Construyendo datos del P&G ($($datos.Count) meses)..." -ForegroundColor Cyan

$lineasJs = [System.Collections.Generic.List[string]]::new()
foreach ($d in $datos) {
    $lineasJs.Add(
        "    { periodo:`"$($d.periodo)`", mes:`"$($d.mes)`"," +
        " ventas:$($d.ventas), costo_ventas:$($d.costo_ventas)," +
        " margen_pesos:$($d.margen_pesos), margen_pct:$($d.margen_pct)," +
        " personal:$($d.personal), arriendo:$($d.arriendo)," +
        " servicios:$($d.servicios), honorarios:$($d.honorarios)," +
        " impuestos:$($d.impuestos), gastos_ventas:$($d.gastos_ventas)," +
        " total_gastos:$($d.total_gastos), resultado:$($d.resultado) }"
    )
}

$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
$cuerpo    = $lineasJs -join ",`n"

# Construir data.js sin heredoc dentro de bloque condicional
$jsLineas = [System.Collections.Generic.List[string]]::new()
$jsLineas.Add("// Generado automaticamente por retail-etl/ejecutar.ps1")
$jsLineas.Add("// Fuente: Estado de Resultados Comparativo Mes a Mes (2024-2026)")
$jsLineas.Add("// Ultima actualizacion: $timestamp")
$jsLineas.Add("// NO editar manualmente")
$jsLineas.Add("")
$jsLineas.Add("window.DASHBOARD_DATA = {")
$jsLineas.Add("  metadata: {")
$jsLineas.Add("    generado: `"$timestamp`",")
$jsLineas.Add("    fuente: `"Estado de Resultados Comparativo Mes a Mes`",")
$jsLineas.Add("    meses_procesados: $($datos.Count)")
$jsLineas.Add("  },")
$jsLineas.Add("  ventas_mensual: [")
$jsLineas.Add($cuerpo)
$jsLineas.Add("  ]")
$jsLineas.Add("};")
$contenidoPyG = $jsLineas -join [System.Environment]::NewLine

Write-Host "[PASO 3/5] Escribiendo data.js..." -ForegroundColor Cyan
$outPath = $script:CONFIG.Output_Dashboard
$outDir  = Split-Path $outPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($outPath, $contenidoPyG, [System.Text.Encoding]::UTF8)
Write-Host "  P&G OK - $($datos.Count) meses -> $outPath" -ForegroundColor Green

# ── PASO 4: Leer Balance General ─────────────────────────────────────────────
Write-Host "[PASO 4/5] Procesando Balance General..." -ForegroundColor Cyan
$balance = ProcesarTodosLosBalances

if ($balance.Count -eq 0) {
    Write-Host "  ADVERTENCIA: No se pudo leer el Balance General." -ForegroundColor Yellow
} else {
    $lineasBal = [System.Collections.Generic.List[string]]::new()
    foreach ($d in $balance) {
        $lineasBal.Add(
            "    { mes:`"$($d.mes)`"," +
            " disponible:$($d.disponible), cxc:$($d.cxc), deudores:$($d.deudores)," +
            " inventarios:$($d.inventarios), activo_cte:$($d.activo_cte), total_activo:$($d.total_activo)," +
            " oblig_fin:$($d.oblig_fin), proveedores:$($d.proveedores), cxp_otros:$($d.cxp_otros)," +
            " pasivo_cte:$($d.pasivo_cte), total_pasivo:$($d.total_pasivo), patrimonio:$($d.patrimonio) }"
        )
    }
    $cuerpoBal = $lineasBal -join ",`n"

    Write-Host "[PASO 5/5] Escribiendo data_balance.js..." -ForegroundColor Cyan

    $balLineas = [System.Collections.Generic.List[string]]::new()
    $balLineas.Add("// Generado automaticamente por retail-etl/ejecutar.ps1")
    $balLineas.Add("// Fuente: Balance General Comparativo (2024-2026)")
    $balLineas.Add("// Ultima actualizacion: $timestamp")
    $balLineas.Add("// NO editar manualmente")
    $balLineas.Add("")
    $balLineas.Add("window.BALANCE_DATA = {")
    $balLineas.Add("  metadata: {")
    $balLineas.Add("    generado: `"$timestamp`",")
    $balLineas.Add("    fuente: `"Balance General Comparativo`",")
    $balLineas.Add("    meses_procesados: $($balance.Count)")
    $balLineas.Add("  },")
    $balLineas.Add("  mensual: [")
    $balLineas.Add($cuerpoBal)
    $balLineas.Add("  ]")
    $balLineas.Add("};")
    $contenidoBal = $balLineas -join [System.Environment]::NewLine

    $outBal = $script:CONFIG.Output_Balance_Dashboard
    [System.IO.File]::WriteAllText($outBal, $contenidoBal, [System.Text.Encoding]::UTF8)
    Write-Host "  Balance OK - $($balance.Count) meses -> $outBal" -ForegroundColor Green
}

Write-Host "[COMPLETADO] P&G=$($datos.Count) meses. Balance=$($balance.Count) meses." -ForegroundColor Green
