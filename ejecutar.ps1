# ── ejecutar.ps1 ─────────────────────────────────────────────────────────────
# Script principal. Doble clic o ejecutar en PowerShell para actualizar el
# dashboard con los datos más recientes del P&G y el Balance General.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\scripts\procesar_pyg.ps1"
. "$PSScriptRoot\scripts\procesar_balance.ps1"

Write-Host "═══════════════════════════════════════════════" -ForegroundColor DarkCyan
Write-Host "  retail-etl — Alternativas El Descuento       " -ForegroundColor Cyan
Write-Host "  Generador de datos para el dashboard BI      " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════`n" -ForegroundColor DarkCyan

# ── 1. P&G ──────────────────────────────────────────────────────────────────
Write-Host "Procesando Estado de Resultados..." -ForegroundColor Cyan
$datos = ProcesarTodosLosPyG

if ($datos.Count -eq 0) {
    Write-Host "`nERROR: No se pudieron leer datos del P&G. Verifica config\rutas.ps1" -ForegroundColor Red
    exit 1
}

$lineasJs = [System.Collections.Generic.List[string]]::new()
foreach ($d in $datos) {
    $linea = "    { periodo:`"$($d.periodo)`", mes:`"$($d.mes)`", " +
             "ventas:$($d.ventas), costo_ventas:$($d.costo_ventas), " +
             "margen_pesos:$($d.margen_pesos), margen_pct:$($d.margen_pct), " +
             "personal:$($d.personal), arriendo:$($d.arriendo), " +
             "servicios:$($d.servicios), honorarios:$($d.honorarios), " +
             "impuestos:$($d.impuestos), gastos_ventas:$($d.gastos_ventas), " +
             "total_gastos:$($d.total_gastos), resultado:$($d.resultado) }"
    $lineasJs.Add($linea)
}

$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
$cuerpo    = $lineasJs -join ",`n"

$contenidoPyG = @"
// Generado automaticamente por retail-etl/ejecutar.ps1
// Fuente: Estado de Resultados Comparativo Mes a Mes (2024-2026)
// Ultima actualizacion: $timestamp
// NO editar manualmente

window.DASHBOARD_DATA = {
  metadata: {
    generado: "$timestamp",
    fuente: "Estado de Resultados Comparativo Mes a Mes",
    meses_procesados: $($datos.Count)
  },
  ventas_mensual: [
$cuerpo
  ]
};
"@

$outPath = $script:CONFIG.Output_Dashboard
$outDir  = Split-Path $outPath
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$contenidoPyG | Out-File -FilePath $outPath -Encoding UTF8 -Force
Write-Host "  P&G OK — $($datos.Count) meses → $outPath" -ForegroundColor Green

# ── 2. Balance General ───────────────────────────────────────────────────────
Write-Host "`nProcesando Balance General..." -ForegroundColor Cyan
$balance = ProcesarTodosLosBalances

if ($balance.Count -eq 0) {
    Write-Host "  ADVERTENCIA: No se pudo leer el Balance General." -ForegroundColor Yellow
} else {
    $lineasBal = [System.Collections.Generic.List[string]]::new()
    foreach ($d in $balance) {
        $linea = "    { mes:`"$($d.mes)`", " +
                 "disponible:$($d.disponible), cxc:$($d.cxc), deudores:$($d.deudores), " +
                 "inventarios:$($d.inventarios), activo_cte:$($d.activo_cte), total_activo:$($d.total_activo), " +
                 "oblig_fin:$($d.oblig_fin), proveedores:$($d.proveedores), cxp_otros:$($d.cxp_otros), " +
                 "pasivo_cte:$($d.pasivo_cte), total_pasivo:$($d.total_pasivo), patrimonio:$($d.patrimonio) }"
        $lineasBal.Add($linea)
    }

    $cuerpoBal = $lineasBal -join ",`n"
    $contenidoBal = @"
// Generado automaticamente por retail-etl/ejecutar.ps1
// Fuente: Balance General Comparativo (2024-2026)
// Ultima actualizacion: $timestamp
// NO editar manualmente

window.BALANCE_DATA = {
  metadata: {
    generado: "$timestamp",
    fuente: "Balance General Comparativo",
    meses_procesados: $($balance.Count)
  },
  mensual: [
$cuerpoBal
  ]
};
"@

    $outBal = $script:CONFIG.Output_Balance_Dashboard
    $contenidoBal | Out-File -FilePath $outBal -Encoding UTF8 -Force
    Write-Host "  Balance OK — $($balance.Count) meses → $outBal" -ForegroundColor Green
}

# ── Resumen final ─────────────────────────────────────────────────────────────
Write-Host "`n=== RESUMEN ===" -ForegroundColor Yellow
Write-Host "P&G procesado    : $($datos.Count) meses"
Write-Host "Balance procesado: $($balance.Count) meses"
Write-Host "`nAbre el dashboard:"
Write-Host "  G:\Mi unidad\github\retail-cm-ia\dashboard\mando-integral-bi.html`n"
