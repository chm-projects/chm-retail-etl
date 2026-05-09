# ── ejecutar.ps1 ─────────────────────────────────────────────────────────────
# Script principal. Doble clic o ejecutar en PowerShell para actualizar el
# dashboard con los datos más recientes del P&G.
# ─────────────────────────────────────────────────────────────────────────────

. "$PSScriptRoot\scripts\procesar_pyg.ps1"

Write-Host "═══════════════════════════════════════════════" -ForegroundColor DarkCyan
Write-Host "  retail-etl — Alternativas El Descuento       " -ForegroundColor Cyan
Write-Host "  Generador de datos para el dashboard BI      " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════`n" -ForegroundColor DarkCyan

# 1. Leer todos los archivos P&G
$datos = ProcesarTodosLosPyG

if ($datos.Count -eq 0) {
    Write-Host "`nERROR: No se pudieron leer datos. Verifica las rutas en config\rutas.ps1" -ForegroundColor Red
    exit 1
}

# 2. Construir líneas JS para ventas_mensual
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

# 3. Generar el archivo data.js
$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
$cuerpo    = $lineasJs -join ",`n"

$contenido = @"
// Generado automaticamente por retail-etl/ejecutar.ps1
// Fuente: Estado de Resultados Comparativo Mes a Mes (2024-2026)
// Ultima actualizacion: $timestamp
// NO editar manualmente — ejecutar ejecutar.ps1 para actualizar

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

# 4. Escribir al destino
$outPath = $script:CONFIG.Output_Dashboard
$outDir  = Split-Path $outPath
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}
$contenido | Out-File -FilePath $outPath -Encoding UTF8 -Force

Write-Host "`n=== RESUMEN ===" -ForegroundColor Yellow
Write-Host "Meses procesados : $($datos.Count)"
Write-Host "Archivo generado : $outPath"
Write-Host "`nAbre el dashboard en tu navegador para ver los datos actualizados."
Write-Host "  $("G:\Mi unidad\github\retail-cm-ia\dashboard\mando-integral-bi.html")`n"
