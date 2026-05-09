# ── Configuración de rutas ──────────────────────────────────────────────────
# Editar estas rutas si cambia la ubicación de los archivos en tu PC.

$script:CONFIG = @{

    # Carpeta donde están los archivos Excel del P&G
    DataAnalytics = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics"

    # Archivos del Estado de Resultados (uno por año)
    Archivos_PyG = @{
        "2024" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2024.xlsx"
        "2025" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2025.xlsx"
        "2026" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2026.xlsx"
    }

    # Archivo de Costo de Ventas Detallado (para el Top 50 de productos)
    Costo_Ventas_Detallado = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Costo Ventas Detallado_20260509_170131.xlsx"

    # Destino del archivo de datos para el dashboard
    Output_Dashboard = "G:\Mi unidad\github\retail-cm-ia\dashboard\data\data.js"
}
