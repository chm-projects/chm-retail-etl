# ── rutas.ps1 ────────────────────────────────────────────────────────────────
# Edit these paths if the Excel files are moved on this machine.

$script:CONFIG = @{

    DataAnalytics = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics"

    # P&G (Income Statement) - one file per year
    Archivos_PyG = @{
        "2024" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2024.xlsx"
        "2025" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2025.xlsx"
        "2026" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Estado de Resultados Comparativo Mes a Mes 2026.xlsx"
    }

    # Detailed cost of goods (for Top-50 product analysis)
    Costo_Ventas_Detallado = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Costo Ventas Detallado_20260509_170131.xlsx"

    # Balance Sheet - one file per date range
    Archivos_Balance = @{
        "2024-2025" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Balance General Comparativo_20260509_174609.xlsx"
        "2025-2026" = "G:\Mi unidad\ALTERNATIVAS EL DESCUENTO\Data Analytics\Balance General Comparativo_20260509_174656.xlsx"
    }

    # Dashboard output paths
    Output_Dashboard         = "G:\Mi unidad\github\retail-cm-ia\dashboard\data\data.js"
    Output_Balance_Dashboard = "G:\Mi unidad\github\retail-cm-ia\dashboard\data\data_balance.js"
}
