# retail-etl — ARCHIVADO

> **Este repositorio esta archivado y reemplazado por [chm-retail-dwh](https://github.com/chm-projects/chm-retail-dwh).**
>
> El nuevo pipeline produce output identico (byte-equivalente) con arquitectura declarativa (dbt + DuckDB + Python), tests automaticos y deteccion de drift. Para el flujo mensual ver [docs/migration.md](https://github.com/chm-projects/chm-retail-dwh/blob/master/docs/migration.md) en el nuevo repo.
>
> Este repo se conserva como referencia historica. No ejecutar `ejecutar.ps1` si retail-dwh esta activo.

---

**Alternativas El Descuento** — Pipeline de Procesamiento de Datos (LEGADO)
Yarumal, Antioquia · ETL: Excel → JS → Dashboard BI

---

## ¿Qué hace?

Lee los archivos Excel del P&G (Estado de Resultados) y del Balance General directamente desde la carpeta de Data Analytics y genera los archivos `data.js` y `data_balance.js` que alimentan el dashboard BI.

```
Excel P&G (2024 / 2025 / 2026)          Excel Balance (2024-25 / 2025-26)
         │                                          │
         ▼                                          ▼
  procesar_pyg.ps1                        procesar_balance.ps1
  ProcessAllPyG()                         ProcessAllBalances()
  COM Excel → array hashtable/mes         COM Excel → array hashtable/mes
         │                                          │
         └─────────────────┬────────────────────────┘
                           ▼
                     ejecutar.ps1
                 Construye JS sin BOM
                           │
            ┌──────────────┴──────────────┐
            ▼                             ▼
    data/data.js                 data/data_balance.js
    window.DASHBOARD_DATA        window.BALANCE_DATA
            │                             │
            └──────────────┬──────────────┘
                           ▼
             mando-integral-bi.html
             (cargado por retail-cm-ia)
```

---

## Requisitos

- Windows 10/11 con PowerShell 5.1 (ya incluido)
- Microsoft Excel instalado (lectura de .xlsx via COM)
- Los archivos Excel en la ruta configurada en `config/rutas.ps1`

---

## Cómo usar

### Opción 1 — Botón en el dashboard (recomendado)
1. Abrir `retail-cm-ia/INICIAR DASHBOARD.bat` (doble clic)
2. En el dashboard presionar **Actualizar datos** (botón fucsia, esquina superior)
3. La barra de progreso avanza en tiempo real hasta completar

### Opción 2 — Terminal directa
```powershell
powershell -ExecutionPolicy Bypass -File "G:\Mi unidad\github\retail-etl\ejecutar.ps1"
```

### Ejecutar tests
```powershell
powershell -ExecutionPolicy Bypass -Command "Invoke-Pester 'G:\Mi unidad\github\retail-etl\tests\ETL.Tests.ps1'"
# Resultado esperado: Passed: 30, Failed: 0
```

---

## Estructura

```
retail-etl/
├── ejecutar.ps1               # Runner principal del ETL (pasos 1-5)
├── config/
│   └── rutas.ps1              # ÚNICO ARCHIVO A EDITAR si cambian rutas
├── scripts/
│   ├── procesar_pyg.ps1       # Lee Excel P&G → array mensual
│   └── procesar_balance.ps1   # Lee Excel Balance → array mensual
├── tests/
│   └── ETL.Tests.ps1          # 30 tests Pester (unit + integración)
└── .claude/
    └── settings.json          # Permisos y acceso a retail-cm-ia
```

---

## Configurar rutas

Si los archivos Excel cambian de nombre o ubicación, editar **solo** `config/rutas.ps1`:

```powershell
$script:CONFIG = @{
    Archivos_PyG = @{
        "2024" = "ruta\al\P&G_2024.xlsx"
        "2025" = "ruta\al\P&G_2025.xlsx"
        "2026" = "ruta\al\P&G_2026.xlsx"
    }
    Archivos_Balance = @{
        "2024-2025" = "ruta\al\Balance_24-25.xlsx"
        "2025-2026" = "ruta\al\Balance_25-26.xlsx"
    }
    Output_Dashboard         = "ruta\retail-cm-ia\dashboard\data\data.js"
    Output_Balance_Dashboard = "ruta\retail-cm-ia\dashboard\data\data_balance.js"
}
```

---

## Datos que extrae

### P&G — por mes (`data.js`)

| Campo | Fuente en Excel | Uso en dashboard |
|---|---|---|
| `ventas` | Total Operacionales | Gráfica tendencia, KPI ventas |
| `costo_ventas` | Total Costos de ventas | Cálculo margen bruto |
| `margen_pesos` | ventas − costo_ventas | KPI margen en pesos |
| `margen_pct` | (margen / ventas) × 100 | Gráfica Margen Bruto % |
| `personal` | Total Gastos de personal | Tab Costos Fijos |
| `arriendo` | Total Arrendamientos | Tab Costos Fijos |
| `servicios` | Total Servicios | Tab Costos Fijos |
| `honorarios` | Total Honorarios | Tab Costos Fijos |
| `total_gastos` | Total Gastos | Análisis cobertura |
| `resultado` | Resultado del Ejercicio | KPI Utilidad Neta |

### Balance General — por mes (`data_balance.js`)

| Campo | Fuente en Excel | Uso en dashboard |
|---|---|---|
| `disponible` | Total Disponible | Liquidez |
| `cxc` | Total Clientes | Rotación CxC |
| `inventarios` | Total Inventarios | Rotación inventario |
| `activo_cte` | Calculado: caja+deudores+inv | Capital de trabajo |
| `total_activo` | Total Activo | Endeudamiento |
| `oblig_fin` | Total Obligaciones financieras | Deuda bancaria |
| `proveedores` | Total Proveedores | Rotación CxP |
| `pasivo_cte` | Calculado: proveedores+CxP+imp+lab+otros | Capital de trabajo |
| `total_pasivo` | Total Pasivo | Endeudamiento |
| `patrimonio` | Total Patrimonio | Estructura financiera |

---

## Reglas técnicas (aprendidas en producción)

| Regla | Razón |
|---|---|
| Código en inglés ASCII-only | `ñ`/tildes en variables → mojibake en proceso hijo PS5.1 |
| `@{}` con claves string, no int | `[ordered]@{}[$col]` con int → `ArgumentOutOfRangeException` |
| Funciones a nivel de script | Funciones anidadas en loops pierden scope en PS5.1 |
| `UTF8Encoding($false)` para escribir JS | `[System.Text.Encoding]::UTF8` incluye BOM |
| Sin heredocs en bloques condicionales | `@"..."@` dentro de `else{}` → `AmpersandNotAllowed` |

---

## Relación con retail-cm-ia

| Repositorio | Responsabilidad |
|---|---|
| **retail-etl** (este) | Extracción y transformación de datos Excel |
| **retail-cm-ia** | Dashboard, conocimientos de negocio, habilidades IA |

Los archivos `data.js` y `data_balance.js` son generados; están en `.gitignore` de `retail-cm-ia` (datos financieros, no se versionan).
