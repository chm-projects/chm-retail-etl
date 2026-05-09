# retail-etl
**Alternativas El Descuento** — Pipeline de Procesamiento de Datos  
Yarumal, Antioquia · ETL: Excel → JSON → Dashboard BI

> Este repositorio es el motor de datos del proyecto.  
> El repositorio de conocimiento, habilidades IA y el dashboard viven en **[retail-cm-ia](https://github.com/sachavar/retail-cm-ia)**.

---

## ¿Qué hace este repositorio?

Lee los archivos Excel del P&G (Estado de Resultados) directamente desde la carpeta de Data Analytics y genera el archivo `data.js` que alimenta el dashboard BI de forma automática. **Elimina la necesidad de editar el dashboard manualmente cada vez que hay datos nuevos.**

```
[Excel P&G 2024/2025/2026]
         ↓
   ejecutar.ps1          ← Doble clic para actualizar
         ↓
   dashboard/data/data.js  ← Generado automáticamente
         ↓
   mando-integral-bi.html  ← Lee data.js y muestra los datos
```

---

## Requisitos

- Windows con PowerShell (ya incluido en Windows 10/11)
- Microsoft Excel instalado (para leer los archivos .xlsx via COM)
- Los archivos de Data Analytics en la ruta configurada

---

## Cómo Usar

**Opción 1 — Doble clic:**
1. Clic derecho en `ejecutar.ps1`
2. Seleccionar "Ejecutar con PowerShell"
3. Abrir `retail-cm-ia/dashboard/mando-integral-bi.html` en el navegador

**Opción 2 — PowerShell:**
```powershell
cd "G:\Mi unidad\github\retail-etl"
.\ejecutar.ps1
```

---

## Estructura

```
retail-etl/
├── ejecutar.ps1           ← Runner principal (doble clic aquí)
├── config/
│   └── rutas.ps1          ← Configurar rutas si cambian los archivos
├── scripts/
│   └── procesar_pyg.ps1   ← Lógica de extracción del P&G
└── README.md
```

---

## Configurar Rutas

Si los archivos Excel cambian de nombre o ubicación, editar `config/rutas.ps1`:

```powershell
$script:CONFIG = @{
    Archivos_PyG = @{
        "2024" = "ruta\al\archivo_2024.xlsx"
        "2025" = "ruta\al\archivo_2025.xlsx"
        "2026" = "ruta\al\archivo_2026.xlsx"
    }
    Output_Dashboard = "G:\...\retail-cm-ia\dashboard\data\data.js"
}
```

---

## Datos que Extrae del P&G

| Campo | Fuente en el P&G | Uso en Dashboard |
|---|---|---|
| `ventas` | Total Operacionales (ingresos retail) | Gráfica de tendencia |
| `costo_ventas` | Total Costos de ventas | Margen bruto |
| `margen_pct` | (ventas - costo_ventas) / ventas | KPI de margen |
| `personal` | Total Gastos de personal (sección 52) | Costos fijos |
| `arriendo` | Total Arrendamientos (sección 52) | Costos fijos |
| `servicios` | Total Servicios (sección 52) | Costos fijos |
| `honorarios` | Total Honorarios (sección 52) | Costos fijos |
| `resultado` | Resultado del Ejercicio | Utilidad neta |

---

## Relación con retail-cm-ia

| Repositorio | Responsabilidad |
|---|---|
| **retail-cm-ia** | Conocimiento, habilidades IA, dashboard HTML |
| **retail-etl** (este) | Extracción y transformación de datos del P&G |

El archivo `data.js` generado por este pipeline va a `retail-cm-ia/dashboard/data/` y está en el `.gitignore` de ese repo (datos financieros, no se versionan).
