# ADLB – Cantidades UNIQUE 76 (RCI)

| Archivo | Qué es |
|---|---|
| `CUADRO_CANTIDADES_UNIQUE76_RCI.xlsx` | Cuadro de cantidades limpio. Hojas: **DASHBOARD**, **NIVEL**, **BD_REVIT**, **CRUCE RCI** (+ LISTAS oculta). |
| `dashboard/index.html` | Tablero interactivo (tipo Power BI) con los mismos datos. |
| `revit/…_ORIGINAL.xlsx` | Exportación cruzada original (entrada). |
| `scripts/build_cuadro.py` | Regenera el Excel a partir de una nueva exportación Revit. |
| `scripts/build_dashboard.py` | Regenera `dashboard/index.html` desde `dashboard/data.json`. |

## Cómo funciona el Excel
- **NIVEL**: mismo cuadro y numeración. Las cantidades del cap. 7 son `SUMIFS` sobre BD_REVIT (cantidad global). Columna FUENTE indica Revit/Manual.
- **BD_REVIT**: una sola tabla con los 461 elementos (antes 6 hojas). Solo se edita la columna amarilla **Código NIVEL**; Ítem, Und. y Estado se calculan. Incluye “Código sugerido” para los 139 elementos sin asignar.
- **CRUCE RCI**: matriz ítem × nivel de piso con control `Dif.` contra NIVEL (todo en 0).
- **DASHBOARD**: KPIs, selector de ítem (celda amarilla) con tabla y gráfico por nivel, resumen por categoría.

## Actualizar con un nuevo export de Revit
```bash
python scripts/build_cuadro.py nuevo_export.xlsx CUADRO_CANTIDADES_UNIQUE76_RCI.xlsx
```
El export debe tener hojas `RCI - <Categoría>` con encabezados en la fila 6 (formato actual).
