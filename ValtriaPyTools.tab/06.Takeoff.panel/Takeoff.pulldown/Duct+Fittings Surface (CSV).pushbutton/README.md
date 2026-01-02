# Duct + Fittings Surface (CSV)

Herramienta de cuantificación que calcula la superficie exterior en metros cuadrados de ductos y accesorios de HVAC sin escribir parámetros en Revit.

## Uso
1. Elija la **UT global** (vaina, perfil20, perfil30 u otras). Si además carga un CSV con UT por sistema, el valor global actúa como respaldo.
2. Opcionalmente active los mapas de **UT por sistema** y/o **Leq por tipo de fitting** y seleccione los CSV correspondientes.
3. Ajuste los filtros por **sistema** y **fase** si necesita limitar el cálculo.
4. Indique la **carpeta de exportación** y un **prefijo de archivo** para generar los cuatro CSV.
5. Pulse **Calcular y exportar** para crear:
   - `<prefijo>_DUCTS_DETALLE.csv`
   - `<prefijo>_DUCTS_RESUMEN.csv`
   - `<prefijo>_FITTINGS_DETALLE.csv`
   - `<prefijo>_FITTINGS_RESUMEN.csv`

## Formato de los CSV de entrada
- Separador obligatorio: `;`
- **UT por sistema**: `NombreSistema;valor`, donde el valor puede ser `vaina`, `perfil20`, `perfil30`, `otras` o un número en metros (usar `,` para decimales).
- **Leq por tipo**: `NombreTipo;valor`, admitiendo coincidencias exactas o parciales. El valor es la longitud equivalente en metros (decimales con `,`).

## Notas
- Todas las longitudes de Revit en pies se convierten automáticamente a metros dentro del script.
- Para ductos: `Área = Perímetro * (Longitud + UT)`.
- Para fittings: `Área = Perímetro_boca_mayor * (Leq + UT)` con mínimo opcional de **1,00 m² por pieza**.
- Las secciones circulares, rectangulares y ovales se soportan (oval con aproximación de Ramanujan).
- La herramienta no crea ni modifica parámetros en el modelo.
