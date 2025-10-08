# Valtria PyTools

Extensión de pyRevit para los equipos de VALTRIA con utilidades de control de calidad, gestión de impresión y soporte a modelado MEP.

<img width="938" height="100" alt="{A2DD10A5-D693-48D8-BE18-3C697C23DB3E}" src="https://github.com/user-attachments/assets/d57dbd87-e12e-4cee-a787-cc2fdf0b43e9" />


## Requisitos
- Revit 2019 a 2026.
- pyRevit estable con runtime de IronPython 2.7.
- Permisos de lectura/escritura en las carpetas donde se generarán IFC, PDF y CSV.

## Instalación
1. Descarga o clona este repositorio.
2. Copia la carpeta `ValtriaPyTools.extension` en `%APPDATA%\pyRevit\Extensions` o en la carpeta de extensiones corporativa.
3. Ejecuta `pyrevit caches clear` y después `pyrevit reload` (o reinicia Revit) para recargar la extensión.
4. Abre Revit y verifica que aparecen las pestañas **VALTRIA Tools** y **ValtriaPyTools**.

## Guía rápida de configuración de pyRevit
Pensada para alguien sin experiencia en programación:

1. **Instalar pyRevit**  
   - Ve a [pyrevitlabs.io](https://www.pyrevitlabs.io) y descarga el instalador estable.  
   - Durante la instalación elige la opción *Install for Current User*.
2. **Configurar el runtime correcto**  
   - Abre el menú Inicio de Windows y escribe `pyrevit` para abrir la aplicación **pyRevit CLI**.  
   - Ejecuta el comando `pyrevit env`. Comprueba que en la sección `Installed Clones` aparece `default | master | <ruta>` y que el `Runtime` es **IronPython 2.7**.  
   - Si no es IronPython 2.7, ejecuta `pyrevit clones switch master` y luego `pyrevit env` de nuevo para confirmar.
3. **Añadir la extensión Valtria**  
   - Con la CLI abierta, ejecuta `pyrevit extensions add <ruta_al_directorio_ValtriaPyTools.extension>`.  
   - Si prefieres copiar la carpeta manualmente, asegúrate de que se encuentra en `%APPDATA%\pyRevit\Extensions`.
4. **Limpiar cachés y recargar**  
   - Ejecuta `pyrevit caches clear` y luego `pyrevit reload`.  
   - Alternativa: cerrar Revit y volver a abrirlo tras limpiar la caché.
5. **Verificación final en Revit**  
   - Abre Revit.  
   - Comprueba que en la cinta aparecen las pestañas **VALTRIA Tools** y **ValtriaPyTools**.  
   - Pulsa cualquier botón de la pestaña para confirmar que se abre el cuadro de diálogo esperado.

> Consejo: guarda esta guía en PDF y compártela con el equipo para acelerar futuras instalaciones.

## Qué incluye cada botón

### Pestaña: VALTRIA Tools (existente)
- **Ribbon Inspector**: diagnostica pestañas duplicadas y estados de carga.
- **BIM Blog**: abre el artículo "BIM para salas limpias" publicado en valtria.com.
- **Crear Print Set**: genera un Print Set a partir de las hojas seleccionadas (con confirmación si el nombre ya existe).

### Pestaña: ValtriaPyTools (nueva)
- **IFC Export Views**: exporta vistas seleccionadas a archivos IFC. Incluye recordatorio para aislar elementos por vista cuando se requiera precisión total.
- **Print to PDF (Views/Sheets)**: imprime vistas u hojas seleccionadas en PDF usando un ViewSet temporal sin modificar el PrintSet activo.
- **Auto-Fit 3D Section Box**: ajusta el Section Box de la vista 3D activa a la selección (o al modelo visible) con un margen de 150 mm.
- **List HVAC Systems (CSV)**: resume los sistemas de climatización con número de elementos y longitud total en metros.
- **Ducts & Accessories Takeoff (CSV)**: genera un takeoff con conductos, accesorios y uniones, incluyendo categoría, sistema, tamaño, longitud, nivel, workset y comentarios.

## Limitaciones IFC por vista
La API de Revit no aísla completamente los elementos por vista durante la exportación a IFC. Se incluye un TODO para implementar un patrón de aislamiento temporal si se requiere un control absoluto. Para resultados limpios se recomienda duplicar la vista, aislar los elementos deseados manualmente y ejecutar la herramienta.

## Notas de compatibilidad IronPython
- Los scripts evitan anotaciones de tipo, f-strings y cualquier sintaxis exclusiva de Python 3.
- Las dependencias compartidas residen en `_lib/valtria_lib.py` para mantener compatibilidad con IronPython 2.7.
- Las rutas de salida se crean automáticamente si no existen.

## Estructura principal
```
ValtriaPyTools.extension/
  extension.yaml
  ValtriaPyTools.tab/
    Exports.panel/
    Printing.panel/
    Views.panel/
    HVAC.panel/
    Takeoff.panel/
  VALTRIA Tools.tab/  (herramientas históricas)
  _lib/
  README.md
```

## QA Rápido
1. Abrir Revit, recargar pyRevit y confirmar la pestaña **ValtriaPyTools** con los cinco botones nuevos.
2. **Print to PDF (Views/Sheets)**: seleccionar una vista y una hoja, ejecutar la herramienta y validar que genera PDFs en la carpeta seleccionada sin alterar el PrintSet activo.
3. **Auto-Fit 3D Section Box**: en una vista 3D, seleccionar dos elementos y ejecutar la herramienta para comprobar el ajuste con margen.
4. **List HVAC Systems (CSV)**: ejecutar en un modelo con sistemas de climatización y revisar que el CSV incluya nombre, recuento y longitud en metros.
5. **Ducts & Accessories Takeoff (CSV)**: validar que el CSV contiene filas para conductos, uniones y accesorios con las columnas definidas.
6. **IFC Export Views**: seleccionar vistas y confirmar la generación de archivos `.ifc` en la carpeta objetivo (teniendo presente la limitación descrita arriba).

## Resolución de problemas
- Si aparece el error `Can not de/activate native item`, utilizar **Ribbon Inspector** para revisar pestañas duplicadas.
- Ante errores durante la ejecución de los scripts, se mostrará un `forms.alert` y el detalle quedará en la consola de pyRevit.
- Usar `pyrevit env` para verificar rutas cargadas si alguna pestaña no aparece.

## Licencia
Incluye aquí la licencia que corresponda al proyecto.
