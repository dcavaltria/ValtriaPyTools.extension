# Valtria PyTools

Extensión de pyRevit pensada para los equipos de VALTRIA con utilidades de control de calidad, documentación, IA aplicada y soporte al modelado MEP.

<img width="1870" height="98" alt="{D9F984DE-C044-4F1E-9197-80F89B8D41E1}" src="https://github.com/user-attachments/assets/6030207b-0369-4a95-b237-962c7b4209d4" />


## Índice
- [Requisitos](#requisitos)
- [Instalación rápida](#instalación-rápida)
- [Configuración de pyRevit](#configuración-de-pyrevit)
- [Paneles y herramientas](#paneles-y-herramientas)
- [Librerías compartidas](#librerías-compartidas)
- [Validación sugerida](#validación-sugerida)
- [Resolución de problemas](#resolución-de-problemas)
- [Licencia](#licencia)

## Requisitos
- Revit 2019 a 2026 (validado en entorno corporativo).
- pyRevit estable con runtime de IronPython 2.7 habilitado.
- Permisos de lectura y escritura en las carpetas donde se genera IFC, PDF y CSV.
- Para las herramientas de IA: acceso a internet, Python 3 para ejecutar `install_anthropic.bat` y una API key válida de Anthropic.

## Instalación rápida
1. Descarga o clona este repositorio.
2. Copia `ValtriaPyTools.extension` en `%APPDATA%\pyRevit\Extensions` o en la carpeta de extensiones compartida de tu organización.
3. (Opcional) Ejecuta en la CLI `pyrevit extensions add <ruta_completa_a_ValtriaPyTools.extension>` para registrar la extensión sin copiar archivos.
4. Refresca pyRevit con `pyrevit caches clear` seguido de `pyrevit reload` (o reinicia Revit).
5. Abre Revit y comprueba que aparece la pestaña **ValtriaPyTools**.

## Configuración de pyRevit
1. Instala pyRevit desde [pyrevitlabs.io](https://www.pyrevitlabs.io) usando la opción *Install for Current User*.
2. En la CLI `pyrevit`, valida el entorno con `pyrevit env` y asegúrate de que el clone `master` usa **IronPython 2.7**. Si no es así, ejecuta `pyrevit clones switch master`.
3. Si necesitas la pestaña en varias estaciones, sincroniza esta carpeta dentro del repositorio corporativo y usa `pyrevit extensions add` apuntando a la ruta UNC.
4. Tras copiar o actualizar la extensión, limpia cachés (`pyrevit caches clear`) y recarga (`pyrevit reload`) para aplicar los cambios.

## Paneles y herramientas

### 00.AI.panel
- **Ephyra**: asistente conversacional conectado a la API de Anthropic que lee la selección actual, puede generar resúmenes, revisar parámetros y ejecutar acciones controladas. Guarda caché diaria de respuestas y contador de consultas.
- **Consultar Claude** *(opcional)*: cliente ligero en WPF para consultas rápidas a Claude. Ejecuta con Python 3 (pyRevit `py3button`). Para activarlo elimina la extensión `.disable` del directorio `ClaudeAI.py3pushbutton.disable`.

Para ambos asistentes configura `api_config.json` con tu `api_key`. Los instaladores `install_anthropic.bat` y `instalar_anthropic_SIN_ADMIN.bat` añaden la librería `anthropic` sin necesidad de privilegios elevados.

### 01.QA.panel
- **Valtria BIM Blog**: abre el artículo corporativo de referencia sobre BIM para salas limpias.

### 02.Sheets.panel
- **Alinear Schedules**: alinea columnas y ajusta el ancho de schedules seleccionados.
- **Colocar Leyenda en Hojas**: inserta una vista de leyenda seleccionada en todas las hojas elegidas.
- **Crear Planos CSV**: crea planos desde un CSV con `Sheet Number` y `Sheet Name`, asignando `CATEGORIA PLANO` y titleblock.
- **Crear Print Set**: genera un `ViewSheetSet` sin modificar el Print Set activo, permite reutilizar nombres y eliminar sets previos antes de crear el nuevo.
- **Renombrar Hojas**: renombra `Sheet Number` y/o `Sheet Name` vía prefijo, sufijo o búsqueda y reemplazo, con previsualización y control de duplicados.

### 03.Printing.panel
- **Print to PDF (Views/Sheets)**: imprime vistas y hojas seleccionadas a PDF manteniendo intacto el diálogo de impresión. Selecciona impresora, carpeta de destino y monitoriza la creación del archivo antes de continuar.

### 04.Views.panel
- **Add Filter to View**: selecciona un filtro existente y lo aplica en la vista activa.
- **Auto-Fit 3D Section Box**: ajusta automáticamente el `SectionBox` de la vista 3D activa a la selección o al modelo visible, añadiendo un margen de 150 mm.
- **Crop View**: dibuja un marco de líneas de detalle siguiendo el recorte de las vistas seleccionadas en la hoja y oculta la visibilidad del crop.
- **Masking Region**: marca o desmarca "Masking" en los tipos de regiones rellenadas de los Detail Items seleccionados.
- **Transfer Project Browser**: copia configuraciones de organización del Project Browser entre proyectos abiertos.

### 05.HVAC.panel
- **Crear Cobre A**: crea o actualiza un `PipeScheduleType` con diámetros normalizados de tubería de cobre tipo A (ASTM B88, tipo L).
- **List HVAC Systems (CSV)**: exporta un resumen CSV de sistemas de climatización.

### 06.Takeoff.panel
- **Ducts & Accessories Takeoff (CSV)**: extrae conductos, accesorios y uniones a CSV con datos clave del modelo.

### 07.Model.Panel
- **Batch Shared Parameters**: agrega varios parámetros compartidos de una sola vez a categorías seleccionadas (instancia o tipo).
- **Cargar Familias**: selecciona y carga archivos RFA desde una carpeta con confirmación.
- **Corregir Links en Bulk**: herramienta en desarrollo para corregir múltiples archivos RVT con links incorrectos.
- **Mostrar Burbujas Grid**: activa las burbujas de rejilla en ambos extremos para todas las grids visibles de la vista activa.
- **Renombrar Tipos**: búsqueda y reemplazo masivo sobre nombres de tipos (o parámetros relacionados), respetando conflictos y campos de solo lectura.
- **Super Renamer**: edición guiada del parámetro `Mark` (u otro) sobre la selección actual añadiendo prefijos, sufijos o reemplazos con vista previa.
- **Crear Worksets CSV**: crea worksets de usuario desde un CSV de ejemplo.
- **Delete Workset**: elimina o transfiere contenidos de un workset de usuario con asistencia paso a paso y log de elementos afectados.
- **Renombrar Worksets**: busca y reemplaza texto en los nombres de worksets de usuario.

### 08.Link.panel
- **Cargar Links en Bulk**: carga múltiples archivos Revit (RVT) con configuraciones controladas.
- **Corregir Links Cloud**: reconvierte links RVT cargados como ruta absoluta de Desktop Connector a ruta cloud (Autodesk Docs/BIM 360 Docs).
- **Dump Cloud IDs**: extrae `hubId`, `projectId` y `modelId` de los links cloud ya cargados.

### 99.CheckInterferences
- **Check Interferences**: detector de colisiones MEP/Estructura con tolerancias configuradas (25/50/100 mm), reporte en la consola de pyRevit y exportación opcional a CSV para ACC.

### Quantification.panel
- **Duct + Fittings Surface (CSV)**: calcula superficie exterior [m²] de conductos y accesorios (circular/rect/oval) sin crear parámetros; exporta detalle y resumen a CSV.

## Librerías compartidas
- `_lib/valtria_lib.py`: funciones reutilizadas por varios scripts (manejo de selección, conversión de unidades, helpers de parámetros). Cualquier actualización debe mantenerse compatible con IronPython 2.7.
- `lib/`: utilidades auxiliares usadas por herramientas históricas (`valtria_utils`, etc.). Evita mover o renombrar sin revisar las referencias internas.

## Validación sugerida
1. Recargar pyRevit y comprobar que la pestaña **ValtriaPyTools** aparece con todos los paneles listados.
2. **Crear Print Set**: seleccionar varias hojas, crear un set y verificar que el conjunto original permanece intacto.
3. **Print to PDF (Views/Sheets)**: generar PDF de una vista y confirmar que se guarda en la carpeta elegida.
4. **Auto-Fit 3D Section Box**: con una selección en 3D, ejecutar y validar que el `SectionBox` se ajusta con margen.
5. **List HVAC Systems (CSV)**, **Ducts & Accessories Takeoff (CSV)** y **Duct + Fittings Surface (CSV)**: exportar archivos y revisar el contenido estructurado.
6. **Crear Planos CSV**: generar hojas desde un CSV de prueba y validar `CATEGORIA PLANO` y titleblock.
7. **Check Interferences**: lanzar en una vista de coordinación y revisar el reporte y la exportación CSV.
8. **Ephyra**: configurar `api_config.json`, ejecutar una consulta y comprobar la actualización del contador en `contador_consultas.json`.

## Resolución de problemas
- Ejecuta `pyrevit env` para confirmar rutas cargadas y runtime. Si la extensión no aparece, revisa rutas de `Extensions` y permisos.
- Si la CLI devuelve errores de importación de `anthropic`, ejecuta `install_anthropic.bat` (con Python 3 en PATH) o la versión sin administrador.
- Ante fallos en impresión a PDF, valida que la impresora elegida genera PDF y no requiere diálogo extra; la herramienta registrará los nombres en la consola de pyRevit.
- Muchos scripts escriben trazas de error en la ventana de resultados (`script.get_output()`); revisa allí detalles adicionales antes de volver a ejecutar.
- Comprueba que las carpetas de salida (IFC/PDF/CSV) existen o que tienes permisos para crearlas desde Revit.

## Licencia
Incluye aquí la licencia que corresponda al proyecto (MIT, propietario, etc.).
