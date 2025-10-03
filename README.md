# VALTRIA PyTools

Extensión de pyRevit que aporta utilidades de control de calidad y gestión de impresión para los equipos de VALTRIA.

## Características
- **VALTRIA Tools**: pestaña propia en el ribbon de Revit que agrupa las herramientas.
- **Ribbon Inspector**: botón de diagnóstico que lista las pestañas cargadas y resalta duplicados.
- **Crear Print Set**: genera un conjunto de impresión desde las hojas seleccionadas, con confirmación si el nombre ya existe.

## Instalación rápida
1. Descarga o clona este repositorio.
2. Copia la carpeta ValtriaPyTools.extension en %APPDATA%\pyRevit\Extensions\.
3. Ejecuta pyrevit caches clear y después pyrevit reload (o reinicia Revit).
4. Abre Revit y verifica la pestaña **VALTRIA Tools**.

## Uso
- **Ribbon Inspector**: abre la lista de pestañas activas y marca duplicados potenciales. Útil cuando aparezcan errores tipo Can not de/activate native item.
- **Crear Print Set**: selecciona hojas en el Project Browser, lanza la herramienta, proporciona un nombre y confirma si quieres sobrescribir sets existentes.

## Estructura principal
`
ValtriaPyTools.extension/
  extension.yaml
  VALTRIA Tools.tab/
    QA.panel/
      Ribbon Inspector.pushbutton/
    Sheets.panel/
      Crear Print Set.pushbutton/
  lib/valtria_core/
  _repo/
`

## Desarrollo
- Los módulos comunes viven en lib/valtria_core/.
- Para añadir nuevos botones, sigue el patrón Nombre.panel/Nueva Herramienta.pushbutton/ con su propio undle.yaml y script.py.
- Ejecuta pyrevit env para revisar las rutas cargadas si algo no aparece en el ribbon.

## Resolución de problemas
- Si ves el error Can not de/activate native item: ... RibbonTab, revisa pestañas duplicadas con **Ribbon Inspector**.
- Asegúrate de que sólo exista una carpeta con el mismo nombre de pestaña en %APPDATA% y %PROGRAMDATA%.
- Limpia cachés (pyrevit caches clear) y recarga (pyrevit reload).

## Licencia
Incluye aquí la licencia que corresponda al proyecto.
