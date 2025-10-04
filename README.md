# VALTRIA PyTools

Extensi�n de pyRevit que aporta utilidades de control de calidad y gesti�n de impresi�n para los equipos de VALTRIA.

## Caracter�sticas
- **VALTRIA Tools**: pesta�a propia en el ribbon de Revit que agrupa las herramientas.
- **Ribbon Inspector**: bot�n de diagn�stico que lista las pesta�as cargadas y resalta duplicados.
- **BIM Blog**: acceso directo al artículo "BIM para salas limpias" publicado en valtria.com.
- **Crear Print Set**: genera un conjunto de impresi�n desde las hojas seleccionadas, con confirmaci�n si el nombre ya existe.

## Instalaci�n r�pida
1. Descarga o clona este repositorio.
2. Copia la carpeta ValtriaPyTools.extension en %APPDATA%\pyRevit\Extensions\.
3. Ejecuta pyrevit caches clear y despu�s pyrevit reload (o reinicia Revit).
4. Abre Revit y verifica la pesta�a **VALTRIA Tools**.
5. quiero añadir un linea mas en mi githbu para controld e veriones 

## Uso
- **Ribbon Inspector**: abre la lista de pesta�as activas y marca duplicados potenciales. �til cuando aparezcan errores tipo Can not de/activate native item.
- **BIM Blog**: abre el navegador predeterminado en el artículo informativo sobre BIM para salas limpias.
- **Crear Print Set**: selecciona hojas en el Project Browser, lanza la herramienta, proporciona un nombre y confirma si quieres sobrescribir sets existentes.

## Estructura principal
`
ValtriaPyTools.extension/
  extension.yaml
  VALTRIA Tools.tab/
    QA.panel/
      Ribbon Inspector.pushbutton/
      BIM Blog.pushbutton/
    Sheets.panel/
      Crear Print Set.pushbutton/
  lib/valtria_core/
  _repo/
`

## Desarrollo
- Los m�dulos comunes viven en lib/valtria_core/.
- Para a�adir nuevos botones, sigue el patr�n Nombre.panel/Nueva Herramienta.pushbutton/ con su propio undle.yaml y script.py.
- Ejecuta pyrevit env para revisar las rutas cargadas si algo no aparece en el ribbon.

## Resoluci�n de problemas
- Si ves el error Can not de/activate native item: ... RibbonTab, revisa pesta�as duplicadas con **Ribbon Inspector**.
- Aseg�rate de que s�lo exista una carpeta con el mismo nombre de pesta�a en %APPDATA% y %PROGRAMDATA%.
- Limpia cach�s (pyrevit caches clear) y recarga (pyrevit reload).

## Licencia
Incluye aqu� la licencia que corresponda al proyecto.
