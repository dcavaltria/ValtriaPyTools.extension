# Quick Start - 🜂 Ephyra en ValtriaPyTools

## Pasos básicos

### 1. Configurar la API key

Guarda tu clave de Anthropic en:

```
%APPDATA%\pyrevit_claude.json
```

Contenido de ejemplo:

```json
{
    "ANTHROPIC_API_KEY": "sk-ant-xxxxxxxxxxxxxxxx"
}
```

También puedes definir la variable de entorno de usuario `ANTHROPIC_API_KEY`. El script usa primero la variable y después el archivo.

### 2. Recargar pyRevit

- pyRevit → Reload  
- o reinicia Revit

### 3. Usar el botón

1. Selecciona elementos en Revit (muros, puertas, MEP, etc.).
2. Haz clic en **ValtriaPyTools → AI → 🜂 Ephyra**.
3. Escribe tu pregunta y confirma.
4. Revisa la consola de pyRevit para ver el modelo usado, la respuesta y cualquier acción ejecutada.

## Consejos

- La herramienta detecta automáticamente el mejor modelo disponible (`/v1/models`).
- Si Anthropic devuelve un error, el detalle aparece en la consola y en `%APPDATA%\pyrevit_claude_error.log`.
- Para exportar datos, puedes pedir a Ephyra `{"action":"export_selection","format":"excel"}` (o `csv`/`json`) y elegir la ruta de guardado.
- Asegúrate de tener créditos activos en tu cuenta Anthropic.

Listo. Puedes adaptar el prompt o ampliar el contexto según tus necesidades.
