# coding: utf-8
import csv
import io
import os

from Autodesk.Revit.DB import (  # type: ignore
    FilteredWorksetCollector,
    Transaction,
    Workset,
    WorksetDefaultVisibilitySettings,
    WorksetKind,
)
from pyrevit import forms, script  # type: ignore

TITLE = "Crear Worksets desde CSV"
BUNDLE_DIR = os.path.dirname(__file__)
SAMPLE_FILE = os.path.join(BUNDLE_DIR, "worksets_csv_ejemplo.csv")

uidoc = __revit__.ActiveUIDocument  # type: ignore[attr-defined]
if uidoc is None:
    forms.alert(
        "No hay un documento activo. Abre un modelo antes de ejecutar la herramienta.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

doc = uidoc.Document
if doc is None:
    forms.alert(
        "No es posible acceder al documento activo.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if not doc.IsWorkshared:
    forms.alert(
        "El documento actual no esta en modo worksharing. Habilita los worksets antes de continuar.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()


def _read_csv_text(path):
    last_exc = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with io.open(path, "r", encoding=encoding) as handle:
                return handle.read()
        except Exception as exc:
            last_exc = exc
    raise last_exc or Exception("No se pudo leer el archivo CSV.")


def _extract_workset_names(path):
    raw_text = _read_csv_text(path)
    if not raw_text.strip():
        return []

    sample_block = raw_text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample_block, delimiters=";,")
        reader = csv.reader(io.StringIO(raw_text), dialect)
    except Exception:
        delimiter = ";" if sample_block.count(";") > sample_block.count(",") else ","
        reader = csv.reader(io.StringIO(raw_text), delimiter=delimiter)

    names = []
    header_checked = False
    workset_col = None

    for row in reader:
        if not row:
            continue

        cleaned_row = [value.strip() for value in row]
        if not any(cleaned_row):
            continue

        if not header_checked:
            header_checked = True
            for idx, value in enumerate(cleaned_row):
                if value and "workset" in value.lower():
                    workset_col = idx
                    break
            if workset_col is not None:
                continue

        index = workset_col if workset_col is not None else len(cleaned_row) - 1
        if index >= len(cleaned_row):
            continue

        candidate = cleaned_row[index]
        if candidate:
            names.append(candidate)

    return names


output = script.get_output()
logger = script.get_logger()

sample_hint = ""
if os.path.exists(SAMPLE_FILE):
    sample_hint = " (ejemplo: {0})".format(os.path.basename(SAMPLE_FILE))

csv_path = forms.pick_file(
    file_ext="csv",
    restore_dir=True,
    title="Selecciona el CSV con la lista de worksets{0}".format(sample_hint),
)

if not csv_path:
    script.exit()

try:
    csv_names = _extract_workset_names(csv_path)
except Exception as read_exc:
    forms.alert(
        "No se pudo leer el CSV seleccionado:\n{0}".format(read_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if not csv_names:
    forms.alert(
        "El CSV no contiene valores reconocibles. Asegurate de que exista una columna llamada 'Workset'.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

existing_names = {
    ws.Name.strip().lower()
    for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
}
unique_names = []
duplicates_in_file = []
already_existing = []

seen = set()
for raw_name in csv_names:
    normalized = raw_name.strip()
    if not normalized:
        continue

    lowered = normalized.lower()
    if lowered in seen:
        duplicates_in_file.append(normalized)
        continue
    seen.add(lowered)

    if lowered in existing_names:
        already_existing.append(normalized)
        continue

    unique_names.append(normalized)

if not unique_names:
    message_parts = ["No hay worksets nuevos que crear."]
    if already_existing:
        message_parts.append(
            "- {0} nombres ya existen en el modelo.".format(len(already_existing))
        )
    if duplicates_in_file:
        message_parts.append(
            "- {0} nombres estaban duplicados en el CSV.".format(
                len(duplicates_in_file)
            )
        )

    forms.alert(
        "\n".join(message_parts),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

preview_limit = 10
preview_lines = ["- {0}".format(name) for name in unique_names[:preview_limit]]
if len(unique_names) > preview_limit:
    preview_lines.append("... ({0} en total)".format(len(unique_names)))

summary_lines = [
    "Archivo seleccionado: {0}".format(os.path.basename(csv_path)),
    "Worksets nuevos a crear: {0}".format(len(unique_names)),
]
if already_existing:
    summary_lines.append(
        "Omitidos por existir en el modelo: {0}".format(len(already_existing))
    )
if duplicates_in_file:
    summary_lines.append(
        "Duplicados detectados en el CSV: {0}".format(len(duplicates_in_file))
    )
summary_lines.append("")
summary_lines.append("Vista previa:")
summary_lines.extend(preview_lines)

confirm = forms.alert(
    "\n".join(summary_lines) + "\n\nDeseas continuar?",
    yes=True,
    no=True,
    ok=False,
    title=TITLE,
)

if not confirm:
    script.exit()

open_state_default = False
visible_default = False

if forms.alert(
    "Por defecto los nuevos worksets se crean cerrados y no visibles.\n\n¿Quieres crearlos abiertos y visibles?",
    yes=True,
    no=True,
    ok=False,
    title=TITLE,
):
    open_state_default = True
    visible_default = True

try:
    visibility_settings = WorksetDefaultVisibilitySettings.GetWorksetDefaultVisibilitySettings(
        doc
    )
except Exception as vis_exc:
    forms.alert(
        "No se pudo acceder a la configuracion de visibilidad:\n{0}".format(vis_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if visibility_settings is None:
    forms.alert(
        "La configuracion de visibilidad no esta disponible en este documento.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

transaction = Transaction(doc, "Crear worksets desde CSV")
transaction_started = False
created_worksets = []
creation_errors = []
visibility_errors = []
state_errors = []

try:
    transaction.Start()
    transaction_started = True

    for name in unique_names:
        try:
            new_workset = Workset.Create(doc, name)
            created_worksets.append(new_workset)

            try:
                visibility_settings.SetWorksetVisibility(new_workset.Id, visible_default)
            except Exception as vis_exc:
                visibility_errors.append((name, vis_exc))

            try:
                if new_workset.IsOpen != open_state_default:
                    new_workset.IsOpen = open_state_default
            except Exception as state_exc:
                state_errors.append((name, state_exc))
        except Exception as create_exc:
            creation_errors.append((name, create_exc))

    transaction.Commit()
except Exception as run_exc:
    if transaction_started:
        transaction.RollBack()

    forms.alert(
        "No se pudieron crear los worksets:\n{0}".format(run_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if created_worksets:
    output.print_md("### Worksets creados ({0})".format(len(created_worksets)))
    for ws in created_worksets:
        output.print_md("- {0}".format(ws.Name))

if creation_errors:
    output.print_md("### Errores de creacion ({0})".format(len(creation_errors)))
    for name, err in creation_errors:
        logger.error("No se pudo crear el workset %s: %s", name, err)
        output.print_md("- {0} -> {1}".format(name, err))

if already_existing:
    output.print_md(
        "### Omitidos por existir ({0})".format(len(already_existing))
    )
    for name in already_existing:
        output.print_md("- {0}".format(name))

if duplicates_in_file:
    output.print_md(
        "### Duplicados en CSV ({0})".format(len(duplicates_in_file))
    )
    for name in duplicates_in_file:
        output.print_md("- {0}".format(name))

if visibility_errors:
    output.print_md("### Avisos de visibilidad ({0})".format(len(visibility_errors)))
    for name, err in visibility_errors:
        logger.warning("No se pudo ajustar la visibilidad de %s: %s", name, err)
        output.print_md("- {0} -> {1}".format(name, err))

if state_errors:
    output.print_md("### Avisos de estado abierto ({0})".format(len(state_errors)))
    for name, err in state_errors:
        logger.warning("No se pudo ajustar el estado abierto de %s: %s", name, err)
        output.print_md("- {0} -> {1}".format(name, err))

output.print_md("Archivo procesado: `{0}`".format(csv_path))

final_message = "Se crearon {0} worksets.".format(len(created_worksets))
if creation_errors:
    final_message += " Revisa la ventana de resultados para los detalles."

forms.alert(
    final_message + "\nConsulta la ventana de resultados para mas informacion.",
    title=TITLE,
)
