# coding: utf-8
import csv
import io
import os
import unicodedata

from Autodesk.Revit.DB import (  # type: ignore
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    StorageType,
    Transaction,
    ViewSheet,
)
from pyrevit import forms, script  # type: ignore
from valtria_core.text import ensure_text as safe_text


TITLE = "Crear planos desde CSV"
PARAM_NAME = "CATEGORIA PLANO"
MAX_PREVIEW_LINES = 15



class SimpleItem(object):
    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


def normalize_header(value):
    text = safe_text(value).strip().lower()
    if not text:
        return ""
    try:
        text = unicodedata.normalize("NFKD", text)
        text = u"".join([ch for ch in text if ord(ch) < 128])
    except Exception:
        pass
    for ch in (" ", "_", "-", ".", "#", "/", "\\", "(", ")", "[", "]", "{", "}", ":"):
        text = text.replace(ch, "")
    return text


def read_csv_text(path):
    last_exc = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with io.open(path, "r", encoding=encoding) as handle:
                return handle.read()
        except Exception as exc:
            last_exc = exc
    raise last_exc or Exception("No se pudo leer el archivo CSV.")


def read_csv_rows(path):
    raw_text = read_csv_text(path)
    if not raw_text.strip():
        return []

    sample_block = raw_text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample_block, delimiters=";,")
        reader = csv.reader(io.StringIO(raw_text), dialect)
    except Exception:
        delimiter = ";" if sample_block.count(";") > sample_block.count(",") else ","
        reader = csv.reader(io.StringIO(raw_text), delimiter=delimiter)

    rows = []
    for row in reader:
        if not row:
            continue
        cleaned = [safe_text(value).strip() for value in row]
        if not any(cleaned):
            continue
        rows.append(cleaned)
    return rows


def resolve_columns(rows):
    if not rows:
        return [], None, None, False

    number_keys = {
        "sheetnumber",
        "sheetno",
        "sheetnum",
        "sheet#",
        "numerohoja",
        "numerodehoja",
        "numhoja",
        "nhoja",
        "nrohoja",
    }
    name_keys = {
        "sheetname",
        "nombrehoja",
        "nombredelhoja",
    }

    header = rows[0]
    normalized = [normalize_header(value) for value in header]

    number_idx = None
    name_idx = None
    for idx, value in enumerate(normalized):
        if number_idx is None and value in number_keys:
            number_idx = idx
        if name_idx is None and value in name_keys:
            name_idx = idx

    if number_idx is not None and name_idx is not None:
        return rows[1:], number_idx, name_idx, True

    number_idx = 0
    name_idx = 1 if len(rows[0]) > 1 else None
    return rows, number_idx, name_idx, False


def get_parameter_by_name(element, name):
    target = safe_text(name).strip().lower()
    if not target:
        return None
    for param in element.Parameters:
        try:
            def_name = safe_text(param.Definition.Name).strip().lower()
        except Exception:
            continue
        if def_name == target:
            return param
    return None


def get_param_value(param):
    try:
        if param.StorageType == StorageType.String:
            return safe_text(param.AsString()).strip()
        value = safe_text(param.AsValueString()).strip()
        if value:
            return value
        return safe_text(param.AsInteger()).strip()
    except Exception:
        return ""


def collect_param_values(doc, param_name):
    values = set()
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        param = get_parameter_by_name(sheet, param_name)
        if not param:
            continue
        value = get_param_value(param)
        if value:
            values.add(value)
    return sorted(values)


def select_categoria_value(doc):
    values = collect_param_values(doc, PARAM_NAME)
    if values:
        items = [SimpleItem("__custom__", "Ingresar valor manual")]
        items.extend([SimpleItem(value, value) for value in values])
        picked = forms.SelectFromList.show(
            items,
            title=TITLE + " - " + PARAM_NAME,
            multiselect=False,
            button_name="Seleccionar",
            name_attr="name",
        )
        if not picked:
            return None
        if isinstance(picked, list):
            picked = picked[0]
        if picked.value != "__custom__":
            return picked.value

    value = forms.ask_for_string(
        default="",
        prompt="Valor para {0}:".format(PARAM_NAME),
        title=TITLE,
    )
    if value is None:
        return None
    value = safe_text(value).strip()
    if not value:
        return ""
    return value


def select_titleblock(doc):
    types = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsElementType()
        .ToElements()
    )
    if not types:
        return None

    items = []
    for symbol in types:
        items.append(SimpleItem(symbol, format_titleblock_label(symbol)))

    items = sorted(items, key=lambda x: safe_text(x.label).lower())
    picked = forms.SelectFromList.show(
        items,
        title=TITLE + " - Titleblock",
        multiselect=False,
        button_name="Seleccionar",
        name_attr="name",
    )
    if not picked:
        return None
    if isinstance(picked, list):
        picked = picked[0]
    return picked.value


def format_titleblock_label(symbol):
    family_name = ""
    try:
        family = getattr(symbol, "Family", None)
        if family is not None:
            family_name = safe_text(getattr(family, "Name", "")).strip()
    except Exception:
        family_name = ""
    if not family_name:
        family_name = safe_text(getattr(symbol, "FamilyName", "")).strip()

    type_name = safe_text(getattr(symbol, "Name", "")).strip()
    if not type_name or type_name.lower() == family_name.lower():
        for param_id in (
            BuiltInParameter.SYMBOL_NAME_PARAM,
            BuiltInParameter.ALL_MODEL_TYPE_NAME,
        ):
            try:
                param = symbol.get_Parameter(param_id)
                candidate = safe_text(param.AsString()).strip() if param else ""
            except Exception:
                candidate = ""
            if candidate:
                type_name = candidate
                break
    if not type_name:
        for pname in ("Type Name", "Nombre de tipo"):
            try:
                param = symbol.LookupParameter(pname)
                candidate = safe_text(param.AsString()).strip() if param else ""
            except Exception:
                candidate = ""
            if candidate:
                type_name = candidate
                break

    if family_name and type_name:
        return "{0} - {1}".format(family_name, type_name)
    return family_name or type_name or "Titleblock"


def prompt_sheet_selection(entries):
    if not entries:
        return []
    items = [
        SimpleItem(entry, u"{0} | {1}".format(entry["number"], entry["name"]))
        for entry in entries
    ]
    picked = forms.SelectFromList.show(
        items,
        title=TITLE + " - Seleccion de planos",
        multiselect=True,
        button_name="Continuar",
        name_attr="name",
    )
    if not picked:
        return []
    if not isinstance(picked, list):
        picked = [picked]
    return [item.value for item in picked]


def build_preview(entries):
    preview = []
    for idx, entry in enumerate(entries):
        if idx >= MAX_PREVIEW_LINES:
            preview.append("... ({0} en total)".format(len(entries)))
            break
        preview.append("- {0} | {1}".format(entry["number"], entry["name"]))
    return preview


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


csv_path = forms.pick_file(
    file_ext="csv",
    restore_dir=True,
    title="Selecciona el CSV con Sheet Number y Sheet Name",
)

if not csv_path:
    script.exit()

try:
    rows = read_csv_rows(csv_path)
except Exception as read_exc:
    forms.alert(
        "No se pudo leer el CSV seleccionado:\n{0}".format(read_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if not rows:
    forms.alert(
        "El CSV esta vacio o no contiene datos validos.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

data_rows, number_idx, name_idx, used_header = resolve_columns(rows)
if name_idx is None:
    forms.alert(
        "No se encontraron columnas suficientes. El CSV debe tener Sheet Number y Sheet Name.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

entries = []
invalid_rows = []
row_offset = 2 if used_header else 1

for idx, row in enumerate(data_rows):
    row_number = idx + row_offset
    if number_idx >= len(row) or name_idx >= len(row):
        invalid_rows.append((row_number, "Fila sin columnas suficientes"))
        continue
    number = safe_text(row[number_idx]).strip()
    name = safe_text(row[name_idx]).strip()
    if not number or not name:
        missing = []
        if not number:
            missing.append("sin Sheet Number")
        if not name:
            missing.append("sin Sheet Name")
        invalid_rows.append((row_number, ", ".join(missing)))
        continue
    entries.append({"number": number, "name": name, "row": row_number})

if not entries:
    forms.alert(
        "No se encontraron filas validas para crear planos.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

existing_by_number = {}
for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
    number = safe_text(sheet.SheetNumber).strip()
    if number:
        existing_by_number[number.lower()] = sheet

seen_numbers = set()
unique_entries = []
duplicates_in_csv = []
conflicts_in_model = []

for entry in entries:
    normalized = entry["number"].lower()
    if normalized in seen_numbers:
        duplicates_in_csv.append(entry)
        continue
    seen_numbers.add(normalized)

    existing_sheet = existing_by_number.get(normalized)
    if existing_sheet:
        conflicts_in_model.append((entry, existing_sheet))
        continue
    unique_entries.append(entry)

if not unique_entries:
    summary_lines = ["No hay planos nuevos para crear."]
    if invalid_rows:
        summary_lines.append("Filas con errores: {0}".format(len(invalid_rows)))
    if duplicates_in_csv:
        summary_lines.append("Duplicados en CSV: {0}".format(len(duplicates_in_csv)))
    if conflicts_in_model:
        summary_lines.append(
            "Sheet Number ya existe en el modelo: {0}".format(len(conflicts_in_model))
        )
    forms.alert(
        "\n".join(summary_lines),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

selected_entries = prompt_sheet_selection(unique_entries)
if not selected_entries:
    forms.alert(
        "Operacion cancelada. No se seleccionaron planos.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

unique_entries = selected_entries

titleblock = select_titleblock(doc)
if titleblock is None:
    forms.alert("No se selecciono un titleblock.", exitscript=True, title=TITLE)
    script.exit()

categoria_value = select_categoria_value(doc)
if categoria_value is None:
    forms.alert("Operacion cancelada. No se ingreso {0}.".format(PARAM_NAME), title=TITLE)
    script.exit()
if categoria_value == "":
    forms.alert("{0} no puede estar vacio.".format(PARAM_NAME), exitscript=True, title=TITLE)
    script.exit()

summary_lines = [
    "Archivo: {0}".format(os.path.basename(csv_path)),
    "Planos a crear: {0}".format(len(unique_entries)),
    "Titleblock: {0}".format(format_titleblock_label(titleblock)),
    "{0}: {1}".format(PARAM_NAME, categoria_value),
]
if invalid_rows:
    summary_lines.append("Filas con errores: {0}".format(len(invalid_rows)))
if duplicates_in_csv:
    summary_lines.append("Duplicados en CSV: {0}".format(len(duplicates_in_csv)))
if conflicts_in_model:
    summary_lines.append(
        "Sheet Number ya existe en el modelo: {0}".format(len(conflicts_in_model))
    )

summary_lines.append("")
summary_lines.append("Vista previa:")
summary_lines.extend(build_preview(unique_entries))

confirm = forms.alert(
    "\n".join(summary_lines) + "\n\nDeseas continuar?",
    yes=True,
    no=True,
    ok=False,
    title=TITLE,
)

if not confirm:
    script.exit()

output = script.get_output()
logger = script.get_logger()

transaction = Transaction(doc, "Crear planos desde CSV")
created_sheets = []
create_errors = []
param_errors = []

try:
    transaction.Start()

    if hasattr(titleblock, "IsActive") and not titleblock.IsActive:
        titleblock.Activate()
        doc.Regenerate()

    for entry in unique_entries:
        try:
            sheet = ViewSheet.Create(doc, titleblock.Id)
            sheet.SheetNumber = entry["number"]
            sheet.Name = entry["name"]

            param = get_parameter_by_name(sheet, PARAM_NAME)
            if param:
                try:
                    param.Set(categoria_value)
                except Exception as param_exc:
                    param_errors.append((entry, param_exc))
            else:
                param_errors.append((entry, "Parametro no encontrado"))

            created_sheets.append(sheet)
        except Exception as create_exc:
            create_errors.append((entry, create_exc))

    if created_sheets:
        transaction.Commit()
    else:
        transaction.RollBack()
except Exception as run_exc:
    transaction.RollBack()
    forms.alert(
        "No se pudieron crear los planos:\n{0}".format(run_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if created_sheets:
    output.print_md("### Planos creados ({0})".format(len(created_sheets)))
    for sheet in created_sheets:
        output.print_md("- {0} | {1}".format(sheet.SheetNumber, sheet.Name))

if create_errors:
    output.print_md("### Errores de creacion ({0})".format(len(create_errors)))
    for entry, err in create_errors:
        logger.error("Error creando %s: %s", entry["number"], err)
        output.print_md("- {0} | {1} -> {2}".format(entry["number"], entry["name"], err))

if param_errors:
    output.print_md(
        "### Problemas con {0} ({1})".format(PARAM_NAME, len(param_errors))
    )
    for entry, err in param_errors:
        logger.warning("No se pudo asignar %s en %s: %s", PARAM_NAME, entry["number"], err)
        output.print_md("- {0} | {1} -> {2}".format(entry["number"], entry["name"], err))

if invalid_rows:
    output.print_md("### Filas con errores en CSV ({0})".format(len(invalid_rows)))
    for row_number, reason in invalid_rows[:MAX_PREVIEW_LINES]:
        output.print_md("- Fila {0}: {1}".format(row_number, reason))
    if len(invalid_rows) > MAX_PREVIEW_LINES:
        output.print_md("... ({0} filas adicionales)".format(len(invalid_rows) - MAX_PREVIEW_LINES))

if duplicates_in_csv:
    output.print_md("### Duplicados en CSV ({0})".format(len(duplicates_in_csv)))
    for entry in duplicates_in_csv[:MAX_PREVIEW_LINES]:
        output.print_md("- {0} | {1}".format(entry["number"], entry["name"]))
    if len(duplicates_in_csv) > MAX_PREVIEW_LINES:
        output.print_md("... ({0} duplicados adicionales)".format(len(duplicates_in_csv) - MAX_PREVIEW_LINES))

if conflicts_in_model:
    output.print_md("### Sheet Number ya existe ({0})".format(len(conflicts_in_model)))
    for entry, sheet in conflicts_in_model[:MAX_PREVIEW_LINES]:
        output.print_md(
            "- {0} | {1} -> existe en {2}".format(
                entry["number"], entry["name"], safe_text(sheet.Name)
            )
        )
    if len(conflicts_in_model) > MAX_PREVIEW_LINES:
        output.print_md(
            "... ({0} conflictos adicionales)".format(
                len(conflicts_in_model) - MAX_PREVIEW_LINES
            )
        )

final_message = "Se crearon {0} planos.".format(len(created_sheets))
if create_errors or param_errors:
    final_message += " Revisa la ventana de resultados para mas detalles."

forms.alert(final_message, title=TITLE)
