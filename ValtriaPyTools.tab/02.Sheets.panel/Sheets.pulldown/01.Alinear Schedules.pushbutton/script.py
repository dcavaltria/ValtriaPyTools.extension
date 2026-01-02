# -*- coding: utf-8 -*-
"""
Alinea columnas de schedules y ajusta su ancho para evitar saltos de linea.

Flujo:
- Permite seleccionar uno o varios schedules del documento.
- Solicita al usuario la alineacion horizontal y vertical deseada.
- Aplica la alineacion en las secciones principales del schedule.
- Calcula un ancho aproximado por contenido (sin cortes) y lo asigna.

Pensado para Revit 2019+ ejecutado desde pyRevit.
"""

from Autodesk.Revit.DB import (
    ElementId,
    HorizontalAlignmentStyle,
    ScheduleSheetInstance,
    SectionType,
    TableCellStyle,
    TableCellStyleOverrideOptions,
    Transaction,
    VerticalAlignmentStyle,
    ViewSchedule,
    ViewDuplicateOption,
    FilteredElementCollector,
)
from pyrevit import forms, revit


SECTION_ORDER = (
    SectionType.Header,
    SectionType.Body,
    SectionType.Summary,
    SectionType.Footer,
)

# Factor aproximado para convertir caracteres a pixeles (depende de la fuente usada en Revit).
AVERAGE_CHAR_PIXELS = 7.0
PADDING_PIXELS = 20
MIN_FALLBACK_PIXELS = 60
MAX_ROWS_TO_SCAN = 2000


class ListItem(object):
    """Contenedor simple para SelectFromList."""

    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


def gather_selected_schedules(doc):
    uidoc = revit.uidoc
    selections = []
    seen_keys = set()

    if uidoc is not None:
        element_ids = list(uidoc.Selection.GetElementIds())
        for element_id in element_ids:
            element = doc.GetElement(element_id)
            if element is None:
                continue
            if isinstance(element, ViewSchedule) and not element.IsTemplate:
                key = (element.Id.IntegerValue, None)
                if key not in seen_keys:
                    selections.append((element, None))
                    seen_keys.add(key)
            elif isinstance(element, ScheduleSheetInstance):
                schedule = doc.GetElement(element.ScheduleId)
                if (
                    isinstance(schedule, ViewSchedule)
                    and not schedule.IsTemplate
                    and element.Id.IntegerValue not in seen_keys
                ):
                    key = (schedule.Id.IntegerValue, element.Id.IntegerValue)
                    if key not in seen_keys:
                        selections.append((schedule, element.Id.IntegerValue))
                        seen_keys.add(key)

    active_view = getattr(doc, "ActiveView", None)
    if (
        isinstance(active_view, ViewSchedule)
        and not active_view.IsTemplate
        and (active_view.Id.IntegerValue, None) not in seen_keys
    ):
        selections.append((active_view, None))
        seen_keys.add((active_view.Id.IntegerValue, None))

    return selections


def ask_alignment(title, options, default_key):
    ordered = sorted(options, key=lambda opt: 0 if opt[0] == default_key else 1)
    items = [ListItem(key, label) for key, label in ordered]
    picked = forms.SelectFromList.show(
        items,
        title=title,
        multiselect=False,
        button_name="Aceptar",
        name_attr="name",
    )
    if not picked:
        return None
    if isinstance(picked, list):
        picked = picked[0]
    return picked.value


def get_section(table_data, section_type):
    try:
        return table_data.GetSectionData(section_type)
    except Exception:
        return None


def clone_style(section, row_number, column_number):
    try:
        base_style = section.GetTableCellStyle(row_number, column_number)
        if base_style:
            return TableCellStyle(base_style)
        return None
    except Exception:
        return None


def apply_alignment_to_column(table_data, column_number, horizontal_style, vertical_style):
    for section_type in SECTION_ORDER:
        if section_type == SectionType.Footer:
            continue
        section = get_section(table_data, section_type)
        if section is None:
            continue
        if not section.NumberOfColumns:
            continue
        if not section.IsValidColumnNumber(column_number):
            continue
        try:
            reference_row = section.FirstRowNumber
        except Exception:
            continue
        style = clone_style(section, reference_row, column_number)
        if style is None:
            continue
        try:
            override = style.GetCellStyleOverrideOptions()
        except Exception:
            override = None
        if override is None:
            override = TableCellStyleOverrideOptions()
        override.HorizontalAlignment = True
        override.VerticalAlignment = True
        style.SetCellStyleOverrideOptions(override)
        style.FontHorizontalAlignment = horizontal_style
        style.FontVerticalAlignment = vertical_style
        section.SetCellStyle(column_number, style)


def iter_column_texts(table_data, column_number):
    for section_type in SECTION_ORDER:
        if section_type == SectionType.Footer:
            continue
        section = get_section(table_data, section_type)
        if section is None:
            continue
        if not section.NumberOfColumns:
            continue
        if not section.IsValidColumnNumber(column_number):
            continue
        first_row = section.FirstRowNumber
        last_row = section.LastRowNumber
        if last_row < first_row:
            continue
        rows_to_scan = min(MAX_ROWS_TO_SCAN, last_row - first_row + 1)
        last_row = first_row + rows_to_scan - 1
        for row in range(first_row, last_row + 1):
            try:
                raw_text = section.GetCellText(row, column_number)
            except Exception:
                raw_text = ""
            if not raw_text:
                continue
            normalized = raw_text.replace("\r\n", "\n").replace("\r", "\n")
            for fragment in normalized.split("\n"):
                fragment = fragment.strip()
                if fragment:
                    yield fragment


def compute_column_width_pixels(schedule, table_data, body_section, column_number):
    max_chars = 0
    for text in iter_column_texts(table_data, column_number):
        length = len(text)
        if length > max_chars:
            max_chars = length
    base = PADDING_PIXELS
    if max_chars > 0:
        base += int(round(max_chars * AVERAGE_CHAR_PIXELS))
    min_width = getattr(schedule, "MinimumColumnWidth", None)
    if not min_width or min_width <= 0:
        min_width = MIN_FALLBACK_PIXELS
    width_pixels = max(base, min_width)
    max_width = getattr(schedule, "MaximumColumnWidth", None)
    if max_width and max_width > 0:
        width_pixels = min(width_pixels, max_width)
    return width_pixels


def pixels_to_feet(body_section, column_number, target_pixels):
    try:
        current_pixels = body_section.GetColumnWidthInPixels(column_number)
        current_feet = body_section.GetColumnWidth(column_number)
        if current_pixels and current_feet:
            scale = current_feet / float(current_pixels)
            return target_pixels * scale
    except Exception:
        pass
    return None


def capture_row_heights(table_data):
    heights = []
    for section_type in SECTION_ORDER:
        section = get_section(table_data, section_type)
        if section is None:
            continue
        first_row = section.FirstRowNumber
        last_row = section.LastRowNumber
        if last_row < first_row:
            continue
        for row in range(first_row, last_row + 1):
            try:
                height = section.GetRowHeight(row)
                heights.append((section_type, row, height))
            except Exception:
                continue
    return heights


def restore_row_heights(table_data, heights):
    for section_type, row, height in heights:
        section = get_section(table_data, section_type)
        if section is None:
            continue
        try:
            current = section.GetRowHeight(row)
        except Exception:
            current = None
        if current is not None and abs(current - height) < 1e-6:
            continue
        try:
            section.SetRowHeight(row, height)
        except Exception:
            pass


def format_schedule(schedule, horizontal_style, vertical_style):
    table_data = schedule.GetTableData()
    body_section = get_section(table_data, SectionType.Body)
    if body_section is None or not body_section.NumberOfColumns:
        return 0, 0

    definition = schedule.Definition
    visible_fields = []
    try:
        field_count = definition.GetFieldCount()
    except Exception:
        field_count = 0
    for index in range(field_count):
        try:
            field_id = definition.GetFieldId(index)
            field = definition.GetField(field_id)
        except Exception:
            continue
        if field is None:
            continue
        try:
            if field.IsHidden:
                continue
        except Exception:
            pass
        visible_fields.append(field)

    original_heights = capture_row_heights(table_data)

    first_column = body_section.FirstColumnNumber
    last_column = body_section.LastColumnNumber
    columns_processed = 0
    width_updates = 0

    for column_number in range(first_column, last_column + 1):
        columns_processed += 1
        apply_alignment_to_column(table_data, column_number, horizontal_style, vertical_style)

        width_pixels = compute_column_width_pixels(schedule, table_data, body_section, column_number)
        current_pixels = None
        try:
            current_pixels = body_section.GetColumnWidthInPixels(column_number)
        except Exception:
            current_pixels = None

        target_pixels = width_pixels
        if current_pixels and abs(target_pixels - current_pixels) < 2:
            continue

        target_feet = pixels_to_feet(body_section, column_number, target_pixels)

        updated = False
        try:
            body_section.SetColumnWidthInPixels(column_number, int(round(target_pixels)))
            updated = True
        except Exception:
            updated = False

        if target_feet is not None:
            try:
                body_section.SetColumnWidth(column_number, target_feet)
                updated = True
            except Exception:
                pass

        if updated:
            width_updates += 1
            resulting_width = None
            try:
                resulting_width = body_section.GetColumnWidth(column_number)
            except Exception:
                resulting_width = target_feet

            column_index = column_number - first_column
            if resulting_width and column_index < len(visible_fields):
                field = visible_fields[column_index]
                try:
                    field.GridColumnWidth = resulting_width
                except Exception:
                    pass
                try:
                    field.SheetColumnWidth = resulting_width
                except Exception:
                    pass

    restore_row_heights(table_data, original_heights)

    return columns_processed, width_updates


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No hay documento activo.", title="Alinear schedules", exitscript=True)

    selected_entries = gather_selected_schedules(doc)
    if not selected_entries:
        forms.alert(
            "Selecciona una o varias schedules directamente en Revit (Project Browser o instancia en hoja) y vuelve a ejecutar.",
            title="Alinear schedules",
            exitscript=True,
        )

    instances_by_schedule = {}
    for instance in FilteredElementCollector(doc).OfClass(ScheduleSheetInstance):
        schedule_id = instance.ScheduleId
        if schedule_id is None or schedule_id == ElementId.InvalidElementId:
            continue
        instances_by_schedule.setdefault(schedule_id.IntegerValue, []).append(instance.Id.IntegerValue)

    schedule_data = {}
    for schedule, instance_id in selected_entries:
        sid = schedule.Id.IntegerValue
        record = schedule_data.setdefault(
            sid,
            {
                "original": schedule,
                "selected_instances": set(),
                "sources": set(),
            },
        )
        if instance_id is None:
            record["sources"].add("view")
        else:
            record["sources"].add("instance")
            record["selected_instances"].add(instance_id)

    targets = []
    skipped = []
    duplicate_notes = []

    for sid, data in schedule_data.items():
        schedule = data["original"]
        selected_instances = data["selected_instances"]
        all_instances = set(instances_by_schedule.get(sid, []))

        if not selected_instances:
            if all_instances:
                skipped.append(
                    (
                        schedule,
                        "La schedule está colocada en otras hojas. Selecciona las instancias deseadas para evitar modificar otras.",
                    )
                )
                continue
            data["processed"] = schedule
            targets.append((schedule, data))
            continue

        invalid_selected = selected_instances - all_instances
        if invalid_selected:
            skipped.append(
                (
                    schedule,
                    "No se encontró alguna instancia seleccionada en las hojas.",
                )
            )
            continue

        missing = all_instances - selected_instances
        schedule_to_use = schedule

        if missing:
            data["requires_split"] = True
        else:
            data["requires_split"] = False

        data["processed"] = schedule_to_use
        targets.append((schedule_to_use, data))

    horizontal_key = ask_alignment(
        "Alineacion horizontal",
        [("left", "Izquierda"), ("center", "Centro"), ("right", "Derecha")],
        default_key="center",
    )
    if horizontal_key is None:
        forms.alert("Operacion cancelada (no se eligio alineacion horizontal).", title="Alinear schedules", exitscript=True)

    vertical_key = ask_alignment(
        "Alineacion vertical",
        [("top", "Superior"), ("middle", "Centro"), ("bottom", "Inferior")],
        default_key="middle",
    )
    if vertical_key is None:
        forms.alert("Operacion cancelada (no se eligio alineacion vertical).", title="Alinear schedules", exitscript=True)

    horizontal_map = {
        "left": HorizontalAlignmentStyle.Left,
        "center": HorizontalAlignmentStyle.Center,
        "right": HorizontalAlignmentStyle.Right,
    }
    vertical_map = {
        "top": VerticalAlignmentStyle.Top,
        "middle": VerticalAlignmentStyle.Middle,
        "bottom": VerticalAlignmentStyle.Bottom,
    }

    horizontal_style = horizontal_map.get(horizontal_key, HorizontalAlignmentStyle.Center)
    vertical_style = vertical_map.get(vertical_key, VerticalAlignmentStyle.Middle)

    processed = []
    errors = []

    transaction = Transaction(doc, "Alinear schedules y ajustar columnas")
    transaction.Start()

    try:
        for schedule, data in targets:
            try:
                if data.get("requires_split"):
                    selected_instances = data["selected_instances"]
                    try:
                        duplicate_id = schedule.Duplicate(ViewDuplicateOption.Duplicate)
                        new_schedule = doc.GetElement(duplicate_id)
                        if new_schedule is None:
                            raise RuntimeError("No se pudo obtener la copia de la schedule.")
                        instance_infos = []
                        for inst_id in selected_instances:
                            instance = doc.GetElement(ElementId(inst_id))
                            if instance is None:
                                continue
                            sheet_id = getattr(instance, "OwnerViewId", None)
                            if sheet_id is None or sheet_id == ElementId.InvalidElementId:
                                continue
                            point = getattr(instance, "Point", None)
                            if point is None:
                                continue
                            rotation = getattr(instance, "Rotation", None)
                            instance_infos.append(
                                {
                                    "original_id": inst_id,
                                    "sheet_id": sheet_id,
                                    "point": point,
                                    "rotation": rotation,
                                }
                            )
                        data["selected_instances"].clear()
                        for info in instance_infos:
                            doc.Delete(ElementId(info["original_id"]))
                            new_instance = ScheduleSheetInstance.Create(
                                doc, info["sheet_id"], new_schedule.Id, info["point"]
                            )
                            if info["rotation"] is not None:
                                try:
                                    new_instance.Rotation = info["rotation"]
                                except Exception:
                                    pass
                            data["selected_instances"].add(new_instance.Id.IntegerValue)
                        duplicate_notes.append((schedule, new_schedule, len(selected_instances)))
                        data["processed"] = new_schedule
                        schedule = new_schedule
                    except Exception as dup_err:
                        skipped.append((schedule, "Error al duplicar para aislar instancias: {0}".format(dup_err)))
                        continue

                columns, width_changes = format_schedule(schedule, horizontal_style, vertical_style)
                processed.append((schedule, columns, width_changes))
            except Exception as err:
                errors.append((schedule, str(err)))
        transaction.Commit()
    except Exception as err:
        transaction.RollBack()
        forms.alert("Error inesperado:\n{0}".format(err), title="Alinear schedules")
        return

    lines = []
    if processed:
        lines.append("Schedules formateados:")
        for schedule, columns, width_changes in processed:
            lines.append(
                "- {0}: columnas alineadas {1}, columnas con ancho ajustado {2}".format(
                    schedule.Name, columns, width_changes
                )
            )
    if duplicate_notes:
        lines.append("")
        lines.append("Schedules duplicadas para aislar instancias:")
        for original, duplicate, count in duplicate_notes:
            lines.append(
                "- {0} -> {1} (instancias reasignadas: {2})".format(
                    original.Name, duplicate.Name, count
                )
            )
    if errors:
        lines.append("")
        lines.append("Schedules con errores:")
        for schedule, errmsg in errors:
            lines.append("- {0}: {1}".format(schedule.Name, errmsg))
    if skipped:
        lines.append("")
        lines.append("Schedules omitidas:")
        for schedule, reason in skipped:
            lines.append("- {0}: {1}".format(schedule.Name, reason))

    if not lines:
        lines.append("No se realizaron cambios.")

    forms.alert("\n".join(lines), title="Alinear schedules")


if __name__ == "__main__":
    main()
