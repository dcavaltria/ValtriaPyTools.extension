# -*- coding: utf-8 -*-
"""
Clona una leyenda entre hojas respetando la posicion tomada de una hoja de referencia.

Pasos:
1. Solicita la hoja de referencia y lista las leyendas colocadas en ella.
2. Captura la posicion de las leyendas elegidas.
3. Pide las hojas destino.
4. Inserta las leyendas en cada hoja destino usando la misma posicion.

Compatible con Revit 2019+ (probado con pyRevit).
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    View,
    ViewSheet,
    ViewType,
    XYZ,
)

try:
    from Autodesk.Revit.DB import Viewport
except Exception:
    Viewport = None

from pyrevit import forms, revit


class SelectListItem(object):
    """Wrapper para mostrar nombre formateado en los selectores."""

    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


class LegendPlacement(object):
    """Informacion de la leyenda colocada en una hoja."""

    def __init__(self, view, viewport, location, viewport_type_id=None):
        self.view = view
        self.viewport = viewport
        self.location = location
        self.viewport_type_id = viewport_type_id


def feet_to_mm(value):
    return value * 304.8


def format_sheet_label(sheet):
    number = sheet.SheetNumber or ""
    name = sheet.Name or ""
    return u"{0} | {1}".format(number, name)


def format_location(location):
    if location is None:
        return "sin posicion registrada"
    return "X={0:.0f} mm, Y={1:.0f} mm".format(
        feet_to_mm(location.X),
        feet_to_mm(location.Y),
    )


def collect_all_sheets(doc, exclude_ids=None):
    exclude_ids = exclude_ids or set()
    sheets = []
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        if sheet.IsPlaceholder:
            continue
        if sheet.Id in exclude_ids:
            continue
        sheets.append(sheet)
    return sheets


def select_reference_sheet(doc):
    sheets = collect_all_sheets(doc)
    if not sheets:
        return None
    items = [SelectListItem(sheet, format_sheet_label(sheet)) for sheet in sheets]
    picked = forms.SelectFromList.show(
        items,
        title="Selecciona hoja de referencia",
        multiselect=False,
        button_name="Usar hoja",
        name_attr="name",
    )
    if not picked:
        return None
    if isinstance(picked, list):
        return picked[0].value
    return picked.value


def build_viewport_map(doc, sheet):
    mapping = {}
    if Viewport is None:
        return mapping
    try:
        for vp_id in sheet.GetAllViewports():
            viewport = doc.GetElement(vp_id)
            if viewport is not None:
                mapping[viewport.ViewId] = viewport
    except Exception:
        pass
    return mapping


def collect_legend_instances_on_sheet(doc, sheet):
    instances = []
    viewport_by_view = build_viewport_map(doc, sheet)
    try:
        placed_ids = list(sheet.GetAllPlacedViews())
    except Exception:
        placed_ids = []

    for view_id in placed_ids:
        view = doc.GetElement(view_id)
        if view is None:
            continue
        try:
            if view.ViewType != ViewType.Legend or view.IsTemplate:
                continue
        except Exception:
            continue

        viewport = viewport_by_view.get(view.Id)
        location = None
        viewport_type_id = None
        if viewport is not None:
            try:
                location = viewport.GetBoxCenter()
            except Exception:
                location = None
            try:
                viewport_type_id = viewport.GetTypeId()
            except Exception:
                viewport_type_id = None

        instances.append(LegendPlacement(view, viewport, location, viewport_type_id))
    return instances


def pick_reference_legend(instances):
    if not instances:
        return []
    items = []
    for instance in instances:
        label = u"{0} ({1})".format(
            instance.view.Name or "(sin nombre)",
            format_location(instance.location),
        )
        items.append(SelectListItem(instance, label))
    picked = forms.SelectFromList.show(
        items,
        title="Selecciona las leyendas a clonar",
        multiselect=True,
        button_name="Usar leyendas",
        name_attr="name",
    )
    if not picked:
        return []
    if isinstance(picked, list):
        return [item.value for item in picked]
    return [picked.value]


def prompt_destination_sheets(doc, reference_sheet):
    exclude_ids = {reference_sheet.Id}
    sheets = collect_all_sheets(doc, exclude_ids=exclude_ids)
    if not sheets:
        return []
    items = [SelectListItem(sheet, format_sheet_label(sheet)) for sheet in sheets]
    picked = forms.SelectFromList.show(
        items,
        title="Selecciona hojas destino",
        multiselect=True,
        button_name="Colocar leyenda",
        name_attr="name",
    )
    if not picked:
        return []
    return [item.value for item in picked]


def legend_already_on_sheet(doc, sheet, legend_view_id):
    try:
        placed_views = list(sheet.GetAllPlacedViews())
        if legend_view_id in placed_views:
            return True
    except Exception:
        pass
    if Viewport is not None:
        try:
            for viewport_id in sheet.GetAllViewports():
                viewport = doc.GetElement(viewport_id)
                if viewport and viewport.ViewId == legend_view_id:
                    return True
        except Exception:
            pass
    return False


def find_new_viewport(doc, sheet, legend_view_id, previous_ids):
    if Viewport is None:
        return None
    try:
        current_ids = list(sheet.GetAllViewports())
    except Exception:
        current_ids = []
    for viewport_id in current_ids:
        try:
            integer_id = viewport_id.IntegerValue
        except Exception:
            integer_id = None
        if integer_id is not None and integer_id in previous_ids:
            continue
        viewport = doc.GetElement(viewport_id)
        if viewport and viewport.ViewId == legend_view_id:
            return viewport
    return None


def apply_viewport_type(viewport, viewport_type_id, detail_steps):
    if viewport is None or viewport_type_id is None:
        return True, None
    try:
        if viewport.GetTypeId() != viewport_type_id:
            viewport.ChangeTypeId(viewport_type_id)
            detail_steps.append("ChangeTypeId")
        return True, None
    except Exception as err:
        return False, "No se pudo asignar tipo de viewport: {0}".format(err)


def place_legend_on_sheet(doc, sheet, legend_view, target_point, viewport_type_id):
    try:
        if hasattr(sheet, "CanAddView") and not sheet.CanAddView(legend_view):
            return False, "La hoja no permite agregar esta leyenda."
    except Exception:
        pass

    detail_steps = []
    viewport = None

    viewport_cls = Viewport
    viewport_create = getattr(viewport_cls, "Create", None) if viewport_cls else None
    if viewport_create is not None:
        try:
            can_add_to_sheet = getattr(viewport_cls, "CanAddViewToSheet", None)
            if callable(can_add_to_sheet):
                if not can_add_to_sheet(doc, sheet.Id, legend_view.Id):
                    return False, "La hoja no permite agregar esta leyenda."
        except Exception:
            pass
        try:
            creation_point = target_point if target_point is not None else XYZ(0, 0, 0)
            try:
                viewport = viewport_create(doc, sheet.Id, legend_view.Id, creation_point)
            except TypeError:
                viewport = viewport_create(doc, sheet.Id, legend_view.Id)
            detail_steps.append("Viewport.Create")
        except AttributeError:
            viewport = None
        except Exception as err:
            return False, str(err)

    if viewport is None:
        add_view = getattr(sheet, "AddView", None)
        if add_view is None:
            return False, (
                "No hay metodo disponible para agregar la leyenda a la hoja "
                "(Viewport.Create/AddView no disponibles)."
            )

        previous_ids = set()
        if viewport_cls is not None:
            try:
                previous_ids = set(
                    viewport_id.IntegerValue for viewport_id in sheet.GetAllViewports()
                )
            except Exception:
                previous_ids = set()

        try:
            add_view(legend_view)
            detail_steps.append("AddView")
        except Exception as err:
            return False, str(err)

        if viewport_cls is not None:
            viewport = find_new_viewport(doc, sheet, legend_view.Id, previous_ids)

    type_ok, type_detail = apply_viewport_type(viewport, viewport_type_id, detail_steps)
    if not type_ok:
        return False, type_detail

    if target_point is None:
        detail = " + ".join(detail_steps) if detail_steps else "Colocada"
        return True, detail

    if viewport is None:
        detail = " + ".join(detail_steps) if detail_steps else "Colocada"
        return (
            False,
            "No se pudo localizar el viewport creado para mover la leyenda. ({0})".format(
                detail
            ),
        )

    try:
        viewport.SetBoxCenter(target_point)
        detail_steps.append("SetBoxCenter")
        return True, " + ".join(detail_steps)
    except Exception as err:
        return False, "No se pudo mover la leyenda: {0}".format(err)


def build_summary_lines(title, entries):
    lines = []
    if entries:
        lines.append(title)
        for sheet, detail in entries:
            lines.append(
                u"- {0} ({1})".format(format_sheet_label(sheet), detail)
            )
        lines.append("")
    return lines


def build_legend_lines(title, instances):
    lines = []
    if instances:
        lines.append(title)
        for instance in instances:
            lines.append(u"- {0}".format(instance.view.Name or "(sin nombre)"))
        lines.append("")
    return lines


def main():
    doc = revit.doc
    if doc is None:
        forms.alert(
            "No hay documento activo.",
            title="Colocar leyenda",
            exitscript=True,
        )

    reference_sheet = select_reference_sheet(doc)
    if reference_sheet is None:
        forms.alert(
            "Operacion cancelada (no se eligio hoja de referencia).",
            title="Colocar leyenda",
            exitscript=True,
        )

    reference_instances = collect_legend_instances_on_sheet(doc, reference_sheet)
    if not reference_instances:
        forms.alert(
            "La hoja seleccionada no contiene leyendas colocadas.",
            title="Colocar leyenda",
            exitscript=True,
        )

    legend_instances = pick_reference_legend(reference_instances)
    if not legend_instances:
        forms.alert(
            "Operacion cancelada (no se eligieron leyendas).",
            title="Colocar leyenda",
            exitscript=True,
        )

    legends_without_location = [
        instance for instance in legend_instances if instance.location is None
    ]
    legend_instances = [
        instance for instance in legend_instances if instance.location is not None
    ]
    if not legend_instances:
        forms.alert(
            "No se pudo obtener la posicion de las leyendas seleccionadas.",
            title="Colocar leyenda",
            exitscript=True,
        )

    destination_sheets = prompt_destination_sheets(doc, reference_sheet)
    if not destination_sheets:
        forms.alert(
            "No se seleccionaron hojas destino.",
            title="Colocar leyenda",
            exitscript=True,
        )

    placed = []
    skipped = []
    errors = []

    transaction = Transaction(
        doc, "Colocar leyendas replicando posicion"
    )
    transaction.Start()

    try:
        for sheet in destination_sheets:
            if sheet.IsPlaceholder:
                skipped.append((sheet, "Sheet placeholder"))
                continue
            for legend_instance in legend_instances:
                legend_view = legend_instance.view
                target_point = legend_instance.location
                viewport_type_id = legend_instance.viewport_type_id

                if legend_already_on_sheet(doc, sheet, legend_view.Id):
                    skipped.append(
                        (sheet, "Leyenda '{0}': ya contiene la leyenda".format(legend_view.Name))
                    )
                    continue

                success, detail = place_legend_on_sheet(
                    doc, sheet, legend_view, target_point, viewport_type_id
                )
                if success:
                    placed.append(
                        (sheet, "Leyenda '{0}': {1}".format(legend_view.Name, detail))
                    )
                else:
                    errors.append(
                        (sheet, "Leyenda '{0}': {1}".format(legend_view.Name, detail))
                    )

        transaction.Commit()
    except Exception as err:
        transaction.RollBack()
        forms.alert(
            "Error inesperado:\n{0}".format(err),
            title="Colocar leyenda",
        )
        return

    message = []
    message.extend(build_summary_lines("Leyendas colocadas:", placed))
    message.extend(build_summary_lines("Sheets omitidas:", skipped))
    message.extend(build_summary_lines("Errores:", errors))
    message.extend(
        build_legend_lines(
            "Leyendas omitidas sin posicion:", legends_without_location
        )
    )

    if not message:
        message.append("No se realizaron cambios.")

    forms.alert("\n".join(message), title="Colocar leyenda")


if __name__ == "__main__":
    main()
