# -*- coding: utf-8 -*-
"""Mostrar historial de edicion del elemento seleccionado.

Uso:
1) Ejecuta el comando.
2) Selecciona uno o varios elementos en el modelo.
3) Revisa el TaskDialog con el historial (fecha y usuario).
"""
from __future__ import print_function

from Autodesk.Revit.DB import WorksharingUtils
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import forms, revit, script

logger = script.get_logger()


def _get_first_attr(obj, names):
    for name in names:
        if hasattr(obj, name):
            value = getattr(obj, name)
            if value:
                return value
    return None


def _format_date(value):
    if value is None:
        return "N/D"
    try:
        return value.ToString()
    except Exception:
        return str(value)


def _format_user(value):
    return value if value else "N/D"


def _build_user_info(info):
    return {
        "creator": _format_user(getattr(info, "Creator", None)),
        "last_editor": _format_user(getattr(info, "LastChangedBy", None)),
    }


def main():
    uidoc = revit.uidoc
    doc = revit.doc

    try:
        forms.alert(
            "Selecciona uno o varios elementos en el modelo.",
            title="VALTRIA Tools",
        )
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, "Selecciona uno o varios elementos"
        )
    except Exception:
        script.exit()

    elements = []
    for ref in refs:
        element = doc.GetElement(ref.ElementId)
        if element is not None:
            elements.append(element)

    if not elements:
        forms.alert(
            "No se pudieron obtener elementos validos.",
            title="VALTRIA Tools",
            exitscript=True,
        )
        return

    if not doc.IsWorkshared:
        forms.alert(
            "El modelo no esta en modo compartido.\n"
            "No hay historial de edicion disponible.",
            title="VALTRIA Tools",
            exitscript=True,
        )
        return

    all_lines = [
        "Revit no expone fecha/hora de modificacion por elemento.",
        "Se muestra solo el ultimo editor disponible.",
        "",
    ]
    for element in elements:
        try:
            info = WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
        except Exception as error:  # pragma: no cover - entorno de Revit
            logger.error("No se pudo leer el historial: %s", error)
            all_lines.append(
                "Elemento: {0} (Id {1})".format(element.Name, element.Id.IntegerValue)
            )
            all_lines.append("No se pudo leer el historial: {0}".format(error))
            all_lines.append("")
            continue

        user_info = _build_user_info(info)
        all_lines.append(
            "Elemento: {0} (Id {1})".format(element.Name, element.Id.IntegerValue)
        )
        all_lines.append(
            "Creador: {0} | Ultimo editor: {1}".format(
                user_info["creator"], user_info["last_editor"]
            )
        )
        all_lines.append("")

    TaskDialog.Show("Historial de ediciones", "\n".join(all_lines).strip())


if __name__ == "__main__":
    main()
