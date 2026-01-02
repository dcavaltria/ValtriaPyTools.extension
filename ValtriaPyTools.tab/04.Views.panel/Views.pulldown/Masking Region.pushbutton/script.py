# -*- coding: utf-8 -*-
"""Activa o desactiva Masking en los tipos de region rellenada seleccionados."""

import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import FilledRegion, FilledRegionType, Transaction
from pyrevit import forms

from valtria_lib import (
    get_doc,
    get_uidoc,
    get_selected_elements,
    get_element_type,
    log_exception,
    log_to_file,
)


TITLE = "Masking Region"
LOG_TOOL = "masking_region_toggle"


try:
    unicode
except NameError:
    unicode = str  # type: ignore


def safe_text(value):
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


def log_line(message):
    log_to_file(LOG_TOOL, message)


def format_names(items, max_items=6):
    names = [safe_text(it).strip() for it in items if safe_text(it).strip()]
    if not names:
        return u""
    if len(names) <= max_items:
        return u", ".join(names)
    return u"{0} (+{1} mas)".format(u", ".join(names[:max_items]), len(names) - max_items)


def collect_filled_regions(elements):
    regions = []
    skipped = 0
    seen = set()
    for elem in elements or []:
        if elem is None:
            continue
        try:
            if isinstance(elem, FilledRegion):
                int_id = getattr(getattr(elem, "Id", None), "IntegerValue", None)
                if int_id is not None and int_id in seen:
                    continue
                if int_id is not None:
                    seen.add(int_id)
                regions.append(elem)
                continue
        except Exception:
            pass
        skipped += 1
    return regions, skipped


def collect_types(regions):
    type_map = {}
    for region in regions or []:
        try:
            fr_type = get_element_type(region)
        except Exception:
            fr_type = None
        if not isinstance(fr_type, FilledRegionType):
            continue
        tid = getattr(getattr(fr_type, "Id", None), "IntegerValue", None)
        if tid is None or tid in type_map:
            continue
        type_map[tid] = fr_type
    return type_map


def type_label(fr_type):
    name = safe_text(getattr(fr_type, "Name", u"")).strip()
    if name:
        return name
    try:
        return safe_text(getattr(getattr(fr_type, "Family", None), "Name", u""))
    except Exception:
        return u"(sin nombre)"


def main():
    log_line("----")
    log_line("Inicio Masking Region")
    try:
        doc = get_doc()
        uidoc = get_uidoc()
    except Exception as ctx_err:
        log_exception(ctx_err)
        forms.alert("No hay documento activo.", title=TITLE, warn_icon=True)
        return

    if doc is None or uidoc is None:
        forms.alert("No hay documento activo.", title=TITLE, warn_icon=True)
        log_line("Abortado: doc o uidoc nulos")
        return

    selected = get_selected_elements()
    if not selected:
        forms.alert(
            "Selecciona uno o varios Detail Items (regiones rellenadas) y vuelve a ejecutar.",
            title=TITLE,
            warn_icon=True,
        )
        log_line("Abortado: sin seleccion")
        return

    regions, skipped_count = collect_filled_regions(selected)
    if not regions:
        forms.alert(
            "La seleccion no contiene regiones rellenadas.",
            title=TITLE,
            warn_icon=True,
        )
        log_line("Abortado: sin FilledRegion en la seleccion")
        return

    type_map = collect_types(regions)
    if not type_map:
        forms.alert(
            "No se pudieron identificar los tipos de las regiones seleccionadas.",
            title=TITLE,
            warn_icon=True,
        )
        log_line("Abortado: sin FilledRegionType asociados")
        return

    action = forms.CommandSwitchWindow.show(
        ["Marcar (Masking = Si)", "Desmarcar (Masking = No)"],
        message=u"Se encontraron {0} tipo(s) de region rellenada.\nQue estado aplicar?".format(len(type_map)),
        title=TITLE,
    )
    if not action:
        log_line("Cancelado por el usuario")
        return
    target_state = action.startswith("Marcar")
    state_label = "Si" if target_state else "No"
    log_line("Solicitado Masking={0} para {1} tipos".format(state_label, len(type_map)))

    tx = Transaction(doc, TITLE)
    tx.Start()
    updated = []
    unchanged = []
    failed = []

    for fr_type in type_map.values():
        name = type_label(fr_type)
        try:
            current = bool(getattr(fr_type, "IsMasking", False))
        except Exception:
            current = None
        try:
            if current is not None and current == target_state:
                unchanged.append(name)
                continue
            fr_type.IsMasking = target_state
            updated.append(name)
        except Exception as err:
            failed.append((name, err))
            log_line("Error al actualizar {0}: {1}".format(name, safe_text(err)))

    if updated:
        try:
            tx.Commit()
        except Exception as err:
            tx.RollBack()
            log_exception(err)
            forms.alert(
                "No se pudo completar la operacion:\n{0}".format(safe_text(err)),
                title=TITLE,
                warn_icon=True,
            )
            return
    else:
        tx.RollBack()
        msg_parts = []
        if unchanged:
            msg_parts.append(
                "Los tipos seleccionados ya tenian Masking = {0}.".format(state_label)
            )
        if failed:
            msg_parts.append(
                "Errores al intentar aplicar el cambio:\n{0}".format(
                    format_names([u"{0}: {1}".format(n, safe_text(e)) for n, e in failed])
                )
            )
        if not msg_parts:
            msg_parts.append("No se realizaron cambios.")
        forms.alert(u"\n".join(msg_parts), title=TITLE, warn_icon=True)
        return

    summary = []
    summary.append("Masking establecido en: {0}".format(state_label))
    summary.append("Tipos actualizados: {0}".format(len(updated)))
    summary.append(format_names(updated))
    if unchanged:
        summary.append("")
        summary.append("Sin cambios (ya estaban asi): {0}".format(len(unchanged)))
        summary.append(format_names(unchanged))
    if failed:
        summary.append("")
        summary.append("Errores al aplicar: {0}".format(len(failed)))
        summary.append(format_names([u"{0}: {1}".format(n, safe_text(e)) for n, e in failed]))
    if skipped_count:
        summary.append("")
        summary.append("Omitidos por no ser regiones rellenadas: {0}".format(skipped_count))

    forms.alert(u"\n".join([line for line in summary if line is not None]), title=TITLE, warn_icon=bool(failed))


if __name__ == "__main__":
    try:
        main()
    except Exception as main_error:
        try:
            tb = traceback.format_exc()
            log_line("Error inesperado: {0}\n{1}".format(safe_text(main_error), safe_text(tb)))
        except Exception:
            pass
        log_exception(main_error)

