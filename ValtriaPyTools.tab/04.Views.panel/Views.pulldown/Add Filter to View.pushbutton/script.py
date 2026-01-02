# -*- coding: utf-8 -*-
"""Aplica rapidamente un filtro existente a la vista activa."""

import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..'))
LIB_DIR = os.path.join(EXTENSION_DIR, 'lib')
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    ParameterFilterElement,
    Transaction,
)
from pyrevit import forms

from valtria_lib import get_doc, get_uidoc, log_exception, log_to_file


TITLE = "Agregar filtro a vista"
LOG_TOOL = "add_filter_to_view"


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
    """Write to shared log for this tool."""
    log_to_file(LOG_TOOL, message)


def filter_label(filter_elem, doc):
    """Construye la etiqueta a mostrar para un filtro."""
    name = safe_text(getattr(filter_elem, "Name", u"")).strip() or u"(sin nombre)"
    try:
        cat_ids = filter_elem.GetCategories()
    except Exception:
        cat_ids = []
    cat_names = []
    if cat_ids:
        try:
            categories = doc.Settings.Categories
        except Exception:
            categories = None
        for cid in cat_ids:
            try:
                cat = categories.get_Item(cid) if categories else None
            except Exception:
                cat = None
            if cat:
                cat_names.append(safe_text(getattr(cat, "Name", u"")))
    if cat_names:
        cat_names = sorted([c for c in cat_names if c], key=lambda val: val.lower())
        if cat_names:
            return u"{0}  [{1}]".format(name, u", ".join(cat_names))
    return name


class FilterOption(object):
    """Adapter para SelectFromList."""

    def __init__(self, filter_elem, label):
        self.value = filter_elem
        self.label = label

    @property
    def name(self):
        return self.label


def _format_names(names, max_items=6):
    """Compacta listas largas para mostrarlas en alertas."""
    if not names:
        return u""
    if len(names) <= max_items:
        return u", ".join(names)
    return u"{0} (+{1} mas)".format(u", ".join(names[:max_items]), len(names) - max_items)


def collect_candidates(doc, view):
    """Devuelve los filtros que se pueden agregar a la vista junto con diagnosticos."""
    existing = set()
    try:
        for fid in view.GetFilters():
            existing.add(fid.IntegerValue)
    except Exception:
        pass

    collector = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    candidates = []
    diagnostics = {
        "existing": [],
        "unmapped": [],  # sin categorias asignadas
        "incompatible": [],  # la vista no los acepta
        "errors": [],
    }
    for filt in collector:
        if filt is None:
            continue
        fid = getattr(getattr(filt, "Id", None), "IntegerValue", None)
        if fid is None:
            continue
        label = filter_label(filt, doc)
        if fid in existing:
            diagnostics["existing"].append(label)
            continue
        try:
            cat_ids = list(filt.GetCategories()) or []
        except Exception as cat_err:
            cat_ids = []
            diagnostics["errors"].append((label, cat_err))
        if not cat_ids:
            diagnostics["unmapped"].append(label)
        try:
            can_add = bool(getattr(view, "CanAddFilter")(filt.Id))
        except Exception as err:
            can_add = False
            diagnostics["errors"].append((label, err))
        if not can_add:
            diagnostics["incompatible"].append(label)
            continue
        candidates.append(filt)
    return candidates, diagnostics


def pick_filters(candidates, doc):
    """Pide al usuario uno o varios filtros."""
    if not candidates:
        return []
    if len(candidates) == 1:
        return [candidates[0]]

    ordered = sorted(candidates, key=lambda f: safe_text(getattr(f, "Name", u"")).lower())
    items = []
    for filt in ordered:
        items.append(FilterOption(filt, filter_label(filt, doc)))
    picked = forms.SelectFromList.show(
        items,
        title=TITLE,
        multiselect=True,
        button_name="Agregar",
        name_attr="name",
    )
    if not picked:
        return []
    if isinstance(picked, (list, tuple)):
        selected = []
        for p in picked:
            selected.append(p.value if hasattr(p, "value") else p)
        return selected
    return [picked.value if hasattr(picked, "value") else picked]


def main():
    log_line("----")
    log_line("Inicio Add Filter to View")
    doc = get_doc()
    uidoc = get_uidoc()
    if doc is None or uidoc is None:
        forms.alert("No hay documento activo.", title=TITLE, warn_icon=True)
        log_line("Abortado: no hay documento o uidoc")
        return

    active_view = uidoc.ActiveView
    if active_view is None:
        forms.alert("No hay vista activa.", title=TITLE, warn_icon=True)
        log_line("Abortado: no hay vista activa")
        return
    if getattr(active_view, "IsTemplate", False):
        forms.alert("La vista activa es una plantilla. Abre una vista de proyecto.", title=TITLE, warn_icon=True)
        log_line("Abortado: la vista activa es plantilla")
        return

    doc_name = safe_text(getattr(doc, "Title", u""))
    view_name = safe_text(getattr(active_view, "Name", u""))
    log_line("Doc: {0} | Vista activa: {1}".format(doc_name, view_name))

    target_view = active_view
    applied_to_primary = False
    applied_to_template = False
    template_view = None
    tpl_name = u""

    # Si la vista es dependiente, los filtros deben agregarse en la vista principal.
    primary_view = None
    try:
        get_primary = getattr(active_view, "GetPrimaryViewId", None)
    except Exception:
        get_primary = None
    primary_view_id = None
    if callable(get_primary):
        try:
            primary_view_id = get_primary()
        except Exception:
            primary_view_id = None
    if primary_view_id and primary_view_id != ElementId.InvalidElementId:
        try:
            primary_view = doc.GetElement(primary_view_id)
        except Exception:
            primary_view = None
    if primary_view is not None:
        pv_name = safe_text(getattr(primary_view, "Name", u""))
        log_line("Vista dependiente detectada; vista principal: {0}".format(pv_name))
        proceed = forms.alert(
            u"La vista activa es dependiente de \"{0}\".\nLos filtros se gestionan en la vista principal.\nAgregar el filtro en la vista principal?".format(pv_name),
            title=TITLE,
            yes=True,
            no=True,
            warn_icon=False,
        )
        if not proceed:
            forms.alert(
                u"No se pueden agregar filtros directamente en vistas dependientes. Abre la vista principal \"{0}\" e intentalo de nuevo.".format(pv_name),
                title=TITLE,
                warn_icon=True,
            )
            log_line("Usuario cancelo: vista dependiente, no se cambio a principal.")
            return
        target_view = primary_view
        applied_to_primary = True

    try:
        template_id = getattr(target_view, "ViewTemplateId", None)
    except Exception:
        template_id = None
    if template_id and template_id != ElementId.InvalidElementId:
        try:
            template_view = doc.GetElement(template_id)
        except Exception:
            template_view = None
    if template_view is not None:
        tpl_name = safe_text(getattr(template_view, "Name", u""))
        use_template = forms.alert(
            u"La vista usa la plantilla \"{0}\".\nAgregar el filtro a la plantilla?".format(tpl_name),
            title=TITLE,
            yes=True,
            no=True,
            warn_icon=False,
        )
        if use_template:
            target_view = template_view
            applied_to_template = True
            log_line("Aplicando sobre plantilla: {0}".format(tpl_name))

    candidates, diagnostics = collect_candidates(doc, target_view)
    log_line("Candidatos: {0}, existentes: {1}, sin categorias: {2}, incompatibles: {3}, errores: {4}".format(
        len(candidates),
        len(diagnostics.get("existing") or []),
        len(diagnostics.get("unmapped") or []),
        len(diagnostics.get("incompatible") or []),
        len(diagnostics.get("errors") or []),
    ))
    if not candidates:
        msg = [u"No hay filtros disponibles para esta vista o ya estan aplicados."]
        if diagnostics.get("existing"):
            msg.append(u"{0} filtro(s) ya aplicados: {1}".format(
                len(diagnostics["existing"]), _format_names(diagnostics["existing"])
            ))
        if diagnostics.get("unmapped"):
            msg.append(
                u"{0} filtro(s) no tienen categorias asignadas. Revisa Gestionar > Filtros y mapea categorias: {1}".format(
                    len(diagnostics["unmapped"]), _format_names(diagnostics["unmapped"])
                )
            )
        if diagnostics.get("incompatible"):
            msg.append(
                u"{0} filtro(s) no son compatibles con esta vista/plantilla (categorias no admitidas).".format(
                    len(diagnostics["incompatible"])
                )
            )
        if applied_to_primary:
            msg.append(
                u"Se trabaja sobre la vista principal \"{0}\" porque la vista activa es dependiente.".format(
                    safe_text(getattr(target_view, "Name", u""))
                )
            )
        if template_view is not None and not applied_to_template:
            msg.append(
                u"La vista usa la plantilla \"{0}\"; prueba a agregarlos en la plantilla.".format(tpl_name)
            )
        forms.alert(u"\n".join([m for m in msg if m]), title=TITLE, warn_icon=True)
        log_line("Sin candidatos. Diagnosticos: {0}".format(msg))
        return

    auto_add_all = False
    if len(candidates) > 1:
        auto_add_all = forms.alert(
            u"Se encontraron {0} filtros compatibles.\nQuieres agregarlos todos a la vista seleccionada?".format(len(candidates)),
            title=TITLE,
            yes=True,
            no=True,
            warn_icon=False,
        )
        log_line("Usuario seleccion auto_add_all={0}".format(bool(auto_add_all)))
    else:
        log_line("Solo un candidato disponible; se seleccionara automaticamente.")

    selected_filters = candidates if auto_add_all else pick_filters(candidates, doc)
    if not selected_filters:
        return

    tx = Transaction(doc, TITLE)
    tx.Start()
    added = []
    failed = []
    for filt in selected_filters:
        if filt is None:
            continue
        try:
            target_view.AddFilter(filt.Id)
            try:
                target_view.SetFilterVisibility(filt.Id, True)
            except Exception:
                pass
            added.append(filt)
        except Exception as err:
            failed.append((safe_text(getattr(filt, "Name", u"")), err))
            log_line("Fallo al agregar filtro {0}: {1}".format(safe_text(getattr(filt, "Name", u"")), safe_text(err)))
    if added:
        try:
            tx.Commit()
        except Exception as err:
            tx.RollBack()
            log_exception(err)
            forms.alert("No se pudo completar la operacion:\n{0}".format(safe_text(err)), title=TITLE, warn_icon=True)
            return
    else:
        tx.RollBack()
        forms.alert("No se pudo agregar ningun filtro.", title=TITLE, warn_icon=True)
        return

    target_name = safe_text(getattr(target_view, "Name", u""))
    added_names = _format_names([safe_text(getattr(f, "Name", u"")) for f in added])
    log_line("Agregados {0} filtro(s) a {1}. Fallidos: {2}".format(
        len(added),
        target_name,
        len(failed)
    ))
    if applied_to_template:
        message = u'{0} filtro(s) agregado(s) a la plantilla "{1}": {2}.'.format(len(added), target_name, added_names)
        if applied_to_primary:
            message += u"\nLa vista dependiente activa heredara el cambio."
    elif applied_to_primary:
        message = u'{0} filtro(s) agregado(s) a la vista principal "{1}": {2}.'.format(len(added), target_name, added_names)
        message += u"\nLas vistas dependientes (incluida la activa) heredaran el filtro."
    else:
        message = u'{0} filtro(s) agregado(s) a la vista "{1}": {2}.'.format(len(added), target_name, added_names)
    if failed:
        message += u"\nNo se pudieron agregar {0} filtro(s): {1}".format(
            len(failed), _format_names([name for name, _ in failed])
        )
    forms.alert(message, title=TITLE, warn_icon=False)


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




