# -*- coding: utf-8 -*-
"""Auto-fit section box to selection or visible elements in active 3D view."""

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..'))
LIB_DIR = os.path.join(EXTENSION_DIR, 'lib')
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    FilteredElementCollector,
    Transaction,
    View3D,
    XYZ,
)
from pyrevit import forms

from valtria_lib import (
    get_doc,
    get_uidoc,
    get_all_visible_model_boundingbox,
    mm_to_internal,
    log_exception,
)


DEFAULT_PADDING_MM = 150.0


try:
    unicode
except NameError:
    unicode = str  # type: ignore


def expand_bbox(bbox, padding):
    if bbox is None:
        return None
    min_point = bbox.Min
    max_point = bbox.Max
    pad = padding
    new_box = BoundingBoxXYZ()
    new_box.Min = XYZ(min_point.X - pad, min_point.Y - pad, min_point.Z - pad)
    new_box.Max = XYZ(max_point.X + pad, max_point.Y + pad, max_point.Z + pad)
    return new_box


def get_selected_elements(doc, uidoc):
    element_ids = uidoc.Selection.GetElementIds()
    elements = []
    for elid in element_ids:
        element = doc.GetElement(elid)
        if element is not None:
            elements.append(element)
    return elements


def ensure_text(value):
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


def collect_available_3d_views(doc):
    collector = FilteredElementCollector(doc).OfClass(View3D)
    views = []
    for view in collector:
        if view is None:
            continue
        if getattr(view, 'IsTemplate', False):
            continue
        views.append(view)
    return views


def _normalized_view_name(view):
    name = ensure_text(getattr(view, 'Name', u""))
    return name.strip().lower()


def get_default_user_view(views, active_view):
    if not views:
        return None
    for view in views:
        name = _normalized_view_name(view)
        if '3d' in name and ('usuario' in name or 'user' in name):
            return view
    for view in views:
        name = _normalized_view_name(view).replace(' ', '')
        if name == '{3d}':
            return view
    if isinstance(active_view, View3D):
        for view in views:
            if view.Id == active_view.Id:
                return view
    return views[0]


class ViewSelectionItem(object):
    """Adapter item so SelectFromList can show View3D options."""

    def __init__(self, view, is_default=False):
        self.value = view
        label = ensure_text(getattr(view, 'Name', u"")).strip() or u"(sin nombre)"
        if getattr(view, 'IsPerspective', False):
            label = u"{0} (Perspectiva)".format(label)
        if is_default:
            label = u"{0}  [Predeterminada]".format(label)
        self.label = label

    @property
    def name(self):
        return self.label


def prompt_view_selection(views, default_view):
    if not views:
        return None
    if len(views) == 1:
        return views[0]

    def sort_key(view):
        return (
            0 if default_view and view.Id == default_view.Id else 1,
            _normalized_view_name(view),
        )

    ordered = sorted(views, key=sort_key)
    items = []
    for view in ordered:
        items.append(ViewSelectionItem(view, is_default=(default_view and view.Id == default_view.Id)))
    picked = forms.SelectFromList.show(
        items,
        title='Selecciona la vista 3D',
        multiselect=False,
        button_name='Usar vista',
        name_attr='name',
    )
    if not picked:
        return None
    return picked.value if hasattr(picked, 'value') else picked


def prompt_padding_mm(default_mm):
    default_m = default_mm / 1000.0
    default_text = "{0:.3f}".format(default_m).rstrip('0').rstrip('.')
    if not default_text:
        default_text = "0"
    response = forms.ask_for_string(
        default=default_text,
        prompt='Introduce el desfase (en metros) a aplicar al Section Box:',
        title='Auto-Fit Section Box',
    )
    if response is None:
        return None
    response = response.strip()
    if not response:
        return default_mm
    normalized = response.replace(',', '.')
    try:
        value_m = float(normalized)
    except Exception:
        forms.alert('Introduce un numero valido para el desfase (por ejemplo 1.0).', title='Auto-Fit Section Box', warn_icon=True)
        return None
    if value_m < 0:
        forms.alert('El desfase debe ser mayor o igual a 0.', title='Auto-Fit Section Box', warn_icon=True)
        return None
    return value_m * 1000.0


def main():
    doc = get_doc()
    uidoc = get_uidoc()
    available_views = collect_available_3d_views(doc)
    if not available_views:
        forms.alert('No se encontraron vistas 3D disponibles.', title='Auto-Fit Section Box', warn_icon=True)
        return
    default_view = get_default_user_view(available_views, doc.ActiveView)
    target_view = prompt_view_selection(available_views, default_view)
    if target_view is None:
        return
    padding_mm = prompt_padding_mm(DEFAULT_PADDING_MM)
    if padding_mm is None:
        return
    selection = get_selected_elements(doc, uidoc)
    if selection:
        bbox = get_all_visible_model_boundingbox(doc, target_view, elements=selection)
    else:
        bbox = get_all_visible_model_boundingbox(doc, target_view)
    if bbox is None:
        forms.alert('No se pudo calcular un contorno para la vista.', title='Auto-Fit Section Box', warn_icon=True)
        return
    padding = mm_to_internal(padding_mm)
    new_box = expand_bbox(bbox, padding)
    transaction = Transaction(doc, 'Auto-Fit Section Box')
    transaction.Start()
    try:
        if not target_view.IsSectionBoxActive:
            target_view.IsSectionBoxActive = True
        target_view.SetSectionBox(new_box)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise
    message = u'Section Box ajustado correctamente en "{0}".'.format(ensure_text(getattr(target_view, 'Name', u"")))
    forms.alert(message, title='Auto-Fit Section Box', warn_icon=False)


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)


