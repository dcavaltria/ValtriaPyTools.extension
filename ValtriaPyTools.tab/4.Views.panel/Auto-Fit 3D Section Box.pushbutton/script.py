# -*- coding: utf-8 -*-
"""Auto-fit section box to selection or visible elements in active 3D view."""

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..'))
LIB_DIR = os.path.join(EXTENSION_DIR, '_lib')
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    Transaction,
    View3D,
    XYZ,
)
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    get_uidoc,
    get_all_visible_model_boundingbox,
    mm_to_internal,
    log_exception,
)


PADDING_MM = 150.0


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


def main():
    doc = get_doc()
    uidoc = get_uidoc()
    view = doc.ActiveView
    if not isinstance(view, View3D):
        forms.alert('Activa una vista 3D antes de ejecutar la herramienta.', title='Auto-Fit Section Box', warn_icon=True)
        return
    selection = get_selected_elements(doc, uidoc)
    if selection:
        bbox = get_all_visible_model_boundingbox(doc, view, elements=selection)
    else:
        bbox = get_all_visible_model_boundingbox(doc, view)
    if bbox is None:
        forms.alert('No se pudo calcular un contorno para la vista.', title='Auto-Fit Section Box', warn_icon=True)
        return
    padding = mm_to_internal(PADDING_MM)
    new_box = expand_bbox(bbox, padding)
    transaction = Transaction(doc, 'Auto-Fit Section Box')
    transaction.Start()
    try:
        if not view.IsSectionBoxActive:
            view.IsSectionBoxActive = True
        view.SetSectionBox(new_box)
        transaction.Commit()
    except Exception:
        transaction.RollBack()
        raise
    forms.alert('Section Box ajustado correctamente.', title='Auto-Fit Section Box')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

