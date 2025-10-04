# -*- coding: utf-8 -*-
"""Export selected views to IFC files."""

import os
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    View,
    ViewSet,
    IFCExportOptions,
)
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    log_exception,
)


def is_exportable(view):
    if view.IsTemplate:
        return False
    if getattr(view, 'ViewType', None) is None:
        return False
    if view.ViewType.ToString() in ['Schedule', 'ProjectBrowser', 'DrawingSheet']:
        return False
    if view.ViewType.ToString() == 'Legend':
        return False
    return True


def clean_name(name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    if not safe:
        safe = 'View'
    return safe


def main():
    doc = get_doc()
    collector = FilteredElementCollector(doc).OfClass(View)
    options = []
    for view in collector:
        if is_exportable(view):
            label = '{0} ({1})'.format(view.Name, view.ViewType)
            options.append(forms.SelectFromListItem(name=label, value=view))
    if not options:
        forms.alert('No hay vistas de modelo disponibles para exportar.', title='IFC Export Views')
        return
    selected = forms.SelectFromList.show(options, multiselect=True, title='Selecciona vistas a exportar')
    if not selected:
        return
    folder = forms.pick_folder()
    if not folder:
        return
    if not os.path.exists(folder):
        os.makedirs(folder)
    errors = []
    for item in selected:
        view = item.value
        viewset = ViewSet()
        viewset.Insert(view)
        file_name = clean_name(view.Name)
        try:
            # TODO: Implement temporal isolation to ensure view-specific IFC exports.
            if not doc.Export(folder, file_name, viewset, IFCExportOptions()):
                errors.append(view.Name)
        except Exception as export_error:
            log_exception(export_error)
            errors.append(view.Name)
    if errors:
        forms.alert('Algunas vistas no se exportaron: {0}'.format(', '.join(errors)), title='IFC Export Views', warn_icon=True)
    else:
        forms.alert('Exportación IFC completada.', title='IFC Export Views')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

