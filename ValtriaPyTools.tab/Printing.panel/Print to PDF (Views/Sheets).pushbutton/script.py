# -*- coding: utf-8 -*-
"""Print selected views or sheets to PDF without altering persistent print sets."""

import os
import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    View,
    ViewSet,
    PrintRange,
    ViewSheet,
)
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    log_exception,
)


def sanitize_name(text):
    return re.sub(r'[\\/:*?"<>|]', '_', text)


def describe_view(view):
    if isinstance(view, ViewSheet):
        return '{0} - {1}'.format(view.SheetNumber, view.Name)
    return view.Name


def collect_printable_views(doc):
    views = []
    for view in FilteredElementCollector(doc).OfClass(View):
        try:
            if not view.CanBePrinted:
                continue
        except Exception:
            continue
        if view.IsTemplate:
            continue
        views.append(view)
    return views


def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)


def print_views(doc, views, folder):
    print_manager = doc.PrintManager
    original_range = print_manager.PrintRange
    original_to_file = print_manager.PrintToFile
    original_file = print_manager.PrintToFileName
    print_manager.PrintRange = PrintRange.Select
    print_manager.Apply()
    try:
        for view in views:
            filename = sanitize_name(describe_view(view)) + '.pdf'
            target = os.path.join(folder, filename)
            ensure_folder(os.path.dirname(target))
            viewset = ViewSet()
            viewset.Insert(view)
            transaction = Transaction(doc, 'Temporary Print ViewSet')
            transaction.Start()
            try:
                setting = print_manager.ViewSheetSetting
                setting.CurrentViewSheetSet.Views = viewset
                transaction.Commit()
            except Exception:
                transaction.RollBack()
                raise
            print_manager.PrintToFile = True
            print_manager.PrintToFileName = target
            print_manager.Apply()
            print_manager.SubmitPrint()
    finally:
        print_manager.PrintToFile = original_to_file
        print_manager.PrintToFileName = original_file
        print_manager.PrintRange = original_range
        print_manager.Apply()


def main():
    doc = get_doc()
    printable = collect_printable_views(doc)
    if not printable:
        forms.alert('No se encontraron vistas imprimibles.', title='Print to PDF')
        return
    items = []
    for view in printable:
        label = describe_view(view)
        items.append(forms.SelectFromListItem(name=label, value=view))
    selected = forms.SelectFromList.show(items, multiselect=True, title='Selecciona vistas/planos a imprimir')
    if not selected:
        return
    folder = forms.pick_folder()
    if not folder:
        return
    ensure_folder(folder)
    views = [item.value for item in selected]
    print_views(doc, views, folder)
    forms.alert('Impresión a PDF completada.', title='Print to PDF')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

