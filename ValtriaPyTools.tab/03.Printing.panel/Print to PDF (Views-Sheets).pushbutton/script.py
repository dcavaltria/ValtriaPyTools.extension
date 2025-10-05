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
    return re.sub(r"[\\/:*?\"<>|]", "_", text)


def describe_view(view):
    if isinstance(view, ViewSheet):
        return "{0} - {1}".format(view.SheetNumber, view.Name)
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


class TempPrintSession(object):
    """Encapsulate temporary print settings and ensure restoration."""

    def __init__(self, doc):
        self.doc = doc
        self.print_manager = doc.PrintManager
        self._settings = self.print_manager.ViewSheetSetting
        self._transaction = Transaction(doc, 'Temp Print Session')
        self._started = False
        self._transaction.Start()
        self._started = True
        self._original_range = self.print_manager.PrintRange
        self._original_to_file = self.print_manager.PrintToFile
        self._original_file = self.print_manager.PrintToFileName
        self._original_views = self._clone_current_views()

    def _clone_current_views(self):
        results = []
        try:
            current_set = self._settings.CurrentViewSheetSet
            if current_set:
                for view in current_set.Views:
                    results.append(view)
        except Exception:
            pass
        return results

    def _restore_current_views(self):
        try:
            viewset = ViewSet()
            for view in self._original_views:
                if view is not None:
                    viewset.Insert(view)
            self._settings.CurrentViewSheetSet.Views = viewset
        except Exception:
            pass

    def dispose(self):
        try:
            self.print_manager.PrintRange = self._original_range
        except Exception:
            pass
        try:
            self.print_manager.PrintToFile = self._original_to_file
            self.print_manager.PrintToFileName = self._original_file
        except Exception:
            pass
        self._restore_current_views()
        try:
            self.print_manager.Apply()
        except Exception:
            pass
        if self._started:
            try:
                self._transaction.RollBack()
            except Exception:
                pass


def print_views(doc, views, folder):
    session = TempPrintSession(doc)
    print_manager = session.print_manager
    print_manager.PrintRange = PrintRange.Select
    print_manager.Apply()
    try:
        for view in views:
            filename = sanitize_name(describe_view(view)) + '.pdf'
            target = os.path.join(folder, filename)
            ensure_folder(os.path.dirname(target))
            viewset = ViewSet()
            viewset.Insert(view)
            session._settings.CurrentViewSheetSet.Views = viewset
            print_manager.PrintToFile = True
            print_manager.PrintToFileName = target
            print_manager.Apply()
            print_manager.SubmitPrint()
    finally:
        session.dispose()


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
