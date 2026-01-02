# -*- coding: utf-8 -*-
"""Print selected views or sheets directly to PDF keeping the print dialog unchanged."""

import os
import re
import sys
import time

import clr
clr.AddReference('System.Drawing')
from System.Drawing.Printing import PrinterSettings

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXT_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..'))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
for _p in (EXT_DIR, LIB_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    View,
    ViewSheet,
    ViewSet,
    PrintRange,
)
from pyrevit import forms

from valtria_lib import get_doc, log_exception


# ---------- Helpers ---------- #

class SelectListItem(object):
    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


def sanitize_name(text):
    return re.sub(r"[\\/:*?\"<>|]", "_", (text or '').strip())[:190]


def describe_view(view):
    if isinstance(view, ViewSheet):
        return u"{} - {}".format(view.SheetNumber or '', view.Name or '')
    return view.Name or ''


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
    if path and not os.path.exists(path):
        os.makedirs(path)


def unique_pdf_path(folder, view):
    base = sanitize_name(describe_view(view)) or 'Vista'
    target = os.path.join(folder, base + '.pdf')
    if not os.path.exists(target):
        return target
    index = 1
    while True:
        candidate = os.path.join(folder, '{}_{:02d}.pdf'.format(base, index))
        if not os.path.exists(candidate):
            return candidate
        index += 1


def wait_for_pdf(path, timeout=120.0, poll=0.25, stable_cycles=3, min_size=1024):
    start = time.time()
    last_size = None
    stable_count = 0

    while time.time() - start < timeout:
        if not os.path.exists(path):
            time.sleep(poll)
            continue

        try:
            current_size = os.path.getsize(path)
        except OSError:
            current_size = 0

        if current_size < min_size:
            last_size = None
            stable_count = 0
            time.sleep(poll)
            continue

        if last_size == current_size:
            stable_count += 1
        else:
            last_size = current_size
            stable_count = 1

        try:
            with open(path, 'rb') as stream:
                header = stream.read(5)
                if not header.startswith(b'%PDF-'):
                    time.sleep(poll)
                    continue

                if stable_count >= stable_cycles:
                    seek_pos = max(0, current_size - 4096)
                    stream.seek(seek_pos)
                    tail = stream.read()
                    if b'%%EOF' in tail:
                        return True
        except Exception:
            pass

        time.sleep(poll)

    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception:
            pass
    return False


def installed_printers():
    names = []
    try:
        for printer in PrinterSettings.InstalledPrinters:
            names.append(str(printer))
    except Exception:
        pass
    return names


# ---------- Print session ---------- #

class TempPrintSession(object):
    def __init__(self, doc):
        self.doc = doc
        self.print_manager = doc.PrintManager
        self._transaction = Transaction(doc, 'Print to PDF session')
        self._transaction.Start()
        self._started = True

        self._original_range = self.print_manager.PrintRange
        self._original_to_file = self.print_manager.PrintToFile
        self._original_file = self.print_manager.PrintToFileName
        self._original_printer = self.print_manager.PrinterName

        self._settings = None
        self._original_views = []
        if self._original_range == PrintRange.Select:
            self._capture_views()

    def _capture_views(self):
        try:
            self._settings = self.print_manager.ViewSheetSetting
            current_set = self._settings.CurrentViewSheetSet
            if current_set:
                self._original_views = [view for view in current_set.Views]
        except Exception:
            self._settings = None
            self._original_views = []

    def ensure_select_range(self):
        self.print_manager.PrintRange = PrintRange.Select
        self.print_manager.Apply()
        if self._settings is None:
            self._settings = self.print_manager.ViewSheetSetting

    def set_printer(self, printer_name):
        if not printer_name:
            return
        current = self.print_manager.PrinterName or ''
        if printer_name == current:
            return
        self.print_manager.SelectNewPrintDriver(printer_name)
        self.print_manager.Apply()

    def set_views(self, views):
        if self._settings is None:
            self._settings = self.print_manager.ViewSheetSetting
        viewset = ViewSet()
        for view in views:
            viewset.Insert(view)
        self._settings.CurrentViewSheetSet.Views = viewset

    def set_filename(self, filepath):
        self.print_manager.PrintToFile = True
        self.print_manager.PrintToFileName = filepath
        self.print_manager.Apply()

    def submit(self):
        self.print_manager.SubmitPrint()

    def dispose(self):
        try:
            if self._settings is not None and self._original_views:
                viewset = ViewSet()
                for view in self._original_views:
                    if view is not None:
                        viewset.Insert(view)
                self._settings.CurrentViewSheetSet.Views = viewset
        except Exception:
            pass
        try:
            if self._original_printer and self.print_manager.PrinterName != self._original_printer:
                self.print_manager.SelectNewPrintDriver(self._original_printer)
        except Exception:
            pass
        try:
            self.print_manager.PrintToFile = self._original_to_file
            self.print_manager.PrintToFileName = self._original_file
            self.print_manager.PrintRange = self._original_range
            self.print_manager.Apply()
        except Exception:
            pass
        if self._started:
            try:
                self._transaction.RollBack()
            except Exception:
                pass


# ---------- User interaction ---------- #

def prompt_printer(print_manager):
    printers = installed_printers()
    if not printers:
        raise RuntimeError('No se encontraron impresoras instaladas en el sistema.')
    current = print_manager.PrinterName or ''
    printers = sorted(printers, key=lambda name: (0 if 'pdf' in name.lower() else 1, name.lower()))
    items = [SelectListItem(name, u"{}{}".format('? ' if name == current else '', name)) for name in printers]
    selection = forms.SelectFromList.show(
        items,
        title='Selecciona la impresora',
        multiselect=False,
        button_name='Usar impresora',
        name_attr='name',
    )
    if not selection:
        raise RuntimeError('Operación cancelada por el usuario.')
    printer_name = getattr(selection[0], 'value', selection[0]) if isinstance(selection, list) else getattr(selection, 'value', selection)
    if 'pdf' not in (printer_name or '').lower():
        answer = forms.alert(
            u"La impresora seleccionada ({0}) no parece generar PDF.\n¿Deseas continuar de todas formas?".format(printer_name),
            title='Print to PDF',
            options=['Continuar', 'Cancelar'],
            warn_icon=True,
        )
        if answer != 'Continuar':
            raise RuntimeError('Impresión cancelada: impresora no PDF.')
    return printer_name


def prompt_views(doc):
    printable = collect_printable_views(doc)
    if not printable:
        forms.alert('No se encontraron vistas imprimibles.', title='Print to PDF')
        return []
    items = [SelectListItem(view, u"{} | {}".format(view.Id.IntegerValue, describe_view(view))) for view in printable]
    selection = forms.SelectFromList.show(
        items,
        multiselect=True,
        title='Selecciona vistas/planos a imprimir',
        name_attr='name',
    )
    return [item.value for item in selection] if selection else []


# ---------- Core printing ---------- #

def print_to_pdf(doc, views, folder, printer_name):
    ok_paths = []
    failed = []
    for view in views:
        session = TempPrintSession(doc)
        target_path = None
        try:
            session.ensure_select_range()
            session.set_printer(printer_name)
            session.set_views([view])
            target_path = unique_pdf_path(folder, view)
            session.set_filename(target_path)
            session.submit()
        except Exception:
            target_path = None
        finally:
            session.dispose()
        if target_path and wait_for_pdf(target_path):
            ok_paths.append(target_path)
        else:
            failed.append(describe_view(view))
    return ok_paths, failed


# ---------- Main ---------- #

def main():
    doc = get_doc()

    try:
        printer_name = prompt_printer(doc.PrintManager)
        views = prompt_views(doc)
        if not views:
            return
        folder = forms.pick_folder()
        if not folder:
            return
        ensure_folder(folder)
        successes, failures = print_to_pdf(doc, views, folder, printer_name)
        if successes:
            message = [u"PDFs generados ({}):".format(len(successes))]
            message.extend([u"  • {}".format(os.path.basename(p)) for p in successes[:10]])
            if len(successes) > 10:
                message.append(u"  … y {} más.".format(len(successes) - 10))
        else:
            message = [u"No se generó ningún PDF."]
        if failures:
            message.append(u"\nNo se pudieron imprimir:")
            message.extend([u"  • {}".format(name) for name in failures])
        forms.alert("\n".join(message), title='Print to PDF')
    except Exception as error:
        log_exception(error)


if __name__ == '__main__':
    main()

