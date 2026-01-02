# -*- coding: utf-8 -*-
"""
Crear Print Set desde Sheets seleccionadas (o selector si no hay seleccion),
SIN dejar el set como activo en el dialogo de impresion.

- Pide nombre (default con fecha/hora)
- Ordena por SheetNumber
- Guarda el ViewSheetSet con SaveAs
- Restaura PrintRange y CurrentViewSheetSet.Views originales

Autor: VALTRIA / DCA.DynamoPython.helper
Compat: pyRevit (IronPython), Revit 2019+
"""
import clr
import traceback
from datetime import datetime

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    PrintRange,
    Transaction,
    ViewSet,
    ViewSheet,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms, revit


class SelectListItem(object):
    """Wrapper with display name for selection dialogs."""

    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


def ask_string(title, prompt, default_text):
    try:
        return forms.ask_for_string(default=default_text, prompt=prompt, title=title)
    except Exception:
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("System.Drawing")
        from System.Drawing import Point, Size
        from System.Windows.Forms import Button, DialogResult, Form, Label, TextBox

        form = Form()
        form.Text = title
        form.FormBorderStyle = 3
        form.StartPosition = 1
        form.ClientSize = Size(420, 140)
        form.MinimizeBox = False
        form.MaximizeBox = False

        label = Label()
        label.Text = prompt
        label.Size = Size(380, 20)
        label.Location = Point(20, 15)
        form.Controls.Add(label)

        textbox = TextBox()
        textbox.Text = default_text or ""
        textbox.Size = Size(380, 25)
        textbox.Location = Point(20, 45)
        form.Controls.Add(textbox)

        ok_button = Button()
        ok_button.Text = "OK"
        ok_button.DialogResult = DialogResult.OK
        ok_button.Location = Point(230, 90)
        form.Controls.Add(ok_button)

        cancel_button = Button()
        cancel_button.Text = "Cancelar"
        cancel_button.DialogResult = DialogResult.Cancel
        cancel_button.Location = Point(320, 90)
        form.Controls.Add(cancel_button)

        form.AcceptButton = ok_button
        form.CancelButton = cancel_button

        return textbox.Text.strip() if form.ShowDialog() == DialogResult.OK else None


def get_selected_sheets_from_browser(doc):
    uidoc = revit.uidoc
    if uidoc is None:
        return []
    element_ids = list(uidoc.Selection.GetElementIds())
    if not element_ids:
        return []
    sheets = []
    for element_id in element_ids:
        element = doc.GetElement(element_id)
        if isinstance(element, ViewSheet) and not element.IsPlaceholder:
            sheets.append(element)
    return sheets


def prompt_user_pick_sheets(doc):
    all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))
    if not all_sheets:
        return []
    items = []
    for sheet in all_sheets:
        label = u"{0} | {1}".format(sheet.SheetNumber or "", sheet.Name or "")
        items.append(SelectListItem(sheet, label))
    picked = forms.SelectFromList.show(
        items,
        title="Selecciona Sheets",
        multiselect=True,
        button_name="Usar estas",
        name_attr='name',
    )
    return [item.value for item in picked] if picked else []


def sheet_sort_key(sheet):
    sheet_number = sheet.SheetNumber or ""
    digits = [ch for ch in sheet_number if ch.isdigit()]
    if digits:
        try:
            return (0, int("".join(digits)), sheet_number)
        except Exception:
            pass
    return (1, sheet_number)


def build_viewset(sheets):
    viewset = ViewSet()
    for sheet in sorted(sheets, key=sheet_sort_key):
        viewset.Insert(sheet)
    return viewset


def clone_current_views(view_sheet_setting):
    cloned = []
    try:
        current = view_sheet_setting.CurrentViewSheetSet
        if current:
            for view in current.Views:
                cloned.append(view)
    except Exception:
        pass
    return cloned


def restore_current_views(view_sheet_setting, views):
    try:
        viewset = ViewSet()
        for view in views:
            if view is not None:
                viewset.Insert(view)
        view_sheet_setting.CurrentViewSheetSet.Views = viewset
    except Exception:
        pass


def delete_set_if_exists(view_sheet_setting, name):
    try:
        view_sheet_setting.Delete(name)
        return True
    except Exception:
        pass
    iterator = None
    try:
        iterator = view_sheet_setting.ViewSheetSets
    except Exception:
        iterator = None
    if iterator is None:
        return False
    try:
        iterator.Reset()
    except Exception:
        pass
    while True:
        try:
            has_next = iterator.MoveNext()
        except Exception:
            break
        if not has_next:
            break
        viewset = None
        try:
            viewset = iterator.Current
        except Exception:
            viewset = None
        if viewset is None:
            continue
        try:
            vs_name = viewset.Name
        except Exception:
            vs_name = None
        if vs_name == name:
            try:
                view_sheet_setting.Delete(viewset)
                return True
            except Exception:
                continue
    return False


def list_existing_print_sets(doc):
    names = []
    try:
        print_manager = doc.PrintManager
    except Exception:
        return names

    original_range = print_manager.PrintRange
    view_sheet_setting = None
    switched_range = False

    try:
        view_sheet_setting = print_manager.ViewSheetSetting
    except Exception:
        view_sheet_setting = None

    if view_sheet_setting is None:
        try:
            print_manager.PrintRange = PrintRange.Select
            print_manager.Apply()
            view_sheet_setting = print_manager.ViewSheetSetting
            switched_range = True
        except Exception:
            view_sheet_setting = None

    if view_sheet_setting is not None:
        iterator = None
        try:
            iterator = view_sheet_setting.ViewSheetSets
        except Exception:
            iterator = None
        if iterator is not None:
            try:
                iterator.Reset()
            except Exception:
                pass
            while True:
                try:
                    has_next = iterator.MoveNext()
                except Exception:
                    break
                if not has_next:
                    break
                viewset = None
                try:
                    viewset = iterator.Current
                except Exception:
                    viewset = None
                if viewset is None:
                    continue
                try:
                    name = viewset.Name
                except Exception:
                    name = None
                if name:
                    names.append(name)

    if switched_range:
        try:
            print_manager.PrintRange = original_range
            print_manager.Apply()
        except Exception:
            pass

    return sorted(set(names))


def prompt_delete_existing_print_sets(doc):
    existing = list_existing_print_sets(doc)
    if not existing:
        return []
    try:
        answer = forms.alert(
            "Quieres eliminar algun Print Set existente antes de crear uno nuevo?",
            title="Crear Print Set",
            options=["Eliminar", "Continuar"],
            exitscript=False,
        )
    except Exception:
        answer = None
    if answer != "Eliminar":
        return []
    items = [SelectListItem(name, name) for name in existing]
    picked = forms.SelectFromList.show(
        items,
        title="Selecciona Print Sets a eliminar",
        multiselect=True,
        button_name="Eliminar",
        name_attr='name',
    )
    return [item.value for item in picked] if picked else []


def delete_print_sets(doc, names):
    if not names:
        return []
    try:
        print_manager = doc.PrintManager
    except Exception:
        return []

    original_range = print_manager.PrintRange
    view_sheet_setting = None
    switched_range = False

    try:
        view_sheet_setting = print_manager.ViewSheetSetting
    except Exception:
        view_sheet_setting = None

    if view_sheet_setting is None:
        try:
            print_manager.PrintRange = PrintRange.Select
            print_manager.Apply()
            view_sheet_setting = print_manager.ViewSheetSetting
            switched_range = True
        except Exception:
            view_sheet_setting = None

    deleted = []
    if view_sheet_setting is not None:
        for name in names:
            if name and delete_set_if_exists(view_sheet_setting, name):
                deleted.append(name)

    if switched_range:
        try:
            print_manager.PrintRange = original_range
            print_manager.Apply()
        except Exception:
            pass

    return deleted


def save_print_set_without_changing_current(doc, viewset, set_name):
    print_manager = doc.PrintManager
    original_range = print_manager.PrintRange
    restore_allowed = (original_range == PrintRange.Select)

    view_sheet_setting = None
    original_views = []

    if restore_allowed:
        try:
            view_sheet_setting = print_manager.ViewSheetSetting
            original_views = clone_current_views(view_sheet_setting)
        except Exception:
            view_sheet_setting = None
            original_views = []

    if view_sheet_setting is None:
        print_manager.PrintRange = PrintRange.Select
        print_manager.Apply()
        view_sheet_setting = print_manager.ViewSheetSetting

    try:
        view_sheet_setting.CurrentViewSheetSet.Views = viewset
        try:
            view_sheet_setting.SaveAs(set_name)
        except Exception:
            delete_set_if_exists(view_sheet_setting, set_name)
            view_sheet_setting.SaveAs(set_name)
    finally:
        if restore_allowed and view_sheet_setting is not None:
            restore_current_views(view_sheet_setting, original_views)
        try:
            print_manager.PrintRange = original_range
            print_manager.Apply()
        except Exception:
            pass


def format_sheet_list(viewset):
    names = []
    for sheet in viewset:
        try:
            number = sheet.SheetNumber or ''
            name = sheet.Name or ''
            names.append(u"{0} | {1}".format(number, name))
        except Exception:
            names.append(str(sheet))
    return names


def main():
    doc = revit.doc
    if doc is None:
        TaskDialog.Show("Crear Print Set", "No hay documento activo.")
        return
    try:
        sheets = get_selected_sheets_from_browser(doc)
        if not sheets:
            sheets = prompt_user_pick_sheets(doc)
        if not sheets:
            TaskDialog.Show("Crear Print Set", "No hay Sheets seleccionadas ni elegidas en el selector.")
            return

        sets_to_delete = prompt_delete_existing_print_sets(doc)

        default_name = "VAL_{0}".format(datetime.now().strftime("%Y%m%d-%H%M%S"))
        if len(sets_to_delete) == 1:
            default_name = sets_to_delete[0]
        set_name = ask_string("Crear Print Set", "Nombre del Print Set:", default_name)
        if not set_name or not set_name.strip():
            TaskDialog.Show("Crear Print Set", "Operacion cancelada. No se proporciono nombre.")
            return
        set_name = set_name.strip()

        viewset = build_viewset(sheets)

        transaction = Transaction(doc, "Crear Print Set (sin activar)")
        deleted_sets = []
        try:
            transaction.Start()
            if sets_to_delete:
                deleted_sets = delete_print_sets(doc, sets_to_delete)
            save_print_set_without_changing_current(doc, viewset, set_name)
            transaction.Commit()
        except Exception:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            raise

        sheet_lines = "\n".join(format_sheet_list(viewset))
        message = "Print Set creado (no activado):\n- Nombre: {0}\n- N\u00ba hojas: {1}\n\nHojas incluidas:\n{2}".format(
            set_name,
            viewset.Size,
            sheet_lines,
        )
        if deleted_sets:
            message += "\n\nPrint Sets eliminados:\n{0}".format("\n".join("- {0}".format(name) for name in deleted_sets))
        TaskDialog.Show("Crear Print Set", message)
    except Exception as exc:
        try:
            if 'transaction' in locals() and transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
        except Exception:
            pass
        TaskDialog.Show(
            "Crear Print Set - Error",
            u"{0}\n\nTrace:\n{1}".format(exc, traceback.format_exc()),
        )


if __name__ == "__main__":
    main()


