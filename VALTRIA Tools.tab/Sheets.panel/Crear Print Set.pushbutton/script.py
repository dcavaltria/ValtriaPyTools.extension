# -*- coding: utf-8 -*-
"""
Crear Print Set desde Sheets seleccionadas (o selector si no hay selección),
SIN dejar el set como activo en el Print dialog.

- Pide nombre (default con fecha/hora)
- Ordena por SheetNumber
- Guarda el ViewSheetSet con SaveAs
- Restaura PrintRange y CurrentViewSheetSet.Views originales

Autor: VALTRIA / DCA.DynamoPython.helper
Compat: pyRevit (IronPython), Revit 2019+
"""
import clr, traceback
from datetime import datetime

from pyrevit import revit, DB, forms
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    ViewSheet, ViewSet, Transaction, PrintRange, FilteredElementCollector
)

# ---------- UI helpers ----------
def ask_string(title, prompt, default_text):
    try:
        return forms.ask_for_string(default=default_text, prompt=prompt, title=title)
    except:
        # WinForms fallback
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("System.Drawing")
        from System.Windows.Forms import (Form, Label, TextBox, Button, DialogResult)
        from System.Drawing import Size, Point, ContentAlignment
        f = Form(); f.Text = title; f.FormBorderStyle = 3; f.StartPosition = 1
        f.ClientSize = Size(420,140); f.MinimizeBox = False; f.MaximizeBox = False
        lab = Label(); lab.Text = prompt; lab.Size = Size(380,20); lab.Location = Point(20,15)
        f.Controls.Add(lab)
        tb = TextBox(); tb.Text = default_text or ""; tb.Size = Size(380,25); tb.Location = Point(20,45)
        f.Controls.Add(tb)
        ok = Button(); ok.Text = "OK"; ok.DialogResult = DialogResult.OK; ok.Location = Point(230,90)
        f.Controls.Add(ok)
        ca = Button(); ca.Text = "Cancelar"; ca.DialogResult = DialogResult.Cancel; ca.Location = Point(320,90)
        f.Controls.Add(ca)
        f.AcceptButton = ok; f.CancelButton = ca
        return tb.Text.strip() if f.ShowDialog() == DialogResult.OK else None

# ---------- selección ----------
def get_selected_sheets_from_browser(doc):
    uidoc = revit.uidoc
    if uidoc is None:
        return []
    ids = list(uidoc.Selection.GetElementIds())
    if not ids:
        return []
    out = []
    for eid in ids:
        el = doc.GetElement(eid)
        if isinstance(el, ViewSheet) and not el.IsPlaceholder:
            out.append(el)
    return out

def prompt_user_pick_sheets(doc):
    all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
    if not all_sheets:
        return []
    items = [forms.TemplateListItem(s, name=u"{0} | {1}".format(s.SheetNumber, s.Name))
             for s in all_sheets]
    picked = forms.SelectFromList.show(items, title="Selecciona Sheets", multiselect=True, button_name="Usar estas")
    return [p.item for p in picked] if picked else []

# ---------- util ----------
def sheet_sort_key(sheet):
    sn = sheet.SheetNumber or ""
    digits = [ch for ch in sn if ch.isdigit()]
    if digits:
        try:
            return (0, int("".join(digits)), sn)
        except:
            pass
    return (1, sn)

def build_viewset(sheets):
    vset = ViewSet()
    for s in sorted(sheets, key=sheet_sort_key):
        vset.Insert(s)
    return vset

def clone_current_views(vss):
    """Devuelve lista de Views actualmente seleccionadas en el print dialog."""
    original = []
    try:
        curset = vss.CurrentViewSheetSet
        if curset:
            it = curset.Views  # ViewSet enumerable
            for v in it:
                original.append(v)
    except:
        pass
    return original

def restore_current_views(vss, views_list):
    """Restaura las vistas previas como 'CurrentViewSheetSet'."""
    try:
        vs = ViewSet()
        for v in views_list:
            if v: vs.Insert(v)
        vss.CurrentViewSheetSet.Views = vs
    except:
        pass

def delete_set_if_exists(vss, name):
    """Borra un set existente sin depender de ViewSheetSets (seguro entre versiones)."""
    # 1) Intento directo por nombre (si la API lo permite en tu versión)
    try:
        vss.Delete(name)
        return True
    except:
        pass
    # 2) Si existe la colección, intenta por objeto
    try:
        sets_iter = vss.ViewSheetSets  # puede no existir
        for s in sets_iter:
            try:
                if s.Name == name:
                    vss.Delete(s)
                    return True
            except:
                continue
    except:
        pass
    return False

def save_print_set_without_changing_current(doc, vset, set_name):
    """Guarda el print set y RESTAURA el estado original del print dialog."""
    pm = doc.PrintManager
    vss = pm.ViewSheetSetting

    # --- snapshot estado original ---
    original_range = pm.PrintRange
    original_views = clone_current_views(vss)

    try:
        # usar Select de forma temporal para poder asignar Views
        pm.PrintRange = PrintRange.Select
        vss.CurrentViewSheetSet.Views = vset

        # primer intento: guardar
        try:
            vss.SaveAs(set_name)
        except:
            # borrar si ya existe y reintentar
            delete_set_if_exists(vss, set_name)
            vss.SaveAs(set_name)
    finally:
        # --- RESTAURAR estado original pase lo que pase ---
        try:
            restore_current_views(vss, original_views)
        except:
            pass
        try:
            pm.PrintRange = original_range
        except:
            pass

# ---------- main ----------
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
            TaskDialog.Show("Crear Print Set",
                            "No hay Sheets seleccionadas ni elegidas en el selector.")
            return

        default_name = "VAL_{0}".format(datetime.now().strftime("%Y%m%d-%H%M%S"))
        set_name = ask_string("Crear Print Set", "Nombre del Print Set:", default_name)
        if not set_name or not set_name.strip():
            TaskDialog.Show("Crear Print Set", "Operación cancelada. No se proporcionó nombre.")
            return
        set_name = set_name.strip()

        vset = build_viewset(sheets)

        t = Transaction(doc, "Crear Print Set (sin activar)")
        t.Start()
        save_print_set_without_changing_current(doc, vset, set_name)
        t.Commit()

        TaskDialog.Show("Crear Print Set",
                        "Print Set creado (no activado):\n- Nombre: {0}\n- Nº hojas: {1}".format(set_name, vset.Size))
    except Exception as exc:
        try:
            if 't' in locals() and t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except:
            pass
        TaskDialog.Show("Crear Print Set - Error",
                        u"{0}\n\nTrace:\n{1}".format(exc, traceback.format_exc()))

if __name__ == "__main__":
    main()
