# -*- coding: utf-8 -*-
"""
Renombrar sheet number y sheet name de varias hojas utilizando prefijo,
sufijo o busqueda y reemplazo.

Pasos:
- Toma la seleccion actual y permite confirmarla.
- Solicita el modo de renombrado y los textos necesarios.
- Muestra una vista previa y pide confirmacion.
- Aplica los cambios en una unica transaccion.
"""

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, ViewSheet
from pyrevit import forms, revit
from valtria_core.text import ensure_text as safe_text


TITLE = "Renombrar hojas"
MODE_PREFIX = "prefix"
MODE_SUFFIX = "suffix"
MODE_REPLACE = "replace"
TARGET_NUMBER = "number"
TARGET_NAME = "name"
MAX_PREVIEW_LINES = 40



class SimpleItem(object):
    """Item generico para SelectFromList."""

    def __init__(self, value, label, description=""):
        self.value = value
        self.label = label
        self._description = description

    @property
    def name(self):
        return self.label

    @property
    def description(self):
        return self._description


def format_sheet_label(sheet):
    number = safe_text(getattr(sheet, "SheetNumber", u"")).strip() or u"(sin numero)"
    name = safe_text(getattr(sheet, "Name", u"")).strip() or u"(sin nombre)"
    return u"{0} | {1}".format(number, name)


def collect_selected_sheets(doc):
    selection = revit.get_selection()
    if not selection:
        return []
    sheets = []
    seen_ids = set()
    for elem_id in selection.element_ids:
        if elem_id in seen_ids:
            continue
        seen_ids.add(elem_id)
        element = doc.GetElement(elem_id)
        if isinstance(element, ViewSheet):
            sheets.append(element)
    return sheets


def prompt_sheet_selection(sheets):
    if not sheets:
        return []
    items = [
        SimpleItem(sheet, format_sheet_label(sheet))
        for sheet in sorted(sheets, key=lambda x: (safe_text(x.SheetNumber), safe_text(x.Name)))
    ]
    picked = forms.SelectFromList.show(
        items,
        title=TITLE + " - Seleccion actual",
        multiselect=True,
        button_name="Continuar",
        name_attr="name",
    )
    if not picked:
        return []
    if not isinstance(picked, list):
        picked = [picked]
    return [item.value for item in picked]


def prompt_mode():
    options = [
        SimpleItem(MODE_PREFIX, "Agregar prefijo"),
        SimpleItem(MODE_SUFFIX, "Agregar sufijo"),
        SimpleItem(MODE_REPLACE, "Buscar y reemplazar"),
    ]
    picked = forms.SelectFromList.show(
        options,
        title=TITLE + " - Modo",
        multiselect=False,
        button_name="Seleccionar",
        name_attr="name",
    )
    if not picked:
        return None
    if isinstance(picked, list):
        picked = picked[0]
    return picked.value


def prompt_targets(allow_multiple=True):
    options = [
        SimpleItem(TARGET_NUMBER, "Sheet Number"),
        SimpleItem(TARGET_NAME, "Sheet Name"),
    ]
    picked = forms.SelectFromList.show(
        options,
        title=TITLE + " - Propiedad",
        multiselect=allow_multiple,
        button_name="Continuar",
        name_attr="name",
    )
    if not picked:
        return []
    if allow_multiple:
        if not isinstance(picked, list):
            picked = [picked]
        return [item.value for item in picked]
    if isinstance(picked, list):
        picked = picked[0]
    return [picked.value]


def prompt_prefix_suffix(mode):
    prompt = "Ingresa el prefijo a agregar:" if mode == MODE_PREFIX else "Ingresa el sufijo a agregar:"
    text = forms.ask_for_string(
        default="",
        prompt=prompt,
        title=TITLE,
    )
    if text is None:
        return None
    return safe_text(text)


def prompt_find_replace():
    search = forms.ask_for_string(
        default="",
        prompt="Texto a buscar:",
        title=TITLE,
    )
    if search is None or search == "":
        return None, None
    replace = forms.ask_for_string(
        default="",
        prompt="Texto de reemplazo:",
        title=TITLE,
    )
    if replace is None:
        return None, None
    return safe_text(search), safe_text(replace)


def compute_new_value(original, mode, data):
    base = safe_text(original)
    if mode == MODE_PREFIX:
        return safe_text(data.get("text", u"")) + base
    if mode == MODE_SUFFIX:
        return base + safe_text(data.get("text", u""))
    if mode == MODE_REPLACE:
        search = safe_text(data.get("search", u""))
        replace = safe_text(data.get("replace", u""))
        if not search:
            return base
        return base.replace(search, replace)
    return base


def build_proposals(sheets, mode, targets, data):
    proposals = []
    for sheet in sheets:
        number_old = safe_text(sheet.SheetNumber).strip()
        name_old = safe_text(sheet.Name).strip()
        number_new = number_old
        name_new = name_old

        if TARGET_NUMBER in targets:
            number_new = safe_text(compute_new_value(number_old, mode, data)).strip()
        if TARGET_NAME in targets:
            name_new = safe_text(compute_new_value(name_old, mode, data)).strip()

        if number_new != number_old or name_new != name_old:
            proposals.append(
                {
                    "sheet": sheet,
                    "number_old": number_old,
                    "number_new": number_new,
                    "name_old": name_old,
                    "name_new": name_new,
                }
            )
    return proposals


def split_conflicts(doc, proposals, targets):
    if not proposals:
        return [], []

    valid = []
    conflicts = []

    existing_by_number = {}
    if TARGET_NUMBER in targets:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            number = safe_text(sheet.SheetNumber).strip()
            if number:
                existing_by_number[number] = sheet

    assigned_numbers = {}

    for proposal in proposals:
        issues = []
        sheet = proposal["sheet"]
        if TARGET_NUMBER in targets:
            number_new = proposal["number_new"]
            if not number_new:
                issues.append("El sheet number no puede quedar vacio.")
            else:
                owner = existing_by_number.get(number_new)
                if owner is not None and owner.Id != sheet.Id:
                    issues.append("Conflicto con {0}".format(format_sheet_label(owner)))
                other = assigned_numbers.get(number_new)
                if other is not None and other.Id != sheet.Id:
                    issues.append("Conflicto con {0}".format(format_sheet_label(other)))
                assigned_numbers[number_new] = sheet
        if TARGET_NAME in targets:
            name_new = proposal["name_new"]
            if not name_new:
                issues.append("El sheet name no puede quedar vacio.")

        if issues:
            conflict = dict(proposal)
            conflict["issues"] = issues
            conflicts.append(conflict)
        else:
            valid.append(proposal)
    return valid, conflicts


def format_preview_lines(accepted, conflicts):
    lines = []
    if accepted:
        lines.append("Cambios propuestos ({0}):".format(len(accepted)))
        for idx, proposal in enumerate(accepted):
            if idx >= MAX_PREVIEW_LINES:
                lines.append("... ({0} cambios adicionales no listados)".format(len(accepted) - MAX_PREVIEW_LINES))
                break
            lines.append(
                "- {0} -> {1} | {2}".format(
                    "{0} | {1}".format(proposal["number_old"] or "-", proposal["name_old"] or "-"),
                    proposal["number_new"] or "-",
                    proposal["name_new"] or "-",
                )
            )
    if conflicts:
        if lines:
            lines.append("")
        lines.append("Conflictos detectados ({0}):".format(len(conflicts)))
        for idx, proposal in enumerate(conflicts):
            if idx >= MAX_PREVIEW_LINES:
                lines.append("... ({0} conflictos adicionales no listados)".format(len(conflicts) - MAX_PREVIEW_LINES))
                break
            issue_text = "; ".join(proposal.get("issues", []))
            lines.append(
                "- {0} -> {1} | {2} [{3}]".format(
                    format_sheet_label(proposal["sheet"]),
                    proposal["number_new"] or "-",
                    proposal["name_new"] or "-",
                    issue_text or "sin detalle",
                )
            )
    return lines


def apply_changes(doc, accepted, targets, mode_label):
    if not accepted:
        return [], []

    transaction = Transaction(doc, "Renombrar hojas ({0})".format(mode_label))
    renamed = []
    failures = []
    try:
        transaction.Start()
    except Exception as err:
        forms.alert("No se pudo iniciar la transaccion:\n{0}".format(safe_text(err)), title=TITLE)
        return [], []

    for proposal in accepted:
        sheet = proposal["sheet"]
        try:
            if TARGET_NUMBER in targets and proposal["number_old"] != proposal["number_new"]:
                sheet.SheetNumber = proposal["number_new"]
            if TARGET_NAME in targets and proposal["name_old"] != proposal["name_new"]:
                sheet.Name = proposal["name_new"]
            renamed.append(proposal)
        except Exception as err:
            failures.append((proposal, safe_text(err)))

    try:
        if renamed:
            transaction.Commit()
        else:
            transaction.RollBack()
    except Exception as err:
        forms.alert("Error al finalizar la transaccion:\n{0}".format(safe_text(err)), title=TITLE)
        return [], failures

    return renamed, failures


def summarize_results(renamed, failures, conflicts):
    lines = []
    if renamed:
        lines.append("Hojas actualizadas: {0}".format(len(renamed)))
        for idx, proposal in enumerate(renamed):
            if idx >= MAX_PREVIEW_LINES:
                lines.append("... ({0} cambios adicionales no listados)".format(len(renamed) - MAX_PREVIEW_LINES))
                break
            lines.append(
                "- {0} -> {1} | {2}".format(
                    format_sheet_label(proposal["sheet"]),
                    proposal["number_new"] or "-",
                    proposal["name_new"] or "-",
                )
            )
    if failures:
        if lines:
            lines.append("")
        lines.append("Errores durante la actualizacion: {0}".format(len(failures)))
        for idx, (proposal, reason) in enumerate(failures):
            if idx >= MAX_PREVIEW_LINES:
                lines.append("... ({0} errores adicionales no listados)".format(len(failures) - MAX_PREVIEW_LINES))
                break
            lines.append(
                "- {0} -> {1} | {2} | {3}".format(
                    format_sheet_label(proposal["sheet"]),
                    proposal["number_new"] or "-",
                    proposal["name_new"] or "-",
                    reason,
                )
            )
    if conflicts:
        if lines:
            lines.append("")
        lines.append("Omitidas por conflicto previo: {0}".format(len(conflicts)))
    if not lines:
        lines.append("No se realizaron cambios.")
    return lines


def get_mode_label(mode):
    if mode == MODE_PREFIX:
        return "prefijo"
    if mode == MODE_SUFFIX:
        return "sufijo"
    if mode == MODE_REPLACE:
        return "busqueda y reemplazo"
    return mode


def main():
    doc = revit.doc
    if doc is None:
        forms.alert("No hay documento activo.", title=TITLE)
        return

    sheets = collect_selected_sheets(doc)
    if not sheets:
        forms.alert("Selecciona al menos una hoja antes de ejecutar.", title=TITLE)
        return

    confirmed_sheets = prompt_sheet_selection(sheets)
    if not confirmed_sheets:
        forms.alert("Operacion cancelada. No hay hojas seleccionadas.", title=TITLE)
        return

    mode = prompt_mode()
    if mode is None:
        forms.alert("Operacion cancelada. No se selecciono modo.", title=TITLE)
        return

    allow_multiple_targets = mode != MODE_REPLACE
    targets = prompt_targets(allow_multiple=allow_multiple_targets)
    if not targets:
        forms.alert("Operacion cancelada. No se selecciono propiedad a modificar.", title=TITLE)
        return

    data = {}
    if mode in (MODE_PREFIX, MODE_SUFFIX):
        text_value = prompt_prefix_suffix(mode)
        if text_value is None:
            forms.alert("Operacion cancelada. No se proporciono texto.", title=TITLE)
            return
        if text_value == "":
            forms.alert("Operacion cancelada. El texto no puede quedar vacio.", title=TITLE)
            return
        data["text"] = text_value
    else:
        search, replace = prompt_find_replace()
        if search is None:
            forms.alert("Operacion cancelada. No se proporciono texto de busqueda.", title=TITLE)
            return
        data["search"] = search
        data["replace"] = replace if replace is not None else u""

    proposals = build_proposals(confirmed_sheets, mode, targets, data)
    if not proposals:
        forms.alert("No se detectaron cambios con los valores indicados.", title=TITLE)
        return

    accepted, conflicts = split_conflicts(doc, proposals, targets)
    preview_lines = format_preview_lines(accepted, conflicts)
    if not preview_lines:
        forms.alert("No hay cambios disponibles para aplicar.", title=TITLE)
        return

    preview_lines.append("")
    preview_lines.append("Se aplicaran solo los cambios sin conflicto.")
    preview_lines.append("Aplicar cambios?")

    response = forms.alert(
        u"\n".join(preview_lines),
        title=TITLE,
        yes=True,
        no=True,
        warn_icon=bool(conflicts),
    )
    if not response:
        forms.alert("Operacion cancelada por el usuario.", title=TITLE)
        return

    renamed, failures = apply_changes(doc, accepted, targets, get_mode_label(mode))
    summary_lines = summarize_results(renamed, failures, conflicts)
    forms.alert(u"\n".join(summary_lines), title=TITLE, warn_icon=bool(failures or conflicts))


if __name__ == "__main__":
    main()
