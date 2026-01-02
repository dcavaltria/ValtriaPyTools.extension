# -*- coding: utf-8 -*-
"""
Renombrar tipos por categoría mediante búsqueda y reemplazo.

Flujo:
- Solicita la categoría a procesar.
- Pide texto a buscar y su reemplazo.
- Muestra el resumen de cambios detectados (incluye conflictos).
- Solicita confirmación antes de renombrar.

Compatibilidad: pyRevit (IronPython), Revit 2019+
"""
import re

from Autodesk.Revit.DB import (
    BuiltInParameter,
    FilteredElementCollector,
    Transaction,
)
from pyrevit import forms, revit


TITLE = "Renombrar tipos"
MAX_PREVIEW_LINES = 40

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


def get_category_name(category):
    name = safe_text(getattr(category, "Name", u""))
    return name or u"(Sin categoría)"


def get_family_name(element_type):
    try:
        family = getattr(element_type, "Family", None)
        if family is not None:
            name = safe_text(getattr(family, "Name", u""))
            if name:
                return name
    except Exception:
        pass
    try:
        param = element_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if param is not None and param.HasValue:
            name = safe_text(param.AsString())
            if name:
                return name
    except Exception:
        pass
    return u""


def get_type_name(element_type):
    name = safe_text(getattr(element_type, "Name", u""))
    if name:
        return name
    try:
        param = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if param is not None and param.HasValue:
            name = safe_text(param.AsString())
            if name:
                return name
    except Exception:
        pass
    try:
        param = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if param is not None and param.HasValue:
            name = safe_text(param.AsString())
            if name:
                return name
    except Exception:
        pass
    return u""


def set_type_name(element_type, new_value):
    text = safe_text(new_value).strip()
    if not text:
        raise ValueError(u"El nombre de tipo no puede quedar vacío.")
    element_type.Name = text


def read_string_parameter(element_type, builtin):
    try:
        param = element_type.get_Parameter(builtin)
    except Exception:
        param = None
    if param is None:
        return u""
    try:
        if param.HasValue:
            value = param.AsString()
            if value is None:
                value = param.AsValueString()
            return safe_text(value)
    except Exception:
        pass
    return u""


def write_string_parameter(element_type, builtin, new_value):
    text = safe_text(new_value)
    try:
        param = element_type.get_Parameter(builtin)
    except Exception:
        param = None
    if param is None:
        raise ValueError(u"El parámetro seleccionado no está disponible en este tipo.")
    if param.IsReadOnly:
        raise ValueError(u"El parámetro seleccionado es de solo lectura.")
    try:
        result = param.Set(text)
    except Exception as err:
        raise err
    if result is False:
        raise ValueError(u"No se pudo asignar el valor indicado.")


def normalize(text):
    return safe_text(text).strip().lower()


def family_key_for(element_type, family_name):
    category = getattr(element_type, "Category", None)
    category_id = -1
    if category is not None:
        category_id = getattr(getattr(category, "Id", None), "IntegerValue", None)
        if category_id is None:
            category_id = -1
    return (category_id, normalize(family_name))


class ParameterOption(object):
    """Configura cómo leer y escribir cada parámetro soportado."""

    def __init__(self, key, label, getter=None, setter=None, builtin=None, enforce_unique=False):
        self.key = key
        self.label = label
        self.enforce_unique = enforce_unique
        self._builtin = builtin
        if builtin is not None:
            self._getter = getter or (lambda element_type: read_string_parameter(element_type, builtin))
            self._setter = setter or (lambda element_type, value: write_string_parameter(element_type, builtin, value))
        else:
            self._getter = getter
            self._setter = setter

    def get_value(self, element_type):
        if self._getter is None:
            return u""
        return safe_text(self._getter(element_type))

    def set_value(self, element_type, value):
        if self._setter is None:
            raise ValueError(u"No hay acción definida para modificar este parámetro.")
        self._setter(element_type, value)

    @property
    def name(self):
        return self.label


PARAMETER_OPTIONS = [
    ParameterOption(
        key="type_name",
        label=u"Type Name (Nombre de Tipo)",
        getter=get_type_name,
        setter=set_type_name,
        enforce_unique=True,
    ),
    ParameterOption(
        key="type_mark",
        label=u"Type Mark (Marca de Tipo)",
        builtin=BuiltInParameter.ALL_MODEL_TYPE_MARK,
    ),
    ParameterOption(
        key="manufacturer",
        label=u"Manufacturer (Fabricante)",
        builtin=BuiltInParameter.ALL_MODEL_MANUFACTURER,
    ),
    ParameterOption(
        key="model",
        label=u"Model (Modelo)",
        builtin=BuiltInParameter.ALL_MODEL_MODEL,
    ),
]


class CategoryOption(object):
    """Elemento visualizable en SelectFromList."""

    def __init__(self, category, types):
        self.category = category
        self.types = list(types)
        self.label = u"{0} ({1} tipos)".format(get_category_name(category), len(self.types))

    @property
    def name(self):
        return self.label


def collect_types_by_category(doc):
    grouped = {}
    collector = FilteredElementCollector(doc).WhereElementIsElementType()
    for element_type in collector:
        category = getattr(element_type, "Category", None)
        if category is None:
            continue
        try:
            cat_id = category.Id.IntegerValue
        except Exception:
            continue
        grouped.setdefault(cat_id, {"category": category, "types": []})
        grouped[cat_id]["types"].append(element_type)
    return grouped


def build_proposals(types, option, search_text, replace_text, ignore_case):
    proposals = []
    if ignore_case:
        pattern = re.compile(re.escape(search_text), re.IGNORECASE)
    for element_type in types:
        current_value = option.get_value(element_type)
        if current_value is None:
            current_value = u""
        if not safe_text(current_value):
            continue
        if ignore_case:
            if not pattern.search(current_value):
                continue
            new_value = pattern.sub(replace_text, current_value)
        else:
            if search_text not in current_value:
                continue
            new_value = current_value.replace(search_text, replace_text)
        if current_value == new_value:
            continue
        if not safe_text(new_value).strip():
            continue
        proposals.append(
            {
                "type": element_type,
                "old_value": safe_text(current_value),
                "new_value": safe_text(new_value).strip(),
                "family_name": get_family_name(element_type),
            }
        )
    return proposals


def separate_conflicts(types, proposals, option):
    if not proposals:
        return [], []
    if not option.enforce_unique:
        return proposals, []
    existing = {}
    for element_type in types:
        family_name = get_family_name(element_type)
        key = family_key_for(element_type, family_name)
        existing.setdefault(key, set()).add(normalize(option.get_value(element_type)))
    for proposal in proposals:
        key = family_key_for(proposal["type"], proposal["family_name"])
        existing.setdefault(key, set()).discard(normalize(proposal["old_value"]))
    accepted = []
    conflicts = []
    for proposal in proposals:
        key = family_key_for(proposal["type"], proposal["family_name"])
        names = existing.setdefault(key, set())
        new_norm = normalize(proposal["new_value"])
        if not new_norm or new_norm in names:
            conflicts.append(proposal)
            continue
        names.add(new_norm)
        accepted.append(proposal)
    return accepted, conflicts


def format_preview_lines(accepted, conflicts, option):
    lines = []
    if accepted:
        lines.append(u"Cambios propuestos en {0} ({1}):".format(option.label, len(accepted)))
        for idx, proposal in enumerate(accepted):
            if idx >= MAX_PREVIEW_LINES:
                lines.append(u"... ({0} cambios adicionales no listados)".format(len(accepted) - MAX_PREVIEW_LINES))
                break
            family = proposal["family_name"] or u"(Sin familia)"
            lines.append(u"- {0} :: {1} -> {2}".format(family, proposal["old_value"], proposal["new_value"]))
    if conflicts:
        if lines:
            lines.append(u"")
        lines.append(u"Omitidos por duplicar nombres existentes en {0} ({1}):".format(option.label, len(conflicts)))
        for idx, proposal in enumerate(conflicts):
            if idx >= MAX_PREVIEW_LINES:
                lines.append(u"... ({0} conflictos adicionales no listados)".format(len(conflicts) - MAX_PREVIEW_LINES))
                break
            family = proposal["family_name"] or u"(Sin familia)"
            lines.append(u"- {0} :: {1} -> {2}".format(family, proposal["old_value"], proposal["new_value"]))
    if not lines:
        lines.append(u"No se encontraron coincidencias para los criterios proporcionados.")
    return lines


def show_summary(title, lines):
    message = u"\n".join(lines)
    return forms.alert(message, title=title, yes=True, no=True, warn_icon=False)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert(u"No hay documento activo.", title=TITLE)
        return

    grouped = collect_types_by_category(doc)
    if not grouped:
        forms.alert(u"No se encontraron tipos en el documento.", title=TITLE)
        return

    options = [
        CategoryOption(entry["category"], sorted(entry["types"], key=lambda t: (normalize(get_family_name(t)), normalize(get_type_name(t)))))
        for entry in grouped.values()
        if entry["types"]
    ]
    if not options:
        forms.alert(u"No hay categorías con tipos para renombrar.", title=TITLE)
        return

    options.sort(key=lambda opt: normalize(opt.label))
    picked = forms.SelectFromList.show(
        options,
        title=TITLE,
        multiselect=False,
        button_name="Continuar",
        name_attr="name",
    )
    if not picked:
        forms.alert(u"Operación cancelada. No se seleccionó categoría.", title=TITLE)
        return
    if isinstance(picked, list):
        picked = picked[0]

    parameter_option = forms.SelectFromList.show(
        PARAMETER_OPTIONS,
        title=TITLE + u" - Parámetro",
        multiselect=False,
        button_name="Continuar",
        name_attr="name",
    )
    if not parameter_option:
        forms.alert(u"Operación cancelada. No se seleccionó parámetro.", title=TITLE)
        return
    if isinstance(parameter_option, list):
        parameter_option = parameter_option[0]

    search_text = forms.ask_for_string(
        default="",
        prompt=u"Ingrese el texto a buscar en {0}:".format(parameter_option.label),
        title=TITLE,
    )
    if search_text is None or search_text == "":
        forms.alert(u"Operación cancelada. No se proporcionó texto de búsqueda.", title=TITLE)
        return
    search_text = safe_text(search_text)

    replace_text = forms.ask_for_string(
        default="",
        prompt=u"Ingrese el texto de reemplazo para {0} (puede quedar vacío):".format(parameter_option.label),
        title=TITLE,
    )
    if replace_text is None:
        forms.alert(u"Operación cancelada. No se proporcionó texto de reemplazo.", title=TITLE)
        return
    replace_text = safe_text(replace_text)

    ignore_case_response = forms.alert(
        u"¿Desea ignorar mayúsculas/minúsculas durante la búsqueda?",
        title=TITLE,
        yes=True,
        no=True,
        warn_icon=False,
    )
    ignore_case = False
    if ignore_case_response:
        response_text = safe_text(ignore_case_response).lower()
        ignore_case = response_text in ("yes", "si", u"sí", "true", "ok", "1")

    proposals = build_proposals(picked.types, parameter_option, search_text, replace_text, ignore_case)
    if not proposals:
        forms.alert(u"No se encontraron tipos que coincidan con el texto indicado.", title=TITLE)
        return

    accepted, conflicts = separate_conflicts(picked.types, proposals, parameter_option)
    lines = format_preview_lines(accepted, conflicts, parameter_option)
    preview_response = show_summary(
        TITLE + u" - Confirmación",
        lines + [u"", u"¿Aplicar cambios en {0}?".format(parameter_option.label)],
    )
    if not preview_response:
        forms.alert(u"Operación cancelada por el usuario.", title=TITLE)
        return

    if not accepted:
        forms.alert(u"No hay cambios seguros para aplicar (todos generan conflicto).", title=TITLE)
        return

    transaction = Transaction(
        doc,
        u"Renombrar {0} ({1})".format(parameter_option.label, get_category_name(picked.category)),
    )
    renamed = []
    failures = []
    try:
        transaction.Start()
    except Exception as start_err:
        forms.alert(u"No se pudo iniciar la transacción:\n{0}".format(safe_text(start_err)), title=TITLE)
        return

    for proposal in accepted:
        try:
            parameter_option.set_value(proposal["type"], proposal["new_value"])
            renamed.append(proposal)
        except Exception as err:
            failures.append((proposal, safe_text(err)))

    try:
        if renamed:
            transaction.Commit()
        else:
            transaction.RollBack()
    except Exception as end_err:
        forms.alert(u"Error al finalizar la transacción:\n{0}".format(safe_text(end_err)), title=TITLE)
        return

    summary_lines = []
    if renamed:
        summary_lines.append(u"{0} actualizados: {1}".format(parameter_option.label, len(renamed)))
        for idx, proposal in enumerate(renamed):
            if idx >= MAX_PREVIEW_LINES:
                summary_lines.append(u"... ({0} cambios adicionales no listados)".format(len(renamed) - MAX_PREVIEW_LINES))
                break
            family = proposal["family_name"] or u"(Sin familia)"
            summary_lines.append(u"- {0} :: {1} -> {2}".format(family, proposal["old_value"], proposal["new_value"]))
    if failures:
        if summary_lines:
            summary_lines.append(u"")
        summary_lines.append(u"No se pudieron actualizar {0} elementos:".format(len(failures)))
        for idx, (proposal, reason) in enumerate(failures):
            if idx >= MAX_PREVIEW_LINES:
                summary_lines.append(u"... ({0} errores adicionales no listados)".format(len(failures) - MAX_PREVIEW_LINES))
                break
            family = proposal["family_name"] or u"(Sin familia)"
            summary_lines.append(u"- {0} :: {1} -> {2} | {3}".format(family, proposal["old_value"], proposal["new_value"], reason))
    if conflicts:
        if summary_lines:
            summary_lines.append(u"")
        summary_lines.append(u"Cambios omitidos por conflicto en {0}: {1}".format(parameter_option.label, len(conflicts)))

    forms.alert(u"\n".join(summary_lines), title=TITLE, warn_icon=False)


if __name__ == "__main__":
    main()
