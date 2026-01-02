# -*- coding: utf-8 -*-
"""
Super Renamer: ajustes masivos al parametro Mark (u otro) sobre la seleccion actual.

Flujo:
- Toma la seleccion activa o permite elegir elementos si no hay seleccion.
- Selecciona el parametro a modificar desde la seleccion disponible.
- Permite elegir operaciones: copiar valor desde otro parametro, prefijo, buscar/reemplazar, sufijo.
- Muestra los cambios detectados antes de aplicar.
- Ejecuta la transaccion con los cambios confirmados.

Compatibilidad: pyRevit (IronPython), Revit 2019+
"""
import re

from Autodesk.Revit.DB import (
    BuiltInParameter,
    StorageType,
    Transaction,
)
from Autodesk.Revit.UI import Selection
from pyrevit import forms, revit


TITLE = "Super Renamer"
MAX_PREVIEW_LINES = 40

try:
    unicode
except NameError:
    unicode = str  # type: ignore


class OperationOption(object):
    """Opcion de operacion a aplicar."""

    def __init__(self, key, label):
        self.key = key
        self.label = label

    @property
    def name(self):
        return self.label


OPERATIONS = [
    OperationOption("copy", u"Copiar valor desde otro parametro"),
    OperationOption("prefix", u"Añadir prefijo"),
    OperationOption("search", u"Buscar y reemplazar"),
    OperationOption("suffix", u"Añadir sufijo"),
]


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


def get_parameter(element, param_name):
    if not element:
        return None
    param = element.LookupParameter(param_name)
    if param is None and param_name.strip().lower() == "mark":
        try:
            param = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        except Exception:
            param = None
    return param


def read_parameter(param):
    if param is None:
        return u""
    try:
        text = param.AsString()
        if text is not None:
            return safe_text(text)
    except Exception:
        pass
    try:
        text = param.AsValueString()
        if text is not None:
            return safe_text(text)
    except Exception:
        pass
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    try:
        if storage == StorageType.Double:
            return safe_text(param.AsDouble())
        if storage == StorageType.Integer:
            return safe_text(param.AsInteger())
        if storage == StorageType.ElementId:
            element_id = param.AsElementId()
            return safe_text(getattr(element_id, "IntegerValue", element_id))
    except Exception:
        pass
    return u""


def assign_parameter(param, new_value):
    if param is None:
        raise ValueError("Parametro no disponible.")
    text = safe_text(new_value)
    try:
        storage = param.StorageType
    except Exception:
        storage = None
    try:
        if storage == StorageType.String or storage == StorageType.None:
            result = param.Set(text)
        else:
            if hasattr(param, "SetValueString"):
                result = param.SetValueString(text)
            else:
                result = param.Set(text)
    except Exception as err:
        raise err
    if result is False:
        raise ValueError("No se pudo asignar el valor indicado.")


def get_available_parameter_names(elements):
    name_map = {}
    for element in elements:
        if element is None:
            continue
        try:
            params = element.Parameters
        except Exception:
            params = None
        if not params:
            continue
        try:
            for param in params:
                if param is None:
                    continue
                try:
                    definition = param.Definition
                except Exception:
                    definition = None
                name = safe_text(getattr(definition, "Name", ""))
                if not name:
                    continue
                key = name.lower()
                if key not in name_map:
                    name_map[key] = name
        except Exception:
            try:
                iterator = params.GetEnumerator()
            except Exception:
                iterator = None
            if iterator is None:
                continue
            try:
                while iterator.MoveNext():
                    param = iterator.Current
                    if param is None:
                        continue
                    try:
                        definition = param.Definition
                    except Exception:
                        definition = None
                    name = safe_text(getattr(definition, "Name", ""))
                    if not name:
                        continue
                    key = name.lower()
                    if key not in name_map:
                        name_map[key] = name
            except Exception:
                continue
    return sorted(name_map.values(), key=lambda n: n.lower())


def select_parameter_from_list(param_names, title):
    if not param_names:
        return None
    return forms.SelectFromList.show(
        param_names,
        title=title,
        multiselect=False,
        button_name="Seleccionar",
    )


def apply_operations(original_value, operations_config):
    value = safe_text(original_value)
    for config in operations_config:
        if config["key"] == "prefix":
            value = safe_text(config["value"]) + value
        elif config["key"] == "suffix":
            value = value + safe_text(config["value"])
        elif config["key"] == "search":
            search_text = config["search"]
            replace_text = config["replace"]
            if config.get("ignore_case"):
                pattern = re.compile(re.escape(search_text), re.IGNORECASE)
                value = pattern.sub(replace_text, value)
            else:
                value = value.replace(search_text, replace_text)
    return safe_text(value)


def element_label(element):
    if element is None:
        return u"(Elemento nulo)"
    name = safe_text(getattr(element, "Name", u""))
    if not name:
        name = element.__class__.__name__
    try:
        elem_id = element.Id.IntegerValue
    except Exception:
        elem_id = "?"
    return u"{0} [Id:{1}]".format(name, elem_id)


def format_preview(changes, missing, readonly, unchanged, missing_source, param_name, source_param_name):
    lines = []
    lines.append(u"Parámetro: {0}".format(param_name))
    if source_param_name:
        lines.append(u"Parámetro origen: {0}".format(source_param_name))
    lines.append(u"Elementos con cambios: {0}".format(len(changes)))
    if missing:
        lines.append(u"Sin parámetro ({0}): {1}".format(len(missing), ", ".join(str(e.Id.IntegerValue) for e in missing)))
    if missing_source:
        lines.append(u"Sin parámetro origen ({0}): {1}".format(len(missing_source), ", ".join(str(e.Id.IntegerValue) for e in missing_source)))
    if readonly:
        lines.append(u"Solo lectura ({0}): {1}".format(len(readonly), ", ".join(str(e.Id.IntegerValue) for e in readonly)))
    if unchanged:
        lines.append(u"Sin cambios necesarios: {0}".format(len(unchanged)))
    if changes:
        lines.append(u"")
        for idx, info in enumerate(changes):
            if idx >= MAX_PREVIEW_LINES:
                lines.append(u"... ({0} cambios adicionales no listados)".format(len(changes) - MAX_PREVIEW_LINES))
                break
            lines.append(
                u"- {0}: '{1}' -> '{2}'".format(
                    element_label(info["element"]),
                    info["old_value"],
                    info["new_value"],
                )
            )
    return lines


def ensure_elements(doc, uidoc):
    if doc is None or uidoc is None:
        forms.alert("No hay documento activo.", title=TITLE)
        return []
    selection_ids = list(uidoc.Selection.GetElementIds())
    elements = [doc.GetElement(eid) for eid in selection_ids if doc.GetElement(eid) is not None]
    if elements:
        return elements
    try:
        picked = uidoc.Selection.PickObjects(Selection.ObjectType.Element, "Selecciona elementos para renombrar")
    except Selection.OperationCanceledException:
        forms.alert("Operación cancelada. No hay elementos seleccionados.", title=TITLE)
        return []
    except Exception as pick_err:
        forms.alert("No se pudieron seleccionar elementos:\n{0}".format(pick_err), title=TITLE)
        return []
    elements = [doc.GetElement(ref.ElementId) for ref in picked if doc.GetElement(ref.ElementId) is not None]
    if not elements:
        forms.alert("No se seleccionaron elementos válidos.", title=TITLE)
    return elements


def configure_operations(param_names):
    picked = forms.SelectFromList.show(
        OPERATIONS,
        title=TITLE + " - Operaciones",
        multiselect=True,
        button_name="Configurar",
        name_attr="name",
    )
    if not picked:
        forms.alert("Operación cancelada. No se eligieron acciones.", title=TITLE)
        return []
    if not isinstance(picked, list):
        picked = [picked]
    selected_keys = set(opt.key for opt in picked)
    ordered = [opt for opt in OPERATIONS if opt.key in selected_keys]
    config = []
    for opt in ordered:
        if opt.key == "copy":
            if not param_names:
                forms.alert("No hay parametros disponibles para elegir el origen.", title=TITLE)
                return []
            source_param = select_parameter_from_list(
                param_names,
                TITLE + " - Parametro origen",
            )
            if not source_param:
                forms.alert("Operacion cancelada. No se eligio parametro origen.", title=TITLE)
                return []
            config.append({"key": "copy", "source_param": source_param})
        elif opt.key == "prefix":
            prefix = forms.ask_for_string(
                default="",
                prompt="Prefijo a añadir (se antepone al valor actual):",
                title=TITLE,
            )
            if prefix is None:
                return []
            if prefix == "":
                continue
            config.append({"key": "prefix", "value": safe_text(prefix)})
        elif opt.key == "suffix":
            suffix = forms.ask_for_string(
                default="",
                prompt="Sufijo a añadir (se agrega al final del valor actual):",
                title=TITLE,
            )
            if suffix is None:
                return []
            if suffix == "":
                continue
            config.append({"key": "suffix", "value": safe_text(suffix)})
        elif opt.key == "search":
            search_text = forms.ask_for_string(
                default="",
                prompt="Texto a buscar:",
                title=TITLE,
            )
            if search_text is None or search_text == "":
                forms.alert("La operación de búsqueda requiere un texto a buscar. Se omitirá.", title=TITLE)
                continue
            replace_text = forms.ask_for_string(
                default="",
                prompt="Texto de reemplazo (puede quedar vacío):",
                title=TITLE,
            )
            if replace_text is None:
                return []
            ignore_case = forms.alert(
                "¿Ignorar mayúsculas/minúsculas en la búsqueda?",
                title=TITLE,
                yes=True,
                no=True,
                warn_icon=False,
            )
            config.append(
                {
                    "key": "search",
                    "search": safe_text(search_text),
                    "replace": safe_text(replace_text),
                    "ignore_case": bool(ignore_case),
                }
            )
    return config


def get_copy_source_param(operations_config):
    for config in operations_config:
        if config.get("key") == "copy":
            return config.get("source_param")
    return None


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    elements = ensure_elements(doc, uidoc)
    if not elements:
        return

    param_names = get_available_parameter_names(elements)
    if not param_names:
        forms.alert("No se encontraron parametros en la seleccion.", title=TITLE)
        return
    param_name = select_parameter_from_list(
        param_names,
        TITLE + " - Parametro a modificar",
    )
    if not param_name:
        forms.alert("Operacion cancelada. No se eligio parametro.", title=TITLE)
        return
    param_name = safe_text(param_name).strip()

    operations_config = configure_operations(param_names)
    if not operations_config:
        forms.alert("No se configuraron operaciones válidas.", title=TITLE)
        return

    changes = []
    missing = []
    missing_source = []
    readonly = []
    unchanged = []
    source_param_name = get_copy_source_param(operations_config)

    processed_ids = set()
    for element in elements:
        if element is None:
            continue
        elem_id = getattr(getattr(element, "Id", None), "IntegerValue", None)
        if elem_id is not None:
            if elem_id in processed_ids:
                continue
            processed_ids.add(elem_id)
        param = get_parameter(element, param_name)
        if param is None:
            missing.append(element)
            continue
        if param.IsReadOnly:
            readonly.append(element)
            continue
        current_value = read_parameter(param)
        base_value = current_value
        if source_param_name:
            source_param = get_parameter(element, source_param_name)
            if source_param is None:
                missing_source.append(element)
                continue
            base_value = read_parameter(source_param)
        new_value = apply_operations(base_value, operations_config)
        if current_value == new_value:
            unchanged.append(element)
            continue
        changes.append(
            {
                "element": element,
                "param": param,
                "old_value": current_value,
                "new_value": new_value,
            }
        )

    if not changes:
        message = ["No hay cambios para aplicar en el parámetro {0}.".format(param_name)]
        if missing:
            message.append("Elementos sin parámetro: {0}".format(len(missing)))
        if missing_source:
            message.append("Elementos sin parámetro origen: {0}".format(len(missing_source)))
        if readonly:
            message.append("Elementos con parámetro de solo lectura: {0}".format(len(readonly)))
        if unchanged:
            message.append("Elementos sin modificaciones necesarias: {0}".format(len(unchanged)))
        forms.alert("\n".join(message), title=TITLE)
        return

    preview_lines = format_preview(
        changes, missing, readonly, unchanged, missing_source, param_name, source_param_name
    )
    preview_lines.append(u"")
    preview_lines.append(u"¿Aplicar estos cambios?")
    confirm = forms.alert("\n".join(preview_lines), title=TITLE, yes=True, no=True, warn_icon=False)
    if not confirm:
        forms.alert("Operación cancelada por el usuario.", title=TITLE)
        return

    transaction = Transaction(doc, "Super Renamer ({0})".format(param_name))
    applied = []
    failures = []

    try:
        transaction.Start()
    except Exception as start_err:
        forms.alert("No se pudo iniciar la transacción:\n{0}".format(safe_text(start_err)), title=TITLE)
        return

    for info in changes:
        try:
            assign_parameter(info["param"], info["new_value"])
            applied.append(info)
        except Exception as err:
            failures.append((info, safe_text(err)))

    try:
        if applied:
            transaction.Commit()
        else:
            transaction.RollBack()
    except Exception as end_err:
        forms.alert("Error al finalizar la transacción:\n{0}".format(safe_text(end_err)), title=TITLE)
        return

    summary = []
    summary.append("Parámetro: {0}".format(param_name))
    if source_param_name:
        summary.append("Parámetro origen: {0}".format(source_param_name))
    summary.append("Elementos actualizados: {0}".format(len(applied)))
    for idx, info in enumerate(applied):
        if idx >= MAX_PREVIEW_LINES:
            summary.append("... ({0} cambios adicionales no listados)".format(len(applied) - MAX_PREVIEW_LINES))
            break
        summary.append(
            "- {0}: '{1}' -> '{2}'".format(
                element_label(info["element"]),
                info["old_value"],
                info["new_value"],
            )
        )
    if failures:
        summary.append("")
        summary.append("No se pudieron actualizar {0} elementos:".format(len(failures)))
        for idx, (info, reason) in enumerate(failures):
            if idx >= MAX_PREVIEW_LINES:
                summary.append("... ({0} errores adicionales no listados)".format(len(failures) - MAX_PREVIEW_LINES))
                break
            summary.append(
                "- {0}: '{1}' -> '{2}' | {3}".format(
                    element_label(info["element"]),
                    info["old_value"],
                    info["new_value"],
                    reason,
                )
            )
    if missing:
        summary.append("")
        summary.append("Elementos sin parámetro {0}: {1}".format(param_name, len(missing)))
    if missing_source:
        summary.append("Elementos sin parámetro origen: {0}".format(len(missing_source)))
    if readonly:
        summary.append("Elementos con parámetro de solo lectura: {0}".format(len(readonly)))
    if unchanged:
        summary.append("Elementos sin cambios necesarios: {0}".format(len(unchanged)))

    forms.alert("\n".join(summary), title=TITLE, warn_icon=False)


if __name__ == "__main__":
    main()
