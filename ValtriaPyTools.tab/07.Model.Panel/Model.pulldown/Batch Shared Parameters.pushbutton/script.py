# -*- coding: utf-8 -*-
"""Agrega varios parametros compartidos de una sola vez a categorias seleccionadas."""

import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", ".."))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (  # type: ignore
    BuiltInParameterGroup,
    CategorySet,
    InstanceBinding,
    LabelUtils,
    Transaction,
    TypeBinding,
)
from pyrevit import forms

from valtria_lib import get_app, get_doc, log_exception, log_to_file


TITLE = "Batch Shared Parameters"
LOG_TOOL = "batch_shared_parameters"

COMMON_GROUPS = [
    BuiltInParameterGroup.PG_DATA,
    BuiltInParameterGroup.PG_TEXT,
    BuiltInParameterGroup.PG_CONSTRAINTS,
    BuiltInParameterGroup.PG_GEOMETRY,
    BuiltInParameterGroup.PG_GENERAL,
    BuiltInParameterGroup.PG_IFC,
    BuiltInParameterGroup.PG_IDENTITY_DATA,
    BuiltInParameterGroup.PG_MATERIALS,
]


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


def log_line(message):
    log_to_file(LOG_TOOL, message)


def iterate(net_collection):
    """Itera sobre colecciones .NET de forma segura."""
    if net_collection is None:
        return
    try:
        for item in net_collection:
            yield item
        return
    except Exception:
        pass
    try:
        iterator = net_collection.GetEnumerator()
    except Exception:
        iterator = None
    if iterator is None:
        return
    while iterator.MoveNext():
        yield iterator.Current


class ParameterOption(object):
    """Opcion de parametro compartido para SelectFromList."""

    def __init__(self, definition, group_name, type_label):
        self.value = definition
        self.group_name = group_name
        self.type_label = type_label

    @property
    def name(self):
        return u"{0}  [{1}] ({2})".format(
            safe_text(getattr(self.value, "Name", u"")),
            safe_text(self.group_name) or u"(sin grupo)",
            safe_text(self.type_label) or u"?",
        )


class CategoryOption(object):
    """Opcion de categoria para SelectFromList."""

    def __init__(self, category):
        self.value = category
        self.raw_name = safe_text(getattr(category, "Name", u""))

    @property
    def name(self):
        return self.raw_name or u"(Sin categoria)"


class GroupOption(object):
    """Opcion de BuiltInParameterGroup."""

    def __init__(self, group):
        self.value = group
        self.raw_name = describe_group(group)

    @property
    def name(self):
        return self.raw_name


def describe_group(group):
    try:
        return safe_text(LabelUtils.GetLabelFor(group))
    except Exception:
        return safe_text(group)


def describe_definition(definition):
    """Devuelve un texto corto con el tipo de dato del parametro."""
    try:
        ptype = getattr(definition, "ParameterType", None)
        if ptype is not None:
            return safe_text(ptype)
    except Exception:
        pass
    try:
        dtype = definition.GetDataType()
        if dtype is not None:
            type_id = safe_text(getattr(dtype, "TypeId", None))
            return type_id or safe_text(dtype)
    except Exception:
        pass
    return u""


def prompt_shared_param_file(app):
    current = safe_text(getattr(app, "SharedParametersFilename", u"")).strip()
    if current and os.path.exists(current):
        use_current = forms.alert(
            u"Usar el archivo de parametros compartidos actual?\n{0}".format(current),
            title=TITLE,
            yes=True,
            no=True,
            warn_icon=False,
        )
        if use_current:
            return current
    picked = forms.pick_file(
        file_ext="txt",
        restore_dir=True,
        title="Selecciona el archivo de parametros compartidos (.txt)",
    )
    if picked:
        app.SharedParametersFilename = picked
        return picked
    return None


def collect_parameter_options(definition_file):
    options = []
    for group in iterate(getattr(definition_file, "Groups", None)):
        group_name = safe_text(getattr(group, "Name", u""))
        for definition in iterate(getattr(group, "Definitions", None)):
            options.append(
                ParameterOption(
                    definition,
                    group_name,
                    describe_definition(definition),
                )
            )
    options.sort(key=lambda opt: opt.name.lower())
    return options


def collect_bindable_categories(doc):
    categories = []
    cat_collection = getattr(getattr(doc, "Settings", None), "Categories", None)
    for cat in iterate(cat_collection):
        if cat is None:
            continue
        try:
            if not getattr(cat, "AllowsBoundParameters", False):
                continue
        except Exception:
            continue
        try:
            if getattr(cat, "Parent", None) is not None:
                continue  # evitar subcategorias
        except Exception:
            pass
        categories.append(cat)
    categories.sort(key=lambda c: safe_text(getattr(c, "Name", u"")).lower())
    return categories


def count_categories(category_set):
    """Cuenta elementos en un CategorySet que no implementa __len__."""
    total = 0
    for _ in iterate(category_set):
        total += 1
    return total


def binding_covers_categories(binding, category_set):
    """Comprueba si un binding ya incluye todas las categorias seleccionadas."""
    cats = getattr(binding, "Categories", None)
    if cats is None:
        return False
    for cat in iterate(category_set):
        try:
            if not cats.Contains(cat):
                return False
        except Exception:
            return False
    return True


def find_binding_for_definition(bindings, definition):
    """Devuelve el binding existente para una definicion o None."""
    if bindings is None or definition is None:
        return None
    try:
        it = bindings.ForwardIterator()
    except Exception:
        return None
    while it.MoveNext():
        try:
            defn = it.Key
        except Exception:
            defn = None
        if defn == definition:
            try:
                return it.Current
            except Exception:
                return None
    return None


def pick_parameter_group():
    options = [GroupOption(group) for group in COMMON_GROUPS]
    picked = forms.SelectFromList.show(
        options,
        title=TITLE + " - Grupo de parametros",
        multiselect=False,
        button_name="Usar grupo",
        name_attr="name",
    )
    if picked:
        return picked.value if hasattr(picked, "value") else picked
    return BuiltInParameterGroup.PG_DATA


def format_list(items, max_items=6):
    values = [safe_text(it).strip() for it in items if safe_text(it).strip()]
    if not values:
        return u""
    if len(values) <= max_items:
        return u", ".join(values)
    return u"{0} (+{1} mas)".format(u", ".join(values[:max_items]), len(values) - max_items)


def main():
    log_line("----")
    log_line("Inicio Batch Shared Parameters")
    try:
        doc = get_doc()
        app = get_app()
    except Exception as ctx_err:
        log_exception(ctx_err)
        forms.alert("No hay documento activo.", title=TITLE, warn_icon=True)
        return

    if doc is None or app is None:
        forms.alert("No hay documento activo.", title=TITLE, warn_icon=True)
        log_line("Abortado: doc o app nulos")
        return

    shared_path = prompt_shared_param_file(app)
    if not shared_path:
        log_line("Abortado: no se selecciono archivo de parametros compartidos")
        return

    try:
        definition_file = app.OpenSharedParameterFile()
    except Exception as err:
        definition_file = None
        log_exception(err)
    if definition_file is None:
        forms.alert(
            "No se pudo abrir el archivo de parametros compartidos:\n{0}".format(shared_path),
            title=TITLE,
            warn_icon=True,
        )
        return

    param_options = collect_parameter_options(definition_file)
    if not param_options:
        forms.alert("El archivo no contiene parametros disponibles.", title=TITLE, warn_icon=True)
        log_line("Abortado: archivo sin definiciones")
        return

    selected_params = forms.SelectFromList.show(
        param_options,
        title=TITLE + " - Parametros",
        multiselect=True,
        button_name="Agregar",
        name_attr="name",
    )
    if not selected_params:
        log_line("Abortado: usuario no selecciono parametros")
        return
    if not isinstance(selected_params, (list, tuple)):
        selected_params = [selected_params]
    definitions = [getattr(opt, "value", opt) for opt in selected_params]

    categories = collect_bindable_categories(doc)
    if not categories:
        forms.alert("No hay categorias disponibles para vincular parametros.", title=TITLE, warn_icon=True)
        log_line("Abortado: sin categorias")
        return

    cat_options = [CategoryOption(cat) for cat in categories]
    selected_cats = forms.SelectFromList.show(
        cat_options,
        title=TITLE + " - Categorias",
        multiselect=True,
        button_name="Usar categorias",
        name_attr="name",
    )
    if not selected_cats:
        log_line("Abortado: usuario no selecciono categorias")
        return
    if not isinstance(selected_cats, (list, tuple)):
        selected_cats = [selected_cats]
    category_set = CategorySet()
    for opt in selected_cats:
        cat = getattr(opt, "value", opt)
        try:
            category_set.Insert(cat)
        except Exception:
            pass

    scope_action = forms.CommandSwitchWindow.show(
        ["Instancia", "Tipo"],
        message="Como quieres crear los parametros seleccionados?",
        title=TITLE,
    )
    if not scope_action:
        log_line("Abortado: usuario no eligio instancia/tipo")
        return
    is_instance = scope_action.startswith("Instancia")

    param_group = pick_parameter_group()

    summary_lines = []
    bindings = getattr(doc, "ParameterBindings", None)

    to_process = []
    already_bound = []
    for definition in definitions:
        existing_binding = find_binding_for_definition(bindings, definition)
        if existing_binding is not None:
            if is_instance and isinstance(existing_binding, InstanceBinding):
                if binding_covers_categories(existing_binding, category_set):
                    already_bound.append(safe_text(getattr(definition, "Name", u"")))
                    continue
            if (not is_instance) and isinstance(existing_binding, TypeBinding):
                if binding_covers_categories(existing_binding, category_set):
                    already_bound.append(safe_text(getattr(definition, "Name", u"")))
                    continue
        to_process.append(definition)

    summary_lines.append(u"Parametros a agregar: {0}".format(len(to_process)))
    summary_lines.append(format_list([safe_text(getattr(d, "Name", u"")) for d in to_process]))
    summary_lines.append(u"Categorias: {0}".format(count_categories(category_set)))
    summary_lines.append(
        format_list([safe_text(getattr(getattr(opt, "value", opt), "Name", u"")) for opt in selected_cats])
    )
    summary_lines.append(u"Alcance: {0}".format("Instancia" if is_instance else "Tipo"))
    summary_lines.append(u"Grupo: {0}".format(describe_group(param_group)))
    if already_bound:
        summary_lines.append(u"Ya estaban agregados y se omitiran: {0}".format(len(already_bound)))
        summary_lines.append(format_list(already_bound))
    summary_lines.append(u"\nAplicar estos parametros?")
    if not forms.alert(u"\n".join([line for line in summary_lines if line]), title=TITLE, yes=True, no=True, warn_icon=False):
        log_line("Abortado en confirmacion")
        return

    tx = Transaction(doc, TITLE)
    tx.Start()
    inserted = []
    reinserted = []
    failed = []

    binding_type = InstanceBinding if is_instance else TypeBinding

    if not to_process:
        forms.alert(
            "Todos los parametros seleccionados ya estaban agregados para las categorias y alcance indicados.",
            title=TITLE,
            warn_icon=False,
        )
        return

    for definition in to_process:
        name = safe_text(getattr(definition, "Name", u""))
        try:
            binding = binding_type(category_set)
            ok = bindings.Insert(definition, binding, param_group)
            if ok:
                inserted.append(name)
                continue
            ok = bindings.ReInsert(definition, binding, param_group)
            if ok:
                reinserted.append(name)
            else:
                failed.append((name, u"No se pudo insertar ni actualizar."))
        except Exception as err:
            failed.append((name, safe_text(err)))
            log_line("Error al procesar {0}: {1}".format(name, safe_text(err)))

    if inserted or reinserted:
        try:
            tx.Commit()
        except Exception as err:
            tx.RollBack()
            log_exception(err)
            forms.alert(
                "No se pudo completar la operacion:\n{0}".format(safe_text(err)),
                title=TITLE,
                warn_icon=True,
            )
            return
    else:
        tx.RollBack()
        forms.alert("No se realizaron cambios.", title=TITLE, warn_icon=True)
        return

    summary = []
    summary.append(u"Parametros nuevos: {0}".format(len(inserted)))
    if inserted:
        summary.append(format_list(inserted))
    if reinserted:
        summary.append(u"")
        summary.append(u"Actualizados/ReInsertados: {0}".format(len(reinserted)))
        summary.append(format_list(reinserted))
    if failed:
        summary.append(u"")
        summary.append(u"Errores: {0}".format(len(failed)))
        summary.append(format_list([u"{0}: {1}".format(n, e) for n, e in failed]))
    forms.alert(u"\n".join([line for line in summary if line is not None]), title=TITLE, warn_icon=bool(failed))


if __name__ == "__main__":
    try:
        main()
    except Exception as main_error:
        try:
            tb = traceback.format_exc()
            log_line("Error inesperado: {0}\n{1}".format(safe_text(main_error), safe_text(tb)))
        except Exception:
            pass
        log_exception(main_error)

