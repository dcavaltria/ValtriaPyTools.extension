# -*- coding: utf-8 -*-
"""Export a CSV takeoff of duct elements filtered by system."""

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXT_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", ".."))
LIB_DIR = os.path.join(EXT_DIR, "lib")
for _path in (EXT_DIR, LIB_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
)
from pyrevit import forms

from valtria_lib import (
    get_doc,
    ask_save_csv,
    write_csv,
    feet_to_m,
    param_str,
    system_name_of,
    log_exception,
)


CATEGORIES = (
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctCurves,
)
HEADERS = [
    'Category',
    'Element Id',
    'System',
    'Type Name',
    'Size',
    'Length (m)',
    'Level',
    'Workset',
    'Comments',
]


class SystemOption(object):
    def __init__(self, system_name):
        self.system_name = system_name or ''

    @property
    def label(self):
        return self.system_name or '<Sin sistema>'


def collect_elements(doc):
    elements = []
    for category in CATEGORIES:
        collector = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType()
        for element in collector:
            elements.append(element)
    return elements


def sort_key(name):
    name = name or ''
    return (name == '', name.lower())


def prompt_systems(elements):
    found_names = set()
    for element in elements:
        system_name = system_name_of(element) or ''
        found_names.add(system_name)
    if not found_names:
        return []
    options = [SystemOption(name) for name in sorted(found_names, key=sort_key)]
    selection = forms.SelectFromList.show(
        options,
        title='Seleccionar sistemas HVAC',
        name_attr='label',
        multiselect=True,
        button_name='Filtrar',
    )
    if selection is None:
        return None
    if isinstance(selection, (list, tuple)):
        return [opt.system_name for opt in selection]
    return [selection.system_name]


def get_type_name(doc, element):
    try:
        type_id = element.GetTypeId()
    except Exception:
        type_id = None
    if isinstance(type_id, ElementId) and type_id.IntegerValue > 0:
        element_type = doc.GetElement(type_id)
        if element_type is not None:
            return getattr(element_type, 'Name', '') or ''
    return ''


def get_length_in_m(element):
    try:
        param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
    except Exception:
        param = None
    if param is None:
        return ''
    try:
        value = param.AsDouble()
    except Exception:
        value = 0.0
    if not value:
        return ''
    return '{0:.2f}'.format(feet_to_m(value))


def get_level_name(doc, element):
    try:
        level_id = element.LevelId
    except Exception:
        level_id = None
    if isinstance(level_id, ElementId) and level_id.IntegerValue > 0:
        level = doc.GetElement(level_id)
        if level is not None:
            return getattr(level, 'Name', '') or ''
    return ''


def element_row(doc, element, system_name):
    category_name = getattr(getattr(element, 'Category', None), 'Name', '') or ''
    element_id = element.Id.IntegerValue if element and element.Id else ''
    type_name = get_type_name(doc, element)
    size_value = ''
    try:
        size_param = element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)
    except Exception:
        size_param = None
    if size_param is not None:
        size_value = param_str(size_param)
    length_value = get_length_in_m(element)
    level_name = get_level_name(doc, element)
    try:
        workset_param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    except Exception:
        workset_param = None
    workset_name = param_str(workset_param) if workset_param is not None else ''
    try:
        comments_param = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    except Exception:
        comments_param = None
    comments = param_str(comments_param) if comments_param is not None else ''
    return [
        category_name,
        element_id,
        system_name,
        type_name,
        size_value,
        length_value,
        level_name,
        workset_name,
        comments,
    ]


def main():
    doc = get_doc()
    elements = collect_elements(doc)
    if not elements:
        forms.alert('No se encontraron elementos de conductos.', title='Duct Takeoff')
        return

    selected_systems = prompt_systems(elements)
    if selected_systems is None:
        return
    selected_set = set(selected_systems) if selected_systems else None

    filtered = []
    for element in elements:
        system_name = system_name_of(element) or ''
        if selected_set is None or not selected_set:
            filtered.append((element, system_name))
        elif system_name in selected_set:
            filtered.append((element, system_name))
    if not filtered:
        forms.alert('No hay elementos en los sistemas seleccionados.', title='Duct Takeoff')
        return

    filepath = ask_save_csv('Duct_Takeoff.csv')
    if not filepath:
        return

    rows = []
    for element, system_name in filtered:
        rows.append(element_row(doc, element, system_name))
    write_csv(filepath, HEADERS, rows)
    forms.alert('CSV generado: {0}'.format(filepath), title='Duct Takeoff')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)



