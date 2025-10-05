# -*- coding: utf-8 -*-
"""Export a takeoff of ducts, fittings, and accessories to CSV."""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
)
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    ask_save_csv,
    write_csv,
    feet_to_m,
    param_str,
    system_name_of,
    log_exception,
)


CATEGORIES = [
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
]


def collect_elements(doc):
    elements = []
    for category in CATEGORIES:
        collector = FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType()
        for element in collector:
            elements.append(element)
    return elements


def element_data(doc, element):
    category_name = ''
    if element.Category:
        category_name = element.Category.Name
    element_id = element.Id.IntegerValue
    system_name = system_name_of(element)
    type_name = ''
    type_id = element.GetTypeId()
    if type_id and type_id.IntegerValue > 0:
        element_type = doc.GetElement(type_id)
        if element_type is not None:
            type_name = element_type.Name
    size = param_str(element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE))
    length = ''
    param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
    if param is not None:
        try:
            length_value = param.AsDouble()
            if length_value:
                length = feet_to_m(length_value)
        except Exception:
            length = ''
    level_name = ''
    try:
        level_id = element.LevelId
        if level_id and level_id.IntegerValue > 0:
            level = doc.GetElement(level_id)
            if level is not None:
                level_name = level.Name
    except Exception:
        pass
    workset = param_str(element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM))
    comments = param_str(element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS))
    return [
        category_name,
        element_id,
        system_name,
        type_name,
        size,
        length,
        level_name,
        workset,
        comments,
    ]


def main():
    doc = get_doc()
    elements = collect_elements(doc)
    if not elements:
        forms.alert('No se encontraron elementos de conductos ni accesorios.', title='Duct Takeoff')
        return
    filepath = ask_save_csv('Ducts_Takeoff.csv')
    if not filepath:
        return
    headers = ['Category', 'Id', 'System', 'Type Name', 'Size', 'Length (m)', 'Level', 'Workset', 'Comments']
    rows = []
    for element in elements:
        rows.append(element_data(doc, element))
    write_csv(filepath, headers, rows)
    forms.alert('CSV generado: {0}'.format(filepath), title='Duct Takeoff')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

