# -*- coding: utf-8 -*-
"""Export a CSV summary of HVAC duct systems."""

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXT_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", ".."))
LIB_DIR = os.path.join(EXT_DIR, "_lib")
for _path in (EXT_DIR, LIB_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    BuiltInCategory,
    MEPSystem,
    WorksetTable,
)
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    ask_save_csv,
    write_csv,
    feet_to_m,
    log_exception,
    param_str,
)


DUCT_SYSTEM_CATEGORY = int(BuiltInCategory.OST_DuctSystem)


FLOW_PARAM_NAMES = (
    'RBS_SYSTEM_FLOW_PARAM',
    'RBS_SYSTEM_AIRFLOW_PARAM',
    'RBS_DUCT_FLOW_PARAM',
    'RBS_DUCT_VOLUME_FLOW_PARAM',
)
STATIC_PRESSURE_PARAM_NAMES = (
    'RBS_SYSTEM_STATIC_PRESSURE_PARAM',
    'RBS_SYSTEM_INLET_STATIC_PRESSURE_PARAM',
    'RBS_SYSTEM_OUTLET_STATIC_PRESSURE_PARAM',
    'RBS_SYS_STATIC_PRES',
)
FLOW_FALLBACK_NAMES = (
    'Flow',
    'Caudal',
)
STATIC_PRESSURE_FALLBACK_NAMES = (
    'Static Pressure',
    'Presion estatica',
)


def gather_systems(doc):
    systems = []
    collector = FilteredElementCollector(doc).OfClass(MEPSystem)
    for system in collector:
        if system.Category and system.Category.Id.IntegerValue == DUCT_SYSTEM_CATEGORY:
            systems.append(system)
    return systems


def select_systems(systems):
    if not systems:
        return None
    sorted_systems = sorted(systems, key=lambda sys: (sys.Name or "").lower())
    selection = forms.SelectFromList.show(
        sorted_systems,
        title='Seleccionar sistemas HVAC',
        name_attr='Name',
        multiselect=True,
        button_name='Exportar',
    )
    if selection is None:
        return None
    if isinstance(selection, (list, tuple)):
        return list(selection)
    return [selection]


def resolve_workset_name(doc, workset_id):
    if workset_id is None:
        return ''
    try:
        if hasattr(workset_id, 'IntegerValue') and workset_id.IntegerValue > 0:
            workset = WorksetTable.GetWorkset(doc, workset_id)
            if workset is not None:
                name = getattr(workset, 'Name', '')
                if name:
                    return name
    except Exception:
        pass
    return ''


def resolve_system_type(doc, system):
    type_name = ''
    type_id = None
    try:
        type_id = system.GetTypeId()
    except Exception:
        type_id = None
    if type_id is not None and hasattr(type_id, 'IntegerValue') and type_id.IntegerValue > 0:
        type_element = doc.GetElement(type_id)
        if type_element is not None:
            type_name = getattr(type_element, 'Name', '') or ''
    if type_name:
        return type_name
    try:
        type_name = param_str(system.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM))
    except Exception:
        type_name = ''
    return type_name


def join_values(values):
    if not values:
        return ''
    return ', '.join(sorted(values))


def _get_builtin_parameter(element, name):
    if element is None or not name:
        return None
    builtin = getattr(BuiltInParameter, name, None)
    if builtin is None:
        return None
    try:
        return element.get_Parameter(builtin)
    except Exception:
        return None


def get_parameter_value(element, builtin_names, fallback_names):
    for builtin_name in builtin_names:
        param = _get_builtin_parameter(element, builtin_name)
        if param is None:
            continue
        value = param_str(param)
        if value:
            return value
    if element is not None:
        for param_name in fallback_names:
            if not param_name:
                continue
            try:
                param = element.LookupParameter(param_name)
            except Exception:
                param = None
            if param is None:
                continue
            value = param_str(param)
            if value:
                return value
    return ''


def system_stats(doc, system):
    try:
        raw_items = list(system.Elements)
    except Exception:
        raw_items = []
    length_feet = 0.0
    elements = []
    workset_names = set()
    material_names = set()
    for item in raw_items:
        element = item
        if hasattr(item, 'IntegerValue'):
            element = doc.GetElement(item)
        if element is None:
            continue
        elements.append(element)
        try:
            param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        except Exception:
            param = None
        if param is not None:
            try:
                value = param.AsDouble()
            except Exception:
                value = 0.0
            if value:
                length_feet += value
        try:
            ws_param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        except Exception:
            ws_param = None
        workset_value = param_str(ws_param) if ws_param is not None else ''
        if workset_value:
            workset_names.add(workset_value)
        try:
            material_ids = element.GetMaterialIds(False)
        except Exception:
            material_ids = []
        for mat_id in material_ids:
            if hasattr(mat_id, 'IntegerValue') and mat_id.IntegerValue > 0:
                material = doc.GetElement(mat_id)
                if material is not None:
                    mat_name = getattr(material, 'Name', '') or ''
                    if mat_name:
                        material_names.add(mat_name)
    if not workset_names:
        fallback = resolve_workset_name(doc, getattr(system, 'WorksetId', None))
        if fallback:
            workset_names.add(fallback)
    return len(elements), feet_to_m(length_feet), join_values(workset_names), join_values(material_names)


def main():
    doc = get_doc()
    systems = gather_systems(doc)
    if not systems:
        forms.alert('No se encontraron sistemas de HVAC.', title='List HVAC Systems')
        return

    systems_to_export = select_systems(systems)
    if systems_to_export is None:
        return
    if not systems_to_export:
        forms.alert('Seleccione al menos un sistema para exportar.', title='List HVAC Systems')
        return

    filepath = ask_save_csv('HVAC_Systems.csv')
    if not filepath:
        return

    headers = [
        'System Name',
        'System Type',
        'Flow',
        'Static Pressure',
        'Worksets',
        'Materials',
        'Element Count',
        'Total Length (m)'
    ]
    rows = []
    for system in systems_to_export:
        name = system.Name
        system_type = resolve_system_type(doc, system)
        count, length_m, worksets, materials = system_stats(doc, system)
        flow_value = get_parameter_value(system, FLOW_PARAM_NAMES, FLOW_FALLBACK_NAMES)
        pressure_value = get_parameter_value(system, STATIC_PRESSURE_PARAM_NAMES, STATIC_PRESSURE_FALLBACK_NAMES)
        rows.append([
            name,
            system_type,
            flow_value,
            pressure_value,
            worksets,
            materials,
            count,
            '{0:.2f}'.format(length_m),
        ])
    write_csv(filepath, headers, rows)
    forms.alert('CSV generado: {0}'.format(filepath), title='List HVAC Systems')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)



