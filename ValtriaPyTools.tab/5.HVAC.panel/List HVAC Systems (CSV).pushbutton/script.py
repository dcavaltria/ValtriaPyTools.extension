# -*- coding: utf-8 -*-
"""Export a CSV summary of HVAC duct systems."""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
)
from Autodesk.Revit.DB.Mechanical import DuctSystem
from pyrevit import forms

from _lib.valtria_lib import (
    get_doc,
    ask_save_csv,
    write_csv,
    feet_to_m,
    log_exception,
)


def gather_systems(doc):
    systems = []
    collector = FilteredElementCollector(doc).OfClass(DuctSystem)
    for system in collector:
        systems.append(system)
    return systems


def system_stats(doc, system):
    element_ids = list(system.Elements)
    length_feet = 0.0
    for elid in element_ids:
        element = doc.GetElement(elid)
        if element is None:
            continue
        try:
            param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        except Exception:
            param = None
        if param is None:
            continue
        try:
            value = param.AsDouble()
        except Exception:
            value = 0.0
        if value:
            length_feet += value
    return len(element_ids), feet_to_m(length_feet)


def main():
    doc = get_doc()
    systems = gather_systems(doc)
    if not systems:
        forms.alert('No se encontraron sistemas de HVAC.', title='List HVAC Systems')
        return
    filepath = ask_save_csv('HVAC_Systems.csv')
    if not filepath:
        return
    headers = ['System Name', 'Element Count', 'Total Length (m)']
    rows = []
    for system in systems:
        name = system.Name
        count, length_m = system_stats(doc, system)
        rows.append([name, count, length_m])
    write_csv(filepath, headers, rows)
    forms.alert('CSV generado: {0}'.format(filepath), title='List HVAC Systems')


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

