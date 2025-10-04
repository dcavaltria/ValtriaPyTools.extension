# -*- coding: utf-8 -*-
"""Shared helpers for Valtria pyRevit tools (IronPython 2.7 compatible)."""

import os
import csv
import codecs
import traceback

from Autodesk.Revit.DB import (
    FilteredElementCollector,
)
from RevitServices.Persistence import DocumentManager
from pyrevit import forms


_FEET_PER_METER = 3.28083989501312
_MM_PER_FOOT = 304.8


def get_app():
    """Return the Revit application object."""
    return DocumentManager.Instance.CurrentUIApplication.Application


def get_uidoc():
    """Return the active Revit UIDocument."""
    return DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument


def get_doc():
    """Return the active Revit Document."""
    return DocumentManager.Instance.CurrentDBDocument


def mm_to_internal(value_mm):
    """Convert millimetres to Revit internal units (feet)."""
    try:
        return float(value_mm) / _MM_PER_FOOT
    except Exception:
        return 0.0


def feet_to_m(value_feet):
    """Convert Revit internal units (feet) to metres."""
    try:
        return float(value_feet) / _FEET_PER_METER
    except Exception:
        return 0.0


def ask_save_csv(default_name):
    """Prompt the user for a CSV destination path."""
    return forms.save_file(file_ext='csv', default_name=default_name)


def write_csv(filepath, headers, rows):
    """Write rows to CSV using UTF-8 BOM."""
    if not filepath:
        return
    directory = os.path.dirname(filepath)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)
    stream = codecs.open(filepath, 'w', 'utf-8-sig')
    try:
        writer = csv.writer(stream)
        if headers:
            writer.writerow(headers)
        for row in rows:
            writer.writerow([to_unicode(value) for value in row])
    finally:
        stream.close()


def to_unicode(value):
    if value is None:
        return ''
    if isinstance(value, unicode):
        return value
    return unicode(value)


def select_views(doc, view_ids):
    """Return view elements matching the provided ids."""
    results = []
    for vid in view_ids:
        view = doc.GetElement(vid)
        if view is not None:
            results.append(view)
    return results


def get_all_visible_model_boundingbox(doc, view3d, elements=None):
    """Return a combined bounding box of all model elements visible in the view."""
    if elements is None:
        collector = FilteredElementCollector(doc, view3d.Id)
        elements = collector.WhereElementIsNotElementType().ToElements()
    min_x = None
    min_y = None
    min_z = None
    max_x = None
    max_y = None
    max_z = None
    for element in elements:
        try:
            bbox = element.get_BoundingBox(view3d)
        except Exception:
            bbox = None
        if bbox is None:
            continue
        if bbox.Min is None or bbox.Max is None:
            continue
        if min_x is None:
            min_x = bbox.Min.X
            min_y = bbox.Min.Y
            min_z = bbox.Min.Z
            max_x = bbox.Max.X
            max_y = bbox.Max.Y
            max_z = bbox.Max.Z
        else:
            if bbox.Min.X < min_x:
                min_x = bbox.Min.X
            if bbox.Min.Y < min_y:
                min_y = bbox.Min.Y
            if bbox.Min.Z < min_z:
                min_z = bbox.Min.Z
            if bbox.Max.X > max_x:
                max_x = bbox.Max.X
            if bbox.Max.Y > max_y:
                max_y = bbox.Max.Y
            if bbox.Max.Z > max_z:
                max_z = bbox.Max.Z
    if min_x is None:
        return None
    from Autodesk.Revit.DB import BoundingBoxXYZ, XYZ
    bbox = BoundingBoxXYZ()
    bbox.Min = XYZ(min_x, min_y, min_z)
    bbox.Max = XYZ(max_x, max_y, max_z)
    return bbox


def param_str(param):
    """Return a readable string for a Revit parameter."""
    if param is None:
        return ''
    try:
        if param.StorageType.ToString() == 'String':
            value = param.AsString()
            if value:
                return value
        value = param.AsValueString()
        if value:
            return value
    except Exception:
        try:
            value = param.AsString()
            if value:
                return value
        except Exception:
            pass
    try:
        value = param.AsInteger()
        return unicode(value)
    except Exception:
        pass
    try:
        value = param.AsDouble()
        return unicode(value)
    except Exception:
        pass
    return ''


def system_name_of(element):
    """Return the associated MEP system name if available."""
    try:
        mep_system = getattr(element, 'MEPSystem', None)
        if mep_system is not None:
            name = getattr(mep_system, 'Name', None)
            if name:
                return name
    except Exception:
        pass
    connector_manager = None
    try:
        if hasattr(element, 'MEPModel') and element.MEPModel:
            connector_manager = element.MEPModel.ConnectorManager
    except Exception:
        connector_manager = None
    if connector_manager is None and hasattr(element, 'ConnectorManager'):
        try:
            connector_manager = element.ConnectorManager
        except Exception:
            connector_manager = None
    if connector_manager is not None:
        try:
            for connector in connector_manager.Connectors:
                try:
                    system = connector.MEPSystem
                    if system is not None and system.Name:
                        return system.Name
                except Exception:
                    continue
        except Exception:
            pass
    return ''


def log_exception(exc):
    """Print stack trace and alert the user."""
    traceback.print_exc()
    message = unicode(exc)
    forms.alert(message, title='Valtria PyTools', warn_icon=True)

