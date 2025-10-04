# -*- coding: utf-8 -*-
"""Shared helpers for Valtria pyRevit tools (IronPython 2.7 compatible)."""

import codecs
import csv
import os
import traceback

from Autodesk.Revit.DB import FilteredElementCollector

try:
    from RevitServices.Persistence import DocumentManager
except Exception:
    DocumentManager = None

try:
    from pyrevit import revit
except Exception:
    revit = None

try:
    from pyrevit import forms
except Exception:
    forms = None

try:
    unicode
except NameError:
    unicode = str  # type: ignore


_SENTINEL = object()
_FEET_PER_METER = 3.28083989501312
_MM_PER_FOOT = 304.8


# ----------------------------------------------------------------------
# Internal helpers
# ----------------------------------------------------------------------
def _safe_getattr(source, name):
    if source is None:
        return None
    try:
        return getattr(source, name)
    except Exception:
        return None


def _ensure_text(value):
    if value is None:
        return ''
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return ''


def _ensure_directory(path):
    directory = os.path.dirname(path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)


def _alert(message, title='Valtria PyTools', warn_icon=True):
    if forms is None:
        return
    try:
        forms.alert(message, title=title, warn_icon=warn_icon)
    except Exception:
        pass


class _RevitContext(object):
    """Lazily resolves Revit handles and caches them for reuse."""

    def __init__(self):
        self.reset()

    def reset(self):
        self._app = _SENTINEL
        self._uiapp = _SENTINEL
        self._uidoc = _SENTINEL
        self._doc = _SENTINEL

    def _resolve_uiapp(self):
        if DocumentManager is not None:
            dm = _safe_getattr(DocumentManager, 'Instance')
            uiapp = _safe_getattr(dm, 'CurrentUIApplication')
            if uiapp is not None:
                return uiapp
        if revit is not None:
            uiapp = _safe_getattr(revit, 'uiapp')
            if uiapp is not None:
                return uiapp
        global_revit = globals().get('__revit__')
        if global_revit is not None:
            uiapp = _safe_getattr(global_revit, 'Application')
            if uiapp is not None:
                return uiapp
        return None

    def uiapp(self):
        if self._uiapp is _SENTINEL:
            self._uiapp = self._resolve_uiapp()
        return self._uiapp

    def _resolve_uidoc(self):
        uiapp = self.uiapp()
        uidoc = _safe_getattr(uiapp, 'ActiveUIDocument')
        if uidoc is not None:
            return uidoc
        if revit is not None:
            uidoc = _safe_getattr(revit, 'uidoc')
            if uidoc is not None:
                return uidoc
        global_revit = globals().get('__revit__')
        if global_revit is not None:
            uidoc = _safe_getattr(global_revit, 'ActiveUIDocument')
            if uidoc is not None:
                return uidoc
        return None

    def uidoc(self):
        if self._uidoc is _SENTINEL:
            self._uidoc = self._resolve_uidoc()
        return self._uidoc

    def _resolve_doc(self):
        if DocumentManager is not None:
            dm = _safe_getattr(DocumentManager, 'Instance')
            doc = _safe_getattr(dm, 'CurrentDBDocument')
            if doc is not None:
                return doc
        if revit is not None:
            doc = _safe_getattr(revit, 'doc')
            if doc is not None:
                return doc
        uidoc = self.uidoc()
        doc = _safe_getattr(uidoc, 'Document')
        if doc is not None:
            return doc
        global_revit = globals().get('__revit__')
        if global_revit is not None:
            uidoc = _safe_getattr(global_revit, 'ActiveUIDocument')
            doc = _safe_getattr(uidoc, 'Document')
            if doc is not None:
                return doc
        return None

    def doc(self):
        if self._doc is _SENTINEL:
            self._doc = self._resolve_doc()
        return self._doc

    def _resolve_app(self):
        uiapp = self.uiapp()
        app = _safe_getattr(uiapp, 'Application')
        if app is not None:
            return app
        doc = self.doc()
        app = _safe_getattr(doc, 'Application')
        if app is not None:
            return app
        global_revit = globals().get('__revit__')
        if global_revit is not None:
            app = _safe_getattr(global_revit, 'Application')
            if app is not None:
                return app
        return None

    def app(self):
        if self._app is _SENTINEL:
            self._app = self._resolve_app()
        return self._app


_CONTEXT = _RevitContext()


# ----------------------------------------------------------------------
# Public API
# ----------------------------------------------------------------------
def refresh_revit_context():
    """Clear cached handles so they are resolved again on next access."""
    _CONTEXT.reset()
    return _CONTEXT


def get_uiapp():
    """Return the current UIApplication instance."""
    uiapp = _CONTEXT.uiapp()
    if uiapp is None:
        raise RuntimeError('Unable to resolve Revit UIApplication.')
    return uiapp


def get_app():
    """Return the Revit application object."""
    app = _CONTEXT.app()
    if app is None:
        raise RuntimeError('Unable to resolve Revit application object.')
    return app


def get_uidoc():
    """Return the active Revit UIDocument."""
    uidoc = _CONTEXT.uidoc()
    if uidoc is None:
        raise RuntimeError('Unable to resolve active UIDocument.')
    return uidoc


def get_doc():
    """Return the active Revit Document."""
    doc = _CONTEXT.doc()
    if doc is None:
        raise RuntimeError('Unable to resolve active Revit document.')
    return doc


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
    if forms is None:
        raise RuntimeError('pyRevit forms module not available.')
    return forms.save_file(file_ext='csv', default_name=_ensure_text(default_name))


def write_csv(filepath, headers, rows):
    """Write rows to CSV using UTF-8 BOM."""
    if not filepath:
        return
    _ensure_directory(filepath)
    stream = codecs.open(filepath, 'w', 'utf-8-sig')
    try:
        writer = csv.writer(stream)
        if headers:
            writer.writerow([_ensure_text(value) for value in headers])
        for row in rows or []:
            writer.writerow([_ensure_text(value) for value in row])
    finally:
        stream.close()


def to_unicode(value):
    """Backward compatible helper that coerces values to unicode strings."""
    return _ensure_text(value)


def select_views(doc, view_ids):
    """Return view elements matching the provided ids."""
    results = []
    if not view_ids:
        return results
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
        if element is None:
            continue
        try:
            bbox = element.get_BoundingBox(view3d)
        except Exception:
            bbox = None
        if bbox is None or bbox.Min is None or bbox.Max is None:
            continue
        if min_x is None:
            min_x = bbox.Min.X
            min_y = bbox.Min.Y
            min_z = bbox.Min.Z
            max_x = bbox.Max.X
            max_y = bbox.Max.Y
            max_z = bbox.Max.Z
            continue
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
        storage_type = _safe_getattr(param, 'StorageType')
        if storage_type is not None and storage_type.ToString() == 'String':
            value = param.AsString()
            if value:
                return _ensure_text(value)
        value = param.AsValueString()
        if value:
            return _ensure_text(value)
    except Exception:
        pass
    for accessor_name in ('AsString', 'AsInteger', 'AsDouble'):
        accessor = _safe_getattr(param, accessor_name)
        if accessor is None:
            continue
        try:
            value = accessor()
        except Exception:
            continue
        if value is not None:
            return _ensure_text(value)
    return ''


def system_name_of(element):
    """Return the associated MEP system name if available."""
    if element is None:
        return ''
    mep_system = _safe_getattr(element, 'MEPSystem')
    name = _safe_getattr(mep_system, 'Name')
    if name:
        return _ensure_text(name)
    connector_manager = None
    mep_model = _safe_getattr(element, 'MEPModel')
    if mep_model is not None:
        connector_manager = _safe_getattr(mep_model, 'ConnectorManager')
    if connector_manager is None:
        connector_manager = _safe_getattr(element, 'ConnectorManager')
    if connector_manager is None:
        return ''
    try:
        connectors = connector_manager.Connectors
    except Exception:
        return ''
    for connector in connectors:
        system = _safe_getattr(connector, 'MEPSystem')
        name = _safe_getattr(system, 'Name')
        if name:
            return _ensure_text(name)
    return ''


def log_exception(exc, title='Valtria PyTools'):
    """Print stack trace and alert the user."""
    traceback.print_exc()
    message = _ensure_text(exc)
    if not message:
        message = 'Unexpected error.'
    _alert(message, title=title, warn_icon=True)


__all__ = [
    'get_app',
    'get_uiapp',
    'get_uidoc',
    'get_doc',
    'refresh_revit_context',
    'mm_to_internal',
    'feet_to_m',
    'ask_save_csv',
    'write_csv',
    'to_unicode',
    'select_views',
    'get_all_visible_model_boundingbox',
    'param_str',
    'system_name_of',
    'log_exception',
]