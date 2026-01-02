# -*- coding: utf-8 -*-
"""Shared helpers for Valtria pyRevit tools (IronPython 2.7 compatible)."""

import codecs
import csv
import datetime
import os
import json
import traceback

import clr

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    ElementId,
    StorageType,
    Transaction,
    ExternalDefinitionCreationOptions,
    InstanceBinding,
    TypeBinding,
    CategorySet,
    BuiltInParameterGroup,
)

try:
    clr.AddReference('Microsoft.Office.Interop.Excel')
    from Microsoft.Office.Interop import Excel as ExcelInterop
except Exception:
    ExcelInterop = None

try:
    clr.AddReference('System.Runtime.InteropServices')
    from System.Runtime.InteropServices import Marshal
except Exception:
    Marshal = None

try:
    clr.AddReference('System.Windows.Forms')
    from System.Windows.Forms import SaveFileDialog, DialogResult
except Exception:
    SaveFileDialog = None
    DialogResult = None

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
_METERS_PER_FOOT = 0.3048
_SQUARE_METERS_PER_SQUARE_FOOT = _METERS_PER_FOOT * _METERS_PER_FOOT
_CUBIC_METERS_PER_CUBIC_FOOT = _METERS_PER_FOOT * _METERS_PER_FOOT * _METERS_PER_FOOT


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


def _release_com(obj):
    if obj is None or Marshal is None:
        return
    try:
        Marshal.FinalReleaseComObject(obj)
    except Exception:
        try:
            Marshal.ReleaseComObject(obj)
        except Exception:
            pass

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
_EXTENSION_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))
_LOG_DIR = os.path.join(_EXTENSION_DIR, '_logs')


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


def get_selected_element_ids():
    """Return list of selected element ids or empty list."""
    uidoc = get_uidoc()
    if uidoc is None:
        return []
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        return []
    return list(ids)


def get_selected_elements():
    """Return list of selected Element instances."""
    doc = get_doc()
    if doc is None:
        return []
    elements = []
    for elem_id in get_selected_element_ids():
        try:
            element = doc.GetElement(elem_id)
        except Exception:
            element = None
        if element is not None:
            elements.append(element)
    return elements


def get_elements_in_active_view(include_types=False):
    """Return elements visible in the active view."""
    doc = get_doc()
    if doc is None:
        return []
    view = get_uidoc().ActiveView if get_uidoc() else None
    if view is None:
        return []
    collector = FilteredElementCollector(doc, view.Id)
    if not include_types:
        collector = collector.WhereElementIsNotElementType()
    try:
        return list(collector)
    except Exception:
        return []


def get_element_type(element):
    """Return the ElementType of the given element if possible."""
    if element is None:
        return None
    doc = get_doc()
    if doc is None:
        return None
    try:
        type_id = element.GetTypeId()
    except Exception:
        type_id = None
    if not type_id or type_id == ElementId.InvalidElementId:
        return None
    try:
        return doc.GetElement(type_id)
    except Exception:
        return None


def get_element_category_bic(element):
    """Return BuiltInCategory for the given element if available."""
    cat = _safe_getattr(element, 'Category')
    if cat is None:
        return None
    try:
        return cat.BuiltInCategory
    except Exception:
        return None


def _parameter_to_value(param):
    if param is None:
        return None
    try:
        stype = param.StorageType
    except Exception:
        stype = None
    if stype == StorageType.String:
        try:
            return param.AsString()
        except Exception:
            return None
    if stype == StorageType.Double:
        try:
            return param.AsDouble()
        except Exception:
            return None
    if stype == StorageType.Integer:
        try:
            return param.AsInteger()
        except Exception:
            return None
    if stype == StorageType.ElementId:
        try:
            eid = param.AsElementId()
        except Exception:
            return None
        if eid and eid != ElementId.InvalidElementId:
            doc = get_doc()
            if doc is not None:
                try:
                    target = doc.GetElement(eid)
                    name = _safe_getattr(target, 'Name')
                    if name:
                        return _ensure_text(name)
                except Exception:
                    pass
            return eid.IntegerValue
        return None
    try:
        return param.AsValueString()
    except Exception:
        return None


def get_param_value(element, name, default=None):
    """Return parameter value by name or default."""
    if element is None:
        return default
    try:
        param = element.LookupParameter(name)
    except Exception:
        param = None
    value = _parameter_to_value(param)
    return default if value is None else value


def set_param_value(element, name, value):
    """Set parameter value by name converting to appropriate storage type."""
    if element is None:
        raise Exception(u"No element provided.")
    try:
        param = element.LookupParameter(name)
    except Exception:
        param = None
    if param is None:
        raise Exception(u"Parametro '{0}' no existe en el elemento {1}".format(name, element.Id))
    if param.IsReadOnly:
        raise Exception(u"Parametro '{0}' es de solo lectura".format(name))

    stype = param.StorageType
    if stype == StorageType.String:
        return param.Set(u'' if value is None else unicode(value))
    if stype == StorageType.Double:
        return param.Set(0.0 if value is None else float(value))
    if stype == StorageType.Integer:
        return param.Set(0 if value is None else int(value))
    if stype == StorageType.ElementId:
        if value is None:
            return param.Set(ElementId.InvalidElementId)
        if isinstance(value, ElementId):
            return param.Set(value)
        return param.Set(ElementId(int(value)))
    raise Exception(u"Tipo de almacenamiento no soportado para '{0}'".format(name))


def ensure_shared_parameter(param_name, param_type, group_name, categories,
                            is_instance=True, param_group=BuiltInParameterGroup.PG_DATA):
    """Ensure a shared parameter exists and is bound to given categories."""
    app = get_app()
    doc = get_doc()
    if app is None or doc is None:
        raise Exception(u"No se pudo obtener el contexto de Revit.")

    sp_path = app.SharedParametersFilename
    if not sp_path or not os.path.exists(sp_path):
        sp_path = os.path.join(os.environ.get("APPDATA", ""), "pyRevit_shared_params.txt")
        if not os.path.exists(sp_path):
            with open(sp_path, "w") as fh:
                fh.write("# pyRevit shared params\n")
        app.SharedParametersFilename = sp_path
        info(u"Usando shared parameters: {0}".format(sp_path))

    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception(u"No se pudo abrir el archivo de shared parameters.")

    group = None
    for grp in def_file.Groups:
        if grp.Name == group_name:
            group = grp
            break
    if group is None:
        group = def_file.Groups.Create(group_name)

    definition = None
    for definition_item in group.Definitions:
        if definition_item.Name == param_name:
            definition = definition_item
            break
    if definition is None:
        from Autodesk.Revit.DB import ParameterType
        pt_map = {
            "Text": ParameterType.Text,
            "Number": ParameterType.Number,
            "YesNo": ParameterType.YesNo,
        }
        parameter_type = pt_map.get(param_type, ParameterType.Text)
        options = ExternalDefinitionCreationOptions(param_name, parameter_type)
        definition = group.Definitions.Create(options)

    catset = CategorySet()
    for bic in categories:
        category = Category.GetCategory(doc, bic)
        if category:
            catset.Insert(category)

    binding = InstanceBinding(catset) if is_instance else TypeBinding(catset)
    tx = Transaction(doc, u"Asegurar parametro '{0}'".format(param_name))
    tx.Start()
    try:
        bindings = doc.ParameterBindings
        if not bindings.Insert(definition, binding, param_group):
            bindings.ReInsert(definition, binding, param_group)
        tx.Commit()
    except Exception:
        tx.RollBack()
        raise


def read_length(element):
    """Return element length in internal units if available."""
    value = get_param_value(element, "Length")
    if isinstance(value, (int, float)):
        return float(value)
    try:
        param = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
    except Exception:
        param = None
    return _parameter_to_value(param)


def read_area(element):
    """Return element area in internal units if available."""
    value = get_param_value(element, "Area")
    if isinstance(value, (int, float)):
        return float(value)
    for bip in (BuiltInParameter.HOST_AREA_COMPUTED, BuiltInParameter.RBS_CURVE_SURFACE_AREA):
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        result = _parameter_to_value(param)
        if result is not None:
            return result
    return None


def read_volume(element):
    """Return element volume in internal units if available."""
    value = get_param_value(element, "Volume")
    if isinstance(value, (int, float)):
        return float(value)
    candidate_names = [
        "HOST_VOLUME_COMPUTED",
        "RBS_CURVE_VOLUME",
        "RBS_PIPE_VOLUME_PARAM",
        "RBS_DUCT_VOLUME_PARAM",
        "RBS_MECHANICAL_EQUIPMENT_VOLUME"
    ]
    for name in candidate_names:
        bip = getattr(BuiltInParameter, name, None)
        if bip is None:
            continue
        try:
            param = element.get_Parameter(bip)
        except Exception:
            param = None
        result = _parameter_to_value(param)
        if isinstance(result, (int, float)):
            return float(result)
    return None


def measure_elements(elements):
    """Return measurement summary (length/area/volume) for elements list."""
    summary = {
        "count": len(elements),
        "sum_length_ft": 0.0,
        "sum_area_ft2": 0.0,
        "sum_volume_ft3": 0.0,
    }
    for elem in elements:
        summary["sum_length_ft"] += read_length(elem) or 0.0
        summary["sum_area_ft2"] += read_area(elem) or 0.0
        summary["sum_volume_ft3"] += read_volume(elem) or 0.0
    summary["sum_length_m"] = (summary["sum_length_ft"] or 0.0) * _METERS_PER_FOOT
    summary["sum_area_m2"] = (summary["sum_area_ft2"] or 0.0) * _SQUARE_METERS_PER_SQUARE_FOOT
    summary["sum_volume_m3"] = (summary["sum_volume_ft3"] or 0.0) * _CUBIC_METERS_PER_CUBIC_FOOT
    return summary


def mep_attributes(element):
    """Return dictionary with common MEP attributes if present."""
    data = {}
    if element is None:
        return data
    system_name = system_name_of(element)
    if system_name:
        data["system_name"] = system_name

    name_param = None
    try:
        name_param = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    except Exception:
        name_param = None
    val = _parameter_to_value(name_param)
    if val and not data.get("system_name"):
        data["system_name"] = _ensure_text(val)

    abbrev = None
    try:
        abbrev = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
    except Exception:
        abbrev = None
    val = _parameter_to_value(abbrev)
    if val:
        data["system_abbreviation"] = _ensure_text(val)

    doc = get_doc()
    system_type = None
    try:
        system_type_param = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_TYPE_PARAM)
    except Exception:
        system_type_param = None
    if system_type_param is not None:
        st_value = _parameter_to_value(system_type_param)
        if isinstance(st_value, (int, float)):
            if doc is not None:
                try:
                    type_element = doc.GetElement(ElementId(int(st_value)))
                    st_value = _safe_getattr(type_element, "Name")
                except Exception:
                    pass
        if st_value:
            system_type = _ensure_text(st_value)
    if system_type:
        data["system_type"] = system_type

    classification = None
    try:
        classification = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
    except Exception:
        classification = None
    val = _parameter_to_value(classification)
    if val:
        data["system_classification"] = _ensure_text(val)

    flow_params = [
        getattr(BuiltInParameter, "RBS_DUCT_FLOW_PARAM", None),
        getattr(BuiltInParameter, "RBS_PIPE_FLOW_PARAM", None),
        getattr(BuiltInParameter, "RBS_FLOW_PARAM", None),
    ]
    for bip in flow_params:
        if bip is None:
            continue
        try:
            p = element.get_Parameter(bip)
        except Exception:
            p = None
        val = _parameter_to_value(p)
        if isinstance(val, (int, float)):
            data["flow"] = float(val)
            data["flow_parameter"] = bip.ToString()
            break

    size_param = None
    try:
        size_param = element.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)
    except Exception:
        size_param = None
    val = _parameter_to_value(size_param)
    if val:
        data["calculated_size"] = _ensure_text(val)

    return data


def element_snapshot(element, max_params=8, include_mep=True):
    """Return dictionary snapshot of key element data."""
    snapshot = {
        "id": None,
        "category": "",
        "name": "",
        "type_name": "",
        "parameters": {},
    }
    if element is None:
        return snapshot
    try:
        snapshot["id"] = int(element.Id.IntegerValue)
    except Exception:
        snapshot["id"] = None
    snapshot["category"] = _ensure_text(_safe_getattr(_safe_getattr(element, "Category"), "Name"))
    snapshot["name"] = _ensure_text(_safe_getattr(element, "Name"))
    etype = get_element_type(element)
    snapshot["type_name"] = _ensure_text(_safe_getattr(etype, "Name"))

    captured = 0
    for param in getattr(element, "Parameters", []):
        if captured >= max_params:
            break
        try:
            definition = _safe_getattr(param, "Definition")
            pname = _safe_getattr(definition, "Name")
        except Exception:
            pname = None
        if not pname:
            continue
        value = _parameter_to_value(param)
        if value is None:
            continue
        text_value = _ensure_text(value)
        if not text_value or len(text_value) > 80:
            continue
        snapshot["parameters"][pname] = text_value
        captured += 1

    if include_mep:
        snapshot["mep"] = mep_attributes(element)
    return snapshot


def build_context_for_elements(elements, max_elements=None, max_params=8, char_limit=12000):
    """
    Build human-readable context text for the provided elements plus summary.
    Returns (text, summary_dict).
    """
    if not elements:
        return "No hay elementos seleccionados.", {"count": 0}

    elements = list(elements)
    if max_elements and max_elements > 0:
        selected = elements[:max_elements]
    else:
        selected = elements
    notes = []
    if max_elements and max_elements > 0 and len(elements) > max_elements:
        notes.append("Se limitaron los primeros {0} de {1} elementos.".format(max_elements, len(elements)))

    summary = measure_elements(selected)

    blocks = []
    rows = []
    for idx, elem in enumerate(selected, 1):
        snap = element_snapshot(elem, max_params=max_params, include_mep=True)
        lines = [
            "Elemento {0}".format(idx),
            "    Categoria: {0}".format(snap.get("category", "")),
            "    Nombre: {0}".format(snap.get("name", "")),
        ]
        if snap.get("type_name"):
            lines.append("    Tipo: {0}".format(snap.get("type_name", "")))
        for pname, pvalue in snap.get("parameters", {}).items():
            lines.append("    {0}: {1}".format(pname, pvalue))
        mep = snap.get("mep") or {}
        for key in ("system_name", "system_abbreviation", "system_type", "system_classification", "calculated_size"):
            if mep.get(key):
                lines.append("    {0}: {1}".format(key, mep[key]))
        if "flow" in mep:
            flow_key = mep.get("flow_parameter", "flow")
            lines.append("    {0}: {1}".format(flow_key, mep["flow"]))
        blocks.append("\n".join(lines))

        row = {
            "id": snap.get("id"),
            "category": snap.get("category", ""),
            "name": snap.get("name", ""),
            "type": snap.get("type_name", ""),
        }
        length_ft = read_length(elem)
        if length_ft:
            row["length_m"] = float(length_ft) * _METERS_PER_FOOT
        area_ft2 = read_area(elem)
        if area_ft2:
            row["area_m2"] = float(area_ft2) * _SQUARE_METERS_PER_SQUARE_FOOT
        volume_ft3 = read_volume(elem)
        if volume_ft3:
            row["volume_m3"] = float(volume_ft3) * _CUBIC_METERS_PER_CUBIC_FOOT
        for pname, pvalue in snap.get("parameters", {}).items():
            row[pname] = pvalue
        for key, value in mep.items():
            row[key] = value
        rows.append(row)

    detail_text = "\n\n".join(blocks)
    if len(detail_text) > char_limit:
        trimmed = detail_text[:char_limit]
        last_break = trimmed.rfind("\n")
        if last_break > 0:
            trimmed = trimmed[:last_break]
        detail_text = trimmed
        notes.append("Se truncaron los datos para respetar el limite de {0} caracteres.".format(char_limit))

    summary_lines = [
        "Resumen (metric):",
        " - Elementos: {0}".format(summary.get("count", 0)),
        " - Largo total: {0:.3f} m".format(summary.get("sum_length_m", 0.0)),
        " - Area total: {0:.3f} m2".format(summary.get("sum_area_m2", 0.0)),
        " - Volumen total: {0:.3f} m3".format(summary.get("sum_volume_m3", 0.0)),
    ]
    if notes:
        summary_lines.append("Nota: " + " ".join(notes))

    context_text = "\n".join(summary_lines) + "\n\nElementos:\n" + detail_text
    return context_text, summary, rows


def log_exception(exc, title='Valtria PyTools'):
    """Print stack trace and alert the user."""
    traceback.print_exc()
    message = _ensure_text(exc)
    if not message:
        message = 'Unexpected error.'
    try:
        tb = traceback.format_exc()
        log_to_file('errors', message + "\n" + (tb or ''))
    except Exception:
        pass
    _alert(message, title=title, warn_icon=True)


def _collect_columns(rows):
    columns = []
    seen = set()
    preferred = [
        "id",
        "category",
        "name",
        "type",
        "length_m",
        "area_m2",
        "volume_m3",
        "system_name",
        "system_abbreviation",
        "system_type",
        "system_classification",
        "flow",
        "calculated_size",
    ]
    for col in preferred:
        for row in rows:
            if col in row and col not in seen:
                seen.add(col)
                columns.append(col)
                break
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                columns.append(key)
    return columns


def _prepare_value(value):
    if value is None:
        return ''
    if isinstance(value, (int, float)):
        return value
    return _ensure_text(value)


def _log_file_path(tool_name):
    safe = _ensure_text(tool_name or 'general')
    filtered = []
    for ch in safe:
        if ch.isalnum() or ch in ('-', '_'):
            filtered.append(ch)
        else:
            filtered.append('_')
    filename = ''.join(filtered) or 'general'
    return os.path.join(_LOG_DIR, filename + '.log')


def log_to_file(tool_name, message):
    """Append message to a per-tool log file under _logs."""
    try:
        if not os.path.isdir(_LOG_DIR):
            os.makedirs(_LOG_DIR)
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        line = u"[{0}] {1}".format(timestamp, _ensure_text(message))
        path = _log_file_path(tool_name)
        stream = codecs.open(path, 'a', 'utf-8')
        try:
            stream.write(line + u'\n')
        finally:
            stream.close()
        return path
    except Exception:
        return None


def export_rows_to_csv(rows, path, columns=None):
    if columns is None:
        columns = _collect_columns(rows)
    if not columns:
        columns = ['data']
    _ensure_directory(path)
    with codecs.open(path, 'w', encoding='utf-8-sig') as fh:
        writer = csv.writer(fh)
        writer.writerow(columns)
        for row in rows:
            writer.writerow([_prepare_value(row.get(col)) for col in columns])
    return path


def export_rows_to_json(rows, path):
    _ensure_directory(path)
    with codecs.open(path, 'w', encoding='utf-8') as fh:
        json.dump(rows, fh, ensure_ascii=False, indent=2)
    return path


def export_rows_to_excel(rows, path, columns=None):
    if ExcelInterop is None:
        raise Exception(u"No se puede exportar a Excel: componente interop no disponible.")
    if columns is None:
        columns = _collect_columns(rows)
    if not columns:
        columns = ['data']
    if not path.lower().endswith('.xlsx'):
        path += '.xlsx'
    _ensure_directory(path)
    excel = None
    wb = None
    ws = None
    try:
        excel = ExcelInterop.ApplicationClass()
        excel.Visible = False
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets.Item[1]
        for c, column in enumerate(columns, 1):
            ws.Cells(1, c).Value2 = column
        for r, row in enumerate(rows, 2):
            for c, column in enumerate(columns, 1):
                value = _prepare_value(row.get(column))
                ws.Cells(r, c).Value2 = value
        wb.SaveAs(path)
        wb.Close(False)
        excel.Quit()
    finally:
        _release_com(ws)
        _release_com(wb)
        _release_com(excel)
    return path


def export_rows(rows, fmt='csv', path=None):
    if rows is None:
        raise Exception(u"No hay datos para exportar.")
    fmt = (fmt or 'csv').lower()
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    default_dir = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
    ext_map = {
        'csv': 'csv',
        'json': 'json',
        'excel': 'xlsx',
        'xlsx': 'xlsx',
        'xls': 'xlsx'
    }
    ext = ext_map.get(fmt, 'csv')
    target_path = path.strip() if path else ''
    prompt_user = False
    if target_path:
        if not os.path.isabs(target_path):
            target_path = os.path.join(default_dir, target_path)
    else:
        filename = 'claude_export_{0}.{1}'.format(timestamp, ext)
        target_path = os.path.join(default_dir, filename)
        prompt_user = True

    base, current_ext = os.path.splitext(target_path)
    if not current_ext:
        target_path = base + '.' + ext

    if prompt_user:
        target_path = _prompt_save_path(ext, default_dir, os.path.basename(target_path))

    try:
        if fmt == 'csv':
            return export_rows_to_csv(rows, target_path)
        if fmt == 'json':
            return export_rows_to_json(rows, target_path)
        if fmt in ('excel', 'xlsx', 'xls'):
            return export_rows_to_excel(rows, target_path)
        raise Exception(u"Formato de exportacion no soportado: {0}".format(fmt))
    except IOError as io_err:
        fallback = _prompt_save_path(ext, default_dir, os.path.basename(target_path))
        try:
            if fmt == 'csv':
                return export_rows_to_csv(rows, fallback)
            if fmt == 'json':
                return export_rows_to_json(rows, fallback)
            if fmt in ('excel', 'xlsx', 'xls'):
                return export_rows_to_excel(rows, fallback)
        except Exception:
            pass
        raise Exception(u"No se pudo escribir el archivo ({0})".format(io_err))


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
    'get_selected_element_ids',
    'get_selected_elements',
    'get_elements_in_active_view',
    'get_element_type',
    'get_element_category_bic',
    'get_param_value',
    'set_param_value',
    'ensure_shared_parameter',
    'read_length',
    'read_area',
    'read_volume',
    'measure_elements',
    'mep_attributes',
    'element_snapshot',
    'build_context_for_elements',
    'export_rows',
    'export_rows_to_csv',
    'export_rows_to_json',
    'export_rows_to_excel',
    'log_exception',
    'log_to_file',
]

def _prompt_save_path(fmt, default_dir, default_name):
    if SaveFileDialog is None or DialogResult is None:
        return os.path.join(default_dir, default_name)
    dialog = SaveFileDialog()
    dialog.Title = u"Guardar exportacion"
    dialog.FileName = default_name
    if default_dir and os.path.isdir(default_dir):
        dialog.InitialDirectory = default_dir
    filter_text = "CSV (*.csv)|*.csv|JSON (*.json)|*.json|Excel (*.xlsx)|*.xlsx"
    dialog.Filter = filter_text
    fmt_index = {
        'csv': 1,
        'json': 2,
        'excel': 3,
        'xlsx': 3,
        'xls': 3
    }.get(fmt, 1)
    dialog.FilterIndex = fmt_index
    result = dialog.ShowDialog()
    if result == DialogResult.OK:
        return dialog.FileName
    return os.path.join(default_dir, default_name)
