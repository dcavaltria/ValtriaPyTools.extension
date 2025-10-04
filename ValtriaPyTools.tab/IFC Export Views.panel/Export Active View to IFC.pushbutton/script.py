# -*- coding: utf-8 -*-
"""Export the active view to IFC using safe Revit API overloads."""

import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..'))
LIB_DIR = os.path.join(EXTENSION_DIR, '_lib')
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import ElementId, IFCExportOptions
try:
    from Autodesk.Revit.DB.IFC import (
        ExporterIFC,
        IFCExportConfiguration,
        IFCExportConfigurationsMap,
        IFCVersion,
    )
except ImportError:
    ExporterIFC = None
    IFCExportConfiguration = None
    IFCExportConfigurationsMap = None
    IFCVersion = None

from pyrevit import forms

try:
    from _lib.valtria_lib import get_doc, log_exception
except ImportError:
    forms.alert(
        'No se pudo cargar _lib.valtria_lib. Actualiza la extension.',
        title='Export Active View to IFC',
        warn_icon=True,
    )
    raise

try:
    unicode
except NameError:  # Python 3 compatibility when running tests
    unicode = str  # type: ignore


_BLOCKED_VIEWTYPES = {
    'Schedule',
    'ProjectBrowser',
    'DrawingSheet',
    'Legend',
}

_DEFAULT_IFC_VERSION = None
if IFCVersion is not None:
    for candidate in ('IFC2x3CV2', 'IFC2x3'):
        if hasattr(IFCVersion, candidate):
            _DEFAULT_IFC_VERSION = getattr(IFCVersion, candidate)
            break


class DebugCollector(object):
    """Collects messages and shows them at the end when enabled."""

    def __init__(self, enabled=False):
        self.enabled = enabled
        self.messages = []

    def add(self, message):
        if not self.enabled:
            return
        try:
            self.messages.append(unicode(message))
        except Exception:
            self.messages.append(str(message))

    def dump(self, title='IFC Export Debug'):
        if not self.enabled or not self.messages:
            return
        try:
            forms.alert('\n'.join(self.messages), title=title)
        except Exception:
            for message in self.messages:
                print(message)


def _ensure_text(value):
    if value is None:
        return ''
    return value if isinstance(value, unicode) else unicode(value)


def is_exportable(view):
    if view is None or getattr(view, 'IsTemplate', False):
        return False
    view_type = getattr(view, 'ViewType', None)
    if view_type is None:
        return False
    return view_type.ToString() not in _BLOCKED_VIEWTYPES


def clean_name(name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', name or '')
    return safe or 'View'


def load_ifc_configurations(doc, debug):
    configs = []
    if IFCExportConfiguration is None or IFCExportConfigurationsMap is None:
        debug.add('IFC configuration API not available.')
        return configs
    try:
        config_map = IFCExportConfigurationsMap()
    except Exception as map_error:
        debug.add('Could not create configurations map: {0}'.format(map_error))
        return configs
    loaders = (
        ('AddBuiltInConfigurations', (config_map,)),
        ('LoadDefaultConfigurations', (config_map,)),
        ('LoadConfigurations', (doc, config_map)),
        ('LoadSavedConfigurations', (doc, config_map)),
        ('LoadUserConfigurations', (config_map,)),
        ('AddSavedConfigurations', (doc, config_map)),
    )
    for method_name, args in loaders:
        method = getattr(IFCExportConfiguration, method_name, None)
        if method is None:
            continue
        try:
            method(*args)
        except Exception as call_error:
            debug.add('{0} failed: {1}'.format(method_name, call_error))
    try:
        enumerator = config_map.GetEnumerator()
    except Exception as enum_error:
        debug.add('Could not enumerate configurations: {0}'.format(enum_error))
        return configs
    seen = set()
    while enumerator.MoveNext():
        current = enumerator.Current
        config = getattr(current, 'Value', None)
        if config is None:
            continue
        name = _ensure_text(getattr(config, 'Name', None))
        if name in seen:
            continue
        seen.add(name)
        configs.append(config)
    debug.add('Found {0} IFC configurations.'.format(len(configs)))
    return configs


def pick_ifc_configuration(doc, debug):
    configs = load_ifc_configurations(doc, debug)
    if not configs:
        debug.add('Using default IFC options (no configuration found).')
        return None
    labels = []
    mapping = {}
    for config in configs:
        name = _ensure_text(getattr(config, 'Name', 'Configuration'))
        schema = _ensure_text(getattr(config, 'IFCVersion', ''))
        label = '{0} ({1})'.format(name, schema) if schema else name
        labels.append(label)
        mapping[label] = config
    choice = forms.ask_for_one_item(
        labels,
        default=labels[0],
        prompt='Select IFC configuration',
        title='Export Active View to IFC',
    )
    debug.add('Selected configuration: {0}'.format(choice or 'None'))
    if not choice:
        return None
    return mapping.get(choice)


def _apply_view_to_options(options, view, debug):
    for attr_name in ('FilterViewId', 'ActiveViewId'):
        if hasattr(options, attr_name):
            try:
                setattr(options, attr_name, view.Id)
                debug.add('Set {0} to view id {1}'.format(attr_name, view.Id.IntegerValue))
                return
            except Exception as assign_error:
                debug.add('Could not set {0}: {1}'.format(attr_name, assign_error))
    try:
        options.AddOption('ActiveViewId', unicode(view.Id.IntegerValue))
        debug.add('Stored ActiveViewId via AddOption.')
    except Exception as add_error:
        debug.add('Could not store ActiveViewId option: {0}'.format(add_error))


def _apply_configuration(configuration, doc, view, options, debug):
    if configuration is None:
        return options
    update = getattr(configuration, 'UpdateOptions', None)
    if update is not None:
        for args in ((options, doc, view.Id), (options, doc, view), (options, doc)):
            try:
                update(*args)
                debug.add('Applied configuration with UpdateOptions{0}.'.format(args))
                return options
            except TypeError:
                continue
            except Exception as call_error:
                debug.add('UpdateOptions failed: {0}'.format(call_error))
    get_opts = getattr(configuration, 'GetExportOptions', None)
    if get_opts is not None:
        for args in ((doc, view.Id), (doc, view), (doc,)):
            try:
                clone = get_opts(*args)
            except TypeError:
                continue
            except Exception as call_error:
                debug.add('GetExportOptions failed: {0}'.format(call_error))
                continue
            if isinstance(clone, IFCExportOptions):
                debug.add('Cloned options from configuration.')
                return clone
    debug.add('Configuration could not be applied; using defaults.')
    return options


def build_ifc_options(doc, view, configuration, debug):
    options = IFCExportOptions()
    if _DEFAULT_IFC_VERSION is not None:
        try:
            options.FileVersion = _DEFAULT_IFC_VERSION
            debug.add('Default IFC version: {0}'.format(_DEFAULT_IFC_VERSION))
        except Exception as version_error:
            debug.add('Could not assign default IFC version: {0}'.format(version_error))
    options = _apply_configuration(configuration, doc, view, options, debug)
    _apply_view_to_options(options, view, debug)
    return options


def _ensure_ifc_path(folder, name):
    base = clean_name(name)[:150]
    full_path = os.path.join(folder, base)
    if not full_path.lower().endswith('.ifc'):
        full_path += '.ifc'
    return full_path, base


def export_view(doc, view, target_folder, configuration, debug):
    options = build_ifc_options(doc, view, configuration, debug)
    full_path, file_name = _ensure_ifc_path(target_folder, view.Name)
    debug.add('Target IFC path: {0}'.format(full_path))
    exporter_success = False
    if ExporterIFC is not None:
        try:
            ExporterIFC.ExportDoc(doc, full_path, options)
            debug.add('ExporterIFC.ExportDoc finished without raising.')
            exporter_success = True
        except Exception as exporter_error:
            debug.add('ExporterIFC.ExportDoc failed: {0}'.format(exporter_error))
            debug.add(traceback.format_exc())
    if exporter_success:
        return True
    try:
        result = doc.Export(target_folder, file_name, options)
        debug.add('doc.Export(folder, filename, options) returned: {0}'.format(result))
        return bool(result)
    except Exception as export_error:
        debug.add('doc.Export(folder, filename, options) failed: {0}'.format(export_error))
        debug.add(traceback.format_exc())
        return False


def main():
    doc = get_doc()
    active_view = doc.ActiveView
    debug_choice = forms.ask_for_one_item(
        ['No', 'Yes'],
        default='No',
        prompt='Mostrar informacion de depuracion?',
        title='Export Active View to IFC',
    )
    debug = DebugCollector(enabled=(debug_choice == 'Yes'))
    debug.add('Debugging enabled: {0}'.format(debug.enabled))
    debug.add('Active view: {0} ({1})'.format(
        _ensure_text(getattr(active_view, 'Name', '')), getattr(active_view, 'Id', ElementId.InvalidElementId).IntegerValue if active_view else 'None'))
    if not is_exportable(active_view):
        forms.alert('Activa una vista de modelo (no plantilla) antes de exportar.', title='Export Active View to IFC', warn_icon=True)
        debug.add('Active view not exportable.')
        debug.dump()
        return
    configuration = pick_ifc_configuration(doc, debug)
    export_folder = forms.pick_folder(title='Selecciona carpeta destino IFC')
    debug.add('Export folder: {0}'.format(export_folder or 'None'))
    if not export_folder:
        debug.dump()
        return
    if not os.path.exists(export_folder):
        os.makedirs(export_folder)
        debug.add('Created export folder.')
    try:
        success = export_view(doc, active_view, export_folder, configuration, debug)
    except Exception as exc:
        log_exception(exc)
        debug.add('Unexpected exception: {0}'.format(exc))
        debug.add(traceback.format_exc())
        debug.dump()
        return
    if success:
        forms.alert('Exportacion IFC completada.', title='Export Active View to IFC')
    else:
        forms.alert('No se pudo exportar la vista activa.', title='Export Active View to IFC', warn_icon=True)
    debug.dump()


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)


