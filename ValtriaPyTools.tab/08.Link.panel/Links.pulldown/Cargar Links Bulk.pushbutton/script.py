# -*- coding: utf-8 -*-
"""Bulk load Revit links (.rvt) from Desktop Connector as local paths."""

import os
import traceback
from datetime import datetime

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    ImportPlacement,
    ModelPathUtils,
    RevitLinkInstance,
    RevitLinkOptions,
    RevitLinkType,
    Transaction,
    Transform,
    WorksetKind,
)

from pyrevit import forms, revit

LOG_FILE = os.path.join(os.path.dirname(__file__), 'cargar_links_bulk.log')

PLACEMENT_LABELS = {
    'origin': 'Origin to Origin',
    'shared': 'Shared Coordinates',
}

LINK_TYPE_DEFINITIONS = [
    {
        'key': 'rvt',
        'label': 'Revit (.rvt)',
        'extensions': {'.rvt'},
        'placement_modes': ['origin', 'shared'],
    },
]


def log_message(text):
    try:
        stamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(LOG_FILE, 'a') as fh:
            fh.write("[{0}] {1}\n".format(stamp, text))
    except Exception:
        pass


class SelectListItem(object):
    def __init__(self, value, label):
        self.value = value
        self.label = label

    @property
    def name(self):
        return self.label


def _normalize_user_path(path):
    if not path:
        return None
    return os.path.normpath(path).replace('\\', '/').lower().rstrip('/')


def is_user_visible_cloud_path(path):
    if not path:
        return False
    low = path.lower()
    return low.startswith('autodesk docs://') or low.startswith('bim 360 docs://') or low.startswith('bim 360://')


def _desktop_connector_parts(file_path):
    """Return (idx, parts, root_type) where root_type is 'autodesk' or None."""
    parts = [p for p in os.path.abspath(file_path).replace('\\', '/').split('/') if p]
    for idx, part in enumerate(parts):
        key = part.replace(' ', '').lower()
        if 'accdocs' in key or 'autodeskdocs' in key or 'dcaccdocs' in key:
            return idx, parts, 'autodesk'
        if 'bim360' in key:
            return idx, parts, 'bim360'
    return None, parts, None


def build_cloud_user_visible_path(file_path):
    """Convert Desktop Connector local path to Autodesk Docs/BIM 360 user-visible path."""
    idx, parts, root_type = _desktop_connector_parts(file_path)
    if idx is None or root_type is None:
        return None
    remainder = parts[idx + 1 :]
    if len(remainder) < 2:
        return None
    base = '/'.join(remainder)
    prefix = 'Autodesk Docs://' if root_type == 'autodesk' else 'BIM 360 Docs://'
    return prefix + base


def resolve_paths(file_path):
    """
    Devuelve un diccionario con variantes de path:
    {
      'local_user': ruta local,
      'local_model': ModelPath local,
      'cloud_user': ruta Autodesk Docs/BIM360 o None,
      'cloud_model': ModelPath de cloud_user (si conversion ok) o None,
      'cloud_is_cloud': bool
    }
    """
    result = {
        'local_user': file_path,
        'local_model': None,
        'cloud_user': None,
        'cloud_model': None,
        'cloud_is_cloud': False,
    }
    try:
        result['local_model'] = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
    except Exception as ex_local:
        log_message("resolve_paths | fallo al convertir local: {0}".format(ex_local))

    cloud_path = build_cloud_user_visible_path(file_path)
    if cloud_path:
        candidates = [cloud_path]
        if cloud_path.startswith('Autodesk Docs://') and not cloud_path.startswith('Autodesk Docs:///'):
            candidates.append(cloud_path.replace('Autodesk Docs://', 'Autodesk Docs:///'))
        if cloud_path.startswith('BIM 360 Docs://') and not cloud_path.startswith('BIM 360 Docs:///'):
            candidates.append(cloud_path.replace('BIM 360 Docs://', 'BIM 360 Docs:///'))

        for cand in candidates:
            log_message("resolve_paths | cloud candidate={0}".format(cand))
            try:
                mp_cloud = ModelPathUtils.ConvertUserVisiblePathToModelPath(cand)
                try:
                    is_cloud = ModelPathUtils.IsCloudPath(mp_cloud)
                except Exception:
                    try:
                        is_cloud = 'Cloud' in (mp_cloud.GetType().Name or '')
                    except Exception:
                        is_cloud = False
                if is_cloud:
                    result['cloud_user'] = cand
                    result['cloud_model'] = mp_cloud
                    result['cloud_is_cloud'] = True
                    log_message("resolve_paths | cloud reconocido")
                    break
                else:
                    # Aunque no lo reconozca, guardar el primero para intentar LoadFrom mas tarde
                    if result['cloud_user'] is None:
                        result['cloud_user'] = cand
                        result['cloud_model'] = mp_cloud
                        log_message("resolve_paths | cloud no reconocido pero guardado para reintento")
            except Exception as ex:
                log_message("resolve_paths | cloud conversion failed | {0}".format(ex))

    return result


def set_parameter(element, bip, value):
    if element is None or bip is None:
        return False
    try:
        param = element.get_Parameter(bip)
    except Exception:
        param = None
    if param and not param.IsReadOnly:
        try:
            param.Set(value)
            return True
        except Exception:
            try:
                if isinstance(value, str):
                    param.SetValueString(value)
                    return True
            except Exception:
                pass
    return False


def set_named_parameter(element, names, value):
    if element is None:
        return False
    targets = [n.strip().lower() for n in names if n]
    try:
        iterator = element.Parameters
    except Exception:
        iterator = []
    for param in iterator:
        try:
            definition = param.Definition
        except Exception:
            definition = None
        if definition is None:
            continue
        try:
            name = (definition.Name or '').strip().lower()
        except Exception:
            name = ''
        if name in targets and not param.IsReadOnly:
            try:
                param.Set(value)
                return True
            except Exception:
                continue
    return False


def set_element_name(element, file_name, allow_rename=False):
    if element is None:
        return False
    bip_candidates = [
        getattr(BuiltInParameter, 'RVT_LINK_INSTANCE_NAME', None),
        getattr(BuiltInParameter, 'ALL_MODEL_TYPE_NAME', None),
        getattr(BuiltInParameter, 'ELEM_NAME_PARAM', None),
        getattr(BuiltInParameter, 'SYMBOL_NAME_PARAM', None),
    ]
    for bip in bip_candidates:
        if bip and set_parameter(element, bip, file_name):
            return True
    if set_named_parameter(element, ['Name', 'Nombre'], file_name):
        return True
    if allow_rename:
        try:
            element.Name = file_name
            return True
        except Exception:
            pass
    return False


def apply_name_and_mark(element, file_name, allow_rename=False):
    if element is None:
        return
    set_parameter(element, BuiltInParameter.ALL_MODEL_MARK, file_name)
    set_element_name(element, file_name, allow_rename=allow_rename)


def apply_workset(element, workset_id):
    if element is None or workset_id is None:
        return False
    try:
        target_id = ElementId(workset_id.IntegerValue)
    except Exception:
        return False
    return set_parameter(element, BuiltInParameter.ELEM_PARTITION_PARAM, target_id)


def resolve_shared_transform(host_doc, link_doc=None):
    def get_tr(document):
        try:
            loc = document.ActiveProjectLocation
            return loc.GetTransform() if loc else None
        except Exception:
            return None

    host_tr = get_tr(host_doc)
    link_tr = get_tr(link_doc) if link_doc else None
    if host_tr and link_tr:
        try:
            inv_host = host_tr.Inverse
            if inv_host:
                return inv_host.Multiply(link_tr)
        except Exception:
            pass
    return host_tr or Transform.Identity


def find_revit_link_type(doc, file_path):
    target = _normalize_user_path(os.path.abspath(file_path))
    cloud_path = build_cloud_user_visible_path(file_path)
    if cloud_path:
        target_cloud = _normalize_user_path(cloud_path)
    else:
        target_cloud = None
    collector = FilteredElementCollector(doc).OfClass(RevitLinkType)
    for link_type in collector:
        try:
            ext_ref = link_type.GetExternalFileReference()
        except Exception:
            ext_ref = None
        if ext_ref is None:
            continue
        try:
            stored = ModelPathUtils.ConvertModelPathToUserVisiblePath(ext_ref.GetAbsolutePath())
        except Exception:
            stored = None
        norm = _normalize_user_path(stored) if stored else None
        if norm and (norm == target or (target_cloud and norm == target_cloud)):
            return link_type
    return None


# UI pickers
def pick_link_type():
    options = LINK_TYPE_DEFINITIONS
    items = [SelectListItem(opt, opt['label']) for opt in options]
    selection = forms.SelectFromList.show(
        items,
        title='Selecciona tipo de links a cargar',
        multiselect=False,
        button_name='Continuar',
        name_attr='name',
    )
    if not selection:
        return None
    chosen = selection[0] if isinstance(selection, list) else selection
    return chosen.value


def pick_files(definition):
    dc_root = os.path.join(os.environ.get("USERPROFILE", ""), "DC", "ACCDocs")
    if os.path.isdir(dc_root):
        forms.alert("Tip: navega desde\n{0}".format(dc_root), title="Desktop Connector", exitscript=False)

    folder = forms.pick_folder(title='Selecciona carpeta con los RVT')
    if not folder:
        return []
    exts = definition['extensions']
    files = []
    for entry in os.listdir(folder):
        path = os.path.join(folder, entry)
        if os.path.isfile(path) and os.path.splitext(entry)[1].lower() in exts:
            files.append(path)
    if not files:
        forms.alert('No se encontraron archivos .rvt en la carpeta.')
        return []
    items = [SelectListItem(p, os.path.basename(p)) for p in sorted(files)]
    selected = forms.SelectFromList.show(
        items,
        title='Selecciona archivos a cargar ({0})'.format(definition['label']),
        multiselect=True,
        button_name='Usar seleccion',
        name_attr='name',
    )
    if not selected:
        return []
    return [itm.value for itm in selected]


def pick_placement(definition):
    modes = definition.get('placement_modes') or ['origin']
    items = [SelectListItem(m, PLACEMENT_LABELS.get(m, m)) for m in modes]
    selection = forms.SelectFromList.show(
        items,
        title='Selecciona tipo de insercion',
        multiselect=False,
        button_name='Continuar',
        name_attr='name',
    )
    if not selection:
        return None
    return selection[0].value if isinstance(selection, list) else selection.value


def pick_workset(doc):
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    worksets = list(collector)
    if not worksets:
        return None
    items = [SelectListItem(ws, ws.Name) for ws in sorted(worksets, key=lambda w: w.Name.lower())]
    selection = forms.SelectFromList.show(
        items,
        title='Selecciona el workset destino',
        multiselect=False,
        button_name='Usar workset',
        name_attr='name',
    )
    if not selection:
        return None
    return selection[0].value if isinstance(selection, list) else selection.value


def confirm_summary(definition, files, placement_key, workset):
    lines = [
        "Tipo de link: {0}".format(definition['label']),
        "Archivos: {0}".format(len(files)),
        "Insercion: {0}".format(PLACEMENT_LABELS.get(placement_key, placement_key)),
    ]
    if workset:
        lines.append("Workset: {0}".format(workset.Name))
    lines.append("")
    lines.append("Deseas continuar?")
    answer = forms.alert(
        "\n".join(lines),
        title='Confirmar carga de links',
        options=['Continuar', 'Cancelar'],
        exitscript=False,
    )
    return answer == 'Continuar'


# Core logic
def place_revit_link(doc, file_path, placement_key, workset):
    if not os.path.exists(file_path):
        raise IOError("El archivo local no existe o no esta descargado: {0}".format(file_path))

    workset_id = workset.Id if workset else None
    log_message("place_revit_link | src={0}".format(file_path))
    paths = resolve_paths(file_path)
    primary_model = paths.get('cloud_model') if paths.get('cloud_is_cloud') else paths.get('local_model')
    primary_user = paths.get('cloud_user') if paths.get('cloud_is_cloud') else paths.get('local_user')
    log_message("place_revit_link | user_visible usado={0}".format(primary_user))

    options = RevitLinkOptions(False)
    if workset_id:
        try:
            options.WorksetId = workset_id
            log_message("place_revit_link | workset_id aplicado={0}".format(workset_id.IntegerValue))
        except Exception:
            pass

    created_new = False
    try:
        link_type_id = RevitLinkType.Create(doc, primary_model, options)
        link_type = doc.GetElement(link_type_id)
        created_new = True
        log_message("place_revit_link | link type creado")
    except Exception as ex_create:
        log_message("place_revit_link | Create fallo, intentar reutilizar | err={0}".format(ex_create))
        link_type = find_revit_link_type(doc, file_path)
        if link_type is None:
            raise
        link_type_id = link_type.Id
        log_message("place_revit_link | link type reutilizado id={0}".format(link_type_id.IntegerValue))

    transform = Transform.Identity
    import_placement = ImportPlacement.Shared if placement_key == 'shared' else ImportPlacement.Origin
    instance = None
    applied_transform = False

    try:
        created = RevitLinkInstance.Create(doc, link_type_id, import_placement)
        instance = created if isinstance(created, RevitLinkInstance) else doc.GetElement(created)
        applied_transform = placement_key != 'shared'
        log_message("place_revit_link | instancia creada via ImportPlacement {0}".format(import_placement))
    except Exception as ex_inst1:
        log_message("place_revit_link | Create con ImportPlacement fallo | err={0}".format(ex_inst1))
        instance = None
        applied_transform = False

    if instance is None:
        try:
            created = RevitLinkInstance.Create(doc, link_type_id, transform)
            instance = created if isinstance(created, RevitLinkInstance) else doc.GetElement(created)
            applied_transform = True
            log_message("place_revit_link | instancia creada con transform identity")
        except Exception as ex_inst2:
            log_message("place_revit_link | Create con transform fallo | err={0}".format(ex_inst2))
            try:
                created = RevitLinkInstance.Create(doc, link_type_id)
                instance = created if isinstance(created, RevitLinkInstance) else doc.GetElement(created)
                log_message("place_revit_link | instancia creada sin transform")
            except Exception as ex_inst3:
                log_message("place_revit_link | Create sin transform fallo | err={0}".format(ex_inst3))
                instance = None
        if instance is None:
            for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
                if inst.GetTypeId() == link_type_id:
                    instance = inst
                    break
            if instance is None:
                raise

    if placement_key == 'shared' and instance is not None:
        try:
            doc.Regenerate()
        except Exception:
            pass
        try:
            link_doc = instance.GetLinkDocument()
        except Exception:
            link_doc = None
        shared_transform = resolve_shared_transform(doc, link_doc)
        if shared_transform:
            try:
                if getattr(instance, 'Pinned', False):
                    instance.Pinned = False
            except Exception:
                pass
            try:
                instance.SetTransform(shared_transform)
                applied_transform = True
                log_message("place_revit_link | transform shared aplicado")
            except Exception as ex_tr:
                log_message("place_revit_link | fallo al aplicar transform shared | err={0}".format(ex_tr))

    # Intentar recargar a cloud si existe una ruta cloud candidata
    cloud_model = paths.get('cloud_model')
    cloud_user = paths.get('cloud_user')
    if cloud_user and link_type:
        try:
            log_message("place_revit_link | intento LoadFrom cloud={0}".format(cloud_user))
            link_type.LoadFrom(cloud_model, options)
            log_message("place_revit_link | LoadFrom cloud OK")
        except Exception as ex_load:
            log_message("place_revit_link | LoadFrom cloud fallo | err={0}".format(ex_load))

    file_name = os.path.basename(file_path)
    if workset_id:
        apply_workset(instance, workset_id)
    apply_name_and_mark(instance, file_name, allow_rename=True)
    if link_type:
        apply_name_and_mark(link_type, file_name, allow_rename=True)

    if not applied_transform and instance is not None:
        try:
            if getattr(instance, 'Pinned', False):
                instance.Pinned = False
        except Exception:
            pass
        try:
            instance.SetTransform(transform)
        except Exception:
            pass

    return instance


def process_links(doc, definition, files, placement_key, workset):
    summaries = []
    log_message(
        "Inicio carga | files={0} | placement={1} | workset={2}".format(
            len(files), placement_key, getattr(workset, 'Name', None)
        )
    )
    transaction = Transaction(doc, 'Cargar links en bulk')
    transaction.Start()
    any_success = False
    try:
        for file_path in files:
            file_name = os.path.basename(file_path)
            try:
                log_message("Procesar archivo | {0}".format(file_path))
                place_revit_link(doc, file_path, placement_key, workset)
                summaries.append((file_name, True, 'Insertado'))
                any_success = True
                log_message("OK | {0}".format(file_name))
            except Exception as exc:
                summaries.append((file_name, False, str(exc)))
                log_message("ERROR | {0} | {1}\n{2}".format(file_name, exc, traceback.format_exc()))
        if any_success:
            transaction.Commit()
            log_message("Transaccion commit")
        else:
            transaction.RollBack()
            log_message("Transaccion rollback (sin exitos)")
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
            log_message("Transaccion rollback por excepcion general")
        raise
    return summaries


def build_result_message(summaries):
    success = [s for s in summaries if s[1]]
    failed = [s for s in summaries if not s[1]]
    lines = []
    if success:
        lines.append(u"Links insertados ({0}):".format(len(success)))
        for name, _, _ in success[:10]:
            lines.append(u"  - {0}".format(name))
        if len(success) > 10:
            lines.append(u"  ... y {0} mas.".format(len(success) - 10))
        lines.append(u"")
    if failed:
        lines.append(u"Links con errores ({0}):".format(len(failed)))
        for name, _, msg in failed:
            lines.append(u"  - {0}: {1}".format(name, msg))
    if not lines:
        lines.append('No se procesaron archivos.')
    return "\n".join(lines)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert('No hay documento activo.')
        return
    try:
        log_message("=== EJECUCION INICIADA ===")
        definition = pick_link_type()
        if not definition:
            return
        files = pick_files(definition)
        if not files:
            return
        placement_key = pick_placement(definition)
        if not placement_key:
            return
        workset = pick_workset(doc)
        if workset is None:
            answer = forms.alert(
                'No seleccionaste un workset. Deseas continuar usando el actual?',
                title='Workset',
                options=['Continuar', 'Cancelar'],
                exitscript=False,
            )
            if answer != 'Continuar':
                return
        if not confirm_summary(definition, files, placement_key, workset):
            return
        summaries = process_links(doc, definition, files, placement_key, workset)
        forms.alert(build_result_message(summaries), title='Resultado carga links')
    except Exception as exc:
        log_message("FALLO GENERAL: {0}\n{1}".format(exc, traceback.format_exc()))
        forms.alert(
            u"Error al cargar links:\n{0}\n\nTrace:\n{1}".format(exc, traceback.format_exc()),
            title='Cargar links',
        )
    finally:
        log_message("=== EJECUCION FINALIZADA ===")


if __name__ == '__main__':
    main()
