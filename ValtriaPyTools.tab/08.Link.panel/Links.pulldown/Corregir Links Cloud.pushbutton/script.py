# -*- coding: utf-8 -*-
"""Corrige links RVT cargados con ruta absoluta de Desktop Connector a ruta cloud."""

import os
import traceback
from datetime import datetime

import clr

try:
    clr.AddReference('RevitAPI')
    clr.AddReference('RevitAPIUI')
except Exception:
    pass

from Autodesk.Revit.DB import (  # noqa
    FilteredElementCollector,
    ModelPathUtils,
    RevitLinkOptions,
    RevitLinkType,
    Transaction,
)

from pyrevit import forms, revit

DESKTOP_CONNECTOR_ROOT_KEYS = (
    'autodeskdocs',
    'accdocs',
    'dcaccdocs',
    'bim360docs',
    'bim360',
)

LOG_FILE = os.path.join(os.path.dirname(__file__), 'corregir_links_cloud.log')


def log_message(text):
    """Append a timestamped line to the local log; ignore errors to not block UI."""
    try:
        stamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(LOG_FILE, 'a') as fh:
            fh.write("[{0}] {1}\n".format(stamp, text))
    except Exception:
        pass


def _normalize_user_path(path):
    if not path:
        return None
    return os.path.normpath(path).replace('\\', '/').lower().rstrip('/')


def _desktop_connector_parts(file_path):
    """Devuelve (idx, parts, root_type) donde root_type es 'autodesk' o 'bim360'."""
    parts = [p for p in os.path.abspath(file_path).replace('\\', '/').split('/') if p]
    root_type = None

    for idx, part in enumerate(parts):
        key = part.replace(' ', '').lower()

        # Autodesk Docs / ACC
        if 'autodeskdocs' in key or 'accdocs' in key or 'dcaccdocs' in key:
            root_type = 'autodesk'
            return idx, parts, root_type

        # BIM 360 Docs
        if 'bim360' in key:
            root_type = 'bim360'
            return idx, parts, root_type

    return None, parts, None


def build_cloud_user_visible_path(file_path):
    """Convierte una ruta de Desktop Connector en un user-visible path de Autodesk Docs/BIM 360."""
    idx, parts, root_type = _desktop_connector_parts(file_path)
    if idx is None or root_type is None:
        return None

    remainder = parts[idx + 1 :]
    if len(remainder) < 2:
        # Aseguramos al menos <account>/<project>/...
        return None

    base = '/'.join(remainder)

    if root_type == 'autodesk':
        prefix = 'Autodesk Docs://'
    else:
        prefix = 'BIM 360 Docs://'

    cloud_path = prefix + base
    log_message("build cloud | root={0} | base={1} | path={2}".format(root_type, base, cloud_path))
    return cloud_path


def is_cloud_path(model_path):
    try:
        return ModelPathUtils.IsCloudPath(model_path)
    except Exception:
        pass
    try:
        return 'Cloud' in (model_path.GetType().Name or '')
    except Exception:
        return False


def resolve_model_path(file_path):
    """Return a ModelPath, prefiriendo cloud si viene de Desktop Connector."""
    cloud_user_path = build_cloud_user_visible_path(file_path)
    log_message("convert model path | src={0} | cloud_user_path={1}".format(file_path, cloud_user_path))

    if cloud_user_path:
        try:
            mp_cloud = ModelPathUtils.ConvertUserVisiblePathToModelPath(cloud_user_path)
            try:
                is_cloud = ModelPathUtils.IsCloudPath(mp_cloud)
            except Exception:
                is_cloud = 'Cloud' in (mp_cloud.GetType().Name or '')

            if is_cloud:
                try:
                    uv = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp_cloud)
                except Exception:
                    uv = None
                log_message("cloud conversion ok | user_visible={0}".format(uv))
                return mp_cloud, cloud_user_path
        except Exception as ex:
            log_message("cloud conversion failed | path={0} | err={1}".format(cloud_user_path, ex))

    # Fallback: local/absoluto
    mp_local = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
    try:
        user_visible = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp_local)
    except Exception:
        user_visible = None

    log_message(
        "local conversion | user_visible={0} | exists_local={1}".format(
            user_visible, os.path.exists(user_visible or '')
        )
    )
    return mp_local, os.path.abspath(file_path)


def find_revit_link_type(doc, file_path):
    """Busca si ya existe un RevitLinkType que apunte a esa ruta (local o cloud)."""
    candidates = {_normalize_user_path(os.path.abspath(file_path))}

    cloud_user_path = build_cloud_user_visible_path(file_path)
    if cloud_user_path:
        candidates.add(_normalize_user_path(cloud_user_path))

    collector = FilteredElementCollector(doc).OfClass(RevitLinkType)

    for link_type in collector:
        try:
            ext_ref = link_type.GetExternalFileReference()
        except Exception:
            ext_ref = None

        if ext_ref is None:
            continue

        try:
            stored_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(ext_ref.GetAbsolutePath())
        except Exception:
            stored_path = None

        if stored_path and _normalize_user_path(stored_path) in candidates:
            return link_type

    return None


def gather_absolute_links(doc):
    """Return list of (RevitLinkType, stored_user_path, cloud_candidate)."""
    results = []
    collector = FilteredElementCollector(doc).OfClass(RevitLinkType)
    for link_type in collector:
        try:
            ext_ref = link_type.GetExternalFileReference()
        except Exception:
            ext_ref = None
        if ext_ref is None:
            continue
        try:
            stored_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(ext_ref.GetAbsolutePath())
        except Exception:
            stored_path = None
        if not stored_path:
            continue
        # Saltar si ya es cloud
        try:
            if ModelPathUtils.IsCloudPath(ext_ref.GetAbsolutePath()):
                continue
        except Exception:
            pass
        cloud_candidate = build_cloud_user_visible_path(stored_path)
        log_message(
            "found link | name={0} | stored={1} | cloud_candidate={2}".format(
                getattr(link_type, 'Name', '<sin nombre>'), stored_path, cloud_candidate
            )
        )
        results.append((link_type, stored_path, cloud_candidate))
    return results


def relink_to_cloud(doc, targets):
    """Reload link types using cloud paths. targets: list of tuples (link_type, cloud_path)."""
    success = []
    failed = []
    t = Transaction(doc, 'Corregir links a cloud')
    t.Start()
    try:
        for link_type, cloud_path in targets:
            name = getattr(link_type, 'Name', '<sin nombre>')
            try:
                log_message("relink attempt | name={0} | cloud_path={1}".format(name, cloud_path))
                mp_cloud, user_path_used = resolve_model_path(cloud_path)
                try:
                    visible = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp_cloud)
                except Exception:
                    visible = None
                cloud_flag = is_cloud_path(mp_cloud)
                log_message(
                    "modelpath visible={0} | is_cloud={1} | user_path_used={2}".format(
                        visible, cloud_flag, user_path_used
                    )
                )
                options = RevitLinkOptions(False)
                load_method = getattr(link_type, 'LoadFrom', None)
                if callable(load_method):
                    load_method(mp_cloud, options)
                else:
                    link_type.Reload()
                success.append(name)
                log_message("OK relink {0} -> {1}".format(name, cloud_path))
            except Exception as ex:
                failed.append((name, str(ex)))
                log_message("ERROR relink {0} -> {1}: {2}\n{3}".format(name, cloud_path, ex, traceback.format_exc()))
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        raise
    return success, failed


def build_summary(candidates):
    lines = []
    for idx, (lt, stored, cloud) in enumerate(candidates, 1):
        lines.append("{0}. {1}".format(idx, getattr(lt, 'Name', '<sin nombre>')))
        lines.append("    actual: {0}".format(stored))
        lines.append("    cloud?: {0}".format(cloud or "no deducible"))
    return "\n".join(lines)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert('No hay documento activo.')
        return
    try:
        candidates = gather_absolute_links(doc)
        candidates = [(lt, stored, cloud) for lt, stored, cloud in candidates if cloud]
        if not candidates:
            forms.alert('No se encontraron links con ruta absoluta de Desktop Connector para corregir.')
            return

        msg_lines = [
            "Se encontraron {0} links con ruta absoluta de Desktop Connector.".format(len(candidates)),
            "Se intentara convertirlas a rutas cloud.",
            "",
            build_summary(candidates[:10]),
        ]
        if len(candidates) > 10:
            msg_lines.append("... y {0} mas.".format(len(candidates) - 10))
        msg_lines.append("")
        msg_lines.append("Quieres continuar?")

        answer = forms.alert(
            "\n".join(msg_lines),
            title='Corregir rutas a cloud',
            options=['Continuar', 'Cancelar'],
            exitscript=False,
        )
        if answer != 'Continuar':
            return

        targets = []
        for lt, stored, cloud in candidates:
            name = getattr(lt, 'Name', '<sin nombre>')
            prompt_lines = [
                "Link: {0}".format(name),
                "Ruta actual: {0}".format(stored),
                "Ruta cloud propuesta: {0}".format(cloud or "no deducible"),
                "",
                "Selecciona el archivo correcto en el explorador (ACC/Desktop Connector).",
                "Si cancelas, se usara la propuesta.",
            ]
            forms.alert("\n".join(prompt_lines), title='Selecciona archivo para ruta cloud', exitscript=False)
            # Selector simple: solo filtra por *.rvt para evitar errores de cadena de filtro
            picked = forms.pick_file(file_ext='*.rvt')
            final_cloud = cloud
            if picked:
                manual_cloud = build_cloud_user_visible_path(picked) or picked
                final_cloud = manual_cloud
                log_message(
                    "ruta elegida por explorador | name={0} | picked={1} | final_cloud={2}".format(
                        name, picked, final_cloud
                    )
                )
            else:
                log_message("ruta propuesta usada | name={0} | cloud={1}".format(name, final_cloud))
            targets.append((lt, final_cloud))

        success, failed = relink_to_cloud(doc, targets)

        result_lines = []
        if success:
            result_lines.append("Links corregidos ({0}):".format(len(success)))
            for name in success[:10]:
                result_lines.append("  - {0}".format(name))
            if len(success) > 10:
                result_lines.append("  ... y {0} mas.".format(len(success) - 10))
            result_lines.append("")
        if failed:
            result_lines.append("Links con error ({0}):".format(len(failed)))
            for name, err in failed:
                result_lines.append("  - {0}: {1}".format(name, err))
        if not result_lines:
            result_lines.append("No se procesaron links.")

        forms.alert("\n".join(result_lines), title='Corregir rutas a cloud')
    except Exception as exc:
        log_message("FALLO GENERAL: {0}\n{1}".format(exc, traceback.format_exc()))
        forms.alert("Error:\n{0}\n\nTrace:\n{1}".format(exc, traceback.format_exc()), title='Corregir rutas a cloud')


if __name__ == '__main__':
    main()
