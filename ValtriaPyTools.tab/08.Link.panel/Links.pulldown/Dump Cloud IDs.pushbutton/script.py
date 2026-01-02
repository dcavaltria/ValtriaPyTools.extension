# -*- coding: utf-8 -*-
"""List existing cloud link IDs (hub/project/model) to help build ExternalResourceReference."""

import os
import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkType,
    ModelPathUtils,
)

from pyrevit import forms, revit

LOG_FILE = os.path.join(os.path.dirname(__file__), 'dump_cloud_ids.log')


def log(line):
    try:
        with open(LOG_FILE, 'a') as fh:
            fh.write(line + '\n')
    except Exception:
        pass


def main():
    doc = revit.doc
    if doc is None:
        forms.alert('No hay documento activo.')
        return

    lines = []
    total = 0
    for lt in FilteredElementCollector(doc).OfClass(RevitLinkType):
        total += 1
        try:
            ext = lt.GetExternalFileReference()
        except Exception:
            ext = None
        if not ext:
            continue

        try:
            mp = ext.GetAbsolutePath()
        except Exception as ex_path:
            log("skip {0} | error GetAbsolutePath: {1}".format(getattr(lt, 'Name', '<sin nombre>'), ex_path))
            continue

        # Intentar user-visible para referencia
        try:
            uv = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
        except Exception:
            uv = None

        name = getattr(lt, 'Name', '<sin nombre>')
        is_cloud = False
        try:
            is_cloud = ModelPathUtils.IsCloudPath(mp)
        except Exception as ex_cloudflag:
            log("fail IsCloudPath {0} | err={1} | uv={2}".format(name, ex_cloudflag, uv))

        hub = proj = model = None
        try:
            # Algunos ExternalFileReference devuelven info adicional
            info = ext.GetReferenceInformation()
            if info:
                for k in info.Keys:
                    log(u"info {0} | {1}={2}".format(name, k, info[k]))
                hub = info.get("HubId", None) or hub
                proj = info.get("ProjectId", None) or proj
                model = info.get("ModelId", None) or model
        except Exception as ex_info:
            log("fail GetReferenceInformation {0} | err={1}".format(name, ex_info))

        if is_cloud:
            line = u"{0} | cloud=True | uv={1}".format(name, uv)
            if hub or proj or model:
                line = u"{0} | hub={1} | project={2} | model={3}".format(line, hub, proj, model)
            lines.append(line)
            log(line)
        else:
            log("not cloud {0} | uv={1}".format(name, uv))

    if not lines:
        forms.alert("No se encontraron links cloud en el modelo.\nRevisa dump_cloud_ids.log para detalles.", title='Dump Cloud IDs')
        return

    preview = "\n".join(lines[:10])
    if len(lines) > 10:
        preview += u"\n... y {0} más (ver {1})".format(len(lines) - 10, LOG_FILE)
    forms.alert(preview, title='Dump Cloud IDs', exitscript=False)


if __name__ == '__main__':
    main()
