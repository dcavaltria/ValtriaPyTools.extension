# -*- coding: utf-8 -*-
"""Transfiere configuraciones de Project Browser entre proyectos abiertos."""

import codecs
import datetime
import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from System.Collections.Generic import List  # type: ignore

from Autodesk.Revit.DB import (  # type: ignore
    BrowserOrganization,
    BrowserOrganizationType,
    BuiltInParameter,
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    Transaction,
    Transform,
)
from pyrevit import forms

from valtria_lib import get_app, get_doc, log_exception, log_to_file
from valtria_core.text import ensure_text as safe_text


TITLE = "Transfer Project Browser"
LOG_TOOL = "transfer_project_browser"
LOG_DIR = os.path.join(EXTENSION_DIR, "_logs")


def log_line(message):
    log_to_file(LOG_TOOL, message)


def iterate(net_collection):
    """Itera sobre colecciones .NET de forma segura."""
    if net_collection is None:
        return
    try:
        for item in net_collection:
            yield item
        return
    except Exception:
        pass
    try:
        iterator = net_collection.GetEnumerator()
    except Exception:
        iterator = None
    if iterator is None:
        return
    while iterator.MoveNext():
        yield iterator.Current


def doc_label(doc):
    title = safe_text(getattr(doc, "Title", u"")).strip() or u"(Sin titulo)"
    path = safe_text(getattr(doc, "PathName", u"")).strip()
    if not path:
        path = u"(No guardado)"
    return u"{0} [{1}]".format(title, path)


def org_name(org):
    """Devuelve el nombre visible del esquema de browser."""
    candidates = []
    try:
        candidates.append(getattr(org, "Name", None))
    except Exception:
        pass
    try:
        p = org.get_Parameter(BuiltInParameter.ELEM_NAME_PARAM)
        if p:
            candidates.append(p.AsString())
    except Exception:
        pass
    try:
        p = org.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p:
            candidates.append(p.AsString())
    except Exception:
        pass
    for cand in candidates:
        text = safe_text(cand).strip()
        if text:
            return text
    return u""


def list_open_documents(app):
    docs = []
    for doc in iterate(getattr(app, "Documents", None)):
        if doc is None:
            continue
        try:
            if getattr(doc, "IsFamilyDocument", False):
                continue
        except Exception:
            pass
        docs.append(doc)
    return docs


def collect_browser_organizations(doc):
    collector = FilteredElementCollector(doc).OfClass(BrowserOrganization)
    items = []
    for org in collector:
        if org is None:
            continue
        items.append(org)
    return items


def org_type_value(org_type):
    try:
        return int(org_type)
    except Exception:
        try:
            return int(getattr(org_type, "value__", None))
        except Exception:
            return safe_text(org_type)


def org_type_label(org_type):
    try:
        if org_type == BrowserOrganizationType.Views:
            return u"Vistas"
        if org_type == BrowserOrganizationType.Sheets:
            return u"Hojas"
        if org_type == BrowserOrganizationType.Schedules:
            return u"Tablas"
    except Exception:
        pass
    return safe_text(org_type) or u"? "


def org_type_from_element(org):
    if org is None:
        return None
    try:
        tval = getattr(org, "Type", None)
    except Exception:
        tval = None
    if tval is not None:
        return tval
    try:
        return getattr(org, "BrowserOrganizationType", None)
    except Exception:
        return None


def org_key(org):
    name = org_name(org).lower()
    tval = org_type_value(org_type_from_element(org))
    return (tval, name)


class DocOption(object):
    def __init__(self, doc):
        self.value = doc
        self.label = doc_label(doc)

    @property
    def name(self):
        return self.label


class OrgOption(object):
    def __init__(self, org):
        self.value = org
        self.name_value = org_name(org) or u"(sin nombre)"
        self.type_value = org_type_from_element(org)
        self.type_label = org_type_label(self.type_value)

    @property
    def name(self):
        return u"{0} | {1}".format(self.type_label, self.name_value)


def pick_document(documents, title, default_doc=None, exclude=None):
    options = []
    for doc in documents:
        if exclude is not None and doc == exclude:
            continue
        options.append(DocOption(doc))
    if not options:
        return None
    default_option = None
    if default_doc is not None:
        for opt in options:
            if opt.value == default_doc:
                default_option = opt
                break
    picked = forms.SelectFromList.show(
        options,
        title=title,
        multiselect=False,
        button_name="Seleccionar",
        name_attr="name",
        default=default_option,
    )
    if not picked:
        return None
    return picked.value if hasattr(picked, "value") else picked


def pick_orgs(orgs):
    ordered = sorted(orgs, key=lambda o: (org_type_label(org_type_from_element(o)), safe_text(getattr(o, "Name", u"")).lower()))
    items = [OrgOption(o) for o in ordered]
    picked = forms.SelectFromList.show(
        items,
        title=u"Selecciona las configuraciones a transferir",
        multiselect=True,
        button_name="Transferir",
        name_attr="name",
    )
    if not picked:
        return []
    if isinstance(picked, (list, tuple)):
        return [p.value if hasattr(p, "value") else p for p in picked]
    return [picked.value if hasattr(picked, "value") else picked]


def ensure_log_dir():
    if not os.path.isdir(LOG_DIR):
        os.makedirs(LOG_DIR)


def write_report(origin_doc, dest_doc, rows):
    ensure_log_dir()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = "transfer_project_browser_{0}.txt".format(timestamp)
    path = os.path.join(LOG_DIR, filename)
    origin_label = doc_label(origin_doc) if origin_doc else u"(desconocido)"
    dest_label = doc_label(dest_doc) if dest_doc else u"(desconocido)"

    lines = [
        u"Transferencia de Project Browser",
        u"Fecha: {0}".format(timestamp),
        u"Origen: {0}".format(origin_label),
        u"Destino: {0}".format(dest_label),
        u"",
        u"Resultado:",
    ]
    for row in rows:
        status = row.get("status", u"?")
        name = row.get("name", u"?")
        typ = row.get("type", u"?")
        detail = row.get("detail", u"")
        base = u"- [{0}] {1} | {2}".format(status, typ, name)
        if detail:
            base += u" -- {0}".format(detail)
        lines.append(base)
    with codecs.open(path, "w", "utf-8") as fh:
        fh.write(u"\n".join(lines))
    return path


def main():
    log_line("----")
    log_line("Inicio transferencia Project Browser")
    app = get_app()
    active_doc = get_doc()
    open_docs = list_open_documents(app)
    if len(open_docs) < 2:
        forms.alert(u"Abre al menos dos proyectos Revit (.rvt) para copiar configuraciones de Project Browser.", title=TITLE, warn_icon=True)
        log_line("Abortado: menos de dos documentos abiertos")
        return

    source_doc = pick_document(open_docs, u"Selecciona proyecto origen", default_doc=active_doc)
    if source_doc is None:
        log_line("Cancelado por usuario al elegir origen")
        return

    available_orgs = collect_browser_organizations(source_doc)
    if not available_orgs:
        forms.alert(u"El proyecto origen no tiene configuraciones de Project Browser para copiar.", title=TITLE, warn_icon=True)
        log_line("Abortado: sin configuraciones en origen")
        return

    selected_orgs = pick_orgs(available_orgs)
    if not selected_orgs:
        log_line("Cancelado por usuario al seleccionar configuraciones")
        return

    dest_doc = pick_document(open_docs, u"Selecciona proyecto destino", exclude=source_doc)
    if dest_doc is None:
        log_line("Cancelado por usuario al elegir destino")
        return

    dest_existing = {}
    for org in collect_browser_organizations(dest_doc):
        dest_existing[org_key(org)] = org

    to_copy = []
    report_rows = []
    for org in selected_orgs:
        name = org_name(org) or u"(sin nombre)"
        tlabel = org_type_label(org_type_from_element(org))
        key = org_key(org)
        if key in dest_existing:
            report_rows.append({"status": u"omitido", "name": name, "type": tlabel, "detail": u"Ya existe en destino"})
            continue
        to_copy.append(org)

    if not to_copy:
        forms.alert(u"Todas las configuraciones seleccionadas ya existen en el proyecto destino.", title=TITLE, warn_icon=True)
        log_line("Nada que copiar: todas existen en destino")
        return

    confirm = forms.alert(
        u"Se copiaran {0} configuraciones desde \"{1}\" hacia \"{2}\".\nContinuar?".format(
            len(to_copy), doc_label(source_doc), doc_label(dest_doc)
        ),
        title=TITLE,
        yes=True,
        no=True,
        warn_icon=False,
    )
    if not confirm:
        log_line("Cancelado por usuario en confirmacion")
        return

    tx = Transaction(dest_doc, TITLE)
    tx.Start()
    options = CopyPasteOptions()
    for org in to_copy:
        name = org_name(org) or u"(sin nombre)"
        tlabel = org_type_label(org_type_from_element(org))
        try:
            ids = List[ElementId]()
            ids.Add(org.Id)
            copied_ids = ElementTransformUtils.CopyElements(
                source_doc,
                ids,
                dest_doc,
                Transform.Identity,
                options,
            )
            new_ids = list(copied_ids) if copied_ids is not None else []
            if new_ids:
                report_rows.append({"status": u"copiado", "name": name, "type": tlabel, "detail": u""})
            else:
                report_rows.append({"status": u"error", "name": name, "type": tlabel, "detail": u"No se devolvieron IDs"})
        except Exception as err:
            report_rows.append({"status": u"error", "name": name, "type": tlabel, "detail": safe_text(err)})
            log_line("Fallo al copiar {0} ({1}): {2}".format(name, tlabel, safe_text(err)))
    try:
        tx.Commit()
    except Exception as err:
        tx.RollBack()
        log_exception(err, title=TITLE)
        forms.alert(u"No se pudo completar la transferencia:\n{0}".format(safe_text(err)), title=TITLE, warn_icon=True)
        return

    report_path = write_report(source_doc, dest_doc, report_rows)
    copied_count = len([r for r in report_rows if r.get("status") == u"copiado"])
    skipped_count = len([r for r in report_rows if r.get("status") == u"omitido"])
    error_count = len([r for r in report_rows if r.get("status") == u"error"])

    summary_lines = [
        u"Configuraciones copiadas: {0}".format(copied_count),
        u"Omitidas (ya existian): {0}".format(skipped_count),
        u"Errores: {0}".format(error_count),
        u"Reporte: {0}".format(report_path),
    ]
    forms.alert(u"\n".join(summary_lines), title=TITLE, warn_icon=False)
    log_line("Fin transferencia. Copiadas={0}, Omitidas={1}, Errores={2}, Reporte={3}".format(
        copied_count, skipped_count, error_count, report_path
    ))


if __name__ == "__main__":
    try:
        main()
    except Exception as main_error:
        try:
            log_line("Error inesperado: {0}".format(safe_text(main_error)))
        except Exception:
            pass
        log_exception(main_error, title=TITLE)

