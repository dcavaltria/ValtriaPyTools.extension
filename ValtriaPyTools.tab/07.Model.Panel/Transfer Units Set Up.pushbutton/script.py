# -*- coding: utf-8 -*-
"""Transfiere la configuracion de unidades entre proyectos abiertos."""

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import Transaction, Units  # type: ignore
from pyrevit import forms

from valtria_lib import get_app, get_doc, log_exception, log_to_file
from valtria_core.text import ensure_text as safe_text


TITLE = "Transfer Units Set Up"
LOG_TOOL = "transfer_units_setup"


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


class DocOption(object):
    def __init__(self, doc):
        self.value = doc
        self.label = doc_label(doc)

    @property
    def name(self):
        return self.label


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


def clone_units(source_units):
    if source_units is None:
        return None
    try:
        clone = source_units.Clone()
        if clone is not None:
            return clone
    except Exception:
        pass
    try:
        return Units(source_units)
    except Exception:
        pass
    return source_units


def main():
    log_line("----")
    log_line("Inicio transferencia de unidades")
    app = get_app()
    active_doc = get_doc()
    open_docs = list_open_documents(app)
    if len(open_docs) < 2:
        forms.alert(
            u"Abre al menos dos proyectos Revit (.rvt) para copiar las unidades.",
            title=TITLE,
            warn_icon=True,
        )
        log_line("Abortado: menos de dos documentos abiertos")
        return

    source_doc = pick_document(open_docs, u"Selecciona proyecto origen", default_doc=active_doc)
    if source_doc is None:
        log_line("Cancelado por usuario al elegir origen")
        return

    dest_doc = pick_document(open_docs, u"Selecciona proyecto destino", exclude=source_doc)
    if dest_doc is None:
        log_line("Cancelado por usuario al elegir destino")
        return

    if source_doc == dest_doc:
        forms.alert(u"El proyecto destino no puede ser igual al origen.", title=TITLE, warn_icon=True)
        log_line("Abortado: origen y destino iguales")
        return

    try:
        if getattr(dest_doc, "IsReadOnly", False):
            forms.alert(u"El proyecto destino esta en modo lectura y no puede modificarse.", title=TITLE, warn_icon=True)
            log_line("Abortado: destino en solo lectura")
            return
    except Exception:
        pass

    confirm = forms.alert(
        u"Se copiaran las unidades desde:\n{0}\n\nHacia:\n{1}\n\nContinuar?".format(
            doc_label(source_doc), doc_label(dest_doc)
        ),
        title=TITLE,
        yes=True,
        no=True,
        warn_icon=False,
    )
    if not confirm:
        log_line("Cancelado por usuario en confirmacion")
        return

    try:
        source_units = source_doc.GetUnits()
    except Exception as err:
        log_line("No se pudieron leer las unidades del origen: {0}".format(safe_text(err)))
        forms.alert(u"No se pudieron leer las unidades del proyecto origen.", title=TITLE, warn_icon=True)
        return

    units_to_apply = clone_units(source_units)
    if units_to_apply is None:
        forms.alert(u"No se pudo preparar la configuracion de unidades para transferir.", title=TITLE, warn_icon=True)
        log_line("Abortado: unidades de origen vacias")
        return

    tx = Transaction(dest_doc, TITLE)
    tx.Start()
    try:
        dest_doc.SetUnits(units_to_apply)
        tx.Commit()
    except Exception as err:
        try:
            tx.RollBack()
        except Exception:
            pass
        log_exception(err, title=TITLE)
        forms.alert(u"No se pudo aplicar la configuracion de unidades en destino.", title=TITLE, warn_icon=True)
        return

    forms.alert(u"Unidades transferidas correctamente.", title=TITLE, warn_icon=False)
    log_line("Transferencia completada: {0} -> {1}".format(doc_label(source_doc), doc_label(dest_doc)))


if __name__ == "__main__":
    try:
        main()
    except Exception as main_error:
        try:
            log_line("Error inesperado: {0}".format(safe_text(main_error)))
        except Exception:
            pass
        log_exception(main_error, title=TITLE)
