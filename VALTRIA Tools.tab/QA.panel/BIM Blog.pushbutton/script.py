# -*- coding: utf-8 -*-
"""Abrir el artículo de BIM para salas limpias en el sitio de VALTRIA."""
from __future__ import print_function

import webbrowser

from pyrevit import forms, script

URL = "https://valtria.com/en/blog/bim-para-salas-limpias/"

logger = script.get_logger()

try:
    webbrowser.open(URL, new=2)
    logger.info("Abriendo el artículo: %s", URL)
except Exception as error:  # pragma: no cover - entorno de Revit
    logger.error("No se pudo abrir el navegador: %s", error)
    forms.alert(
        "No se pudo abrir el navegador.\n\n{0}".format(error),
        title="VALTRIA Tools",
        exitscript=True,
    )
