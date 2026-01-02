# -*- coding: utf-8 -*-
"""Placeholder: informa que la herramienta no esta disponible."""

from pyrevit import forms


TITLE = "Corregir Links en Bulk"


def main():
    forms.alert(
        "Esta herramienta aun no tiene implementacion.\n"
        "Si la necesitas, avisa al equipo BIM para priorizarla.",
        title=TITLE,
        warn_icon=True,
    )


if __name__ == "__main__":
    main()
