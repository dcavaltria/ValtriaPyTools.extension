# -*- coding: utf-8 -*-
"""
Script para crear segmentos de tuberia Cobre A en Revit.
Basado en normativa ASTM B88 - Tipo L.
"""

__title__ = "Crear\nCobre A"
__author__ = "MEP Team"

import sys

from Autodesk.Revit.DB import ElementId, FilteredElementCollector, Transaction
from Autodesk.Revit.DB.Plumbing import PipeScheduleType
from pyrevit import forms, revit

doc = revit.doc

# Datos de tuberia Cobre A (Tipo L - Estandar)
# Formato: (Nominal_mm, ID_mm, OD_mm)
COPPER_A_DATA = [
    ("6.000 mm", 8.00, 9.52),      # 1/4"
    ("8.000 mm", 10.67, 12.70),    # 3/8"
    ("10.000 mm", 13.84, 15.88),   # 1/2"
    ("15.000 mm", 17.27, 19.05),   # 5/8"
    ("20.000 mm", 20.19, 22.22),   # 3/4"
    ("25.000 mm", 26.42, 28.58),   # 1"
    ("32.000 mm", 32.89, 34.92),   # 1 1/4"
    ("40.000 mm", 38.99, 41.28),   # 1 1/2"
    ("50.000 mm", 51.05, 53.98),   # 2"
    ("65.000 mm", 63.50, 66.68),   # 2 1/2"
    ("80.000 mm", 76.20, 79.38),   # 3"
    ("90.000 mm", 88.90, 92.08),   # 3 1/2"
]


def mm_to_feet(mm_value):
    """Convierte milimetros a pies (unidad interna de Revit)."""
    return mm_value / 304.8


def get_or_create_pipe_schedule_type(document, name):
    """Obtiene un PipeScheduleType existente o crea uno nuevo."""
    existing_id = PipeScheduleType.GetPipeScheduleId(document, name)
    if existing_id and existing_id != ElementId.InvalidElementId:
        print("Encontrado PipeScheduleType existente '{}'.".format(name))
        return document.GetElement(existing_id)

    schedule_types = list(
        FilteredElementCollector(document)
        .OfClass(PipeScheduleType)
        .WhereElementIsElementType()
    )

    for schedule_type in schedule_types:
        if schedule_type.Name == name:
            return schedule_type

    try:
        new_schedule = PipeScheduleType.Create(document, name)
        if new_schedule:
            print("Creado nuevo PipeScheduleType via API.")
            return new_schedule
    except Exception as err:
        print("No se pudo crear el PipeScheduleType directamente: {}".format(err))

    if schedule_types:
        try:
            duplicate_id = schedule_types[0].Duplicate(name)
            duplicated = document.GetElement(duplicate_id)
            print("Duplicado PipeScheduleType existente.")
            return duplicated
        except Exception as err:
            print("No se pudo duplicar PipeScheduleType existente: {}".format(err))

    return None


def create_pipe_sizes(schedule_type, pipe_data):
    """Actualiza los tamanos de tuberia en el schedule type indicado."""
    try:
        existing_sizes = list(schedule_type.GetSizes())
        for size in existing_sizes:
            schedule_type.RemoveSize(size)
    except Exception:
        pass

    created_count = 0
    for nominal_str, inner_dia, outer_dia in pipe_data:
        try:
            nominal_value = float(nominal_str.split()[0])

            nominal_ft = mm_to_feet(nominal_value)
            id_ft = mm_to_feet(inner_dia)
            od_ft = mm_to_feet(outer_dia)

            schedule_type.AddSize(nominal_ft, id_ft, od_ft)
            created_count += 1

            print(
                "Creado: {} - ID: {} mm, OD: {} mm".format(
                    nominal_str, inner_dia, outer_dia
                )
            )
        except Exception as err:
            print("Error en {}: {}".format(nominal_str, err))
            continue

    return created_count


def main():
    """Funcion principal del script."""
    if doc.IsFamilyDocument:
        forms.alert(
            "Este script solo funciona en proyectos.",
            title="Error",
            exitscript=True,
        )

    schedule_name = "Copper - A"
    confirm = forms.alert(
        "Se crearan {} segmentos de tuberia para:\n\n'{}'\n\nContinuar?".format(
            len(COPPER_A_DATA), schedule_name
        ),
        title="Crear Segmentos Cobre A",
        ok=True,
        cancel=True,
    )

    if not confirm:
        sys.exit()

    transaction = Transaction(doc, "Crear Segmentos Cobre A")
    transaction.Start()

    try:
        print("\n" + "=" * 50)
        print("Iniciando creacion de segmentos Cobre A...")
        print("=" * 50 + "\n")

        schedule_type = get_or_create_pipe_schedule_type(doc, schedule_name)

        if not schedule_type:
            transaction.RollBack()
            forms.alert(
                "No se pudo crear el Schedule Type",
                title="Error",
                exitscript=True,
            )

        print("Schedule Type '{}' listo\n".format(schedule_name))

        created_count = create_pipe_sizes(schedule_type, COPPER_A_DATA)

        transaction.Commit()

        print("\n" + "=" * 50)
        print("Proceso completado")
        print("=" * 50)

        result_lines = [
            "Segmentos creados exitosamente.",
            "Schedule Type: {}".format(schedule_name),
            "Total de segmentos: {}".format(created_count),
            "",
            "Tamanos creados:",
        ]

        for nominal, inner_dia, outer_dia in COPPER_A_DATA:
            result_lines.append(
                "- {} -> ID: {} mm / OD: {} mm".format(nominal, inner_dia, outer_dia)
            )

        result_lines.extend(
            [
                "",
                "Siguiente paso:",
                (
                    "Mechanical Settings -> Pipe Settings -> Segments and Sizes -> "
                    "seleccionar '{}'".format(schedule_name)
                ),
            ]
        )

        forms.alert("\n".join(result_lines), title="Creacion Exitosa")

    except Exception as err:
        transaction.RollBack()
        error_msg = "Error al crear los segmentos:\n\n{}".format(err)
        forms.alert(error_msg, title="Error")
        print("\nError: {}".format(err))


if __name__ == "__main__":
    main()
