from Autodesk.Revit.DB import (  # type: ignore
    DatumEnds,
    FilteredElementCollector,
    Grid,
    Transaction,
)
from pyrevit import forms, script  # type: ignore


TITLE = "Mostrar burbujas de grids"

uidoc = __revit__.ActiveUIDocument  # type: ignore[attr-defined]
if uidoc is None:
    forms.alert(
        "No hay un documento activo. Abre un modelo antes de ejecutar la herramienta.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

doc = uidoc.Document
active_view = uidoc.ActiveView

if doc is None or active_view is None:
    forms.alert(
        "No es posible acceder al documento o la vista activa.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if doc.IsFamilyDocument:
    forms.alert(
        "Esta herramienta solo se puede ejecutar sobre un modelo de proyecto.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

collector = FilteredElementCollector(doc, active_view.Id).OfClass(Grid)
grids = [grid for grid in collector if isinstance(grid, Grid)]

if not grids:
    forms.alert(
        "La vista activa no contiene grids disponibles.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

transaction = Transaction(doc, "Mostrar burbujas de grids")
output = script.get_output()

updated = 0
skipped = []
transaction_started = False

try:
    transaction.Start()
    transaction_started = True

    for grid in grids:
        try:
            changed = False
            if not grid.IsBubbleVisibleInView(DatumEnds.End0, active_view):
                can_show = True
                can_show_method = getattr(grid, "CanShowBubbleInView", None)
                if callable(can_show_method):
                    can_show = can_show_method(DatumEnds.End0, active_view)
                if can_show:
                    grid.ShowBubbleInView(DatumEnds.End0, active_view)
                    changed = True
            if not grid.IsBubbleVisibleInView(DatumEnds.End1, active_view):
                can_show = True
                can_show_method = getattr(grid, "CanShowBubbleInView", None)
                if callable(can_show_method):
                    can_show = can_show_method(DatumEnds.End1, active_view)
                if can_show:
                    grid.ShowBubbleInView(DatumEnds.End1, active_view)
                    changed = True
            if changed:
                updated += 1
        except Exception as grid_exc:
            skipped.append((grid, grid_exc))

    transaction.Commit()
except Exception as run_exc:
    if transaction_started:
        transaction.RollBack()
    forms.alert(
        "La operacion no pudo completarse:\n{0}".format(run_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

output.print_md("### Burbujas de grids visibles")
output.print_md("- Grids procesados: {0}".format(len(grids)))
output.print_md("- Grids actualizados: {0}".format(updated))
output.print_md("- Grids sin cambios: {0}".format(len(grids) - updated - len(skipped)))

if skipped:
    output.print_md("### Grids con incidencias")
    for grid, cause in skipped:
        name = getattr(grid, "Name", None)
        identifier = "{0} | Id {1}".format(name or "Grid", grid.Id.IntegerValue)
        output.print_md("- {0} -> {1}".format(identifier, cause))

forms.alert(
    "Proceso finalizado.\nGrids actualizados: {0}\nIncidencias: {1}".format(
        updated, len(skipped)
    ),
    title=TITLE,
)
