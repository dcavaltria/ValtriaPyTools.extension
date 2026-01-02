from collections import namedtuple

from Autodesk.Revit.DB import (  # type: ignore
    BuiltInParameter,
    DeleteWorksetOption,
    DeleteWorksetSettings,
    ElementWorksetFilter,
    FilteredElementCollector,
    FilteredWorksetCollector,
    Transaction,
    WorksetKind,
    WorksetTable,
)
from pyrevit import forms, script  # type: ignore

uidoc = __revit__.ActiveUIDocument  # type: ignore[attr-defined]
if uidoc is None:
    forms.alert(
        "No hay un documento activo. Abre un modelo antes de ejecutar Delete Workset.",
        exitscript=True,
        title="Delete Workset",
    )
    script.exit()

doc = uidoc.Document
if doc is None:
    forms.alert(
        "No es posible acceder al documento activo.",
        exitscript=True,
        title="Delete Workset",
    )
    script.exit()

output = script.get_output()
logger = script.get_logger()

if not doc.IsWorkshared:
    forms.alert(
        "El documento actual no esta en modo worksharing. Abre un modelo "
        "compartido para gestionar sus worksets.",
        exitscript=True,
        title="Delete Workset",
    )

user_worksets = list(FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))

if not user_worksets:
    forms.alert(
        "Solo existen worksets del sistema y no pueden eliminarse.",
        exitscript=True,
        title="Delete Workset",
    )

WorksetChoice = namedtuple("WorksetChoice", ["name", "workset"])
workset_choices = sorted(
    [WorksetChoice(ws.Name, ws) for ws in user_worksets],
    key=lambda item: item.name.lower(),
)

selected_choice = forms.SelectFromList.show(
    workset_choices,
    name_attr="name",
    title="Selecciona el workset a eliminar",
    button_name="Continuar",
)

if not selected_choice:
    script.exit()

selected_ws = selected_choice.workset

instance_elements = list(
    FilteredElementCollector(doc)
    .WherePasses(ElementWorksetFilter(selected_ws.Id))
    .WhereElementIsNotElementType()
)
type_elements = list(
    FilteredElementCollector(doc)
    .WherePasses(ElementWorksetFilter(selected_ws.Id))
    .WhereElementIsElementType()
)
elements_to_process = instance_elements + type_elements

element_count = len(instance_elements)
type_count = len(type_elements)

action = forms.CommandSwitchWindow.show(
    ["Eliminar elementos y workset", "Transferir elementos a otro workset"],
    message="Que deseas hacer con los elementos del workset seleccionado?",
    title="Opciones de contenido",
)

if not action:
    script.exit()

transfer_mode = action.startswith("Transferir")
target_ws = None

if transfer_mode:
    destination_choices = [
        item for item in workset_choices if item.workset.Id != selected_ws.Id
    ]
    if not destination_choices:
        forms.alert(
            "No hay otro workset disponible para transferir los elementos.",
            exitscript=True,
            title="Delete Workset",
        )

    target_choice = forms.SelectFromList.show(
        destination_choices,
        name_attr="name",
        title="Selecciona el workset de destino",
        button_name="Transferir",
    )

    if not target_choice:
        script.exit()

    target_ws = target_choice.workset

summary_lines = [
    "Workset origen: {0}".format(selected_ws.Name),
    "Elementos detectados: {0}".format(element_count),
]

if type_count:
    summary_lines.append("Tipos detectados: {0}".format(type_count))

if transfer_mode and target_ws:
    summary_lines.append(
        "Accion: Transferencia de elementos a '{0}'".format(target_ws.Name)
    )
else:
    summary_lines.append("Accion: Eliminacion de elementos")

confirm = forms.alert(
    "\n".join(summary_lines) + "\n\nDeseas continuar?",
    yes=True,
    no=True,
    ok=False,
    title="Confirmar operacion",
)

if not confirm:
    script.exit()

deleted_ids = []
failed_ops = []
transferred = 0

transaction = Transaction(doc, "Delete Workset :: {0}".format(selected_ws.Name))
transaction_started = False

try:
    transaction.Start()
    transaction_started = True

    if transfer_mode and target_ws:
        target_id = target_ws.Id.IntegerValue
        for element in elements_to_process:
            ws_param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
            if ws_param and not ws_param.IsReadOnly:
                try:
                    ws_param.Set(target_id)
                    transferred += 1
                except Exception as item_exc:
                    failed_ops.append((element, item_exc))
            else:
                failed_ops.append(
                    (element, Exception("Parametro de workset no editable"))
                )
    else:
        element_ids = [element.Id for element in elements_to_process]
        if element_ids:
            try:
                deleted_collection = doc.Delete(element_ids)
                if deleted_collection:
                    deleted_ids = list(deleted_collection)
            except Exception as delete_exc:
                failed_ops.append((None, delete_exc))

    delete_settings = DeleteWorksetSettings()
    try:
        if transfer_mode and target_ws:
            delete_settings.DeleteWorksetOption = DeleteWorksetOption.MoveElementsToWorkset
            delete_settings.WorksetId = target_ws.Id
        else:
            delete_settings.DeleteWorksetOption = DeleteWorksetOption.DeleteAllElements

        if not WorksetTable.CanDeleteWorkset(doc, selected_ws.Id, delete_settings):
            raise Exception(
                "No fue posible eliminar el workset indicado con la configuracion actual."
            )

        WorksetTable.DeleteWorkset(doc, selected_ws.Id, delete_settings)
    finally:
        delete_settings.Dispose()

    transaction.Commit()
except Exception as run_exc:
    if transaction_started:
        transaction.RollBack()

    forms.alert(
        "No se pudo completar la operacion:\n{0}".format(run_exc),
        exitscript=True,
        title="Delete Workset",
    )
    script.exit()

if transfer_mode:
    output.print_md("### Transferencia realizada")
    output.print_md("- Elementos transferidos: {0}".format(transferred))
else:
    output.print_md("### Eliminacion realizada")
    output.print_md("- Elementos eliminados: {0}".format(len(set(deleted_ids))))

if failed_ops:
    output.print_md("### Elementos con incidencias")
    for element, cause in failed_ops:
        elem_repr = "Desconocido"
        if element:
            elem_name = None
            try:
                elem_name = element.Name
            except Exception:
                elem_name = None

            if elem_name:
                elem_repr = "{0} | Id {1}".format(elem_name, element.Id.IntegerValue)
            else:
                elem_repr = "{0} | Id {1}".format(
                    element.GetType().Name, element.Id.IntegerValue
                )

        logger.warning(
            "Elemento %s no se proceso correctamente. Motivo: %s",
            elem_repr,
            cause,
        )
        output.print_md("- {0} -> {1}".format(elem_repr, cause))
else:
    output.print_md("- No se registraron incidencias.")

issues_count = len(failed_ops)
if transfer_mode:
    final_summary = "Elementos transferidos: {0}\nIncidencias: {1}".format(
        transferred, issues_count
    )
else:
    final_summary = "Elementos eliminados: {0}\nIncidencias: {1}".format(
        len(set(deleted_ids)), issues_count
    )

forms.alert(
    "Proceso finalizado.\n{0}\nConsulta la ventana de resultados para el detalle.".format(
        final_summary
    ),
    title="Delete Workset",
)
