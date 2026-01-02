from Autodesk.Revit.DB import (  # type: ignore
    FilteredWorksetCollector,
    Transaction,
    WorksetKind,
    WorksetTable,
)
from pyrevit import forms, script  # type: ignore

TITLE = "Renombrar Worksets"

uidoc = __revit__.ActiveUIDocument  # type: ignore[attr-defined]
if uidoc is None:
    forms.alert(
        "No hay un documento activo. Abre un modelo antes de ejecutar la herramienta.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

doc = uidoc.Document
if doc is None:
    forms.alert(
        "No es posible acceder al documento activo.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

if not doc.IsWorkshared:
    forms.alert(
        "El documento actual no esta en modo worksharing. Habilita los worksets antes de continuar.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

user_worksets = list(FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
if not user_worksets:
    forms.alert(
        "No hay worksets de usuario en el modelo.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

search_text = forms.ask_for_string(
    prompt="Texto a buscar en el nombre del workset:",
    title="Buscar y reemplazar",
)
if not search_text:
    script.exit()

replace_text = forms.ask_for_string(
    prompt="Texto de reemplazo (puede ser vacio):",
    title="Buscar y reemplazar",
)
if replace_text is None:
    script.exit()

operations = []
for ws in user_worksets:
    old_name = ws.Name
    new_name = old_name.replace(search_text, replace_text)
    if new_name != old_name:
        operations.append({"workset": ws, "old": old_name, "new": new_name})

if not operations:
    forms.alert(
        "No se encontraron worksets con el texto indicado.",
        exitscript=True,
        title=TITLE,
    )
    script.exit()

id_to_name = {ws.Id.IntegerValue: ws.Name for ws in user_worksets}
ops_by_id = {op["workset"].Id.IntegerValue: op for op in operations}
name_owner = {}
conflicts = []
invalid_names = []

for ws in user_worksets:
    target = ops_by_id.get(ws.Id.IntegerValue)
    final_name = target["new"] if target else ws.Name

    if not final_name.strip():
        invalid_names.append((ws.Name, final_name))

    owner_id = name_owner.get(final_name)
    if owner_id is not None and owner_id != ws.Id.IntegerValue:
        conflicts.append((ws.Name, id_to_name[owner_id], final_name))
    else:
        name_owner[final_name] = ws.Id.IntegerValue

if invalid_names:
    lines = ["Hay nombres vacios tras aplicar el reemplazo. Revisa:"]
    for old, new in invalid_names:
        lines.append("- {0} -> '{1}'".format(old, new))
    forms.alert("\n".join(lines), exitscript=True, title=TITLE)
    script.exit()

if conflicts:
    lines = ["Los siguientes nombres resultantes se duplican:"]
    for current, other, final in conflicts:
        lines.append("- {0} / {1} -> {2}".format(current, other, final))
    lines.append("\nAjusta el texto de reemplazo para evitar duplicados.")
    forms.alert("\n".join(lines), exitscript=True, title=TITLE)
    script.exit()

ops_sorted = sorted(operations, key=lambda op: op["old"].lower())
preview_lines = ["Cambios propuestos ({0}):".format(len(ops_sorted))]
preview_lines.extend(
    "- {0} -> {1}".format(op["old"], op["new"]) for op in ops_sorted
)

confirm = forms.alert(
    "\n".join(preview_lines) + "\n\nDeseas continuar?",
    yes=True,
    no=True,
    ok=False,
    title="Confirmar renombrado",
)
if not confirm:
    script.exit()

output = script.get_output()
transaction = Transaction(doc, "Renombrar worksets (buscar/reemplazar)")
transaction_started = False

try:
    transaction.Start()
    transaction_started = True

    for op in ops_sorted:
        WorksetTable.RenameWorkset(doc, op["workset"].Id, op["new"])

    transaction.Commit()
except Exception as run_exc:  # noqa: BLE001
    if transaction_started:
        transaction.RollBack()

    forms.alert(
        "No se pudo completar el renombrado:\n{0}".format(run_exc),
        exitscript=True,
        title=TITLE,
    )
    script.exit()

output.print_md("### Worksets renombrados ({0})".format(len(ops_sorted)))
for op in ops_sorted:
    output.print_md("- {0} -> {1}".format(op["old"], op["new"]))

forms.alert(
    "Se renombraron {0} worksets.\nConsulta la ventana de resultados para el detalle.".format(
        len(ops_sorted)
    ),
    title=TITLE,
)
