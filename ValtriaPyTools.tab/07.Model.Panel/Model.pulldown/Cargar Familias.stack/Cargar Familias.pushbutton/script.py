# -*- coding: utf-8 -*-
"""Selecciona, confirma y carga familias RFA desde una carpeta."""

import os
import traceback

import clr

try:
    clr.AddReference('RevitAPI')
    clr.AddReference('RevitAPIUI')
except Exception:
    pass

from Autodesk.Revit.DB import IFamilyLoadOptions, Transaction  # noqa
from Autodesk.Revit.UI import TaskDialog  # noqa

from pyrevit import forms, revit


class FamilyLoadOptions(IFamilyLoadOptions):
    """Reload families without prompting the user."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        if hasattr(overwriteParameterValues, 'Value'):
            overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        if hasattr(overwriteParameterValues, 'Value'):
            overwriteParameterValues.Value = True
        return True


class SelectListItem(object):
    """Adapter item for pyRevit selection dialogs."""

    def __init__(self, value):
        self.value = value
        self.label = os.path.basename(value)

    @property
    def name(self):
        return self.label


def pick_families():
    folder = forms.pick_folder(title='Selecciona la carpeta con familias RFA')
    if not folder:
        return None, []
    candidates = []
    for entry in os.listdir(folder):
        path = os.path.join(folder, entry)
        if os.path.isfile(path) and entry.lower().endswith('.rfa'):
            candidates.append(path)
    if not candidates:
        forms.alert('No se encontraron archivos .rfa en la carpeta seleccionada.')
        return folder, []
    items = [SelectListItem(path) for path in sorted(candidates)]
    selected = forms.SelectFromList.show(
        items,
        title='Selecciona familias a cargar',
        multiselect=True,
        button_name='Usar seleccion',
        name_attr='name',
    )
    if not selected:
        return folder, []
    return folder, [item.value for item in selected]


def confirm_families(file_paths):
    preview_lines = []
    for file_path in file_paths[:10]:
        preview_lines.append('  - {0}'.format(os.path.basename(file_path)))
    if len(file_paths) > 10:
        preview_lines.append('  ... y {0} mas.'.format(len(file_paths) - 10))
    message_lines = [
        'Se cargarán {0} familias:'.format(len(file_paths)),
        '',
    ]
    message_lines.extend(preview_lines)
    message_lines.append('')
    message_lines.append('Deseas continuar?')
    answer = forms.alert(
        '\n'.join(message_lines),
        title='Confirmar carga de familias',
        options=['Cargar', 'Cancelar'],
        exitscript=False,
    )
    return answer == 'Cargar'


def load_families(doc, files):
    """Load the given families and return result tuples."""
    options = FamilyLoadOptions()
    results = []
    transaction = Transaction(doc, 'Cargar familias desde carpeta')
    transaction.Start()
    try:
        for file_path in files:
            file_name = os.path.basename(file_path)
            try:
                loaded = doc.LoadFamily(file_path, options)
                if loaded:
                    results.append((file_name, True, 'Cargada/Actualizada'))
                else:
                    results.append((file_name, False, 'Revit devolvió False'))
            except Exception as family_error:
                results.append((file_name, False, str(family_error)))
        transaction.Commit()
    except Exception:
        if transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
        raise
    return results


def build_summary(results):
    success = [item for item in results if item[1]]
    failed = [item for item in results if not item[1]]
    lines = []
    lines.append('Familias procesadas: {0}'.format(len(results)))
    lines.append('Exitosas: {0}'.format(len(success)))
    lines.append('Con errores: {0}'.format(len(failed)))
    lines.append('')
    if success:
        lines.append('Cargadas:')
        for name, _, message in success[:10]:
            lines.append('  - {0} ({1})'.format(name, message))
        if len(success) > 10:
            lines.append('  ... y {0} mas.'.format(len(success) - 10))
        lines.append('')
    if failed:
        lines.append('Errores:')
        for name, _, message in failed[:10]:
            lines.append('  - {0}: {1}'.format(name, message))
        if len(failed) > 10:
            lines.append('  ... y {0} mas.'.format(len(failed) - 10))
    return '\n'.join(lines)


def main():
    doc = revit.doc
    if doc is None:
        forms.alert('No hay un documento activo.')
        return
    try:
        _, files = pick_families()
        if not files:
            return
        if not confirm_families(files):
            return
        results = load_families(doc, files)
        TaskDialog.Show('Carga de Familias', build_summary(results))
    except Exception as exc:
        forms.alert(
            'Error al cargar familias:\n{0}\n\nTrace:\n{1}'.format(exc, traceback.format_exc()),
            title='Carga de Familias',
        )


if __name__ == '__main__':
    main()
