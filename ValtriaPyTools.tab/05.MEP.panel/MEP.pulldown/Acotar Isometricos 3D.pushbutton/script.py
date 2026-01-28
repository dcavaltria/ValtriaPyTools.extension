# -*- coding: utf-8 -*-
"""
Acotar Isometricos 3D

Steps:
1) Validate active view is 3D (not perspective/template).
2) Select pipe segments (linear only).
3) Create pipe tags or fallback to text notes with length in mm.
"""

__title__ = "Acotar\nIsometricos 3D"
__author__ = "Valtria"

import traceback

from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    IndependentTag,
    Line,
    Reference,
    TagMode,
    TagOrientation,
    TextNote,
    TextNoteType,
    Transaction,
    ViewType,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from pyrevit import forms, revit, script

import valtria_lib as vlib


TITLE = "Acotar Isometricos 3D"
MM_PER_FOOT = 304.8
OFFSET_MM = 300.0
STACK_MM = 200.0
MAX_COLS = 6


class PipeSelectionFilter(ISelectionFilter):
    def __init__(self):
        self._cat_id = int(BuiltInCategory.OST_PipeCurves)

    def AllowElement(self, elem):
        if elem is None or elem.Category is None:
            return False
        return elem.Category.Id.IntegerValue == self._cat_id

    def AllowReference(self, reference, position):
        return False


def add_xyz(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def scale_xyz(v, scale):
    return XYZ(v.X * scale, v.Y * scale, v.Z * scale)


def feet_to_mm(value_feet):
    try:
        return float(value_feet) * MM_PER_FOOT
    except Exception:
        return None


def format_length_mm(length_feet):
    length_mm = feet_to_mm(length_feet)
    if length_mm is None:
        return None
    return "{0:.0f} mm".format(length_mm)


def is_pipe(elem):
    if elem is None or elem.Category is None:
        return False
    return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves)


def get_line_curve(elem):
    loc = getattr(elem, "Location", None)
    curve = getattr(loc, "Curve", None)
    if curve is None:
        return None
    if isinstance(curve, Line):
        return curve
    return None


def curve_midpoint(curve):
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            return curve.GetEndPoint(0)
        except Exception:
            return None


def get_pipe_tag_type_id(doc):
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_PipeTags
    ).WhereElementIsElementType()
    for tag_type in collector:
        return tag_type.Id
    return None


def get_text_type_id(doc):
    collector = FilteredElementCollector(doc).OfClass(TextNoteType)
    for text_type in collector:
        return text_type.Id
    return None


def get_offset(view, index):
    base = vlib.mm_to_internal(OFFSET_MM)
    step = vlib.mm_to_internal(STACK_MM)
    col = index % MAX_COLS
    row = index // MAX_COLS
    right = scale_xyz(view.RightDirection, base + col * step)
    up = scale_xyz(view.UpDirection, row * step)
    return add_xyz(right, up)


def get_preselected_pipes(uidoc, doc):
    elements = []
    try:
        ids = uidoc.Selection.GetElementIds()
    except Exception:
        ids = []
    for elem_id in ids:
        elem = doc.GetElement(elem_id)
        if is_pipe(elem):
            elements.append(elem)
    return elements


def pick_pipes(uidoc):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, PipeSelectionFilter(), "Select pipe segments"
        )
    except OperationCanceledException:
        return []
    if not refs:
        return []
    return refs


def collect_linear_pipes(doc, elements):
    pipes = []
    skipped_non_linear = 0
    skipped_no_center = 0
    for elem in elements:
        curve = get_line_curve(elem)
        if curve is None:
            skipped_non_linear += 1
            continue
        center = curve_midpoint(curve)
        if center is None:
            skipped_no_center += 1
            continue
        length_feet = vlib.read_length(elem)
        if length_feet is None:
            try:
                length_feet = curve.Length
            except Exception:
                length_feet = None
        pipes.append((elem, curve, center, length_feet))
    return pipes, skipped_non_linear, skipped_no_center


def create_pipe_tag(doc, view, elem, point, tag_type_id):
    try:
        ref = Reference(elem)
    except Exception:
        ref = None
    if ref is None:
        return None
    try:
        tag = IndependentTag.Create(
            doc,
            view.Id,
            ref,
            True,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            point,
        )
    except Exception:
        return None
    if tag is None:
        return None
    if tag_type_id is not None:
        try:
            if tag.GetTypeId() != tag_type_id:
                tag.ChangeTypeId(tag_type_id)
        except Exception:
            pass
    return tag


def create_text_note(doc, view, point, text, text_type_id):
    if text_type_id is None:
        return None
    try:
        return TextNote.Create(doc, view.Id, point, text, text_type_id)
    except Exception:
        return None


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None or uidoc is None:
        forms.alert("No active document.", exitscript=True)
        return

    view = doc.ActiveView
    if view is None or view.IsTemplate:
        forms.alert("Active view is not valid.", exitscript=True)
        return
    if view.ViewType != ViewType.ThreeD:
        forms.alert("Active view must be a 3D view.", exitscript=True)
        return
    if getattr(view, "IsPerspective", False):
        forms.alert("Perspective 3D views are not supported.", exitscript=True)
        return

    elements = get_preselected_pipes(uidoc, doc)
    if not elements:
        forms.alert("Select pipe segments to annotate.", title=TITLE, exitscript=False)
        refs = pick_pipes(uidoc)
        if not refs:
            return
        elements = [doc.GetElement(ref.ElementId) for ref in refs]

    pipes, skipped_non_linear, skipped_no_center = collect_linear_pipes(doc, elements)
    if not pipes:
        forms.alert("No linear pipe segments found.", exitscript=True)
        return

    tag_type_id = get_pipe_tag_type_id(doc)
    text_type_id = get_text_type_id(doc)

    created_tags = 0
    created_text = 0
    skipped_no_length = 0

    t = Transaction(doc, "Acotar Isometricos 3D")
    t.Start()
    try:
        for index, (elem, curve, center, length_feet) in enumerate(pipes):
            offset = get_offset(view, index)
            head = add_xyz(center, offset)
            tag = create_pipe_tag(doc, view, elem, head, tag_type_id)
            if tag is not None:
                created_tags += 1
                continue

            text = format_length_mm(length_feet)
            if not text:
                skipped_no_length += 1
                continue
            note = create_text_note(doc, view, head, text, text_type_id)
            if note is not None:
                created_text += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    message = [
        "Tags created: {0}".format(created_tags),
        "Text notes created: {0}".format(created_text),
    ]
    if skipped_non_linear:
        message.append("Skipped non-linear pipes: {0}".format(skipped_non_linear))
    if skipped_no_center:
        message.append("Skipped no-center pipes: {0}".format(skipped_no_center))
    if skipped_no_length:
        message.append("Skipped without length: {0}".format(skipped_no_length))
    forms.alert("\n".join(message), title=TITLE)


logger = script.get_logger()

try:
    main()
except Exception as ex:
    logger.error("AcotarIsometricos3D failed: {0}".format(ex))
    logger.error(traceback.format_exc())
    forms.alert("Error: {0}".format(ex), title=TITLE)
