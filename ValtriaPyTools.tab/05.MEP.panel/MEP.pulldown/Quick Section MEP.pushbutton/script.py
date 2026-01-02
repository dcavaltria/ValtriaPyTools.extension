# -*- coding: utf-8 -*-
"""
Quick Section MEP

Steps:
1) Select a linear MEP element (duct or pipe).
2) Create a section view parallel to the element axis.
3) Crop 1m around the element.
"""

__title__ = "Quick\nSection"
__author__ = "Valtria"

import os
import re
import traceback

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    Transaction,
    Transform,
    ViewFamily,
    ViewFamilyType,
    ViewSection,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from pyrevit import forms, revit, script


TITLE = "Quick Section MEP"
FEET_PER_METER = 3.28083989501312
MARGIN_M = 1.0


class LinearMepFilter(ISelectionFilter):
    def __init__(self):
        self._cats = set([
            int(BuiltInCategory.OST_DuctCurves),
            int(BuiltInCategory.OST_PipeCurves),
        ])

    def AllowElement(self, elem):
        if elem is None or elem.Category is None:
            return False
        return elem.Category.Id.IntegerValue in self._cats

    def AllowReference(self, reference, position):
        return False


def normalize_xyz(vec):
    try:
        return vec.Normalize()
    except Exception:
        return None


def scale_xyz(vec, scale):
    return XYZ(vec.X * scale, vec.Y * scale, vec.Z * scale)


def add_xyz(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def get_axis_curve_and_dir(elem):
    try:
        loc = elem.Location
    except Exception:
        loc = None
    curve = getattr(loc, "Curve", None)
    if curve is None:
        return None, None
    axis_dir = None
    try:
        axis_dir = normalize_xyz(curve.Direction)
    except Exception:
        axis_dir = None
    if axis_dir is None:
        try:
            deriv = curve.ComputeDerivatives(0.5, True)
            axis_dir = normalize_xyz(deriv.BasisX)
        except Exception:
            axis_dir = None
    return curve, axis_dir


def get_section_view_type(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
    for vtype in collector:
        try:
            if vtype.ViewFamily == ViewFamily.Section:
                return vtype
        except Exception:
            continue
    return None


def get_size_params(elem):
    width = None
    height = None
    diameter = None

    for bip in (
        BuiltInParameter.RBS_CURVE_WIDTH_PARAM,
        BuiltInParameter.RBS_CURVE_HEIGHT_PARAM,
    ):
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if param is None:
            continue
        try:
            value = param.AsDouble()
        except Exception:
            value = None
        if value and value > 0:
            if bip == BuiltInParameter.RBS_CURVE_WIDTH_PARAM:
                width = value
            else:
                height = value

    for bip in (
        BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
        getattr(BuiltInParameter, "RBS_PIPE_DIAMETER_PARAM", None),
        getattr(BuiltInParameter, "RBS_DUCT_DIAMETER_PARAM", None),
    ):
        if bip is None:
            continue
        try:
            param = elem.get_Parameter(bip)
        except Exception:
            param = None
        if param is None:
            continue
        try:
            value = param.AsDouble()
        except Exception:
            value = None
        if value and value > 0:
            diameter = value
            break

    return width, height, diameter


def build_section_box(center, axis_dir, length, half_y, half_z):
    up = XYZ.BasisZ
    if abs(axis_dir.DotProduct(up)) > 0.9:
        up = XYZ.BasisY
    view_dir = normalize_xyz(axis_dir.CrossProduct(up))
    if view_dir is None:
        view_dir = XYZ.BasisX
    up = normalize_xyz(view_dir.CrossProduct(axis_dir))
    if up is None:
        up = XYZ.BasisZ

    transform = Transform.Identity
    transform.Origin = center
    transform.BasisX = axis_dir
    transform.BasisY = up
    transform.BasisZ = view_dir

    bbox = BoundingBoxXYZ()
    bbox.Transform = transform
    bbox.Min = XYZ(-length, -half_y, -half_z)
    bbox.Max = XYZ(length, half_y, half_z)
    return bbox


def next_view_name(doc, username, elem_id):
    clean_user = re.sub(r"[^A-Za-z0-9_\\-]", "_", username or "user")
    prefix = "quicksection_{0}_".format(clean_user)
    collector = FilteredElementCollector(doc).OfClass(ViewSection)
    max_idx = 0
    pattern = re.compile(r"^" + re.escape(prefix) + r"(\\d{3})_elemid:")
    for view in collector:
        try:
            name = view.Name
        except Exception:
            name = ""
        if not name:
            continue
        match = pattern.match(name)
        if match:
            try:
                idx = int(match.group(1))
            except Exception:
                idx = 0
            if idx > max_idx:
                max_idx = idx
    new_idx = max_idx + 1
    return "{0}{1:03d}_elemid_{2}".format(prefix, new_idx, elem_id)


def pick_linear_mep(uidoc):
    forms.alert(
        "Select a duct or pipe to create the section.",
        title=TITLE,
        exitscript=False,
    )
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            LinearMepFilter(),
            "Select a duct or pipe",
        )
    except OperationCanceledException:
        return None
    if ref is None:
        return None
    return uidoc.Document.GetElement(ref.ElementId)


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    if doc is None or uidoc is None:
        forms.alert("No active document.", exitscript=True)
        return

    element = pick_linear_mep(uidoc)
    if element is None:
        return

    curve, axis_dir = get_axis_curve_and_dir(element)
    if curve is None or axis_dir is None:
        forms.alert("Selected element has no valid axis.", exitscript=True)
        return

    try:
        center = curve.Evaluate(0.5, True)
    except Exception:
        try:
            center = curve.GetEndPoint(0)
        except Exception:
            forms.alert("Unable to compute element center.", exitscript=True)
            return

    try:
        half_len = curve.Length / 2.0
    except Exception:
        half_len = 0.0

    width, height, diameter = get_size_params(element)
    if diameter:
        half_y = diameter / 2.0
        half_z = diameter / 2.0
    else:
        half_y = (height / 2.0) if height else 0.0
        half_z = (width / 2.0) if width else 0.0
        if half_y == 0.0 and half_z == 0.0:
            half_y = 0.5 * FEET_PER_METER
            half_z = 0.5 * FEET_PER_METER

    margin = MARGIN_M * FEET_PER_METER
    half_len = half_len + margin
    half_y = half_y + margin
    half_z = half_z + margin

    vtype = get_section_view_type(doc)
    if vtype is None:
        forms.alert("No section view type found.", exitscript=True)
        return

    bbox = build_section_box(center, axis_dir, half_len, half_y, half_z)

    username = os.environ.get("USERNAME", "user")
    view_name = next_view_name(doc, username, element.Id.IntegerValue)

    t = Transaction(doc, "Quick Section MEP")
    t.Start()
    try:
        section = ViewSection.CreateSection(doc, vtype.Id, bbox)
        section.Name = view_name
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    try:
        uidoc.ActiveView = section
    except Exception:
        pass


logger = script.get_logger()

try:
    main()
except Exception as ex:
    logger.error("QuickSection failed: {0}".format(ex))
    logger.error(traceback.format_exc())
    forms.alert("Error: {0}".format(ex), title=TITLE)
