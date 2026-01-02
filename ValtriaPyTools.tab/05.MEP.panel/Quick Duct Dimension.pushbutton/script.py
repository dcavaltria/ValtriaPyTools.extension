# -*- coding: utf-8 -*-
"""
Quick Duct Dimension (Auto)

Steps:
1) Select a duct.
2) Auto-find nearest grids perpendicular to the duct axis and duct width refs.
3) Pick an insertion point for the dimension line.
4) Create the dimension.
"""

__title__ = "Quick\nDuct Dim"
__author__ = "Valtria"

import traceback

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    DatumExtentType,
    DimensionStyleType,
    DimensionType,
    FilteredElementCollector,
    GeometryInstance,
    Line,
    Options,
    Plane,
    PlanarFace,
    ReferenceArray,
    Reference,
    SketchPlane,
    Solid,
    Transaction,
    UV,
    ViewType,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

try:
    from Autodesk.Revit.Exceptions import OperationCanceledException
except Exception:
    OperationCanceledException = Exception

from pyrevit import forms, revit, script


TITLE = "Quick Duct Dimension"
GRID_DIR_TOL = 0.9
FACE_AXIS_TOL = 0.1
FACE_PERP_TOL = 0.8


class DuctSelectionFilter(ISelectionFilter):
    def __init__(self):
        self._cat_id = int(BuiltInCategory.OST_DuctCurves)

    def AllowElement(self, elem):
        if elem is None or elem.Category is None:
            return False
        return elem.Category.Id.IntegerValue == self._cat_id

    def AllowReference(self, reference, position):
        return False


def add_xyz(a, b):
    return XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)


def sub_xyz(a, b):
    return XYZ(a.X - b.X, a.Y - b.Y, a.Z - b.Z)


def scale_xyz(v, scale):
    return XYZ(v.X * scale, v.Y * scale, v.Z * scale)


def normalize_xyz(v):
    try:
        return v.Normalize()
    except Exception:
        return None


def project_to_plane(vec, normal):
    try:
        dot = vec.DotProduct(normal)
    except Exception:
        return None
    proj = sub_xyz(vec, scale_xyz(normal, dot))
    return normalize_xyz(proj)


def ensure_view_work_plane(doc, view):
    try:
        current = view.SketchPlane
    except Exception:
        current = None
    if current is not None:
        return True
    try:
        plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin)
    except Exception:
        return False
    t = Transaction(doc, "Set work plane for dimension")
    t.Start()
    try:
        sketch = SketchPlane.Create(doc, plane)
        view.SketchPlane = sketch
        t.Commit()
        return True
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return False


def iter_solids(geometry):
    if geometry is None:
        return
    for obj in geometry:
        if isinstance(obj, Solid):
            try:
                if obj.Faces.Size > 0:
                    yield obj
            except Exception:
                yield obj
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
            except Exception:
                inst_geom = None
            for solid in iter_solids(inst_geom):
                yield solid


def get_axis_curve_and_dir(duct):
    try:
        loc = duct.Location
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


def get_default_dimension_type(doc):
    collector = FilteredElementCollector(doc).OfClass(DimensionType)
    for dim_type in collector:
        try:
            if dim_type.StyleType == DimensionStyleType.Linear:
                return dim_type
        except Exception:
            continue
    return None


def iter_grid_curves_in_view(grid, view):
    curves = None
    try:
        curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
    except Exception:
        curves = None
    if curves:
        for curve in curves:
            yield curve
        return
    try:
        curves = grid.GetCurvesInView(DatumExtentType.Model, view)
    except Exception:
        curves = None
    if curves:
        for curve in curves:
            yield curve
        return
    curve = getattr(grid, "Curve", None)
    if curve is not None:
        yield curve


def curve_direction(curve):
    try:
        return normalize_xyz(curve.Direction)
    except Exception:
        pass
    try:
        deriv = curve.ComputeDerivatives(0.5, True)
        return normalize_xyz(deriv.BasisX)
    except Exception:
        return None


def get_grid_reference_from_curve_or_plane(grid, curve):
    if curve is not None:
        try:
            ref = curve.Reference
        except Exception:
            ref = None
        if ref is not None:
            return ref
    for method_name in ("GetPlaneReference", "GetReference"):
        method = getattr(grid, method_name, None)
        if callable(method):
            try:
                ref = method()
            except Exception:
                ref = None
            if ref is not None:
                return ref
    try:
        return Reference(grid)
    except Exception:
        return None


def curve_point(curve):
    try:
        return curve.Evaluate(0.5, True)
    except Exception:
        try:
            return curve.GetEndPoint(0)
        except Exception:
            return None


def find_nearest_grids(doc, view, center, axis_dir, perp_dir):
    neg_ref = None
    pos_ref = None
    neg_dist = None
    pos_dist = None

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Grids
    ).WhereElementIsNotElementType()

    for grid in collector:
        ref_fallback = get_grid_reference_from_curve_or_plane(grid, None)
        for curve in iter_grid_curves_in_view(grid, view):
            if curve is None:
                continue
            ref = get_grid_reference_from_curve_or_plane(grid, curve) or ref_fallback
            if ref is None:
                continue
            grid_dir = curve_direction(curve)
            if grid_dir is None:
                continue
            if grid_dir is None:
                continue
            try:
                alignment = abs(grid_dir.DotProduct(axis_dir))
            except Exception:
                alignment = 0.0
            if alignment < GRID_DIR_TOL:
                continue
            pt = curve_point(curve)
            if pt is None:
                continue
            signed = sub_xyz(center, pt).DotProduct(perp_dir)
            if signed > 0:
                if pos_dist is None or abs(signed) < abs(pos_dist):
                    pos_dist = signed
                    pos_ref = ref
            elif signed < 0:
                if neg_dist is None or abs(signed) < abs(neg_dist):
                    neg_dist = signed
                    neg_ref = ref

    return neg_ref, pos_ref, neg_dist, pos_dist


def get_rectangular_side_face_refs(doc, view, duct, axis_dir, perp_dir):
    options = Options()
    options.ComputeReferences = True
    options.IncludeNonVisibleObjects = True
    try:
        options.View = view
    except Exception:
        pass
    try:
        geometry = duct.get_Geometry(options)
    except Exception:
        geometry = None

    pos_face = None
    neg_face = None
    pos_dot = -1.0
    neg_dot = 1.0

    for solid in iter_solids(geometry):
        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue
            normal = face.FaceNormal
            try:
                if abs(normal.DotProduct(axis_dir)) > FACE_AXIS_TOL:
                    continue
            except Exception:
                continue
            try:
                dot = normal.Normalize().DotProduct(perp_dir)
            except Exception:
                continue
            if abs(dot) < FACE_PERP_TOL:
                continue
            if dot > 0 and dot > pos_dot:
                pos_dot = dot
                pos_face = face
            elif dot < 0 and dot < neg_dot:
                neg_dot = dot
                neg_face = face

    neg_ref = neg_face.Reference if neg_face is not None else None
    pos_ref = pos_face.Reference if pos_face is not None else None
    return neg_ref, pos_ref


def get_duct_diameter(duct):
    candidates = [
        BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
        getattr(BuiltInParameter, "RBS_DUCT_DIAMETER_PARAM", None),
        getattr(BuiltInParameter, "RBS_PIPE_DIAMETER_PARAM", None),
    ]
    for bip in candidates:
        if bip is None:
            continue
        try:
            param = duct.get_Parameter(bip)
        except Exception:
            param = None
        if param is None:
            continue
        try:
            value = param.AsDouble()
        except Exception:
            value = None
        if value and value > 0:
            return value
    try:
        param = duct.LookupParameter("Diameter")
    except Exception:
        param = None
    if param is not None:
        try:
            value = param.AsDouble()
        except Exception:
            value = None
        if value and value > 0:
            return value
    return None


def get_invisible_line_style(doc):
    try:
        lines_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        lines_cat = None
    if lines_cat is None:
        return None
    for subcat in lines_cat.SubCategories:
        try:
            name = subcat.Name
        except Exception:
            name = None
        if name and name.lower() == "invisible lines":
            return subcat.GetGraphicsStyle(1)
    return None


def create_round_edge_refs(doc, view, center, axis_dir, perp_dir, radius, length):
    style = get_invisible_line_style(doc)
    half = max(length / 2.0, radius * 2.0, 5.0)
    start_a = add_xyz(add_xyz(center, scale_xyz(axis_dir, -half)), scale_xyz(perp_dir, -radius))
    end_a = add_xyz(add_xyz(center, scale_xyz(axis_dir, half)), scale_xyz(perp_dir, -radius))
    start_b = add_xyz(add_xyz(center, scale_xyz(axis_dir, -half)), scale_xyz(perp_dir, radius))
    end_b = add_xyz(add_xyz(center, scale_xyz(axis_dir, half)), scale_xyz(perp_dir, radius))

    curve_a = doc.Create.NewDetailCurve(view, Line.CreateBound(start_a, end_a))
    curve_b = doc.Create.NewDetailCurve(view, Line.CreateBound(start_b, end_b))

    if style is not None:
        try:
            curve_a.LineStyle = style
            curve_b.LineStyle = style
        except Exception:
            pass

    ref_a = curve_a.GeometryCurve.Reference
    ref_b = curve_b.GeometryCurve.Reference
    return ref_a, ref_b


def pick_duct(doc, uidoc):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, DuctSelectionFilter(), "Select a duct"
        )
    except OperationCanceledException:
        return None
    if ref is None:
        return None
    return doc.GetElement(ref.ElementId)


def pick_insertion_point(doc, uidoc, view):
    if not ensure_view_work_plane(doc, view):
        forms.alert("No work plane set in current view.", title=TITLE)
        return None
    try:
        return uidoc.Selection.PickPoint("Pick dimension insertion point")
    except OperationCanceledException:
        return None


def build_dimension_line(view, point, direction, length=200.0):
    start = add_xyz(point, scale_xyz(direction, -length))
    end = add_xyz(point, scale_xyz(direction, length))
    return Line.CreateBound(start, end)


def append_unique(refs, ref, doc):
    if ref is None:
        return
    try:
        key = ref.ConvertToStableRepresentation(doc)
    except Exception:
        try:
            key = "{0}".format(ref.ElementId.IntegerValue)
        except Exception:
            key = None
    if key:
        for existing in refs:
            try:
                if key == existing.ConvertToStableRepresentation(doc):
                    return
            except Exception:
                continue
    refs.append(ref)


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    view = doc.ActiveView

    if view is None or view.IsTemplate or view.ViewType == ViewType.ThreeD:
        forms.alert("Active view is not valid for dimensioning.", exitscript=True)
        return

    dim_type = get_default_dimension_type(doc)
    if dim_type is None:
        forms.alert("No linear dimension type found.", exitscript=True)
        return

    duct = pick_duct(doc, uidoc)
    if duct is None:
        return

    axis_curve, axis_dir = get_axis_curve_and_dir(duct)
    if axis_curve is None or axis_dir is None:
        forms.alert("Could not resolve duct axis.", exitscript=True)
        return

    axis_proj = project_to_plane(axis_dir, view.ViewDirection)
    if axis_proj is None:
        forms.alert("Duct axis is parallel to view direction.", exitscript=True)
        return

    perp_dir = normalize_xyz(view.ViewDirection.CrossProduct(axis_proj))
    if perp_dir is None:
        forms.alert("Could not resolve perpendicular direction.", exitscript=True)
        return

    try:
        center = axis_curve.Evaluate(0.5, True)
    except Exception:
        try:
            center = axis_curve.GetEndPoint(0)
        except Exception:
            forms.alert("Could not resolve duct center.", exitscript=True)
            return

    neg_grid_ref, pos_grid_ref, neg_grid_dist, pos_grid_dist = find_nearest_grids(
        doc, view, center, axis_proj, perp_dir
    )
    if neg_grid_ref is None and pos_grid_ref is None:
        forms.alert(
            "No aligned grids found in this view. Only duct width will be dimensioned.",
            title=TITLE,
            exitscript=False,
        )

    neg_face_ref, pos_face_ref = get_rectangular_side_face_refs(
        doc, view, duct, axis_proj, perp_dir
    )

    use_round = False
    round_radius = None
    if neg_face_ref is None or pos_face_ref is None:
        diameter = get_duct_diameter(duct)
        if diameter:
            use_round = True
            round_radius = diameter / 2.0
        else:
            try:
                neg_face_ref = axis_curve.Reference
                pos_face_ref = None
            except Exception:
                pass

    insertion_point = pick_insertion_point(doc, uidoc, view)
    if insertion_point is None:
        return

    grid_span = max(abs(neg_grid_dist or 0.0), abs(pos_grid_dist or 0.0))
    try:
        offset = abs(sub_xyz(insertion_point, center).DotProduct(perp_dir))
    except Exception:
        offset = 0.0
    line_length = max(grid_span + offset + 20.0, 30.0)
    line = build_dimension_line(view, insertion_point, perp_dir, length=line_length)

    t = Transaction(doc, "Quick Duct Dimension")
    t.Start()
    try:
        if use_round and round_radius:
            neg_face_ref, pos_face_ref = create_round_edge_refs(
                doc,
                view,
                center,
                axis_proj,
                perp_dir,
                round_radius,
                axis_curve.Length,
            )

        refs = []
        append_unique(refs, neg_grid_ref, doc)
        append_unique(refs, neg_face_ref, doc)
        append_unique(refs, pos_face_ref, doc)
        append_unique(refs, pos_grid_ref, doc)

        if len(refs) < 2:
            t.RollBack()
            forms.alert("Not enough valid references.", exitscript=True)
            return

        ref_array = ReferenceArray()
        for ref in refs:
            ref_array.Append(ref)

        dim = doc.Create.NewDimension(view, line, ref_array)
        try:
            dim.DimensionType = dim_type
        except Exception:
            pass

        t.Commit()
    except Exception:
        t.RollBack()
        raise


logger = script.get_logger()

try:
    main()
except Exception as ex:
    logger.error("QuickDuctDim failed: {0}".format(ex))
    logger.error(traceback.format_exc())
    forms.alert("Error: {0}".format(ex), title=TITLE)
