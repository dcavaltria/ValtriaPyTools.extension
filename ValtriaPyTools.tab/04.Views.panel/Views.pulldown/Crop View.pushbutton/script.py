# -*- coding: utf-8 -*-
"""Crea un marco de lineas de detalle siguiendo el crop de vistas en una hoja."""

import os
import sys
import math

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..'))
LIB_DIR = os.path.join(EXTENSION_DIR, 'lib')
for _path in (EXTENSION_DIR, LIB_DIR):
    if _path and _path not in sys.path:
        sys.path.insert(0, _path)

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    GraphicsStyleType,
    Line,
    Transaction,
    Transform,
    ViewSheet,
    ViewType,
    Viewport,
    XYZ,
)
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import forms

from valtria_lib import get_doc, get_uidoc, log_exception


TITLE = "Marco por Crop"
MAX_LISTED = 20
ANNOTATION_OFFSET_CM = 1.0
CM_TO_FEET = 0.03280839895

try:
    unicode
except NameError:
    unicode = str  # type: ignore


def safe_text(value):
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u""


class ViewportSelectionFilter(ISelectionFilter):
    """Permite seleccionar unicamente viewports."""

    def AllowElement(self, element):
        try:
            return isinstance(element, Viewport)
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


class StyleOption(object):
    """Envuelve un estilo para usarlo en SelectFromList."""

    def __init__(self, style, name, is_default=False):
        self.value = style
        self.raw_name = name
        self._is_default = is_default

    @property
    def name(self):
        label = safe_text(self.raw_name) or u"(sin nombre)"
        if self._is_default:
            label = u"{0}  [Predeterminado]".format(label)
        return label


def iterate(net_collection):
    """Itera sobre colecciones .NET sin lanzar excepciones."""
    if net_collection is None:
        return
    try:
        for item in net_collection:
            yield item
        return
    except Exception:
        pass
    try:
        iterator = net_collection.GetEnumerator()
    except Exception:
        iterator = None
    if iterator is None:
        return
    while iterator.MoveNext():
        yield iterator.Current


def collect_selected_viewports(doc, uidoc):
    ids = []
    try:
        ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        ids = []
    viewports = []
    seen = set()
    for eid in ids:
        if eid in seen:
            continue
        seen.add(eid)
        element = doc.GetElement(eid)
        if isinstance(element, Viewport):
            viewports.append(element)
    return viewports


def prompt_viewports(doc, uidoc):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            ViewportSelectionFilter(),
            "Selecciona los viewports en la hoja",
        )
    except Selection.OperationCanceledException:
        return []
    except Exception:
        return []
    viewports = []
    for ref in refs:
        element = doc.GetElement(ref.ElementId)
        if isinstance(element, Viewport):
            viewports.append(element)
    return viewports


def extract_views(doc, viewports):
    views = []
    seen = set()
    for viewport in viewports:
        try:
            view = doc.GetElement(viewport.ViewId)
        except Exception:
            view = None
        if view is None or getattr(view, "IsTemplate", False):
            continue
        vid = getattr(getattr(view, "Id", None), "IntegerValue", None)
        if vid in seen:
            continue
        seen.add(vid)
        views.append(view)
    return views


def map_view_to_sheet_label(doc, viewports):
    mapping = {}
    for viewport in viewports:
        view_id_int = getattr(getattr(viewport, "ViewId", None), "IntegerValue", None)
        if view_id_int is None:
            continue
        try:
            sheet = doc.GetElement(viewport.SheetId)
        except Exception:
            sheet = None
        number = safe_text(getattr(sheet, "SheetNumber", u"")).strip() if sheet else u""
        name = safe_text(getattr(sheet, "Name", u"")).strip() if sheet else u""
        label = u"{0} | {1}".format(number, name).strip(" |")
        mapping[view_id_int] = label
    return mapping


def collect_line_styles(doc):
    styles = []
    try:
        cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
    except Exception:
        cat = None
    if cat is None:
        return styles
    for subcat in iterate(getattr(cat, "SubCategories", None)):
        if subcat is None:
            continue
        name = safe_text(getattr(subcat, "Name", u""))
        try:
            gs = subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
        except Exception:
            gs = None
        if gs is None:
            continue
        styles.append((name, gs))
    styles.sort(key=lambda pair: safe_text(pair[0]).lower())
    return styles


def find_default_style(styles):
    preferred = set(
        [
            "<centerline>",
            "centerline",
            "center line",
            "center",
            "centro",
            "centrolinea",
        ]
    )
    for name, gs in styles:
        norm = safe_text(name).strip().lower()
        simplified = norm.replace("<", "").replace(">", "")
        if norm in preferred or simplified in preferred:
            return gs, name
    return None, None


def prompt_style(styles, default_style):
    default_id = getattr(getattr(default_style, "Id", None), "IntegerValue", None) if default_style else None
    items = []
    for name, gs in styles:
        is_default = False
        gs_id = getattr(getattr(gs, "Id", None), "IntegerValue", None)
        if default_id is not None and gs_id == default_id:
            is_default = True
        items.append(StyleOption(gs, name, is_default=is_default))
    picked = forms.SelectFromList.show(
        items,
        title=TITLE + " - Estilo de linea",
        multiselect=False,
        button_name="Usar estilo",
        name_attr="name",
    )
    if not picked:
        return None, None
    style = picked.value if hasattr(picked, "value") else picked
    style_name = safe_text(getattr(picked, "raw_name", None))
    if not style_name:
        style_name = safe_text(getattr(style, "Name", u""))
    return style, style_name


def view_accepts_detail_lines(view):
    vt = getattr(view, "ViewType", None)
    disallowed = set(
        [
            ViewType.ThreeD,
            ViewType.Schedule,
            ViewType.DrawingSheet,
            ViewType.ProjectBrowser,
            ViewType.SystemBrowser,
        ]
    )
    return vt not in disallowed


def clone_curve(curve):
    if curve is None:
        return None
    for attr in ("Clone",):
        method = getattr(curve, attr, None)
        if callable(method):
            try:
                return method()
            except Exception:
                pass
    try:
        return curve.CreateTransformed(Transform.Identity)
    except Exception:
        pass
    try:
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        return Line.CreateBound(start, end)
    except Exception:
        return None


def curve_points(curve):
    """Devuelve puntos representativos de una curva para calcular limites."""
    pts = []
    try:
        tess = curve.Tessellate()
        if tess:
            pts.extend(list(tess))
    except Exception:
        pass
    for idx in (0, 1):
        try:
            pt = curve.GetEndPoint(idx)
            if pt:
                pts.append(pt)
        except Exception:
            pass
    return pts


def bounds_from_curves(curves):
    """Calcula un bounding box simple en XY a partir de una lista de curvas."""
    minx = miny = minz = float("inf")
    maxx = maxy = maxz = float("-inf")
    for curve in curves:
        for pt in curve_points(curve):
            minx = min(minx, pt.X)
            miny = min(miny, pt.Y)
            minz = min(minz, pt.Z)
            maxx = max(maxx, pt.X)
            maxy = max(maxy, pt.Y)
            maxz = max(maxz, pt.Z)
    if minx == float("inf"):
        return None, None
    return XYZ(minx, miny, minz), XYZ(maxx, maxy, maxz)


def transform_curve_to_sheet(curve, transform):
    """Transforma una curva y ofrece un fallback discreto cuando falla el API."""
    is_closed = False
    try:
        is_closed = bool(getattr(curve, "IsClosed", False))
    except Exception:
        is_closed = False

    def is_flat_on_sheet(test_curve):
        try:
            start = test_curve.GetEndPoint(0)
            end = test_curve.GetEndPoint(1)
            if start is None or end is None:
                return False
            return abs(start.Z) < 1e-6 and abs(end.Z) < 1e-6 and abs(start.Z - end.Z) < 1e-6
        except Exception:
            return False

    try:
        cloned = clone_curve(curve)
        if cloned is not None:
            direct = cloned.CreateTransformed(transform)
            if direct is not None and is_flat_on_sheet(direct):
                return [direct]
    except Exception:
        pass

    points = []
    try:
        tess = curve.Tessellate()
        if tess:
            points.extend(list(tess))
    except Exception:
        pass
    if not points:
        for idx in (0, 1):
            try:
                pt = curve.GetEndPoint(idx)
                if pt:
                    points.append(pt)
            except Exception:
                continue
    transformed_points = []
    for pt in points:
        try:
            tpt = transform.OfPoint(pt)
            if tpt is not None:
                transformed_points.append(XYZ(tpt.X, tpt.Y, 0.0))
        except Exception:
            continue
    if len(transformed_points) < 2:
        return []

    curves = []
    for idx in range(len(transformed_points) - 1):
        start = transformed_points[idx]
        end = transformed_points[idx + 1]
        try:
            if start.DistanceTo(end) < 1e-9:
                continue
        except Exception:
            pass
        try:
            line = Line.CreateBound(start, end)
            if line is not None:
                curves.append(line)
        except Exception:
            continue
    if is_closed:
        try:
            first_pt = transformed_points[0]
            last_pt = transformed_points[-1]
            if first_pt and last_pt and not first_pt.IsAlmostEqualTo(last_pt):
                closing = Line.CreateBound(last_pt, first_pt)
                if closing is not None:
                    curves.append(closing)
        except Exception:
            pass
    return curves


def viewport_center(viewport):
    """Obtiene el centro del viewport en coordenadas de hoja."""
    try:
        outline = viewport.GetBoxOutline()
        if outline:
            min_pt = outline.MinimumPoint
            max_pt = outline.MaximumPoint
            return XYZ(
                (min_pt.X + max_pt.X) / 2.0,
                (min_pt.Y + max_pt.Y) / 2.0,
                (min_pt.Z + max_pt.Z) / 2.0,
            )
    except Exception:
        pass
    getter = getattr(viewport, "GetBoxCenter", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            return None
    return None


def rotation_to_angle(rotation):
    """Convierte la enumeracion de rotacion del viewport en un angulo en radianes."""
    try:
        from Autodesk.Revit.DB import ViewportRotation

        none_val = getattr(ViewportRotation, "None", None)
        cw_val = getattr(ViewportRotation, "Clockwise", getattr(ViewportRotation, "Right", None))
        ccw_val = getattr(ViewportRotation, "Counterclockwise", getattr(ViewportRotation, "Left", None))
        up_val = getattr(ViewportRotation, "Up", None)
        down_val = getattr(ViewportRotation, "Down", None)
        angle180_val = getattr(ViewportRotation, "Rotate180", None)

        if rotation == none_val:
            return 0.0
        if rotation == cw_val:
            return -math.pi / 2.0
        if rotation == ccw_val:
            return math.pi / 2.0
        if rotation == up_val:
            return 0.0
        if rotation == down_val:
            return math.pi
        if rotation == angle180_val:
            return math.pi
    except Exception:
        pass

    text = safe_text(rotation).lower()
    if "180" in text or "upside" in text or "half" in text:
        return math.pi
    if "270" in text:
        return -math.pi / 2.0
    if "90" in text or "left" in text or "counter" in text or "ccw" in text:
        return math.pi / 2.0
    if "clock" in text or "right" in text or "cw" in text:
        return -math.pi / 2.0
    return 0.0


def build_transform_fallback(viewport, view, curves):
    """Construye una transformacion view->sheet cuando GetTransform no existe."""
    center_sheet = viewport_center(viewport)
    if center_sheet is None:
        return None

    min_pt, max_pt = bounds_from_curves(curves) if curves else (None, None)
    if min_pt is None or max_pt is None:
        try:
            crop = view.CropBox
            if crop:
                min_pt = crop.Min
                max_pt = crop.Max
        except Exception:
            min_pt = max_pt = None
    if min_pt is None or max_pt is None:
        min_pt = XYZ(-0.5, -0.5, 0.0)
        max_pt = XYZ(0.5, 0.5, 0.0)

    view_center = XYZ(
        (min_pt.X + max_pt.X) / 2.0,
        (min_pt.Y + max_pt.Y) / 2.0,
        (min_pt.Z + max_pt.Z) / 2.0,
    )

    try:
        view_scale = float(getattr(view, "Scale", 1) or 1)
    except Exception:
        view_scale = 1.0
    view_scale = view_scale if view_scale else 1.0
    scale = 1.0 / view_scale

    angle = rotation_to_angle(getattr(viewport, "Rotation", None))
    cos_a = math.cos(angle)
    sin_a = math.sin(angle)

    basis_x = XYZ(cos_a * scale, sin_a * scale, 0.0)
    basis_y = XYZ(-sin_a * scale, cos_a * scale, 0.0)
    basis_z = XYZ(0.0, 0.0, 1.0)

    origin = XYZ(
        center_sheet.X - (basis_x.X * view_center.X + basis_y.X * view_center.Y),
        center_sheet.Y - (basis_x.Y * view_center.X + basis_y.Y * view_center.Y),
        center_sheet.Z - (basis_x.Z * view_center.X + basis_y.Z * view_center.Y),
    )

    transform = Transform.Identity
    transform.Origin = origin
    transform.BasisX = basis_x
    transform.BasisY = basis_y
    transform.BasisZ = basis_z
    return transform


def get_viewport_transform(viewport, view, curves):
    """Obtiene una transformacion de vista a hoja compatible con varias versiones."""
    for method_name in ("GetTransform", "GetBoxTransform", "GetViewToSheetTransform"):
        method = getattr(viewport, method_name, None)
        if callable(method):
            try:
                transform = method()
                if transform:
                    return transform
            except Exception:
                continue

    try:
        outline = viewport.GetBoxOutline()
    except Exception:
        outline = None
    if outline:
        min_pt = outline.MinimumPoint
        max_pt = outline.MaximumPoint
        min_view, max_view = bounds_from_curves(curves) if curves else (None, None)
        if min_view is not None and max_view is not None:
            view_center = XYZ(
                (min_view.X + max_view.X) / 2.0,
                (min_view.Y + max_view.Y) / 2.0,
                (min_view.Z + max_view.Z) / 2.0,
            )
            sheet_center = XYZ(
                (min_pt.X + max_pt.X) / 2.0,
                (min_pt.Y + max_pt.Y) / 2.0,
                (min_pt.Z + max_pt.Z) / 2.0,
            )
            width_view = max(1e-9, max_view.X - min_view.X)
            height_view = max(1e-9, max_view.Y - min_view.Y)
            width_sheet = max(1e-9, max_pt.X - min_pt.X)
            height_sheet = max(1e-9, max_pt.Y - min_pt.Y)

            angle = rotation_to_angle(getattr(viewport, "Rotation", None))
            cos_a = math.cos(angle)
            sin_a = math.sin(angle)
            if abs(cos_a) < abs(sin_a):
                scale_x = width_sheet / height_view
                scale_y = height_sheet / width_view
            else:
                scale_x = width_sheet / width_view
                scale_y = height_sheet / height_view
            basis_x = XYZ(cos_a * scale_x, sin_a * scale_x, 0.0)
            basis_y = XYZ(-sin_a * scale_y, cos_a * scale_y, 0.0)
            basis_z = XYZ(0.0, 0.0, 1.0)

            origin = XYZ(
                sheet_center.X - (basis_x.X * view_center.X + basis_y.X * view_center.Y),
                sheet_center.Y - (basis_x.Y * view_center.X + basis_y.Y * view_center.Y),
                sheet_center.Z - (basis_x.Z * view_center.X + basis_y.Z * view_center.Y),
            )

            transform = Transform.Identity
            transform.Origin = origin
            transform.BasisX = basis_x
            transform.BasisY = basis_y
            transform.BasisZ = basis_z
            return transform

    return build_transform_fallback(viewport, view, curves)


def get_crop_curves(view):
    manager = None
    getter = getattr(view, "GetCropRegionShapeManager", None)
    if getter is None:
        return []
    try:
        manager = getter()
    except Exception:
        manager = None
    if manager is None:
        return []
    loops = []
    for method_name in ("GetCropRegionShape", "GetCropShape"):
        method = getattr(manager, method_name, None)
        if not callable(method):
            continue
        try:
            shape = method()
        except Exception:
            shape = None
        if shape is None:
            continue
        try:
            for loop in shape:
                if loop:
                    loops.append(loop)
        except Exception:
            if shape:
                loops.append(shape)
        if loops:
            break
    if not loops:
        return []
    for loop in loops:
        curves = []
        for curve in iterate(loop):
            cloned = clone_curve(curve)
            if cloned is not None:
                curves.append(cloned)
        if curves:
            return curves
    return []


def crop_box_outline_curves(view):
    """Crea un rectangulo a partir del CropBox (coordenadas de modelo)."""
    try:
        crop = getattr(view, "CropBox", None)
        if crop is None:
            return []
        transform = getattr(crop, "Transform", None)
        if transform is None:
            return []
        min_pt = getattr(crop, "Min", None)
        max_pt = getattr(crop, "Max", None)
        if min_pt is None or max_pt is None:
            return []
        corners_local = [
            XYZ(min_pt.X, min_pt.Y, min_pt.Z),
            XYZ(max_pt.X, min_pt.Y, min_pt.Z),
            XYZ(max_pt.X, max_pt.Y, min_pt.Z),
            XYZ(min_pt.X, max_pt.Y, min_pt.Z),
        ]
        corners_model = []
        for pt in corners_local:
            try:
                corners_model.append(transform.OfPoint(pt))
            except Exception:
                return []
        curves = []
        for idx in range(len(corners_model)):
            start = corners_model[idx]
            end = corners_model[(idx + 1) % len(corners_model)]
            try:
                line = Line.CreateBound(start, end)
                if line is not None:
                    curves.append(line)
            except Exception:
                return []
        return curves
    except Exception:
        return []


def cm_to_feet(cm_value):
    try:
        return float(cm_value) * CM_TO_FEET
    except Exception:
        return 0.0


def ensure_annotation_offset(view, offset_ft):
    """Activa el recorte de anotaciones y aplica un desfase."""
    updated = False
    try:
        if hasattr(view, "AnnotationCropActive"):
            if not view.AnnotationCropActive:
                view.AnnotationCropActive = True
                updated = True
    except Exception:
        pass
    try:
        param_active = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
        if param_active and not param_active.IsReadOnly:
            if param_active.AsInteger() == 0:
                param_active.Set(1)
                updated = True
    except Exception:
        pass
    try:
        offset_param = view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_OFFSET)
        if offset_param and not offset_param.IsReadOnly:
            offset_param.Set(offset_ft)
            updated = True
    except Exception:
        pass
    return updated


def viewport_label(doc, viewport, view):
    sheet_label = u""
    try:
        sheet = doc.GetElement(viewport.SheetId)
    except Exception:
        sheet = None
    if sheet:
        number = safe_text(getattr(sheet, "SheetNumber", u"")).strip()
        name = safe_text(getattr(sheet, "Name", u"")).strip()
        sheet_label = u"{0} | {1}".format(number, name).strip(" |")
    return format_view_label(view, sheet_label)


def format_view_label(view, sheet_label):
    view_name = safe_text(getattr(view, "Name", u"")).strip() or u"(sin nombre)"
    if sheet_label:
        return u"{0} -> {1}".format(sheet_label, view_name)
    return view_name


def confirm_targets(labels, style_name):
    lines = []
    lines.append(u"Viewports a enmarcar: {0}".format(len(labels)))
    for idx, label in enumerate(labels):
        if idx >= MAX_LISTED:
            lines.append(u"... ({0} viewports adicionales)".format(len(labels) - MAX_LISTED))
            break
        lines.append(u"- {0}".format(label))
    lines.append(u"Estilo de linea: {0}".format(safe_text(style_name) or u"(sin nombre)"))
    lines.append(u"Las lineas se dibujaran en la hoja activa alrededor del viewport.")
    lines.append(u"Se ocultara 'Crop View Visible' y se aplicara un desfase de anotaciones de 1 cm.")
    lines.append(u"Aplicar el marco?")
    return forms.alert(u"\n".join(lines), title=TITLE, yes=True, no=True, warn_icon=False)


def summarize(created, skipped, failed):
    lines = []
    lines.append(u"Marcos creados: {0}".format(len(created)))
    for idx, label in enumerate(created):
        if idx >= MAX_LISTED:
            lines.append(u"... ({0} vistas adicionales)".format(len(created) - MAX_LISTED))
            break
        lines.append(u"- {0}".format(label))
    if skipped:
        lines.append(u"")
        lines.append(u"Omitidas: {0}".format(len(skipped)))
        for label, reason in skipped:
            lines.append(u"- {0} ({1})".format(label, reason))
    if failed:
        lines.append(u"")
        lines.append(u"Errores: {0}".format(len(failed)))
        for label, reason in failed:
            lines.append(u"- {0}: {1}".format(label, reason))
    return lines


def main():
    doc = get_doc()
    uidoc = get_uidoc()
    if doc is None or uidoc is None:
        forms.alert("No hay documento activo.", title=TITLE)
        return

    active_view = uidoc.ActiveView
    if not isinstance(active_view, ViewSheet):
        forms.alert("Abre una hoja y selecciona los viewports a enmarcar.", title=TITLE)
        return

    viewports = collect_selected_viewports(doc, uidoc)
    if not viewports:
        viewports = prompt_viewports(doc, uidoc)
    if not viewports:
        forms.alert("No se seleccionaron viewports.", title=TITLE)
        return

    line_styles = collect_line_styles(doc)
    if not line_styles:
        forms.alert("No hay estilos de lineas de detalle disponibles en el modelo.", title=TITLE)
        return

    default_style, default_name = find_default_style(line_styles)
    selected_style = None
    selected_style_name = None

    if default_style is not None:
        use_default = forms.alert(
            u"Usar el estilo predeterminado '{0}'?".format(safe_text(default_name)),
            title=TITLE,
            yes=True,
            no=True,
            warn_icon=False,
        )
        if use_default:
            selected_style = default_style
            selected_style_name = default_name

    if selected_style is None:
        selected_style, selected_style_name = prompt_style(line_styles, default_style)
        if selected_style is None:
            forms.alert("Operacion cancelada. No se eligio estilo de linea.", title=TITLE)
            return

    targets = []
    labels = []
    for viewport in viewports:
        try:
            if viewport.SheetId != active_view.Id:
                continue
        except Exception:
            continue
        try:
            view = doc.GetElement(viewport.ViewId)
        except Exception:
            view = None
        if view is None or getattr(view, "IsTemplate", False):
            continue
        label = viewport_label(doc, viewport, view)
        targets.append((viewport, view, label))
        labels.append(label)

    if not targets:
        forms.alert("No se encontraron viewports validos en la hoja activa.", title=TITLE)
        return

    if not confirm_targets(labels, selected_style_name):
        forms.alert("Operacion cancelada por el usuario.", title=TITLE)
        return

    created = []
    skipped = []
    failed = []

    transaction = Transaction(doc, TITLE)
    transaction.Start()
    try:
        annotation_offset_ft = cm_to_feet(ANNOTATION_OFFSET_CM)
        for viewport, view, label in targets:
            try:
                if getattr(view, "IsTemplate", False):
                    skipped.append((label, "La vista es una plantilla."))
                    continue
            except Exception:
                pass

            try:
                if not getattr(view, "CropBoxActive", True):
                    view.CropBoxActive = True
            except Exception:
                skipped.append((label, "El recorte no esta activo y no se pudo habilitar."))
                continue

            curves = get_crop_curves(view)
            if not curves:
                curves = crop_box_outline_curves(view)
            if not curves:
                failed.append((label, "No se pudo obtener el contorno del crop."))
                continue

            try:
                transform = get_viewport_transform(viewport, view, curves)
            except Exception as err:
                failed.append((label, safe_text(err)))
                continue

            if transform is None:
                failed.append((label, "No se pudo obtener la transformacion del viewport."))
                continue

            transformed_curves = []
            ok = True
            for curve in curves:
                partial = transform_curve_to_sheet(curve, transform)
                if not partial:
                    alt_curves = crop_box_outline_curves(view)
                    partial = []
                    for alt_curve in alt_curves:
                        partial.extend(transform_curve_to_sheet(alt_curve, transform))
                if not partial:
                    ok = False
                    failed.append((label, "No se pudo transformar el contorno al espacio de la hoja."))
                    break
                transformed_curves.extend(partial)

            if not ok:
                continue

            for transformed_curve in transformed_curves:
                try:
                    detail = doc.Create.NewDetailCurve(active_view, transformed_curve)
                    try:
                        detail.LineStyle = selected_style
                    except Exception:
                        pass
                except Exception as err:
                    ok = False
                    failed.append((label, safe_text(err)))
                    break

            if ok:
                try:
                    view.CropBoxVisible = False
                except Exception:
                    pass
                try:
                    ensure_annotation_offset(view, annotation_offset_ft)
                except Exception:
                    pass
                created.append(label)

        transaction.Commit()
    except Exception as err:
        transaction.RollBack()
        log_exception(err)
        forms.alert("Error inesperado:\n{0}".format(safe_text(err)), title=TITLE)
        return

    summary = summarize(created, skipped, failed)
    forms.alert(u"\n".join(summary), title=TITLE, warn_icon=bool(failed))


if __name__ == '__main__':
    try:
        main()
    except Exception as main_error:
        log_exception(main_error)

