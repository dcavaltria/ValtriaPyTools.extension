# -*- coding: utf-8 -*-
"""
Check Interferences - Detector de colisiones MEP/Estructura
Cumplimiento RITE/CTE con exportación CSV para ACC
"""
__title__ = "Check\nInterferences"
__author__ = "Valtria"
__doc__ = "Detecta interferencias MEP-Estructura con validación normativa"
__cleanengine__ = True

import sys
import os

# Agregar lib al path
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import revit, DB, script, forms
from System.Collections.Generic import List as DotNetList

try:
    from valtria_utils import (
        mm_to_feet, feet_to_mm, 
        get_elements_by_categories,
        get_element_center,
        check_bbox_intersection
    )
except ImportError:
    # Fallback si lib no está disponible
    def mm_to_feet(mm):
        return mm / 304.8
    
    def feet_to_mm(feet):
        return feet * 304.8

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# Tolerancias según normativa
TOLERANCE_CRITICAL = 25  # mm - Colisión crítica
TOLERANCE_WARNING = 50   # mm - Advertencia
TOLERANCE_CHECK = 100    # mm - Revisar

# ============================================================================
# COLECTORES
# ============================================================================

def get_mep_elements_in_view():
    """Obtiene elementos MEP en vista activa"""
    view_id = doc.ActiveView.Id
    collector = DB.FilteredElementCollector(doc, view_id)
    
    mep_cats = [
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_DuctFitting,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_CableTray,
        DB.BuiltInCategory.OST_Conduit
    ]
    
    filters = DotNetList[DB.ElementFilter](
        [DB.ElementCategoryFilter(cat) for cat in mep_cats]
    )
    
    return list(
        collector
        .WherePasses(DB.LogicalOrFilter(filters))
        .WhereElementIsNotElementType()
        .ToElements()
    )


def get_structural_elements_in_view():
    """Obtiene elementos estructurales en vista activa"""
    view_id = doc.ActiveView.Id
    collector = DB.FilteredElementCollector(doc, view_id)
    
    struct_cats = [
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_StructuralFoundation,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Walls
    ]
    
    filters = DotNetList[DB.ElementFilter](
        [DB.ElementCategoryFilter(cat) for cat in struct_cats]
    )
    
    return list(
        collector
        .WherePasses(DB.LogicalOrFilter(filters))
        .WhereElementIsNotElementType()
        .ToElements()
    )


# ============================================================================
# ANÁLISIS INTERFERENCIAS
# ============================================================================

def analyze_clashes(mep_elements, struct_elements, tolerance_mm):
    """
    Detecta y clasifica interferencias
    
    Returns:
        dict con claves: 'critical', 'warning', 'check'
    """
    clashes = {'critical': [], 'warning': [], 'check': []}
    tolerance_ft = mm_to_feet(tolerance_mm)
    
    for mep_elem in mep_elements:
        bbox_mep = mep_elem.get_BoundingBox(None)
        if not bbox_mep:
            continue
        
        for struct_elem in struct_elements:
            bbox_struct = struct_elem.get_BoundingBox(None)
            if not bbox_struct:
                continue
            
            # Expandir bbox MEP con tolerancia
            bbox_mep_exp = DB.BoundingBoxXYZ()
            bbox_mep_exp.Min = DB.XYZ(
                bbox_mep.Min.X - tolerance_ft,
                bbox_mep.Min.Y - tolerance_ft,
                bbox_mep.Min.Z - tolerance_ft
            )
            bbox_mep_exp.Max = DB.XYZ(
                bbox_mep.Max.X + tolerance_ft,
                bbox_mep.Max.Y + tolerance_ft,
                bbox_mep.Max.Z + tolerance_ft
            )
            
            # Chequeo intersección
            if (bbox_mep_exp.Max.X > bbox_struct.Min.X and 
                bbox_mep_exp.Min.X < bbox_struct.Max.X and
                bbox_mep_exp.Max.Y > bbox_struct.Min.Y and 
                bbox_mep_exp.Min.Y < bbox_struct.Max.Y and
                bbox_mep_exp.Max.Z > bbox_struct.Min.Z and 
                bbox_mep_exp.Min.Z < bbox_struct.Max.Z):
                
                # Calcular distancia real
                center_mep = (bbox_mep.Min + bbox_mep.Max) / 2
                center_struct = (bbox_struct.Min + bbox_struct.Max) / 2
                distance_mm = feet_to_mm(center_mep.DistanceTo(center_struct))
                
                # Obtener nivel
                level_name = "N/A"
                try:
                    if mep_elem.LevelId and mep_elem.LevelId != DB.ElementId.InvalidElementId:
                        level = doc.GetElement(mep_elem.LevelId)
                        level_name = level.Name if level else "N/A"
                except:
                    pass
                
                # Obtener sistema MEP
                system_name = "N/A"
                try:
                    if hasattr(mep_elem, 'MEPSystem') and mep_elem.MEPSystem:
                        system_name = mep_elem.MEPSystem.Name
                except:
                    pass
                
                clash_data = {
                    'mep_elem': mep_elem,
                    'struct_elem': struct_elem,
                    'distance_mm': distance_mm,
                    'level': level_name,
                    'system': system_name
                }
                
                # Clasificar
                if distance_mm < TOLERANCE_CRITICAL:
                    clashes['critical'].append(clash_data)
                elif distance_mm < TOLERANCE_WARNING:
                    clashes['warning'].append(clash_data)
                else:
                    clashes['check'].append(clash_data)
    
    return clashes


# ============================================================================
# EXPORTACIÓN CSV
# ============================================================================

def export_clashes_to_csv(clashes_dict, filepath):
    """Exporta interferencias a CSV compatible con ACC"""
    import csv
    
    with open(filepath, 'wb') as f:
        writer = csv.writer(f, delimiter=';')
        
        # Headers
        writer.writerow([
            'Clash_ID',
            'Status',
            'MEP_ID',
            'MEP_Category',
            'MEP_System',
            'Struct_ID',
            'Struct_Category',
            'Distance_mm',
            'Level',
            'View',
            'Timestamp'
        ])
        
        clash_id = 1
        view_name = doc.ActiveView.Name
        timestamp = DB.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        
        # Críticas
        for clash in clashes_dict['critical']:
            writer.writerow([
                'CLASH_{:05d}'.format(clash_id),
                'CRITICAL',
                clash['mep_elem'].Id.IntegerValue,
                clash['mep_elem'].Category.Name,
                clash['system'],
                clash['struct_elem'].Id.IntegerValue,
                clash['struct_elem'].Category.Name,
                '{:.1f}'.format(clash['distance_mm']),
                clash['level'],
                view_name,
                timestamp
            ])
            clash_id += 1
        
        # Advertencias
        for clash in clashes_dict['warning']:
            writer.writerow([
                'CLASH_{:05d}'.format(clash_id),
                'WARNING',
                clash['mep_elem'].Id.IntegerValue,
                clash['mep_elem'].Category.Name,
                clash['system'],
                clash['struct_elem'].Id.IntegerValue,
                clash['struct_elem'].Category.Name,
                '{:.1f}'.format(clash['distance_mm']),
                clash['level'],
                view_name,
                timestamp
            ])
            clash_id += 1
        
        # Chequeos
        for clash in clashes_dict['check']:
            writer.writerow([
                'CLASH_{:05d}'.format(clash_id),
                'CHECK',
                clash['mep_elem'].Id.IntegerValue,
                clash['mep_elem'].Category.Name,
                clash['system'],
                clash['struct_elem'].Id.IntegerValue,
                clash['struct_elem'].Category.Name,
                '{:.1f}'.format(clash['distance_mm']),
                clash['level'],
                view_name,
                timestamp
            ])
            clash_id += 1


# ============================================================================
# UI & REPORTE
# ============================================================================

def print_clash_report(clashes_dict):
    """Genera reporte visual en output"""
    
    output.print_md("# 🔍 Valtria Clash Detection")
    output.print_md("---")
    
    total = len(clashes_dict['critical']) + len(clashes_dict['warning']) + len(clashes_dict['check'])
    
    # Resumen
    output.print_md("## 📊 Resumen")
    output.print_md("| Status | Cantidad | Tolerancia |")
    output.print_md("|--------|----------|------------|")
    output.print_md("| 🔴 **CRÍTICO** | {} | < {}mm |".format(
        len(clashes_dict['critical']), TOLERANCE_CRITICAL))
    output.print_md("| 🟡 **ADVERTENCIA** | {} | < {}mm |".format(
        len(clashes_dict['warning']), TOLERANCE_WARNING))
    output.print_md("| 🔵 **REVISAR** | {} | < {}mm |".format(
        len(clashes_dict['check']), TOLERANCE_CHECK))
    output.print_md("| **TOTAL** | **{}** | |".format(total))
    
    # Detalle críticas
    if clashes_dict['critical']:
        output.print_md("\n## 🔴 Interferencias Críticas")
        for clash in clashes_dict['critical'][:10]:  # Mostrar máximo 10
            output.print_md(
                "- **{}** {} ↔ **{}** {} | {:.0f}mm | Nivel: {} | Sistema: {}".format(
                    clash['mep_elem'].Category.Name,
                    output.linkify(clash['mep_elem'].Id),
                    clash['struct_elem'].Category.Name,
                    output.linkify(clash['struct_elem'].Id),
                    clash['distance_mm'],
                    clash['level'],
                    clash['system']
                )
            )
        
        if len(clashes_dict['critical']) > 10:
            output.print_md("\n*... y {} más críticas*".format(
                len(clashes_dict['critical']) - 10))
    
    # Detalle advertencias (primeras 5)
    if clashes_dict['warning']:
        output.print_md("\n## 🟡 Advertencias (primeras 5)")
        for clash in clashes_dict['warning'][:5]:
            output.print_md(
                "- **{}** {} ↔ **{}** {} | {:.0f}mm".format(
                    clash['mep_elem'].Category.Name,
                    output.linkify(clash['mep_elem'].Id),
                    clash['struct_elem'].Category.Name,
                    output.linkify(clash['struct_elem'].Id),
                    clash['distance_mm']
                )
            )


# ============================================================================
# MAIN
# ============================================================================

def main():
    """Función principal"""
    
    # Verificar vista activa
    if not doc.ActiveView:
        forms.alert('No hay vista activa', exitscript=True)
    
    # Recolección
    output.print_md("🔍 **Recolectando elementos en vista:** {}".format(doc.ActiveView.Name))
    
    mep_elems = get_mep_elements_in_view()
    struct_elems = get_structural_elements_in_view()
    
    output.print_md("- MEP: **{}** elementos".format(len(mep_elems)))
    output.print_md("- Estructura: **{}** elementos".format(len(struct_elems)))
    
    if not mep_elems:
        forms.alert('No hay elementos MEP en la vista activa', exitscript=True)
    
    if not struct_elems:
        forms.alert('No hay elementos estructurales en la vista activa', exitscript=True)
    
    # Análisis
    output.print_md("\n⚙️ **Analizando interferencias...**")
    
    clashes = analyze_clashes(mep_elems, struct_elems, TOLERANCE_CHECK)
    
    # Reporte
    print_clash_report(clashes)
    
    # Exportación
    total_clashes = sum(len(v) for v in clashes.values())
    if total_clashes > 0:
        if forms.alert('¿Exportar {} interferencias a CSV?'.format(total_clashes), 
                      yes=True, no=True):
            filepath = forms.save_file(
                file_ext='csv',
                default_name='ValtriaCLASHES_{}'.format(doc.ActiveView.Name)
            )
            if filepath:
                export_clashes_to_csv(clashes, filepath)
                output.print_md("\n✅ **Exportado correctamente:** `{}`".format(filepath))
                forms.alert('CSV exportado con éxito', title='Valtria Tools')
    
    output.print_md("\n---")
    output.print_md("*Valtria BIM Tools v1.0 - Cumplimiento RITE/CTE*")


# ============================================================================
# EJECUCIÓN
# ============================================================================

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        output.print_md("## ❌ Error")
        output.print_md("```\n{}\n```".format(str(e)))
        
        import traceback
        output.print_md("### Stack Trace")
        output.print_md("```python\n{}\n```".format(traceback.format_exc()))
        
        forms.alert('Error en ejecución. Ver output para detalles.', 
                   title='Valtria Tools - Error')