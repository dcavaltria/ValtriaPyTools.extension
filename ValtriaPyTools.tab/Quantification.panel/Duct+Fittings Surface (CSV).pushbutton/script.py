# -*- coding: utf-8 -*-
# pyRevit | IronPython 2.x
# Duct + Fittings Surface (CSV): cuantificación por superficie [m²] sin crear parámetros (UNE/ANFACA).
# - Ducts: Perímetro * (Longitud + UT)
# - Fittings: Perímetro_boca_mayor * (Leq + UT) con mínimo 1,00 m²/pieza
# - UT global o por sistema (CSV)
# - Leq por tipo de pieza (CSV) o heurística si no hay valor

import os
import csv
import math
import sys
import traceback

from pyrevit import forms, script

import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    StorageType,
    Domain,
    ConnectorProfileType
)

try:
    clr.AddReference('PresentationFramework')
    clr.AddReference('PresentationCore')
    from System.Windows import Window
    from System.Windows.Markup import XamlReader
    from System.IO import StringReader
except Exception:
    forms.alert("No se pudo cargar WPF (.NET). Usa Revit 2019+.", exitscript=True)

XAML = r"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Duct + Fittings Surface (CSV) – Valtria"
        Height="470" Width="650"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <GroupBox Header="Unión Transversal (UT)" Grid.Row="0" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="140"/>
          <ColumnDefinition Width="220"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="UT global:" VerticalAlignment="Center"/>
        <ComboBox Name="cmbUT" Grid.Column="1" Height="26" Margin="8,0,8,0">
          <ComboBoxItem Content="vaina"/>
          <ComboBoxItem Content="perfil20"/>
          <ComboBoxItem Content="perfil30"/>
          <ComboBoxItem Content="otras"/>
        </ComboBox>
        <TextBlock Grid.Column="2" VerticalAlignment="Center"
                   Text="Si hay mapa por sistema, este valor es fallback." />
      </Grid>
    </GroupBox>

    <GroupBox Header="UT por Sistema (CSV opcional)" Grid.Row="1" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="140"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <CheckBox Name="chkUTMap" Grid.Row="0" Grid.ColumnSpan="3" Content="Usar CSV de UT por sistema" Margin="0,0,0,6"/>
        <TextBlock Text="CSV UT mapa:" Grid.Row="1" VerticalAlignment="Center"/>
        <TextBox Name="txtUTCsv" Grid.Row="1" Grid.Column="1" Height="24" Margin="8,0,8,0"/>
        <Button Name="btnUTCsv" Grid.Row="1" Grid.Column="2" Content="..." Width="28" Height="24"/>
      </Grid>
    </GroupBox>

    <GroupBox Header="Leq por Tipo de Fitting (CSV opcional, metros)" Grid.Row="2" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="200"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <CheckBox Name="chkLeqMap" Grid.Row="0" Grid.ColumnSpan="3" Content="Usar CSV de Leq por tipo" Margin="0,0,0,6"/>
        <TextBlock Text="CSV Leq mapa:" Grid.Row="1" VerticalAlignment="Center"/>
        <TextBox Name="txtLeqCsv" Grid.Row="1" Grid.Column="1" Height="24" Margin="8,0,8,0"/>
        <Button Name="btnLeqCsv" Grid.Row="1" Grid.Column="2" Content="..." Width="28" Height="24"/>

        <CheckBox Name="chkMin1" Grid.Row="2" Grid.ColumnSpan="3" Content="Aplicar mínimo 1,00 m² por pieza (fittings)" Margin="0,6,0,0" IsChecked="True"/>
      </Grid>
    </GroupBox>

    <GroupBox Header="Filtros (opcionales)" Grid.Row="3" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="160"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="160"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Sistema contiene:" VerticalAlignment="Center"/>
        <TextBox Name="txtFilterSys" Grid.Column="1" Height="24" Margin="8,0,8,0"/>
        <TextBlock Text="Fase contiene:" Grid.Column="2" VerticalAlignment="Center"/>
        <TextBox Name="txtFilterPh" Grid.Column="3" Height="24" Margin="8,0,0,0"/>
      </Grid>
    </GroupBox>

    <GroupBox Header="Exportación CSV" Grid.Row="4" Margin="0,0,0,10">
      <Grid Margin="8">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="140"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Carpeta destino:" VerticalAlignment="Center"/>
        <TextBox Name="txtOutDir" Grid.Column="1" Height="24" Margin="8,0,8,0"/>
        <Button Name="btnOutDir" Grid.Column="2" Content="..." Width="28" Height="24"/>
        <TextBlock Text="Prefijo archivo:" Grid.Row="1" VerticalAlignment="Center"/>
        <TextBox Name="txtPrefix" Grid.Row="1" Grid.Column="1" Height="24" Margin="8,0,8,0"/>
      </Grid>
    </GroupBox>

    <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnOk" Content="Calcular y exportar" Width="160" Height="26" Margin="0,0,8,0"/>
      <Button Name="btnCancel" Content="Cancelar" Width="100" Height="26"/>
    </StackPanel>
  </Grid>
</Window>
"""

def _load_wpf():
    sr = StringReader(XAML)
    return XamlReader.Load(sr)

FT_TO_M = 0.3048

def _as_utf8(x):
    if isinstance(x, unicode):
        return x.encode('utf-8')
    return x

from Autodesk.Revit.DB import ConnectorProfileType

def get_dbl(e, bip):
    try:
        p = e.get_Parameter(bip)
        if p and p.HasValue and p.StorageType == StorageType.Double:
            return p.AsDouble()
    except:
        pass
    return None

def get_str(e, bip, fallbacks=None):
    try:
        p = e.get_Parameter(bip)
        if p and p.HasValue:
            s = p.AsString()
            if s: return s
    except:
        pass
    if fallbacks:
        for n in fallbacks:
            try:
                q = e.LookupParameter(n)
                if q and q.HasValue:
                    s = q.AsString()
                    if s: return s
            except:
                pass
    return ""

def system_name(elem):
    s = get_str(elem, BuiltInParameter.RBS_SYSTEM_NAME_PARAM, ["System","System Name","Sistema"])
    return s or ""

def level_name(elem):
    try:
        lid = elem.LevelId
        if lid and lid.IntegerValue > 0:
            lv = elem.Document.GetElement(lid)
            if lv: return lv.Name
    except:
        pass
    return ""

def phase_name(elem):
    try:
        ph = elem.Document.GetElement(elem.CreatedPhaseId)
        return ph.Name if ph else ""
    except:
        return ""

def perim_rect(w_m, h_m):
    return max(0.0, 2.0*(max(0.0,w_m)+max(0.0,h_m)))

def perim_circ(d_m):
    return max(0.0, math.pi*max(0.0,d_m))

def perim_oval(a_m, b_m):
    a = max(0.0,a_m); b = max(0.0,b_m)
    if a == 0.0 and b == 0.0: return 0.0
    return math.pi*(3*(a+b) - math.sqrt((3*a+b)*(a+3*b)))

def ut_value(perim_m, ut_mode_or_val):
    try:
        if isinstance(ut_mode_or_val, float) or isinstance(ut_mode_or_val, int):
            return max(0.0, float(ut_mode_or_val))
        m = (ut_mode_or_val or "").strip().lower()
        if m == "vaina":    return 0.024
        if m == "perfil20": return 0.120
        if m == "perfil30": return 0.170
        if m == "otras":    return max(perim_m, 0.0)
    except:
        pass
    return 0.120

def read_map_csv(path):
    m = {}
    if not path or not os.path.isfile(path):
        return m
    try:
        with open(path, 'rb') as f:
            data = f.read().decode('utf-8').splitlines()
        rdr = csv.reader(data, delimiter=';')
        for row in rdr:
            if not row or len(row) < 2:
                continue
            k = row[0].strip()
            v = row[1].strip()
            if not k:
                continue
            try:
                vnum = float(v.replace(',', '.'))
                m[k] = vnum
            except:
                m[k] = v
    except Exception as ex:
        forms.alert(u"Error leyendo CSV:\n{0}".format(ex))
    return m

def ut_for_system(perim_m, sysname, ut_global, ut_map):
    if ut_map:
        if sysname in ut_map:
            return ut_value(perim_m, ut_map[sysname])
        ls = sysname.lower()
        for k,v in ut_map.items():
            if k and k.lower() in ls:
                return ut_value(perim_m, v)
    return ut_value(perim_m, ut_global)

from Autodesk.Revit.DB import Domain

def largest_connector_info(fi):
    try:
        cm = fi.MEPModel.ConnectorManager
        conns = [c for c in cm.Connectors]
    except:
        return 0.0, 0.0, "Unknown"
    max_perim = 0.0
    radius_guess = 0.0
    shape = "Unknown"
    for c in conns:
        try:
            if c.Domain != Domain.DomainHvac:
                continue
            shp = c.Shape
            if shp == ConnectorProfileType.Round:
                r_m = (c.Radius or 0.0) * FT_TO_M
                per_m = 2.0 * math.pi * r_m
                if per_m > max_perim:
                    max_perim = per_m; shape = "Round"
                if r_m > radius_guess:
                    radius_guess = r_m
            elif shp == ConnectorProfileType.Rectangular:
                w_m = (c.Width or 0.0) * FT_TO_M
                h_m = (c.Height or 0.0) * FT_TO_M
                per_m = 2.0 * (w_m + h_m)
                if per_m > max_perim:
                    max_perim = per_m; shape = "Rect"
            elif shp == ConnectorProfileType.Oval:
                a_m = (c.Width or 0.0) * FT_TO_M
                b_m = (c.Height or 0.0) * FT_TO_M
                per_m = perim_oval(a_m, b_m)
                if per_m > max_perim:
                    max_perim = per_m; shape = "Oval"
        except:
            continue
    return max_perim, radius_guess, shape

def leq_guess(fi, radius_m):
    tname = fi.Symbol.Family.Name + " : " + fi.Symbol.Name
    name_l = tname.lower()
    if "codo" in name_l or "elbow" in name_l:
        R = radius_m if radius_m > 0 else 0.5
        if "45" in name_l: return 0.25 * math.pi * R
        return 0.5 * math.pi * R
    if "redu" in name_l or "reduc" in name_l or "cone" in name_l:
        return 0.6
    if "tee" in name_l or "t-" in name_l or "deriv" in name_l or "takeoff" in name_l or "wye" in name_l:
        return 0.6
    return 0.5

def leq_from_map(fi, leq_map):
    if not leq_map:
        return None
    key_full = fi.Symbol.Family.Name + " : " + fi.Symbol.Name
    if key_full in leq_map:
        try: return float(leq_map[key_full])
        except: return None
    if fi.Symbol.Name in leq_map:
        try: return float(leq_map[fi.Symbol.Name])
        except: return None
    lower = key_full.lower()
    for k,v in leq_map.items():
        try:
            if k and k.lower() in lower:
                return float(v)
        except:
            pass
    return None

def collect_ducts(filter_sys=None, filter_phase=None):
    ducts = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_DuctCurves)\
        .WhereElementIsNotElementType()\
        .ToElements()
    out = []
    fs = (filter_sys or "").lower().strip()
    fp = (filter_phase or "").lower().strip()
    for d in ducts:
        ok = True
        if fs: ok = fs in system_name(d).lower()
        if ok and fp: ok = fp in phase_name(d).lower()
        if ok: out.append(d)
    return out

def collect_fittings(filter_sys=None, filter_phase=None):
    fits = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_DuctFitting)\
        .WhereElementIsNotElementType()\
        .ToElements()
    out = []
    fs = (filter_sys or "").lower().strip()
    fp = (filter_phase or "").lower().strip()
    for f in fits:
        ok = True
        if fs: ok = fs in system_name(f).lower()
        if ok and fp: ok = fp in phase_name(f).lower()
        if ok: out.append(f)
    return out

def calc_duct_rows(ducts, ut_global, ut_map):
    rows = [["Category","ElementId","System","Level","Phase","Shape",
             "Width_m","Height_m","Diameter_m","Length_m",
             "Perimeter_m","UT_m","Area_m2"]]
    total_area = 0.0
    for d in ducts:
        try:
            d_ft = get_dbl(d, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) or 0.0
            w_ft = get_dbl(d, BuiltInParameter.RBS_CURVE_WIDTH_PARAM) or 0.0
            h_ft = get_dbl(d, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM) or 0.0
            L_ft = get_dbl(d, BuiltInParameter.CURVE_ELEM_LENGTH) or 0.0

            D = d_ft*FT_TO_M; W = w_ft*FT_TO_M; H = h_ft*FT_TO_M; L = L_ft*FT_TO_M

            if D > 0.0: shape = "Circular";    perim = perim_circ(D)
            else:
                if W>0 and H>0 and abs(W-H)>1e-6:
                    shape = "Oval";            perim = perim_oval(W,H)
                else:
                    shape = "Rectangular";     perim = perim_rect(W,H)

            sysn = system_name(d); lvl = level_name(d); ph = phase_name(d)
            ut_m = ut_for_system(perim, sysn, ut_global, ut_map)
            area = perim * max(0.0, (L + ut_m)); total_area += area

            rows.append(["Duct", int(d.Id.IntegerValue), sysn, lvl, ph, shape,
                         round(W,6), round(H,6), round(D,6), round(L,6),
                         round(perim,6), round(ut_m,6), round(area,6)])
        except Exception as ex:
            rows.append(["Duct", int(d.Id.IntegerValue), "ERROR: "+str(ex),"","","",0,0,0,0,0,0,0])
    return rows, total_area

def calc_fitting_rows(fittings, ut_global, ut_map, leq_map, min1m2=True):
    rows = [["Category","ElementId","System","Level","Phase","Type","ConnShapeMax",
             "PerimeterMax_m","Leq_m","UT_m","Area_m2","Area_m2_Final"]]
    total_area = 0.0
    for f in fittings:
        try:
            per_max, r_guess, connshape = largest_connector_info(f)
            sysn = system_name(f); lvl = level_name(f); ph = phase_name(f)
            typ = f.Symbol.Family.Name + " : " + f.Symbol.Name
            ut_m = ut_for_system(per_max, sysn, ut_global, ut_map)

            leq_map_val = leq_from_map(f, leq_map)
            leq_m = leq_map_val if leq_map_val is not None else leq_guess(f, r_guess)

            area = max(0.0, per_max) * max(0.0, (leq_m + ut_m))
            area_fin = max(1.0, area) if min1m2 else area; total_area += area_fin

            rows.append(["Duct Fitting", int(f.Id.IntegerValue), sysn, lvl, ph, typ, connshape,
                         round(per_max,6), round(leq_m,6), round(ut_m,6),
                         round(area,6), round(area_fin,6)])
        except Exception as ex:
            rows.append(["Duct Fitting", int(f.Id.IntegerValue), "ERROR: "+str(ex),"","","","","",0,0,0,0])
    return rows, total_area

def summarize(rows, catlabel, col_system, col_level, col_phase, col_area):
    hdr = rows[0]; res = {}
    for r in rows[1:]:
        try:
            sysn = r[col_system]; lvl = r[col_level]; ph = r[col_phase]; a = float(r[col_area])
            key = (catlabel, sysn, lvl, ph); res[key] = res.get(key, 0.0) + a
        except:
            continue
    table = [["Category","System","Level","Phase","Area_m2_total"]]
    for (cat,sysn,lvl,ph),a in sorted(res.items(), key=lambda k:(k[0][1],k[0][2],k[0][3],k[0][0])):
        table.append([cat, sysn, lvl, ph, round(a,6)])
    return table

def write_csv(path, rows):
    with open(path, 'wb') as f:
        w = csv.writer(f, delimiter=';')
        for r in rows:
            w.writerow([_as_utf8(x) for x in r])

w = _load_wpf()
w.cmbUT.SelectedIndex = 1
w.chkUTMap.IsChecked = False
w.chkLeqMap.IsChecked = False
w.chkMin1.IsChecked = True
w.txtUTCsv.Text = ""
w.txtLeqCsv.Text = ""
w.txtFilterSys.Text = ""
w.txtFilterPh.Text = ""
w.txtOutDir.Text = script.get_output().get_envvar("last_outdir") or ""
w.txtPrefix.Text = "duct_and_fittings_surface"

def pick_csv_into(tb):
    p = forms.pick_file(file_ext='csv', restore_dir=True, title="Seleccionar CSV (clave;valor)")
    if p: tb.Text = p

def on_btnUTCsv(sender, e):  pick_csv_into(w.txtUTCsv)
def on_btnLeqCsv(sender, e): pick_csv_into(w.txtLeqCsv)
def on_btnOutDir(sender, e):
    d = forms.pick_folder(title="Seleccionar carpeta de exportación")
    if d: w.txtOutDir.Text = d

def on_ok(sender, e):
    try:
        ut_global = (w.cmbUT.Text or "perfil20").strip().lower()
        use_utmap = bool(w.chkUTMap.IsChecked);  utcsv = w.txtUTCsv.Text.strip()
        use_leqmap = bool(w.chkLeqMap.IsChecked); leqcsv = w.txtLeqCsv.Text.strip()
        min1 = bool(w.chkMin1.IsChecked)
        filt_sys = w.txtFilterSys.Text.strip();  filt_ph = w.txtFilterPh.Text.strip()
        outdir = w.txtOutDir.Text.strip();       prefix = (w.txtPrefix.Text or "duct_and_fittings_surface").strip()

        if not outdir or not os.path.isdir(outdir):
            forms.alert("Carpeta de exportación no válida."); return

        ut_map = read_map_csv(utcsv) if use_utmap else {}
        tmp = read_map_csv(leqcsv) if use_leqmap else {}
        leq_map = {}
        for k,v in tmp.items():
            try:
                if isinstance(v,(int,float)): leq_map[k] = float(v)
                else: leq_map[k] = float(str(v).replace(',','.'))
            except:
                pass

        ducts = collect_ducts(filt_sys, filt_ph)
        fittings = collect_fittings(filt_sys, filt_ph)

        duct_rows, area_ducts = calc_duct_rows(ducts, ut_global, ut_map)
        fit_rows, area_fits   = calc_fitting_rows(fittings, ut_global, ut_map, leq_map, min1m2=min1)

        duct_sum = summarize(duct_rows, "Duct", 2, 3, 4, 12)
        fit_sum  = summarize(fit_rows,  "Duct Fitting", 2, 3, 4, 11)

        d_det = os.path.join(outdir, prefix + "_DUCTS_DETALLE.csv")
        d_sum = os.path.join(outdir, prefix + "_DUCTS_RESUMEN.csv")
        f_det = os.path.join(outdir, prefix + "_FITTINGS_DETALLE.csv")
        f_sum = os.path.join(outdir, prefix + "_FITTINGS_RESUMEN.csv")

        write_csv(d_det, duct_rows); write_csv(d_sum, duct_sum)
        write_csv(f_det, fit_rows);  write_csv(f_sum, fit_sum)

        script.get_output().set_envvar("last_outdir", outdir)

        msg = u"Exportado:\n- {}\n- {}\n- {}\n- {}\n\nTotal DUCTS [m²]: {:.3f}\nTotal FITTINGS [m²]: {:.3f}".format(
            d_det, d_sum, f_det, f_sum, area_ducts, area_fits
        )
        forms.alert(msg, ok=True); w.Close()
    except Exception as ex:
        forms.alert(u"Error:\n{}\n\n{}".format(ex, traceback.format_exc()), ok=True)

def on_cancel(sender, e): w.Close()

w.btnUTCsv.Click += on_btnUTCsv
w.btnLeqCsv.Click += on_btnLeqCsv
w.btnOutDir.Click += on_btnOutDir
w.btnOk.Click += on_ok
w.btnCancel.Click += on_cancel

w.ShowDialog()
