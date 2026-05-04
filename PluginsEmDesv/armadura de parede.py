# -*- coding: utf-8 -*-
__title__ = "Reforço de Parede"
__author__ = "Samuel"
__version__ = "Versão 6.0"

import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from pyrevit import forms, revit, script
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
out = script.get_output()

CM_TO_FT = 1.0 / 30.48

def cm_to_ft(v):
    return float(v) * CM_TO_FT

def ask_float(prompt, default_val):
    txt = forms.ask_for_string(default=str(default_val),
                               prompt=prompt,
                               title="Reforço de Parede")
    if not txt:
        script.exit()
    return float(txt)

def bar_type_name(bt):
    p = bt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if p and p.HasValue:
        return p.AsString()
    return "Tipo_{}".format(bt.Id.IntegerValue)

def create_rebar(host, bar_type, normal, p1, p2):
    curves = List[Curve]()
    curves.Add(Line.CreateBound(p1, p2))
    return Rebar.CreateFromCurves(
        doc,
        RebarStyle.Standard,
        bar_type,
        None,
        None,
        host,
        normal,
        curves,
        RebarHookOrientation.Left,
        RebarHookOrientation.Right,
        True,
        True
    )

def wall_info(wall):
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    axis = (p1 - p0).Normalize()
    normal = axis.CrossProduct(XYZ.BasisZ).Normalize()

    bb = wall.get_BoundingBox(None)

    return {
        "curve": curve,
        "axis": axis,
        "normal": normal,
        "zmin": bb.Min.Z,
        "zmax": bb.Max.Z
    }

def projected_width(bb, winfo):
    base = winfo["curve"].GetEndPoint(0)
    axis = winfo["axis"]

    pts = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z)
    ]

    vals = [(p - base).DotProduct(axis) for p in pts]
    return abs(vals[1] - vals[0])

with forms.WarningBar(title="Selecione as paredes"):
    walls = revit.pick_elements_by_category(BuiltInCategory.OST_Walls)

btypes = list(FilteredElementCollector(doc).OfClass(RebarBarType))
bt_map = {bar_type_name(bt): bt for bt in btypes}

bt_sel = forms.ask_for_one_item(sorted(bt_map.keys()),
                                prompt="Escolha o vergalhão",
                                title="Reforço")

bar_type = bt_map[bt_sel]

cover = cm_to_ft(ask_float("Cobrimento lateral (cm)", 3))
edge_ext = cm_to_ft(ask_float("Excedente horizontal (cm)", 30))
vert_ext = cm_to_ft(ask_float("Excedente vertical (cm)", 30))
diag_len = cm_to_ft(ask_float("Comprimento diagonal (cm)", 25))

with Transaction(doc, "Reforço Abertura v6.0") as t:
    t.Start()

    for wall in walls:
        info = wall_info(wall)
        axis = info["axis"]
        normal = info["normal"]
        face_shift = normal.Multiply(cover)

        inserts = wall.FindInserts(True, True, True, True)

        for iid in inserts:
            ins = doc.GetElement(iid)
            bb = ins.get_BoundingBox(None)
            if not bb:
                continue

            width = projected_width(bb, info)
            half = width * 0.5

            mid = XYZ(
                (bb.Min.X + bb.Max.X)/2,
                (bb.Min.Y + bb.Max.Y)/2,
                (bb.Min.Z + bb.Max.Z)/2
            )

            proj = info["curve"].Project(mid)
            center = proj.XYZPoint

            z0 = bb.Min.Z + cover
            z1 = bb.Max.Z - cover

            left = center - axis.Multiply(half - cover)
            right = center + axis.Multiply(half - cover)

            left_ext = left - axis.Multiply(edge_ext)
            right_ext = right + axis.Multiply(edge_ext)

            v_bot = z0 - vert_ext
            v_top = z1 + vert_ext

            # Verticais
            pL_bot = XYZ(left.X, left.Y, v_bot)
            pL_top = XYZ(left.X, left.Y, v_top)
            pR_bot = XYZ(right.X, right.Y, v_bot)
            pR_top = XYZ(right.X, right.Y, v_top)

            create_rebar(wall, bar_type, normal, pL_bot+face_shift, pL_top+face_shift)
            create_rebar(wall, bar_type, normal, pR_bot+face_shift, pR_top+face_shift)

            # Horizontais (inferior e superior)
            pB_l = XYZ(left_ext.X, left_ext.Y, z0)
            pB_r = XYZ(right_ext.X, right_ext.Y, z0)

            pT_l = XYZ(left_ext.X, left_ext.Y, z1)
            pT_r = XYZ(right_ext.X, right_ext.Y, z1)

            create_rebar(wall, bar_type, normal, pB_l+face_shift, pB_r+face_shift)
            create_rebar(wall, bar_type, normal, pT_l+face_shift, pT_r+face_shift)

            # Diagonais saindo do VÉRTICE estrutural
            diag_dir1 = (axis + XYZ.BasisZ).Normalize()
            diag_dir2 = (-axis + XYZ.BasisZ).Normalize()

            create_rebar(wall, bar_type, normal,
                         pL_top+face_shift,
                         pL_top + diag_dir1.Multiply(diag_len) + face_shift)

            create_rebar(wall, bar_type, normal,
                         pR_top+face_shift,
                         pR_top + diag_dir2.Multiply(diag_len) + face_shift)

            create_rebar(wall, bar_type, normal,
                         pL_bot+face_shift,
                         pL_bot - diag_dir2.Multiply(diag_len) + face_shift)

            create_rebar(wall, bar_type, normal,
                         pR_bot+face_shift,
                         pR_bot - diag_dir1.Multiply(diag_len) + face_shift)

    t.Commit()

forms.alert("Reforço criado com sucesso!", warn_icon=False)
