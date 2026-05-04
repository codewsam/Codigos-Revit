# -*- coding: utf-8 -*-
__title__ = "Reforço de Parede"
__author__ = "Samuel"
__version__ = "Versão 6.6"

import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from pyrevit import forms, revit, script
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

CM_TO_FT = 1.0 / 30.48

def cm_to_ft(v):
    return float(v) * CM_TO_FT

def ask_float(prompt, default_val):
    txt = forms.ask_for_string(default=str(default_val),
                               prompt=prompt,
                               title="Reforço de Parede")
    if not txt:
        script.exit()
    try:
        return float(txt)
    except:
        forms.alert("Valor inválido.", exitscript=True)

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
        True,
    )

# =============================

with forms.WarningBar(title="Selecione as paredes"):
    walls = revit.pick_elements_by_category(BuiltInCategory.OST_Walls)

if not walls:
    forms.alert("Nenhuma parede selecionada.", exitscript=True)

btypes = list(FilteredElementCollector(doc).OfClass(RebarBarType))
if not btypes:
    forms.alert("Não há tipos de vergalhão no projeto.", exitscript=True)

bt_map = {}
for bt in btypes:
    p = bt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if p and p.HasValue:
        bt_map[p.AsString()] = bt

bt_sel = forms.ask_for_one_item(sorted(bt_map.keys()),
                                prompt="Escolha o tipo de vergalhão")

if not bt_sel:
    script.exit()

bar_type = bt_map[bt_sel]

cover = cm_to_ft(ask_float("Cobrimento lateral (cm)", 3))
edge_ext = cm_to_ft(ask_float("Extensão horizontal além da abertura (cm)", 30))
vert_ext = cm_to_ft(ask_float("Extensão vertical além da abertura (cm)", 30))
diag_len = cm_to_ft(ask_float("Comprimento diagonal (cm)", 40))

with Transaction(doc, "Reforço de Aberturas v6.6") as t:
    t.Start()

    for wall in walls:

        loc = wall.Location
        if not hasattr(loc, "Curve"):
            continue

        curve = loc.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)

        axis = (p1 - p0).Normalize()
        normal = axis.CrossProduct(XYZ.BasisZ).Normalize()

        inserts = wall.FindInserts(True, True, True, True)

        for iid in inserts:
            ins = doc.GetElement(iid)
            if not ins:
                continue

            bb = ins.get_BoundingBox(None)
            if not bb:
                continue

            z0 = bb.Min.Z + cover
            z1 = bb.Max.Z - cover

            mid = XYZ((bb.Min.X + bb.Max.X) * 0.5,
                      (bb.Min.Y + bb.Max.Y) * 0.5,
                      (bb.Min.Z + bb.Max.Z) * 0.5)

            proj = curve.Project(mid)
            if not proj:
                continue

            center = proj.XYZPoint

            width = max(bb.Max.X - bb.Min.X,
                        bb.Max.Y - bb.Min.Y)

            half_w = width * 0.5 - cover

            left_face = center - axis.Multiply(half_w)
            right_face = center + axis.Multiply(half_w)

            left_ext = left_face - axis.Multiply(edge_ext)
            right_ext = right_face + axis.Multiply(edge_ext)

            # ===== VERTICAIS =====
            create_rebar(wall, bar_type, normal,
                         XYZ(left_face.X, left_face.Y, z0 - vert_ext),
                         XYZ(left_face.X, left_face.Y, z1 + vert_ext))

            create_rebar(wall, bar_type, normal,
                         XYZ(right_face.X, right_face.Y, z0 - vert_ext),
                         XYZ(right_face.X, right_face.Y, z1 + vert_ext))

            # ===== HORIZONTAIS =====
            create_rebar(wall, bar_type, XYZ.BasisZ,
                         XYZ(left_ext.X, left_ext.Y, z0),
                         XYZ(right_ext.X, right_ext.Y, z0))

            create_rebar(wall, bar_type, XYZ.BasisZ,
                         XYZ(left_ext.X, left_ext.Y, z1),
                         XYZ(right_ext.X, right_ext.Y, z1))

            half = diag_len * 0.5

            # ===== 4 DIAGONAIS =====

            points = [
                (XYZ(left_face.X, left_face.Y, z1),  axis - XYZ.BasisZ),  # sup esquerda
                (XYZ(right_face.X, right_face.Y, z1), -axis - XYZ.BasisZ), # sup direita
                (XYZ(left_face.X, left_face.Y, z0),  axis + XYZ.BasisZ),  # inf esquerda
                (XYZ(right_face.X, right_face.Y, z0), -axis + XYZ.BasisZ) # inf direita
            ]

            for pt, vec in points:
                direction = vec.Normalize()
                start_pt = pt - direction.Multiply(half)
                end_pt   = pt + direction.Multiply(half)

                create_rebar(wall, bar_type, normal,
                             start_pt,
                             end_pt)

    t.Commit()

forms.alert("Reforço criado com sucesso!", warn_icon=False)
