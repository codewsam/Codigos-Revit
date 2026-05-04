# -*- coding: utf-8 -*-
__title__ = "Colocar Ferro na Parede"
__title__ = "Reforço de Parede"
__author__ = "Samuel"
__version__ = "Versão 2.0"
__version__ = "Versão 4.7"
__doc__ = """
_____________________________________________________________________
Descrição:

Selecione paredes e insira famílias estruturais automaticamente nelas.

_____________________________________________________________________
Passo a passo:

1. Selecione a família estrutural desejada
2. Selecione as paredes
3. O plugin insere automaticamente

_____________________________________________________________________
Última atualização:
- [Versão 2.0] - SIMPLIFICADA

Cria reforço de aberturas no padrão de detalhamento:
2 verticais, 1 horizontal inferior e diagonais nos cantos.
"""
# ___  __  __  ____    ___   ____   _____  ____  
#|_ _||  \/  ||  _ \  / _ \ |  _ \ |_   _|/ ___| 
# | | | |\/| || |_) || | | || |_) |  | |  \___ \ 
# | | | |  | ||  __/ | |_| ||  _ <   | |   ___) |
#|___||_|  |_||_|     \___/ |_| \_\  |_|  |____/ 
#=================================================

# Importações

import clr
import os
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, revit, script
from Autodesk.Revit.DB.Structure import *
from pyrevit import revit, forms, script
from System.Collections.Generic import List
from Autodesk.Revit.DB.Structure import *
from pyrevit import revit, forms, script
doc = __revit__.ActiveUIDocument.Document
out = script.get_output()
CM_TO_FT = 1.0 / 30.48



def cm_to_ft(v):
    return float(v) * CM_TO_FT

def ask_float(prompt, default_val):
    txt = forms.ask_for_string(default=str(default_val), prompt=prompt, title="Reforço de Parede")
    if not txt:
        script.exit()
    try:
        return float(txt)
    except Exception:
        forms.alert("Valor inválido: {}".format(txt), exitscript=True)


def bar_type_name(bt):
    p = bt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    if p and p.HasValue and p.AsString():
        return p.AsString()
    return "Tipo_{}".format(bt.Id.IntegerValue)


def create_line_rebar(host, bar_type, normal, p1, p2, tag):
    curves = List[Curve]()
    curves.Add(Line.CreateBound(p1, p2))
    rb = Rebar.CreateFromCurves(
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
    if rb:
        out.print_md("- OK [{}] RebarId={}".format(tag, rb.Id.IntegerValue))
    return rb


def wall_info(wall):
    loc = wall.Location
    if not loc or not hasattr(loc, "Curve") or loc.Curve is None:
        return None

    c = loc.Curve
    p0 = c.GetEndPoint(0)
    p1 = c.GetEndPoint(1)
    axis = (p1 - p0).Normalize()
    normal = axis.CrossProduct(XYZ.BasisZ).Normalize()
    bb = wall.get_BoundingBox(None)
    if not bb:
        return None

    return {
        "curve": c,
        "axis": axis,
        "normal": normal,
        "zmin": bb.Min.Z,
        "zmax": bb.Max.Z,
    }


def opening_points(winfo, bb, cover, edge_ext, vert_ext, diag_len):
    z0 = bb.Min.Z + cover
    z1 = bb.Max.Z - cover
    if z1 <= z0:
        return None

    mid = XYZ((bb.Min.X + bb.Max.X) * 0.5, (bb.Min.Y + bb.Max.Y) * 0.5, (bb.Min.Z + bb.Max.Z) * 0.5)
    proj = winfo["curve"].Project(mid)
    if proj is None:
        return None

    center = proj.XYZPoint
    x = bb.Max.X - bb.Min.X
    y = bb.Max.Y - bb.Min.Y
    opening_w = x if x >= y else y
    half_w = max(opening_w * 0.5 - cover, cm_to_ft(2))

    left = center - winfo["axis"].Multiply(half_w + edge_ext)
    right = center + winfo["axis"].Multiply(half_w + edge_ext)

    # verticals com ancoragem para cima e para baixo
    v_bot = max(z0 - vert_ext, winfo["zmin"] + cover)
    v_top = min(z1 + vert_ext, winfo["zmax"] - cover)

    pL_bot = XYZ(left.X, left.Y, v_bot)
    pL_top = XYZ(left.X, left.Y, v_top)
    pR_bot = XYZ(right.X, right.Y, v_bot)
    pR_top = XYZ(right.X, right.Y, v_top)

    pB_l = XYZ(left.X, left.Y, z0)
    pB_r = XYZ(right.X, right.Y, z0)
    pT_l = XYZ(left.X, left.Y, z1)
    pT_r = XYZ(right.X, right.Y, z1)

    d = max(diag_len, cover)
    diag = [
        (pT_l, pT_l - winfo["axis"].Multiply(d) + XYZ(0, 0, d)),
        (pT_r, pT_r + winfo["axis"].Multiply(d) + XYZ(0, 0, d)),
        (pB_l, pB_l - winfo["axis"].Multiply(d) - XYZ(0, 0, d)),
        (pB_r, pB_r + winfo["axis"].Multiply(d) - XYZ(0, 0, d)),
    ]

    return {
        "left_v": (pL_bot, pL_top),
        "right_v": (pR_bot, pR_top),
        "bottom_h": (pB_l, pB_r),
        "diag": diag,
        "opening_w_cm": opening_w / CM_TO_FT,
        "opening_h_cm": (bb.Max.Z - bb.Min.Z) / CM_TO_FT,
    }


out.print_md("## Reforço de Parede v4.7 (padrão da imagem)")

with forms.WarningBar(title="Selecione as paredes"):
    walls = revit.pick_elements_by_category(BuiltInCategory.OST_Walls)

if not walls:
    forms.alert("Nenhuma parede selecionada.", exitscript=True)

btypes = list(FilteredElementCollector(doc).OfClass(RebarBarType))
if not btypes:
    forms.alert("Não há RebarBarType no projeto.", exitscript=True)

bt_map = {}
for bt in btypes:
    name = bar_type_name(bt)
    if name in bt_map:
        name = "{} ({})".format(name, bt.Id.IntegerValue)
    bt_map[name] = bt

bt_sel = forms.ask_for_one_item(sorted(bt_map.keys()), prompt="Escolha o tipo de vergalhão", title="Reforço de Parede")
if not bt_sel:
    script.exit()
bar_type = bt_map[bt_sel]

cover = cm_to_ft(ask_float("Distancia da parede (cm)", 3.5))
edge_ext = cm_to_ft(ask_float("Extensão horizontal além da abertura (cm)", 30))
vert_ext = cm_to_ft(ask_float("Extensão vertical além da abertura (cm)", 30))
diag_len = cm_to_ft(ask_float("Comprimento da diagonal (cm)", 25))

ok = 0
err = 0

with Transaction(doc, "Reforço de Aberturas v4.7") as t:
    t.Start()

    for wall in walls:
        try:
            info = wall_info(wall)
            if not info:
                err += 1
                continue

            face_shift = info["normal"].Multiply(cover * 0.20)
            ins_ids = wall.FindInserts(True, True, True, True)
            for iid in ins_ids:
                ins = doc.GetElement(iid)
                if not ins:
                    continue
                bb = ins.get_BoundingBox(None)
                if not bb:
                    continue


                pts = opening_points(info, bb, cover, edge_ext, vert_ext, diag_len)
                if not pts:
                    continue


                out.print_md("- Abertura ID {}: largura={:.1f}cm altura={:.1f}cm".format(
                    ins.Id.IntegerValue, pts["opening_w_cm"], pts["opening_h_cm"]))


                lv0, lv1 = pts["left_v"]
                rv0, rv1 = pts["right_v"]
                bh0, bh1 = pts["bottom_h"]

                if create_line_rebar(wall, bar_type, info["normal"], lv0 + face_shift, lv1 + face_shift, "ABERT-V-ESQ"):
                    ok += 1
                if create_line_rebar(wall, bar_type, info["normal"], rv0 + face_shift, rv1 + face_shift, "ABERT-V-DIR"):
                    ok += 1
                if create_line_rebar(wall, bar_type, XYZ.BasisZ, bh0 + face_shift, bh1 + face_shift, "ABERT-H-INF"):
                    ok += 1

                k = 1
                for p1, p2 in pts["diag"]:
                    if create_line_rebar(wall, bar_type, info["normal"], p1 + face_shift, p2 + face_shift, "ABERT-DIAG-{}".format(k)):
                        ok += 1
                    k += 1

        except Exception as ex:
            err += 1
            out.print_md("- Erro parede {}: {}".format(wall.Id.IntegerValue, ex))

    t.Commit()

out.print_md("## Concluído")
out.print_md("- Barras criadas: {}".format(ok))
out.print_md("- Erros: {}".format(err))
forms.alert("Concluído!\n\nBarras criadas: {}\nErros: {}".format(ok, err), warn_icon=False)
