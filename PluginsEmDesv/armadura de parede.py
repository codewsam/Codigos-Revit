# -*- coding: utf-8 -*-
__title__ = "Reforço de Parede"
__author__ = "Samuel"
__version__ = "Versão 4.6"
__doc__ = """
Cria apenas reforço de aberturas (horizontais, verticais e diagonais).
"""

import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from pyrevit import revit, forms, script
from System.Collections.Generic import List


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
        "p0": p0,
        "p1": p1,
        "axis": axis,
        "normal": normal,
        "length": c.Length,
        "zmin": bb.Min.Z,
        "zmax": bb.Max.Z,
    }


def opening_frame_points(winfo, bb, cover, diag_extra, edge_ext):
    # Projeta centro da abertura no eixo da parede para manter barras dentro do host
    z_base = bb.Min.Z + cover
    z_top = bb.Max.Z - cover
    if z_top <= z_base:
        return None

    mid = XYZ((bb.Min.X + bb.Max.X) * 0.5, (bb.Min.Y + bb.Max.Y) * 0.5, (bb.Min.Z + bb.Max.Z) * 0.5)
    proj = winfo["curve"].Project(mid)
    if proj is None:
        return None

    center_on_axis = proj.XYZPoint

    # largura útil da abertura na direção do eixo da parede
    x = bb.Max.X - bb.Min.X
    y = bb.Max.Y - bb.Min.Y
    opening_width = x if x >= y else y
    half_w = max(opening_width * 0.5 - cover, cm_to_ft(2))

    left = center_on_axis - winfo["axis"].Multiply(half_w + edge_ext)
    right = center_on_axis + winfo["axis"].Multiply(half_w + edge_ext)

    pL0 = XYZ(left.X, left.Y, z_base)
    pR0 = XYZ(right.X, right.Y, z_base)
    pL1 = XYZ(left.X, left.Y, z_top)
    pR1 = XYZ(right.X, right.Y, z_top)

    # diagonais ancoradas um pouco além do canto
    d = max(diag_extra, cover)
    pDL_a = pL0
    pDL_b = pL0 - winfo["axis"].Multiply(d) + XYZ(0, 0, d)

    pDR_a = pR0
    pDR_b = pR0 + winfo["axis"].Multiply(d) + XYZ(0, 0, d)

    pUL_a = pL1
    pUL_b = pL1 - winfo["axis"].Multiply(d) - XYZ(0, 0, d)

    pUR_a = pR1
    pUR_b = pR1 + winfo["axis"].Multiply(d) - XYZ(0, 0, d)

    return {
        "base_l": pL0,
        "base_r": pR0,
        "top_l": pL1,
        "top_r": pR1,
        "diag": [(pDL_a, pDL_b), (pDR_a, pDR_b), (pUL_a, pUL_b), (pUR_a, pUR_b)],
    }


out.print_md("## Reforço de Parede v4.6 (somente aberturas)")

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

spacing_v = cm_to_ft(ask_float("Espaçamento vertical (cm)", 20))
spacing_h = cm_to_ft(ask_float("Espaçamento horizontal (cm)", 20))
cover = cm_to_ft(ask_float("Cobrimento (cm)", 3.5))
diag_extra = cm_to_ft(ask_float("Ancoragem diagonal (cm)", 25))
edge_ext = cm_to_ft(30)  # 30cm além do vão
pair_gap = cm_to_ft(3)   # 3cm entre barras paralelas

ok = 0
err = 0

with Transaction(doc, "Reforço de Aberturas v4.6") as t:
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

                frame = opening_frame_points(info, bb, cover, diag_extra, edge_ext)
                if not frame:
                    continue

                base_l = frame["base_l"] + face_shift
                base_r = frame["base_r"] + face_shift
                top_l = frame["top_l"] + face_shift
                top_r = frame["top_r"] + face_shift

                if create_line_rebar(wall, bar_type, XYZ.BasisZ, base_l, base_r, "ABERT-BASE-1"):
                    ok += 1
                if create_line_rebar(wall, bar_type, XYZ.BasisZ, base_l + XYZ(0, 0, pair_gap), base_r + XYZ(0, 0, pair_gap), "ABERT-BASE-2"):
                    ok += 1
                if create_line_rebar(wall, bar_type, XYZ.BasisZ, top_l, top_r, "ABERT-TOPO-1"):
                    ok += 1
                if create_line_rebar(wall, bar_type, XYZ.BasisZ, top_l - XYZ(0, 0, pair_gap), top_r - XYZ(0, 0, pair_gap), "ABERT-TOPO-2"):
                    ok += 1

                if create_line_rebar(wall, bar_type, info["normal"], base_l, top_l, "ABERT-V-ESQ"):
                    ok += 1
                if create_line_rebar(wall, bar_type, info["normal"], base_r, top_r, "ABERT-V-DIR"):
                    ok += 1

                k = 1
                for dpa, dpb in frame["diag"]:
                    if create_line_rebar(wall, bar_type, info["normal"], dpa + face_shift, dpb + face_shift, "ABERT-DIAG-{}".format(k)):
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



esse primeiro é como ta ficando e o segundo é como quero que fique
oq ta fazendo isso da errado? to colocando valor errado?

vc consegue identificar o tamanho da jenela/porta e colocar para mim ?
