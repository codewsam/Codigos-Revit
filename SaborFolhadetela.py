# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versao 4.3 - Furo no vao"

"""
_____________________________________________________________________
Descrição:

Selecione as paredes onde deseja aplicar a folha de tela soldada.
O script cria UMA FabricArea por parede, com furos nos vãos de
portas e janelas (a tela não é gerada sobre aberturas).
_____________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import List
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

CM_TO_FT = 1.0 / 30.48
MM_TO_FT = 1.0 / 304.8


def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)


# ── HELPERS ───────────────────────────────────────────────────

def get_wall_base_z(wall):
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Min.Z
    base_level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
    base_level    = doc.GetElement(base_level_id)
    base_elev     = base_level.Elevation if base_level else 0.0
    offset_param  = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    base_offset   = offset_param.AsDouble() if offset_param else 0.0
    return base_elev + base_offset


def get_wall_height(wall):
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Max.Z - bb.Min.Z
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    return h_param.AsDouble() if h_param else (2.7 / 0.3048)


def get_wall_axis(wall):
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    L  = (dx * dx + dy * dy) ** 0.5
    return XYZ(dx / L, dy / L, 0.0)


def get_wall_length(wall):
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    return (dx * dx + dy * dy) ** 0.5


# ── ABERTURAS ─────────────────────────────────────────────────

def get_aberturas(wall):
    """
    Retorna {u_min, u_max, z_min, z_max} de cada porta/janela
    hospedada na parede, em coordenadas locais da parede.
    """
    curve = wall.Location.Curve
    p0    = curve.GetEndPoint(0)
    axis  = get_wall_axis(wall)

    result = []
    for bic in [BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows]:
        for el in (FilteredElementCollector(doc)
                   .OfCategory(bic)
                   .OfClass(FamilyInstance)
                   .WhereElementIsNotElementType()
                   .ToElements()):

            if el.Host is None or el.Host.Id != wall.Id:
                continue
            bb = el.get_BoundingBox(None)
            if bb is None:
                continue

            corners = [
                XYZ(bb.Min.X, bb.Min.Y, 0), XYZ(bb.Max.X, bb.Min.Y, 0),
                XYZ(bb.Min.X, bb.Max.Y, 0), XYZ(bb.Max.X, bb.Max.Y, 0),
            ]
            projs = [(c.X - p0.X) * axis.X + (c.Y - p0.Y) * axis.Y for c in corners]

            result.append({
                "u_min": min(projs),
                "u_max": max(projs),
                "z_min": bb.Min.Z,
                "z_max": bb.Max.Z,
            })

    return result


# ── LOOP EXTERNO (sentido anti-horário visto de frente) ───────

def criar_loop_externo(p0, axis, u0, u1, z0, z1):
    """Loop no sentido anti-horário = contorno externo (sólido)."""
    bl = XYZ(p0.X + axis.X * u0, p0.Y + axis.Y * u0, z0)
    br = XYZ(p0.X + axis.X * u1, p0.Y + axis.Y * u1, z0)
    tr = XYZ(p0.X + axis.X * u1, p0.Y + axis.Y * u1, z1)
    tl = XYZ(p0.X + axis.X * u0, p0.Y + axis.Y * u0, z1)

    loop = CurveLoop()
    loop.Append(Line.CreateBound(bl, br))
    loop.Append(Line.CreateBound(br, tr))
    loop.Append(Line.CreateBound(tr, tl))
    loop.Append(Line.CreateBound(tl, bl))
    return loop, bl


def criar_loop_furo(p0, axis, u_min, u_max, z_min, z_max):
    """
    Loop no sentido HORÁRIO = furo (orientação inversa ao externo).
    O Revit interpreta loops internos com orientação oposta como vazios.
    """
    bl = XYZ(p0.X + axis.X * u_min, p0.Y + axis.Y * u_min, z_min)
    br = XYZ(p0.X + axis.X * u_max, p0.Y + axis.Y * u_max, z_min)
    tr = XYZ(p0.X + axis.X * u_max, p0.Y + axis.Y * u_max, z_max)
    tl = XYZ(p0.X + axis.X * u_min, p0.Y + axis.Y * u_min, z_max)

    # Sentido horário: bl → tl → tr → br → bl
    loop = CurveLoop()
    loop.Append(Line.CreateBound(bl, tl))
    loop.Append(Line.CreateBound(tl, tr))
    loop.Append(Line.CreateBound(tr, br))
    loop.Append(Line.CreateBound(br, bl))
    return loop


# ── MONTAR CURVE LOOPS DA PAREDE ─────────────────────────────

def montar_curve_loops(wall, transpasse_ft):
    """
    Retorna (List[CurveLoop], origem):
      - índice 0: loop externo da parede inteira
      - índices 1..N: um loop-furo por abertura
    """
    base_z    = get_wall_base_z(wall)
    height_ft = get_wall_height(wall)
    top_z     = base_z + height_ft + transpasse_ft
    L         = get_wall_length(wall)
    p0        = wall.Location.Curve.GetEndPoint(0)
    axis      = get_wall_axis(wall)

    loop_ext, origem = criar_loop_externo(p0, axis, 0.0, L, base_z, top_z)

    curve_loops = List[CurveLoop]()
    curve_loops.Add(loop_ext)

    for ab in get_aberturas(wall):
        loop_furo = criar_loop_furo(
            p0, axis,
            ab["u_min"], ab["u_max"],
            ab["z_min"], ab["z_max"]
        )
        curve_loops.Add(loop_furo)

    return curve_loops, origem


# ── 1. COLETAR TIPOS ─────────────────────────────────────────
fat_list = list(FilteredElementCollector(doc).OfClass(FabricAreaType).ToElements())
fst_list = list(FilteredElementCollector(doc).OfClass(FabricSheetType).ToElements())

if not fat_list:
    forms.alert("Nenhum FabricAreaType encontrado no projeto.", exitscript=True)
if not fst_list:
    forms.alert("Nenhum FabricSheetType encontrado no projeto.", exitscript=True)

fat_map = {get_name(t): t for t in fat_list}
fst_map = {get_name(t): t for t in fst_list}

# ── 2. TIPO DE TELA ───────────────────────────────────────────
fat_name = forms.SelectFromList.show(
    sorted(fat_map.keys()),
    title="Tipo de Tela Soldada",
    multiselect=False
)
if not fat_name:
    script.exit()

selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

sheet_suffix = fat_name.replace("Tela POP ", "").strip()
selected_fst = fst_map.get(sheet_suffix)
if not selected_fst:
    for k, v in fst_map.items():
        if sheet_suffix in k or k in sheet_suffix:
            selected_fst = v
            break
if not selected_fst:
    forms.alert(u"Nao foi possivel encontrar a folha '{}' automaticamente.".format(sheet_suffix), exitscript=True)

fabric_sheet_type_id = selected_fst.Id

# ── 3. TRANSPASSE ─────────────────────────────────────────────
fazer_transpasse = forms.alert(
    u"Deseja adicionar transpasse?",
    title="Transpasse", yes=True, no=True
)

transpasse_ft  = 0.0
transpasse_txt = "Nao"

if fazer_transpasse:
    txt = forms.ask_for_string(
        default="20",
        prompt=u"Valor do transpasse (cm):",
        title="Transpasse"
    )
    if not txt:
        script.exit()
    try:
        transpasse_ft  = float(txt) * CM_TO_FT
        transpasse_txt = "{} cm".format(txt)
    except Exception:
        forms.alert("Valor invalido.", exitscript=True)

# ── 4. SELECIONAR PAREDES ─────────────────────────────────────
class WallFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, Wall)
    def AllowReference(self, ref, pt):
        return False

with forms.WarningBar(title="Selecione as paredes e pressione Enter"):
    try:
        refs  = uidoc.Selection.PickObjects(ObjectType.Element, WallFilter(), "Selecione as paredes")
        walls = [doc.GetElement(r.ElementId) for r in refs]
        walls = [w for w in walls if isinstance(w, Wall)]
    except Exception:
        walls = []

if not walls:
    forms.alert("Nenhuma parede selecionada.", exitscript=True)

RECOBRIMENTO_FT = 22.0 * MM_TO_FT

# ── 5. CRIAR TELAS ────────────────────────────────────────────
criados = 0
erros   = []

with revit.Transaction("Folha de Tela Soldada"):
    for wall in walls:
        try:
            axis = get_wall_axis(wall)
            curve_loops, origem = montar_curve_loops(wall, transpasse_ft)

            fa = FabricArea.Create(
                doc, wall, curve_loops,
                axis, origem,
                fabric_area_type_id, fabric_sheet_type_id
            )

            p_recob = fa.LookupParameter(u"Deslocamento adicional da recobrimento")
            if p_recob and not p_recob.IsReadOnly:
                p_recob.Set(RECOBRIMENTO_FT)

            criados += 1

        except Exception as e:
            erros.append(u"Parede {}: {}".format(wall.Id.IntegerValue, str(e)))

# ── 6. RESUMO ─────────────────────────────────────────────────
msg = (
    u"Tela aplicada!\n\n"
    u"Tipo         : {}\n"
    u"Folha        : {}\n"
    u"Paredes      : {}/{}\n"
    u"Recobrimento : 22 mm\n"
    u"Transpasse   : {}"
).format(fat_name, get_name(selected_fst), criados, len(walls), transpasse_txt)

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Folha de Tela")
