# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versao 4.4"

"""
_____________________________________________________________________
Descrição:

Selecione as paredes onde deseja aplicar a folha de tela soldada.
O script cria UMA FabricArea por parede. O Revit calcula
automaticamente a geometria da parede, respeitando os vãos.
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


def get_wall_axis(wall):
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    L  = (dx * dx + dy * dy) ** 0.5
    return XYZ(dx / L, dy / L, 0.0)


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


def get_wall_length(wall):
    curve = wall.Location.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    return (dx * dx + dy * dy) ** 0.5


def get_wall_height(wall):
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Max.Z - bb.Min.Z
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    return h_param.AsDouble() if h_param else (2.7 / 0.3048)


def get_face_loops_da_parede(wall):
    """
    Extrai os CurveLoops da face frontal da parede diretamente
    da geometria do Revit. Esses loops já contêm os furos das
    aberturas calculados pelo próprio Revit, sem nenhuma
    manipulação manual — eliminando qualquer risco de interseção.

    Retorna (List[CurveLoop], origem) ou (None, None) se falhar.
    """
    opts = Options()
    opts.ComputeReferences     = False
    opts.IncludeNonVisibleObjects = False
    opts.DetailLevel           = ViewDetailLevel.Fine

    geom_elem = wall.get_Geometry(opts)
    if geom_elem is None:
        return None, None

    # Direção normal esperada da face frontal
    axis   = get_wall_axis(wall)
    normal = XYZ(-axis.Y, axis.X, 0.0)   # 90° anti-horário = face exterior

    melhor_face   = None
    melhor_area   = 0.0

    for obj in geom_elem:
        solid = None
        if isinstance(obj, Solid):
            solid = obj
        elif isinstance(obj, GeometryInstance):
            for sub in obj.GetInstanceGeometry():
                if isinstance(sub, Solid) and sub.Volume > 1e-9:
                    solid = sub
                    break

        if solid is None or solid.Volume < 1e-9:
            continue

        for face in solid.Faces:
            fn = face.FaceNormal
            # Aceita a face cuja normal é aproximadamente paralela
            # ao vetor normal da parede (frente ou verso)
            dot = abs(fn.X * normal.X + fn.Y * normal.Y)
            if dot < 0.9:
                continue
            # Pega a maior face (parede inteira, não fragmentos)
            area = face.Area
            if area > melhor_area:
                melhor_area = area
                melhor_face = face

    if melhor_face is None:
        return None, None

    loops = melhor_face.GetEdgesAsCurveLoops()
    if not loops:
        return None, None

    curve_loops = List[CurveLoop]()
    origem      = None

    for loop in loops:
        curve_loops.Add(loop)
        if origem is None:
            # Pega o primeiro ponto do primeiro loop como origem
            enumerator = loop.GetEnumerator()
            if enumerator.MoveNext():
                origem = enumerator.Current.GetEndPoint(0)

    if curve_loops.Count == 0 or origem is None:
        return None, None

    return curve_loops, origem


def criar_loop_simples(wall, transpasse_ft):
    """
    Fallback: loop retangular simples cobrindo a parede inteira
    (usado apenas se get_face_loops_da_parede falhar).
    """
    p0     = wall.Location.Curve.GetEndPoint(0)
    axis   = get_wall_axis(wall)
    base_z = get_wall_base_z(wall)
    top_z  = base_z + get_wall_height(wall) + transpasse_ft
    L      = get_wall_length(wall)

    bl = XYZ(p0.X,                   p0.Y,                   base_z)
    br = XYZ(p0.X + axis.X * L,      p0.Y + axis.Y * L,      base_z)
    tr = XYZ(p0.X + axis.X * L,      p0.Y + axis.Y * L,      top_z)
    tl = XYZ(p0.X,                   p0.Y,                   top_z)

    loop = CurveLoop()
    loop.Append(Line.CreateBound(bl, br))
    loop.Append(Line.CreateBound(br, tr))
    loop.Append(Line.CreateBound(tr, tl))
    loop.Append(Line.CreateBound(tl, bl))

    loops = List[CurveLoop]()
    loops.Add(loop)
    return loops, bl


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

            # Tenta usar a geometria real da parede (com furos nativos)
            curve_loops, origem = get_face_loops_da_parede(wall)

            # Fallback: loop simples sem furos
            if curve_loops is None:
                curve_loops, origem = criar_loop_simples(wall, transpasse_ft)

            fa = FabricArea.Create(
                doc, wall, curve_loops,
                axis, origem,
                fabric_area_type_id, fabric_sheet_type_id
            )

            p_recob = fa.LookupParameter(u"Deslocamento adicional da recobrimento")
            if p_recob and not p_recob.IsReadOnly:
                p_recob.Set(RECOBRIMENTO_FT)

            if fazer_transpasse:
                p_lap = fa.get_Parameter(BuiltInParameter.LAP_SPLICE_LENGTH)
                if p_lap and not p_lap.IsReadOnly:
                    p_lap.Set(transpasse_ft)

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
