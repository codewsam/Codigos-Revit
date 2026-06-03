# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versao 5.1 - Complete"

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

CM_TO_FT  = 1.0 / 30.48
MM_TO_FT  = 1.0 / 304.8

def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)

# ── 1. COLETAR TIPOS ─────────────────────────────────────────
fat_list = list(FilteredElementCollector(doc).OfClass(FabricAreaType).ToElements())
fst_list = list(FilteredElementCollector(doc).OfClass(FabricSheetType).ToElements())

if not fat_list:
    forms.alert("Nenhum FabricAreaType encontrado no projeto.", exitscript=True)
if not fst_list:
    forms.alert("Nenhum FabricSheetType encontrado no projeto.", exitscript=True)

fat_map = {get_name(t): t for t in fat_list}
fst_map = {get_name(t): t for t in fst_list}

# ── 2. SELECIONAR TIPO DE TELA ───────────────────────────────
fat_name = forms.SelectFromList.show(
    sorted(fat_map.keys()),
    title="Tipo de Tela Soldada",
    multiselect=False
)
if not fat_name:
    script.exit()

selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

# FabricSheetType automatico: "Tela POP Q92 (2,45 x 6,00)" -> "Q92 (2,45 x 6,00)"
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
    u"Deseja adicionar transpasse acima do limite da parede?",
    title="Transpasse",
    yes=True, no=True
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
    except:
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
    except:
        walls = []

if not walls:
    forms.alert("Nenhuma parede selecionada.", exitscript=True)

# ─────────────────────────────────────────────
RECUO_FT        = 0.0  * CM_TO_FT   
RECOBRIMENTO_FT = 22.0 * MM_TO_FT   

# ─────────────────────────────────────────
criados = 0
erros   = []

with revit.Transaction("Folha de Tela Soldada"):
    for wall in walls:
        try:
            loc   = wall.Location
            curve = loc.Curve
            p0    = curve.GetEndPoint(0)
            p1    = curve.GetEndPoint(1)
            dx    = p1.X - p0.X
            dy    = p1.Y - p0.Y
            L     = (dx*dx + dy*dy) ** 0.5
            axis  = XYZ(dx/L, dy/L, 0.0)

            
            h_param   = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
            height_ft = h_param.AsDouble() if h_param else (2.7 / 0.3048)

            
            bot_left  = XYZ(p0.X + axis.X*RECUO_FT, p0.Y + axis.Y*RECUO_FT, p0.Z)
            bot_right = XYZ(p1.X - axis.X*RECUO_FT, p1.Y - axis.Y*RECUO_FT, p1.Z)
            top_right = XYZ(p1.X - axis.X*RECUO_FT, p1.Y - axis.Y*RECUO_FT, p1.Z + height_ft + transpasse_ft)
            top_left  = XYZ(p0.X + axis.X*RECUO_FT, p0.Y + axis.Y*RECUO_FT, p0.Z + height_ft + transpasse_ft)

            loop = CurveLoop()
            loop.Append(Line.CreateBound(bot_left,  bot_right))
            loop.Append(Line.CreateBound(bot_right, top_right))
            loop.Append(Line.CreateBound(top_right, top_left))
            loop.Append(Line.CreateBound(top_left,  bot_left))

            curve_loops = List[CurveLoop]()
            curve_loops.Add(loop)

            # Create com geometria customizada
            fa = FabricArea.Create(
                doc, wall, curve_loops,
                axis, bot_left,
                fabric_area_type_id, fabric_sheet_type_id
            )

            # Deslocamento adicional do recobrimento = 22 mm
            p_recob = fa.LookupParameter(u"Deslocamento adicional da recobrimento")
            if p_recob and not p_recob.IsReadOnly:
                p_recob.Set(RECOBRIMENTO_FT)

            criados += 1

        except Exception as e:
            erros.append(u"Parede {}: {}".format(wall.Id.IntegerValue, str(e)))

# ── 7. RESUMO ─────────────────────────────────────────────────
msg = (
    u"Tela aplicada!\n\n"
    u"Tipo    : {}\n"
    u"Folha   : {}\n"
    u"Paredes : {}/{}\n"
    u"Recuo lateral : 5 cm cada lado\n"
    u"Recobrimento  : 22 mm\n"
    u"Transpasse    : {}"
).format(fat_name, get_name(selected_fst), criados, len(walls), transpasse_txt)

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Folha de Tela")
