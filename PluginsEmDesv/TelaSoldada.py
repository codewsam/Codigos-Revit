# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versao 3.0 - AutoSheet"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import *
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ── 1. COLETAR TIPOS ──────────────────────────────────────────
fat_list = list(FilteredElementCollector(doc).OfClass(FabricAreaType).ToElements())
fst_list = list(FilteredElementCollector(doc).OfClass(FabricSheetType).ToElements())

if not fat_list:
    forms.alert("Nenhum FabricAreaType encontrado no projeto.", exitscript=True)
if not fst_list:
    forms.alert("Nenhum FabricSheetType encontrado no projeto.", exitscript=True)

def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)

def get_dims(el):
    """Retorna (largura, comprimento) em pés, ou (0,0) se nao achar."""
    w, l = 0.0, 0.0
    for pname in ["Largura", "Width", "Comprimento", "Length",
                  "Sheet Width", "Sheet Length", "Fabric Width", "Fabric Length"]:
        p = el.LookupParameter(pname)
        if p and p.StorageType == StorageType.Double:
            v = p.AsDouble()
            if "larg" in pname.lower() or "width" in pname.lower():
                w = v
            else:
                l = v
    return (w, l)

def similarity_score(name_a, name_b):
    """Score simples: quantos tokens de name_a aparecem em name_b."""
    a_tokens = set(name_a.lower().replace('-', ' ').replace('_', ' ').split())
    b_tokens = set(name_b.lower().replace('-', ' ').replace('_', ' ').split())
    if not a_tokens:
        return 0
    return len(a_tokens & b_tokens) / float(len(a_tokens))

def find_best_sheet(fat, fst_list):
    """Acha o FabricSheetType mais parecido com o FabricAreaType dado."""
    fat_name = get_name(fat)
    fat_w, fat_l = get_dims(fat)

    best = None
    best_score = -1.0

    for fst in fst_list:
        fst_name = get_name(fst)
        score = similarity_score(fat_name, fst_name)

        # Bonus se as dimensões forem próximas (dentro de 10%)
        fst_w, fst_l = get_dims(fst)
        if fat_w > 0 and fst_w > 0:
            diff_w = abs(fat_w - fst_w) / max(fat_w, fst_w)
            if diff_w < 0.10:
                score += 0.3
        if fat_l > 0 and fst_l > 0:
            diff_l = abs(fat_l - fst_l) / max(fat_l, fst_l)
            if diff_l < 0.10:
                score += 0.3

        if score > best_score:
            best_score = score
            best = fst

    # Fallback: primeiro da lista
    return best if best else fst_list[0]

fat_map = {get_name(t): t for t in fat_list}

# ── 2. SELECIONAR TIPO DE AREA ────────────────────────────────
fat_name = forms.SelectFromList.show(
    sorted(fat_map.keys()),
    title="Tipo de Area de Tela Soldada",
    multiselect=False
)
if not fat_name:
    script.exit()

selected_fat = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

# ── 3. SHEET AUTOMÁTICO ───────────────────────────────────────
selected_fst = find_best_sheet(selected_fat, fst_list)
fabric_sheet_type_id = selected_fst.Id
matched_fst_name = get_name(selected_fst)

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

# ── 5. CRIAR FABRICAREA ───────────────────────────────────────
criados = 0
erros   = []

OFFSET_COBERTURA = 0.022 / 0.3048 
"""Deslocamento adicional da cobertura em pés (ex: 0.022m = 22mm)"""

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
            major_dir = XYZ(dx / L, dy / L, 0.0)

            fabric_area = FabricArea.Create(doc, wall, major_dir, fabric_area_type_id, fabric_sheet_type_id)

            param = None
            for param_name in ["Deslocamento adicional da recobrimento",
                              "Additional Coverage Offset",
                              "Additional Fabric Offset",
                              "Coverage Offset"]:
                try:
                    param = fabric_area.LookupParameter(param_name)
                    if param and not param.IsReadOnly:
                        param.Set(OFFSET_COBERTURA)
                        break
                except:
                    continue

            criados += 1
        except Exception as e:
            erros.append("Parede {}: {}".format(wall.Id.IntegerValue, str(e)))

# ── 6. RESUMO ─────────────────────────────────────────────────
msg = u"Tela aplicada!\n\nParedes: {}/{}\nTipo de Area: {}\nFolha (auto): {}\n".format(
    criados, len(walls), fat_name, matched_fst_name
)
if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Folha de Tela")
