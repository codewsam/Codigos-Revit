# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versao 2.0"

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

fat_map = {get_name(t): t for t in fat_list}
fst_map = {get_name(t): t for t in fst_list}

# ── 2. SELECIONAR TIPO DE AREA ────────────────────────────────
fat_name = forms.SelectFromList.show(
    sorted(fat_map.keys()),
    title="Tipo de Area de Tela Soldada",
    multiselect=False
)
if not fat_name:
    script.exit()
fabric_area_type_id = fat_map[fat_name].Id

# ── 3. SELECIONAR FOLHA ───────────────────────────────────────
fst_name = forms.SelectFromList.show(
    sorted(fst_map.keys()),
    title="Folha de Tela Soldada",
    multiselect=False
)
if not fst_name:
    script.exit()
fabric_sheet_type_id = fst_map[fst_name].Id

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

# Deslocamento adicional de cobertura: 22mm
# Convertendo para pés: 22mm = 0.022m / 0.3048 = 0.0722 pés
OFFSET_COBERTURA = 0.022 / 0.3048

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
            
            # Aplicar deslocamento adicional de cobertura de 22mm
            # Tentando diferentes nomes de parâmetros
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
msg = u"Tela aplicada!\n\nParedes: {}/{}\nTipo: {}\nFolha: {}".format(
    criados, len(walls), fat_name, fst_name
)
if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Folha de Tela")
