# -*- coding: utf-8 -*-
__title__   = "Folha de Tela"
__author__  = "Samuel"
__version__ = "Versão 1.2.1 - básico bem feito"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import *
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ──────────────────────────────────────────────────────────────
# 1. pegando o FabricAreaType e FabricSheetType do projeto
# ──────────────────────────────────────────────────────────────
fat_list = list(FilteredElementCollector(doc).OfClass(FabricAreaType).ToElements())
fst_list = list(FilteredElementCollector(doc).OfClass(FabricSheetType).ToElements())

if not fat_list:
    forms.alert("Nenhum (FabricAreaType) encontrado no projeto.", exitscript=True)
if not fst_list:
    forms.alert("Nenhum (FabricSheetType) encontrado no projeto.", exitscript=True)

def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)

fat_map = {get_name(t): t for t in fat_list}
fst_map = {get_name(t): t for t in fst_list}

# ──────────────────────────────────────────────────────────────
# 2. SELECIONAR TIPO DE ÁREA (FabricAreaType)
# ──────────────────────────────────────────────────────────────
fat_name = forms.SelectFromList.show(
    sorted(fat_map.keys()),
    title="Seleciona aquele X do centro (???)",
    multiselect=False
)
if not fat_name:
    script.exit()

fabric_area_type    = fat_map[fat_name]
fabric_area_type_id = fabric_area_type.Id

# ──────────────────────────────────────────────────────────────
# 3. SELECIONAR FOLHA (FabricSheetType)
# ──────────────────────────────────────────────────────────────
fst_name = forms.SelectFromList.show(
    sorted(fst_map.keys()),
    title="Seleciona o restante (???)",
    multiselect=False
)
if not fst_name:
    script.exit()

fabric_sheet_type_id = fst_map[fst_name].Id

# ──────────────────────────────────────────────────────────────
# 4. SELECIONAR PAREDES
# ──────────────────────────────────────────────────────────────
class WallFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, Wall)
    def AllowReference(self, ref, pt):
        return False

with forms.WarningBar(title="Selecione as paredes e pressione Enter"):
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallFilter(),
            "Selecione as paredes"
        )
        walls = [doc.GetElement(r.ElementId) for r in refs]
        walls = [w for w in walls if isinstance(w, Wall)]
    except:
        walls = []

if not walls:
    forms.alert("Nenhuma parede selecionada.", exitscript=True)

# ──────────────────────────────────────────────────────────────
# 5. CRIAR FABRICAREA EM CADA PAREDE
# ──────────────────────────────────────────────────────────────
criados = 0
erros   = []

with revit.Transaction("Folha de Tela Soldada"):
    for wall in walls:
        try:
            loc = wall.Location
            if not hasattr(loc, 'Curve'):
                erros.append("Parede {} ignorada (sem LocationCurve)".format(wall.Id.IntegerValue))
                continue

            # Direção principal = ao longo da parede (horizontal)
            curve  = loc.Curve
            p0     = curve.GetEndPoint(0)
            p1     = curve.GetEndPoint(1)
            delta  = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
            length = (delta.X**2 + delta.Y**2) ** 0.5
            major_dir = XYZ(delta.X / length, delta.Y / length, 0.0)

            FabricArea.Create(
                doc,
                wall,
                major_dir,
                fabric_area_type_id,
                fabric_sheet_type_id
            )
            criados += 1

        except Exception as e:
            erros.append("Parede {}: {}".format(wall.Id.IntegerValue, str(e)))

# ──────────────────────────────────────────────────────────────
# 6. RESUMO
# ──────────────────────────────────────────────────────────────
msg = u"Tela Soldada aplicada!\n\nParedes: {}/{}\nTipo: {}\nFolha: {}".format(
    criados, len(walls), fat_name, fst_name
)
if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Folha de Tela")
