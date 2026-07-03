# -*- coding: utf-8 -*-
__title__ = "Detalhe de Curvatura"
__author__ = "Samuel"
__version__ = "Versao 1.1 - diagnostico + criacao automatica"
__doc__ = "Cria automaticamente os Bending Details (Detalhe de Curvatura) para os rebars de reforco de aberturas visiveis na view ativa"

import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from pyrevit import forms, revit, script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

logger = script.get_logger()


# ------------------------------------------------------------------
# Validacoes iniciais
# ------------------------------------------------------------------

if view.ViewType not in (
    ViewType.Section,
    ViewType.DraftingView,
    ViewType.Detail,
    ViewType.EngineeringPlan,
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.Elevation,
    ViewType.AreaPlan,
):
    forms.alert(
        "Abra uma view 2D (Planta, Corte, Elevacao ou Detail View) antes de rodar este comando.",
        exitscript=True,
    )

bending_types = list(FilteredElementCollector(doc).OfClass(RebarBendingDetailType))
if not bending_types:
    forms.alert(
        "Nao ha nenhum tipo de Bending Detail (Detalhe de Curvatura) carregado no projeto.",
        exitscript=True,
    )

bending_type = bending_types[0]


# ------------------------------------------------------------------
# Identificar rebars de reforco de abertura visiveis na view
# ------------------------------------------------------------------
# Criterio: pega TODOS os rebars hospedados em uma Wall que sejam
# visiveis na view ativa. (Reforco de aberturas sempre fica em paredes,
# entao isso cobre o caso sem depender de marcacao adicional.)

all_rebars = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_Rebar)
    .WhereElementIsNotElementType()
    .ToElements()
)

debug_lines = []
debug_lines.append("Total de Rebar na view: {}".format(len(all_rebars)))

target_rebars = []
no_host = 0
host_not_wall = 0

for rb in all_rebars:
    host = rb.GetHostId()
    if host == ElementId.InvalidElementId:
        no_host += 1
        continue

    host_elem = doc.GetElement(host)
    if not host_elem or not isinstance(host_elem, Wall):
        host_not_wall += 1
        continue

    target_rebars.append(rb)

debug_lines.append("Sem host: {}".format(no_host))
debug_lines.append("Host nao e parede: {}".format(host_not_wall))
debug_lines.append("Rebars selecionados: {}".format(len(target_rebars)))

if not target_rebars:
    forms.alert(
        "Nenhum rebar encontrado na view ativa.\n\nDiagnostico:\n" + "\n".join(debug_lines),
        exitscript=True,
    )


# ------------------------------------------------------------------
# Criacao dos Bending Details
# ------------------------------------------------------------------

SPACING_FT = 2.0
COLUMNS = 6

origin = XYZ(0, 0, 0)
try:
    crop_box = view.CropBox
    if crop_box:
        origin = XYZ(crop_box.Min.X, crop_box.Min.Y, 0)
except:
    pass

created_count = 0
skipped = []

with Transaction(doc, "Criar Detalhes de Curvatura") as t:
    t.Start()

    for i, rb in enumerate(target_rebars):
        col = i % COLUMNS
        row = i // COLUMNS

        pos = XYZ(
            origin.X + col * SPACING_FT,
            origin.Y - row * SPACING_FT,
            0,
        )

        try:
            RebarBendingDetail.Create(
                doc,
                view.Id,
                rb.Id,
                0,
                bending_type,
                pos,
                0.0,
            )
            created_count += 1
        except Exception as e:
            skipped.append((rb.Id.IntegerValue, str(e)))
            logger.debug("Falha ao criar bending detail para rebar {}: {}".format(rb.Id, e))

    t.Commit()


# ------------------------------------------------------------------
# Resultado
# ------------------------------------------------------------------

if created_count:
    msg = "{} detalhe(s) de curvatura criado(s) com sucesso.".format(created_count)
    if skipped:
        msg += "\n{} rebar(s) nao puderam ser processados.".format(len(skipped))
        for rid, err in skipped[:5]:
            msg += "\n- Rebar {}: {}".format(rid, err)
    forms.alert(msg, warn_icon=False)
else:
    msg = "Nenhum detalhe de curvatura pode ser criado.\n\nDiagnostico:\n" + "\n".join(debug_lines)
    msg += "\n\nErros:\n"
    for rid, err in skipped[:5]:
        msg += "- Rebar {}: {}\n".format(rid, err)
    forms.alert(msg, warn_icon=True)
