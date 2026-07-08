# -*- coding: utf-8 -*-
__title__ = "Detalhe de Curvatura"
__author__ = "Samuel"
__version__ = "Versao 1.3 - somente barra horizontal de baixo (teste de logica)"
__doc__ = ("Cria o Bending Detail apenas para a barra horizontal de BAIXO de cada "
           "abertura visivel na view ativa (1 detalhe por abertura). Etapa de teste "
           "antes de expandir para cima/esquerda/direita.")

import math
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
# FIM DOS IMPORTS
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
# Sistema de eixos 2D da view (u = direita, v = cima).
# ------------------------------------------------------------------

right_dir = view.RightDirection
up_dir = view.UpDirection
view_origin = view.Origin


def to_uv(pt):
    vec = pt - view_origin
    u = vec.DotProduct(right_dir)
    v = vec.DotProduct(up_dir)
    return u, v


def from_uv(u, v):
    return view_origin + right_dir.Multiply(u) + up_dir.Multiply(v)


def bbox_corners(bbox):
    return [
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z), XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z), XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z), XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z), XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
    ]


def bbox_uv_range(bbox):
    us, vs = [], []
    for c in bbox_corners(bbox):
        u, v = to_uv(c)
        us.append(u)
        vs.append(v)
    return min(us), max(us), min(vs), max(vs)


def get_bbox(elem):
    bb = elem.get_BoundingBox(view)
    if not bb:
        bb = elem.get_BoundingBox(None)
    return bb


# ------------------------------------------------------------------
# Coleta todos os rebars hospedados em Wall, visiveis na view.
# ------------------------------------------------------------------

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
# Agrupamento por abertura (cluster horizontal na mesma parede).
# ------------------------------------------------------------------

GAP_FT = 13.5
CLUSTER_GAP_FT = 4.0
VERTICAL_MARGIN = 1.3    # so classifica como barra vertical se for CLARAMENTE mais alta que larga

rebar_uv = {}
rebars_by_wall = {}

for rb in target_rebars:
    bb = get_bbox(rb)
    if not bb:
        continue
    rebar_uv[rb.Id.IntegerValue] = bbox_uv_range(bb)
    rebars_by_wall.setdefault(rb.GetHostId(), []).append(rb)


def cluster_rebars(rebars):
    rebars_sorted = sorted(rebars, key=lambda r: sum(rebar_uv[r.Id.IntegerValue][0:2]) / 2.0)
    clusters = []
    current = []
    last_ru = None
    for rb in rebars_sorted:
        u_min, u_max, v_min, v_max = rebar_uv[rb.Id.IntegerValue]
        ru = (u_min + u_max) / 2.0
        if last_ru is not None and (ru - last_ru) > CLUSTER_GAP_FT:
            clusters.append(current)
            current = []
        current.append(rb)
        last_ru = ru
    if current:
        clusters.append(current)
    return clusters




placements = []  # (rebar, posicao XYZ, rotacao)

for wall_id, rebars in rebars_by_wall.items():
    clusters = cluster_rebars(rebars)
    wall = doc.GetElement(wall_id)
    wall_bb = get_bbox(wall) if wall else None
    wall_v_min = bbox_uv_range(wall_bb)[2] if wall_bb else None

    for cluster in clusters:
        u_op_min = min(rebar_uv[r.Id.IntegerValue][0] for r in cluster)
        u_op_max = max(rebar_uv[r.Id.IntegerValue][1] for r in cluster)
        v_op_min = min(rebar_uv[r.Id.IntegerValue][2] for r in cluster)
        v_op_max = max(rebar_uv[r.Id.IntegerValue][3] for r in cluster)
        center_v = (v_op_min + v_op_max) / 2.0

        candidatas = []
        for rb in cluster:
            bu_min, bu_max, bv_min, bv_max = rebar_uv[rb.Id.IntegerValue]
            du = bu_max - bu_min
            dv = bv_max - bv_min
            is_vertical = dv > du * VERTICAL_MARGIN
            rv = (bv_min + bv_max) / 2.0
            if not is_vertical and rv < center_v:
                candidatas.append(rb)

        if not candidatas:
            continue

        # entre as candidatas, pega a mais de baixo (menor v) -> 1 barra por abertura
        barra_baixo = min(candidatas, key=lambda r: rebar_uv[r.Id.IntegerValue][2])

        bu_min, bu_max, bv_min, bv_max = rebar_uv[barra_baixo.Id.IntegerValue]
        u = (bu_min + bu_max) / 2.0
        base_v = wall_v_min if wall_v_min is not None else v_op_min
        v = base_v - GAP_FT
        pos = from_uv(u, v)
        rotation = 0.0  

        placements.append((barra_baixo, pos, rotation))

if not placements:
    forms.alert(
        "Nao foi encontrada nenhuma barra horizontal de baixo nas aberturas da view.\n\nDiagnostico:\n" + "\n".join(debug_lines),
        exitscript=True,
    )




created_count = 0
skipped = []

with Transaction(doc, "Criar Detalhes de Curvatura (baixo)") as t:
    t.Start()

    for rb, pos, rotation in placements:
        try:
            RebarBendingDetail.Create(
                doc,
                view.Id,
                rb.Id,
                0,
                bending_type,
                pos,
                rotation,
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
    msg = "{} detalhe(s) de curvatura (barra de baixo) criado(s) com sucesso.".format(created_count)
    if skipped:
        msg += "\n\n{} rebar(s) nao puderam ser processados.".format(len(skipped))
        for rid, err in skipped[:5]:
            msg += "\n- Rebar {}: {}".format(rid, err)
    forms.alert(msg, warn_icon=False)
else:
    msg = "Nenhum detalhe de curvatura pode ser criado.\n\nDiagnostico:\n" + "\n".join(debug_lines)
    msg += "\n\nErros:\n"
    for rid, err in skipped[:5]:
        msg += "- Rebar {}: {}\n".format(rid, err)
    forms.alert(msg, warn_icon=True)
