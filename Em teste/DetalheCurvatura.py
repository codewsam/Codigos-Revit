# -*- coding: utf-8 -*-
__title__ = "Detalhe de Curvatura"
__author__ = "Samuel"
__version__ = "Versao 1.2 - posicionamento por lado da parede/abertura"
__doc__ = ("Cria automaticamente os Bending Details (Detalhe de Curvatura) para os "
           "rebars de reforco de aberturas visiveis na view ativa, posicionando cada "
           "detalhe no lado correspondente da parede: barras de cima ficam acima da "
           "parede, barras de baixo ficam abaixo, e barras de esquerda/direita ficam "
           "para os respectivos lados. A rotacao do detalhe tambem e alinhada com a "
           "direcao da barra.")

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
# Sistema de eixos 2D da view (u = direita, v = cima).
# Projetar tudo nesses eixos permite o script funcionar tanto em
# planta quanto em corte/elevacao, sem depender de X/Y/Z fixos.
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
    """Retorna (u_min, u_max, v_min, v_max) projetando a bbox 3D nos eixos da view."""
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
# Criterio: pega TODOS os rebars hospedados em uma Wall que sejam
# visiveis na view ativa. (Reforco de aberturas sempre fica em paredes,
# entao isso cobre o caso sem depender de marcacao adicional.)
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
# Agrupa os rebars por parede hospedeira e calcula a caixa (em u/v)
# da propria parede. Essa caixa e a referencia dos 4 lados
# (cima / baixo / esquerda / direita).
# ------------------------------------------------------------------

GAP_FT = 1.5        # distancia entre o detalhe e a face da parede
MIN_STEP_FT = 1.2   # distancia minima entre detalhes vizinhos no mesmo lado

walls_uv_range = {}   # wall_id -> (u_min, u_max, v_min, v_max)
rebars_by_wall = {}   # wall_id -> lista de rebars

for rb in target_rebars:
    wall_id = rb.GetHostId()
    rebars_by_wall.setdefault(wall_id, []).append(rb)

    if wall_id in walls_uv_range:
        continue

    wall_elem = doc.GetElement(wall_id)
    bb = get_bbox(wall_elem)
    if not bb:
        continue

    walls_uv_range[wall_id] = bbox_uv_range(bb)


# ------------------------------------------------------------------
# 
# ------------------------------------------------------------------

# controla os "slots" ja usados em cada (parede, lado) para nao
# sobrepor varios detalhes um em cima do outro
occupied = {}  # (wall_id, lado) -> lista de coordenadas ja usadas


def free_slot(wall_id, side, coord):
    key = (wall_id, side)
    used = occupied.setdefault(key, [])
    c = coord
    while any(abs(c - u) < MIN_STEP_FT for u in used):
        c += MIN_STEP_FT
    used.append(c)
    return c


placements = []  # (rebar, posicao XYZ, rotacao, lado)
no_bbox_rebars = 0

for wall_id, rebars in rebars_by_wall.items():
    if wall_id not in walls_uv_range:
        no_bbox_rebars += len(rebars)
        continue

    u_min, u_max, v_min, v_max = walls_uv_range[wall_id]

    for rb in rebars:
        bb = get_bbox(rb)
        if not bb:
            no_bbox_rebars += 1
            continue

        bu_min, bu_max, bv_min, bv_max = bbox_uv_range(bb)
        ru = (bu_min + bu_max) / 2.0
        rv = (bv_min + bv_max) / 2.0
        du = bu_max - bu_min  
        dv = bv_max - bv_min   

        
        dist_top = v_max - rv
        dist_bottom = rv - v_min
        dist_left = ru - u_min
        dist_right = u_max - ru

        side, _ = min(
            (
                ("cima", dist_top),
                ("baixo", dist_bottom),
                ("esquerda", dist_left),
                ("direita", dist_right),
            ),
            key=lambda pair: pair[1],
        )

        if side == "cima":
            u = free_slot(wall_id, side, ru)
            v = v_max + GAP_FT
        elif side == "baixo":
            u = free_slot(wall_id, side, ru)
            v = v_min - GAP_FT
        elif side == "esquerda":
            u = u_min - GAP_FT
            v = free_slot(wall_id, side, rv)
        else:  # direita
            u = u_max + GAP_FT
            v = free_slot(wall_id, side, rv)

        pos = from_uv(u, v)

        # alinha a rotacao do detalhe com a direcao da propria barra:
        # barra mais "deitada" (du >= dv) -> 0 graus, barra mais "em pe" -> 90 graus
        rotation = 0.0 if du >= dv else math.pi / 2.0

        placements.append((rb, pos, rotation, side))

if no_bbox_rebars:
    debug_lines.append("Rebars sem bounding box valida: {}".format(no_bbox_rebars))

if not placements:
    forms.alert(
        "Nao foi possivel calcular a posicao de nenhum detalhe.\n\nDiagnostico:\n" + "\n".join(debug_lines),
        exitscript=True,
    )


# ------------------------------------------------------------------
# Criacao dos Bending Details
# ------------------------------------------------------------------

created_count = 0
skipped = []
by_side_count = {"cima": 0, "baixo": 0, "esquerda": 0, "direita": 0}

with Transaction(doc, "Criar Detalhes de Curvatura") as t:
    t.Start()

    for rb, pos, rotation, side in placements:
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
            by_side_count[side] += 1
        except Exception as e:
            skipped.append((rb.Id.IntegerValue, str(e)))
            logger.debug("Falha ao criar bending detail para rebar {} (lado {}): {}".format(rb.Id, side, e))

    t.Commit()


# ------------------------------------------------------------------
# Resultado
# ------------------------------------------------------------------

if created_count:
    msg = "{} detalhe(s) de curvatura criado(s) com sucesso.".format(created_count)
    msg += "\n  Cima: {}  |  Baixo: {}  |  Esquerda: {}  |  Direita: {}".format(
        by_side_count["cima"], by_side_count["baixo"], by_side_count["esquerda"], by_side_count["direita"]
    )
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
