# -*- coding: utf-8 -*-
__title__ = "Cotar Elevacao"
__version__ = "2.1"
__doc__ = (
    "Cota, em UM clique, a parede + o transpasse da tela soldada em vistas\n"
    "de elevacao/corte - reproduz exatamente o padrao usado manualmente no\n"
    "projeto (ver cotas de referencia 1858218/1858250).\n\n"
    "Fluxo: selecione a(s) parede(s) na elevacao e rode. Sem menus.\n\n"
    "Para cada parede selecionada:\n"
    "  1. Acha automaticamente a FabricArea (tela soldada) que se sobrepoe\n"
    "     a ela em planta (bounding box, ignorando Z), escolhendo a que\n"
    "     realmente pertence aquela parede (base mais proxima da base\n"
    "     dela, nao a primeira que aparecer);\n"
    "  2. Monta UMA corrente de cota SEGMENTADA com 3 pontos, na ordem\n"
    "     vertical: base da tela -> topo da parede -> topo da tela, o que\n"
    "     gera automaticamente os 2 segmentos (altura da parede + o\n"
    "     transpasse);\n"
    "  3. Monta tambem UMA cota TOTAL (base da tela -> topo da tela),\n"
    "     por fora da segmentada, mostrando a soma dos dois trechos;\n"
    "  4. Joga as linhas de cota um pouco pra fora da parede (por padrao,\n"
    "     a esquerda - Config.LADO)."
)

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    Options, Solid, PlanarFace, Line, ReferenceArray, XYZ,
    DimensionType, BuiltInParameter, FilteredElementCollector, Wall,
)
from Autodesk.Revit.DB.Structure import FabricArea
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import revit, forms, script

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()
logger = script.get_logger()

try:
    from Autodesk.Revit.DB import UnitTypeId, UnitUtils
    def to_ft(cm):
        return UnitUtils.ConvertToInternalUnits(cm, UnitTypeId.Centimeters)
    def to_cm(ft):
        return UnitUtils.ConvertFromInternalUnits(ft, UnitTypeId.Centimeters)
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType, UnitUtils
    def to_ft(cm):
        return UnitUtils.ConvertToInternalUnits(cm, DisplayUnitType.DUT_CENTIMETERS)
    def to_cm(ft):
        return UnitUtils.ConvertFromInternalUnits(ft, DisplayUnitType.DUT_CENTIMETERS)


# ============================================================
# CONFIG
# ============================================================
class Config(object):
    NOME_TIPO_COTA_PADRAO = "Cota - 2 mm (cm) - 1 casa decimal vermelha"
    ESCALA_BASE = 50.0

    TOL_DIM_ZERO_CM   = 1.0   # pontos mais proximos que isso = duplicata
    COLA_ELEMENTO_CM  = 15.0  # distancia da linha de cota ate a parede
    PASSO_LINHAS_CM   = 12.0  # espacamento entre paredes diferentes selecionadas juntas
    MARGEM_PONTA_CM   = 6.0   # quanto a linha de cota estica alem das pontas
    TOL_OVERLAP_2D_FT = 0.3   # folga (pes) pra considerar parede/area "relacionadas" em planta
    GAP_TOTAL_CM      = 20.0  # afastamento extra da cota TOTAL, por fora da segmentada

    LADO = -1.0  # -1 = cota para a esquerda da parede; 1 = direita

    GLOBAL_Z = XYZ(0, 0, 1)  # vertical real do modelo (nao da vista) p/ achar topo/base


def fator_escala(view):
    escala = float(getattr(view, "Scale", None) or Config.ESCALA_BASE)
    return escala / Config.ESCALA_BASE


class Tolerancias(object):
    def __init__(self, view):
        f = fator_escala(view)
        self.fator = f
        self.tol_dim_zero  = to_ft(Config.TOL_DIM_ZERO_CM)
        self.cola_elemento = to_ft(Config.COLA_ELEMENTO_CM * f)
        self.passo_linhas  = to_ft(Config.PASSO_LINHAS_CM * f)
        self.margem_ponta  = to_ft(Config.MARGEM_PONTA_CM * f)
        self.gap_total     = to_ft(Config.GAP_TOTAL_CM * f)


def dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


# ============================================================
# ETAPA 1 - Coleta de paredes selecionadas
# ============================================================
def coletar_paredes():
    sel_ids = list(uidoc.Selection.GetElementIds())
    if not sel_ids:
        try:
            picked = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Selecione a(s) parede(s) para cotar na elevacao e clique Concluir/Finish"
            )
            sel_ids = [r.ElementId for r in picked]
        except Exception as e:
            logger.debug("Selecao cancelada: {}".format(e))
            forms.alert("Nenhuma parede selecionada. Operacao cancelada.", exitscript=True)

    paredes = []
    for eid in sel_ids:
        el = doc.GetElement(eid)
        if isinstance(el, Wall):
            paredes.append(el)

    if not paredes:
        forms.alert(
            "Nenhuma parede valida na selecao.\n"
            "Selecione a(s) parede(s) que quer cotar (a tela soldada e "
            "encontrada automaticamente).",
            exitscript=True,
        )

    output.print_md("## Cotar Elevacao - **{} parede(s) selecionada(s)**".format(len(paredes)))
    return paredes


# ============================================================
# ETAPA 2 - Encontrar a(s) FabricArea relacionada(s) a cada parede
# ============================================================
def _overlap_2d(bb1, bb2, tol):
    return not (
        bb1.Max.X + tol < bb2.Min.X or bb1.Min.X - tol > bb2.Max.X or
        bb1.Max.Y + tol < bb2.Min.Y or bb1.Min.Y - tol > bb2.Max.Y
    )


def coletar_areas_da_vista(view):
    """Coleta as FabricArea da vista NA HORA - sem cache global. A versao
    anterior guardava isso numa variavel de modulo (_areas_cache), que o
    pyRevit reaproveita entre cliques do botao (engine cacheado). Isso
    fazia o script, ao rodar numa 2a elevacao/parede na mesma sessao,
    usar as FabricArea da 1a vista - pegando telas de outra parede/lift
    e gerando cadeias de pontos invalidas. Foi essa a causa real do erro
    'as referencias nao estao mais paralelas' e da cota quebrada com um
    valor so (144.0) que apareceu na Imagem 150. Recolher direto aqui, a
    cada chamada, elimina isso."""
    return list(FilteredElementCollector(doc, view.Id).OfClass(FabricArea).ToElements())


def encontrar_areas_relacionadas(wall, view, tol, areas_da_vista):
    bb_wall = wall.get_BoundingBox(view)
    if bb_wall is None:
        return []
    relacionadas = []
    for area in areas_da_vista:
        bb_a = area.get_BoundingBox(view)
        if bb_a is None:
            continue
        if _overlap_2d(bb_wall, bb_a, tol):
            relacionadas.append(area)
    return relacionadas


def escolher_area_da_parede(wall, view, areas_relacionadas):
    """Entre as FabricArea relacionadas (pode haver varias - parede
    vizinha, lift de cima, etc, todas se sobrepondo em planta), escolhe
    a que realmente pertence a ESTA parede: a cuja base (Z minimo) mais
    se aproxima da base da propria parede. A versao anterior pegava
    sempre areas[0] - a primeira da colecao, sem nenhuma relacao com
    qual delas e a certa."""
    bb_wall = wall.get_BoundingBox(view)
    if bb_wall is None or not areas_relacionadas:
        return None
    base_parede = bb_wall.Min.Z

    melhor, melhor_diff = None, None
    for area in areas_relacionadas:
        bb_a = area.get_BoundingBox(view)
        if bb_a is None:
            continue
        diff = abs(bb_a.Min.Z - base_parede)
        if melhor is None or diff < melhor_diff:
            melhor, melhor_diff = area, diff
    return melhor


# ============================================================
# ETAPA 3 - Extracao dos pontos verticais (topo da parede + bordas da tela)
# ============================================================
def topo_face_parede(wall):
    """Retorna (Z, Reference) da face de topo (normal ~ Z global) mais alta
    do solido da parede."""
    opt = Options()
    opt.ComputeReferences = True
    try:
        geom = wall.get_Geometry(opt)
    except Exception as e:
        logger.debug("Falha geometria parede {}: {}".format(wall.Id.IntegerValue, e))
        return None
    if geom is None:
        return None

    melhor = None
    for g in geom:
        if isinstance(g, Solid) and g.Volume > 0:
            for face in g.Faces:
                if face.Reference is None or not isinstance(face, PlanarFace):
                    continue
                if dot(face.FaceNormal, Config.GLOBAL_Z) <= 0.9:
                    continue
                z = face.Origin.Z
                if melhor is None or z > melhor[0]:
                    melhor = (z, face.Reference)
    return melhor


def bordas_horizontais_area(area):
    """Retorna ((Z_base, Reference), (Z_topo, Reference)) a partir das
    linhas de contorno horizontais (Z constante) da FabricArea."""
    opt = Options()
    opt.ComputeReferences = True
    try:
        geom = area.get_Geometry(opt)
    except Exception as e:
        logger.debug("Falha geometria area {}: {}".format(area.Id.IntegerValue, e))
        return None, None
    if geom is None:
        return None, None

    pontos = []
    for g in geom:
        if isinstance(g, Line) and g.Reference is not None:
            p0, p1 = g.GetEndPoint(0), g.GetEndPoint(1)
            if abs(p0.Z - p1.Z) < 1e-4:
                pontos.append((p0.Z, g.Reference))
    if not pontos:
        return None, None
    pontos.sort(key=lambda t: t[0])
    return pontos[0], pontos[-1]


def montar_cadeia_de_pontos(wall, view, tolz, areas_da_vista):
    """Monta a cadeia ordenada [(Z, Reference), ...] para UMA parede:
    base da tela -> topo da parede -> topo da tela. Se nao achar tela
    relacionada, cai no fallback so com topo/base da propria parede
    (cota so da altura, sem transpasse)."""
    wall_top = topo_face_parede(wall)
    if wall_top is None:
        output.print_md("  [AVISO] parede {} sem face de topo detectavel - pulada.".format(
            wall.Id.IntegerValue))
        return []

    areas = encontrar_areas_relacionadas(wall, view, to_ft(0) + Config.TOL_OVERLAP_2D_FT, areas_da_vista)
    if not areas:
        output.print_md(
            "  [INFO] parede {}: nenhuma FabricArea relacionada encontrada - "
            "cotando so altura da parede.".format(wall.Id.IntegerValue)
        )
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return []
        return [(bb.Min.Z, None), wall_top]  # None = sem Reference, tratado depois

    # Entre as areas relacionadas, escolhe a que pertence de fato a esta
    # parede (base da area ~ base da parede), nao a primeira da colecao.
    area = escolher_area_da_parede(wall, view, areas)
    if area is None:
        output.print_md("  [AVISO] parede {}: nao foi possivel escolher a FabricArea certa - "
                         "cotando so altura da parede.".format(wall.Id.IntegerValue))
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return []
        return [(bb.Min.Z, None), wall_top]

    if len(areas) > 1:
        output.print_md(
            "  [INFO] parede {}: {} FabricAreas se sobrepoem em planta - "
            "usando a {} (base mais proxima da base da parede).".format(
                wall.Id.IntegerValue, len(areas), area.Id.IntegerValue)
        )

    area_base, area_topo = bordas_horizontais_area(area)
    if area_base is None or area_topo is None:
        output.print_md("  [AVISO] FabricArea {} sem bordas horizontais validas - "
                         "cotando so altura da parede.".format(area.Id.IntegerValue))
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return []
        return [(bb.Min.Z, None), wall_top]

    cadeia = [area_base, wall_top, area_topo]
    cadeia.sort(key=lambda t: t[0])
    cadeia = dedupe_por_z(cadeia, tolz.tol_dim_zero)

    if len(cadeia) < 2:
        output.print_md(
            "  [AVISO] parede {}: apos remover pontos coincidentes, sobrou menos de "
            "2 pontos validos (topo da parede e a base/topo da tela ficaram na mesma "
            "altura). Pulando essa parede - confira se a FabricArea escolhida ({}) "
            "e mesmo a certa.".format(wall.Id.IntegerValue, area.Id.IntegerValue)
        )
    return cadeia


def dedupe_por_z(cadeia, tol):
    if not cadeia:
        return []
    aceitos = [cadeia[0]]
    for z, ref in cadeia[1:]:
        if abs(z - aceitos[-1][0]) > tol:
            aceitos.append((z, ref))
    return aceitos


# ============================================================
# ETAPA 4 - Layout (offset por parede, pra nao colidir quando ha varias)
# ============================================================
def resolver_layout(cadeias_por_parede, sinal, tolz):
    """cadeias_por_parede: lista de (wall, cadeia, perp_wall). Retorna
    lista de tarefas prontas com perp_pos definido - cada parede numa
    linha propria (segmentada), mais uma cota TOTAL por parede (do
    primeiro ao ultimo ponto da cadeia), posicionada por fora da
    segmentada. A total so e criada quando a cadeia tem mais de 2
    pontos (2 pontos = 1 segmento so, a total seria identica a
    segmentada - redundante)."""
    tarefas = []
    usados = []
    for wall, cadeia, perp_wall in cadeias_por_parede:
        if len(cadeia) < 2:
            continue
        nivel = len(usados)
        perp_pos_seg = perp_wall + sinal * (tolz.cola_elemento + nivel * tolz.passo_linhas)
        usados.append(perp_pos_seg)
        tarefas.append({
            "wall_id": wall.Id.IntegerValue,
            "pares": cadeia,
            "perp_pos": perp_pos_seg,
            "tipo": "segmentada",
        })

        if len(cadeia) > 2:
            perp_pos_total = perp_pos_seg + sinal * tolz.gap_total
            tarefas.append({
                "wall_id": wall.Id.IntegerValue,
                "pares": [cadeia[0], cadeia[-1]],
                "perp_pos": perp_pos_total,
                "tipo": "total",
            })
    return tarefas


# ============================================================
# ETAPA 5 - Criacao das cotas
# ============================================================
def find_dim_type_by_name(nome):
    dtypes = list(FilteredElementCollector(doc).OfClass(DimensionType).ToElements())
    for dt in dtypes:
        try:
            nm = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except Exception:
            nm = None
        if nm == nome:
            return dt
    if not dtypes:
        return None
    output.print_md("  [AVISO] tipo de cota '{}' nao encontrado - usando o primeiro disponivel.".format(nome))
    return dtypes[0]


def _mpt(axis, perp, r, perp_pos):
    return XYZ(
        axis.X * r + perp.X * perp_pos,
        axis.Y * r + perp.Y * perp_pos,
        axis.Z * r + perp.Z * perp_pos,
    )


def _cria_dim_line(axis, perp, valores_axis, perp_pos, margem):
    r_min, r_max = min(valores_axis) - margem, max(valores_axis) + margem
    pt1, pt2 = _mpt(axis, perp, r_min, perp_pos), _mpt(axis, perp, r_max, perp_pos)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


def criar_cotas_no_revit(tarefas, view, axis, perp, tolz, dim_type):
    """Cria 1 dimension por tarefa (segmentada ou total). Se algum ponto
    da cadeia nao tiver Reference valida (fallback sem tela), a tarefa e
    reportada e pulada (nao da pra cotar sem referencia real)."""
    criadas, erros = 0, 0
    por_tipo = {"segmentada": 0, "total": 0}
    with revit.Transaction("Cotar Elevacao"):
        for t in tarefas:
            pares = t["pares"]
            # posicao no eixo: como Z e a vertical GLOBAL e a vista de
            # elevacao/corte tem UpDirection = Z global no caso comum,
            # projeta o ponto (0,0,z) no eixo da vista.
            valores_axis = [dot(XYZ(0, 0, z), axis) for z, _ in pares]
            dim_line = _cria_dim_line(axis, perp, valores_axis, t["perp_pos"], tolz.margem_ponta)
            if dim_line is None:
                logger.debug("Linha degenerada para parede {} ({}) - pulada.".format(t["wall_id"], t["tipo"]))
                continue

            ra = ReferenceArray()
            valido = True
            for _, ref in pares:
                if ref is None:
                    valido = False
                    break
                ra.Append(ref)
            if not valido:
                output.print_md("[ERRO] parede {} ({}): referencia ausente, pulando.".format(
                    t["wall_id"], t["tipo"]))
                erros += 1
                continue

            try:
                nd = doc.Create.NewDimension(view, dim_line, ra)
                if dim_type:
                    try:
                        nd.DimensionType = dim_type
                    except Exception as e:
                        logger.debug("Falha ao aplicar DimensionType: {}".format(e))
                criadas += 1
                por_tipo[t["tipo"]] = por_tipo.get(t["tipo"], 0) + 1
            except Exception as e:
                erros += 1
                output.print_md("[ERRO] parede {}: falha ao criar cota: {}".format(t["wall_id"], e))
    return criadas, erros, por_tipo


# ============================================================
# MAIN
# ============================================================
def main():
    view = doc.ActiveView
    axis = view.UpDirection
    perp = view.RightDirection

    paredes = coletar_paredes()
    tolz = Tolerancias(view)
    areas_da_vista = coletar_areas_da_vista(view)

    cadeias_por_parede = []
    for wall in paredes:
        cadeia = montar_cadeia_de_pontos(wall, view, tolz, areas_da_vista)
        if len(cadeia) < 2:
            continue
        bb = wall.get_BoundingBox(view)
        centro = XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0, (bb.Min.Z + bb.Max.Z) / 2.0)
        perp_wall = dot(centro, perp)
        cadeias_por_parede.append((wall, cadeia, perp_wall))

    if not cadeias_por_parede:
        forms.alert(
            "Nao foi possivel montar nenhuma cadeia de cota valida.\n"
            "Verifique se a(s) parede(s) tem geometria solida normal.",
            exitscript=True,
        )

    tarefas = resolver_layout(cadeias_por_parede, Config.LADO, tolz)

    dim_type = find_dim_type_by_name(Config.NOME_TIPO_COTA_PADRAO)

    try:
        criadas, erros, por_tipo = criar_cotas_no_revit(tarefas, view, axis, perp, tolz, dim_type)
    except Exception as e:
        logger.error("Falha critica ao criar cotas: {}".format(e))
        forms.alert("Falha critica ao criar cotas:\n{}".format(e), exitscript=True)
        return

    output.print_md("---")
    output.print_md("## {} cota(s) criada(s): {} segmentada(s) + {} total(is).".format(
        criadas, por_tipo.get("segmentada", 0), por_tipo.get("total", 0)
    ))
    if erros:
        output.print_md("**{} parede(s) falharam.**".format(erros))


if __name__ == "__main__":
    main()
