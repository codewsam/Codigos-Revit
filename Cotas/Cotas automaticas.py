# -*- coding: utf-8 -*-
__title__ = "Cotar Selecao"
__version__ = "2.0"
__doc__ = (
    "Cota automaticamente os elementos selecionados (paredes, pisos, escadas,\n"
    "portas, janelas, qualquer coisa com geometria solida).\n\n"
    "Fluxo:\n"
    " 1. Selecione os elementos ANTES de rodar (ou selecione quando pedido)\n"
    " 2. Escolha a direcao da cota (Horizontal / Vertical, conforme a vista)\n"
    " 3. Clique no lado para onde as cotas devem ser jogadas\n\n"
    "O script varre TODAS as faces planas dos elementos selecionados alinhadas\n"
    "ao eixo escolhido e agora organiza o resultado do mesmo jeito que se\n"
    "costuma cotar manualmente:\n"
    "  - uma corrente de cota PARA CADA alinhamento de parede (cada 'rua' de\n"
    "    paredes vira uma linha de cota propria, coladinha nela, com todos os\n"
    "    vãos/aberturas daquele alinhamento);\n"
    "  - uma cota GERAL por fora de tudo, ponta a ponta (a parede inteira),\n"
    "    igual a linha mais externa que aparece nas vistas do projeto.\n"
    "Cada linha fica numa distancia diferente da anterior (calculada a partir\n"
    "da escala da vista), entao elas nao ficam uma em cima da outra."
)

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    Options, Solid, PlanarFace, ReferenceArray, Line, XYZ,
    Transaction, DimensionType, BuiltInParameter,
    FilteredElementCollector,
)
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

TOL_DIM_ZERO = to_ft(1.0)  # 1 cm - faces mais proximas que isso sao a mesma face (duplicata)

# Nome do tipo de cota padrao do projeto. Ajuste aqui se mudar de padrao.
NOME_TIPO_COTA_PADRAO = "Cota - 2 mm (cm) - 1 casa decimal vermelha"

# ------------------------------------------------------------
# Parametros de organizacao das cotas (todos em cm, na escala 1:50).
# Sao escalados automaticamente pela escala da vista ativa, entao se
# a vista estiver em 1:100 os afastamentos dobram, em 1:25 caem pela
# metade, etc. Ajuste os valores base aqui se quiser cotas mais
# afastadas/coladas.
# ------------------------------------------------------------
ESCALA_BASE = 50.0
escala_vista = float(getattr(revit.doc.ActiveView, "Scale", None) or ESCALA_BASE)
FATOR = escala_vista / ESCALA_BASE

CLUSTER_TOL_CM   = 50.0 * FATOR   # faces a mais que isso uma da outra = alinhamentos/paredes diferentes
COLA_PAREDE_CM   = 15.0 * FATOR   # distancia da linha de cota ate a propria parede (cota "colada")
GAP_GERAL_CM     = 60.0 * FATOR   # afastamento extra da cota GERAL (a mais externa) alem da ultima parede
MARGEM_PONTA_CM  = 20.0 * FATOR   # quanto a linha de cota estica alem do primeiro/ultimo ponto

CLUSTER_TOL  = to_ft(CLUSTER_TOL_CM)
COLA_PAREDE  = to_ft(COLA_PAREDE_CM)
GAP_GERAL    = to_ft(GAP_GERAL_CM)
MARGEM_PONTA = to_ft(MARGEM_PONTA_CM)

# ============================================================
# ETAPA 1 - Vista ativa
# ============================================================
view = doc.ActiveView
right = view.RightDirection
up    = view.UpDirection

def dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z

# ============================================================
# ETAPA 2 - Elementos selecionados
# ============================================================
sel_ids = list(uidoc.Selection.GetElementIds())

if not sel_ids:
    try:
        picked = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Selecione os elementos para cotar (paredes, pisos, escadas, "
            "portas, janelas...) e clique Concluir/Finish"
        )
        sel_ids = [r.ElementId for r in picked]
    except Exception:
        forms.alert("Nenhum elemento selecionado. Operacao cancelada.", exitscript=True)

elements = [doc.GetElement(eid) for eid in sel_ids]
elements = [el for el in elements if el is not None]

if not elements:
    forms.alert("Nenhum elemento valido selecionado.", exitscript=True)

output.print_md("## Cotar Selecao - **{} elementos selecionados**".format(len(elements)))

# ============================================================
# ETAPA 3 - Direcao da cota
# ============================================================
escolha = forms.CommandSwitchWindow.show(
    ["Horizontal (eixo X da vista)", "Vertical (eixo Y da vista)"],
    message="Direcao da cota:"
)
if not escolha:
    forms.alert("Nenhuma direcao escolhida. Operacao cancelada.", exitscript=True)

if escolha.startswith("Horizontal"):
    axis = right
    perp = up
else:
    axis = up
    perp = right

# ============================================================
# HELPERS DE GEOMETRIA
# ============================================================
def faces_by_axis(element, axis_dir, perp_dir, threshold=0.8):
    """Retorna (pos_no_eixo, pos_perpendicular, Reference) de cada face
    plana do elemento cujo normal esteja alinhado (dentro do threshold)
    com axis_dir."""
    opt = Options()
    opt.ComputeReferences = True
    result = []
    try:
        geom = element.get_Geometry(opt)
    except Exception:
        return result
    if geom is None:
        return result
    for g in geom:
        if not isinstance(g, Solid) or g.Volume <= 0:
            continue
        for face in g.Faces:
            if face.Reference is None or not isinstance(face, PlanarFace):
                continue
            d = dot(face.FaceNormal, axis_dir)
            if abs(d) > threshold:
                pos_axis = dot(face.Origin, axis_dir)
                pos_perp = dot(face.Origin, perp_dir)
                result.append((pos_axis, pos_perp, face.Reference))
    return result

def dedupe_por_posicao(trio_list, tol=TOL_DIM_ZERO):
    """Remove faces praticamente coincidentes no eixo de cota (mesma posicao).
    trio_list: lista de (pos_axis, pos_perp, Reference) ja ordenada por pos_axis."""
    if not trio_list:
        return []
    aceitos = [trio_list[0]]
    for pos_axis, pos_perp, ref in trio_list[1:]:
        diff = abs(pos_axis - aceitos[-1][0])
        if diff > tol:
            aceitos.append((pos_axis, pos_perp, ref))
        else:
            output.print_md(
                "  [FILTRO] face a {:.2f}cm da anterior - descartada (duplicata)".format(to_cm(diff))
            )
    return aceitos

def agrupar_por_alinhamento(trio_list, tol=CLUSTER_TOL):
    """Agrupa faces por proximidade da coordenada PERPENDICULAR ao eixo de
    cota. Cada grupo = uma 'rua'/alinhamento de parede diferente, que vai
    virar uma corrente de cota propria (igual acontece quando se cota
    manualmente parede por parede na vista)."""
    if not trio_list:
        return []
    ordenado = sorted(trio_list, key=lambda t: t[1])  # por pos_perp
    grupos = [[ordenado[0]]]
    for item in ordenado[1:]:
        if abs(item[1] - grupos[-1][-1][1]) <= tol:
            grupos[-1].append(item)
        else:
            grupos.append([item])
    return grupos

# ============================================================
# ETAPA 4 - Coleta de TODAS as faces alinhadas ao eixo escolhido
# ============================================================
trios = []
for el in elements:
    trios += faces_by_axis(el, axis, perp)

if len(trios) < 2:
    forms.alert(
        "Menos de 2 referencias encontradas na direcao escolhida.\n"
        "Tente selecionar mais elementos ou trocar a direcao.",
        exitscript=True,
    )

# Lista global ordenada pelo eixo de cota (serve para a cota GERAL, ponta a ponta)
global_ordenado = sorted(trios, key=lambda t: t[0])
global_dedup = dedupe_por_posicao(global_ordenado)

if len(global_dedup) < 2:
    forms.alert(
        "Menos de 2 referencias distintas encontradas.\n"
        "Tente selecionar mais elementos ou trocar a direcao.",
        exitscript=True,
    )

# Agrupamento por alinhamento (cada "parede"/rua vira uma corrente propria)
grupos = agrupar_por_alinhamento(trios, CLUSTER_TOL)

correntes = []  # cada item: {"pares": [(pos,ref),...], "perp": float}
for grupo in grupos:
    grupo_ordenado = sorted(grupo, key=lambda t: t[0])
    grupo_dedup = dedupe_por_posicao(grupo_ordenado)
    if len(grupo_dedup) < 2:
        continue
    perp_medio = sum(t[1] for t in grupo_dedup) / len(grupo_dedup)
    pares = [(t[0], t[2]) for t in grupo_dedup]
    correntes.append({"pares": pares, "perp": perp_medio})

output.print_md("**Alinhamentos de parede encontrados:** {}".format(len(correntes)))
for i, c in enumerate(correntes):
    output.print_md("  - Alinhamento {}: {} referencias".format(i + 1, len(c["pares"])))

# ============================================================
# ETAPA 5 - Ponto de referencia (lado para onde a cota vai)
# ============================================================
try:
    pt_click = uidoc.Selection.PickPoint(
        "Clique do lado para onde as cotas devem ser jogadas"
    )
except Exception:
    forms.alert("Nenhum ponto escolhido. Operacao cancelada.", exitscript=True)

perp_click = dot(pt_click, perp)
perp_medio_geral = sum(t[1] for t in trios) / len(trios)
sinal = 1.0 if perp_click >= perp_medio_geral else -1.0

def mpt(r, perp_pos):
    """Ponto na linha de cota: posicao 'r' no eixo escolhido, mantendo a
    posicao perpendicular 'perp_pos' (fixa para toda a corrente)."""
    return XYZ(
        axis.X * r + perp.X * perp_pos,
        axis.Y * r + perp.Y * perp_pos,
        axis.Z * r + perp.Z * perp_pos,
    )

def cria_dim_line(pares, perp_pos):
    vals = [v for v, _ in pares]
    r_min = min(vals) - MARGEM_PONTA
    r_max = max(vals) + MARGEM_PONTA
    pt1, pt2 = mpt(r_min, perp_pos), mpt(r_max, perp_pos)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)

# ============================================================
# ETAPA 6 - Tipo de cota (busca pelo nome padrao, com fallback)
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
    return dtypes[0] if dtypes else None

dim_type = find_dim_type_by_name(NOME_TIPO_COTA_PADRAO)

# ============================================================
# ETAPA 7 - Montagem das correntes de cota
#   - uma corrente por alinhamento de parede, colada nela
#     (perp_pos = perp real da parede + pequeno afastamento "sinal")
#   - uma corrente GERAL (ponta a ponta de tudo), por fora de todas
# ============================================================
tarefas = []  # cada item: (nome, pares[(pos,ref)], perp_pos)

for c in correntes:
    perp_pos = c["perp"] + sinal * COLA_PAREDE
    tarefas.append(("parede", c["pares"], perp_pos))

# Cota geral (parede inteira / fora de tudo), so faz sentido se agregar
# algo alem do que uma unica corrente de parede ja mostra.
pares_geral = [(t[0], t[2]) for t in global_dedup]
if len(correntes) != 1 or len(pares_geral) != len(correntes[0]["pares"]):
    perp_extremo = max(t["perp"] for t in correntes) if sinal > 0 else min(t["perp"] for t in correntes)
    perp_pos_geral = perp_extremo + sinal * (COLA_PAREDE + GAP_GERAL)
    tarefas.append(("geral", pares_geral, perp_pos_geral))

# ============================================================
# ETAPA 8 - Criacao das cotas
# ============================================================
criadas = 0
with revit.Transaction("Cotar Selecao"):
    for nome, pares, perp_pos in tarefas:
        dim_line = cria_dim_line(pares, perp_pos)
        if dim_line is None:
            continue
        ra = ReferenceArray()
        for _, ref in pares:
            ra.Append(ref)
        try:
            nd = doc.Create.NewDimension(view, dim_line, ra)
            if dim_type:
                try:
                    nd.DimensionType = dim_type
                except Exception as e:
                    logger.debug("Falha ao aplicar DimensionType: {}".format(e))
            criadas += 1
        except Exception as e:
            output.print_md("[ERRO] Falha ao criar cota '{}': {}".format(nome, str(e)))

output.print_md("---")
output.print_md("## {} cota(s) criada(s): {} de parede + {} geral.".format(
    criadas,
    len(correntes),
    1 if len(tarefas) > len(correntes) else 0,
))

