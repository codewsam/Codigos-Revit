# -*- coding: utf-8 -*-
__title__ = "Cotar Selecao"
__version__ = "1.0"
__doc__ = (
    "Cota automaticamente os elementos selecionados (paredes, pisos, escadas,\n"
    "portas, janelas, qualquer coisa com geometria solida).\n\n"
    "Fluxo:\n"
    " 1. Selecione os elementos ANTES de rodar (ou selecione quando pedido)\n"
    " 2. Escolha a direcao da cota (Horizontal / Vertical, conforme a vista)\n"
    " 3. Clique no ponto onde a linha de cota deve ficar\n\n"
    "O script varre TODAS as faces planas dos elementos selecionados alinhadas\n"
    "ao eixo escolhido e cota tudo numa unica corrente - sem se limitar as\n"
    "extremidades de cada elemento."
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

TOL_DIM_ZERO = to_ft(1.0)  # 1 cm - faces mais proximas que isso sao mescladas

# Nome do tipo de cota padrao do projeto. Ajuste aqui se mudar de padrao.
NOME_TIPO_COTA_PADRAO = "Cota - 2 mm (cm) - 1 casa decimal vermelha"

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
def faces_by_axis(element, axis_dir, threshold=0.8):
    """Retorna todas as faces planas do elemento cujo normal esteja
    alinhado (dentro do threshold) com axis_dir."""
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
                pos = dot(face.Origin, axis_dir)
                result.append((pos, face.Reference))
    return result

def filter_zero_segments(pairs):
    """Remove faces praticamente coincidentes (dentro de TOL_DIM_ZERO)."""
    if not pairs:
        return []
    accepted = [pairs[0]]
    for val, ref in pairs[1:]:
        diff = abs(val - accepted[-1][0])
        if diff > TOL_DIM_ZERO:
            accepted.append((val, ref))
        else:
            output.print_md("  [FILTRO] face a {:.2f}cm da anterior - descartada".format(to_cm(diff)))
    return accepted

# ============================================================
# ETAPA 4 - Coleta de TODAS as faces alinhadas ao eixo escolhido
# ============================================================
pairs = []
for el in elements:
    pairs += faces_by_axis(el, axis)

pairs.sort(key=lambda x: x[0])
pairs = filter_zero_segments(pairs)
refs = [r for _, r in pairs]

output.print_md("**Referencias encontradas:** {}".format(len(refs)))

if len(refs) < 2:
    forms.alert(
        "Menos de 2 referencias encontradas na direcao escolhida.\n"
        "Tente selecionar mais elementos ou trocar a direcao.",
        exitscript=True,
    )

# ============================================================
# ETAPA 5 - Onde colocar a linha de cota
# ============================================================
try:
    pt_click = uidoc.Selection.PickPoint("Clique onde a linha de cota deve ficar")
except Exception:
    forms.alert("Nenhum ponto escolhido. Operacao cancelada.", exitscript=True)

perp_pos = dot(pt_click, perp)

vals = [v for v, _ in pairs]
r_min = min(vals) - to_ft(30)
r_max = max(vals) + to_ft(30)

def mpt(r):
    # ponto na linha de cota: posicao 'r' no eixo escolhido,
    # mantendo a posicao perpendicular do clique do usuario
    return XYZ(
        axis.X * r + perp.X * perp_pos,
        axis.Y * r + perp.Y * perp_pos,
        axis.Z * r + perp.Z * perp_pos,
    )

pt1, pt2 = mpt(r_min), mpt(r_max)
if pt1.DistanceTo(pt2) < 1e-6:
    forms.alert("Linha de cota invalida (pontos coincidentes).", exitscript=True)

dim_line = Line.CreateBound(pt1, pt2)

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
# ETAPA 7 - Criar a cota
# ============================================================
with revit.Transaction("Cotar Selecao"):
    ra = ReferenceArray()
    for r in refs:
        ra.Append(r)
    try:
        nd = doc.Create.NewDimension(view, dim_line, ra)
        if dim_type:
            try:
                nd.DimensionType = dim_type
            except Exception as e:
                logger.debug("Falha ao aplicar DimensionType: {}".format(e))
        output.print_md("---")
        output.print_md("## Cota criada com **{}** referencias.".format(ra.Size))
    except Exception as e:
        forms.alert("Erro ao criar a cota:\n{}".format(str(e)))
        
