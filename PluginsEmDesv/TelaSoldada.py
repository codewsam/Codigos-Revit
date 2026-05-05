# -*- coding: utf-8 -*-
__title__ = "Folha de Tela"
__author__ = "Samuel"
__version__ = "Versão 1.0"


"""
Folha de Tela Soldada - AreaReinforcement em Paredes
Automatiza a criação de painéis de tela soldada (AreaReinforcement)
em paredes selecionadas, com layout de painéis e transpasse configurável.

Requisitos:
- Revit 2022+
- pyRevit instalado
- AreaReinforcementType pré-carregado no projeto
- Paredes retas (LocationCurve)


"""

# ──────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────────────────────────────────────
import clr
import math

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

import System
from System.Collections.Generic import List

# pyRevit
from pyrevit import revit, DB, forms, script

# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO INICIAL
# ──────────────────────────────────────────────────────────────────────────────
doc   = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

# Fator de conversão: metros → pés (unidade interna Revit)
M2FT = 3.28084

# ──────────────────────────────────────────────────────────────────────────────
# FILTROS DE SELEÇÃO
# ──────────────────────────────────────────────────────────────────────────────
class WallSelectionFilter(ISelectionFilter):
    """Permite selecionar apenas paredes."""
    def AllowElement(self, element):
        return isinstance(element, Wall)
    def AllowReference(self, ref, point):
        return False


# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def get_element_name(element):
    """Retorna o nome de um elemento com segurança."""
    try:
        p = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p:
            return p.AsString()
        p2 = element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p2:
            return p2.AsString()
        return "Id_{}".format(element.Id.IntegerValue)
    except:
        return "Id_{}".format(element.Id.IntegerValue)


def ft_to_m(ft):
    return ft / M2FT


def m_to_ft(m):
    return m * M2FT


def normalize(v):
    """Normaliza um XYZ."""
    length = math.sqrt(v.X**2 + v.Y**2 + v.Z**2)
    if length < 1e-9:
        return XYZ(0, 0, 0)
    return XYZ(v.X / length, v.Y / length, v.Z / length)


def cross_product(a, b):
    return XYZ(
        a.Y * b.Z - a.Z * b.Y,
        a.Z * b.X - a.X * b.Z,
        a.X * b.Y - a.Y * b.X
    )


# ──────────────────────────────────────────────────────────────────────────────
# COLETA DE TIPOS DO PROJETO
# ──────────────────────────────────────────────────────────────────────────────
def get_area_reinf_types():
    """Retorna dict {nome: elemento} de AreaReinforcementType."""
    collector = FilteredElementCollector(doc).OfClass(AreaReinforcementType)
    result = {}
    for t in collector:
        name = get_element_name(t)
        result[name] = t
    return result


def get_rebar_bar_types():
    """Retorna dict {nome: elemento} de RebarBarType."""
    collector = FilteredElementCollector(doc).OfClass(RebarBarType)
    result = {}
    for t in collector:
        name = get_element_name(t)
        result[name] = t
    return result


def get_rebar_hook_types():
    """Retorna dict {nome: elemento} de RebarHookType."""
    collector = FilteredElementCollector(doc).OfClass(RebarHookType)
    result = {"(Nenhum)": None}
    for t in collector:
        name = get_element_name(t)
        result[name] = t
    return result


# ──────────────────────────────────────────────────────────────────────────────
# GEOMETRIA DE PAREDE
# ──────────────────────────────────────────────────────────────────────────────
def get_wall_geometry(wall):
    """
    Retorna dados geométricos da parede:
      - origin: ponto base esquerdo inferior (XYZ, pés)
      - axis:   vetor unitário ao longo da parede (horizontal)
      - normal: vetor normal à face da parede
      - width_ft:  comprimento da parede em pés
      - height_ft: altura da parede em pés
      - base_level_id: ElementId do nível base
    """
    loc = wall.Location
    if not hasattr(loc, 'Curve'):
        return None

    curve = loc.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    # Eixo horizontal ao longo da parede
    axis = normalize(XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0))

    # Normal da parede (perpendicular ao eixo, no plano XY)
    # Revit: normal aponta para o exterior (frente) da parede
    up = XYZ(0, 0, 1)
    normal = cross_product(axis, up)
    # Verificar orientação com a normal da parede (Revit API)
    try:
        wall_normal = wall.Orientation
        if wall_normal.DotProduct(normal) < 0:
            normal = XYZ(-normal.X, -normal.Y, -normal.Z)
    except:
        pass

    # Comprimento
    width_ft = curve.Length

    # Altura (parâmetro de altura desconectada ou nível topo)
    height_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    if height_param and height_param.AsDouble() > 0:
        height_ft = height_param.AsDouble()
    else:
        # Tentar calcular pela diferença de níveis
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        top_offset  = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)
        height_ft = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
        height_ft = height_ft.AsDouble() if height_ft else m_to_ft(2.7)

    # Nível base
    base_level_param = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
    base_level_id = base_level_param.AsElementId() if base_level_param else ElementId.InvalidElementId

    # Ponto de origem = início da LocationCurve na base
    origin = XYZ(p0.X, p0.Y, p0.Z)

    return {
        'origin':        origin,
        'axis':          axis,
        'normal':        normal,
        'width_ft':      width_ft,
        'height_ft':     height_ft,
        'base_level_id': base_level_id,
    }


# ──────────────────────────────────────────────────────────────────────────────
# LAYOUT DE PAINÉIS
# ──────────────────────────────────────────────────────────────────────────────
def compute_panel_layout(total_length_ft, panel_size_ft, overlap_ft):
    """
    Calcula as posições de início de cada painel ao longo de um eixo.

    Parâmetros (todos em pés):
      total_length_ft : comprimento total da parede
      panel_size_ft   : largura (ou altura) do painel
      overlap_ft      : transpasse/overlap entre painéis adjacentes

    Retorna: lista de (start_ft, end_ft) de cada painel
    Garante que o último painel cobre até o fim da parede.
    """
    panels = []
    pos = 0.0
    step = panel_size_ft - overlap_ft  # avanço efetivo por painel

    if step <= 0:
        # Overlap maior que painel: coloca um painel só
        panels.append((0.0, min(panel_size_ft, total_length_ft)))
        return panels

    while pos < total_length_ft - 1e-6:
        end = pos + panel_size_ft
        if end > total_length_ft:
            end = total_length_ft
        panels.append((pos, end))
        if end >= total_length_ft:
            break
        pos += step

    return panels


# ──────────────────────────────────────────────────────────────────────────────
# CRIAÇÃO DE CURVELOOP PARA UM PAINEL RETANGULAR
# ──────────────────────────────────────────────────────────────────────────────
def make_panel_curveloop(origin, axis, up, h_start_ft, h_end_ft, v_start_ft, v_end_ft):
    """
    Cria um CurveLoop retangular no plano da parede.

    origin    : ponto base da parede (XYZ)
    axis      : vetor unitário horizontal ao longo da parede
    up        : vetor unitário vertical (0,0,1)
    h_start_ft: início horizontal do painel (pés)
    h_end_ft  : fim horizontal do painel (pés)
    v_start_ft: início vertical do painel (pés, desde base da parede)
    v_end_ft  : fim vertical do painel (pés)

    Retorna CurveLoop ou None em caso de erro.
    """
    try:
        # 4 cantos do retângulo
        p0 = XYZ(
            origin.X + axis.X * h_start_ft + up.X * v_start_ft,
            origin.Y + axis.Y * h_start_ft + up.Y * v_start_ft,
            origin.Z + axis.Z * h_start_ft + up.Z * v_start_ft
        )
        p1 = XYZ(
            origin.X + axis.X * h_end_ft + up.X * v_start_ft,
            origin.Y + axis.Y * h_end_ft + up.Y * v_start_ft,
            origin.Z + axis.Z * h_end_ft + up.Z * v_start_ft
        )
        p2 = XYZ(
            origin.X + axis.X * h_end_ft + up.X * v_end_ft,
            origin.Y + axis.Y * h_end_ft + up.Y * v_end_ft,
            origin.Z + axis.Z * h_end_ft + up.Z * v_end_ft
        )
        p3 = XYZ(
            origin.X + axis.X * h_start_ft + up.X * v_end_ft,
            origin.Y + axis.Y * h_start_ft + up.Y * v_end_ft,
            origin.Z + axis.Z * h_start_ft + up.Z * v_end_ft
        )

        loop = CurveLoop()
        loop.Append(Line.CreateBound(p0, p1))
        loop.Append(Line.CreateBound(p1, p2))
        loop.Append(Line.CreateBound(p2, p3))
        loop.Append(Line.CreateBound(p3, p0))
        return loop
    except Exception as e:
        logger.error("Erro ao criar CurveLoop: {}".format(str(e)))
        return None


# ──────────────────────────────────────────────────────────────────────────────
# INTERFACE DO USUÁRIO
# ──────────────────────────────────────────────────────────────────────────────
def show_input_form():
    """
    Exibe formulário de configuração dos painéis.
    Retorna dict com parâmetros ou None se cancelado.
    """
    # 1. AreaReinforcementType
    art_dict = get_area_reinf_types()
    if not art_dict:
        forms.alert(
            "Nenhum AreaReinforcementType encontrado no projeto.\n"
            "Carregue um tipo de Armadura de Área antes de usar este script.",
            title="Folha de Tela Soldada",
            exitscript=True
        )
        return None

    art_name = forms.SelectFromList.show(
        sorted(art_dict.keys()),
        title="Selecione o Tipo de Armadura de Área",
        multiselect=False
    )
    if not art_name:
        return None
    area_reinf_type = art_dict[art_name]

    # 2. RebarBarType
    rbt_dict = get_rebar_bar_types()
    if not rbt_dict:
        forms.alert(
            "Nenhum RebarBarType encontrado no projeto.",
            title="Folha de Tela Soldada",
            exitscript=True
        )
        return None

    # Ordenar e sugerir CA-60 (tela soldada) ou o primeiro disponível
    rbt_names = sorted(rbt_dict.keys())
    default_rbt = next(
        (n for n in rbt_names if "CA-60" in n or "4.2" in n or "5 CA" in n),
        rbt_names[0]
    )
    rbt_name = forms.SelectFromList.show(
        rbt_names,
        title="Selecione o Diâmetro da Armadura (RebarBarType)",
        multiselect=False
    )
    if not rbt_name:
        return None
    rebar_bar_type = rbt_dict[rbt_name]

    # 3. RebarHookType (opcional — tela soldada geralmente sem gancho)
    hook_dict = get_rebar_hook_types()
    hook_name = forms.SelectFromList.show(
        sorted(hook_dict.keys()),
        title="Selecione o Tipo de Gancho (ou '(Nenhum)')",
        multiselect=False
    )
    if hook_name is None:
        hook_name = "(Nenhum)"
    hook_type = hook_dict.get(hook_name)
    hook_type_id = hook_type.Id if hook_type else ElementId.InvalidElementId

    # 4. Parâmetros numéricos via GetValueWindow
    values = forms.ask_for_string(
        prompt=(
            "Parâmetros do painel (separe por vírgula):\n"
            "Largura painel (m), Altura painel (m), Transpasse (cm), "
            "Direção principal (H=horizontal / V=vertical), "
            "Cobertura nominal (cm)\n\n"
            "Exemplo padrão: 2.40, 1.20, 10, H, 2.5"
        ),
        title="Configuração dos Painéis",
        default="2.40, 1.20, 10, H, 2.5"
    )
    if not values:
        return None

    try:
        parts = [p.strip() for p in values.split(',')]
        panel_w_m   = float(parts[0])   # largura painel (m)
        panel_h_m   = float(parts[1])   # altura painel (m)
        overlap_cm  = float(parts[2])   # transpasse (cm)
        direction   = parts[3].upper()  # 'H' ou 'V'
        cover_cm    = float(parts[4])   # cobertura (cm)
    except Exception as e:
        forms.alert(
            "Erro ao ler parâmetros: {}\n\nUse o formato: 2.40, 1.20, 10, H, 2.5".format(str(e)),
            title="Folha de Tela Soldada"
        )
        return None

    if direction not in ('H', 'V'):
        direction = 'H'

    return {
        'area_reinf_type':  area_reinf_type,
        'area_reinf_type_id': area_reinf_type.Id,
        'rebar_bar_type_id': rebar_bar_type.Id,
        'hook_type_id':     hook_type_id,
        'panel_w_ft':       m_to_ft(panel_w_m),
        'panel_h_ft':       m_to_ft(panel_h_m),
        'overlap_ft':       m_to_ft(overlap_cm / 100.0),
        'direction':        direction,
        'cover_ft':         m_to_ft(cover_cm / 100.0),
        'panel_w_m':        panel_w_m,
        'panel_h_m':        panel_h_m,
        'overlap_cm':       overlap_cm,
    }


# ──────────────────────────────────────────────────────────────────────────────
# SELEÇÃO DE PAREDES
# ──────────────────────────────────────────────────────────────────────────────
def pick_walls():
    """
    Permite ao usuário selecionar paredes no modelo.
    Retorna lista de Wall ou lista vazia.
    """
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallSelectionFilter(),
            "Selecione as paredes para aplicar Tela Soldada (Esc para cancelar)"
        )
        walls = [doc.GetElement(r.ElementId) for r in refs]
        return [w for w in walls if isinstance(w, Wall)]
    except Exception:
        return []


# ──────────────────────────────────────────────────────────────────────────────
# CRIAÇÃO DE AREA REINFORCEMENT EM UMA PAREDE
# ──────────────────────────────────────────────────────────────────────────────
def create_area_reinforcement_in_wall(wall, params, t):
    """
    Cria AreaReinforcement(s) em uma parede, dividida em painéis.

    Usa o overload:
      AreaReinforcement.Create(doc, hostElement, IList<CurveLoop>,
                               majorDirection, areaReinforementTypeId,
                               rebarBarTypeId, rebarHookTypeId)

    Retorna o número de painéis criados.
    """
    geo = get_wall_geometry(wall)
    if not geo:
        logger.warning("Parede {} sem LocationCurve. Ignorada.".format(wall.Id.IntegerValue))
        return 0

    origin    = geo['origin']
    axis      = geo['axis']
    normal    = geo['normal']
    width_ft  = geo['width_ft']
    height_ft = geo['height_ft']
    up        = XYZ(0, 0, 1)

    panel_w_ft  = params['panel_w_ft']
    panel_h_ft  = params['panel_h_ft']
    overlap_ft  = params['overlap_ft']
    direction   = params['direction']
    art_type_id = params['area_reinf_type_id']
    rbt_id      = params['rebar_bar_type_id']
    hook_id     = params['hook_type_id']

    # Direção principal da tela
    # H = barras principais horizontais → major_dir = axis (ao longo da parede)
    # V = barras principais verticais   → major_dir = up
    if direction == 'H':
        major_dir = axis
    else:
        major_dir = up

    # Layout de painéis horizontais (ao longo do comprimento)
    h_panels = compute_panel_layout(width_ft, panel_w_ft, overlap_ft)
    # Layout de painéis verticais (ao longo da altura)
    v_panels = compute_panel_layout(height_ft, panel_h_ft, overlap_ft)

    count = 0
    errors = []

    for (h_start, h_end) in h_panels:
        for (v_start, v_end) in v_panels:
            # Verificar dimensão mínima (evitar painéis degenerados)
            if (h_end - h_start) < m_to_ft(0.05) or (v_end - v_start) < m_to_ft(0.05):
                continue

            loop = make_panel_curveloop(
                origin, axis, up,
                h_start, h_end,
                v_start, v_end
            )
            if loop is None:
                errors.append("Parede {}: falha ao criar CurveLoop painel ({:.2f}-{:.2f}, {:.2f}-{:.2f})".format(
                    wall.Id.IntegerValue, h_start, h_end, v_start, v_end))
                continue

            # IList<CurveLoop>
            curve_loops = List[CurveLoop]()
            curve_loops.Add(loop)

            try:
                ar = AreaReinforcement.Create(
                    doc,
                    wall,
                    curve_loops,
                    major_dir,
                    art_type_id,
                    rbt_id,
                    hook_id
                )
                # Aplicar cobertura (cover)
                cover_ft = params.get('cover_ft', m_to_ft(0.025))
                try:
                    for bip in [
                        BuiltInParameter.REBAR_COVER_BOTTOM,
                        BuiltInParameter.REBAR_COVER_TOP,
                        BuiltInParameter.REBAR_COVER_OTHER
                    ]:
                        cover_param = ar.get_Parameter(bip)
                        if cover_param and not cover_param.IsReadOnly:
                            cover_param.Set(cover_ft)
                except:
                    pass  # Cobertura pode ser controlada pelo tipo

                count += 1
            except Exception as e:
                errors.append("Parede {}: erro ao criar AreaReinforcement: {}".format(
                    wall.Id.IntegerValue, str(e)))

    for err in errors:
        logger.error(err)

    return count


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────
def main():
    # 1. Parâmetros
    params = show_input_form()
    if not params:
        script.exit()
        return

    # 2. Seleção de paredes
    walls = pick_walls()
    if not walls:
        forms.alert(
            "Nenhuma parede selecionada. Script encerrado.",
            title="Folha de Tela Soldada"
        )
        script.exit()
        return

    # 3. Criação dentro de uma Transaction
    total_panels  = 0
    total_walls   = 0
    skipped_walls = 0

    with revit.Transaction("Folha de Tela Soldada"):
        for wall in walls:
            try:
                n = create_area_reinforcement_in_wall(wall, params, None)
                if n > 0:
                    total_panels += n
                    total_walls  += 1
                else:
                    skipped_walls += 1
            except Exception as e:
                logger.error("Erro na parede {}: {}".format(wall.Id.IntegerValue, str(e)))
                skipped_walls += 1

    # 4. Resumo
    msg = (
        u"✅ Folha de Tela Soldada — Concluído!\n\n"
        u"Paredes processadas : {}/{}\n"
        u"Painéis criados     : {}\n"
        u"Paredes ignoradas   : {}\n\n"
        u"Parâmetros usados:\n"
        u"  Painel: {:.2f} m × {:.2f} m\n"
        u"  Transpasse: {:.0f} cm\n"
        u"  Direção principal: {}"
    ).format(
        total_walls, len(walls),
        total_panels,
        skipped_walls,
        params['panel_w_m'],
        params['panel_h_m'],
        params['overlap_cm'],
        "Horizontal (H)" if params['direction'] == 'H' else "Vertical (V)"
    )

    if total_panels > 0:
        forms.alert(msg, title="Folha de Tela Soldada")
    else:
        forms.alert(
            u"Nenhum painel criado.\n\n"
            u"Verifique:\n"
            u"• Se as paredes selecionadas são retas (LocationCurve)\n"
            u"• Se os tipos de armadura estão carregados no projeto\n"
            u"• O log do pyRevit para detalhes do erro",
            title="Folha de Tela Soldada — Sem Resultados"
        )


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    main()
