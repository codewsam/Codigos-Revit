# -*- coding: utf-8 -*-
__title__ = "Cotar(sel)"
__version__ = "2.3"
__doc__ = ("cota - v2.3 (Fase 1 - causa real do bug de posicao): a causa "
           "verdadeira da cota nascer longe da parede selecionada NAO era "
           "o calculo de lado (v2.2) - era 'PlanarFace.Origin'. Em faces "
           "de ponta/topo geradas por encontro (miter/join) com outra "
           "parede, a API pode devolver um ponto do plano que fica FORA "
           "dos limites reais da face (as vezes coincidindo com a "
           "posicao de uma parede vizinha inteiramente diferente). Como "
           "TODO pos_axis/pos_perp do pipeline vem de "
           "'extrair_faces_referenciaveis', isso bastava para jogar a "
           "cota pra longe da parede real. Corrigido usando um ponto "
           "avaliado de verdade dentro da face (bounding box em UV + "
           "Evaluate) em vez do Origin cru.\n\n"
           "cota - v2.2 (mantido, ainda ajuda em outros casos): o calculo "
           "de qual lado da parede recebe a cota individual (parede_total/"
           "vaos) usa apenas as faces de contexto que sobrepoem o trecho "
           "da propria parede, em vez da media de TODAS as paredes "
           "paralelas da view.\n\n"
           "cota - v2.1: mantem a unidade principal = PAREDE (v2.0). Toda "
           "parede e' processada individualmente e SEMPRE recebe cota "
           "geral (+ subcotas quando houver detalhe). Correntes/alinhamentos "
           "continuam servindo so para layout de alinhamento_total/geral.\n\n"
           "Novidades v2.1 (cadeia de referencias completa, como a Cota "
           "Automatica do Revit):\n"
           " - Faces de extremidade da propria parede agora sao marcadas com "
           "ref_tipo 'extremidade_parede' (antes ficavam como 'parede' e a "
           "prioridade nunca era aplicada corretamente).\n"
           " - Pontos de intersecao/cruzamento (T, L, X e prolongamentos "
           "colineares nas pontas) agora sao marcados com ref_tipo "
           "'intersecao' + wall_intersecao_id, entao entram na cadeia com a "
           "prioridade certa (logo depois das extremidades).\n"
           " - Aberturas (portas/janelas/outras hospedadas) continuam usando "
           "as faces de inicio/fim (largura) ja extraidas da geometria - "
           "nada de duplicar essa logica.\n"
           " - Centro da abertura agora e' de fato implementado (antes so "
           "existia a flag Config.USAR_CENTRO_ABERTURA sem efeito nenhum). "
           "Continua DESLIGADO por padrao.\n"
           " - Cada parede processada agora imprime a lista ordenada de "
           "referencias (reaproveitando ordenar_deduplicar_por_prioridade / "
           "_descrever_referencia_parede / _imprimir_lista_referencias_parede, "
           "que ja existiam no codigo mas nunca eram chamadas)."
           )

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    Options, Solid, PlanarFace, ReferenceArray, Line, XYZ, UV,
    DimensionType, BuiltInParameter, BuiltInCategory,
    FilteredElementCollector, Dimension, Wall, FamilyInstance,
    ElementCategoryFilter, LogicalOrFilter, WallFunction,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# FamilyInstanceReferenceType so existe em versoes mais novas da API (2018+)
# e so faz sentido se Config.USAR_CENTRO_ABERTURA estiver ligado - por isso
# o import e' protegido e falha em silencio (feature fica so indisponivel).
try:
    from Autodesk.Revit.DB import FamilyInstanceReferenceType
    _TEM_CENTRO_ABERTURA_API = True
except ImportError:
    FamilyInstanceReferenceType = None
    _TEM_CENTRO_ABERTURA_API = False

from pyrevit import revit, forms, script

doc    = revit.doc
uidoc  = revit.uidoc
output = script.get_output()
logger = script.get_logger()

# ------------------------------------------------------------
# Unidades (compat Revit 2021- / 2022+)
# ------------------------------------------------------------
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

    # Tolerancias base (cm), escaladas pela escala da vista em runtime
    TOL_DIM_ZERO_CM      = 1.0    # faces mais proximas = duplicata (mesmo ponto)
    CLUSTER_TOL_CM       = 35.0   # separa alinhamentos/paredes diferentes
    COLA_ELEMENTO_CM     = 20.0   # distancia da cota "colada" ate o elemento
    GAP_NIVEL_CM         = 20.0   # espacamento entre niveis de linha (0->1->2)
    GAP_GERAL_CM         = 80.0   # afastamento extra da cota GERAL (nivel 3)
    MARGEM_PONTA_CM      = 20.0   # quanto a linha de cota estica alem das pontas
    SUBCOTA_MIN_SEGMENTO_CM = 6.0 # evita segmentos pequenos/repetidos tipo espessura 10
    SUBCOTA_FACE_INTERNA_MAX_CM = 40.0 # ate aqui considera face interna da ponta

    # [v1.2] Cruzamentos perpendiculares (sub-cotas dentro da corrente)
    TOL_CRUZAMENTO_EXTENSAO_CM = 5.0    # quanto o cruzamento pode "estourar" a
                                          # extensao da parede (cantos/juntas)
    LARGURA_CRUZAMENTO_PADRAO_CM = 30.0  # fallback se a Width da parede
                                          # perpendicular nao puder ser lida
    FATOR_LARGURA_CRUZAMENTO = 1.5       # margem de busca da face mais
                                          # proxima = largura da parede * isso

    # Filtro de ruido: ignora faces menores que isso (m2)
    AREA_MINIMA_FACE_M2 = 0.02

    # [v2.1] Centro da abertura como referencia extra (opcional). Precisa de
    # familia com reference plane "Center (Left/Right)" e API >= 2018.
    USAR_CENTRO_ABERTURA = False
    IMPRIMIR_LISTA_REFERENCIAS = True

    # Categorias auto-coletadas quando nao ha selecao previa.
    # Escadas ficam fora de proposito: seus degraus geram muitas faces
    # repetidas e deixam a planta poluida.
    CATEGORIAS_AUTO = [
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Floors,
    ]


def fator_escala(view):
    escala = float(getattr(view, "Scale", None) or Config.ESCALA_BASE)
    return escala / Config.ESCALA_BASE


class Tolerancias(object):
    def __init__(self, view):
        f = fator_escala(view)
        self.fator = f
        self.tol_dim_zero  = to_ft(Config.TOL_DIM_ZERO_CM)
        self.cluster_tol   = to_ft(Config.CLUSTER_TOL_CM * f)
        self.cola_elemento = to_ft(Config.COLA_ELEMENTO_CM * f)
        self.gap_nivel     = to_ft(Config.GAP_NIVEL_CM * f)
        self.gap_geral     = to_ft(Config.GAP_GERAL_CM * f)
        self.margem_ponta  = to_ft(Config.MARGEM_PONTA_CM * f)
        self.subcota_min_segmento = to_ft(Config.SUBCOTA_MIN_SEGMENTO_CM)
        self.subcota_face_interna_max = to_ft(Config.SUBCOTA_FACE_INTERNA_MAX_CM)
        # cruzamento nao escala com a vista - e' geometria real do modelo
        self.tol_cruzamento_extensao = to_ft(Config.TOL_CRUZAMENTO_EXTENSAO_CM)
        self.largura_cruzamento_padrao = to_ft(Config.LARGURA_CRUZAMENTO_PADRAO_CM)


def dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


def _tipo_referencia_do_elemento(el, host_id):
    """Classificacao BASE pelo tipo de elemento. Isso e' refinado depois
    (faces de extremidade viram 'extremidade_parede', faces de cruzamento
    viram 'intersecao') - ver processar_paredes_individualmente."""
    if isinstance(el, Wall):
        return "parede"
    if isinstance(el, FamilyInstance):
        try:
            cat = el.Category
            if cat is not None:
                cat_id = cat.Id.IntegerValue
                if cat_id == int(BuiltInCategory.OST_Doors):
                    return "porta"
                if cat_id == int(BuiltInCategory.OST_Windows):
                    return "janela"
        except Exception:
            pass
        if host_id is not None:
            return "abertura_hospedada"
    if host_id is not None:
        return "abertura_hospedada"
    return "outro"


def _prioridade_referencia(it):
    """Ordem de prioridade da cadeia de cotagem (pedido explicito):
    1) extremidades da parede, 2) intersecoes, 3) portas, 4) janelas,
    5) demais aberturas hospedadas, 6) outras faces de parede, 7) resto."""
    tipo = it.get("ref_tipo")
    prioridades = {
        "extremidade_parede": 0,
        "intersecao": 1,
        "porta": 2,
        "porta_centro": 2,
        "janela": 3,
        "janela_centro": 3,
        "abertura_hospedada": 4,
        "parede": 5,
        "outro": 6,
    }
    return prioridades.get(tipo, 6)


def ordenar_deduplicar_por_prioridade(itens, tol):
    """Ordena por posicao no eixo e, quando duas referencias caem no mesmo
    ponto (dentro de tol), mantem a de maior prioridade (extremidade >
    intersecao > porta > janela > abertura > parede > outro) em vez de
    simplesmente a primeira encontrada na geometria."""
    if not itens:
        return []
    ordenados = sorted(itens, key=lambda t: (t["pos_axis"], _prioridade_referencia(t)))
    aceitos = [ordenados[0]]
    for it in ordenados[1:]:
        if abs(it["pos_axis"] - aceitos[-1]["pos_axis"]) > tol:
            aceitos.append(it)
        else:
            logger.debug("Referencia a {:.2f}cm da anterior - mantendo a de maior prioridade.".format(
                to_cm(abs(it["pos_axis"] - aceitos[-1]["pos_axis"]))))
    return aceitos


def _descrever_referencia_parede(it, p_ini, p_fim, tol):
    pos = it["pos_axis"]
    tipo = it.get("ref_tipo")
    if tipo == "extremidade_parede" or abs(pos - p_ini["pos_axis"]) <= tol:
        return "Face inicial da parede"
    if abs(pos - p_fim["pos_axis"]) <= tol:
        return "Face final da parede"
    if tipo == "intersecao":
        return "Intersecao com parede {}".format(it.get("wall_intersecao_id", "?"))
    if tipo == "porta":
        return "Porta {} (limite)".format(it.get("source_element_id", "?"))
    if tipo == "porta_centro":
        return "Porta {} (centro)".format(it.get("source_element_id", "?"))
    if tipo == "janela":
        return "Janela {} (limite)".format(it.get("source_element_id", "?"))
    if tipo == "janela_centro":
        return "Janela {} (centro)".format(it.get("source_element_id", "?"))
    if tipo == "abertura_hospedada":
        return "Abertura hospedada {} (limite)".format(it.get("source_element_id", "?"))
    return "Referencia complementar {}".format(it.get("source_element_id", "?"))


def _imprimir_lista_referencias_parede(wall_id, pontos_sub, p_ini, p_fim, tol):
    if not Config.IMPRIMIR_LISTA_REFERENCIAS:
        return
    output.print_md("  **Parede {} - referencias ordenadas:**".format(wall_id))
    for it in pontos_sub:
        output.print_md("  - {}".format(_descrever_referencia_parede(it, p_ini, p_fim, tol)))


def identificar_paredes_exteriores(elementos):
    """Retorna o set de ElementId.IntegerValue das paredes cuja Function
    (parametro nativo do Revit, o mesmo usado no filtro 'Exterior/Interior'
    do proprio software) e' Exterior. Usado pra decidir quais alinhamentos
    sao 'perimetro do predio' (cota deve ir pra fora da casa) versus
    alinhamentos internos (cota fica colada no proprio elemento, onde ja
    estava). Se a leitura da Function falhar por qualquer motivo, a
    parede e' tratada como interior (comportamento antigo, mais seguro)."""
    exteriores = set()
    for el in elementos:
        if not isinstance(el, Wall):
            continue
        try:
            wt = el.WallType
            if wt is not None and wt.Function == WallFunction.Exterior:
                exteriores.add(el.Id.IntegerValue)
        except Exception as e:
            logger.debug("Nao foi possivel ler Function da parede {}: {}".format(
                el.Id.IntegerValue, e))
    return exteriores


def _eh_escada(el):
    """True para escadas, lances/patamares ou familias/categorias de escada."""
    try:
        cat = el.Category
        if cat is None:
            return False
        cat_id = cat.Id.IntegerValue
        nomes_bic = ("OST_Stairs", "OST_StairsRuns", "OST_StairsLandings",
                     "OST_StairsSupports", "OST_StairsRailing")
        for nome_bic in nomes_bic:
            try:
                if cat_id == int(getattr(BuiltInCategory, nome_bic)):
                    return True
            except Exception:
                pass
        nome = (cat.Name or "").lower()
        return ("escad" in nome) or ("stair" in nome)
    except Exception:
        return False


def filtrar_elementos_cotaveis(elementos):
    filtrados = []
    escadas = 0
    for el in elementos:
        if _eh_escada(el):
            escadas += 1
            continue
        filtrados.append(el)
    if escadas:
        output.print_md("**{} elemento(s) de escada ignorado(s)** para evitar bagunca nas cotas.".format(escadas))
    return filtrados


# ============================================================
# ETAPA 1 - Selecao de paredes + coleta de elementos relacionados
# ============================================================
class _FiltroSelecaoParede(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Wall)

    def AllowReference(self, reference, position):
        return False


def selecionar_paredes_alvo():
    output.print_md("## Cotar Selecao - selecione as paredes e pressione Enter")
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _FiltroSelecaoParede(),
            "Selecione uma ou varias paredes para cotar e pressione Enter",
        )
    except OperationCanceledException:
        forms.alert("Selecao cancelada. Nenhuma cota foi criada.", exitscript=True)
        return []
    except Exception as e:
        logger.error("Falha na selecao de paredes: {}".format(e))
        forms.alert("Falha na selecao de paredes:\n{}".format(e), exitscript=True)
        return []
    

    paredes = []
    vistos = set()
    for r in refs:
        wid = r.ElementId.IntegerValue
        if wid in vistos:
            continue
        el = doc.GetElement(r.ElementId)
        if isinstance(el, Wall):
            paredes.append(el)
            vistos.add(wid)

    if not paredes:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)
    return paredes


def coletar_elementos_relacionados(view, paredes_selecionadas):
    wall_ids = set(w.Id.IntegerValue for w in paredes_selecionadas)
    elementos = list(paredes_selecionadas)

    try:
        candidatos = list(
            FilteredElementCollector(doc, view.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.error("Falha ao coletar elementos da vista: {}".format(e))
        forms.alert("Falha ao coletar elementos da vista:\n{}".format(e), exitscript=True)
        return [], []

    hospedados = []
    qtd_portas, qtd_janelas, qtd_outros = 0, 0, 0
    for el in candidatos:
        if isinstance(el, Wall):
            continue
        try:
            host_id = _obter_host_id(el)
            if host_id not in wall_ids:
                continue
            hospedados.append(el)
            tipo = _tipo_referencia_do_elemento(el, host_id)
            if tipo == "porta":
                qtd_portas += 1
            elif tipo == "janela":
                qtd_janelas += 1
            else:
                qtd_outros += 1
        except Exception as e:
            logger.debug("Falha ao avaliar elemento {}: {}".format(el.Id.IntegerValue, e))

    vistos = set(wall_ids)
    for el in hospedados:
        eid = el.Id.IntegerValue
        if eid in vistos:
            continue
        elementos.append(el)
        vistos.add(eid)

    try:
        paredes_contexto = list(
            FilteredElementCollector(doc, view.Id)
            .OfClass(Wall)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        logger.error("Falha ao coletar paredes de contexto: {}".format(e))
        forms.alert("Falha ao coletar paredes de contexto:\n{}".format(e), exitscript=True)
        return [], []

    if not paredes_contexto:
        paredes_contexto = list(paredes_selecionadas)

    output.print_md(
        "**{} parede(s) selecionada(s)** | **{} porta(s)** | **{} janela(s)** | "
        "**{} abertura(s) hospedada(s) adicional(is)**".format(
            len(paredes_selecionadas), qtd_portas, qtd_janelas, qtd_outros
        )
    )
    return elementos, paredes_contexto


# ============================================================
# ETAPA 2 - Extracao de faces referenciaveis (com cache + host_id)
# ============================================================
_geom_cache = {}  # ElementId.IntegerValue -> lista de Solids ja extraidos


def _solidos_do_elemento(element, opt):
    key = element.Id.IntegerValue
    if key in _geom_cache:
        return _geom_cache[key]

    solidos = []
    try:
        geom = element.get_Geometry(opt)
    except Exception as e:
        logger.debug("Falha ao ler geometria de {}: {}".format(key, e))
        _geom_cache[key] = solidos
        return solidos

    if geom is None:
        _geom_cache[key] = solidos
        return solidos

    for g in geom:
        if isinstance(g, Solid) and g.Volume > 0:
            solidos.append(g)
        elif hasattr(g, "GetInstanceGeometry"):
            try:
                for g2 in g.GetInstanceGeometry():
                    if isinstance(g2, Solid) and g2.Volume > 0:
                        solidos.append(g2)
            except Exception as e:
                logger.debug("Falha ao ler GeometryInstance de {}: {}".format(key, e))

    _geom_cache[key] = solidos
    return solidos


def _area_minima_ft2():
    return Config.AREA_MINIMA_FACE_M2 * 10.7639


def _obter_host_id(el):
    """'Dono' geometrico do elemento, usado para a cota GERAL DE CADA
    PAREDE: Wall -> o proprio Id; porta/janela hospedada -> Id da
    parede-host; demais (piso) -> None (entram so no alinhamento
    por posicao, sem virar 'parede individual')."""
    if isinstance(el, Wall):
        return el.Id.IntegerValue
    if isinstance(el, FamilyInstance):
        try:
            host = el.Host
            if host is not None:
                return host.Id.IntegerValue
        except Exception as e:
            logger.debug("Sem Host para {}: {}".format(el.Id.IntegerValue, e))
    try:
        host = getattr(el, "Host", None)
        if host is not None:
            return host.Id.IntegerValue
    except Exception:
        pass
    return None


def _ponto_referencial_face(face):
    """[Fase 1 - correcao real] 'PlanarFace.Origin' NAO e garantidamente
    um ponto dentro dos limites visiveis da face - e' so um ponto usado
    pela API para definir o plano/sistema de coordenadas. Em faces de
    topo/ponta geradas por um encontro (miter/join) com outra parede,
    isso pode devolver um ponto bem longe da face real, ao ponto de
    coincidir com a posicao de uma parede vizinha completamente diferente
    - foi exatamente essa a causa da cota "subir" pra longe da parede
    selecionada. Aqui usamos o MEIO (em UV) da bounding box da propria
    face e avaliamos nela, o que garante um ponto realmente sobre a
    face. Se a avaliacao falhar por qualquer motivo (face degenerada),
    cai de volta no Origin (comportamento antigo) em vez de descartar a
    face inteira."""
    try:
        bb = face.GetBoundingBox()
        uv_meio = UV((bb.Min.U + bb.Max.U) / 2.0, (bb.Min.V + bb.Max.V) / 2.0)
        return face.Evaluate(uv_meio)
    except Exception as e:
        logger.debug("Falha ao avaliar ponto real da face - usando Origin como fallback: {}".format(e))
        return face.Origin


def extrair_faces_referenciaveis(elementos, axis_dir, perp_dir, threshold=0.985):
    """Retorna lista de dicts {pos_axis, pos_perp, ref, host_id} para cada
    face plana referenciavel alinhada ao eixo escolhido. Qualquer falha de
    geometria em UM elemento e reportada e pulada - nunca derruba a
    execucao inteira."""
    opt = Options()
    opt.ComputeReferences = True
    area_min = _area_minima_ft2()

    resultado = []
    for el in elementos:
        host_id = _obter_host_id(el)
        ref_tipo = _tipo_referencia_do_elemento(el, host_id)
        try:
            solidos = _solidos_do_elemento(el, opt)
        except Exception as e:
            output.print_md("  [AVISO] elemento {} ignorado (erro de geometria): {}".format(
                el.Id.IntegerValue, e))
            continue

        for solid in solidos:
            try:
                faces = solid.Faces
            except Exception as e:
                logger.debug("Solido sem faces legiveis: {}".format(e))
                continue
            for face in faces:
                try:
                    if face.Reference is None or not isinstance(face, PlanarFace):
                        continue
                    if face.Area < area_min:
                        continue
                    d = dot(face.FaceNormal, axis_dir)
                    if abs(d) <= threshold:
                        continue
                    ponto_ref = _ponto_referencial_face(face)
                    pos_axis = dot(ponto_ref, axis_dir)
                    pos_perp = dot(ponto_ref, perp_dir)
                    resultado.append({
                        "pos_axis": pos_axis, "pos_perp": pos_perp,
                        "ref": face.Reference, "host_id": host_id,
                        "normal_axis": 1 if d >= 0 else -1,
                        "ref_tipo": ref_tipo,
                        "source_element_id": el.Id.IntegerValue,
                    })
                except Exception as e:
                    logger.debug("Face ignorada por erro: {}".format(e))
                    continue

    return resultado


def _referencia_centro_abertura(el, axis, host_id, ref_tipo_base):
    """[v2.1] Referencia opcional do CENTRO da abertura (porta/janela),
    usando o reference plane nativo 'Center (Left/Right)' da familia,
    quando existir. So roda se Config.USAR_CENTRO_ABERTURA estiver ligado.
    Nunca derruba a execucao - qualquer falha (familia sem esse plano,
    API antiga etc.) so descarta esse ponto extra, sem afetar o resto da
    cadeia (inicio/fim da abertura continuam vindo das faces normais)."""
    if not Config.USAR_CENTRO_ABERTURA or not _TEM_CENTRO_ABERTURA_API:
        return None
    if not isinstance(el, FamilyInstance):
        return None
    try:
        refs = list(el.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))
    except Exception as e:
        logger.debug("Sem referencia de centro para {}: {}".format(el.Id.IntegerValue, e))
        return None
    if not refs:
        return None
    try:
        bbox = el.get_BoundingBox(None)
        if bbox is None:
            return None
        centro = XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
        pos_axis = dot(centro, axis)
    except Exception as e:
        logger.debug("Falha ao calcular centro de {}: {}".format(el.Id.IntegerValue, e))
        return None

    tipo_centro = "porta_centro" if ref_tipo_base == "porta" else (
        "janela_centro" if ref_tipo_base == "janela" else None)
    if tipo_centro is None:
        return None

    return {
        "pos_axis": pos_axis, "pos_perp": None,
        "ref": refs[0], "host_id": host_id,
        "normal_axis": 0, "ref_tipo": tipo_centro,
        "source_element_id": el.Id.IntegerValue,
    }


def coletar_centros_abertura(elementos, axis, perp_dir):
    """Varre os elementos hospedados (portas/janelas) e devolve os pontos
    de centro opcionais, com pos_perp preenchido (projetado a partir do
    proprio host_id, reaproveitando a mesma convencao das demais faces)."""
    if not Config.USAR_CENTRO_ABERTURA or not _TEM_CENTRO_ABERTURA_API:
        return []
    extras = []
    for el in elementos:
        host_id = _obter_host_id(el)
        ref_tipo_base = _tipo_referencia_do_elemento(el, host_id)
        if ref_tipo_base not in ("porta", "janela"):
            continue
        item = _referencia_centro_abertura(el, axis, host_id, ref_tipo_base)
        if item is None:
            continue
        try:
            bbox = el.get_BoundingBox(None)
            centro = XYZ(
                (bbox.Min.X + bbox.Max.X) / 2.0,
                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                (bbox.Min.Z + bbox.Max.Z) / 2.0,
            )
            item["pos_perp"] = dot(centro, perp_dir)
        except Exception:
            item["pos_perp"] = 0.0
        extras.append(item)
    return extras


# ============================================================
# ETAPA 3 - Deduplicacao por posicao + agrupamento por alinhamento
# ============================================================
def dedupe_por_posicao(itens, tol):
    """itens: lista de dicts ordenada por pos_axis. Remove faces
    praticamente coincidentes no eixo de cota (mesma posicao)."""
    if not itens:
        return []
    aceitos = [itens[0]]
    for it in itens[1:]:
        diff = abs(it["pos_axis"] - aceitos[-1]["pos_axis"])
        if diff > tol:
            aceitos.append(it)
        else:
            logger.debug("Face a {:.2f}cm da anterior - descartada (duplicata)".format(to_cm(diff)))
    return aceitos


def agrupar_por_alinhamento(itens, tol):
    """Agrupa faces por proximidade da coordenada PERPENDICULAR ao eixo -
    cada grupo = um alinhamento (uma ou mais paredes em fileira)."""
    if not itens:
        return []
    ordenado = sorted(itens, key=lambda t: t["pos_perp"])
    grupos = [[ordenado[0]]]
    for item in ordenado[1:]:
        if abs(item["pos_perp"] - grupos[-1][-1]["pos_perp"]) <= tol:
            grupos[-1].append(item)
        else:
            grupos.append([item])
    return grupos


def montar_correntes(itens, tol_cluster, tol_dedup, paredes_exteriores):
    """Cada corrente = {'itens': [...ordenados por pos_axis, dedup...],
    'perp': media, 'host_ids': set de paredes presentes,
    'perimetro': True se alguma parede da corrente for Exterior}."""
    grupos = agrupar_por_alinhamento(itens, tol_cluster)
    correntes = []
    for grupo in grupos:
        grupo_ordenado = sorted(grupo, key=lambda t: t["pos_axis"])
        grupo_dedup = dedupe_por_posicao(grupo_ordenado, tol_dedup)
        if len(grupo_dedup) < 2:
            continue
        perp_medio = sum(t["pos_perp"] for t in grupo_dedup) / len(grupo_dedup)
        host_ids = set(t["host_id"] for t in grupo_ordenado if t["host_id"] is not None)
        perimetro = bool(host_ids & paredes_exteriores)
        correntes.append({
            "itens": grupo_dedup, "perp": perp_medio,
            "itens_raw": grupo_ordenado,
            "host_ids": host_ids, "perimetro": perimetro,
        })

    # Fallback: muitos modelos nao marcam WallType.Function como Exterior.
    # Nesses casos, trata os alinhamentos mais externos de cada eixo como
    # perimetro para empurrar as cotas para fora da planta.
    correntes_com_parede = [c for c in correntes if c["host_ids"]]
    if correntes_com_parede:
        perps = [c["perp"] for c in correntes_com_parede]
        p_min, p_max = min(perps), max(perps)
        tol_borda = tol_cluster * 0.4
        for c in correntes_com_parede:
            if abs(c["perp"] - p_min) <= tol_borda or abs(c["perp"] - p_max) <= tol_borda:
                c["perimetro"] = True
    return correntes


# ============================================================
# ETAPA 3.5 - Interseccoes/cruzamentos (T, L, X, colineares nas pontas)
# ============================================================
def _linha_da_parede(wall):
    """Retorna (p0, p1) da LocationCurve da parede, so se for reta (Line).
    Paredes curvas ou sem LocationCurve sao ignoradas (retorna None) -
    nao entram na logica de cruzamento, mas continuam cotadas normalmente
    pelo resto do pipeline."""
    try:
        loc = wall.Location
        curve = getattr(loc, "Curve", None)
        if curve is None or not isinstance(curve, Line):
            return None
        return curve.GetEndPoint(0), curve.GetEndPoint(1)
    except Exception as e:
        logger.debug("Sem LocationCurve reta para parede {}: {}".format(
            getattr(wall.Id, "IntegerValue", "?"), e))
        return None


def _intersecao_2d(p1, d1n, p2, d2n):
    """Interseccao de duas retas (no plano XY) definidas por ponto+direcao
    UNITARIA. Retorna (t, s) = distancia ao longo de d1n/d2n ate o ponto de
    encontro, ou None se forem paralelas."""
    denom = d1n.X * d2n.Y - d1n.Y * d2n.X
    if abs(denom) < 1e-9:
        return None
    dx = p2.X - p1.X
    dy = p2.Y - p1.Y
    t = (dx * d2n.Y - dy * d2n.X) / denom
    s = (dx * d1n.Y - dy * d1n.X) / denom
    return t, s


def _cruzamentos_perpendiculares(host_wall, candidatas, perp_dir, tol_extensao):
    """Acha, entre 'candidatas', as paredes cuja linha de eixo e'
    aproximadamente perpendicular ao eixo de cota (alinhada com perp_dir)
    E cruza a extensao da parede host_wall (com folga de tol_extensao pra
    pegar cantos/juntas no limite - cobre T no meio da parede, L/X nas
    pontas e prolongamentos colineares que encostam bem na extremidade).
    Retorna lista de (wall, ponto_3d)."""
    resultado = []
    linha_host = _linha_da_parede(host_wall)
    if linha_host is None:
        return resultado
    p0, p1 = linha_host
    d1 = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
    len1 = d1.GetLength()
    if len1 < 1e-6:
        return resultado
    d1n = d1.Normalize()

    for w in candidatas:
        if w.Id.IntegerValue == host_wall.Id.IntegerValue:
            continue
        linha_w = _linha_da_parede(w)
        if linha_w is None:
            continue
        q0, q1 = linha_w
        d2 = XYZ(q1.X - q0.X, q1.Y - q0.Y, 0.0)
        len2 = d2.GetLength()
        if len2 < 1e-6:
            continue
        d2n = d2.Normalize()

        # a parede candidata precisa ser ~perpendicular ao eixo de cota
        # (ou seja, alinhada com a direcao perpendicular ao eixo)
        if abs(dot(d2n, perp_dir)) < 0.8:
            continue

        inter = _intersecao_2d(p0, d1n, q0, d2n)
        if inter is None:
            continue
        t, s = inter
        if t < -tol_extensao or t > len1 + tol_extensao:
            continue
        if s < -tol_extensao or s > len2 + tol_extensao:
            continue

        ponto = XYZ(p0.X + d1n.X * t, p0.Y + d1n.Y * t, p0.Z)
        resultado.append((w, ponto))

    return resultado


def adicionar_cruzamentos_perpendiculares(correntes, itens_todos, elementos, axis, perp, tolz):
    """Para cada corrente, procura paredes perpendiculares que cruzam
    alguma das paredes-host dessa corrente (no meio - T/X - ou na ponta -
    L/colinear) e injeta um ponto de referencia extra ali - reaproveitando
    a face que 'extrair_faces_referenciaveis' ja extraiu para essa parede
    perpendicular nesse mesmo eixo (nunca cria Reference nova). O ponto
    injetado e' marcado como ref_tipo='intersecao' para entrar na
    prioridade certa da cadeia.

    Isso NUNCA corta a corrente: so adiciona pontos intermediarios, entao
    a cota total (parede_total/alinhamento_total) continua ponta-a-ponta
    igual antes; so o nivel 'vaos' ganha mais segmentos (sub-cotas)."""
    paredes_por_id = {}
    for el in elementos:
        if isinstance(el, Wall):
            paredes_por_id[el.Id.IntegerValue] = el
    todas_paredes = list(paredes_por_id.values())

    total_adicionados = 0
    for c in correntes:
        itens_c = c["itens"]
        host_ids = sorted(c.get("host_ids", set()))
        pontos_add = []

        for hid in host_ids:
            host_wall = paredes_por_id.get(hid)
            if host_wall is None:
                continue

            try:
                cruzamentos = _cruzamentos_perpendiculares(
                    host_wall, todas_paredes, perp, tolz.tol_cruzamento_extensao
                )
            except Exception as e:
                logger.debug("Falha ao buscar cruzamentos da parede {}: {}".format(hid, e))
                continue

            for w, ponto in cruzamentos:
                pos_axis_cruz = dot(ponto, axis)
                wid = w.Id.IntegerValue

                candidatos = [it for it in itens_todos if it["host_id"] == wid]
                if not candidatos:
                    continue

                melhor = min(candidatos, key=lambda it: abs(it["pos_axis"] - pos_axis_cruz))

                try:
                    largura = w.Width  # ja vem em pes (unidade interna)
                    limite = largura * Config.FATOR_LARGURA_CRUZAMENTO
                except Exception:
                    limite = tolz.largura_cruzamento_padrao

                if abs(melhor["pos_axis"] - pos_axis_cruz) > limite:
                    # face mais proxima esta longe demais do cruzamento
                    # real - provavelmente pegou a ponta errada da parede,
                    # entao descarta pra nao inventar uma sub-cota errada.
                    continue

                # marca como intersecao (senao ficaria com ref_tipo
                # 'parede' generico e perderia prioridade/descricao corretas)
                melhor = dict(melhor)
                melhor["ref_tipo"] = "intersecao"
                melhor["wall_intersecao_id"] = hid
                pontos_add.append(melhor)

        if pontos_add:
            todos = itens_c + pontos_add
            todos_ordenados = sorted(todos, key=lambda t: t["pos_axis"])
            novo = ordenar_deduplicar_por_prioridade(todos_ordenados, tolz.tol_dim_zero)
            total_adicionados += max(0, len(novo) - len(itens_c))
            c["itens"] = novo

    if total_adicionados:
        output.print_md("  [INFO] {} ponto(s) de intersecao adicionado(s) (sub-cotas).".format(
            total_adicionados))

    return correntes


# ============================================================
# ETAPA 4 - Deduplicacao GLOBAL por referencia estavel (persistente)
# ============================================================
def stable_key(ref):
    """Chave estavel de uma Reference - usada pra saber se duas cotas
    (desta execucao ou de uma anterior, ja no modelo) apontam para
    exatamente a mesma geometria."""
    if ref is None:
        return None
    try:
        return ref.ConvertToStableRepresentation(doc)
    except Exception as e:
        logger.debug("Falha ConvertToStableRepresentation: {}".format(e))
        return None


def assinatura_da_tarefa(itens):
    chaves = []
    for it in itens:
        k = stable_key(it["ref"])
        if k is None:
            return None
        chaves.append(k)
    return frozenset(chaves)


def assinatura_par(it_a, it_b):
    ka = stable_key(it_a["ref"])
    kb = stable_key(it_b["ref"])
    if ka is None or kb is None:
        return None
    return frozenset([ka, kb])


def coletar_assinaturas_existentes(view):
    """Le as Dimension JA existentes na vista (de execucoes anteriores,
    inclusive manuais) e monta o conjunto de assinaturas, pra nunca
    recriar uma cota que ja existe."""
    existentes = set()
    try:
        dims = FilteredElementCollector(doc, view.Id).OfClass(Dimension).ToElements()
    except Exception as e:
        logger.error("Falha ao ler Dimension existentes na view: {}".format(e))
        return existentes

    for d in dims:
        try:
            refs = d.References
        except Exception:
            continue
        chaves = []
        ok = True
        for i in range(refs.Size):
            k = stable_key(refs.get_Item(i))
            if k is None:
                ok = False
                break
            chaves.append(k)
        if ok and chaves:
            existentes.add(frozenset(chaves))
    return existentes


def _filtrar_pontos_subcota_parede(pontos, wall_id, p_ini, p_fim, tolz):
    """Limpa a corrente de detalhe de uma parede.

    Quando existem faces internas perto das extremidades, usa essas faces
    como inicio/fim da sub-cota. Isso gera a cota util (ex.: 750 dentro de
    770) sem criar os pedacos 10 + 750 + 10.

    [v2.1] Agora usa ordenar_deduplicar_por_prioridade em vez do dedupe
    simples, para que, quando duas referencias caem no mesmo ponto (ex.:
    uma intersecao bem colada na extremidade), a de maior prioridade
    (extremidade > intersecao > porta > janela > abertura > parede) seja a
    mantida em vez de qualquer uma na ordem de chegada.
    """
    lo = min(p_ini["pos_axis"], p_fim["pos_axis"])
    hi = max(p_ini["pos_axis"], p_fim["pos_axis"])
    ordenados = ordenar_deduplicar_por_prioridade(
        sorted([it for it in pontos if it["host_id"] is not None], key=lambda t: t["pos_axis"]),
        tolz.tol_dim_zero
    )
    if len(ordenados) <= 2:
        return ordenados

    inicio = ordenados[0]
    fim = ordenados[-1]

    for it in ordenados[1:]:
        dist = it["pos_axis"] - lo
        if dist > tolz.tol_dim_zero and dist <= tolz.subcota_face_interna_max:
            inicio = it
            break

    for it in reversed(ordenados[:-1]):
        dist = hi - it["pos_axis"]
        if dist > tolz.tol_dim_zero and dist <= tolz.subcota_face_interna_max:
            fim = it
            break

    if inicio["pos_axis"] >= fim["pos_axis"]:
        return [ordenados[0], ordenados[-1]]

    candidatos = [
        it for it in ordenados
        if inicio["pos_axis"] - tolz.tol_dim_zero <= it["pos_axis"] <= fim["pos_axis"] + tolz.tol_dim_zero
    ]

    if len(candidatos) <= 2:
        return candidatos

    limpos = [candidatos[0]]
    for it in candidatos[1:-1]:
        if abs(it["pos_axis"] - limpos[-1]["pos_axis"]) <= tolz.subcota_min_segmento:
            continue
        limpos.append(it)

    ultimo = candidatos[-1]
    if abs(ultimo["pos_axis"] - limpos[-1]["pos_axis"]) <= tolz.subcota_min_segmento:
        if len(limpos) > 1:
            limpos.pop()
    limpos.append(ultimo)
    return limpos


# ============================================================
# ETAPA 5 - Geracao das tarefas de cota (vaos / parede / alinhamento / geral)
# ============================================================
def gerar_tarefas_de_cota(correntes, itens_todos, tolz):
    """Camadas por alinhamento - NOTE: 'parede_total' (cota geral de cada
    parede) e as 'vaos' de cada parede NAO sao geradas aqui. Desde a v2.0
    isso e' responsabilidade exclusiva de
    'processar_paredes_individualmente', que trata cada parede como
    unidade independente (garante cota geral mesmo pra parede isolada,
    que antes ficava sem cota por falta de corrente valida).

    O que esta funcao ainda gera, por corrente/alinhamento:
      nivel 0 'vaos'            - so quando a corrente NAO tem nenhuma
                                   parede identificavel (ex.: piso) e tem
                                   >2 referencias;
      'alinhamento_total'       - cota de conjunto da fileira (varias
                                   paredes/elementos alinhados), gerada
                                   DEPOIS que cada parede ja tem sua
                                   propria cota (PASSO 5 da estrategia) -
                                   nunca substitui a cota individual da
                                   parede;
    Fora do loop por corrente:
      nivel 3 'geral'           - uma por eixo, por fora de tudo.
    """
    tarefas = []
    for c in correntes:
        itens_c = c["itens"]
        perimetro = c["perimetro"]
        paredes_na_corrente = sorted(c.get("host_ids", set()))

        # parede_total/vaos por parede NAO sao decididos aqui. Toda parede
        # e' processada de forma independente em
        # processar_paredes_individualmente (PASSO 1-4): ela SEMPRE recebe
        # sua cota geral, mesmo se estiver sozinha nesta corrente (o que
        # antes fazia a corrente inteira ser descartada). A corrente aqui
        # so serve para decidir alinhamento_total/geral (PASSO 5: cotas de
        # conjunto, depois de todas as paredes prontas).

        if not paredes_na_corrente:
            tem_detalhe = len(itens_c) > 2
            if tem_detalhe:
                tarefas.append({
                    "nome": "vaos",
                    "itens": itens_c,
                    "perp_ref": c["perp"], "nivel": 0, "perimetro": perimetro,
                })

        # Cotas grandes internas: quando o alinhamento tem paredes internas,
        # cria uma cota total da fileira dentro da planta. Para perimetro,
        # evita repetir a cota externa da propria parede.
        if paredes_na_corrente and not perimetro and len(itens_c) >= 2:
            tarefas.append({
                "nome": "alinhamento_total",
                "itens": [itens_c[0], itens_c[-1]],
                "perp_ref": c["perp"], "nivel": 2,
                "perimetro": perimetro,
            })

        # Sem paredes identificaveis: usa uma unica cota de alinhamento.
        if not paredes_na_corrente:
            # nenhuma face pertence a uma parede identificavel (ex.: so
            # piso relacionado) - ainda assim cota o alinhamento
            # inteiro, no nivel imediatamente acima do detalhe.
            tarefas.append({
                "nome": "alinhamento_total",
                "itens": [itens_c[0], itens_c[-1]],
                "perp_ref": c["perp"], "nivel": 1 if len(itens_c) > 2 else 0,
                "perimetro": perimetro,
            })

    if itens_todos:
        global_ordenado = sorted(itens_todos, key=lambda t: t["pos_axis"])
        global_dedup = dedupe_por_posicao(global_ordenado, tolz.tol_dim_zero)
        if len(global_dedup) >= 2:
            tarefas.append({
                "nome": "geral", "itens": [global_dedup[0], global_dedup[-1]],
                "perp_ref": None, "nivel": None,
            })

    return tarefas


def remover_tarefas_duplicadas(tarefas, assinaturas_existentes):
    """Remove tarefas cuja assinatura de referencias ja apareceu - seja em
    Dimension ja existente na view (execucao anterior/manual), seja em
    outra tarefa desta mesma execucao (ex.: parede sozinha no alinhamento
    -> parede_total == alinhamento_total; ou os dois eixos H/V gerando a
    mesma cota por coincidencia)."""
    vistos = set(assinaturas_existentes)
    resultado = []
    puladas = 0
    for t in tarefas:
        if t.get("nome") == "vaos" and len(t.get("itens", [])) > 2:
            itens_limpos = [t["itens"][0]]
            for idx, it in enumerate(t["itens"][1:], 1):
                par = assinatura_par(itens_limpos[-1], it)
                eh_ultimo = idx == len(t["itens"]) - 1
                if par is not None and par in vistos:
                    if eh_ultimo and len(itens_limpos) > 1:
                        itens_limpos.pop()
                    elif not eh_ultimo:
                        continue
                itens_limpos.append(it)
            if len(itens_limpos) < 3:
                puladas += 1
                continue
            t["itens"] = itens_limpos

        assinatura = assinatura_da_tarefa(t["itens"])
        if assinatura is None:
            # sem referencia estavel valida - deixa passar (sera pego no
            # try/except da criacao, que ja reporta erro por referencia
            # ausente/invalida) em vez de descartar silenciosamente.
            resultado.append(t)
            continue
        if assinatura in vistos:
            puladas += 1
            continue
        vistos.add(assinatura)
        for i in range(len(t["itens"]) - 1):
            par = assinatura_par(t["itens"][i], t["itens"][i + 1])
            if par is not None:
                vistos.add(par)
        resultado.append(t)
    if puladas:
        output.print_md("  [INFO] {} tarefa(s) descartada(s) por ja existirem (duplicata real).".format(puladas))
    return resultado


# ============================================================
# ETAPA 6 - Layout (posicao perpendicular final de cada tarefa)
# ============================================================
def resolver_layout(tarefas, centro_perp_modelo, tolz, itens=None):
    """Decide o LADO/posicao final de cada alinhamento, sem precisar
    clicar um ponto:

    - Alinhamentos de PERIMETRO (tocam parede com Function=Exterior):
      ancorados no extremo REAL do predio naquele eixo (o menor/maior
      pos_perp entre TODAS as faces de paredes exteriores) - a cota fica
      colada por fora da casa, perto da parede, nao 'flutuando' pra
      dentro so porque um elemento interno prox puxou a media pro lado
      errado. Empilha nivel 0 (vaos/pedacos) -> 1 (parede) -> 2
      (alinhamento) nessa ordem, sempre se afastando mais da casa.

    - Alinhamentos internos (sem parede exterior): mantem o criterio
      antigo (lado que fica mais longe do centro do modelo), colado no
      proprio elemento - e' o esperado, ja que os pontos referenciados
      estao mesmo dentro da planta.
    """
    perp_por_alinhamento = {}
    for t in tarefas:
        if t["nome"] == "geral":
            continue
        perp_por_alinhamento.setdefault(t["perp_ref"], []).append(t)

    # Extremos reais do PERIMETRO (so entre alinhamentos marcados como
    # perimetro=True) - referencia para "fora da casa".
    perps_perimetro = [
        perp_ref for perp_ref, lista in perp_por_alinhamento.items()
        if lista and lista[0]["perimetro"]
    ]
    perimetro_min = min(perps_perimetro) if perps_perimetro else None
    perimetro_max = max(perps_perimetro) if perps_perimetro else None

    perp_extremos = []
    for perp_ref, lista in perp_por_alinhamento.items():
        eh_perimetro = lista[0]["perimetro"]

        if eh_perimetro and perimetro_min is not None and perimetro_max is not None:
            dist_min = abs(perp_ref - perimetro_min)
            dist_max = abs(perimetro_max - perp_ref)
            if dist_min <= dist_max:
                sinal = -1.0
                extremo = perimetro_min
            else:
                sinal = 1.0
                extremo = perimetro_max
            for t in lista:
                perp_pos = extremo + sinal * (tolz.cola_elemento + t["nivel"] * tolz.gap_nivel)
                t["perp_pos"] = perp_pos
                perp_extremos.append((sinal, perp_pos))
        else:
            # comportamento antigo: lado mais longe do centro do modelo,
            # colado na propria posicao do alinhamento.
            sinal = 1.0 if perp_ref >= centro_perp_modelo else -1.0
            for t in lista:
                base = perp_ref + sinal * tolz.cola_elemento
                perp_pos = base + sinal * (t["nivel"] * tolz.gap_nivel)
                t["perp_pos"] = perp_pos
                perp_extremos.append((sinal, perp_pos))

    for t in tarefas:
        if t["nome"] != "geral":
            continue
        if not perp_extremos:
            # Sem nenhuma tarefa de alinhamento nesse eixo (ex.: parede
            # isolada, sem 'alinhamento_total'/'vaos' proprio) - antes
            # isso caia direto em tolz.gap_geral usado como COORDENADA
            # ABSOLUTA do projeto, sem nenhuma relacao com a geometria
            # real. Isso podia colocar a cota 'geral' proxima/em cima da
            # cota 'parede_total'/'vaos' daquela mesma parede (que e'
            # calculada em processar_paredes_individualmente e nao
            # aparece aqui em perp_extremos). Agora usa o mesmo padrao
            # das demais funcoes: extremo real (max/min pos_perp dos
            # itens do eixo) + afastamento gap_geral, do lado que tiver
            # mais elementos (mesmo criterio de sinal do centro do
            # modelo usado em outros lugares do pipeline).
            if itens:
                pos_perp_vals = [it["pos_perp"] for it in itens]
                lado_pos = sum(1 for p in pos_perp_vals if p >= centro_perp_modelo)
                lado_neg = len(pos_perp_vals) - lado_pos
                if lado_pos >= lado_neg:
                    sinal = 1.0
                    extremo = max(pos_perp_vals)
                else:
                    sinal = -1.0
                    extremo = min(pos_perp_vals)
                t["perp_pos"] = extremo + sinal * tolz.gap_geral
            else:
                # fallback do fallback - sem itens disponiveis, mantem o
                # comportamento antigo (nao deveria acontecer na pratica,
                # ja que processar_eixo sempre tem itens nesse ponto).
                t["perp_pos"] = tolz.gap_geral
            continue
        # geral vai por fora de tudo, no lado que tiver mais elementos
        pos_sinal = [p for s, p in perp_extremos if s > 0]
        neg_sinal = [p for s, p in perp_extremos if s < 0]
        if len(pos_sinal) >= len(neg_sinal):
            t["perp_pos"] = (max(pos_sinal) if pos_sinal else 0.0) + tolz.gap_geral
        else:
            t["perp_pos"] = (min(neg_sinal) if neg_sinal else 0.0) - tolz.gap_geral

    return tarefas


# ============================================================
# ETAPA 7 - Criacao das cotas no Revit
# ============================================================
def find_dim_type_by_name(nome):
    try:
        dtypes = list(FilteredElementCollector(doc).OfClass(DimensionType).ToElements())
    except Exception as e:
        logger.error("Falha ao coletar DimensionType: {}".format(e))
        return None
    for dt in dtypes:
        try:
            nm = dt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except Exception:
            nm = None
        if nm == nome:
            return dt
    if not dtypes:
        logger.debug("Nenhum DimensionType encontrado no documento.")
        return None
    output.print_md(
        "  [AVISO] tipo de cota '{}' nao encontrado - usando o primeiro disponivel.".format(nome)
    )
    return dtypes[0]


def _mpt(axis, perp, r, perp_pos):
    return XYZ(
        axis.X * r + perp.X * perp_pos,
        axis.Y * r + perp.Y * perp_pos,
        axis.Z * r + perp.Z * perp_pos,
    )


def _cria_dim_line(axis, perp, itens, perp_pos, margem):
    vals = [it["pos_axis"] for it in itens]
    r_min, r_max = min(vals) - margem, max(vals) + margem
    pt1, pt2 = _mpt(axis, perp, r_min, perp_pos), _mpt(axis, perp, r_max, perp_pos)
    if pt1.DistanceTo(pt2) < 1e-6:
        return None
    return Line.CreateBound(pt1, pt2)


def criar_cotas_no_revit(tarefas, view, axis, perp, tolz, dim_type):
    criadas, erros = 0, 0
    por_nome = {}
    for t in tarefas:
        try:
            axis_t = t.get("_axis", axis)
            perp_t = t.get("_perp", perp)
            dim_line = _cria_dim_line(axis_t, perp_t, t["itens"], t["perp_pos"], tolz.margem_ponta)
            if dim_line is None:
                logger.debug("Linha degenerada para tarefa '{}' - pulada.".format(t["nome"]))
                continue

            ra = ReferenceArray()
            valido = True
            for it in t["itens"]:
                if it["ref"] is None:
                    valido = False
                    break
                ra.Append(it["ref"])
            if not valido:
                output.print_md("[ERRO] tarefa '{}': referencia ausente, pulando.".format(t["nome"]))
                erros += 1
                continue

            with revit.Transaction("Cotar Parede Completo - {}".format(t["nome"])):
                nd = doc.Create.NewDimension(view, dim_line, ra)
                if dim_type:
                    try:
                        nd.DimensionType = dim_type
                    except Exception as e:
                        logger.debug("Falha ao aplicar DimensionType: {}".format(e))
            criadas += 1
            por_nome[t["nome"]] = por_nome.get(t["nome"], 0) + 1
        except Exception as e:
            erros += 1
            output.print_md("[ERRO] falha ao criar cota '{}': {}".format(t.get("nome", "?"), e))
    return criadas, erros, por_nome


# ============================================================
# MAIN
# ============================================================
def processar_eixo(elementos, view, axis, perp, nome_eixo, tolz, assinaturas_existentes, paredes_exteriores):
    """Roda o pipeline completo (extracao -> alinhamento -> cruzamentos ->
    tarefas -> dedup -> layout) para UM eixo (H ou V). Retorna a lista de
    tarefas prontas (com perp_pos definido) - a criacao no Revit e feita
    depois, juntando H + V numa unica transacao."""
    output.print_md("### Eixo {}".format(nome_eixo))

    itens = extrair_faces_referenciaveis(elementos, axis, perp)
    itens += coletar_centros_abertura(elementos, axis, perp)
    if len(itens) < 2:
        output.print_md("  [INFO] menos de 2 referencias nesse eixo - nada a cotar aqui.")
        return []

    correntes = montar_correntes(itens, tolz.cluster_tol, tolz.tol_dim_zero, paredes_exteriores)
    if not correntes:
        output.print_md("  [INFO] nenhum alinhamento valido nesse eixo.")
        return []

    n_perimetro = sum(1 for c in correntes if c["perimetro"])
    output.print_md("  {} alinhamento(s) encontrado(s) nesse eixo ({} de perimetro/exterior).".format(
        len(correntes), n_perimetro))

    # injeta pontos de intersecao (sub-cotas), sem cortar a corrente - so
    # acrescenta referencias no meio/pontas dela.
    correntes = adicionar_cruzamentos_perpendiculares(correntes, itens, elementos, axis, perp, tolz)

    tarefas = gerar_tarefas_de_cota(correntes, itens, tolz)
    tarefas = remover_tarefas_duplicadas(tarefas, assinaturas_existentes)

    centro_perp_modelo = sum(t["pos_perp"] for t in itens) / len(itens)
    tarefas = resolver_layout(tarefas, centro_perp_modelo, tolz, itens)

    for t in tarefas:
        t["_axis"] = axis
        t["_perp"] = perp
    return tarefas


def _direcao_parede_reta(wall):
    linha = _linha_da_parede(wall)
    if linha is None:
        return None
    p0, p1 = linha
    d = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
    if d.GetLength() < 1e-6:
        return None
    return d.Normalize()


def _parede_fora_hv(wall, axis_h, axis_v):
    d = _direcao_parede_reta(wall)
    if d is None:
        return False
    return abs(dot(d, axis_h)) < 0.985 and abs(dot(d, axis_v)) < 0.985


def _elementos_da_parede(elementos, wall_id):
    relacionados = []
    for el in elementos:
        if el.Id.IntegerValue == wall_id:
            relacionados.append(el)
            continue
        if _obter_host_id(el) == wall_id:
            relacionados.append(el)
    return relacionados


def processar_paredes_individualmente(
    elementos_alvo, paredes_alvo, paredes_contexto, view, tolz, assinaturas_existentes, paredes_exteriores
):
    """ROTINA PRINCIPAL DE COTAGEM - a unidade de trabalho e' a PAREDE, nao
    a corrente/alinhamento.

    Para CADA parede reta do modelo (H, V ou inclinada - nao importa a
    orientacao, cada uma usa seu proprio eixo local), monta a mesma cadeia
    de referencias que a Cota Automatica do Revit ofereceria:

      PASSO 1 - acha as duas faces extremas da parede (marcadas como
                'extremidade_parede'). Essas faces SEMPRE geram a cota da
                parede inteira (parede_total). Nenhuma parede fica sem
                essa cota.
      PASSO 2 - localiza as intersecoes (T, L, X, colineares na ponta) com
                paredes de contexto, marcadas como 'intersecao'.
      PASSO 3 - localiza aberturas hospedadas (portas/janelas/outras) e
                usa as faces de inicio/fim (largura) ja extraidas da
                geometria; opcionalmente adiciona o centro da abertura.
      PASSO 4 - ordena tudo pela posicao ao longo do EIXO da propria
                parede (prioridade: extremidade > intersecao > porta >
                janela > abertura > parede > outro) e monta UMA corrente
                de subcotas com todos esses pontos.

    As "correntes"/alinhamentos (ETAPA 5, gerar_tarefas_de_cota) nao
    decidem se uma parede sera cotada - isso e' feito 100% aqui, por
    parede, de forma independente. Correntes continuam existindo so para
    organizar alinhamento_total/geral (PASSO 5 da estrategia).
    """
    paredes = [w for w in paredes_alvo if _direcao_parede_reta(w) is not None]
    if not paredes:
        return []

    output.print_md("### Paredes (cotagem individual - unidade = parede)")
    tarefas = []
    sem_cota = []

    for wall in paredes:
        wid = wall.Id.IntegerValue
        eixo = _direcao_parede_reta(wall)
        if eixo is None:
            continue
        perp = XYZ(-eixo.Y, eixo.X, 0.0)
        relacionados = _elementos_da_parede(elementos_alvo, wid)

        # Busca de referencias (faces externas/internas, portas, janelas,
        # encontros, mudancas de espessura, rebaixos...) - tenta primeiro
        # com o threshold estrito de sempre; se a parede nao render pelo
        # menos 2 faces extremas com isso (geometria atipica), relaxa o
        # threshold antes de desistir - PASSO 1 exige que TODA parede
        # tenha cota geral.
        faces_parede = extrair_faces_referenciaveis([wall], eixo, perp, threshold=0.985)
        faces_todas = extrair_faces_referenciaveis(relacionados, eixo, perp, threshold=0.985)
        faces_contexto = extrair_faces_referenciaveis(paredes_contexto, eixo, perp, threshold=0.985)

        faces_parede_dedup_check = dedupe_por_posicao(
            sorted(faces_parede, key=lambda t: t["pos_axis"]), tolz.tol_dim_zero)
        if len(faces_parede_dedup_check) < 2:
            logger.debug("Parede {}: threshold estrito nao achou 2 faces extremas, relaxando.".format(wid))
            faces_parede = extrair_faces_referenciaveis([wall], eixo, perp, threshold=0.90)
            faces_todas = extrair_faces_referenciaveis(relacionados, eixo, perp, threshold=0.90)
            faces_contexto = extrair_faces_referenciaveis(paredes_contexto, eixo, perp, threshold=0.90)

        # PASSO 3 (extra opcional) - centro das aberturas hospedadas.
        faces_todas += coletar_centros_abertura(relacionados, eixo, perp)

        # PASSO 2 - intersecoes com paredes de contexto (T/L/X/colinear).
        try:
            cruzamentos = _cruzamentos_perpendiculares(
                wall, paredes_contexto, perp, tolz.tol_cruzamento_extensao
            )
        except Exception as e:
            logger.debug("Falha ao buscar cruzamentos inclinados da parede {}: {}".format(wid, e))
            cruzamentos = []

        for w_cruz, ponto in cruzamentos:
            pos_axis_cruz = dot(ponto, eixo)
            candidatos = [it for it in faces_contexto if it["host_id"] == w_cruz.Id.IntegerValue]
            if not candidatos:
                continue
            melhor = min(candidatos, key=lambda it: abs(it["pos_axis"] - pos_axis_cruz))
            try:
                limite = w_cruz.Width * Config.FATOR_LARGURA_CRUZAMENTO
            except Exception:
                limite = tolz.largura_cruzamento_padrao
            if abs(melhor["pos_axis"] - pos_axis_cruz) <= limite:
                melhor = dict(melhor)
                melhor["ref_tipo"] = "intersecao"
                melhor["wall_intersecao_id"] = w_cruz.Id.IntegerValue
                faces_todas.append(melhor)

        faces_parede = dedupe_por_posicao(sorted(faces_parede, key=lambda t: t["pos_axis"]), tolz.tol_dim_zero)
        faces_todas = ordenar_deduplicar_por_prioridade(
            sorted(faces_todas, key=lambda t: t["pos_axis"]), tolz.tol_dim_zero)
        if len(faces_parede) < 2:
            # Mesmo com threshold relaxado nao foi possivel achar 2 faces
            # referenciaveis (geometria muito atipica/curva/sem solido).
            # Nao ha como criar Dimension sem 2 referencias - registra o
            # aviso em vez de falhar silenciosamente.
            sem_cota.append(wid)
            continue

        # PASSO 1 - marca as duas faces extremas da propria parede como
        # 'extremidade_parede' (maxima prioridade na cadeia).
        p_ini, p_fim = faces_parede[0], faces_parede[-1]
        p_ini = dict(p_ini); p_ini["ref_tipo"] = "extremidade_parede"
        p_fim = dict(p_fim); p_fim["ref_tipo"] = "extremidade_parede"

        # garante que a copia com o tipo corrigido substitua a original em
        # faces_todas na mesma posicao (senao a extremidade fica com o
        # tipo generico 'parede' e perde prioridade no dedup por corte).
        def _com_extremidades_marcadas(lista, p_ini, p_fim, tol):
            nova = []
            usou_ini, usou_fim = False, False
            for it in lista:
                if not usou_ini and abs(it["pos_axis"] - p_ini["pos_axis"]) <= tol and it["host_id"] == p_ini["host_id"]:
                    nova.append(p_ini)
                    usou_ini = True
                    continue
                if not usou_fim and abs(it["pos_axis"] - p_fim["pos_axis"]) <= tol and it["host_id"] == p_fim["host_id"]:
                    nova.append(p_fim)
                    usou_fim = True
                    continue
                nova.append(it)
            return nova

        faces_todas = _com_extremidades_marcadas(faces_todas, p_ini, p_fim, tolz.tol_dim_zero)

        pontos_sub = _filtrar_pontos_subcota_parede(faces_todas, wid, p_ini, p_fim, tolz)
        tem_subcota = (
            len(pontos_sub) >= 2 and (
                len(pontos_sub) > 2 or
                abs(pontos_sub[0]["pos_axis"] - p_ini["pos_axis"]) > tolz.tol_dim_zero or
                abs(pontos_sub[-1]["pos_axis"] - p_fim["pos_axis"]) > tolz.tol_dim_zero
            )
        )

        # [v2.1] cadeia de referencias completa, ordenada por prioridade -
        # reaproveita os helpers que ja existiam mas nunca eram chamados.
        _imprimir_lista_referencias_parede(wid, pontos_sub, p_ini, p_fim, tolz.tol_dim_zero)

        perps_parede = [it["pos_perp"] for it in faces_parede]
        meio = sum(perps_parede) / len(perps_parede)

        # [Fase 1 - selecao] O lado da cota nao pode depender de paredes
        # distantes que nada tem a ver com esta (ex.: outra ala do predio,
        # so porque e' aproximadamente paralela). O "centro" usado para
        # decidir o lado agora so considera faces de contexto cujo
        # pos_axis cai dentro do proprio trecho desta parede (+ folga de
        # tolz.cluster_tol) - ou seja, so quem esta de fato "de frente"
        # para este pedaco (mesma sala/corredor), nao o predio inteiro.
        # Sem isso, uma parede selecionada "embaixo" podia herdar o centro
        # medio de paredes la em cima e a cota ia pro lado errado (longe
        # de onde voce selecionou).
        faixa_min = min(p_ini["pos_axis"], p_fim["pos_axis"]) - tolz.cluster_tol
        faixa_max = max(p_ini["pos_axis"], p_fim["pos_axis"]) + tolz.cluster_tol
        faces_contexto_local = [
            t for t in faces_contexto
            if faixa_min <= t["pos_axis"] <= faixa_max
        ]
        centro_perp_modelo = (
            sum(t["pos_perp"] for t in faces_contexto_local) / len(faces_contexto_local)
            if faces_contexto_local else meio
        )
        logger.debug(
            "Parede {}: centro de lado calculado com {}/{} face(s) de contexto "
            "(faixa axis [{:.1f}, {:.1f}]cm).".format(
                wid, len(faces_contexto_local), len(faces_contexto),
                to_cm(faixa_min), to_cm(faixa_max)))
        sinal = 1.0 if meio >= centro_perp_modelo else -1.0
        extremo = max(perps_parede) if sinal > 0 else min(perps_parede)

        total = {
            "nome": "parede_total",
            "itens": [p_ini, p_fim],
            "perp_ref": meio,
            "perp_pos": extremo + sinal * (tolz.cola_elemento + (tolz.gap_nivel if tem_subcota else 0.0)),
            "nivel": 1 if tem_subcota else 0,
            "perimetro": wid in paredes_exteriores,
            "_axis": eixo,
            "_perp": perp,
        }
        tarefas.append(total)

        if tem_subcota:
            tarefas.append({
                "nome": "vaos",
                "itens": pontos_sub,
                "perp_ref": meio,
                "perp_pos": extremo + sinal * tolz.cola_elemento,
                "nivel": 0,
                "perimetro": wid in paredes_exteriores,
                "_axis": eixo,
                "_perp": perp,
            })

    tarefas = remover_tarefas_duplicadas(tarefas, assinaturas_existentes)
    output.print_md("  {} tarefa(s) montada(s) para {} parede(s) processada(s) individualmente.".format(
        len(tarefas), len(paredes)))
    if sem_cota:
        output.print_md(
            "  [AVISO] {} parede(s) sem geometria referenciavel suficiente para cota "
            "(ids: {}).".format(len(sem_cota), ", ".join(str(i) for i in sem_cota)))
    return tarefas

 
def main():
    view = doc.ActiveView

    try:
        eixo_h = (view.RightDirection, view.UpDirection)
    except Exception as e:
        logger.error("Vista sem RightDirection/UpDirection utilizavel: {}".format(e))
        forms.alert(
            "Essa vista nao tem eixos H/V utilizaveis para cota (provavelmente "
            "nao e uma planta/elevacao/corte). Abra a view certa e rode de novo.",
            exitscript=True,
        )
        return

    tolz = Tolerancias(view)
    output.print_md(
        "Vista ativa: **{}** | Fator de escala: **{:.2f}x** (escala 1:{:.0f})".format(
            view.Name, tolz.fator, float(getattr(view, "Scale", None) or Config.ESCALA_BASE))
    )

    paredes_selecionadas = selecionar_paredes_alvo()
    elementos, paredes_contexto = coletar_elementos_relacionados(view, paredes_selecionadas)
    if not elementos:
        forms.alert(
            "Nao foi possivel coletar os elementos relacionados as paredes selecionadas.",
            exitscript=True,
        )
        return


    contexto_exteriores = identificar_paredes_exteriores(paredes_contexto)
    paredes_contexto_todas = paredes_contexto
    paredes_contexto = [w for w in paredes_contexto_todas if w.Id.IntegerValue in contexto_exteriores]
    output.print_md(
        "**Paredes de contexto (cruzamento/intersecao):** {} de {} sao Exterior - "
        "so essas contam como referencia extra; paredes internas sao ignoradas "
        "para esse fim.".format(len(paredes_contexto), len(paredes_contexto_todas))
    )

    paredes_exteriores = identificar_paredes_exteriores(paredes_selecionadas)
    output.print_md("**{} parede(s) identificada(s) como Exterior** (Function do WallType) - "
        "as cotas de parede/pedaco dessas vao pra fora da casa.".format(len(paredes_exteriores)))

    assinaturas_existentes = coletar_assinaturas_existentes(view)
    output.print_md("**{} assinatura(s) de cota ja existente(s)** na vista (nao serao repetidas).".format(
        len(assinaturas_existentes)))

    dim_type = find_dim_type_by_name(Config.NOME_TIPO_COTA_PADRAO)

    axis_h, perp_h = eixo_h
    axis_v, perp_v = perp_h, axis_h  # eixo V e' simplesmente o H trocado

    # PASSO 1-4 da estrategia: cada parede e' cotada de forma INDEPENDENTE
    # primeiro - toda parede sempre ganha sua cota geral (+ subcotas se
    # tiver detalhe), nao importa se esta isolada ou nao tem nenhuma outra
    # parede alinhada com ela.
    tarefas_paredes = processar_paredes_individualmente(
        elementos, paredes_selecionadas, paredes_contexto, view, tolz, assinaturas_existentes, paredes_exteriores
    )

    # PASSO 5: SOMENTE DEPOIS de todas as paredes cotadas individualmente,
    # gera cotas de conjunto (alinhamento_total/geral) por eixo H/V. Essas
    # nunca substituem a cota da parede - a corrente aqui e' so layout.
    tarefas_h = processar_eixo(elementos, view, axis_h, perp_h, "Horizontal", tolz, assinaturas_existentes, paredes_exteriores)
    tarefas_v = processar_eixo(elementos, view, axis_v, perp_v, "Vertical", tolz, assinaturas_existentes, paredes_exteriores)

    # Dedup GLOBAL entre as tres passagens: cada uma ja foi deduplicada
    # contra as cotas PRE-EXISTENTES na vista, mas uma 'alinhamento_total'
    # de uma corrente com uma unica parede pode coincidir exatamente com a
    # 'parede_total' ja criada para aquela parede - o passo abaixo garante
    # que isso nunca vira cota duplicada. tarefas_paredes vem PRIMEIRO na
    # lista, entao a cota da parede sempre tem prioridade sobre a de
    # alinhamento/geral (nunca o contrario).
    todas_tarefas = remover_tarefas_duplicadas(
        tarefas_paredes + tarefas_h + tarefas_v, assinaturas_existentes
    )
    if not todas_tarefas:
        forms.alert(
            "Nao foi possivel montar nenhuma tarefa de cota valida (paredes, H ou V).\n"
            "Verifique se os elementos coletados tem geometria solida normal.",
            exitscript=True,
        )
        return

    try:
        criadas_total, erros_total, por_nome_total = 0, 0, {}

        # Cria tudo numa unica passada - cada tarefa ja carrega seu proprio
        # _axis/_perp (global H/V ou local da parede), e cada Dimension e'
        # criada em sua propria transacao curta dentro de
        # criar_cotas_no_revit, entao um erro isolado nunca contamina as
        # demais tarefas.
        criadas, erros, por_nome = criar_cotas_no_revit(
            todas_tarefas, view, axis_h, perp_h, tolz, dim_type
        )
        criadas_total += criadas
        erros_total += erros
        for k, v in por_nome.items():
            por_nome_total[k] = por_nome_total.get(k, 0) + v
    except Exception as e:
        logger.error("Falha critica ao criar cotas: {}".format(e))
        forms.alert("Falha critica ao criar cotas:\n{}".format(e), exitscript=True)
        return

    output.print_md("---")
    output.print_md(
        "## {} cota(s) criada(s): {} vao(s)/detalhe + {} parede(s) individual(is) + "
        "{} alinhamento(s) + {} geral(is).".format(
            criadas_total,
            por_nome_total.get("vaos", 0),
            por_nome_total.get("parede_total", 0),
            por_nome_total.get("alinhamento_total", 0),
            por_nome_total.get("geral", 0),
        )
    )
    if erros_total:
        output.print_md("**{} tarefa(s) falharam ao criar cota** - ver [ERRO] acima.".format(erros_total))


if __name__ == "__main__":
    main()
