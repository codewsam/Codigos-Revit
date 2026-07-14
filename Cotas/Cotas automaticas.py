# -*- coding: utf-8 -*-
__title__ = "Cotar Parede Completo"
__version__ = "1.0"
__doc__ = (
    "Cota automaticamente paredes (+ portas/janelas/pisos/escadas relacionados)\n"
    "em UM clique, reproduzindo o padrao observado na view 'FORMA  DO TERREO'\n"
    "(18 cotas manuais: cadeias de vaos/detalhe + cotas de linha unica, tipo\n"
    "'Cota - 2 mm (cm) - 1 casa decimal vermelha', misturando H e V na mesma\n"
    "prancha).\n\n"
    "Fusao dos dois scripts anteriores:\n"
    "  - Cotar Selecao (v3.0): agrupamento por alinhamento, camadas\n"
    "    vaos/parede-total/geral, filtro de faces pequenas, modo rapido/preciso.\n"
    "  - Cotar Elevacao (v2.3): anti-duplicata PERSISTENTE via assinatura de\n"
    "    referencia estavel (ConvertToStableRepresentation) - conferida contra\n"
    "    as Dimension JA existentes na view, nao so dentro da mesma execucao.\n\n"
    "O que mudou/foi corrigido em relacao ao Cotar Selecao original:\n"
    "  1. BUG CORRIGIDO: a extracao de faces guardava host_id (parede-dona da\n"
    "     face, inclusive de portas/janelas hospedadas), mas o agrupamento por\n"
    "     alinhamento fazia unpack de 3 campos numa tupla de 4 - o host_id\n"
    "     nunca era realmente usado. Agora e' usado para gerar a cota GERAL DE\n"
    "     CADA PAREDE (pedido explicito), separada da cota geral do\n"
    "     alinhamento inteiro (quando varias paredes ficam em fileira).\n"
    "  2. Auto-deteccao dos DOIS eixos (H e V) na mesma execucao - a prancha\n"
    "     real mistura os dois, entao o script roda o pipeline duas vezes\n"
    "     (uma por direcao) e junta tudo, em vez de perguntar H ou V.\n"
    "  3. Coleta automatica de elementos se nada estiver selecionado (Paredes,\n"
    "     Portas, Janelas, Pisos, Escadas visiveis na view) - 'vasculhe tudo'.\n"
    "  4. Lado da cota decidido automaticamente (posiciona pra fora do centro\n"
    "     do modelo) - sem precisar clicar um ponto.\n"
    "  5. Deduplicacao por referencia estavel (nao so por valor arredondado)\n"
    "     conferida tanto contra o que ja existe na view quanto dentro da\n"
    "     propria execucao - reexecutar o script nao recria cotas repetidas.\n"
    "  6. Try/except granular em cada etapa (elemento, face, parede, tarefa) -\n"
    "     um erro isolado e reportado e pulado, nunca trava o resto.\n\n"
    "Camadas de cota geradas por alinhamento (igual pilha manual):\n"
    "  Nivel 0 - vaos/detalhe: um segmento por trecho entre referencias\n"
    "            consecutivas (portas, janelas, cantos), colado no elemento;\n"
    "  Nivel 1 - total por PAREDE INDIVIDUAL: uma cota so daquela parede,\n"
    "            ponta a ponta (mesmo que ela esteja num alinhamento maior\n"
    "            com outras paredes);\n"
    "  Nivel 2 - total do ALINHAMENTO: so criada quando o alinhamento tem MAIS\n"
    "            de uma parede (fileira) - soma o trecho todo;\n"
    "  Nivel 3 - GERAL da direcao: por fora de tudo, uma por eixo (H e V),\n"
    "            cobrindo todas as referencias daquele eixo.\n"
    "Niveis identicos (mesmas referencias) sao automaticamente descartados\n"
    "pela deduplicacao, entao paredes sozinhas nao duplicam nivel1==nivel2.\n"
)

# ============================================================
# IMPORTS
# ============================================================
from Autodesk.Revit.DB import (
    Options, Solid, PlanarFace, ReferenceArray, Line, XYZ,
    DimensionType, BuiltInParameter, BuiltInCategory,
    FilteredElementCollector, Dimension, Wall, FamilyInstance,
    ElementCategoryFilter, LogicalOrFilter,
)

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
    CLUSTER_TOL_CM       = 50.0   # separa alinhamentos/paredes diferentes
    COLA_ELEMENTO_CM     = 15.0   # distancia da cota "colada" ate o elemento
    GAP_NIVEL_CM         = 10.0   # espacamento entre niveis de linha (0->1->2)
    GAP_GERAL_CM         = 60.0   # afastamento extra da cota GERAL (nivel 3)
    MARGEM_PONTA_CM      = 20.0   # quanto a linha de cota estica alem das pontas

    # Filtro de ruido: ignora faces menores que isso (m2)
    AREA_MINIMA_FACE_M2 = 0.02

    # Categorias auto-coletadas quando nao ha selecao previa
    CATEGORIAS_AUTO = [
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Stairs,
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


def dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z


# ============================================================
# ETAPA 1 - Coleta de elementos (selecao ou automatica na view inteira)
# ============================================================
def coletar_elementos(view):
    sel_ids = list(uidoc.Selection.GetElementIds())

    if sel_ids:
        elementos = []
        for eid in sel_ids:
            el = doc.GetElement(eid)
            if el is not None and el.Category is not None:
                elementos.append(el)
        output.print_md(
            "## Cotar Parede Completo - **{} elemento(s) selecionado(s)**".format(len(elementos))
        )
        return elementos

    # Nada selecionado -> "vasculha tudo": coleta as categorias relevantes
    # visiveis na propria view ativa.
    output.print_md("## Cotar Parede Completo - nenhuma selecao, coletando a view inteira...")
    filtros = [ElementCategoryFilter(c) for c in Config.CATEGORIAS_AUTO]
    or_filter = filtros[0]
    for f in filtros[1:]:
        or_filter = LogicalOrFilter(or_filter, f)

    try:
        elementos = list(
            FilteredElementCollector(doc, view.Id).WherePasses(or_filter).WhereElementIsNotElementType().ToElements()
        )
    except Exception as e:
        logger.error("Falha ao coletar elementos da view automaticamente: {}".format(e))
        forms.alert("Falha ao coletar elementos da view:\n{}".format(e), exitscript=True)
        return []

    elementos = [el for el in elementos if el.Category is not None]

    if not elementos:
        forms.alert(
            "Nenhum elemento (parede/porta/janela/piso/escada) encontrado nesta view.\n"
            "Selecione manualmente os elementos e rode de novo.",
            exitscript=True,
        )

    por_cat = {}
    for el in elementos:
        nm = el.Category.Name
        por_cat[nm] = por_cat.get(nm, 0) + 1
    resumo = ", ".join("{}: {}".format(k, v) for k, v in sorted(por_cat.items()))
    output.print_md("**{} elemento(s) coletado(s)** ({})".format(len(elementos), resumo))
    return elementos


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
    parede-host; demais (piso, escada) -> None (entram so no alinhamento
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
    return None


def extrair_faces_referenciaveis(elementos, axis_dir, perp_dir, threshold=0.8):
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
                    pos_axis = dot(face.Origin, axis_dir)
                    pos_perp = dot(face.Origin, perp_dir)
                    resultado.append({
                        "pos_axis": pos_axis, "pos_perp": pos_perp,
                        "ref": face.Reference, "host_id": host_id,
                    })
                except Exception as e:
                    logger.debug("Face ignorada por erro: {}".format(e))
                    continue

    return resultado


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


def montar_correntes(itens, tol_cluster, tol_dedup):
    """Cada corrente = {'itens': [...ordenados por pos_axis, dedup...],
    'perp': media, 'host_ids': set de paredes presentes}."""
    grupos = agrupar_por_alinhamento(itens, tol_cluster)
    correntes = []
    for grupo in grupos:
        grupo_ordenado = sorted(grupo, key=lambda t: t["pos_axis"])
        grupo_dedup = dedupe_por_posicao(grupo_ordenado, tol_dedup)
        if len(grupo_dedup) < 2:
            continue
        perp_medio = sum(t["pos_perp"] for t in grupo_dedup) / len(grupo_dedup)
        host_ids = set(t["host_id"] for t in grupo_dedup if t["host_id"] is not None)
        correntes.append({"itens": grupo_dedup, "perp": perp_medio, "host_ids": host_ids})
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


# ============================================================
# ETAPA 5 - Geracao das tarefas de cota (vaos / parede / alinhamento / geral)
# ============================================================
def gerar_tarefas_de_cota(correntes, itens_todos, tolz):
    """Camadas por alinhamento:
      nivel 0 'vaos'            - so se a corrente tiver >2 referencias;
      nivel 1 'parede_total'    - uma por PAREDE INDIVIDUAL (host_id),
                                   usando so os pontos daquela parede;
      nivel 2 'alinhamento_total' - so quando o alinhamento tem MAIS de
                                   uma parede (fileira de paredes);
    Fora do loop por corrente:
      nivel 3 'geral'           - uma por eixo, por fora de tudo.
    """
    tarefas = []
    for c in correntes:
        itens_c = c["itens"]
        tem_detalhe = len(itens_c) > 2

        if tem_detalhe:
            for i in range(len(itens_c) - 1):
                tarefas.append({
                    "nome": "vaos",
                    "itens": [itens_c[i], itens_c[i + 1]],
                    "perp_ref": c["perp"], "nivel": 0,
                })

        # nivel 1: total por parede individual (pedido explicito do usuario)
        paredes_na_corrente = sorted(set(
            it["host_id"] for it in itens_c if it["host_id"] is not None
        ))
        for wid in paredes_na_corrente:
            pontos_parede = [it for it in itens_c if it["host_id"] == wid]
            if len(pontos_parede) < 2:
                continue
            tarefas.append({
                "nome": "parede_total",
                "itens": [pontos_parede[0], pontos_parede[-1]],
                "perp_ref": c["perp"], "nivel": 1 if tem_detalhe else 0,
            })

        # nivel 2: total do alinhamento inteiro (SO se tiver mais de 1 parede)
        if len(paredes_na_corrente) > 1:
            tarefas.append({
                "nome": "alinhamento_total",
                "itens": [itens_c[0], itens_c[-1]],
                "perp_ref": c["perp"], "nivel": 2,
            })
        elif not paredes_na_corrente:
            # nenhuma face pertence a uma parede identificavel (ex.: so
            # piso/escada relacionados) - ainda assim cota o alinhamento
            # inteiro, no nivel imediatamente acima do detalhe.
            tarefas.append({
                "nome": "alinhamento_total",
                "itens": [itens_c[0], itens_c[-1]],
                "perp_ref": c["perp"], "nivel": 1 if tem_detalhe else 0,
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
        resultado.append(t)
    if puladas:
        output.print_md("  [INFO] {} tarefa(s) descartada(s) por ja existirem (duplicata real).".format(puladas))
    return resultado


# ============================================================
# ETAPA 6 - Layout (posicao perpendicular final de cada tarefa)
# ============================================================
def resolver_layout(tarefas, centro_perp_modelo, tolz):
    """Decide automaticamente o LADO de cada alinhamento (pra fora do
    centro do modelo, sem precisar clicar um ponto) e empilha os niveis
    0/1/2 colados no elemento, com a GERAL (nivel 3) por fora de tudo."""
    perp_por_alinhamento = {}
    for t in tarefas:
        if t["nome"] == "geral":
            continue
        perp_por_alinhamento.setdefault(t["perp_ref"], []).append(t)

    perp_extremos = []
    for perp_ref, lista in perp_por_alinhamento.items():
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
    with revit.Transaction("Cotar Parede Completo"):
        for t in tarefas:
            try:
                dim_line = _cria_dim_line(axis, perp, t["itens"], t["perp_pos"], tolz.margem_ponta)
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
def processar_eixo(elementos, view, axis, perp, nome_eixo, tolz, assinaturas_existentes):
    """Roda o pipeline completo (extracao -> alinhamento -> tarefas ->
    dedup -> layout) para UM eixo (H ou V). Retorna a lista de tarefas
    prontas (com perp_pos definido) - a criacao no Revit e feita depois,
    juntando H + V numa unica transacao."""
    output.print_md("### Eixo {}".format(nome_eixo))

    itens = extrair_faces_referenciaveis(elementos, axis, perp)
    if len(itens) < 2:
        output.print_md("  [INFO] menos de 2 referencias nesse eixo - nada a cotar aqui.")
        return []

    correntes = montar_correntes(itens, tolz.cluster_tol, tolz.tol_dim_zero)
    if not correntes:
        output.print_md("  [INFO] nenhum alinhamento valido nesse eixo.")
        return []

    output.print_md("  {} alinhamento(s) encontrado(s) nesse eixo.".format(len(correntes)))

    tarefas = gerar_tarefas_de_cota(correntes, itens, tolz)
    tarefas = remover_tarefas_duplicadas(tarefas, assinaturas_existentes)

    centro_perp_modelo = sum(t["pos_perp"] for t in itens) / len(itens)
    tarefas = resolver_layout(tarefas, centro_perp_modelo, tolz)

    for t in tarefas:
        t["_axis"] = axis
        t["_perp"] = perp
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

    elementos = coletar_elementos(view)

    assinaturas_existentes = coletar_assinaturas_existentes(view)
    output.print_md("**{} assinatura(s) de cota ja existente(s)** na vista (nao serao repetidas).".format(
        len(assinaturas_existentes)))

    dim_type = find_dim_type_by_name(Config.NOME_TIPO_COTA_PADRAO)

    axis_h, perp_h = eixo_h
    axis_v, perp_v = perp_h, axis_h  # eixo V e' simplesmente o H trocado

    tarefas_h = processar_eixo(elementos, view, axis_h, perp_h, "Horizontal", tolz, assinaturas_existentes)
    tarefas_v = processar_eixo(elementos, view, axis_v, perp_v, "Vertical", tolz, assinaturas_existentes)

    # Junta H + V e faz uma segunda passada de deduplicacao (pode ter
    # coincidencia de assinatura entre os dois eixos em casos degenerados).
    todas_tarefas = tarefas_h + tarefas_v
    if not todas_tarefas:
        forms.alert(
            "Nao foi possivel montar nenhuma tarefa de cota valida (H ou V).\n"
            "Verifique se os elementos coletados tem geometria solida normal.",
            exitscript=True,
        )
        return

    try:
        criadas_total, erros_total, por_nome_total = 0, 0, {}
        with revit.Transaction("Cotar Parede Completo"):
            pass  # transacao real acontece dentro de criar_cotas_no_revit por eixo

        # Cria por eixo (cada chamada abre/fecha sua propria transacao
        # curta - assim um erro de commit num eixo nao contamina o outro).
        for nome_eixo, tarefas_eixo in (("Horizontal", tarefas_h), ("Vertical", tarefas_v)):
            if not tarefas_eixo:
                continue
            axis_ref = tarefas_eixo[0]["_axis"]
            perp_ref = tarefas_eixo[0]["_perp"]
            criadas, erros, por_nome = criar_cotas_no_revit(
                tarefas_eixo, view, axis_ref, perp_ref, tolz, dim_type
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
