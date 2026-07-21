# -*- coding: utf-8 -*-
__title__ = "Cotar\nFase 4.3"
__doc__ = (
    "FASE 4.3 - mesma base da fase 4.2, adicionando UMA cota nova por "
    "parede: distancia ate a parede PARALELA mais proxima (largura de "
    "comodo/corredor). Isso aproxima o que a cota TEMPORARIA do Revit "
    "sugere quando voce seleciona uma parede perto de outra paralela - "
    "so que permanente e pra TODAS as paredes de uma vez. Nao usa "
    "nenhuma 'API de cota temporaria' porque ela nao existe/nao e' "
    "acessivel (confirmado na doc oficial da Autodesk: 'Temporary "
    "dimensions created while editing an element in the UI are not "
    "accessible') - a logica foi reconstruida na mao com a mesma ideia "
    "(parede parpendicular/paralela mais proxima).\n"
    "Ainda sem: mudanca de espessura, rebaixo, dedup contra cota ja "
    "existente no modelo."
)

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Options, Solid, PlanarFace,
    XYZ, Line, ReferenceArray, DimensionType, DimensionStyleType,
    Floor, Wall, FamilyInstance, Opening, Transaction
)
from pyrevit import revit, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

TOL_DEDUP      = 0.05  # pes (~1.5cm) - faces praticamente no mesmo lugar
OFFSET_TOTAL   = 2.0   # pes (~60cm) - linha de cima: parede inteira
OFFSET_DETALHE = 1.0   # pes (~30cm) - linha do meio: corrente de detalhe
OFFSET_ABERTURA = 0.4  # pes (~12cm) - linha de baixo: colada em cada abertura


# ============================================================
# 1) Achar o(s) piso(s)  [Fase 4.1: agora e' so apoio - opcional]
# ============================================================
def pegar_pisos():
    """[Fase 4.1] O piso deixou de ser filtro de "quais paredes cotar" -
    agora ele so serve de apoio pra saber pra que lado fica 'fora' do
    edificio (ver centro_geral/orientar_perp_para_fora). Por isso, se
    nao achar nenhum piso, o script NAO trava mais - so continua sem
    esse apoio (a cota geral pode acabar saindo do lado "errado" da
    parede em algum caso raro, mas todas as paredes ainda sao cotadas).
      - se tiver piso(s) selecionado(s), usa todos os selecionados;
      - senao, usa TODOS os pisos da vista ativa;
      - se nao achar nenhum, retorna lista vazia (sem alerta)."""
    sel_ids = list(uidoc.Selection.GetElementIds())
    pisos_selecionados = []
    for eid in sel_ids:
        el = doc.GetElement(eid)
        if isinstance(el, Floor):
            pisos_selecionados.append(el)
    if pisos_selecionados:
        return pisos_selecionados

    pisos_da_vista = list(
        FilteredElementCollector(doc, doc.ActiveView.Id)
        .OfCategory(BuiltInCategory.OST_Floors)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return pisos_da_vista


# ============================================================
# 2) Contorno externo do piso  (igual fase 3 - so mudou quem usa)
# ============================================================
def contorno_do_piso(floor):
    opt = Options()
    geom = floor.get_Geometry(opt)
    if geom is None:
        return []

    face_topo, maior_area = None, 0.0
    for g in geom:
        if not isinstance(g, Solid) or g.Volume <= 0:
            continue
        for face in g.Faces:
            if not isinstance(face, PlanarFace):
                continue
            if face.FaceNormal.Z < 0.9:
                continue
            if face.Area > maior_area:
                maior_area = face.Area
                face_topo = face

    if face_topo is None:
        return []

    loops = face_topo.GetEdgesAsCurveLoops()
    if not loops:
        return []

    maior_loop = max(loops, key=lambda lp: sum(c.Length for c in lp))
    return [(c.GetEndPoint(0), c.GetEndPoint(1)) for c in maior_loop]


def contorno_dos_pisos(pisos):
    """Junta o contorno de VARIOS pisos numa lista unica de segmentos.
    [Fase 4.1] Nao serve mais pra filtrar parede - serve so de entrada
    pra centro_geral(), que usa esses segmentos pra estimar o meio do
    edificio."""
    segmentos = []
    ok, falha = 0, 0
    for floor in pisos:
        segs = contorno_do_piso(floor)
        if segs:
            segmentos.extend(segs)
            ok += 1
        else:
            falha += 1
            output.print_md("  [AVISO] piso {}: nao consegui ler o contorno - ignorado.".format(
                floor.Id.IntegerValue))
    if pisos:
        output.print_md("Contorno lido de **{} piso(s)** ({} falharam) - **{} segmento(s)** no total (apoio de direcao).".format(
            ok, falha, len(segmentos)))
    return segmentos


# ============================================================
# [NOVO Fase 4.1] 3) Centro geral dos pisos - so pra saber onde e' "dentro"
# ============================================================
def centro_geral(segmentos):
    """Media de todos os pontos dos contornos dos pisos - um centroide
    aproximado do edificio. Nao precisa ser exato: so serve de
    referencia pra decidir de que lado do eixo da parede fica "fora"
    (ver orientar_perp_para_fora). Se nao tiver segmento nenhum
    (nenhum piso lido), retorna None."""
    if not segmentos:
        return None
    xs, ys = [], []
    for a, b in segmentos:
        xs.append(a.X); xs.append(b.X)
        ys.append(a.Y); ys.append(b.Y)
    return XYZ(sum(xs) / len(xs), sum(ys) / len(ys), 0.0)


def orientar_perp_para_fora(perp, ponto_da_parede, centro):
    """[Fase 4.1] Vira o vetor perpendicular 180 graus se ele estiver
    apontando pra DENTRO do edificio (na direcao do centro geral dos
    pisos) - assim a cota geral/detalhe/abertura sempre nasce do lado
    de fora da parede, mesmo sem usar mais o piso como filtro de
    "quais paredes processar". Se nao tiver centro (nenhum piso lido),
    devolve o perp sem mexer - so nao da pra garantir o lado."""
    if centro is None:
        return perp
    para_centro = XYZ(centro.X - ponto_da_parede.X, centro.Y - ponto_da_parede.Y, 0.0)
    if para_centro.GetLength() < 1e-6:
        return perp
    if perp.X * para_centro.X + perp.Y * para_centro.Y > 0:
        return XYZ(-perp.X, -perp.Y, 0.0)
    return perp


# ============================================================
# 4) Extracao generica de faces referenciaveis alinhadas a um eixo
# ============================================================
def extrair_faces_da_parede(elementos, eixo, perp, threshold=0.985):
    """Pra cada elemento (parede, porta, janela...), acha as faces planas
    referenciaveis cuja normal esta alinhada com 'eixo' - ou seja, faces
    perpendiculares a direcao da parede: pontas da parede, batentes de
    porta/janela, faces de parede perpendicular no encontro.

    [Fase 4.3] Generica o bastante pra ser reaproveitada TROCANDO os
    parametros: chamar com (perp, eixo) em vez de (eixo, perp) acha as
    faces LONGAS da parede (as superficies internas/externas), usadas
    pra medir afastamento entre paredes paralelas - ver
    montar_cota_parede_paralela."""
    opt = Options()
    opt.ComputeReferences = True
    resultado = []
    for el in elementos:
        try:
            geom = el.get_Geometry(opt)
        except Exception:
            continue
        if geom is None:
            continue
        for g in geom:
            if not isinstance(g, Solid) or g.Volume <= 0:
                continue
            try:
                faces = g.Faces
            except Exception:
                continue
            for face in faces:
                if not isinstance(face, PlanarFace) or face.Reference is None:
                    continue
                n = face.FaceNormal
                d = n.X * eixo.X + n.Y * eixo.Y
                if abs(d) < threshold:
                    continue
                pos_axis = face.Origin.X * eixo.X + face.Origin.Y * eixo.Y
                pos_perp = face.Origin.X * perp.X + face.Origin.Y * perp.Y
                resultado.append({"pos_axis": pos_axis, "pos_perp": pos_perp, "face": face})
    return resultado


def dedupe_por_posicao(itens, tol):
    """itens ordenados por pos_axis - remove faces praticamente no mesmo
    lugar (mesma posicao ao longo do eixo)."""
    if not itens:
        return []
    aceitos = [itens[0]]
    for it in itens[1:]:
        if abs(it["pos_axis"] - aceitos[-1]["pos_axis"]) > tol:
            aceitos.append(it)
    return aceitos


# ============================================================
# 5) Detalhes de uma parede: portas/janelas hospedadas
# ============================================================
def elementos_hospedados_na_parede(wall):
    """Portas e janelas cujo Host e' exatamente esta parede."""
    cats = (BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows)
    encontrados = []
    for cat in cats:
        els = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
        for el in els:
            try:
                host = el.Host
            except Exception:
                host = None
            if host is not None and host.Id.IntegerValue == wall.Id.IntegerValue:
                encontrados.append(el)
    return encontrados


def aberturas_genericas_na_parede(wall):
    """Aberturas RETANGULARES SEM porta/janela (a ferramenta
    'Opening -> Wall' do Revit, classe Opening) hospedadas nesta parede.
    Diferente de porta/janela, um Opening nao tem solido/geometria
    proprio (ele so recorta a parede) - por isso a face dele e' achada
    depois, direto na propria parede (ver faces_da_abertura_generica)."""
    try:
        openings = FilteredElementCollector(doc).OfClass(Opening).WhereElementIsNotElementType().ToElements()
    except Exception:
        return []
    encontradas = []
    for op in openings:
        try:
            host = op.Host
        except Exception:
            host = None
        if host is not None and host.Id.IntegerValue == wall.Id.IntegerValue:
            encontradas.append(op)
    return encontradas


# ============================================================
# 6) Encontros com paredes perpendiculares (cruzamentos)
# ============================================================
def paredes_perpendiculares_cruzando(wall, todas_paredes, perp_dir, tol_extensao=0.5):
    """Acha, entre 'todas_paredes', as que cruzam a extensao desta
    parede e sao perpendiculares a ela (alinhadas com perp_dir). Retorna
    lista de (parede_cruzando, ponto_do_encontro)."""
    curve = wall.Location.Curve
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    d1 = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0)
    len1 = d1.GetLength()
    if len1 < 1e-6:
        return []
    d1n = d1.Normalize()

    resultado = []
    for w in todas_paredes:
        if w.Id.IntegerValue == wall.Id.IntegerValue:
            continue
        c2 = getattr(w.Location, "Curve", None)
        if c2 is None or not isinstance(c2, Line):
            continue
        q0, q1 = c2.GetEndPoint(0), c2.GetEndPoint(1)
        d2 = XYZ(q1.X - q0.X, q1.Y - q0.Y, 0.0)
        len2 = d2.GetLength()
        if len2 < 1e-6:
            continue
        d2n = d2.Normalize()

        # so interessa parede realmente perpendicular a esta (alinhada
        # com a direcao perp, nao com o proprio eixo)
        if abs(d2n.X * perp_dir.X + d2n.Y * perp_dir.Y) < 0.8:
            continue

        denom = d1n.X * d2n.Y - d1n.Y * d2n.X
        if abs(denom) < 1e-9:
            continue
        dx, dy = q0.X - p0.X, q0.Y - p0.Y
        t = (dx * d2n.Y - dy * d2n.X) / denom
        s = (dx * d1n.Y - dy * d1n.X) / denom
        if t < -tol_extensao or t > len1 + tol_extensao:
            continue
        if s < -tol_extensao or s > len2 + tol_extensao:
            continue

        ponto = XYZ(p0.X + d1n.X * t, p0.Y + d1n.Y * t, p0.Z)
        resultado.append((w, ponto))
    return resultado


# ============================================================
# [NOVO Fase 4.3] 6.4) Parede PARALELA mais proxima (largura de comodo)
#   - mesma ideia da cota temporaria que o Revit sugere quando voce
#     seleciona uma parede perto de outra paralela (largura de
#     comodo/corredor). Nao existe API pra ler a cota temporaria em si
#     (documentado pela Autodesk: "Temporary dimensions created while
#     editing an element in the UI are not accessible") - entao a logica
#     e' refeita na mao aqui, com a mesma ideia geometrica.
# ============================================================
def _overlap(lo1, hi1, lo2, hi2):
    return max(0.0, min(hi1, hi2) - max(lo1, lo2))


def parede_paralela_mais_proxima(wall, todas_paredes, eixo, perp, lo, hi):
    """Entre 'todas_paredes', acha as paredes PARALELAS (mesma direcao de
    eixo) cuja faixa ao longo do eixo se sobrepoe com a desta parede
    (elas "correm lado a lado", tipo as duas paredes de um corredor ou
    comodo). Retorna lista de (parede, pos_perp_dela, lo2, hi2)."""
    candidatos = []
    for w in todas_paredes:
        if w.Id.IntegerValue == wall.Id.IntegerValue:
            continue
        curve2 = getattr(w.Location, "Curve", None)
        if curve2 is None or not isinstance(curve2, Line):
            continue
        q0, q1 = curve2.GetEndPoint(0), curve2.GetEndPoint(1)
        d2 = XYZ(q1.X - q0.X, q1.Y - q0.Y, 0.0)
        if d2.GetLength() < 1e-6:
            continue
        d2n = d2.Normalize()
        # so interessa parede realmente PARALELA (mesma direcao do eixo -
        # perpendiculares ja sao tratadas por paredes_perpendiculares_cruzando)
        if abs(d2n.X * eixo.X + d2n.Y * eixo.Y) < 0.99:
            continue

        lo2 = min(q0.X * eixo.X + q0.Y * eixo.Y, q1.X * eixo.X + q1.Y * eixo.Y)
        hi2 = max(q0.X * eixo.X + q0.Y * eixo.Y, q1.X * eixo.X + q1.Y * eixo.Y)
        sobreposicao_min = 0.3 * min(hi - lo, hi2 - lo2) if min(hi - lo, hi2 - lo2) > 0 else 0.0
        if _overlap(lo, hi, lo2, hi2) < max(sobreposicao_min, TOL_DEDUP):
            continue

        pos_perp2 = q0.X * perp.X + q0.Y * perp.Y
        candidatos.append((w, pos_perp2, lo2, hi2))

    return candidatos


def montar_cota_parede_paralela(wall, todas_paredes, eixo, perp, lo, hi, pos_perp_parede):
    """[Fase 4.3] Monta a referencia pra uma cota extra de afastamento
    entre esta parede e a parede paralela mais proxima (largura de
    comodo/corredor) - aproximando o que a cota temporaria do Revit
    sugeriria. Retorna (face_esta, face_prox, meio_eixo, w_prox) ou None
    se nao achar par valido."""
    candidatos = parede_paralela_mais_proxima(wall, todas_paredes, eixo, perp, lo, hi)
    if not candidatos:
        return None

    w_prox, pos_perp2, lo2, hi2 = min(candidatos, key=lambda c: abs(c[1] - pos_perp_parede))
    if abs(pos_perp2 - pos_perp_parede) < TOL_DEDUP:
        return None

    meio_eixo = (max(lo, lo2) + min(hi, hi2)) / 2.0

    # extrai as faces LONGAS (normal ~ perp) trocando a ordem dos
    # parametros de extrair_faces_da_parede - ver docstring dela.
    faces_proxima = extrair_faces_da_parede([w_prox], perp, eixo, threshold=0.90)
    if not faces_proxima:
        return None
    face_prox = min(faces_proxima, key=lambda f: abs(f["pos_axis"] - pos_perp_parede))

    faces_esta = extrair_faces_da_parede([wall], perp, eixo, threshold=0.90)
    if not faces_esta:
        return None
    face_esta = min(faces_esta, key=lambda f: abs(f["pos_axis"] - pos_perp2))

    if abs(face_esta["pos_axis"] - face_prox["pos_axis"]) < TOL_DEDUP:
        return None

    return face_esta, face_prox, meio_eixo, w_prox


# ============================================================
# [NOVO Fase 4.1] 6.5) Filtro minimo de validade da parede
# ============================================================
def parede_reta_valida(wall, comp_min=0.1):
    """[Fase 4.1] Antes, so parede com Location.Curve reta (Line) tinha
    chance de entrar na lista, porque o filtro de contorno do piso
    testava isso por baixo dos panos (_parede_em_cima_do_segmento so
    aceitava Line). Agora que processamos TODAS as paredes do modelo,
    isso precisa ser checado na mao aqui - senao parede curva (arco),
    parede cortina sem Location.Curve, ou parede-toco de comprimento
    quase zero geram uma linha de cota degenerada, e o Revit recusa na
    hora de atribuir o DimensionType (erro 'Dimension type is not
    valid for this dimension').
    comp_min = 0.1 pe (~3cm) - abaixo disso considera residual."""
    curve = getattr(wall.Location, "Curve", None)
    if curve is None or not isinstance(curve, Line):
        return False
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    comp = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0).GetLength()
    if comp < comp_min:
        return False
    return True


# ============================================================
# 7) Monta os pontos de cota de UMA parede: pontas + detalhes
# ============================================================
def montar_pontos_de_cota(wall, todas_paredes):
    """Retorna (p_ini, p_fim, detalhes, pontos_parede) onde:
      p_ini/p_fim    = faces das duas pontas da parede (cota geral);
      detalhes       = faces de porta/janela/encontro ENTRE as pontas,
                       ordenadas ao longo do eixo (corrente de subcotas);
      pontos_parede  = TODAS as faces da propria parede alinhadas ao
                       eixo (inclui as pontas E as faces internas que o
                       recorte de qualquer abertura cria na parede) -
                       usada pra achar abertura sem familia.
    Retorna (None, None, None, None) se a parede nao tiver 2 faces de
    ponta."""
    curve = wall.Location.Curve
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    eixo = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0).Normalize()
    perp = XYZ(-eixo.Y, eixo.X, 0.0)

    # pontas da propria parede - so as faces bem alinhadas (threshold
    # estrito), pra nao pegar quina/chanfro por engano.
    pontos_parede = extrair_faces_da_parede([wall], eixo, perp, threshold=0.985)
    pontos_parede = dedupe_por_posicao(sorted(pontos_parede, key=lambda t: t["pos_axis"]), TOL_DEDUP)
    if len(pontos_parede) < 2:
        return None, None, None, None

    p_ini, p_fim = pontos_parede[0], pontos_parede[-1]
    lo, hi = p_ini["pos_axis"], p_fim["pos_axis"]

    # detalhes: portas/janelas hospedadas nesta parede. [Fase 4.2] usa a
    # extracao combinada (geometria propria -> fallback por bbox contra
    # as faces da propria parede), porque familias de porta so-linha
    # (sem solido) sempre voltariam vazias na extracao direta.
    relacionados = elementos_hospedados_na_parede(wall)
    detalhes = []
    for el in relacionados:
        faces = faces_de_elemento_hospedado(el, eixo, perp, pontos_parede, threshold=0.90)
        if faces:
            detalhes.append(faces[0])
            detalhes.append(faces[1])

    # encontros com parede perpendicular: pega a face da parede que
    # cruza, mais perto do ponto real do encontro
    for w_cruz, ponto in paredes_perpendiculares_cruzando(wall, todas_paredes, perp):
        pos_cruz = ponto.X * eixo.X + ponto.Y * eixo.Y
        faces_cruz = extrair_faces_da_parede([w_cruz], eixo, perp, threshold=0.90)
        if not faces_cruz:
            continue
        melhor = min(faces_cruz, key=lambda f: abs(f["pos_axis"] - pos_cruz))
        try:
            limite = w_cruz.Width * 1.5
        except Exception:
            limite = 0.5
        if abs(melhor["pos_axis"] - pos_cruz) <= limite:
            detalhes.append(melhor)

    # so os detalhes que caem DENTRO do trecho da parede (entre as pontas)
    detalhes = [d for d in detalhes if lo - TOL_DEDUP <= d["pos_axis"] <= hi + TOL_DEDUP]
    detalhes = dedupe_por_posicao(sorted(detalhes, key=lambda d: d["pos_axis"]), TOL_DEDUP)

    return p_ini, p_fim, detalhes, pontos_parede


# ============================================================
# 7.5) Faces de CADA abertura, pra cota colada nela
# ============================================================
def faces_extremas_do_elemento(elemento, eixo, perp, threshold=0.90):
    """Duas faces (esquerda/direita) do proprio solido do elemento -
    serve pra porta/janela (FamilyInstance), que tem geometria propria."""
    pts = extrair_faces_da_parede([elemento], eixo, perp, threshold=threshold)
    pts = dedupe_por_posicao(sorted(pts, key=lambda t: t["pos_axis"]), TOL_DEDUP)
    if len(pts) < 2:
        return None
    return pts[0], pts[-1]


def faces_da_abertura_generica(opening, eixo, pontos_parede):
    """Um 'Opening' (vao sem esquadria) nao tem solido proprio - so
    recorta a parede. Entao, em vez de olhar a geometria do Opening,
    usa a BOUNDING BOX dele pra saber onde ele comeca/termina no eixo, e
    pega as faces mais proximas dentro de 'pontos_parede' (a lista
    COMPLETA de faces da propria parede, que ja inclui as faces internas
    criadas pelo recorte do proprio Opening)."""
    try:
        bb = opening.get_BoundingBox(None)
    except Exception:
        bb = None
    if bb is None or len(pontos_parede) < 2:
        return None

    cantos = [
        XYZ(bb.Min.X, bb.Min.Y, 0.0), XYZ(bb.Max.X, bb.Min.Y, 0.0),
        XYZ(bb.Min.X, bb.Max.Y, 0.0), XYZ(bb.Max.X, bb.Max.Y, 0.0),
    ]
    posicoes = [c.X * eixo.X + c.Y * eixo.Y for c in cantos]
    lo, hi = min(posicoes), max(posicoes)

    esq = min(pontos_parede, key=lambda p: abs(p["pos_axis"] - lo))
    dire = min(pontos_parede, key=lambda p: abs(p["pos_axis"] - hi))
    if esq is dire or abs(esq["pos_axis"] - dire["pos_axis"]) < TOL_DEDUP:
        return None
    if esq["pos_axis"] > dire["pos_axis"]:
        esq, dire = dire, esq
    return esq, dire


def faces_de_elemento_hospedado(elemento, eixo, perp, pontos_parede, threshold=0.90):
    """[Fase 4.2] Acha o par de faces (esquerda/direita) de um elemento
    hospedado na parede (porta, janela...), usado tanto pra virar ponto
    de detalhe/subcota quanto pra cota colada nele.

    Primeiro tenta a geometria PROPRIA do elemento (funciona pra janelas
    e familias que realmente tem solido 3D). Se vier vazio - caso comum
    de familias de porta modeladas so com linhas simbolicas, sem nenhum
    solido (ex.: "Abertura de porta") - cai no mesmo truque ja usado pro
    Opening generico: casa a BOUNDING BOX do elemento com as faces que a
    propria parede ja expoe. O recorte da abertura sempre existe no
    solido da parede, mesmo quando a familia da porta nao tem solido
    proprio - entao isso resolve tanto a cota da abertura quanto
    alimenta a corrente de detalhe corretamente."""
    faces_proprias = faces_extremas_do_elemento(elemento, eixo, perp, threshold=threshold)
    if faces_proprias:
        return faces_proprias
    return faces_da_abertura_generica(elemento, eixo, pontos_parede)


def montar_aberturas(wall, pontos_parede, eixo, perp):
    """Pra cada abertura desta parede (porta, janela ou vao sem
    esquadria), acha o par de faces (jambas) dela - serve pra desenhar
    uma cota bem colada na propria abertura, so com a largura dela,
    separada da corrente geral de detalhe."""
    aberturas = []

    for el in elementos_hospedados_na_parede(wall):
        faces = faces_de_elemento_hospedado(el, eixo, perp, pontos_parede, threshold=0.90)
        if faces:
            aberturas.append((el, faces[0], faces[1]))

    for op in aberturas_genericas_na_parede(wall):
        faces = faces_da_abertura_generica(op, eixo, pontos_parede)
        if faces:
            aberturas.append((op, faces[0], faces[1]))

    return aberturas


# ============================================================
# 8) Criar as cotas: linha de baixo (abertura) + meio (detalhe) + cima (total)
# ============================================================
def _linha_de_cota(p0, p1, perp, offset):
    return Line.CreateBound(
        XYZ(p0.X + perp.X * offset, p0.Y + perp.Y * offset, p0.Z),
        XYZ(p1.X + perp.X * offset, p1.Y + perp.Y * offset, p1.Z),
    )


def _criar_dimension(view, p0, p1, perp, offset, pontos, dim_type):
    linha = _linha_de_cota(p0, p1, perp, offset)
    ra = ReferenceArray()
    for pt in pontos:
        ra.Append(pt["face"].Reference)
    nd = doc.Create.NewDimension(view, linha, ra)
    if dim_type:
        nd.DimensionType = dim_type
    return nd


def pegar_dimension_type_linear():
    """Pega um DimensionType do tipo LINEAR - pegar 'o primeiro que
    aparecer' sem filtrar da erro, porque o modelo tem varios estilos de
    cota (angular, raio, diametro, nivel...) e so o Linear serve pra
    Dimension criada entre faces (ponto a ponto)."""
    tipos = FilteredElementCollector(doc).OfClass(DimensionType).ToElements()
    for t in tipos:
        try:
            if t.StyleType == DimensionStyleType.Linear:
                return t
        except Exception:
            continue
    return None


def main():
    # [Fase 4.1] piso agora e' so apoio de direcao - nao filtra mais
    # "quais paredes processar" e nao trava o script se nao existir.
    pisos = pegar_pisos()
    segmentos = contorno_dos_pisos(pisos) if pisos else []
    centro = centro_geral(segmentos)
    if centro is None:
        output.print_md(
            "[AVISO] nenhum piso lido - sem referencia de 'lado de fora'. "
            "As paredes ainda vao ser todas cotadas, so nao da pra garantir "
            "que a cota sai do lado de fora em 100% dos casos."
        )

    # [Fase 4.1] fonte das paredes a cotar = TODAS as paredes do modelo
    # (antes era so as "em cima do contorno do piso" - quebrava com
    # piso fragmentado em varios comodos).
    todas_paredes = list(FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType().ToElements())
    paredes = [w for w in todas_paredes if parede_reta_valida(w)]
    puladas = len(todas_paredes) - len(paredes)
    output.print_md("**{} parede(s)** no modelo - processando **{}** (piso nao filtra mais, so ajuda a jogar a cota pra fora){}.".format(
        len(todas_paredes), len(paredes),
        " - {} pulada(s) por nao ter eixo reto valido (curva/cortina/residual)".format(puladas) if puladas else ""
    ))
    if not paredes:
        forms.alert("Nenhuma parede encontrada no modelo.", exitscript=True)
        return

    dim_type = pegar_dimension_type_linear()
    view = doc.ActiveView

    criadas_total, criadas_detalhe, criadas_abertura, criadas_vizinha = 0, 0, 0, 0
    pares_paralelas_ja_cotados = set()

    t = Transaction(doc, "Cotar Fase 4.3 - parede + subcotas + aberturas + parede paralela")
    t.Start()
    for wall in paredes:
        try:
            p_ini, p_fim, detalhes, pontos_parede = montar_pontos_de_cota(wall, todas_paredes)
            if p_ini is None:
                output.print_md("  [AVISO] parede {}: nao achei as 2 faces de ponta - pulada.".format(
                    wall.Id.IntegerValue))
                continue

            curve = wall.Location.Curve
            c0, c1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            eixo = XYZ(c1.X - c0.X, c1.Y - c0.Y, 0.0).Normalize()
            perp = XYZ(-eixo.Y, eixo.X, 0.0)

            # [Fase 4.1] vira o perp pra fora do edificio (se tiver
            # centro geral) antes de desenhar qualquer uma das 3 linhas
            meio = XYZ((c0.X + c1.X) / 2.0, (c0.Y + c1.Y) / 2.0, 0.0)
            perp_fora = orientar_perp_para_fora(perp, meio, centro)

            # linha de cima: parede inteira (sempre)
            _criar_dimension(view, c0, c1, perp_fora, OFFSET_TOTAL, [p_ini, p_fim], dim_type)
            criadas_total += 1

            # linha do meio: pedacinhos, so se tiver detalhe de verdade
            # (mais que so as 2 pontas repetidas)
            cadeia = dedupe_por_posicao(
                sorted([p_ini] + detalhes + [p_fim], key=lambda t: t["pos_axis"]), TOL_DEDUP
            )
            if len(cadeia) > 2:
                _criar_dimension(view, c0, c1, perp_fora, OFFSET_DETALHE, cadeia, dim_type)
                criadas_detalhe += 1

            # linha de baixo: uma cota SO da largura, bem colada, pra
            # CADA abertura (porta, janela ou vao sem esquadria)
            for elemento, f_a, f_b in montar_aberturas(wall, pontos_parede, eixo, perp):
                try:
                    _criar_dimension(view, c0, c1, perp_fora, OFFSET_ABERTURA, [f_a, f_b], dim_type)
                    criadas_abertura += 1
                except Exception as e:
                    output.print_md("  [ERRO] abertura {} (parede {}): {}".format(
                        elemento.Id.IntegerValue, wall.Id.IntegerValue, e))

            # [Fase 4.3] cota extra: distancia ate a parede PARALELA mais
            # proxima (largura de comodo/corredor) - aproxima o que a
            # cota temporaria do Revit sugere. Dedup por par de paredes
            # pra nao cotar a mesma dupla duas vezes (uma vez processando
            # cada lado).
            resultado_paralela = montar_cota_parede_paralela(
                wall, todas_paredes, eixo, perp, p_ini["pos_axis"], p_fim["pos_axis"], p_ini["pos_perp"]
            )
            if resultado_paralela:
                face_esta, face_prox, meio_eixo, w_prox = resultado_paralela
                par_key = frozenset([wall.Id.IntegerValue, w_prox.Id.IntegerValue])
                if par_key not in pares_paralelas_ja_cotados:
                    pares_paralelas_ja_cotados.add(par_key)
                    try:
                        p0 = XYZ(eixo.X * meio_eixo + perp.X * face_esta["pos_axis"],
                                 eixo.Y * meio_eixo + perp.Y * face_esta["pos_axis"], 0.0)
                        p1 = XYZ(eixo.X * meio_eixo + perp.X * face_prox["pos_axis"],
                                 eixo.Y * meio_eixo + perp.Y * face_prox["pos_axis"], 0.0)
                        _criar_dimension(view, p0, p1, eixo, 0.0, [face_esta, face_prox], dim_type)
                        criadas_vizinha += 1
                    except Exception as e:
                        output.print_md("  [ERRO] cota de parede paralela (parede {} <-> {}): {}".format(
                            wall.Id.IntegerValue, w_prox.Id.IntegerValue, e))

        except Exception as e:
            output.print_md("  [ERRO] parede {}: {}".format(wall.Id.IntegerValue, e))

    # [Fase 4.2.1] Forca a regeneracao do documento ANTES do commit - sem
    # isso, cotas criadas na mesma transacao podem ficar num estado
    # "pendente" (o Revit ainda nao recalculou geometria/graficos pra
    # elas) e aparecer como se fossem temporarias/invisiveis ate a
    # proxima regeneracao manual (trocar de vista, mover algo, etc).
    # Regenerar aqui garante que todas as cotas ja saem 100% consolidadas
    # e permanentes assim que a transacao fecha.
    doc.Regenerate()
    t.Commit()

    output.print_md(
        "## {} cota(s) de parede inteira + {} cota(s) de detalhe/subcota + "
        "{} cota(s) de abertura + {} cota(s) de parede paralela criada(s).".format(
            criadas_total, criadas_detalhe, criadas_abertura, criadas_vizinha)
    )


main()
