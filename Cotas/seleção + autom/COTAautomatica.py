# -*- coding: utf-8 -*-
__title__ = "Cotar(interior)\nFase 2"
__doc__ = (
    "SCRIPT 2 (interior, automatico) - FASE 2.\n"
    "Em cima da Fase 1 (parede paralela mais proxima), escopado pra "
    "view ativa (nao mais o documento inteiro). Roda em QUALQUER "
    "parede - drywall foi so o exemplo usado pra validar o formato da "
    "cota (a 349 que o usuario cotou na mao), nao um filtro.\n"
    "Mudancas desta fase: (1) classifica e reporta cada cota criada "
    "como horizontal ou vertical; (2) mantem o dedup por par ja "
    "existente na Fase 1 (nunca 2 linhas medindo o mesmo par de "
    "paredes).\n"
    "Ainda sem: aberturas, cruzamento perpendicular, agrupamento "
    "explicito por comodo (o '1 H + 1 V por comodo' hoje e' uma "
    "consequencia do dedup por par, nao uma selecao por espaco - "
    "se sobrar/faltar cota em algum comodo especifico, ajustar aqui), "
    "dedup contra cota ja existente manualmente no modelo."
)

from Autodesk.Revit.DB import (
    FilteredElementCollector, Options, Solid, PlanarFace,
    XYZ, Line, ReferenceArray, DimensionType, DimensionStyleType,
    Wall, Transaction, ElementId
)
from pyrevit import revit, forms, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

TOL_DEDUP = 0.05  # pes (~1.5cm) - faces praticamente no mesmo lugar


# ============================================================
# 1) Extracao generica de faces referenciaveis alinhadas a um eixo
# ============================================================
def extrair_faces_da_parede(elementos, eixo, perp, threshold=0.985):
    """Pra cada elemento (parede...), acha as faces planas referenciaveis
    cuja normal esta alinhada com 'eixo' - ou seja, faces perpendiculares
    a direcao passada: chamada com (eixo_da_parede, perp) acha as faces
    de PONTA da parede; chamada com (perp, eixo) - trocado - acha as
    faces LONGAS (internas/externas), usadas pra medir afastamento entre
    paredes paralelas."""
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
# 2) Parede PARALELA mais proxima (largura de comodo/corredor)
#   - mesma ideia da cota temporaria que o Revit sugere quando voce
#     seleciona uma parede perto de outra paralela. Nao existe API pra
#     ler a cota temporaria em si (documentado pela Autodesk: "Temporary
#     dimensions created while editing an element in the UI are not
#     accessible") - entao a logica e' refeita na mao aqui.
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
        # so interessa parede realmente PARALELA (mesma direcao do eixo)
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
    """Monta a referencia pra uma cota de afastamento entre esta parede e
    a parede paralela mais proxima (largura de comodo/corredor).
    Retorna (face_esta, face_prox, meio_eixo, w_prox) ou None se nao
    achar par valido."""
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
# 3) Filtro minimo de validade da parede
# ============================================================
def parede_reta_valida(wall, comp_min=0.1):
    """So parede com Location.Curve reta (Line) e comprimento minimo -
    pula parede curva (arco), parede cortina sem Location.Curve, ou
    parede-toco de comprimento quase zero (geram linha de cota
    degenerada, e o Revit recusa atribuir o DimensionType).
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
# 4) Extremos de UMA parede (so o que a cota de parede paralela precisa)
# ============================================================
def extremos_da_parede(wall):
    """Retorna (p_ini, p_fim, eixo, perp) - as duas faces de PONTA da
    propria parede, e os vetores eixo/perp dela. [Fase 1 - Script 2]
    Sem mais coleta de detalhe/abertura - essa fase so precisa saber o
    alcance (lo/hi) e a posicao perpendicular da parede, entrada pra
    achar a parede paralela mais proxima. Retorna (None, None, eixo,
    perp) se a parede nao tiver 2 faces de ponta."""
    curve = wall.Location.Curve
    p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
    eixo = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0.0).Normalize()
    perp = XYZ(-eixo.Y, eixo.X, 0.0)

    pontas = extrair_faces_da_parede([wall], eixo, perp, threshold=0.985)
    pontas = dedupe_por_posicao(sorted(pontas, key=lambda t: t["pos_axis"]), TOL_DEDUP)
    if len(pontas) < 2:
        return None, None, eixo, perp
    return pontas[0], pontas[-1], eixo, perp


# ============================================================
# 5) Criar a cota
# ============================================================
def _criar_dimension(view, p0, p1, perp, offset, pontos, dim_type):
    linha = Line.CreateBound(
        XYZ(p0.X + perp.X * offset, p0.Y + perp.Y * offset, p0.Z),
        XYZ(p1.X + perp.X * offset, p1.Y + perp.Y * offset, p1.Z),
    )
    ra = ReferenceArray()
    for pt in pontos:
        ra.Append(pt["face"].Reference)
    nd = doc.Create.NewDimension(view, linha, ra)
    if dim_type:
        try:
            nd.DimensionType = dim_type
        except Exception:
            # achado testando ao vivo: o primeiro DimensionType "Linear" que
            # o coletor acha pode nao ser atribuivel (da "Dimension type is
            # not valid for this dimension"). Se isso acontecer, mantem o
            # tipo padrao que o Revit ja atribuiu na criacao (na pratica, o
            # ultimo tipo linear usado no documento) em vez de derrubar a
            # cota inteira.
            pass
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


# ============================================================
# 6) MODO TESTE - cria so 1 cota, entre um par especifico de paredes
#    (fase de validacao: o usuario cotou esse par na mao, ~349/348.5,
#    pra eu confirmar que a logica automatica bate com o que ele quer
#    antes de rodar em tudo. Nao mexe no fluxo automatico normal.)
# ============================================================
IDS_TESTE_PAR = (1661895, 1672375)  # usado so se MODO == "TESTE" (ver fim do arquivo)


def main_teste(id_wall_1, id_wall_2):
    wall1 = doc.GetElement(ElementId(id_wall_1))
    wall2 = doc.GetElement(ElementId(id_wall_2))
    if wall1 is None or wall2 is None:
        forms.alert("Nao achei uma das paredes de teste (Id {} / {}).".format(
            id_wall_1, id_wall_2), exitscript=True)
        return
    if not parede_reta_valida(wall1) or not parede_reta_valida(wall2):
        forms.alert("Uma das paredes de teste nao tem eixo reto valido.", exitscript=True)
        return

    dim_type = pegar_dimension_type_linear()
    view = doc.ActiveView

    t = Transaction(doc, "Cotar(interior) - MODO TESTE (par especifico)")
    t.Start()
    try:
        p_ini, p_fim, eixo, perp = extremos_da_parede(wall1)
        if p_ini is None:
            forms.alert("Parede {} sem as 2 faces de ponta - nao deu pra testar.".format(
                wall1.Id.IntegerValue), exitscript=True)
            return

        # forca a "parede paralela mais proxima" a ser exatamente wall2,
        # passando so ela como candidata (reaproveita a mesma logica do
        # modo automatico, sem duplicar codigo).
        resultado = montar_cota_parede_paralela(
            wall1, [wall2], eixo, perp, p_ini["pos_axis"], p_fim["pos_axis"], p_ini["pos_perp"]
        )
        if not resultado:
            forms.alert("Nao consegui montar a cota entre as 2 paredes de teste "
                        "(nao sobrepoem, ou faces nao encontradas).", exitscript=True)
            return

        face_esta, face_prox, meio_eixo, w_prox = resultado
        p0 = XYZ(eixo.X * meio_eixo + perp.X * face_esta["pos_axis"],
                 eixo.Y * meio_eixo + perp.Y * face_esta["pos_axis"], 0.0)
        p1 = XYZ(eixo.X * meio_eixo + perp.X * face_prox["pos_axis"],
                 eixo.Y * meio_eixo + perp.Y * face_prox["pos_axis"], 0.0)
        nd = _criar_dimension(view, p0, p1, eixo, 0.0, [face_esta, face_prox], dim_type)
        doc.Regenerate()
        output.print_md("## [TESTE] cota criada entre parede {} e {} - valor: **{}**".format(
            wall1.Id.IntegerValue, w_prox.Id.IntegerValue, nd.ValueString))
    finally:
        t.Commit()


def main():
    view = doc.ActiveView
    todas_paredes = list(FilteredElementCollector(doc, view.Id)
                          .OfClass(Wall).WhereElementIsNotElementType().ToElements())
    paredes = [w for w in todas_paredes if parede_reta_valida(w)]
    puladas = len(todas_paredes) - len(paredes)
    output.print_md("**{} parede(s)** na view **{}** - processando **{}**{}.".format(
        len(todas_paredes), view.Name, len(paredes),
        " - {} pulada(s) por nao ter eixo reto valido (curva/cortina/residual)".format(puladas) if puladas else ""
    ))
    if not paredes:
        forms.alert("Nenhuma parede encontrada na view ativa.", exitscript=True)
        return

    dim_type = pegar_dimension_type_linear()

    criadas_h = 0
    criadas_v = 0
    puladas_dup = 0
    pares_paralelas_ja_cotados = set()

    t = Transaction(doc, "Cotar(interior) - paredes paralelas, H+V, sem duplicar")
    t.Start()
    for wall in paredes:
        try:
            p_ini, p_fim, eixo, perp = extremos_da_parede(wall)
            if p_ini is None:
                output.print_md("  [AVISO] parede {}: nao achei as 2 faces de ponta - pulada.".format(
                    wall.Id.IntegerValue))
                continue

            resultado_paralela = montar_cota_parede_paralela(
                wall, todas_paredes, eixo, perp, p_ini["pos_axis"], p_fim["pos_axis"], p_ini["pos_perp"]
            )
            if not resultado_paralela:
                continue

            face_esta, face_prox, meio_eixo, w_prox = resultado_paralela

            # dedup: mesmo par (independente de qual das 2 disparou a
            # busca primeiro) nao cota 2x - nunca 2 linhas medindo a
            # mesma coisa.
            par_key = frozenset([wall.Id.IntegerValue, w_prox.Id.IntegerValue])
            if par_key in pares_paralelas_ja_cotados:
                puladas_dup += 1
                continue
            pares_paralelas_ja_cotados.add(par_key)

            try:
                p0 = XYZ(eixo.X * meio_eixo + perp.X * face_esta["pos_axis"],
                         eixo.Y * meio_eixo + perp.Y * face_esta["pos_axis"], 0.0)
                p1 = XYZ(eixo.X * meio_eixo + perp.X * face_prox["pos_axis"],
                         eixo.Y * meio_eixo + perp.Y * face_prox["pos_axis"], 0.0)
                _criar_dimension(view, p0, p1, eixo, 0.0, [face_esta, face_prox], dim_type)

                # H/V: parede correndo em Y (N-S) -> cota atravessa em X -> horizontal.
                # parede correndo em X (E-W) -> cota atravessa em Y -> vertical.
                if abs(eixo.Y) > abs(eixo.X):
                    criadas_h += 1
                else:
                    criadas_v += 1
            except Exception as e:
                output.print_md("  [ERRO] cota de parede paralela (parede {} <-> {}): {}".format(
                    wall.Id.IntegerValue, w_prox.Id.IntegerValue, e))

        except Exception as e:
            output.print_md("  [ERRO] parede {}: {}".format(wall.Id.IntegerValue, e))

    doc.Regenerate()
    t.Commit()

    output.print_md("## {} horizontal(is) + {} vertical(is) = {} cota(s) criada(s). "
                     "{} par(es) pulado(s) por ja ter cota (duplicado).".format(
                         criadas_h, criadas_v, criadas_h + criadas_v, puladas_dup))


MODO = "TODAS"  # "TESTE" (par especifico, validacao) | "TODAS" (Fase 2 - qualquer parede na view ativa)

if MODO == "TESTE":
    main_teste(*IDS_TESTE_PAR)
else:
    main()
