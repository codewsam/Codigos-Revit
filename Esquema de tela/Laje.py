import sys
import clr
import math
import re

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

doc = DocumentManager.Instance.CurrentDBDocument

# ===============================
# ENTRADAS
# ===============================
elements         = UnwrapElement(IN[0])
nome_largura     = IN[1]
nome_comprimento = IN[2]
nome_folha       = IN[3]

# ===============================
# CONFIGURAÇÃO
# ===============================
CATEGORIA_DESEJADA = BuiltInCategory.OST_Floors
NOME_VISTA         = "Esquema de corte - LAJES"
TOLERANCIA         = 0.01
ARREDONDAMENTO     = 3

# ===============================
# FUNÇÕES
# ===============================

def numero_da_folha(valor):
    try:
        texto = str(valor).upper().strip().replace("N", "")
        numeros = re.findall(r'\d+', texto)
        if numeros:
            return int(numeros[0])
    except:
        pass
    return 999999


def formatar_folha(valor):
    texto = str(valor).upper().strip()
    return texto if texto.startswith("N") else "N" + texto


def get_param_value(el, nomes):
    if el is None:
        return "-"
    for nome in nomes:
        try:
            p = el.LookupParameter(nome)
            if p and p.HasValue:
                if p.StorageType == StorageType.String:
                    v = p.AsString()
                    return v if v else "-"
                elif p.StorageType == StorageType.Integer:
                    return str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    return str(round(p.AsDouble(), 3))
                elif p.StorageType == StorageType.ElementId:
                    return str(p.AsElementId().IntegerValue)
        except:
            pass
    return "-"


def get_host(el):
    try:
        host_id = el.HostId
        if host_id and host_id != ElementId.InvalidElementId:
            return doc.GetElement(host_id)
    except:
        pass
    try:
        if el.Host:
            return el.Host
    except:
        pass
    try:
        p = el.get_Parameter(BuiltInParameter.HOST_ID_PARAM)
        if p:
            host_id = p.AsElementId()
            if host_id and host_id != ElementId.InvalidElementId:
                return doc.GetElement(host_id)
    except:
        pass
    return None


def host_eh_categoria(host, categoria):
    if host is None:
        return False
    if host.Category is None:
        return False
    return host.Category.Id.IntegerValue == int(categoria)


def get_marca_tipo_tela(el):
    try:
        tipo = doc.GetElement(el.GetTypeId())
        p = tipo.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
        if p and p.HasValue:
            v = p.AsString()
            return v if v else "-"
    except:
        pass
    return "-"


def get_marca_hospedeiro(host):
    if host is None:
        return "-"
    for busca in [lambda h: h.LookupParameter("Marca"),
                  lambda h: h.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)]:
        try:
            p = busca(host)
            if p and p.HasValue:
                v = p.AsString()
                return v if v else "-"
        except:
            pass
    return "-"


def get_particao(el):
    try:
        p = el.LookupParameter(u"Parti\xe7\xe3o")
        if p and p.HasValue:
            v = p.AsString()
            return v.strip() if v and v.strip() else u"Sem Parti\xe7\xe3o"
    except:
        pass
    return u"Sem Parti\xe7\xe3o"


def ordem_particao(nome):
    if nome == u"Sem Parti\xe7\xe3o":
        return (1, "")
    return (0, nome.upper())


def ordem_marca(marca):
    nums = re.findall(r'\d+', str(marca))
    if nums:
        return (0, int(nums[0]), marca)
    return (1, 0, marca.upper())


# ===============================
# AGRUPAMENTO
# Chave: (particao, marca_tipo_tela, l_key, c_key)
# → telas de mesma partição + mesmo tipo + mesma dimensão
#   são somadas INDEPENDENTE da laje hospedeira
# ===============================
grupos       = {}
ignorados    = []
sem_dimensao = []

for el in elements:
    host = get_host(el)

    if not host_eh_categoria(host, CATEGORIA_DESEJADA):
        ignorados.append(el.Id)
        continue

    p_larg = el.LookupParameter(nome_largura)
    p_comp = el.LookupParameter(nome_comprimento)

    if not p_larg or not p_comp:
        sem_dimensao.append(el.Id)
        continue

    l_orig = p_larg.AsDouble()
    c_orig = p_comp.AsDouble()

    if l_orig is None or c_orig is None:
        sem_dimensao.append(el.Id)
        continue

    numero_folha_val = get_param_value(el, [nome_folha])
    marca_tipo_tela  = get_marca_tipo_tela(el)
    marca_hospedeiro = get_marca_hospedeiro(host)
    particao         = get_particao(el)

    l_key = round(l_orig, ARREDONDAMENTO)
    c_key = round(c_orig, ARREDONDAMENTO)

    # chave SEM marca_hospedeiro → agrupa entre lajes diferentes
    chave = (particao, marca_tipo_tela, l_key, c_key)

    if chave not in grupos:
        grupos[chave] = {
            "folhas":          [],
            "elementos":       [],
            "lajes":           set(),   # lajes hospedeiras que contribuíram
            "particao":        particao,
            "marca_tipo_tela": marca_tipo_tela,
            "l_orig":          l_orig,
            "c_orig":          c_orig,
        }

    grupos[chave]["folhas"].append(numero_folha_val)
    grupos[chave]["elementos"].append(el)
    grupos[chave]["lajes"].add(marca_hospedeiro)


# ===============================
# ORDENAÇÃO
# Partição (alfabética, "Sem Partição" por último)
# → menor folha → tipo tela → dimensões
# ===============================
grupos_ordenados = sorted(
    grupos.items(),
    key=lambda item: (
        ordem_particao(item[0][0]),
        min([numero_da_folha(f) for f in item[1]["folhas"]]),
        item[0][1].upper() if item[0][1] else "",
        item[0][2],
        item[0][3],
    )
)

# ===============================
# DESENHO
# ===============================
COLUNAS_MAX              = 5
ESPACAMENTO              = 3.0
ESPACAMENTO_ENTRE_GRUPOS = 5.0

cursor_x            = 0.0
cursor_y            = 0.0
altura_max_da_linha = 0.0
contador_coluna     = 0

particao_atual  = None
ids_processados = []

TransactionManager.Instance.EnsureInTransaction(doc)

# --- Vista de detalhamento ---
vt_types      = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
drafting_type = None
for t in vt_types:
    if t.ViewFamily == ViewFamily.Drafting:
        drafting_type = t
        break

nova_vista = ViewDrafting.Create(doc, drafting_type.Id)
nome_final = NOME_VISTA
c = 1
while True:
    try:
        nova_vista.Name = nome_final
        break
    except:
        nome_final = "{} ({})".format(NOME_VISTA, c)
        c += 1

text_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

# --- Filled region sólida ---
filled_region_type = None
col_frt = FilteredElementCollector(doc).OfClass(FilledRegionType)
for frt in col_frt:
    try:
        fp = frt.GetFillPattern()
        if fp and fp.IsSolidFill:
            filled_region_type = frt
            break
    except:
        continue
if filled_region_type is None:
    try:
        filled_region_type = col_frt.FirstElement()
    except:
        filled_region_type = None


# ===============================
# LOOP DE DESENHO
# ===============================
for chave, dados in grupos_ordenados:

    particao, marca_tipo_tela, l_key, c_key = chave
    particao_dados  = dados["particao"]
    folhas          = dados["folhas"]
    elementos_grupo = dados["elementos"]
    quantidade      = len(elementos_grupo)
    lajes_lista     = sorted(dados["lajes"], key=ordem_marca)

    larg_desenho = dados["c_orig"]
    alt_desenho  = dados["l_orig"]

    # --- Cabeçalho de PARTIÇÃO (quando muda) ---
    if particao_dados != particao_atual:
        if particao_atual is not None:
            cursor_x            = 0.0
            cursor_y           -= altura_max_da_linha + ESPACAMENTO_ENTRE_GRUPOS * 2
            altura_max_da_linha = 0.0
            contador_coluna     = 0

        options_pav = TextNoteOptions(text_type_id)
        options_pav.HorizontalAlignment = HorizontalTextAlignment.Left
        TextNote.Create(
            doc, nova_vista.Id,
            XYZ(cursor_x, cursor_y + 3.0, 0),
            u"===== {} =====".format(particao_dados.upper()),
            options_pav
        )
        particao_atual = particao_dados

    # --- Validação tamanho mínimo ---
    if larg_desenho < TOLERANCIA or alt_desenho < TOLERANCIA:
        for el in elementos_grupo:
            sem_dimensao.append(el.Id)
        cursor_x        += larg_desenho + ESPACAMENTO
        contador_coluna += 1
        if contador_coluna >= COLUNAS_MAX:
            cursor_x            = 0.0
            cursor_y           -= altura_max_da_linha + ESPACAMENTO + 2.5
            altura_max_da_linha = 0.0
            contador_coluna     = 0
        continue

    # --- Pontos do retângulo ---
    p1 = XYZ(cursor_x,                cursor_y,               0)
    p2 = XYZ(cursor_x + larg_desenho,  cursor_y,               0)
    p3 = XYZ(cursor_x + larg_desenho,  cursor_y + alt_desenho, 0)
    p4 = XYZ(cursor_x,                cursor_y + alt_desenho, 0)

    # --- Filled region ---
    if filled_region_type:
        try:
            curve_loop = CurveLoop()
            curve_loop.Append(Line.CreateBound(p1, p2))
            curve_loop.Append(Line.CreateBound(p2, p3))
            curve_loop.Append(Line.CreateBound(p3, p4))
            curve_loop.Append(Line.CreateBound(p4, p1))
            FilledRegion.Create(doc, filled_region_type.Id, nova_vista.Id, [curve_loop])
        except:
            pass

    linha_inf = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p1, p2))
    linha_dir = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p2, p3))
    linha_sup = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p3, p4))
    linha_esq = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p4, p1))

    try:
        doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p1, p3))
    except:
        pass

    # --- Texto folha na diagonal ---
    folhas_unicas = sorted(set(formatar_folha(f) for f in folhas if f != "-"), key=numero_da_folha)
    texto_folha   = ", ".join(folhas_unicas)

    ponto_medio = p1.Add(p3).Multiply(0.5)
    angulo_diag = math.atan2(alt_desenho, larg_desenho)

    options_folha = TextNoteOptions(text_type_id)
    options_folha.HorizontalAlignment = HorizontalTextAlignment.Center
    options_folha.Rotation = angulo_diag
    TextNote.Create(doc, nova_vista.Id, ponto_medio, texto_folha, options_folha)

    # --- Info abaixo do retângulo ---
    lajes_str  = ", ".join(lajes_lista) if lajes_lista else "-"
    texto_info = u"Parti\xe7\xe3o: {}\nLajes: {}\nTipo de tela: {}\nQtd: {}".format(
        particao_dados,
        lajes_str,
        marca_tipo_tela,
        quantidade
    )

    options_info = TextNoteOptions(text_type_id)
    options_info.HorizontalAlignment = HorizontalTextAlignment.Left
    TextNote.Create(
        doc, nova_vista.Id,
        XYZ(cursor_x, cursor_y - 2.5, 0),
        texto_info,
        options_info
    )

    # --- Cota largura ---
    try:
        ref_array_larg = ReferenceArray()
        ref_array_larg.Append(linha_esq.GeometryCurve.Reference)
        ref_array_larg.Append(linha_dir.GeometryCurve.Reference)
        doc.Create.NewDimension(nova_vista,
            Line.CreateBound(
                XYZ(cursor_x,                cursor_y - 1.0, 0),
                XYZ(cursor_x + larg_desenho,  cursor_y - 1.0, 0)
            ), ref_array_larg)
    except:
        pass

    # --- Cota altura ---
    try:
        ref_array_alt = ReferenceArray()
        ref_array_alt.Append(linha_inf.GeometryCurve.Reference)
        ref_array_alt.Append(linha_sup.GeometryCurve.Reference)
        doc.Create.NewDimension(nova_vista,
            Line.CreateBound(
                XYZ(cursor_x - 1.0, cursor_y,               0),
                XYZ(cursor_x - 1.0, cursor_y + alt_desenho, 0)
            ), ref_array_alt)
    except:
        pass

    for el in elementos_grupo:
        ids_processados.append(el.Id)

    if alt_desenho > altura_max_da_linha:
        altura_max_da_linha = alt_desenho

    cursor_x        += larg_desenho + ESPACAMENTO
    contador_coluna += 1

    if contador_coluna >= COLUNAS_MAX:
        cursor_x            = 0.0
        cursor_y           -= altura_max_da_linha + ESPACAMENTO + 2.5
        altura_max_da_linha = 0.0
        contador_coluna     = 0

TransactionManager.Instance.TransactionTaskDone()

OUT = nova_vista, ids_processados, ignorados, sem_dimensao
