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
elements = UnwrapElement(IN[0])

nome_largura = IN[1]
nome_comprimento = IN[2]
nome_folha = IN[3]

# ===============================
# CONFIGURAÇÃO
# ===============================
CATEGORIA_DESEJADA = BuiltInCategory.OST_Walls
NOME_VISTA = "Esquema de corte - PAREDES"

# ===============================
# FUNÇÕES
# ===============================

def numero_da_folha(valor):
    try:
        texto = str(valor).upper().strip()
        texto = texto.replace("N", "")
        numeros = re.findall(r'\d+', texto)
        if numeros:
            return int(numeros[0])
    except:
        pass
    return 999999


def formatar_folha(valor):
    texto = str(valor).upper().strip()
    if texto.startswith("N"):
        return texto
    return "N" + texto


def ler_parametro(el, nomes, padrao="-"):
    if el is None:
        return padrao

    for nome in nomes:
        try:
            p = el.LookupParameter(nome)
            if p and p.HasValue:
                if p.StorageType == StorageType.String:
                    valor = p.AsString()
                    return valor if valor else padrao
                elif p.StorageType == StorageType.Integer:
                    return str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    return str(round(p.AsDouble(), 3))
                elif p.StorageType == StorageType.ElementId:
                    return str(p.AsElementId().IntegerValue)
        except:
            pass

    return padrao


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
        tipo_tela = doc.GetElement(el.GetTypeId())
        p_marca_tipo = tipo_tela.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
        if p_marca_tipo and p_marca_tipo.HasValue:
            valor = p_marca_tipo.AsString()
            return valor if valor else "-"
    except:
        pass
    return "-"


def get_marca_hospedeiro(host):
    if host is None:
        return "-"
    try:
        p = host.LookupParameter("Marca")
        if p and p.HasValue:
            valor = p.AsString()
            return valor if valor else "-"
    except:
        pass
    try:
        p = host.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if p and p.HasValue:
            valor = p.AsString()
            return valor if valor else "-"
    except:
        pass
    return "-"


def get_particao(el):
    try:
        p = el.LookupParameter(u"Partição")
        if p and p.HasValue:
            valor = p.AsString()
            return valor if valor else "Sem Partição"
    except:
        pass
    return "Sem Partição"


# ===============================
# NOVO: NIVEL → VISTA
# ===============================
def get_nome_vista_parede(host):
    if host is None:
        return "Nível indefinido"

    nome_nivel = None

    try:
        p = host.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        if p and p.HasValue:
            nivel_id = p.AsElementId()
            nivel = doc.GetElement(nivel_id)
            if nivel:
                nome_nivel = nivel.Name
    except:
        pass

    if not nome_nivel:
        try:
            nivel = doc.GetElement(host.LevelId)
            if nivel:
                nome_nivel = nivel.Name
        except:
            pass

    if not nome_nivel:
        return "Nível indefinido"

    texto = nome_nivel.upper()

    if "COBERTURA" in texto:
        return "Superior"

    if "SUPERIOR" in texto:
        return "Térreo"

    if "TERREO" in texto or "TÉRREO" in texto:
        return "Térreo"

    return "Nível indefinido"


# ===============================
# AGRUPAMENTO
# ===============================
grupos = {}
ignorados = []
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

    numero_folha = ler_parametro(el, [nome_folha])
    marca_tipo_tela = get_marca_tipo_tela(el)
    marca_hospedeiro = get_marca_hospedeiro(host)
    particao = get_particao(el)
    nome_vista_parede = get_nome_vista_parede(host)

    chave = (
        nome_vista_parede,
        particao,
        marca_hospedeiro,
        marca_tipo_tela,
        l_orig,
        c_orig
    )

    if chave not in grupos:
        grupos[chave] = {
            "folhas": [],
            "elementos": [],
            "nome_vista_parede": nome_vista_parede,
            "marca_hospedeiro": marca_hospedeiro,
            "particao": particao
        }

    grupos[chave]["folhas"].append(numero_folha)
    grupos[chave]["elementos"].append(el)


# ===============================
# ORDENAÇÃO
# ===============================
grupos_ordenados = sorted(
    grupos.items(),
    key=lambda item: (
        item[0][0],
        item[0][1],
        item[0][2],
        min([numero_da_folha(f) for f in item[1]["folhas"]]),
        item[0][3],
        item[0][4],
        item[0][5]
    )
)

# ===============================
# DESENHO
# ===============================
COLUNAS_MAX = 5
ESPACAMENTO = 3.0
ESPACAMENTO_ENTRE_GRUPOS = 5.0

cursor_x = 0.0
cursor_y = 0.0
altura_max_da_linha = 0.0
contador_coluna = 0

nome_vista_atual = None
particao_atual = None
grupo_atual = None

ids_processados = []

TransactionManager.Instance.EnsureInTransaction(doc)

vt_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
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

for chave, dados in grupos_ordenados:

    nome_vista_parede, particao, marca_hospedeiro, marca_tipo_tela, l_orig, c_orig = chave

    if nome_vista_parede != nome_vista_atual:

        if nome_vista_atual is not None:
            cursor_x = 0.0
            cursor_y -= altura_max_da_linha + ESPACAMENTO_ENTRE_GRUPOS * 3
            altura_max_da_linha = 0.0
            contador_coluna = 0

        TextNote.Create(
            doc,
            nova_vista.Id,
            XYZ(cursor_x, cursor_y + 5.0, 0),
            "########## VISTA: {} ##########".format(nome_vista_parede.upper()),
            TextNoteOptions(text_type_id)
        )

        nome_vista_atual = nome_vista_parede
        particao_atual = None
        grupo_atual = None

    if particao != particao_atual:

        if particao_atual is not None:
            cursor_x = 0.0
            cursor_y -= altura_max_da_linha + ESPACAMENTO_ENTRE_GRUPOS * 2
            altura_max_da_linha = 0.0
            contador_coluna = 0

        TextNote.Create(
            doc,
            nova_vista.Id,
            XYZ(cursor_x, cursor_y + 3.0, 0),
            "===== {} =====".format(particao.upper()),
            TextNoteOptions(text_type_id)
        )

        particao_atual = particao
        grupo_atual = None

    if marca_hospedeiro != grupo_atual:

        if grupo_atual is not None:
            cursor_x = 0.0
            cursor_y -= altura_max_da_linha + ESPACAMENTO_ENTRE_GRUPOS
            altura_max_da_linha = 0.0
            contador_coluna = 0

        TextNote.Create(
            doc,
            nova_vista.Id,
            XYZ(cursor_x, cursor_y + 1.5, 0),
            "PAREDE - {}".format(marca_hospedeiro),
            TextNoteOptions(text_type_id)
        )

        grupo_atual = marca_hospedeiro

    folhas = dados["folhas"]
    elementos_grupo = dados["elementos"]
    quantidade = len(elementos_grupo)

    larg_desenho = c_orig
    alt_desenho = l_orig

    p1 = XYZ(cursor_x, cursor_y, 0)
    p3 = XYZ(cursor_x + larg_desenho, cursor_y + alt_desenho, 0)

    ponto_medio = p1.Add(p3).Multiply(0.5)

    folhas_unicas = sorted(set(folhas), key=numero_da_folha)
    folhas_texto = [formatar_folha(f) for f in folhas_unicas]

    TextNote.Create(
        doc,
        nova_vista.Id,
        ponto_medio,
        ", ".join(folhas_texto),
        TextNoteOptions(text_type_id)
    )

    texto_info = "Vista: {}\nLocal: {}\nParede: {}\nTipo de tela: {}\nQtd: {}".format(
        nome_vista_parede,
        particao,
        marca_hospedeiro,
        marca_tipo_tela,
        quantidade
    )

    TextNote.Create(
        doc,
        nova_vista.Id,
        XYZ(cursor_x, cursor_y - 2.5, 0),
        texto_info,
        TextNoteOptions(text_type_id)
    )

    for el in elementos_grupo:
        ids_processados.append(el.Id)

    if alt_desenho > altura_max_da_linha:
        altura_max_da_linha = alt_desenho

    cursor_x += larg_desenho + ESPACAMENTO
    contador_coluna += 1

    if contador_coluna >= COLUNAS_MAX:
        cursor_x = 0.0
        cursor_y -= altura_max_da_linha + ESPACAMENTO + 2.5
        altura_max_da_linha = 0.0
        contador_coluna = 0

TransactionManager.Instance.TransactionTaskDone()

OUT = nova_vista, ids_processados, ignorados, sem_dimensao