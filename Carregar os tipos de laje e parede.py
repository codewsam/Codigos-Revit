import sys
import clr
import math

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

doc = DocumentManager.Instance.CurrentDBDocument

# ===============================
# ENTRADAS DO DYNAMO
# ===============================
elements = UnwrapElement(IN[0])

nome_largura = IN[1] # "Largura total do corte"
nome_comprimento = IN[2] # "Comprimento total do corte"
nome_folha = IN[3] # "Número da folha de tela soldada"

# ===============================
# FUNÇÕES AUXILIARES
# ===============================
def get_param_value(el, nomes):
    for nome in nomes:
        try:
            p = el.LookupParameter(nome)
            if p and p.HasValue:
                if p.StorageType == StorageType.String:
                    valor = p.AsString()
                    return valor if valor else "-"
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

def tipo_hospedeiro(host):
    if host is None:
        return "SEM HOSPEDEIRO"

    if host.Category is None:
        return "SEM CATEGORIA"

    cat_id = host.Category.Id.IntegerValue

    if cat_id == int(BuiltInCategory.OST_Walls):
        return "PAREDE"

    if cat_id == int(BuiltInCategory.OST_Floors):
        return "LAJE"

    return host.Category.Name.upper()

# ===============================
# AGRUPAMENTO
# ===============================
grupos = {}
sem_dimensao = []

for el in elements:
    host = get_host(el)

    tipo_host = tipo_hospedeiro(host)

    
    if host:
    tipo_elemento = doc.GetElement(host.GetTypeId())
    marca_tipo = get_param_value(tipo_elemento, ["Marca de tipo", "Type Mark"])
else:
    marca_tipo = "-"

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

    numero_folha = get_param_value(el, [nome_folha])

    if host:
        marca_host = get_param_value(host, ["Marca", "Mark"])
    else:
        marca_host = "-"

    particao = get_param_value(el, [
        "Partição",
        "Partition",
        "Comentários",
        "Comments"
    ])

    # Agora separa por:
    # tipo de hospedeiro + marca do hospedeiro + partição + largura + comprimento
    chave = (tipo_host, marca_tipo, particao, l_orig, c_orig)

    if chave not in grupos:
        grupos[chave] = {
            "folhas": [],
            "elementos": [],
            "host": host
        }

    grupos[chave]["folhas"].append(numero_folha)
    grupos[chave]["elementos"].append(el)

# ===============================
# CONFIGURAÇÕES DO DESENHO
# ===============================
COLUNAS_MAX = 5
ESPACAMENTO = 3.0

cursor_x = 0.0
cursor_y = 0.0
altura_max_da_linha = 0.0
contador_coluna = 0

ids_processados = []

TransactionManager.Instance.EnsureInTransaction(doc)

# ===============================
# CRIAR VISTA DE DESENHO
# ===============================
vt_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
drafting_type = None

for t in vt_types:
    if t.ViewFamily == ViewFamily.Drafting:
        drafting_type = t
        break

nova_vista = ViewDrafting.Create(doc, drafting_type.Id)

nome_base = "Esquema de corte - telas por hospedeiro"
nome_final = nome_base
c = 1

while True:
    try:
        nova_vista.Name = nome_final
        break
    except:
        nome_final = "{} ({})".format(nome_base, c)
        c += 1

text_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

# ===============================
# REGIÃO PREENCHIDA
# ===============================
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
# DESENHAR CADA GRUPO
# ===============================
for chave, dados in grupos.items():

    tipo_host, marca_tipo, particao, l_orig, c_orig = chave

    folhas = dados["folhas"]
    elementos_grupo = dados["elementos"]

    quantidade = len(elementos_grupo)

    larg_desenho = c_orig
    alt_desenho = l_orig

    p1 = XYZ(cursor_x, cursor_y, 0)
    p2 = XYZ(cursor_x + larg_desenho, cursor_y, 0)
    p3 = XYZ(cursor_x + larg_desenho, cursor_y + alt_desenho, 0)
    p4 = XYZ(cursor_x, cursor_y + alt_desenho, 0)

    # Região preenchida
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

    # Linhas externas
    linha_inf = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p1, p2))
    linha_dir = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p2, p3))
    linha_sup = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p3, p4))
    linha_esq = doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p4, p1))

    # Diagonal
    doc.Create.NewDetailCurve(nova_vista, Line.CreateBound(p1, p3))

    # Texto da folha na diagonal
    ponto_medio = p1.Add(p3).Multiply(0.5)
    angulo_diag = math.atan2(alt_desenho, larg_desenho)

    folhas_unicas = sorted(set(folhas))

    folhas_texto = []
    for f in folhas_unicas:
        f = str(f)
        if f.upper().startswith("N"):
            folhas_texto.append(f)
        else:
            folhas_texto.append("N" + f)

    texto_folha = ", ".join(folhas_texto)

    options_folha = TextNoteOptions(text_type_id)
    options_folha.HorizontalAlignment = HorizontalTextAlignment.Center
    options_folha.Rotation = angulo_diag

    TextNote.Create(doc, nova_vista.Id, ponto_medio, texto_folha, options_folha)

    # Texto embaixo do desenho
    texto_info = "Tipo: {}\nHospedeiro: {}\nPartição: {}\nQtd: {}".format(
        tipo_host,
        marca_tipo,
        particao,
        quantidade
    )

    ponto_texto_info = XYZ(cursor_x, cursor_y - 2.0, 0)

    options_info = TextNoteOptions(text_type_id)
    options_info.HorizontalAlignment = HorizontalTextAlignment.Left

    TextNote.Create(doc, nova_vista.Id, ponto_texto_info, texto_info, options_info)

    # Cota horizontal
    ref_array_larg = ReferenceArray()
    ref_array_larg.Append(linha_esq.GeometryCurve.Reference)
    ref_array_larg.Append(linha_dir.GeometryCurve.Reference)

    linha_cota_larg = Line.CreateBound(
        XYZ(cursor_x, cursor_y - 1.0, 0),
        XYZ(cursor_x + larg_desenho, cursor_y - 1.0, 0)
    )

    doc.Create.NewDimension(nova_vista, linha_cota_larg, ref_array_larg)

    # Cota vertical
    ref_array_alt = ReferenceArray()
    ref_array_alt.Append(linha_inf.GeometryCurve.Reference)
    ref_array_alt.Append(linha_sup.GeometryCurve.Reference)

    linha_cota_alt = Line.CreateBound(
        XYZ(cursor_x - 1.0, cursor_y, 0),
        XYZ(cursor_x - 1.0, cursor_y + alt_desenho, 0)
    )

    doc.Create.NewDimension(nova_vista, linha_cota_alt, ref_array_alt)

    for el in elementos_grupo:
        ids_processados.append(el.Id)

    # Atualizar posição dos próximos desenhos
    if alt_desenho > altura_max_da_linha:
        altura_max_da_linha = alt_desenho

    cursor_x += larg_desenho + ESPACAMENTO
    contador_coluna += 1

    if contador_coluna >= COLUNAS_MAX:
        cursor_x = 0.0
        cursor_y -= altura_max_da_linha + ESPACAMENTO + 2.0
        altura_max_da_linha = 0.0
        contador_coluna = 0

TransactionManager.Instance.TransactionTaskDone()

OUT = nova_vista, ids_processados, sem_dimensao