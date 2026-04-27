# -*- coding: utf-8 -*-
"""
CRIAR TABELAS - Rotina Dynamo

IN[0]  - bool  -> Criar tabela de Vergalhões?
IN[1]  - bool  -> Criar tabela de Tela Soldada?
IN[2]  - str   -> Nome da tabela de Vergalhões
IN[3]  - str   -> Nome da tabela de Tela Soldada
IN[4]  - list  -> Campos de Vergalhões a incluir
IN[5]  - list  -> Campos de Tela Soldada a incluir
IN[6]  - str   -> Filtro por LOCAL
IN[7]  - bool  -> Inserir na folha ativa?
"""

import clr

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    ScheduleSortGroupField,
    ScheduleSortOrder,
    ScheduleFilter,
    ScheduleFilterType,
    ElementId,
    ScheduleSheetInstance,
    UV,
    ViewSheet,
)

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument


# ==========================================================
# INPUTS
# ==========================================================

fazer_verg = IN[0] if len(IN) > 0 and IN[0] is not None else False
fazer_tela = IN[1] if len(IN) > 1 and IN[1] is not None else False

nome_verg = IN[2] if len(IN) > 2 and IN[2] else u"TABELA DE ARMACAO ABERTURAS"
nome_tela = IN[3] if len(IN) > 3 and IN[3] else u"TABELA DE ARMADURA DE TELA SOLDADA"

campos_verg = IN[4] if len(IN) > 4 else []
campos_tela = IN[5] if len(IN) > 5 else []

filtro = IN[6] if len(IN) > 6 and IN[6] else None
na_folha = IN[7] if len(IN) > 7 and IN[7] is not None else False


# ==========================================================
# CATEGORIAS
# ==========================================================

CAT_VERGALHAO = -2009000   # Vergalhão / Rebar
CAT_TELA = -2009016        # Tela soldada / Fabric Sheets


# ==========================================================
# CAMPOS PADRÃO
# ==========================================================

CAMPOS_VERG_PADRAO = [
    (u"LOCAL",           u"Parti"),
    (u"POSICAO (P)",     u"Numero do vergalhao"),
    (u"QUANTIDADE",      u"Quantidade por conjunto"),
    (u"DIAMETRO (mm)",   u"Diametro da barra"),
    (u"COMPRIMENTO(cm)", u"Comprimento da barra"),
    (u"COMP.TOTAL (m)",  u"Comprimento total da barra"),
    (u"PESO (kg)",       u"Peso"),
]

CAMPOS_TELA_PADRAO = [
    (u"LOCAL",       u"Marca do hospedeiro"),
    (u"N",           u"Numero da folha de tela"),
    (u"QUANT.",      u"Contagem"),
    (u"TELA",        u"Marca de tipo"),
    (u"LARGURA",     u"Largura total do corte"),
    (u"COMPRIMENTO", u"Comprimento total do corte"),
    (u"PESO (Kgf)",  u"Massa da folha de corte"),
]


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================

def to_python_list(valor):
    """
    Converte lista do Dynamo para lista normal do Python.
    """
    if valor is None:
        return []

    if isinstance(valor, str):
        if valor.strip() == "":
            return []
        return [valor]

    try:
        return list(valor)
    except:
        return [valor]


def get_nome_unico(nome_base):
    """
    Evita erro quando já existe tabela com o mesmo nome.
    Exemplo:
    TABELA
    TABELA (1)
    TABELA (2)
    """
    existentes = set()

    tabelas = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()

    for tabela in tabelas:
        try:
            existentes.add(tabela.Name)
        except:
            pass

    nome = nome_base
    i = 1

    while nome in existentes:
        nome = u"{} ({})".format(nome_base, i)
        i += 1

    return nome


def get_field_by_name(sched, keyword):
    """
    Procura um campo disponível na tabela.
    A busca é parcial.
    Exemplo:
    keyword = 'Peso'
    Pode achar 'Peso da barra' ou 'Peso total'
    """
    if not keyword:
        return None

    sd = sched.Definition
    keyword = keyword.lower()

    for sf in sd.GetSchedulableFields():
        try:
            nome_campo = sf.GetName(doc)
            if keyword in nome_campo.lower():
                return sf
        except:
            pass

    return None


def montar_lista_campos(campos_input, campos_padrao):
    """
    Se IN[4] ou IN[5] estiver vazio, usa os campos padrão.
    Se tiver lista de strings, usa a lista enviada pelo Dynamo.
    """
    campos = to_python_list(campos_input)

    campos_limpos = []

    for campo in campos:
        if campo:
            texto = str(campo).strip()
            if texto:
                campos_limpos.append(texto)

    if len(campos_limpos) == 0:
        return campos_padrao

    return [(kw, kw) for kw in campos_limpos]


def criar_tabela(nome, cat_id_int, lista_campos, filtro_txt=None, filtro_kw=None):
    cat_id = ElementId(cat_id_int)
    nome_final = get_nome_unico(nome)

    sched = ViewSchedule.CreateSchedule(doc, cat_id)
    sched.Name = nome_final

    sd = sched.Definition
    sd.ShowHeaders = True

    primeiro_campo_id = None
    campos_adicionados = 0
    campos_nao_encontrados = []

    for header, kw in lista_campos:
        sf = get_field_by_name(sched, kw)

        if sf:
            campo = sd.AddField(sf)
            campo.ColumnHeading = header

            if primeiro_campo_id is None:
                primeiro_campo_id = campo.FieldId

            campos_adicionados += 1
        else:
            campos_nao_encontrados.append(kw)

    if primeiro_campo_id:
        try:
            ordenacao = ScheduleSortGroupField(
                primeiro_campo_id,
                ScheduleSortOrder.Ascending
            )
            sd.AddSortGroupField(ordenacao)
        except:
            pass

    if filtro_txt and filtro_kw:
        for i in range(sd.GetFieldCount()):
            try:
                f = sd.GetField(i)
                nome_filtro = f.GetName()

                if filtro_kw.lower() in nome_filtro.lower():
                    filtro_tabela = ScheduleFilter(
                        f.FieldId,
                        ScheduleFilterType.Contains,
                        filtro_txt
                    )
                    sd.AddFilter(filtro_tabela)
                    break
            except:
                pass

    return sched, nome_final, campos_adicionados, campos_nao_encontrados


def inserir_na_folha(sched):
    """
    Insere a tabela na folha atualmente aberta.
    Só funciona se a vista ativa for uma folha.
    """
    vista = uidoc.ActiveView

    if not isinstance(vista, ViewSheet):
        return False

    try:
        ScheduleSheetInstance.Create(
            doc,
            vista.Id,
            sched.Id,
            UV(0.15, 0.15)
        )
        return True
    except:
        return False


# ==========================================================
# EXECUÇÃO
# ==========================================================

resultados = []

TransactionManager.Instance.EnsureInTransaction(doc)

try:
    if fazer_verg:
        lista_v = montar_lista_campos(campos_verg, CAMPOS_VERG_PADRAO)

        sched_v, nome_v, qtd_v, nao_v = criar_tabela(
            nome_verg,
            CAT_VERGALHAO,
            lista_v,
            filtro_txt=filtro,
            filtro_kw=u"parti"
        )

        msg = u"OK - VERGALHOES: '{}' | {} campos criados".format(nome_v, qtd_v)

        if len(nao_v) > 0:
            msg += u" | Campos nao encontrados: {}".format(u", ".join(nao_v))

        if na_folha:
            inseriu = inserir_na_folha(sched_v)
            if inseriu:
                msg += u" | Inserida na folha ativa"
            else:
                msg += u" | Nao inserida: abra uma folha antes de rodar"

        resultados.append(msg)

    if fazer_tela:
        lista_t = montar_lista_campos(campos_tela, CAMPOS_TELA_PADRAO)

        sched_t, nome_t, qtd_t, nao_t = criar_tabela(
            nome_tela,
            CAT_TELA,
            lista_t,
            filtro_txt=filtro,
            filtro_kw=u"hospedeiro"
        )

        msg = u"OK - TELA SOLDADA: '{}' | {} campos criados".format(nome_t, qtd_t)

        if len(nao_t) > 0:
            msg += u" | Campos nao encontrados: {}".format(u", ".join(nao_t))

        if na_folha:
            inseriu = inserir_na_folha(sched_t)
            if inseriu:
                msg += u" | Inserida na folha ativa"
            else:
                msg += u" | Nao inserida: abra uma folha antes de rodar"

        resultados.append(msg)

    if not fazer_verg and not fazer_tela:
        resultados.append(u"Nenhum tipo selecionado. Coloque True em IN[0] ou IN[1].")

except Exception as ex:
    resultados.append(u"ERRO: {}".format(str(ex)))

finally:
    TransactionManager.Instance.TransactionTaskDone()

OUT = resultados
