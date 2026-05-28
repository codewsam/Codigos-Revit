# -*- coding: utf-8 -*-

__title__   = "Filtrar Paredes"
__author__  = "Samuel"
__version__ = "Versao 1.0"
"""
Plugin: Grupos de Paredes Espelhadas
Extensão: Samuel PLUGIN
Versão: 1.0.0
Autor: Samuel

Descrição:
    Permite ao usuário selecionar manualmente um conjunto de paredes
    e agrupá-las logicamente como "paredes espelhadas". Os dados são
    persistidos diretamente no arquivo Revit via Extensible Storage.

Persistência:
    Utiliza DataStorage + Extensible Storage (Schema/Entity) para
    salvar os grupos de forma permanente e invisível no modelo.
"""

# ─────────────────────────────────────────────
#  IMPORTS
# ─────────────────────────────────────────────
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import System
from System.Collections.Generic import List
from datetime import datetime

import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI
import Autodesk.Revit.UI.Selection as Sel
from Autodesk.Revit.DB.ExtensibleStorage import (
    Schema, SchemaBuilder, Entity,
    DataStorage, AccessLevel, FieldBuilder
)

from pyrevit import forms, script

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES GLOBAIS
# ─────────────────────────────────────────────

# GUID único do Schema — NÃO alterar após o primeiro uso em produção.
# Gerado uma única vez e fixo para garantir compatibilidade entre sessões.
SCHEMA_GUID = System.Guid("C7E3A912-4F5B-4D8E-9A1C-2B6D0F3E8C5A")
SCHEMA_NAME = "GruposParedesEspelhadas"

# Nome do DataStorage element que será criado no modelo
DATASTORAGE_NAME = "SamuelPlugin_GruposEspelhados"

# Prefixo dos campos no schema
FIELD_GRUPO_ID   = "grupo_id"
FIELD_WALL_IDS   = "wall_ids"
FIELD_DATA       = "data_criacao"
FIELD_NOME       = "nome_grupo"
FIELD_VERSAO     = "versao_schema"

VERSAO_SCHEMA = "1.0.0"


# ─────────────────────────────────────────────
#  SCHEMA — DEFINIÇÃO E RECUPERAÇÃO
# ─────────────────────────────────────────────

def obter_ou_criar_schema():
    """
    Retorna o Schema do Extensible Storage.
    Se já existir no documento, retorna o existente.
    Se não existir, cria e registra um novo.

    O Schema é uma estrutura de dados tipada, semelhante a uma
    tabela de banco de dados, que define os campos a serem armazenados.

    Retorna:
        Schema: objeto Schema registrado
    """
    schema_existente = Schema.Lookup(SCHEMA_GUID)
    if schema_existente:
        return schema_existente

    # Schema não existe: criar
    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.SetDocumentation(
        "Armazena grupos logicos de paredes espelhadas. "
        "Plugin Samuel PLUGIN v1.0.0"
    )

    # Campo: ID único do grupo (string UUID)
    builder.AddSimpleField(FIELD_GRUPO_ID, System.String)

    # Campo: IDs das paredes (lista de ElementId serializada como string CSV)
    # Extensible Storage não suporta List<ElementId> diretamente via IronPython,
    # por isso serializamos como "id1,id2,id3"
    builder.AddSimpleField(FIELD_WALL_IDS, System.String)

    # Campo: data de criação (string ISO)
    builder.AddSimpleField(FIELD_DATA, System.String)

    # Campo: nome amigável do grupo (expansão futura)
    builder.AddSimpleField(FIELD_NOME, System.String)

    # Campo: versão do schema (para migração futura)
    builder.AddSimpleField(FIELD_VERSAO, System.String)

    return builder.Finish()


# ─────────────────────────────────────────────
#  DATASTORAGE — ELEMENTO CONTAINER
# ─────────────────────────────────────────────

def obter_ou_criar_datastorage(doc):
    """
    Retorna o elemento DataStorage dedicado ao plugin.
    Se não existir no documento, cria um novo.

    O DataStorage é um elemento Revit invisível ao usuário,
    criado especificamente para armazenar dados customizados via
    Extensible Storage sem poluir elementos reais do modelo.

    Args:
        doc: documento Revit ativo

    Retorna:
        DataStorage: elemento container dos grupos
    """
    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DataStorage)
        .ToElements()
    )

    for ds in collector:
        if ds.Name == DATASTORAGE_NAME:
            return ds

    # Não encontrado: criar dentro de transação
    # (chamado externamente dentro de transação já aberta)
    ds_novo = DataStorage.Create(doc)
    ds_novo.Name = DATASTORAGE_NAME
    return ds_novo


# ─────────────────────────────────────────────
#  SERIALIZAÇÃO DOS GRUPOS
# ─────────────────────────────────────────────

def serializar_wall_ids(wall_ids):
    """
    Converte lista de ElementId para string CSV.

    Args:
        wall_ids: lista de DB.ElementId

    Retorna:
        str: "123,456,789"
    """
    return ",".join([str(wid.IntegerValue) for wid in wall_ids])


def deserializar_wall_ids(wall_ids_str):
    """
    Converte string CSV de volta para lista de ElementId.

    Args:
        wall_ids_str: str "123,456,789"

    Retorna:
        list[DB.ElementId]
    """
    if not wall_ids_str:
        return []
    return [DB.ElementId(int(id_str)) for id_str in wall_ids_str.split(",") if id_str.strip()]


def carregar_grupos(doc):
    """
    Lê todos os grupos salvos no DataStorage do documento.

    Cada grupo é um dicionário com as chaves:
        - grupo_id (str)
        - wall_ids (list[ElementId])
        - data_criacao (str)
        - nome_grupo (str)
        - versao_schema (str)

    Args:
        doc: documento Revit ativo

    Retorna:
        list[dict]: lista de grupos, ou [] se nenhum salvo
    """
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        return []

    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DataStorage)
        .ToElements()
    )

    for ds in collector:
        if ds.Name != DATASTORAGE_NAME:
            continue

        entity = ds.GetEntity(schema)
        if not entity.IsValid():
            return []

        # A lista de grupos é salva como um único campo JSON-like
        # Neste design usamos múltiplos DataStorage elements, um por grupo,
        # OR um único DS com serialização customizada.
        # Decisão: múltiplos DS elements nomeados por grupo_id (mais limpo e escalável)
        break

    # Abordagem final: um DS por grupo, todos nomeados com prefixo
    grupos = []
    for ds in collector:
        if not ds.Name.startswith(DATASTORAGE_NAME + "__"):
            continue

        entity = ds.GetEntity(schema)
        if not entity.IsValid():
            continue

        grupo = {
            "grupo_id":     entity.Get[System.String](FIELD_GRUPO_ID),
            "wall_ids":     deserializar_wall_ids(entity.Get[System.String](FIELD_WALL_IDS)),
            "data_criacao": entity.Get[System.String](FIELD_DATA),
            "nome_grupo":   entity.Get[System.String](FIELD_NOME),
            "versao_schema": entity.Get[System.String](FIELD_VERSAO),
        }
        grupos.append(grupo)

    return grupos


def salvar_grupo(doc, wall_ids, nome_grupo=""):
    """
    Persiste um novo grupo de paredes no Extensible Storage.

    Cria um elemento DataStorage dedicado para este grupo,
    nomeado com o padrão: DATASTORAGE_NAME__<grupo_id>

    Args:
        doc:       documento Revit ativo
        wall_ids:  list[DB.ElementId] — paredes selecionadas
        nome_grupo: str — nome opcional do grupo

    Retorna:
        str: grupo_id gerado
    """
    schema = obter_ou_criar_schema()

    # Gerar ID único para o grupo
    grupo_id = str(System.Guid.NewGuid())
    data_criacao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not nome_grupo:
        nome_grupo = "Grupo_" + grupo_id[:8]

    # Criar DataStorage dedicado para este grupo
    t = DB.Transaction(doc, "Salvar Grupo Espelhado")
    t.Start()
    try:
        ds = DataStorage.Create(doc)
        ds.Name = DATASTORAGE_NAME + "__" + grupo_id

        # Criar Entity com os dados do grupo
        entity = Entity(schema)
        entity.Set[System.String](FIELD_GRUPO_ID,   grupo_id)
        entity.Set[System.String](FIELD_WALL_IDS,   serializar_wall_ids(wall_ids))
        entity.Set[System.String](FIELD_DATA,        data_criacao)
        entity.Set[System.String](FIELD_NOME,        nome_grupo)
        entity.Set[System.String](FIELD_VERSAO,      VERSAO_SCHEMA)

        ds.SetEntity(entity)

        t.Commit()
        return grupo_id

    except Exception as ex:
        t.RollbackIfOpen()
        raise ex


# ─────────────────────────────────────────────
#  SELEÇÃO DE PAREDES
# ─────────────────────────────────────────────

class FiltroParedes(Sel.ISelectionFilter):
    """
    Filtro de seleção que aceita apenas elementos do tipo Wall.
    Implementa ISelectionFilter da Revit API.
    """

    def AllowElement(self, element):
        return isinstance(element, DB.Wall)

    def AllowReference(self, reference, point):
        return False


def solicitar_selecao_paredes(uidoc):
    """
    Abre o modo de seleção interativa do Revit para o usuário
    escolher paredes. Usa filtro para aceitar apenas Wall.

    Args:
        uidoc: UIDocument ativo

    Retorna:
        list[DB.ElementId]: IDs das paredes selecionadas
        None: se cancelado ou seleção vazia
    """
    filtro = FiltroParedes()

    try:
        referencias = uidoc.Selection.PickObjects(
            Sel.ObjectType.Element,
            filtro,
            "Selecione as paredes do grupo espelhado. [ESC para cancelar]"
        )
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        return None

    if not referencias:
        return None

    return [ref.ElementId for ref in referencias]


# ─────────────────────────────────────────────
#  VALIDAÇÃO
# ─────────────────────────────────────────────

def validar_selecao(doc, wall_ids):
    """
    Valida a seleção antes de salvar.

    Verificações:
    - Mínimo de 2 paredes selecionadas
    - Todas as paredes ainda existem no documento
    - Nenhuma parede já pertence a um grupo existente

    Args:
        doc:      documento Revit ativo
        wall_ids: list[DB.ElementId]

    Retorna:
        (bool, str): (valido, mensagem_erro)
    """
    if len(wall_ids) < 2:
        return False, "Selecione pelo menos 2 paredes para formar um grupo."

    # Verificar existência
    for wid in wall_ids:
        el = doc.GetElement(wid)
        if el is None or not isinstance(el, DB.Wall):
            return False, "Uma ou mais paredes selecionadas não foram encontradas no modelo."

    # Verificar duplicidade em grupos existentes
    grupos_existentes = carregar_grupos(doc)
    ids_em_uso = set()
    for grupo in grupos_existentes:
        for wid in grupo["wall_ids"]:
            ids_em_uso.add(wid.IntegerValue)

    conflitos = [wid for wid in wall_ids if wid.IntegerValue in ids_em_uso]
    if conflitos:
        return False, (
            "{} parede(s) já pertencem a um grupo existente.\n"
            "Cada parede pode pertencer a apenas um grupo."
        ).format(len(conflitos))

    return True, ""


# ─────────────────────────────────────────────
#  PONTO DE ENTRADA DO PLUGIN
# ─────────────────────────────────────────────

def main():
    """
    Função principal do plugin.

    Fluxo:
    1. Solicitar seleção de paredes
    2. Validar seleção
    3. Solicitar nome do grupo (opcional)
    4. Salvar grupo via Extensible Storage
    5. Exibir confirmação
    """
    doc   = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument

    # ── Passo 1: Seleção
    wall_ids = solicitar_selecao_paredes(uidoc)

    if wall_ids is None:
        forms.alert("Operação cancelada.", title="Grupos Espelhados", warn_icon=False)
        return

    # ── Passo 2: Validação
    valido, erro = validar_selecao(doc, wall_ids)
    if not valido:
        forms.alert(erro, title="Grupos Espelhados — Erro de Validação")
        return

    # ── Passo 3: Nome do grupo (opcional)
    nome_grupo = forms.ask_for_string(
        default="",
        prompt="Nome do grupo (deixe em branco para gerar automaticamente):",
        title="Grupos Espelhados"
    )
    if nome_grupo is None:
        # Usuário clicou em Cancelar no diálogo de nome
        forms.alert("Operação cancelada.", title="Grupos Espelhados", warn_icon=False)
        return

    # ── Passo 4: Salvar
    try:
        grupo_id = salvar_grupo(doc, wall_ids, nome_grupo=nome_grupo)
    except Exception as ex:
        forms.alert(
            "Erro ao salvar o grupo:\n\n{}".format(str(ex)),
            title="Grupos Espelhados — Erro"
        )
        return

    # ── Passo 5: Confirmação
    forms.alert(
        "Grupo criado com sucesso!\n\n"
        "ID:      {}\n"
        "Paredes: {}\n"
        "Nome:    {}".format(
            grupo_id[:18] + "...",
            len(wall_ids),
            nome_grupo if nome_grupo else "Grupo_" + grupo_id[:8]
        ),
        title="Grupos Espelhados",
        warn_icon=False
    )


# ─────────────────────────────────────────────
#  EXECUÇÃO
# ─────────────────────────────────────────────
if __name__ == '__main__':
    main()
