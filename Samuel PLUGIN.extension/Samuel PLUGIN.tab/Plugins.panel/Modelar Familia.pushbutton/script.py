# -*- coding: utf-8 -*-
__title__ = "Modelar Famílias\nCentros de Ambientes"
__version__ = "1.0"
__doc__ = "Insere famílias nos centros dos ambientes selecionados, posicionando-as na altura (UnboundedHeight) do ambiente."

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# =============================================================================
# ETAPA 2 — Formulário: selecionar Família e Tipo
# =============================================================================

# Coletar todos os FamilySymbols carregados no projeto
todos_symbols = FilteredElementCollector(doc)\
    .OfClass(FamilySymbol)\
    .ToElements()

# Montar dicionário  "Nome da Família : Nome do Tipo"  →  FamilySymbol
symbol_dict = {}
for sym in todos_symbols:
    familia_nome = sym.FamilyName
    tipo_nome    = sym.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    chave        = "{} : {}".format(familia_nome, tipo_nome)
    symbol_dict[chave] = sym

if not symbol_dict:
    forms.alert("Nenhuma família carregada no projeto.", exitscript=True)

# Formulário para escolher uma família/tipo
chave_selecionada = forms.ask_for_one_item(
    sorted(symbol_dict.keys()),
    default=sorted(symbol_dict.keys())[0],
    prompt="Selecione a Família e o Tipo a inserir:",
    title="Seleção da Família"
)

if not chave_selecionada:
    script.exit()

symbol_escolhido = symbol_dict[chave_selecionada]

# =============================================================================
# ETAPA 3 — Seleção múltipla de Ambientes
# =============================================================================

with forms.WarningBar(title="Selecione os Ambientes e pressione Finish"):
    rooms = revit.pick_elements_by_category(BuiltInCategory.OST_Rooms)

if not rooms:
    forms.alert("Nenhum ambiente selecionado.", exitscript=True)

# =============================================================================
# ETAPAS 4 e 5 — Inserir instância no centro de cada ambiente + definir offset
# =============================================================================

contador_ok    = 0
contador_erro  = 0

with Transaction(doc, "Modelar Luminárias nos Centros de Ambientes") as t:
    t.Start()

    try:
        # Ativar o symbol caso ainda não esteja ativo
        if not symbol_escolhido.IsActive:
            symbol_escolhido.Activate()
            doc.Regenerate()

        for room in rooms:
            try:
                # ETAPA 4 — Obter o centro do ambiente (LocationPoint)
                localizacao = room.Location
                if localizacao is None:
                    contador_erro += 1
                    continue

                centro_room = localizacao.Point   # XYZ com Z = nível do ambiente

                # ETAPA 5 — Calcular Z de inserção = Z do ponto + altura do ambiente
                altura_ambiente = room.UnboundedHeight   # em pés (unidade interna)
                ponto_insercao  = XYZ(centro_room.X, centro_room.Y, centro_room.Z + altura_ambiente)

                # Criar instância da família no centro + altura do ambiente
                instancia = doc.Create.NewFamilyInstance(
                    ponto_insercao,
                    symbol_escolhido,
                    StructuralType.NonStructural
                )

                # ETAPA 5 — Definir o offset de nível (FAMILY_LEVEL_OFFSET_PARAM)
                param_offset = instancia.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                if param_offset and not param_offset.IsReadOnly:
                    param_offset.Set(altura_ambiente)

                contador_ok += 1

            except Exception as e_room:
                contador_erro += 1
                print("Erro no ambiente '{}': {}".format(
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), str(e_room)
                ))

        t.Commit()

    except Exception as e_geral:
        t.RollBack()
        forms.alert("Erro geral:\n{}".format(str(e_geral)), exitscript=True)

# =============================================================================
# Resumo final
# =============================================================================
forms.alert(
    "Concluído!\n\n✔ Inseridas: {}\n✘ Com erro: {}".format(contador_ok, contador_erro),
    warn_icon=False
)