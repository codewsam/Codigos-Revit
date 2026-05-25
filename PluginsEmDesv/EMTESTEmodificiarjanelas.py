# -*- coding: utf-8 -*-
__title__ = "Modificar Janelas\nEm Lote"
__version__ = "1.0"
__doc__ = "Modifica todos os parâmetros das janelas selecionadas do projeto em uma única operação."

from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# =============================================================================
# ETAPA 1 — Coletar todas as aberturas (Janelas)
# =============================================================================

# Filtrar todas as instâncias de janelas do projeto
todas_janelas = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_Windows)\
    .WhereElementIsNotElementType()\
    .ToElements()

if not todas_janelas:
    forms.alert("Nenhuma janela encontrada no projeto.", exitscript=True)

# =============================================================================
# ETAPA 2 — Agrupar janelas por Symbol (Tipo específico: J1, J2, J3...)
# =============================================================================

janelas_dict = {}
for janela in todas_janelas:
    try:
        # Usar o Symbol Name (tipo específico)
        symbol = janela.Symbol
        symbol_name = symbol.Name
        
        if symbol_name not in janelas_dict:
            janelas_dict[symbol_name] = []
        janelas_dict[symbol_name].append(janela)
    except:
        continue

if not janelas_dict:
    forms.alert("Erro ao processar as janelas.", exitscript=True)

# =============================================================================
# ETAPA 3 — Selecionar qual tipo de janela modificar (J1, J2, J3...)
# =============================================================================

tipos_disponiveis = sorted(janelas_dict.keys())
tipo_selecionado = forms.ask_for_one_item(
    tipos_disponiveis,
    default=tipos_disponiveis[0],
    prompt="Selecione qual tipo de janela você deseja modificar:",
    title="Seleção de Janela (J1, J2, J3...)"
)

if not tipo_selecionado:
    script.exit()

janelas_para_modificar = janelas_dict[tipo_selecionado]
qtd_janelas = len(janelas_para_modificar)

# =============================================================================
# ETAPA 4 — Listar parâmetros modificáveis
# =============================================================================

# Pegar uma janela como referência para listar parâmetros
janela_referencia = janelas_para_modificar[0]
parametros_dict = {}

for param in janela_referencia.Parameters:
    if not param.IsReadOnly:
        try:
            valor_atual = param.AsValueString()
            parametros_dict[param.Definition.Name] = param
        except:
            continue

if not parametros_dict:
    forms.alert("Nenhum parâmetro modificável encontrado.", exitscript=True)

# =============================================================================
# ETAPA 5 — Selecionar qual(is) parâmetro(s) modificar
# =============================================================================

parametros_nomes = sorted(parametros_dict.keys())
parametro_selecionado = forms.ask_for_one_item(
    parametros_nomes,
    default=parametros_nomes[0],
    prompt="Selecione qual parâmetro deseja modificar:",
    title="Seleção de Parâmetro"
)

if not parametro_selecionado:
    script.exit()

# =============================================================================
# ETAPA 6 — Solicitar o novo valor
# =============================================================================

novo_valor = forms.ask_for_string(
    default="",
    prompt="Digite o novo valor para '{}':\n\n(Janelas selecionadas: {})".format(
        parametro_selecionado, qtd_janelas
    ),
    title="Novo Valor"
)

if novo_valor is None:
    script.exit()

# =============================================================================
# ETAPA 7 — Aplicar modificação a todas as janelas
# =============================================================================

contador_ok = 0
contador_erro = 0

with Transaction(doc, "Modificar Janelas - {}".format(parametro_selecionado)) as t:
    t.Start()

    try:
        param_obj = parametros_dict[parametro_selecionado]

        for janela in janelas_para_modificar:
            try:
                param_instancia = janela.get_Parameter(param_obj.Definition)
                
                if param_instancia and not param_instancia.IsReadOnly:
                    # Tentar definir o valor conforme o tipo de parâmetro
                    if param_instancia.StorageType == StorageType.String:
                        param_instancia.Set(novo_valor)
                    elif param_instancia.StorageType == StorageType.Double:
                        param_instancia.Set(float(novo_valor))
                    elif param_instancia.StorageType == StorageType.Integer:
                        param_instancia.Set(int(novo_valor))
                    else:
                        param_instancia.SetValueString(novo_valor)
                    
                    contador_ok += 1
                else:
                    contador_erro += 1

            except Exception as e:
                contador_erro += 1
                print("Erro na janela {}: {}".format(janela.Id, str(e)))

        t.Commit()

    except Exception as e_geral:
        t.RollBack()
        forms.alert("Erro geral:\n{}".format(str(e_geral)), exitscript=True)

# =============================================================================
# Resumo final
# =============================================================================

forms.alert(
    "Modificação Concluída!\n\n"
    "Parâmetro: {}\n"
    "Novo valor: {}\n\n"
    "✔ Modificadas: {}\n"
    "✘ Com erro: {}".format(
        parametro_selecionado,
        novo_valor,
        contador_ok,
        contador_erro
    ),
    warn_icon=False
)

# =============================================================================
