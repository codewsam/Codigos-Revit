# -*- coding: utf-8 -*-
__title__ = "Modelar Pisos\nAutomático"
__version__ = "1.0"
__doc__ = "Modela pisos automaticamente a partir do contorno dos ambientes selecionados, usando o tipo de piso escolhido pelo usuário."

from Autodesk.Revit.DB import *
from pyrevit import forms, revit, script
from System.Collections.Generic import List

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# =============================================================================
# ETAPA 2 — Formulário: selecionar Tipo de Piso
# =============================================================================

# Coletar todos os FloorTypes carregados no projeto
floor_types = FilteredElementCollector(doc)\
    .OfClass(FloorType)\
    .ToElements()

# Filtrar apenas pisos normais — excluir lajes de fundação via propriedade nativa
floor_type_dict = {}
for ft in floor_types:
    try:
        if ft.IsFoundationSlab:
            continue
    except:
        pass
    nome = ft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
    if nome:
        floor_type_dict[nome] = ft

if not floor_type_dict:
    forms.alert("Nenhum tipo de piso encontrado no projeto.", exitscript=True)

#  escolher o tipo de piso
tipo_selecionado = forms.ask_for_one_item(
    sorted(floor_type_dict.keys()),
    default=sorted(floor_type_dict.keys())[0],
    prompt="Selecione o Tipo de Piso:",
    title="Seleção do Tipo do Piso"
)

if not tipo_selecionado:
    script.exit()

floor_type = floor_type_dict[tipo_selecionado]

# =============================================================================
# ETAPA 3 — Seleção múltipla de Ambientes
# =============================================================================

with forms.WarningBar(title="Selecione os Ambientes e pressione Finish"):
    rooms = revit.pick_elements_by_category(BuiltInCategory.OST_Rooms)

if not rooms:
    forms.alert("Nenhum ambiente selecionado.", exitscript=True)

# Nível da vista ativa
nivel_ativo = doc.ActiveView.GenLevel
if nivel_ativo is None:
    forms.alert(
        "A vista ativa não possui nível associado.\nAbra uma planta de piso antes de executar.",
        exitscript=True
    )

# Opções de contorno: face de acabamento das paredes
boundary_opts = SpatialElementBoundaryOptions()
boundary_opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish

# =============================================================================
# ETAPA 4 — Criar Piso a partir do contorno de cada ambiente
# =============================================================================

contador_ok   = 0
contador_erro = 0

with Transaction(doc, "Modelar Pisos Automático") as t:
    t.Start()

    try:
        for room in rooms:
            room_nome = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
            try:
                # Obter os segmentos de contorno do ambiente
                segmentos_lista = room.GetBoundarySegments(boundary_opts)

                if not segmentos_lista or segmentos_lista.Count == 0:
                    print("Ambiente '{}' sem contorno definido — ignorado.".format(room_nome))
                    contador_erro += 1
                    continue

                # Montar um CurveLoop por região (suporte a ambientes com ilhas)
                curve_loops = List[CurveLoop]()

                for segmentos in segmentos_lista:
                    lista_curvas = List[Curve]()
                    z = nivel_ativo.Elevation

                    for seg in segmentos:
                        curva = seg.GetCurve()
                        # Projetar cada curva no Z do nível ativo
                        p0 = curva.GetEndPoint(0)
                        p1 = curva.GetEndPoint(1)
                        linha = Line.CreateBound(
                            XYZ(p0.X, p0.Y, z),
                            XYZ(p1.X, p1.Y, z)
                        )
                        lista_curvas.Add(linha)

                    if lista_curvas.Count < 3:
                        continue

                    curve_loop = CurveLoop.Create(lista_curvas)
                    curve_loops.Add(curve_loop)

                if curve_loops.Count == 0:
                    print("Ambiente '{}' sem loops válidos — ignorado.".format(room_nome))
                    contador_erro += 1
                    continue

                # Criar o piso com Floor.Create (Revit 2022+)
                Floor.Create(doc, curve_loops, floor_type.Id, nivel_ativo.Id)
                contador_ok += 1

            except Exception as e_room:
                contador_erro += 1
                print("Erro em '{}': {}".format(room_nome, str(e_room)))

        t.Commit()

    except Exception as e_geral:
        t.RollBack()
        forms.alert("Erro geral:\n{}".format(str(e_geral)), exitscript=True)

# =============================================================================
# Resumo final
# =============================================================================
forms.alert(
    "Concluído!\n\n✔ Pisos criados: {}\n✘ Com erro: {}".format(contador_ok, contador_erro),
    warn_icon=False
)