# -*- coding: utf-8 -*-
__title__ = "Colocar Ferro na Parede"
__author__ = "Samuel"
__version__ = "Versão 2.0"
__doc__ = """
_____________________________________________________________________
Descrição:

Selecione paredes e insira famílias estruturais automaticamente nelas.

_____________________________________________________________________
Passo a passo:

1. Selecione a família estrutural desejada
2. Selecione as paredes
3. O plugin insere automaticamente

_____________________________________________________________________
Última atualização:
- [Versão 2.0] - SIMPLIFICADA

"""
# ___  __  __  ____    ___   ____   _____  ____  
#|_ _||  \/  ||  _ \  / _ \ |  _ \ |_   _|/ ___| 
# | | | |\/| || |_) || | | || |_) |  | |  \___ \ 
# | | | |  | ||  __/ | |_| ||  _ <   | |   ___) |
#|___||_|  |_||_|     \___/ |_| \_\  |_|  |____/ 
#=================================================

# Importações
import clr
import os
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, revit, script
from System.Collections.Generic import List

# Variáveis globais do Revit
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# _____  _   _  _   _   ____  _____  ___   ___   _   _  ____  
#|  ___|| | | || \ | | / ___||_   _||_ _| / _ \ | \ | |/ ___| 
#| |_   | | | ||  \| || |      | |   | | | | | ||  \| |\___ \ 
#|  _|  | |_| || |\  || |___   | |   | | | |_| || |\  | ___) |
#|_|     \___/ |_| \_| \____|  |_|  |___| \___/ |_| \_||____| 

output = script.get_output()
output.print_md("## Iniciar plugin - Colocar Ferro na Parede...")

try:
    # -------------------------------------------------------------------------
    # ETAPA 1 — Coletar famílias estruturais
    # -------------------------------------------------------------------------
    output.print_md("### Coletando famílias estruturais...")
    
    family_symbols = {}
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    
    for fs in collector:
        try:
            if fs.Category:
                cat_id = fs.Category.Id.IntegerValue
                if cat_id in [int(BuiltInCategory.OST_StructuralFraming), 
                              int(BuiltInCategory.OST_StructuralColumns)]:
                    nome = fs.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    if nome:
                        family_symbols[nome] = fs
        except:
            pass
    
    if not family_symbols:
        forms.alert(
            "Nenhuma família estrutural encontrada no projeto.\n\n"
            "Carregue famílias de vigas ou pilares antes de usar.",
            exitscript=True
        )
    
    # Ordenar nomes
    nomes_ordenados = sorted(family_symbols.keys())
    
    # Formulário para selecionar
    familia_escolhida = forms.ask_for_one_item(
        nomes_ordenados,
        default=nomes_ordenados[0],
        prompt="Selecione a Família Estrutural:",
        title="Seleção da Família"
    )
    
    if not familia_escolhida:
        script.exit()
    
    family_symbol = family_symbols[familia_escolhida]
    output.print_md("**Família selecionada:** {}".format(familia_escolhida))
    
    # -------------------------------------------------------------------------
    # ETAPA 2 — Selecionar paredes
    # -------------------------------------------------------------------------
    output.print_md("### Selecione as paredes...")
    
    with forms.WarningBar(title="Selecione as paredes"):
        walls = revit.pick_elements_by_category(BuiltInCategory.OST_Walls)
    
    if not walls:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)
    
    output.print_md("**Paredes selecionadas:** {}".format(len(walls)))
    
    # -------------------------------------------------------------------------
    # ETAPA 3 — Inserir ferros
    # -------------------------------------------------------------------------
    output.print_md("### Inserindo ferros...")
    
    contador_ok = 0
    contador_erro = 0
    
    # Ativar família
    if not family_symbol.IsActive:
        with Transaction(doc, "Ativar Família") as t:
            t.Start()
            family_symbol.Activate()
            t.Commit()
    
    with Transaction(doc, "Inserir Ferro") as t:
        t.Start()
        
        for wall in walls:
            try:
                # Obter curva de localização
                curva = wall.Location.Curve
                if curva is None:
                    contador_erro += 1
                    continue
                
                # Calcular ponto central
                p1 = curva.GetEndPoint(0)
                p2 = curva.GetEndPoint(1)
                ponto = p1.Add(p2).Multiply(0.5)
                
                # Inserir família
                doc.Create.NewFamilyInstance(
                    ponto,
                    family_symbol,
                    wall,
                    StructuralType.Column
                )
                
                contador_ok += 1
                output.print_md("✔ Ferro inserido na parede '{}'".format(wall.Name))
                
            except Exception as e:
                contador_erro += 1
                output.print_md("✘ Erro: {}".format(str(e)))
        
        t.Commit()
    
    # -------------------------------------------------------------------------
    # RESULTADO
    # -------------------------------------------------------------------------
    output.print_md("## Concluído!")
    output.print_md("**Inseridos:** {}".format(contador_ok))
    output.print_md("**Erros:** {}".format(contador_erro))
    
    forms.alert(
        "Concluído!\n\n✔ Inseridos: {}\n✘ Erros: {}".format(contador_ok, contador_erro),
        warn_icon=False
    )


except Exception as e:
    output.print_md("### ERRO: {}".format(str(e)))
    forms.alert("Erro:\n{}".format(str(e)), exitscript=True)

output.print_md("## Fim!")
