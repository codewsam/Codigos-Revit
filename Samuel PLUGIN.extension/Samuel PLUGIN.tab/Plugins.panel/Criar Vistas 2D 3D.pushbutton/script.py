# -*- coding: utf-8 -*-
__title__ = "Criar Vistas\n2D e 3D"
__author__ = "Fellipe Caetano - BIM Coder"
__version__ = "Versão 1.0"
__doc__ = """
_____________________________________________________________________
Descrição:

Cria vistas 2D (planta baixa) e 3D de forma automática a partir da seleção dos ambientes, aplicando como cropbox (planta baixa) e Sectionbox (vista 3D) de acordo com o contorno do ambiente com um offset de 15cm

_____________________________________________________________________
Última atualização:
- [25.03.2025] - VERSÃO 1.0

"""
# ___  __  __  ____    ___   ____   _____  ____  
#|_ _||  \/  ||  _ \  / _ \ |  _ \ |_   _|/ ___| 
# | | | |\/| || |_) || | | || |_) |  | |  \___ \ 
# | | | |  | ||  __/ | |_| ||  _ <   | |   ___) |
#|___||_|  |_||_|     \___/ |_| \_\  |_|  |____/ 
#=================================================

# Importações Python e .NET
import clr
import clr
import os, traceback,math,re
clr.AddReference("System")

# Importações Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Structure import StructuralType

# Importações pyRevit
from pyrevit import forms, script, revit

# Variáveis globais do Revit
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application
rvt_year = int(app.VersionNumber)
PATH_SCRIPT = os.path.dirname(__file__)

# -*- coding: utf-8 -*-

# ============================================================
# IMPORTS NECESSÁRIOS
# ============================================================

# pyRevit
from pyrevit import revit

# Revit API
from Autodesk.Revit.DB import *

# Python
import sys


# ============================================================
# ETAPA 1 — INÍCIO
# Script: Criar vistas 2D (planta) e 3D automaticamente
# ============================================================

doc = revit.doc
uidoc = revit.uidoc


try:

    # ============================================================
    # ETAPA 2 — SELEÇÃO MÚLTIPLA COM FILTRO (AMBIENTES)
    # ============================================================

    ambientes = revit.pick_elements_by_category(
        BuiltInCategory.OST_Rooms,
        message='Selecione múltiplos ambientes'
    )

    if not ambientes:
        raise Exception("Nenhum ambiente foi selecionado")


    # ============================================================
    # OBTER NÍVEL DA VISTA ATUAL
    # ============================================================

    vista_atual = doc.ActiveView

    if not vista_atual.GenLevel:
        raise Exception("A vista atual não possui nível associado")

    nivel = doc.GetElement(vista_atual.GenLevel.Id)


    # ============================================================
    # BUSCAR TIPOS DE VISTA
    # ============================================================

    vft_planta = None
    vft_3d = None

    tipos = FilteredElementCollector(doc).OfClass(ViewFamilyType)

    for vft in tipos:

        if vft.ViewFamily == ViewFamily.FloorPlan and not vft_planta:
            vft_planta = vft

        if vft.ViewFamily == ViewFamily.ThreeDimensional and not vft_3d:
            vft_3d = vft


    if not vft_planta:
        raise Exception("Tipo de vista de planta não encontrado")

    if not vft_3d:
        raise Exception("Tipo de vista 3D não encontrado")


    # ============================================================
    # TRANSACTION (MODIFICAÇÃO DO MODELO)
    # ============================================================

    with Transaction(doc, "Criar vistas automáticas") as t:

        t.Start()

        # ============================================================
        # LOOP EM CADA AMBIENTE
        # ============================================================

        for element in ambientes:

            # ============================================================
            # ETAPA 3 — OBTER BOUNDINGBOX DO ELEMENTO
            # ============================================================

            bbox = element.get_BoundingBox(None)

            if not bbox:
                continue


            # ============================================================
            # ETAPA 4 — APLICAR OFFSET DE 15cm
            # ============================================================

            offset = 0.15 / 0.3048  # conversão metros → pés

            min_pt = XYZ(
                bbox.Min.X - offset,
                bbox.Min.Y - offset,
                bbox.Min.Z - offset
            )

            max_pt = XYZ(
                bbox.Max.X + offset,
                bbox.Max.Y + offset,
                bbox.Max.Z + offset
            )

            bbox_final = BoundingBoxXYZ()
            bbox_final.Min = min_pt
            bbox_final.Max = max_pt


            # ============================================================
            # OBTER NOME DO AMBIENTE
            # ============================================================

            nome_ambiente = element.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
            
            # ============================================================
            # ETAPA 5 — CRIAR VISTA DE PLANTA
            # ============================================================

            vista_planta = ViewPlan.Create(doc, vft_planta.Id, nivel.Id)

            vista_planta.Name = "Planta - " + nome_ambiente

            vista_planta.CropBoxActive = True
            vista_planta.CropBoxVisible = True
            vista_planta.CropBox = bbox_final


            # ============================================================
            # ETAPA 6 — CRIAR VISTA 3D
            # ============================================================

            vista_3d = View3D.CreateIsometric(doc, vft_3d.Id)

            vista_3d.Name = "3D - " + nome_ambiente

            vista_3d.SetSectionBox(bbox_final)


        t.Commit()


    print("Vistas criadas com sucesso!")


except Exception as erro:

    print("Erro ao executar o script:")
    print(str(erro))

      

























