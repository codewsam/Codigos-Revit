# -*- coding: utf-8 -*-
"""
TQS Aberturas -> Portas e Janelas
=================================
Regras de classificacao:
  - Suspensa (base > 5cm)                        -> JANELA (sempre, qualquer tamanho)
  - Toca o chao (base <= 5cm) + largura < 100cm  -> PORTA
  - Toca o chao (base <= 5cm) + largura >= 100cm -> mantém como Opening (vao grande)

Autor: Samuel PLUGIN
"""

# ==============================================================================
# CONFIGURACOES
# ==============================================================================
NOME_FAMILIA_PORTA  = "Abertura de porta"
NOME_FAMILIA_JANELA = "Abertura de Janela"

LIMIAR_PORTA_CM  = 5.0    # base <= este valor -> considera que toca o chao
LIMIAR_GRANDE_CM = 100.0  # largura >= este valor + toca chao -> manter como Opening

DELETAR_OPENING = True
ADICIONAR_PARAM = True

# ==============================================================================
# IMPORTS
# ==============================================================================
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, Opening, FamilySymbol,
    Transaction, BuiltInParameter, XYZ, Level,
    Structure
)

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

CM_TO_FT = 1.0 / 30.48
FT_TO_CM = 30.48

# ==============================================================================
# FUNCOES
# ==============================================================================

def get_family_symbol(nome_familia):
    for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
        if getattr(sym.Family, "Name", "") == nome_familia:
            return sym
    return None

def get_nivel_mais_proximo(elev_ft):
    levels = list(FilteredElementCollector(doc).OfClass(Level))
    escolhido = None
    menor_diff = float("inf")
    for lvl in levels:
        diff = elev_ft - lvl.Elevation
        if 0 <= diff < menor_diff:
            menor_diff = diff
            escolhido = lvl
    if escolhido is None:
        for lvl in levels:
            diff = abs(elev_ft - lvl.Elevation)
            if diff < menor_diff:
                menor_diff = diff
                escolhido = lvl
    return escolhido

def get_dados_opening(op):
    bb = op.get_BoundingBox(None)
    if bb is None:
        return None
    largura_ft = abs(bb.Max.X - bb.Min.X)
    altura_ft  = abs(bb.Max.Z - bb.Min.Z)
    base_ft    = bb.Min.Z
    centro     = XYZ(
        (bb.Max.X + bb.Min.X) / 2.0,
        (bb.Max.Y + bb.Min.Y) / 2.0,
        bb.Min.Z
    )
    return largura_ft, altura_ft, base_ft, centro

def ajustar_dimensoes(inst, largura_ft, altura_ft):
    for nome in ["Largura", "Width", "b", "w"]:
        p = inst.LookupParameter(nome)
        if p and not p.IsReadOnly:
            p.Set(largura_ft)
            break
    for nome in ["Altura", "Height", "h"]:
        p = inst.LookupParameter(nome)
        if p and not p.IsReadOnly:
            p.Set(altura_ft)
            break

def adicionar_classificacao(inst, texto):
    p = inst.LookupParameter("Classificacao Automatica")
    if p and not p.IsReadOnly:
        p.Set(texto)

# ==============================================================================
# MAIN
# ==============================================================================

def main():
    sym_porta  = get_family_symbol(NOME_FAMILIA_PORTA)
    sym_janela = get_family_symbol(NOME_FAMILIA_JANELA)

    if sym_porta is None or sym_janela is None:
        if sym_porta  is None: print("ERRO: familia nao encontrada: " + NOME_FAMILIA_PORTA)
        if sym_janela is None: print("ERRO: familia nao encontrada: " + NOME_FAMILIA_JANELA)
        return

    openings = list(FilteredElementCollector(doc).OfClass(Opening))
    if not openings:
        print("Nenhuma Wall Opening encontrada.")
        return

    print("Wall Openings encontradas: {}".format(len(openings)))

    portas  = []
    janelas = []
    grandes = []
    falhas  = []
    deletar = []

    t = Transaction(doc, "TQS: Converter Aberturas em Portas/Janelas")
    t.Start()

    try:
        if not sym_porta.IsActive:  sym_porta.Activate()
        if not sym_janela.IsActive: sym_janela.Activate()

        for op in openings:
            op_id = op.Id.IntegerValue
            dados = get_dados_opening(op)
            if dados is None:
                falhas.append("ID:{} - sem BoundingBox".format(op_id))
                continue

            largura_ft, altura_ft, base_ft, ponto = dados
            largura_cm = largura_ft * FT_TO_CM
            base_cm    = base_ft    * FT_TO_CM

            toca_chao = base_cm <= LIMIAR_PORTA_CM

            # --- REGRA 1: suspensa -> sempre JANELA ---
            if not toca_chao:
                classificacao = "JANELA"
                sym = sym_janela

            # --- REGRA 2: toca chao + grande -> preservar como Opening ---
            elif largura_cm >= LIMIAR_GRANDE_CM:
                grandes.append(
                    "ID:{} | L:{:.0f}cm | A:{:.0f}cm | Base:{:.0f}cm | mantida como Opening (vao grande)".format(
                        op_id, largura_cm, altura_ft * FT_TO_CM, base_cm))
                continue

            # --- REGRA 3: toca chao + normal -> PORTA ---
            else:
                classificacao = "PORTA"
                sym = sym_porta

            host = op.Host
            if host is None:
                falhas.append("ID:{} - sem parede hospedeira".format(op_id))
                continue

            nivel = get_nivel_mais_proximo(base_ft)
            if nivel is None:
                falhas.append("ID:{} - nenhum nivel encontrado".format(op_id))
                continue

            offset_ft = base_ft - nivel.Elevation

            try:
                inst = doc.Create.NewFamilyInstance(
                    ponto, sym, host, nivel,
                    Structure.StructuralType.NonStructural
                )

                for bip in [BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,
                             BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM]:
                    p = inst.get_Parameter(bip)
                    if p and not p.IsReadOnly:
                        p.Set(offset_ft)
                        break

                ajustar_dimensoes(inst, largura_ft, altura_ft)

                if ADICIONAR_PARAM:
                    adicionar_classificacao(inst, classificacao)

                if DELETAR_OPENING:
                    deletar.append(op.Id)

                info = "ID:{} -> {} | L:{:.0f}cm | A:{:.0f}cm | Base:{:.0f}cm".format(
                    op_id, classificacao, largura_cm, altura_ft * FT_TO_CM, base_cm)

                if classificacao == "PORTA":
                    portas.append(info)
                else:
                    janelas.append(info)

            except Exception as ex:
                falhas.append("ID:{} - erro: {}".format(op_id, str(ex)))

        if DELETAR_OPENING:
            for eid in deletar:
                try:
                    doc.Delete(eid)
                except Exception as ex:
                    falhas.append("Falha ao deletar ID:{} - {}".format(eid.IntegerValue, str(ex)))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        print("ERRO CRITICO - transacao revertida: " + str(ex))
        return

    print("\n========== RESULTADO ==========")
    print("Portas criadas    : {}".format(len(portas)))
    for x in portas:   print("  + " + x)
    print("Janelas criadas   : {}".format(len(janelas)))
    for x in janelas:  print("  + " + x)
    print("Grandes (mantidas): {}".format(len(grandes)))
    for x in grandes:  print("  ~ " + x)
    print("Falhas            : {}".format(len(falhas)))
    for x in falhas:   print("  ! " + x)
    print("================================")

main()
