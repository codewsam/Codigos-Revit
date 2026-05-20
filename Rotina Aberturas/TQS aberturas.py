# -*- coding: utf-8 -*-
"""
TQS Aberturas → Portas e Janelas
=================================
Converte Wall Openings importadas do TQS em famílias de Porta ou Janela,
classificando automaticamente pela elevação da base da abertura.

Regras de classificação:
  - Base <= 5 cm  → PORTA  (usa família "Abertura de porta")
  - Base >= 70 cm → JANELA (usa família "Abertura de Janela")
  - Caso contrário → ignorado (zona ambígua, reportado no log)

Largura mínima filtrada: 20 cm (remove resquícios do TQS com 10 cm)

Autor: Samuel PLUGIN
"""

# ==============================================================================
# CONFIGURAÇÕES — ajuste aqui se necessário
# ==============================================================================
NOME_FAMILIA_PORTA  = "Abertura de porta"   # Nome exato da família no modelo
NOME_FAMILIA_JANELA = "Abertura de Janela"  # Nome exato da família no modelo

LIMIAR_PORTA_CM   = 5.0    # base <= este valor → PORTA
LIMIAR_JANELA_CM  = 70.0   # base >= este valor → JANELA
LARGURA_MIN_CM    = 20.0   # openings menores que isso são ignoradas (resíduos TQS)

DELETAR_OPENING   = True   # True = deleta a Wall Opening após inserir a família
ADICIONAR_PARAM   = True   # True = adiciona parâmetro "Classificação Automática"

# ==============================================================================
# IMPORTS
# ==============================================================================
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, Opening, FamilySymbol, FamilyInstance,
    Transaction, BuiltInParameter, XYZ, Level,
    BuiltInCategory, Structure
)
from Autodesk.Revit.DB import Line as RvtLine

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ==============================================================================
# CONSTANTES DE CONVERSÃO
# ==============================================================================
FT_TO_CM = 30.48
CM_TO_FT = 1.0 / 30.48

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def cm_to_ft(cm):
    return cm * CM_TO_FT

def ft_to_cm(ft):
    return ft * FT_TO_CM

def get_family_symbol(nome_familia):
    """Retorna o primeiro FamilySymbol ativo da família com o nome dado."""
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for sym in collector:
        fam_name = getattr(sym.Family, "Name", "")
        if fam_name == nome_familia:
            return sym
    return None

def get_nivel_mais_proximo(elev_ft):
    """Retorna o Level mais próximo (por baixo) de uma elevação em feet."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    nivel_escolhido = None
    menor_diff = float("inf")
    for lvl in levels:
        diff = elev_ft - lvl.Elevation
        if 0 <= diff < menor_diff:
            menor_diff = diff
            nivel_escolhido = lvl
    # Se nenhum nível ficou abaixo, pega o mais próximo em absoluto
    if nivel_escolhido is None:
        for lvl in levels:
            diff = abs(elev_ft - lvl.Elevation)
            if diff < menor_diff:
                menor_diff = diff
                nivel_escolhido = lvl
    return nivel_escolhido

def get_dimensao_opening(op):
    """
    Retorna (largura_ft, altura_ft, base_elev_ft, centro_XYZ) de uma Wall Opening.
    Usa BoundingBox como fallback confiável.
    """
    bb = op.get_BoundingBox(None)
    if bb is None:
        return None
    largura_ft  = abs(bb.Max.X - bb.Min.X)
    altura_ft   = abs(bb.Max.Z - bb.Min.Z)
    base_ft     = bb.Min.Z
    centro_x    = (bb.Max.X + bb.Min.X) / 2.0
    centro_y    = (bb.Max.Y + bb.Min.Y) / 2.0
    centro_z    = bb.Min.Z  # ponto de inserção na base
    return (largura_ft, altura_ft, base_ft, XYZ(centro_x, centro_y, centro_z))

def ajustar_dimensoes(inst, largura_ft, altura_ft):
    """
    Tenta ajustar Largura e Altura da instância de família pelos parâmetros padrão.
    Suporta nomes em PT-BR e EN.
    """
    nomes_largura = ["Largura", "Width", "b", "w"]
    nomes_altura  = ["Altura", "Height", "h", "Unconnected Height"]

    def set_param(inst, nomes, valor_ft):
        for nome in nomes:
            p = inst.LookupParameter(nome)
            if p and not p.IsReadOnly:
                p.Set(valor_ft)
                return True
        return False

    ok_l = set_param(inst, nomes_largura, largura_ft)
    ok_h = set_param(inst, nomes_altura,  altura_ft)
    return ok_l, ok_h

def adicionar_classificacao(inst, texto):
    """Adiciona ou atualiza o parâmetro 'Classificação Automática' na instância."""
    p = inst.LookupParameter("Classificacao Automatica")
    if p and not p.IsReadOnly:
        p.Set(texto)

# ==============================================================================
# LÓGICA PRINCIPAL
# ==============================================================================

def main():
    # --- Buscar famílias ---
    sym_porta  = get_family_symbol(NOME_FAMILIA_PORTA)
    sym_janela = get_family_symbol(NOME_FAMILIA_JANELA)

    erros = []
    if sym_porta is None:
        erros.append("Família '{}' NÃO encontrada no modelo.".format(NOME_FAMILIA_PORTA))
    if sym_janela is None:
        erros.append("Família '{}' NÃO encontrada no modelo.".format(NOME_FAMILIA_JANELA))
    if erros:
        for e in erros:
            print("ERRO: " + e)
        print("\nCarregue as famílias no modelo antes de executar.")
        return

    # --- Coletar Wall Openings ---
    openings = list(FilteredElementCollector(doc).OfClass(Opening))
    if not openings:
        print("Nenhuma Wall Opening encontrada no modelo.")
        return

    print("Wall Openings encontradas: {}".format(len(openings)))
    print("-" * 60)

    portas_criadas  = []
    janelas_criadas = []
    ignoradas       = []
    falhas          = []
    ids_deletar     = []

    t = Transaction(doc, "TQS: Converter Aberturas em Portas/Janelas")
    t.Start()

    try:
        # Ativar símbolos se necessário
        if not sym_porta.IsActive:
            sym_porta.Activate()
        if not sym_janela.IsActive:
            sym_janela.Activate()

        for op in openings:
            op_id = op.Id.IntegerValue
            dados = get_dimensao_opening(op)

            if dados is None:
                falhas.append("ID:{} - sem BoundingBox".format(op_id))
                continue

            largura_ft, altura_ft, base_ft, ponto_insercao = dados
            largura_cm = ft_to_cm(largura_ft)
            altura_cm  = ft_to_cm(altura_ft)
            base_cm    = ft_to_cm(base_ft)

            # Filtrar resíduos pequenos do TQS
            if largura_cm < LARGURA_MIN_CM:
                ignoradas.append("ID:{} - largura {:.1f}cm (resíduo TQS)".format(op_id, largura_cm))
                continue

            # Classificar
            if base_cm <= LIMIAR_PORTA_CM:
                classificacao = "PORTA"
                sym = sym_porta
            elif base_cm >= LIMIAR_JANELA_CM:
                classificacao = "JANELA"
                sym = sym_janela
            else:
                ignoradas.append("ID:{} - base {:.1f}cm (zona ambígua entre {}cm e {}cm)".format(
                    op_id, base_cm, LIMIAR_PORTA_CM, LIMIAR_JANELA_CM))
                continue

            # Pegar parede hospedeira
            host = op.Host
            if host is None:
                falhas.append("ID:{} - sem parede hospedeira".format(op_id))
                continue

            # Pegar nível mais próximo
            nivel = get_nivel_mais_proximo(base_ft)
            if nivel is None:
                falhas.append("ID:{} - nenhum nível encontrado".format(op_id))
                continue

            # Calcular offset em relação ao nível
            offset_ft = base_ft - nivel.Elevation

            try:
                # Inserir instância na parede
                inst = doc.Create.NewFamilyInstance(
                    ponto_insercao,
                    sym,
                    host,
                    nivel,
                    Structure.StructuralType.NonStructural
                )

                # Ajustar offset do nível (elevação base)
                p_offset = inst.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
                if p_offset and not p_offset.IsReadOnly:
                    p_offset.Set(offset_ft)

                # Ajustar dimensões
                ok_l, ok_h = ajustar_dimensoes(inst, largura_ft, altura_ft)

                # Parâmetro de classificação
                if ADICIONAR_PARAM:
                    adicionar_classificacao(inst, classificacao)

                # Registrar para deleção posterior
                if DELETAR_OPENING:
                    ids_deletar.append(op.Id)

                info = "ID:{} → {} | L:{:.1f}cm A:{:.1f}cm Base:{:.1f}cm | Nível:{}".format(
                    op_id, classificacao, largura_cm, altura_cm, base_cm,
                    getattr(nivel, "Name", "?"))
                if classificacao == "PORTA":
                    portas_criadas.append(info)
                else:
                    janelas_criadas.append(info)

            except Exception as ex:
                falhas.append("ID:{} - erro ao inserir: {}".format(op_id, str(ex)))

        # Deletar openings originais
        if DELETAR_OPENING and ids_deletar:
            for eid in ids_deletar:
                try:
                    doc.Delete(eid)
                except Exception as ex:
                    falhas.append("Falha ao deletar Opening ID:{} - {}".format(
                        eid.IntegerValue, str(ex)))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        print("ERRO CRÍTICO — transação revertida: " + str(ex))
        return

    # --- Relatório final ---
    print("\n========== RELATÓRIO FINAL ==========")
    print("Portas criadas   : {}".format(len(portas_criadas)))
    for x in portas_criadas: print("  ✔ " + x)

    print("\nJanelas criadas  : {}".format(len(janelas_criadas)))
    for x in janelas_criadas: print("  ✔ " + x)

    print("\nIgnoradas        : {}".format(len(ignoradas)))
    for x in ignoradas: print("  — " + x)

    print("\nFalhas           : {}".format(len(falhas)))
    for x in falhas: print("  ✖ " + x)

    print("\nOpenings deletadas: {}".format(len(ids_deletar) if DELETAR_OPENING else 0))
    print("=====================================")

main()
