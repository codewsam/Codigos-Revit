# -*- coding: utf-8 -*-
__title__   = "Checklist NBR 6118 (não finalizado)"
__author__  = "Samuel"
__version__ = "Versao 1.0"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form, Label, ComboBox, TextBox, Button, ListBox, GroupBox,
    DialogResult, FormBorderStyle, FormStartPosition,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon,
    RichTextBoxScrollBars, RichTextBox, Panel, CheckBox, ProgressBar,
    Application, SaveFileDialog
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from pyrevit import forms, revit, script

import sys
import os

# ─────────────────────────────────────────────
#  CONSTANTES NBR 6118
# ─────────────────────────────────────────────

# Cobrimento nominal mínimo (mm) por classe de agressividade
# Tabela 7.2 da NBR 6118:2014
COBRIMENTO_MINIMO = {
    "I - Fraca (rural/suburbano seco)": {
        "laje": 20,
        "viga": 25,
        "pilar": 25,
    },
    "II - Moderada (urbano)": {
        "laje": 25,
        "viga": 30,
        "pilar": 30,
    },
    "III - Forte (marinha/industrial)": {
        "laje": 35,
        "viga": 40,
        "pilar": 40,
    },
    "IV - Muito Forte (submerso/respingos)": {
        "laje": 45,
        "viga": 50,
        "pilar": 50,
    },
}

# Dimensões mínimas (mm) - NBR 6118 itens 13.2 e 18.4
DIM_MIN = {
    "viga_largura":  120,   # largura mínima de viga (mm)
    "pilar_menor":   190,   # menor dimensão de pilar (mm)
    "laje_espessura": 70,   # espessura mínima de laje maciça (mm)
}

# Taxa de armadura longitudinal (%) - NBR 6118 item 17.3.5
TAXA_ARMADURA = {
    "pilar_min": 0.4,
    "pilar_max": 8.0,
    "viga_min":  0.15,
    "viga_max":  4.0,
}

STATUS_OK      = "OK"
STATUS_ALERTA  = "ALERTA"
STATUS_FALHA   = "FALHA"
STATUS_ND      = "N/D"   # parâmetro não encontrado no modelo

# ─────────────────────────────────────────────
#  FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────

def get_param_value(element, names):
    """
    Tenta ler um parâmetro pelo nome (aceita lista de nomes alternativos).
    Retorna o valor numérico em mm ou None se não encontrado.
    """
    if isinstance(names, str):
        names = [names]
    for name in names:
        param = element.LookupParameter(name)
        if param and param.HasValue:
            try:
                # Converte de pés (unidade interna Revit) para mm
                return UnitUtils.ConvertFromInternalUnits(
                    param.AsDouble(),
                    UnitTypeId.Millimeters
                )
            except Exception:
                try:
                    return param.AsDouble() * 304.8  # fallback conversão manual
                except Exception:
                    return None
    return None


def get_element_type_name(element):
    """Retorna o nome do tipo do elemento."""
    try:
        tipo = element.Document.GetElement(element.GetTypeId())
        return tipo.Name if tipo else "Sem Tipo"
    except Exception:
        return "Sem Tipo"


def verificar_cobrimento(element, categoria, classe_ag):
    """
    Verifica o cobrimento nominal do elemento.
    Retorna (status, valor_encontrado, valor_minimo).
    """
    minimo = COBRIMENTO_MINIMO[classe_ag].get(categoria, None)
    if minimo is None:
        return STATUS_ND, None, None

    cobrimento = get_param_value(element, [
        "Cobrimento", "Cobrimento Nominal", "Cover",
        "Structural Cover", "Cobertura"
    ])

    if cobrimento is None:
        return STATUS_ND, None, minimo

    if cobrimento >= minimo:
        return STATUS_OK, cobrimento, minimo
    elif cobrimento >= minimo * 0.9:
        return STATUS_ALERTA, cobrimento, minimo
    else:
        return STATUS_FALHA, cobrimento, minimo


def verificar_dimensoes_viga(element):
    """
    Verifica largura mínima de viga.
    Retorna lista de (descricao, status, valor, minimo).
    """
    resultados = []
    largura = get_param_value(element, ["b", "Largura", "Width", "b_w"])
    minimo  = DIM_MIN["viga_largura"]

    if largura is None:
        resultados.append(("Largura", STATUS_ND, None, minimo))
    elif largura >= minimo:
        resultados.append(("Largura", STATUS_OK, largura, minimo))
    elif largura >= minimo * 0.9:
        resultados.append(("Largura", STATUS_ALERTA, largura, minimo))
    else:
        resultados.append(("Largura", STATUS_FALHA, largura, minimo))

    return resultados


def verificar_dimensoes_pilar(element):
    """
    Verifica menor dimensão de pilar.
    Retorna lista de (descricao, status, valor, minimo).
    """
    resultados = []
    b = get_param_value(element, ["b", "Largura", "Width"])
    h = get_param_value(element, ["h", "Altura", "Depth"])

    minimo = DIM_MIN["pilar_menor"]
    menor  = min(v for v in [b, h] if v is not None) if (b or h) else None

    if menor is None:
        resultados.append(("Menor Dim.", STATUS_ND, None, minimo))
    elif menor >= minimo:
        resultados.append(("Menor Dim.", STATUS_OK, menor, minimo))
    elif menor >= minimo * 0.9:
        resultados.append(("Menor Dim.", STATUS_ALERTA, menor, minimo))
    else:
        resultados.append(("Menor Dim.", STATUS_FALHA, menor, minimo))

    return resultados


def verificar_dimensoes_laje(element):
    """
    Verifica espessura mínima de laje.
    Retorna lista de (descricao, status, valor, minimo).
    """
    resultados = []
    espessura = get_param_value(element, [
        "Espessura", "Thickness", "h", "Structural Layer Thickness"
    ])
    minimo = DIM_MIN["laje_espessura"]

    if espessura is None:
        resultados.append(("Espessura", STATUS_ND, None, minimo))
    elif espessura >= minimo:
        resultados.append(("Espessura", STATUS_OK, espessura, minimo))
    elif espessura >= minimo * 0.9:
        resultados.append(("Espessura", STATUS_ALERTA, espessura, minimo))
    else:
        resultados.append(("Espessura", STATUS_FALHA, espessura, minimo))

    return resultados


# ─────────────────────────────────────────────
#  COLETOR DE ELEMENTOS
# ─────────────────────────────────────────────

def coletar_elementos(doc):
    """Coleta vigas, pilares e lajes estruturais do modelo."""
    vigas   = []
    pilares = []
    lajes   = []

    # Vigas e pilares são FramingElements (categoria OST_StructuralFraming e OST_StructuralColumns)
    vigas_col = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_StructuralFraming)\
        .WhereElementIsNotElementType()\
        .ToElements()

    pilares_col = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_StructuralColumns)\
        .WhereElementIsNotElementType()\
        .ToElements()

    lajes_col = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Floors)\
        .WhereElementIsNotElementType()\
        .ToElements()

    for el in vigas_col:
        vigas.append(el)
    for el in pilares_col:
        pilares.append(el)
    for el in lajes_col:
        lajes.append(el)

    return vigas, pilares, lajes


# ─────────────────────────────────────────────
#  MOTOR DE VERIFICAÇÃO
# ─────────────────────────────────────────────

def executar_verificacoes(doc, classe_ag):
    """
    Roda todas as verificações e retorna lista de resultados.
    Cada item: dict com chave, tipo, nome_tipo, verificacao, status, valor, minimo
    """
    resultados = []
    vigas, pilares, lajes = coletar_elementos(doc)

    # ── VIGAS ──────────────────────────────
    for el in vigas:
        nome_tipo = get_element_type_name(el)
        el_id     = el.Id.IntegerValue

        # Cobrimento
        st, val, mn = verificar_cobrimento(el, "viga", classe_ag)
        resultados.append({
            "id": el_id, "tipo": "Viga", "nome_tipo": nome_tipo,
            "verificacao": "Cobrimento Nominal",
            "status": st, "valor": val, "minimo": mn
        })

        # Dimensões
        for desc, st, val, mn in verificar_dimensoes_viga(el):
            resultados.append({
                "id": el_id, "tipo": "Viga", "nome_tipo": nome_tipo,
                "verificacao": desc,
                "status": st, "valor": val, "minimo": mn
            })

    # ── PILARES ────────────────────────────
    for el in pilares:
        nome_tipo = get_element_type_name(el)
        el_id     = el.Id.IntegerValue

        st, val, mn = verificar_cobrimento(el, "pilar", classe_ag)
        resultados.append({
            "id": el_id, "tipo": "Pilar", "nome_tipo": nome_tipo,
            "verificacao": "Cobrimento Nominal",
            "status": st, "valor": val, "minimo": mn
        })

        for desc, st, val, mn in verificar_dimensoes_pilar(el):
            resultados.append({
                "id": el_id, "tipo": "Pilar", "nome_tipo": nome_tipo,
                "verificacao": desc,
                "status": st, "valor": val, "minimo": mn
            })

    # ── LAJES ──────────────────────────────
    for el in lajes:
        nome_tipo = get_element_type_name(el)
        el_id     = el.Id.IntegerValue

        st, val, mn = verificar_cobrimento(el, "laje", classe_ag)
        resultados.append({
            "id": el_id, "tipo": "Laje", "nome_tipo": nome_tipo,
            "verificacao": "Cobrimento Nominal",
            "status": st, "valor": val, "minimo": mn
        })

        for desc, st, val, mn in verificar_dimensoes_laje(el):
            resultados.append({
                "id": el_id, "tipo": "Laje", "nome_tipo": nome_tipo,
                "verificacao": desc,
                "status": st, "valor": val, "minimo": mn
            })

    return resultados


def resumo(resultados):
    ok     = sum(1 for r in resultados if r["status"] == STATUS_OK)
    alerta = sum(1 for r in resultados if r["status"] == STATUS_ALERTA)
    falha  = sum(1 for r in resultados if r["status"] == STATUS_FALHA)
    nd     = sum(1 for r in resultados if r["status"] == STATUS_ND)
    return ok, alerta, falha, nd


# ─────────────────────────────────────────────
#  EXPORTAÇÃO TXT
# ─────────────────────────────────────────────

def fmt_num(valor):
    """Formata numero float para string com 1 casa decimal (compativel IronPython)."""
    if valor is None:
        return "-"
    return str(round(float(valor), 1))


def exportar_relatorio(resultados, classe_ag, caminho):
    linhas = []
    linhas.append("=" * 70)
    linhas.append("  CHECKLIST NBR 6118 - VERIFICACAO DE CONFORMIDADE")
    linhas.append("  Autor: Samuel | Versao 1.0")
    linhas.append("  Classe de Agressividade: {}".format(classe_ag))
    linhas.append("=" * 70)
    linhas.append("")

    ok, alerta, falha, nd = resumo(resultados)
    linhas.append("RESUMO:")
    linhas.append("  OK      : {}".format(ok))
    linhas.append("  ALERTA  : {}".format(alerta))
    linhas.append("  FALHA   : {}".format(falha))
    linhas.append("  N/D     : {} (parametro nao localizado no modelo)".format(nd))
    linhas.append("")
    linhas.append("-" * 70)
    linhas.append("{:<10} {:<12} {:<30} {:<22} {:<8} {:<10} {:<10}".format(
        "ID", "Tipo", "Nome do Tipo", "Verificacao", "Status", "Valor(mm)", "Min(mm)"
    ))
    linhas.append("-" * 70)

    for r in resultados:
        linhas.append("{:<10} {:<12} {:<30} {:<22} {:<8} {:<10} {:<10}".format(
            str(r["id"]),
            r["tipo"],
            r["nome_tipo"][:28],
            r["verificacao"][:20],
            r["status"],
            fmt_num(r["valor"]),
            fmt_num(r["minimo"])
        ))

    linhas.append("=" * 70)

    with open(caminho, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────

class ChecklistForm(Form):

    def __init__(self, doc):
        self.doc = doc
        self.resultados = []
        self._build_ui()

    def _build_ui(self):
        self.Text            = "Checklist NBR 6118 - Concreto"
        self.Size            = Size(860, 620)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        # ── Classe de agressividade ─────────
        lbl_classe = Label()
        lbl_classe.Text     = "Classe de Agressividade Ambiental (NBR 6118 Tab. 6.1):"
        lbl_classe.Location = Point(12, 14)
        lbl_classe.Size     = Size(420, 20)
        self.Controls.Add(lbl_classe)

        self.cmb_classe = ComboBox()
        self.cmb_classe.Location      = Point(12, 36)
        self.cmb_classe.Size          = Size(500, 24)
        self.cmb_classe.DropDownStyle = ComboBoxStyle.DropDownList
        for classe in COBRIMENTO_MINIMO.keys():
            self.cmb_classe.Items.Add(classe)
        self.cmb_classe.SelectedIndex = 1   # default: Classe II
        self.Controls.Add(self.cmb_classe)

        # ── Botão rodar ─────────────────────
        self.btn_rodar = Button()
        self.btn_rodar.Text     = "Executar Verificacao"
        self.btn_rodar.Location = Point(530, 34)
        self.btn_rodar.Size     = Size(150, 28)
        self.btn_rodar.Click   += self._on_rodar
        self.Controls.Add(self.btn_rodar)

        # ── Botão exportar ──────────────────
        self.btn_exportar = Button()
        self.btn_exportar.Text     = "Exportar Relatorio"
        self.btn_exportar.Location = Point(695, 34)
        self.btn_exportar.Size     = Size(140, 28)
        self.btn_exportar.Enabled  = False
        self.btn_exportar.Click   += self._on_exportar
        self.Controls.Add(self.btn_exportar)

        # ── Resumo ──────────────────────────
        self.lbl_resumo = Label()
        self.lbl_resumo.Text     = "Aguardando execucao..."
        self.lbl_resumo.Location = Point(12, 72)
        self.lbl_resumo.Size     = Size(820, 20)
        self.lbl_resumo.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(self.lbl_resumo)

        # ── Resultado em texto ───────────────
        self.txt_resultado = RichTextBox()
        self.txt_resultado.Location   = Point(12, 98)
        self.txt_resultado.Size       = Size(820, 460)
        self.txt_resultado.ReadOnly   = True
        self.txt_resultado.ScrollBars = RichTextBoxScrollBars.Both
        self.txt_resultado.Font       = Font("Courier New", 8)
        self.txt_resultado.BackColor  = Color.FromArgb(30, 30, 30)
        self.txt_resultado.ForeColor  = Color.White
        self.txt_resultado.WordWrap   = False
        self.Controls.Add(self.txt_resultado)

    # ── EVENTOS ─────────────────────────────────────────────────────────

    def _on_rodar(self, sender, args):
        classe_ag = self.cmb_classe.SelectedItem
        if not classe_ag:
            MessageBox.Show(
                "Selecione a classe de agressividade antes de continuar.",
                "Atencao",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        self.btn_rodar.Enabled    = False
        self.btn_exportar.Enabled = False
        self.txt_resultado.Clear()
        self.lbl_resumo.Text = "Executando verificacoes..."
        Application.DoEvents()

        try:
            self.resultados = executar_verificacoes(self.doc, classe_ag)
            self._preencher_resultado(classe_ag)
            self.btn_exportar.Enabled = True
        except Exception as ex:
            MessageBox.Show(
                "Erro durante a verificacao:\n{}".format(str(ex)),
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
        finally:
            self.btn_rodar.Enabled = True

    def _preencher_resultado(self, classe_ag):
        ok, alerta, falha, nd = resumo(self.resultados)
        total = len(self.resultados)

        self.lbl_resumo.Text = (
            "Total: {}  |  OK: {}  |  ALERTA: {}  |  FALHA: {}  |  N/D: {}"
            .format(total, ok, alerta, falha, nd)
        )

        linhas = []
        linhas.append("=" * 95)
        linhas.append("  CHECKLIST NBR 6118 | Classe: {}".format(classe_ag))
        linhas.append("=" * 95)
        linhas.append("{:<10} {:<10} {:<30} {:<22} {:<8} {:<10} {:<8}".format(
            "ID", "Tipo", "Nome do Tipo", "Verificacao", "Status", "Val(mm)", "Min(mm)"
        ))
        linhas.append("-" * 95)

        for r in self.resultados:
            linha = "{:<10} {:<10} {:<30} {:<22} {:<8} {:<10} {:<8}".format(
                str(r["id"]),
                r["tipo"],
                r["nome_tipo"][:28],
                r["verificacao"][:20],
                r["status"],
                fmt_num(r["valor"]),
                fmt_num(r["minimo"])
            )
            linhas.append(linha)

        linhas.append("=" * 95)
        self.txt_resultado.Text = "\n".join(linhas)

        # Colorização simples: pinta linhas de falha em vermelho
        # (RichTextBox IronPython: colorir por substring)
        self._colorir_linhas()

    def _colorir_linhas(self):
        """Colore linhas do RichTextBox de acordo com status."""
        rtb   = self.txt_resultado
        texto = rtb.Text
        linhas = texto.split("\n")
        rtb.Clear()

        for linha in linhas:
            start = len(rtb.Text)
            rtb.AppendText(linha + "\n")
            end   = len(rtb.Text)

            if STATUS_FALHA in linha:
                rtb.Select(start, end - start)
                rtb.SelectionColor = Color.FromArgb(255, 80, 80)
            elif STATUS_ALERTA in linha:
                rtb.Select(start, end - start)
                rtb.SelectionColor = Color.FromArgb(255, 200, 50)
            elif STATUS_OK in linha:
                rtb.Select(start, end - start)
                rtb.SelectionColor = Color.FromArgb(100, 220, 100)
            elif STATUS_ND in linha:
                rtb.Select(start, end - start)
                rtb.SelectionColor = Color.FromArgb(160, 160, 160)
            else:
                rtb.Select(start, end - start)
                rtb.SelectionColor = Color.White

        rtb.SelectionStart = 0

    def _on_exportar(self, sender, args):
        dlg = SaveFileDialog()
        dlg.Title      = "Salvar Relatorio NBR 6118"
        dlg.Filter     = "Arquivo de texto (*.txt)|*.txt"
        dlg.FileName   = "checklist_nbr6118.txt"

        if dlg.ShowDialog() == DialogResult.OK:
            try:
                classe_ag = self.cmb_classe.SelectedItem
                exportar_relatorio(self.resultados, classe_ag, dlg.FileName)
                MessageBox.Show(
                    "Relatorio exportado com sucesso!\n{}".format(dlg.FileName),
                    "Sucesso",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
            except Exception as ex:
                MessageBox.Show(
                    "Erro ao exportar:\n{}".format(str(ex)),
                    "Erro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                )


# ─────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    doc  = revit.doc
    form = ChecklistForm(doc)
    form.ShowDialog()
