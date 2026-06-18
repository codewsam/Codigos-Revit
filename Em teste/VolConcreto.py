# -*- coding: utf-8 -*-
__title__   = "Volume Concreto(n finalizado)"
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
    Form, Label, ComboBox, Button, GroupBox,
    DialogResult, FormBorderStyle, FormStartPosition,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon,
    RichTextBoxScrollBars, RichTextBox, CheckBox,
    Application, SaveFileDialog
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from pyrevit import forms, revit, script


# ─────────────────────────────────────────────
#  CONSTANTES
# ─────────────────────────────────────────────

CATEGORIAS = {
    "Pilares":  BuiltInCategory.OST_StructuralColumns,
    "Vigas":    BuiltInCategory.OST_StructuralFraming,
    "Lajes":    BuiltInCategory.OST_Floors,
    "Fundacao": BuiltInCategory.OST_StructuralFoundation,
    "Paredes":  BuiltInCategory.OST_Walls,
}


# ─────────────────────────────────────────────
#  FUNÇÕES AUXILIARES
# ─────────────────────────────────────────────

def pes_cubicos_para_m3(valor_interno):
    """Converte pés³ (unidade interna Revit) para m³."""
    try:
        return UnitUtils.ConvertFromInternalUnits(valor_interno, UnitTypeId.CubicMeters)
    except Exception:
        return valor_interno * 0.0283168  # fallback manual

def fmt_vol(valor):
    """Formata volume m³ para string com 3 casas (compatível IronPython)."""
    if valor is None:
        return "-"
    return str(round(float(valor), 3))

def get_fase(element, doc):
    """Retorna o nome da fase de criação do elemento."""
    try:
        fase_id = element.CreatedPhaseId
        if fase_id and fase_id != ElementId.InvalidElementId:
            fase = doc.GetElement(fase_id)
            return fase.Name if fase else "Sem Fase"
    except Exception:
        pass
    return "Sem Fase"

def get_nivel(element, doc):
    """Tenta obter o nível associado ao elemento."""
    try:
        nivel_id = element.LevelId
        if nivel_id and nivel_id != ElementId.InvalidElementId:
            nivel = doc.GetElement(nivel_id)
            return nivel.Name if nivel else "-"
    except Exception:
        pass
    return "-"

def get_volume(element):
    """Lê o parâmetro de volume do elemento e retorna em m³."""
    nomes = ["Volume", "Structural Material Volume"]
    for nome in nomes:
        p = element.LookupParameter(nome)
        if p and p.HasValue:
            try:
                return pes_cubicos_para_m3(p.AsDouble())
            except Exception:
                pass
    # Fallback: tenta parâmetro built-in
    try:
        p = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
        if p and p.HasValue:
            return pes_cubicos_para_m3(p.AsDouble())
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────
#  COLETA E CÁLCULO
# ─────────────────────────────────────────────

def coletar_volumes(doc):
    """
    Coleta todos os elementos das categorias definidas.
    Retorna lista de dicts: categoria, tipo, nivel, fase, volume_m3
    """
    registros = []

    for cat_nome, cat_bic in CATEGORIAS.items():
        elementos = FilteredElementCollector(doc)\
            .OfCategory(cat_bic)\
            .WhereElementIsNotElementType()\
            .ToElements()

        for el in elementos:
            vol = get_volume(el)
            if vol is None or vol <= 0:
                continue

            try:
                tipo_el = doc.GetElement(el.GetTypeId())
                nome_tipo = tipo_el.Name if tipo_el else "Sem Tipo"
            except Exception:
                nome_tipo = "Sem Tipo"

            registros.append({
                "categoria": cat_nome,
                "nome_tipo": nome_tipo,
                "nivel":     get_nivel(el, doc),
                "fase":      get_fase(el, doc),
                "volume":    vol,
                "id":        el.Id.IntegerValue,
            })

    return registros


def agrupar_por(registros, chave):
    """Agrupa registros por uma chave e soma volumes."""
    grupos = {}
    for r in registros:
        k = r[chave]
        if k not in grupos:
            grupos[k] = 0.0
        grupos[k] += r["volume"]
    return sorted(grupos.items(), key=lambda x: x[0])


def agrupar_categoria_fase(registros):
    """Agrupa por categoria + fase para tabela cruzada."""
    grupos = {}
    for r in registros:
        k = (r["categoria"], r["fase"])
        if k not in grupos:
            grupos[k] = 0.0
        grupos[k] += r["volume"]
    return grupos


# ─────────────────────────────────────────────
#  GERAÇÃO DO RELATÓRIO
# ─────────────────────────────────────────────

def montar_relatorio(registros):
    linhas = []

    total_geral = sum(r["volume"] for r in registros)

    linhas.append("=" * 80)
    linhas.append("  VOLUME DE CONCRETO - RELATORIO GERAL")
    linhas.append("  Autor: Samuel | Versao 1.0")
    linhas.append("  Total de elementos: {}".format(len(registros)))
    linhas.append("  Volume Total: {} m3".format(fmt_vol(total_geral)))
    linhas.append("=" * 80)
    linhas.append("")

    # ── Por Categoria ──────────────────────
    linhas.append("[ POR CATEGORIA ]")
    linhas.append("-" * 40)
    for cat, vol in agrupar_por(registros, "categoria"):
        pct = round(float(vol) / float(total_geral) * 100, 1) if total_geral else 0
        linhas.append("  {:<20} {:>10} m3  ({} %)".format(cat, fmt_vol(vol), pct))
    linhas.append("")

    # ── Por Fase (Etapa de Concretagem) ───
    linhas.append("[ POR FASE / ETAPA DE CONCRETAGEM ]")
    linhas.append("-" * 40)
    for fase, vol in agrupar_por(registros, "fase"):
        linhas.append("  {:<30} {:>10} m3".format(fase, fmt_vol(vol)))
    linhas.append("")

    # ── Por Nível ─────────────────────────
    linhas.append("[ POR NIVEL ]")
    linhas.append("-" * 40)
    for nivel, vol in agrupar_por(registros, "nivel"):
        linhas.append("  {:<30} {:>10} m3".format(nivel, fmt_vol(vol)))
    linhas.append("")

    # ── Tabela Categoria x Fase ───────────
    linhas.append("[ CATEGORIA x FASE ]")
    linhas.append("-" * 80)

    fases = sorted(set(r["fase"] for r in registros))
    cats  = sorted(set(r["categoria"] for r in registros))
    grp   = agrupar_categoria_fase(registros)

    # Cabeçalho
    cab = "{:<16}".format("Categoria")
    for f in fases:
        cab += " {:>14}".format(f[:13])
    cab += " {:>14}".format("TOTAL")
    linhas.append(cab)
    linhas.append("-" * 80)

    for cat in cats:
        linha = "{:<16}".format(cat)
        total_cat = 0.0
        for f in fases:
            v = grp.get((cat, f), 0.0)
            total_cat += v
            linha += " {:>14}".format(fmt_vol(v) if v > 0 else "-")
        linha += " {:>14}".format(fmt_vol(total_cat))
        linhas.append(linha)

    linhas.append("-" * 80)
    # Linha de totais por fase
    linha_tot = "{:<16}".format("TOTAL")
    for f in fases:
        v = sum(grp.get((cat, f), 0.0) for cat in cats)
        linha_tot += " {:>14}".format(fmt_vol(v))
    linha_tot += " {:>14}".format(fmt_vol(total_geral))
    linhas.append(linha_tot)
    linhas.append("=" * 80)

    # ── Detalhe por elemento ───────────────
    linhas.append("")
    linhas.append("[ DETALHE POR ELEMENTO ]")
    linhas.append("-" * 80)
    linhas.append("{:<10} {:<14} {:<28} {:<20} {:<12} {:<10}".format(
        "ID", "Categoria", "Tipo", "Nivel", "Fase", "Vol(m3)"
    ))
    linhas.append("-" * 80)
    for r in sorted(registros, key=lambda x: (x["categoria"], x["fase"], x["nivel"])):
        linhas.append("{:<10} {:<14} {:<28} {:<20} {:<12} {:<10}".format(
            str(r["id"]),
            r["categoria"][:13],
            r["nome_tipo"][:27],
            r["nivel"][:19],
            r["fase"][:11],
            fmt_vol(r["volume"])
        ))
    linhas.append("=" * 80)

    return "\n".join(linhas)


def exportar_txt(relatorio, caminho):
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(relatorio)


# ─────────────────────────────────────────────
#  INTERFACE GRÁFICA
# ─────────────────────────────────────────────

class VolumeForm(Form):

    def __init__(self, doc):
        self.doc        = doc
        self.registros  = []
        self.relatorio  = ""
        self._build_ui()

    def _build_ui(self):
        self.Text            = "Volume de Concreto"
        self.Size            = Size(860, 620)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        # ── Título ──────────────────────────
        lbl_titulo = Label()
        lbl_titulo.Text     = "Calculo de Volume de Concreto por Elemento e por Etapa"
        lbl_titulo.Location = Point(12, 14)
        lbl_titulo.Size     = Size(600, 20)
        lbl_titulo.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(lbl_titulo)

        # ── Botão calcular ───────────────────
        self.btn_calcular = Button()
        self.btn_calcular.Text     = "Calcular Volumes"
        self.btn_calcular.Location = Point(530, 34)
        self.btn_calcular.Size     = Size(150, 28)
        self.btn_calcular.Click   += self._on_calcular
        self.Controls.Add(self.btn_calcular)

        # ── Botão exportar ───────────────────
        self.btn_exportar = Button()
        self.btn_exportar.Text     = "Exportar .txt"
        self.btn_exportar.Location = Point(695, 34)
        self.btn_exportar.Size     = Size(140, 28)
        self.btn_exportar.Enabled  = False
        self.btn_exportar.Click   += self._on_exportar
        self.Controls.Add(self.btn_exportar)

        # ── Resumo ───────────────────────────
        self.lbl_resumo = Label()
        self.lbl_resumo.Text     = "Aguardando calculo..."
        self.lbl_resumo.Location = Point(12, 40)
        self.lbl_resumo.Size     = Size(510, 20)
        self.Controls.Add(self.lbl_resumo)

        # ── Área de resultado ─────────────────
        self.txt_resultado = RichTextBox()
        self.txt_resultado.Location   = Point(12, 70)
        self.txt_resultado.Size       = Size(820, 490)
        self.txt_resultado.ReadOnly   = True
        self.txt_resultado.ScrollBars = RichTextBoxScrollBars.Both
        self.txt_resultado.Font       = Font("Courier New", 8)
        self.txt_resultado.BackColor  = Color.FromArgb(30, 30, 30)
        self.txt_resultado.ForeColor  = Color.White
        self.txt_resultado.WordWrap   = False
        self.Controls.Add(self.txt_resultado)

    # ── EVENTOS ──────────────────────────────────────────────────────────

    def _on_calcular(self, sender, args):
        self.btn_calcular.Enabled = False
        self.btn_exportar.Enabled = False
        self.lbl_resumo.Text      = "Coletando elementos..."
        Application.DoEvents()

        try:
            self.registros = coletar_volumes(self.doc)

            if not self.registros:
                MessageBox.Show(
                    "Nenhum elemento com volume encontrado no modelo.\n"
                    "Verifique se os elementos estruturais possuem o parametro 'Volume'.",
                    "Atencao",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return

            self.relatorio = montar_relatorio(self.registros)
            self.txt_resultado.Text = self.relatorio

            total = sum(r["volume"] for r in self.registros)
            self.lbl_resumo.Text = "Elementos: {}  |  Volume Total: {} m3".format(
                len(self.registros), fmt_vol(total)
            )
            self.btn_exportar.Enabled = True

        except Exception as ex:
            MessageBox.Show(
                "Erro durante o calculo:\n{}".format(str(ex)),
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
        finally:
            self.btn_calcular.Enabled = True

    def _on_exportar(self, sender, args):
        dlg          = SaveFileDialog()
        dlg.Title    = "Exportar Volume de Concreto"
        dlg.Filter   = "Arquivo de texto (*.txt)|*.txt"
        dlg.FileName = "volume_concreto.txt"

        if dlg.ShowDialog() == DialogResult.OK:
            try:
                exportar_txt(self.relatorio, dlg.FileName)
                MessageBox.Show(
                    "Relatorio exportado!\n{}".format(dlg.FileName),
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
    form = VolumeForm(doc)
    form.ShowDialog()
