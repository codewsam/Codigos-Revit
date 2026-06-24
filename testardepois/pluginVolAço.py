# -*- coding: utf-8 -*-
__title__   = "Volume de Aco"
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
    Form, Label, Button,
    DialogResult, FormBorderStyle, FormStartPosition,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    RichTextBoxScrollBars, RichTextBox,
    Application, SaveFileDialog
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from pyrevit import forms, revit, script

doc = revit.doc

# ══════════════════════════════════════════════════════════════
#  CONSTANTES DE CONVERSÃO
# ══════════════════════════════════════════════════════════════

def ft_to_m(v):
    try:
        return UnitUtils.ConvertFromInternalUnits(v, UnitTypeId.Meters)
    except Exception:
        return v * 0.3048

def ft_to_m2(v):
    try:
        return UnitUtils.ConvertFromInternalUnits(v, UnitTypeId.SquareMeters)
    except Exception:
        return v * 0.092903

def ft_to_kg(v):
    """Converte lb (unidade interna Revit para peso) para kg."""
    try:
        return UnitUtils.ConvertFromInternalUnits(v, UnitTypeId.Kilograms)
    except Exception:
        return v * 0.453592

def fmt(v, casas=3):
    if v is None:
        return "-"
    return str(round(float(v), casas))

# ══════════════════════════════════════════════════════════════
#  UTILITÁRIOS
# ══════════════════════════════════════════════════════════════

def get_fase(el):
    try:
        fid = el.CreatedPhaseId
        if fid and fid != ElementId.InvalidElementId:
            f = doc.GetElement(fid)
            return f.Name if f else "Sem Fase"
    except Exception:
        pass
    return "Sem Fase"

def get_nivel(el):
    """Tenta LevelId, depois parâmetros comuns de nível."""
    try:
        lid = el.LevelId
        if lid and lid != ElementId.InvalidElementId:
            lv = doc.GetElement(lid)
            if lv:
                return lv.Name
    except Exception:
        pass
    for bip in [
        BuiltInParameter.REBAR_ELEM_HOST_LEVEL,
        BuiltInParameter.FABRIC_AREA_LEVEL,
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        BuiltInParameter.SCHEDULE_LEVEL_PARAM,
    ]:
        try:
            p = el.get_Parameter(bip)
            if p:
                lv = doc.GetElement(p.AsElementId())
                if lv:
                    return lv.Name
        except Exception:
            pass
    return "-"

def get_param_double(el, bips=None, names=None):
    """Tenta ler um parâmetro Double por lista de BIP e/ou nomes."""
    if bips:
        for bip in bips:
            try:
                p = el.get_Parameter(bip)
                if p and p.HasValue and p.StorageType == StorageType.Double:
                    return p.AsDouble()
            except Exception:
                pass
    if names:
        for name in names:
            try:
                p = el.LookupParameter(name)
                if p and p.HasValue and p.StorageType == StorageType.Double:
                    return p.AsDouble()
            except Exception:
                pass
    return None

def get_type_name(el):
    try:
        t = doc.GetElement(el.GetTypeId())
        return t.Name if t else "Sem Tipo"
    except Exception:
        return "Sem Tipo"

def get_diametro_mm(el):
    """Tenta ler diâmetro nominal em mm (Rebar e RebarInSystem)."""
    raw = get_param_double(
        el,
        bips=[BuiltInParameter.REBAR_BAR_DIAMETER],
        names=["Diameter", "Bar Diameter", "Diametro", "Nominal Diameter"]
    )
    if raw is not None:
        return round(ft_to_m(raw) * 1000, 1)
    return None

# ══════════════════════════════════════════════════════════════
#  COLETA POR CATEGORIA
# ══════════════════════════════════════════════════════════════

CATEGORIAS_REBAR = [
    ("Armadura (Rebar)",          BuiltInCategory.OST_Rebar),
    ("Armadura em Sistema",       BuiltInCategory.OST_RebarInSystem),
    ("Armadura de Area",          BuiltInCategory.OST_AreaReinforcement),
    ("Armadura de Caminho",       BuiltInCategory.OST_PathReinforcement),
    ("Armadura Estrutural",       BuiltInCategory.OST_StructuralStiffener),
]

CATEGORIAS_MALHA = [
    ("Tela Soldada (Area)",       BuiltInCategory.OST_FabricAreas),
    ("Tela Soldada (Folha)",      BuiltInCategory.OST_FabricReinforcement),
]

def coletar_rebar(cat_nome, cat_bic):
    registros = []
    try:
        elementos = (
            FilteredElementCollector(doc)
            .OfCategory(cat_bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return registros

    for el in elementos:
        # ── Peso ──────────────────────────────────────────────
        peso_raw = get_param_double(
            el,
            bips=[
                BuiltInParameter.REBAR_ELEM_TOTAL_WEIGHT,
                BuiltInParameter.REBAR_TOTAL_WEIGHT,
            ],
            names=["Total Weight", "Peso Total", "Weight", "Peso"]
        )
        peso_kg = ft_to_kg(peso_raw) if peso_raw is not None else None

        # ── Comprimento ───────────────────────────────────────
        comp_raw = get_param_double(
            el,
            bips=[
                BuiltInParameter.REBAR_ELEM_LENGTH,
                BuiltInParameter.CURVE_ELEM_LENGTH,
            ],
            names=["Length", "Total Length", "Comprimento", "Comprimento Total"]
        )
        comp_m = ft_to_m(comp_raw) if comp_raw is not None else None

        # ── Quantidade de barras ───────────────────────────────
        qtd_raw = get_param_double(
            el,
            bips=[BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS],
            names=["Quantity of Bars", "Quantidade de Barras", "Number of Bars"]
        )
        qtd = int(qtd_raw) if qtd_raw is not None else 1

        # Ignora elementos sem dados úteis
        if peso_kg is None and comp_m is None:
            continue

        diam_mm = get_diametro_mm(el)
        diam_str = "phi{}".format(int(diam_mm)) if diam_mm else "Sem diam."

        registros.append({
            "categoria": cat_nome,
            "tipo":      get_type_name(el),
            "nivel":     get_nivel(el),
            "fase":      get_fase(el),
            "diametro":  diam_str,
            "peso_kg":   peso_kg or 0.0,
            "comp_m":    comp_m  or 0.0,
            "qtd":       qtd,
            "id":        el.Id.IntegerValue,
            "origem":    "rebar",
        })

    return registros


def coletar_malha(cat_nome, cat_bic):
    registros = []
    try:
        elementos = (
            FilteredElementCollector(doc)
            .OfCategory(cat_bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return registros

    for el in elementos:
        # ── Peso (FabricArea e FabricSheet guardam em lb) ─────
        peso_raw = get_param_double(
            el,
            bips=[
                BuiltInParameter.FABRIC_AREA_TOTAL_WEIGHT,
                BuiltInParameter.FABRIC_SHEET_WEIGHT,
            ],
            names=["Total Weight", "Peso Total", "Weight"]
        )
        peso_kg = ft_to_kg(peso_raw) if peso_raw is not None else None

        # ── Área (m²) ─────────────────────────────────────────
        area_raw = get_param_double(
            el,
            bips=[
                BuiltInParameter.FABRIC_AREA_AREA,
                BuiltInParameter.HOST_AREA_COMPUTED,
            ],
            names=["Area", "Fabric Area"]
        )
        area_m2 = ft_to_m2(area_raw) if area_raw is not None else None

        if peso_kg is None and area_m2 is None:
            continue

        registros.append({
            "categoria": cat_nome,
            "tipo":      get_type_name(el),
            "nivel":     get_nivel(el),
            "fase":      get_fase(el),
            "diametro":  "Malha",
            "peso_kg":   peso_kg or 0.0,
            "comp_m":    0.0,           # malha não tem comprimento de barra
            "area_m2":   area_m2 or 0.0,
            "qtd":       1,
            "id":        el.Id.IntegerValue,
            "origem":    "malha",
        })

    return registros


def coletar_todos():
    registros = []
    for nome, bic in CATEGORIAS_REBAR:
        registros.extend(coletar_rebar(nome, bic))
    for nome, bic in CATEGORIAS_MALHA:
        registros.extend(coletar_malha(nome, bic))
    return registros

# ══════════════════════════════════════════════════════════════
#  AGRUPAMENTOS
# ══════════════════════════════════════════════════════════════

def agrupar(registros, chave):
    """Agrupa por chave, somando peso e comprimento."""
    grupos = {}
    for r in registros:
        k = r[chave]
        if k not in grupos:
            grupos[k] = {"peso_kg": 0.0, "comp_m": 0.0, "area_m2": 0.0, "qtd": 0}
        grupos[k]["peso_kg"] += r.get("peso_kg", 0.0)
        grupos[k]["comp_m"]  += r.get("comp_m",  0.0)
        grupos[k]["area_m2"] += r.get("area_m2", 0.0)
        grupos[k]["qtd"]     += r.get("qtd",     1)
    return sorted(grupos.items(), key=lambda x: x[0])


def agrupar_cat_fase(registros):
    grupos = {}
    for r in registros:
        k = (r["categoria"], r["fase"])
        if k not in grupos:
            grupos[k] = {"peso_kg": 0.0, "comp_m": 0.0}
        grupos[k]["peso_kg"] += r.get("peso_kg", 0.0)
        grupos[k]["comp_m"]  += r.get("comp_m",  0.0)
    return grupos

# ══════════════════════════════════════════════════════════════
#  RELATÓRIO
# ══════════════════════════════════════════════════════════════

def montar_relatorio(registros):
    L = []

    total_peso = sum(r.get("peso_kg", 0.0) for r in registros)
    total_comp = sum(r.get("comp_m",  0.0) for r in registros)
    total_area = sum(r.get("area_m2", 0.0) for r in registros)
    total_qtd  = sum(r.get("qtd",     1)   for r in registros)

    L.append("=" * 90)
    L.append("  VOLUME / PESO DE ACO E ARMADURA - RELATORIO GERAL")
    L.append("  Autor: Samuel | Versao 1.0")
    L.append("  Total de registros : {}".format(len(registros)))
    L.append("  Peso Total         : {} kg".format(fmt(total_peso, 2)))
    L.append("  Comprimento Total  : {} m  (barras individuais)".format(fmt(total_comp, 2)))
    L.append("  Area de Malha      : {} m2".format(fmt(total_area, 2)))
    L.append("=" * 90)
    L.append("")

    # ── Por Categoria ──────────────────────────────────────────
    L.append("[ POR CATEGORIA ]")
    L.append("-" * 60)
    L.append("{:<28} {:>12} {:>14} {:>10}".format(
        "Categoria", "Peso (kg)", "Comp. (m)", "Area (m2)"))
    L.append("-" * 60)
    for k, v in agrupar(registros, "categoria"):
        L.append("{:<28} {:>12} {:>14} {:>10}".format(
            k,
            fmt(v["peso_kg"], 2),
            fmt(v["comp_m"],  2) if v["comp_m"] > 0 else "-",
            fmt(v["area_m2"], 2) if v["area_m2"] > 0 else "-",
        ))
    L.append("")

    # ── Por Diâmetro ───────────────────────────────────────────
    L.append("[ POR DIAMETRO / TIPO DE BARRA ]")
    L.append("-" * 60)
    L.append("{:<16} {:>12} {:>14} {:>8}".format(
        "Diametro", "Peso (kg)", "Comp. (m)", "Barras"))
    L.append("-" * 60)
    for k, v in agrupar(registros, "diametro"):
        L.append("{:<16} {:>12} {:>14} {:>8}".format(
            k,
            fmt(v["peso_kg"], 2),
            fmt(v["comp_m"],  2) if v["comp_m"] > 0 else "-",
            str(v["qtd"]),
        ))
    L.append("")

    # ── Por Fase ───────────────────────────────────────────────
    L.append("[ POR FASE / ETAPA DE CONCRETAGEM ]")
    L.append("-" * 60)
    L.append("{:<30} {:>12} {:>14}".format("Fase", "Peso (kg)", "Comp. (m)"))
    L.append("-" * 60)
    for k, v in agrupar(registros, "fase"):
        L.append("{:<30} {:>12} {:>14}".format(
            k,
            fmt(v["peso_kg"], 2),
            fmt(v["comp_m"],  2) if v["comp_m"] > 0 else "-",
        ))
    L.append("")

    # ── Por Nível ──────────────────────────────────────────────
    L.append("[ POR NIVEL ]")
    L.append("-" * 60)
    L.append("{:<30} {:>12} {:>14}".format("Nivel", "Peso (kg)", "Comp. (m)"))
    L.append("-" * 60)
    for k, v in agrupar(registros, "nivel"):
        L.append("{:<30} {:>12} {:>14}".format(
            k,
            fmt(v["peso_kg"], 2),
            fmt(v["comp_m"],  2) if v["comp_m"] > 0 else "-",
        ))
    L.append("")

    # ── Tabela Cruzada Categoria x Fase ───────────────────────
    L.append("[ CATEGORIA x FASE  (Peso em kg) ]")
    L.append("-" * 90)

    fases = sorted(set(r["fase"]      for r in registros))
    cats  = sorted(set(r["categoria"] for r in registros))
    grp   = agrupar_cat_fase(registros)

    cab = "{:<26}".format("Categoria")
    for f in fases:
        cab += " {:>14}".format(f[:13])
    cab += " {:>14}".format("TOTAL")
    L.append(cab)
    L.append("-" * 90)

    for cat in cats:
        linha = "{:<26}".format(cat[:25])
        total_cat = 0.0
        for f in fases:
            v = grp.get((cat, f), {}).get("peso_kg", 0.0)
            total_cat += v
            linha += " {:>14}".format(fmt(v, 2) if v > 0 else "-")
        linha += " {:>14}".format(fmt(total_cat, 2))
        L.append(linha)

    L.append("-" * 90)
    linha_tot = "{:<26}".format("TOTAL")
    for f in fases:
        v = sum(grp.get((cat, f), {}).get("peso_kg", 0.0) for cat in cats)
        linha_tot += " {:>14}".format(fmt(v, 2))
    linha_tot += " {:>14}".format(fmt(total_peso, 2))
    L.append(linha_tot)
    L.append("=" * 90)

    # ── Detalhe por Elemento ───────────────────────────────────
    L.append("")
    L.append("[ DETALHE POR ELEMENTO ]")
    L.append("-" * 90)
    L.append("{:<10} {:<22} {:<16} {:<20} {:<12} {:<10} {:<10} {:<8}".format(
        "ID", "Categoria", "Diametro", "Nivel", "Fase",
        "Peso(kg)", "Comp.(m)", "Barras"
    ))
    L.append("-" * 90)

    for r in sorted(registros, key=lambda x: (x["categoria"], x["fase"], x["nivel"], x["diametro"])):
        comp_str = fmt(r["comp_m"], 2) if r.get("comp_m", 0) > 0 else (
            fmt(r.get("area_m2", 0), 2) + "m2" if r.get("area_m2", 0) > 0 else "-"
        )
        L.append("{:<10} {:<22} {:<16} {:<20} {:<12} {:<10} {:<10} {:<8}".format(
            str(r["id"]),
            r["categoria"][:21],
            r["diametro"][:15],
            r["nivel"][:19],
            r["fase"][:11],
            fmt(r.get("peso_kg", 0), 2),
            comp_str,
            str(r.get("qtd", 1)),
        ))

    L.append("=" * 90)
    return "\n".join(L)


def exportar_txt(relatorio, caminho):
    with open(caminho, "w") as f:
        f.write(relatorio)

# ══════════════════════════════════════════════════════════════
#  INTERFACE
# ══════════════════════════════════════════════════════════════

class AcoForm(Form):

    def __init__(self):
        self.registros = []
        self.relatorio = ""
        self._build_ui()

    def _build_ui(self):
        self.Text            = "Volume de Aco / Armadura"
        self.Size            = Size(900, 640)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        lbl = Label()
        lbl.Text     = "Calculo de Peso e Comprimento de Aco por Categoria, Diametro, Fase e Nivel"
        lbl.Location = Point(12, 14)
        lbl.Size     = Size(700, 20)
        lbl.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(lbl)

        self.btn_calc = Button()
        self.btn_calc.Text     = "Calcular Aco"
        self.btn_calc.Location = Point(560, 34)
        self.btn_calc.Size     = Size(150, 28)
        self.btn_calc.Click   += self._on_calcular
        self.Controls.Add(self.btn_calc)

        self.btn_exp = Button()
        self.btn_exp.Text     = "Exportar .txt"
        self.btn_exp.Location = Point(720, 34)
        self.btn_exp.Size     = Size(140, 28)
        self.btn_exp.Enabled  = False
        self.btn_exp.Click   += self._on_exportar
        self.Controls.Add(self.btn_exp)

        self.lbl_resumo = Label()
        self.lbl_resumo.Text     = "Aguardando calculo..."
        self.lbl_resumo.Location = Point(12, 40)
        self.lbl_resumo.Size     = Size(540, 20)
        self.Controls.Add(self.lbl_resumo)

        self.txt = RichTextBox()
        self.txt.Location   = Point(12, 70)
        self.txt.Size       = Size(860, 510)
        self.txt.ReadOnly   = True
        self.txt.ScrollBars = RichTextBoxScrollBars.Both
        self.txt.Font       = Font("Courier New", 8)
        self.txt.BackColor  = Color.FromArgb(30, 30, 30)
        self.txt.ForeColor  = Color.White
        self.txt.WordWrap   = False
        self.Controls.Add(self.txt)

    def _on_calcular(self, sender, args):
        self.btn_calc.Enabled = False
        self.btn_exp.Enabled  = False
        self.lbl_resumo.Text  = "Coletando elementos de aco..."
        Application.DoEvents()

        try:
            self.registros = coletar_todos()

            if not self.registros:
                MessageBox.Show(
                    "Nenhum elemento de aco/armadura encontrado no modelo.\n"
                    "Verifique se existem Rebar, RebarInSystem ou FabricArea no projeto.",
                    "Atencao",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return

            self.relatorio      = montar_relatorio(self.registros)
            self.txt.Text       = self.relatorio
            self.btn_exp.Enabled = True

            total_peso = sum(r.get("peso_kg", 0.0) for r in self.registros)
            total_comp = sum(r.get("comp_m",  0.0) for r in self.registros)
            self.lbl_resumo.Text = (
                "Registros: {}  |  Peso Total: {} kg  |  Comp. Total: {} m".format(
                    len(self.registros), fmt(total_peso, 2), fmt(total_comp, 2)
                )
            )

        except Exception as ex:
            MessageBox.Show(
                "Erro durante o calculo:\n{}".format(str(ex)),
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
        finally:
            self.btn_calc.Enabled = True

    def _on_exportar(self, sender, args):
        dlg          = SaveFileDialog()
        dlg.Title    = "Exportar Relatorio de Aco"
        dlg.Filter   = "Arquivo de texto (*.txt)|*.txt"
        dlg.FileName = "volume_aco.txt"

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

# ══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    form = AcoForm()
    form.ShowDialog()
