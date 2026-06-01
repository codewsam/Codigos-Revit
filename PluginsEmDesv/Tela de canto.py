# -*- coding: utf-8 -*-
__title__   = "Tela de Canto"
__author__  = "Samuel"
__version__ = "Versao 1.0"

"""
Plugin para insercao automatica de Telas de Canto em cantos de paredes.
Detecta intersecoes entre paredes selecionadas e insere FabricArea em formato L.
Baseado na arquitetura do plugin Folha de Tela Soldada.
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import *
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form, Label, ComboBox, TextBox, Button,
    DialogResult, FormBorderStyle, FormStartPosition,
    ComboBoxStyle, DockStyle, AnchorStyles,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Size, Point, Font, FontStyle, Color
from pyrevit import forms, revit, script

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ── CONSTANTES DE CONVERSAO ───────────────────────────────────
CM_TO_FT  = 1.0 / 30.48
MM_TO_FT  = 1.0 / 304.8
FT_TO_CM  = 30.48

RECOBRIMENTO_FT = 22.0 * MM_TO_FT   # 22 mm padrao


# ── FUNCOES AUXILIARES ────────────────────────────────────────

def get_name(el):
    """Retorna o nome do tipo do elemento via parametro BuiltIn."""
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)


def get_wall_height(wall):
    """Retorna a altura da parede em pes."""
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    return h_param.AsDouble() if h_param else (2.7 / 0.3048)


def get_wall_curve(wall):
    """Retorna a curva de localizacao da parede."""
    return wall.Location.Curve


def get_wall_direction(wall):
    """Retorna o vetor direcional unitario da parede."""
    curve = get_wall_curve(wall)
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    L  = (dx * dx + dy * dy) ** 0.5
    return XYZ(dx / L, dy / L, 0.0)


def ponto_intersecao_2d(p0, d0, p1, d1):
    """
    Calcula o ponto de intersecao 2D entre duas retas infinitas.
    Parametros: ponto inicial e direcao de cada reta.
    Retorna XYZ do ponto de intersecao ou None se paralelas.
    """
    # Resolve sistema: p0 + t*d0 = p1 + s*d1
    # t*d0.X - s*d1.X = p1.X - p0.X
    # t*d0.Y - s*d1.Y = p1.Y - p0.Y
    det = d0.X * (-d1.Y) - (-d1.X) * d0.Y
    if abs(det) < 1e-9:
        return None  # Paredes paralelas, sem canto
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    t  = (dx * (-d1.Y) - (-d1.X) * dy) / det
    return XYZ(p0.X + t * d0.X, p0.Y + t * d0.Y, p0.Z)


def encontrar_canto(wall_a, wall_b):
    """
    Tenta encontrar o ponto de canto entre duas paredes.
    Retorna (ponto_canto, extremidade_a, extremidade_b) ou None.
    - extremidade_a: 0 ou 1 indicando qual ponta da wall_a esta no canto
    - extremidade_b: 0 ou 1 indicando qual ponta da wall_b esta no canto
    """
    curva_a = get_wall_curve(wall_a)
    curva_b = get_wall_curve(wall_b)

    a0 = curva_a.GetEndPoint(0)
    a1 = curva_a.GetEndPoint(1)
    b0 = curva_b.GetEndPoint(0)
    b1 = curva_b.GetEndPoint(1)

    TOLERANCIA = 0.5  # pes (~15 cm)

    # Verifica os 4 pares de extremidades possiveis
    pares = [
        (a0, 0, b0, 0),
        (a0, 0, b1, 1),
        (a1, 1, b0, 0),
        (a1, 1, b1, 1),
    ]

    for pa, idx_a, pb, idx_b in pares:
        dist = pa.DistanceTo(pb)
        if dist < TOLERANCIA:
            # Extremidades proximas: canto encontrado
            ponto_medio = XYZ(
                (pa.X + pb.X) / 2.0,
                (pa.Y + pb.Y) / 2.0,
                min(pa.Z, pb.Z)
            )
            return (ponto_medio, idx_a, idx_b)

    # Tenta intersecao geometrica das retas
    dir_a = get_wall_direction(wall_a)
    dir_b = get_wall_direction(wall_b)
    pt_int = ponto_intersecao_2d(a0, dir_a, b0, dir_b)
    if pt_int is None:
        return None

    # Verifica se o ponto de intersecao esta perto de alguma extremidade
    for pa, idx_a in [(a0, 0), (a1, 1)]:
        for pb, idx_b in [(b0, 0), (b1, 1)]:
            da = pt_int.DistanceTo(pa)
            db = pt_int.DistanceTo(pb)
            if da < TOLERANCIA and db < TOLERANCIA:
                return (pt_int, idx_a, idx_b)

    return None


def criar_loop_tela_canto(wall_a, wall_b, ponto_canto, idx_a, idx_b,
                           largura_ft, altura_ft):
    """
    Cria o CurveLoop em formato L para a Tela de Canto.

    A tela cobre:
      - 'largura_ft' ao longo da wall_a a partir do canto
      - 'largura_ft' ao longo da wall_b a partir do canto
      - 'altura_ft' de altura em ambas as abas

    Retorna uma lista de CurveLoop (uma por aba) ou None em caso de erro.
    """
    dir_a = get_wall_direction(wall_a)
    dir_b = get_wall_direction(wall_b)

    # Sentido a partir do canto: se idx=0 o canto e a ponta 0, entao
    # a direcao para dentro da parede e +dir; caso contrario e -dir
    sinal_a = 1.0 if idx_a == 0 else -1.0
    sinal_b = 1.0 if idx_b == 0 else -1.0

    # Aba 1: ao longo da wall_a
    c0_a = XYZ(ponto_canto.X,
               ponto_canto.Y,
               ponto_canto.Z)
    c1_a = XYZ(ponto_canto.X + sinal_a * dir_a.X * largura_ft,
               ponto_canto.Y + sinal_a * dir_a.Y * largura_ft,
               ponto_canto.Z)
    c2_a = XYZ(c1_a.X, c1_a.Y, c1_a.Z + altura_ft)
    c3_a = XYZ(c0_a.X, c0_a.Y, c0_a.Z + altura_ft)

    loop_a = CurveLoop()
    loop_a.Append(Line.CreateBound(c0_a, c1_a))
    loop_a.Append(Line.CreateBound(c1_a, c2_a))
    loop_a.Append(Line.CreateBound(c2_a, c3_a))
    loop_a.Append(Line.CreateBound(c3_a, c0_a))

    # Aba 2: ao longo da wall_b
    c0_b = XYZ(ponto_canto.X,
               ponto_canto.Y,
               ponto_canto.Z)
    c1_b = XYZ(ponto_canto.X + sinal_b * dir_b.X * largura_ft,
               ponto_canto.Y + sinal_b * dir_b.Y * largura_ft,
               ponto_canto.Z)
    c2_b = XYZ(c1_b.X, c1_b.Y, c1_b.Z + altura_ft)
    c3_b = XYZ(c0_b.X, c0_b.Y, c0_b.Z + altura_ft)

    loop_b = CurveLoop()
    loop_b.Append(Line.CreateBound(c0_b, c1_b))
    loop_b.Append(Line.CreateBound(c1_b, c2_b))
    loop_b.Append(Line.CreateBound(c2_b, c3_b))
    loop_b.Append(Line.CreateBound(c3_b, c0_b))

    return [
        (loop_a, wall_a, dir_a, c0_a),
        (loop_b, wall_b, dir_b, c0_b),
    ]


def coletar_tipos_tela():
    """
    Coleta FabricAreaType e FabricSheetType disponiveis no projeto.
    Retorna (fat_map, fst_map) ou encerra o script se nao encontrar.
    """
    fat_list = list(
        FilteredElementCollector(doc)
        .OfClass(FabricAreaType)
        .ToElements()
    )
    fst_list = list(
        FilteredElementCollector(doc)
        .OfClass(FabricSheetType)
        .ToElements()
    )

    if not fat_list:
        forms.alert("Nenhum FabricAreaType encontrado no projeto.", exitscript=True)
    if not fst_list:
        forms.alert("Nenhum FabricSheetType encontrado no projeto.", exitscript=True)

    fat_map = {get_name(t): t for t in fat_list}
    fst_map = {get_name(t): t for t in fst_list}
    return fat_map, fst_map


def resolver_sheet_type(fat_name, fst_map):
    """
    Tenta resolver automaticamente o FabricSheetType correspondente
    ao FabricAreaType selecionado, usando a mesma logica da Tela Soldada.
    """
    sheet_suffix = fat_name.replace("Tela POP ", "").strip()
    resultado = fst_map.get(sheet_suffix)
    if not resultado:
        for k, v in fst_map.items():
            if sheet_suffix in k or k in sheet_suffix:
                resultado = v
                break
    return resultado


# ── INTERFACE GRAFICA (WinForms) ──────────────────────────────

class JanelaTelaCanto(Form):
    """
    Janela de configuracao do plugin Tela de Canto.
    Campos: Tipo de Tela (ComboBox), Largura (cm), Altura (cm).
    """

    def __init__(self, tipos_disponiveis):
        Form.__init__(self)
        self.Text            = "Tela de Canto"
        self.Size            = Size(360, 260)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        # Resultado publico
        self.tipo_selecionado = None
        self.largura_cm       = None
        self.altura_cm        = None

        padding_x    = 20
        largura_ctrl = 300
        y            = 20

        # ── Label: Tipo de Tela ───────────────────────────────
        lbl_tipo = Label()
        lbl_tipo.Text     = "Tipo de Tela de Canto:"
        lbl_tipo.Location = Point(padding_x, y)
        lbl_tipo.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(lbl_tipo)

        y += 22
        self.cmb_tipo = ComboBox()
        self.cmb_tipo.Location      = Point(padding_x, y)
        self.cmb_tipo.Size          = Size(largura_ctrl, 24)
        self.cmb_tipo.DropDownStyle = ComboBoxStyle.DropDownList
        for nome in sorted(tipos_disponiveis):
            self.cmb_tipo.Items.Add(nome)
        if self.cmb_tipo.Items.Count > 0:
            self.cmb_tipo.SelectedIndex = 0
        self.Controls.Add(self.cmb_tipo)

        y += 36

        # ── Label: Largura ───────────────────────────────────
        lbl_larg = Label()
        lbl_larg.Text     = "Largura da Tela (cm):"
        lbl_larg.Location = Point(padding_x, y)
        lbl_larg.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(lbl_larg)

        y += 22
        self.txt_largura = TextBox()
        self.txt_largura.Location = Point(padding_x, y)
        self.txt_largura.Size     = Size(largura_ctrl, 24)
        self.txt_largura.Text     = "50"
        self.Controls.Add(self.txt_largura)

        y += 36

        # ── Label: Altura ────────────────────────────────────
        lbl_alt = Label()
        lbl_alt.Text     = "Altura da Tela (cm):"
        lbl_alt.Location = Point(padding_x, y)
        lbl_alt.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(lbl_alt)

        y += 22
        self.txt_altura = TextBox()
        self.txt_altura.Location = Point(padding_x, y)
        self.txt_altura.Size     = Size(largura_ctrl, 24)
        self.txt_altura.Text     = "200"
        self.Controls.Add(self.txt_altura)

        y += 44

        # ── Botoes OK / Cancelar ─────────────────────────────
        btn_ok = Button()
        btn_ok.Text     = "OK"
        btn_ok.Size     = Size(90, 30)
        btn_ok.Location = Point(padding_x, y)
        btn_ok.Click   += self.ao_clicar_ok
        self.Controls.Add(btn_ok)

        btn_cancelar = Button()
        btn_cancelar.Text     = "Cancelar"
        btn_cancelar.Size     = Size(90, 30)
        btn_cancelar.Location = Point(padding_x + 100, y)
        btn_cancelar.Click   += self.ao_clicar_cancelar
        self.Controls.Add(btn_cancelar)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancelar

    def ao_clicar_ok(self, sender, e):
        """Valida os campos e fecha com DialogResult.OK."""
        if self.cmb_tipo.SelectedIndex < 0:
            MessageBox.Show(
                "Selecione um tipo de Tela de Canto.",
                "Aviso",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        try:
            largura = float(self.txt_largura.Text.replace(",", "."))
            if largura <= 0:
                raise ValueError("Largura deve ser positiva.")
        except Exception:
            MessageBox.Show(
                "Informe uma largura valida (numero positivo em cm).",
                "Aviso",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        try:
            altura = float(self.txt_altura.Text.replace(",", "."))
            if altura <= 0:
                raise ValueError("Altura deve ser positiva.")
        except Exception:
            MessageBox.Show(
                "Informe uma altura valida (numero positivo em cm).",
                "Aviso",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        self.tipo_selecionado = self.cmb_tipo.SelectedItem
        self.largura_cm       = largura
        self.altura_cm        = altura
        self.DialogResult     = DialogResult.OK
        self.Close()

    def ao_clicar_cancelar(self, sender, e):
        """Fecha o formulario sem executar nada."""
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ── FILTRO DE SELECAO DE PAREDES ──────────────────────────────

class WallFilter(ISelectionFilter):
    """Permite selecionar apenas elementos do tipo Wall."""
    def AllowElement(self, el):
        return isinstance(el, Wall)
    def AllowReference(self, ref, pt):
        return False


# ── FLUXO PRINCIPAL ───────────────────────────────────────────

# 1. Coletar tipos disponiveis no projeto
fat_map, fst_map = coletar_tipos_tela()

# 2. Exibir janela de configuracao
janela = JanelaTelaCanto(sorted(fat_map.keys()))
resultado_janela = janela.ShowDialog()

if resultado_janela != DialogResult.OK:
    script.exit()

fat_name       = janela.tipo_selecionado
largura_ft     = janela.largura_cm * CM_TO_FT
altura_ft      = janela.altura_cm  * CM_TO_FT

# 3. Resolver FabricAreaType e FabricSheetType
selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

selected_fst = resolver_sheet_type(fat_name, fst_map)
if not selected_fst:
    forms.alert(
        u"Nao foi possivel encontrar a folha automaticamente para '{}'.".format(fat_name),
        exitscript=True
    )
fabric_sheet_type_id = selected_fst.Id

# 4. Selecionar paredes (minimo 2 para formar um canto)
with forms.WarningBar(title="Selecione as paredes de canto e pressione Enter"):
    try:
        refs  = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallFilter(),
            "Selecione as paredes de canto (minimo 2)"
        )
        walls = [doc.GetElement(r.ElementId) for r in refs]
        walls = [w for w in walls if isinstance(w, Wall)]
    except Exception:
        walls = []

if len(walls) < 2:
    forms.alert("Selecione pelo menos 2 paredes para formar um canto.", exitscript=True)

# 5. Detectar cantos entre as paredes selecionadas
#    Para N paredes, verifica todos os pares possiveis
cantos_encontrados = []

for i in range(len(walls)):
    for j in range(i + 1, len(walls)):
        resultado_canto = encontrar_canto(walls[i], walls[j])
        if resultado_canto:
            ponto_canto, idx_a, idx_b = resultado_canto
            cantos_encontrados.append((walls[i], walls[j], ponto_canto, idx_a, idx_b))

if not cantos_encontrados:
    forms.alert(
        u"Nenhum canto detectado entre as paredes selecionadas.\n"
        u"Verifique se as paredes se encontram ou estao proximas.",
        exitscript=True
    )

# 6. Inserir Telas de Canto dentro de uma unica Transaction
criados  = 0
erros    = []
cantos_processados = 0

with revit.Transaction("Tela de Canto"):
    for wall_a, wall_b, ponto_canto, idx_a, idx_b in cantos_encontrados:
        cantos_processados += 1
        abas = None

        try:
            abas = criar_loop_tela_canto(
                wall_a, wall_b,
                ponto_canto, idx_a, idx_b,
                largura_ft, altura_ft
            )
        except Exception as e:
            erros.append(
                u"Canto {}: erro ao calcular geometria: {}".format(
                    cantos_processados, str(e)
                )
            )
            continue

        # Insere uma FabricArea para cada aba do L
        for loop, wall_ref, direcao, origem in abas:
            try:
                curve_loops = List[CurveLoop]()
                curve_loops.Add(loop)

                fa = FabricArea.Create(
                    doc,
                    wall_ref,
                    curve_loops,
                    direcao,
                    origem,
                    fabric_area_type_id,
                    fabric_sheet_type_id
                )

                # Aplica recobrimento adicional de 22 mm (mesmo padrao da Tela Soldada)
                p_recob = fa.LookupParameter(u"Deslocamento adicional da recobrimento")
                if p_recob and not p_recob.IsReadOnly:
                    p_recob.Set(RECOBRIMENTO_FT)

                criados += 1

            except Exception as e:
                erros.append(
                    u"Canto {} / parede {}: {}".format(
                        cantos_processados,
                        wall_ref.Id.IntegerValue,
                        str(e)
                    )
                )

# 7. Exibir resumo final
msg = (
    u"Tela de Canto aplicada!\n\n"
    u"Tipo       : {}\n"
    u"Folha      : {}\n"
    u"Largura    : {} cm\n"
    u"Altura     : {} cm\n"
    u"Cantos detectados : {}\n"
    u"Abas criadas      : {}/{}\n"
    u"Recobrimento      : 22 mm"
).format(
    fat_name,
    get_name(selected_fst),
    int(janela.largura_cm),
    int(janela.altura_cm),
    cantos_processados,
    criados,
    cantos_processados * 2
)

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Tela de Canto")
