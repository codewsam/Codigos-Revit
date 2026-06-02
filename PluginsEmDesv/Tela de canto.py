# -*- coding: utf-8 -*-
__title__   = "Tela de Canto"
__author__  = "Samuel"
__version__ = "Versao 1.1"

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
    Form, Label, ComboBox, TextBox, Button, CheckBox,
    DialogResult, FormBorderStyle, FormStartPosition,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Size, Point
from pyrevit import forms, revit, script


doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

CM_TO_FT  = 1.0 / 30.48
MM_TO_FT  = 1.0 / 304.8
FT_TO_CM  = 30.48

RECOBRIMENTO_FT = 22.0 * MM_TO_FT


def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)


def get_wall_base_z(wall):
    """
    Retorna o Z real da base da parede usando o BoundingBox da geometria.
    Mais confiavel que nivel+offset pois reflete o Z exato da geometria,
    eliminando fragmentos causados por desalinhamento de origem.
    """
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Min.Z
    # Fallback por nivel + offset
    base_level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
    base_level    = doc.GetElement(base_level_id)
    base_elev     = base_level.Elevation if base_level else 0.0
    offset_param  = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    base_offset   = offset_param.AsDouble() if offset_param else 0.0
    return base_elev + base_offset


def get_wall_top_z(wall):
    """
    Retorna o Z real do topo da parede usando o BoundingBox da geometria.
    """
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Max.Z
    # Fallback por nivel + altura
    base_z  = get_wall_base_z(wall)
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    height  = h_param.AsDouble() if h_param else (2.7 / 0.3048)
    return base_z + height


def get_wall_height(wall):
    """
    Retorna a altura real da parede como diferenca entre topo e base
    do BoundingBox, garantindo que a tela cobre exatamente a geometria
    sem sobras nem fragmentos.
    """
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Max.Z - bb.Min.Z
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    return h_param.AsDouble() if h_param else (2.7 / 0.3048)


def get_wall_curve(wall):
    return wall.Location.Curve


def get_wall_direction(wall):
    curve = get_wall_curve(wall)
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    L  = (dx * dx + dy * dy) ** 0.5
    return XYZ(dx / L, dy / L, 0.0)


def ponto_intersecao_2d(p0, d0, p1, d1):
    det = d0.X * (-d1.Y) - (-d1.X) * d0.Y
    if abs(det) < 1e-9:
        return None
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    t  = (dx * (-d1.Y) - (-d1.X) * dy) / det
    return XYZ(p0.X + t * d0.X, p0.Y + t * d0.Y, p0.Z)


def encontrar_canto(wall_a, wall_b):
    curva_a = get_wall_curve(wall_a)
    curva_b = get_wall_curve(wall_b)

    a0 = curva_a.GetEndPoint(0)
    a1 = curva_a.GetEndPoint(1)
    b0 = curva_b.GetEndPoint(0)
    b1 = curva_b.GetEndPoint(1)

    TOLERANCIA = 0.5

    pares = [
        (a0, 0, b0, 0),
        (a0, 0, b1, 1),
        (a1, 1, b0, 0),
        (a1, 1, b1, 1),
    ]

    for pa, idx_a, pb, idx_b in pares:
        dist = pa.DistanceTo(pb)
        if dist < TOLERANCIA:
            ponto_medio = XYZ(
                (pa.X + pb.X) / 2.0,
                (pa.Y + pb.Y) / 2.0,
                min(pa.Z, pb.Z)
            )
            return (ponto_medio, idx_a, idx_b)

    dir_a = get_wall_direction(wall_a)
    dir_b = get_wall_direction(wall_b)
    pt_int = ponto_intersecao_2d(a0, dir_a, b0, dir_b)
    if pt_int is None:
        return None

    for pa, idx_a in [(a0, 0), (a1, 1)]:
        for pb, idx_b in [(b0, 0), (b1, 1)]:
            da = pt_int.DistanceTo(pa)
            db = pt_int.DistanceTo(pb)
            if da < TOLERANCIA and db < TOLERANCIA:
                return (pt_int, idx_a, idx_b)

    return None


def criar_loop_tela_canto(wall_a, wall_b, ponto_canto, idx_a, idx_b,
                           largura_ft, altura_ft_a, altura_ft_b):
    """
    Cria o CurveLoop em formato de aba para cada parede do canto.

    CORRECAO:
    - base_z e top_z vem diretamente do BoundingBox da parede (Min.Z / Max.Z),
      que refletem a geometria real independente de nivel, offset ou modo de
      altura configurado na parede.
    - Quando altura_automatica=True, o loop vai de Min.Z a Max.Z exatos.
    - Quando altura manual, o loop vai de Min.Z ate Min.Z + altura_manual.
    - A origem passada para FabricArea.Create tem o mesmo Z de base_z,
      garantindo que o layout de folhas comece exatamente na base e nao
      gere fragmentos residuais.
    """
    dir_a = get_wall_direction(wall_a)
    dir_b = get_wall_direction(wall_b)

    sinal_a = 1.0 if idx_a == 0 else -1.0
    sinal_b = 1.0 if idx_b == 0 else -1.0

    # XY do ponto de canto (apenas coordenadas horizontais)
    cx = ponto_canto.X
    cy = ponto_canto.Y

    # Base Z exata de cada parede (do BoundingBox)
    base_z_a = get_wall_base_z(wall_a)
    base_z_b = get_wall_base_z(wall_b)

    # Topo Z: se altura automatica, usa Max.Z do BoundingBox;
    # se manual, usa base + altura_ft passada
    top_z_a = base_z_a + altura_ft_a
    top_z_b = base_z_b + altura_ft_b

    # ── Aba 1: ao longo da wall_a ─────────────────────────────
    c0_a = XYZ(cx, cy, base_z_a)
    c1_a = XYZ(cx + sinal_a * dir_a.X * largura_ft,
               cy + sinal_a * dir_a.Y * largura_ft,
               base_z_a)
    c2_a = XYZ(c1_a.X, c1_a.Y, top_z_a)
    c3_a = XYZ(c0_a.X, c0_a.Y, top_z_a)

    loop_a = CurveLoop()
    loop_a.Append(Line.CreateBound(c0_a, c1_a))
    loop_a.Append(Line.CreateBound(c1_a, c2_a))
    loop_a.Append(Line.CreateBound(c2_a, c3_a))
    loop_a.Append(Line.CreateBound(c3_a, c0_a))

    # ── Aba 2: ao longo da wall_b ─────────────────────────────
    c0_b = XYZ(cx, cy, base_z_b)
    c1_b = XYZ(cx + sinal_b * dir_b.X * largura_ft,
               cy + sinal_b * dir_b.Y * largura_ft,
               base_z_b)
    c2_b = XYZ(c1_b.X, c1_b.Y, top_z_b)
    c3_b = XYZ(c0_b.X, c0_b.Y, top_z_b)

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

    def __init__(self, tipos_disponiveis):
        Form.__init__(self)
        self.Text            = "Tela de Canto"
        self.Size            = Size(360, 310)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        self.tipo_selecionado   = None
        self.largura_cm         = None
        self.altura_cm          = None
        self.altura_automatica  = False

        padding_x    = 20
        largura_ctrl = 300
        y            = 20

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

        self.lbl_alt = Label()
        self.lbl_alt.Text     = "Altura da Tela (cm):"
        self.lbl_alt.Location = Point(padding_x, y)
        self.lbl_alt.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(self.lbl_alt)

        y += 22
        self.txt_altura = TextBox()
        self.txt_altura.Location = Point(padding_x, y)
        self.txt_altura.Size     = Size(largura_ctrl, 24)
        self.txt_altura.Text     = "200"
        self.Controls.Add(self.txt_altura)

        y += 36

        self.chk_auto = CheckBox()
        self.chk_auto.Text     = "Altura automatica pela parede"
        self.chk_auto.Location = Point(padding_x, y)
        self.chk_auto.Size     = Size(largura_ctrl, 22)
        self.chk_auto.Checked  = False
        self.chk_auto.CheckedChanged += self.ao_mudar_checkbox
        self.Controls.Add(self.chk_auto)

        y += 36

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

    def ao_mudar_checkbox(self, sender, e):
        auto = self.chk_auto.Checked
        self.txt_altura.Enabled  = not auto
        self.lbl_alt.Enabled     = not auto
        if auto:
            self.txt_altura.Text = "(altura da parede)"

    def ao_clicar_ok(self, sender, e):
        if self.cmb_tipo.SelectedIndex < 0:
            MessageBox.Show(
                "Selecione um tipo de Tela de Canto.",
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        try:
            largura = float(self.txt_largura.Text.replace(",", "."))
            if largura <= 0:
                raise ValueError()
        except Exception:
            MessageBox.Show(
                "Informe uma largura valida (numero positivo em cm).",
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        if not self.chk_auto.Checked:
            try:
                altura = float(self.txt_altura.Text.replace(",", "."))
                if altura <= 0:
                    raise ValueError()
            except Exception:
                MessageBox.Show(
                    "Informe uma altura valida (numero positivo em cm).",
                    "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
                )
                return
            self.altura_cm = altura
        else:
            self.altura_cm = None

        self.tipo_selecionado  = self.cmb_tipo.SelectedItem
        self.largura_cm        = largura
        self.altura_automatica = self.chk_auto.Checked
        self.DialogResult      = DialogResult.OK
        self.Close()

    def ao_clicar_cancelar(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ── FILTRO DE SELECAO DE PAREDES ──────────────────────────────

class WallFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, Wall)
    def AllowReference(self, ref, pt):
        return False


# ── FLUXO PRINCIPAL ───────────────────────────────────────────

fat_map, fst_map = coletar_tipos_tela()

janela = JanelaTelaCanto(sorted(fat_map.keys()))
resultado_janela = janela.ShowDialog()

if resultado_janela != DialogResult.OK:
    script.exit()

fat_name          = janela.tipo_selecionado
largura_ft        = janela.largura_cm * CM_TO_FT
altura_automatica = janela.altura_automatica
altura_ft_manual  = janela.altura_cm * CM_TO_FT if not altura_automatica else None

selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

selected_fst = resolver_sheet_type(fat_name, fst_map)
if not selected_fst:
    forms.alert(
        u"Nao foi possivel encontrar a folha automaticamente para '{}'.".format(fat_name),
        exitscript=True
    )
fabric_sheet_type_id = selected_fst.Id

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

criados  = 0
erros    = []
cantos_processados = 0

with revit.Transaction("Tela de Canto"):
    for wall_a, wall_b, ponto_canto, idx_a, idx_b in cantos_encontrados:
        cantos_processados += 1

        # ── Resolve a altura de cada parede ──────────────────
        # Altura automatica: usa diferenca exata Max.Z - Min.Z do BoundingBox.
        # Altura manual: usa o valor do usuario (a partir da base da parede).
        if altura_automatica:
            altura_ft_a = get_wall_height(wall_a)   # bb.Max.Z - bb.Min.Z
            altura_ft_b = get_wall_height(wall_b)
        else:
            altura_ft_a = altura_ft_manual
            altura_ft_b = altura_ft_manual

        abas = None
        try:
            abas = criar_loop_tela_canto(
                wall_a, wall_b,
                ponto_canto, idx_a, idx_b,
                largura_ft, altura_ft_a, altura_ft_b
            )
        except Exception as e:
            erros.append(
                u"Canto {}: erro ao calcular geometria: {}".format(
                    cantos_processados, str(e)
                )
            )
            continue

        for loop, wall_ref, direcao, origem in abas:
            try:
                curve_loops = List[CurveLoop]()
                curve_loops.Add(loop)

                # CORRECAO: majorDirection = Z (vertical)
                # Com majorDirection horizontal (direcao da parede), o Revit
                # empilha folhas pela altura da parede usando a dimensao CURTA
                # da folha (ex: Q138 = 2,45m por camada). Numa parede de 3m:
                #   1 folha (2,45m) + fragmento (0,55m) = BUG
                #
                # Com majorDirection vertical (Z), a dimensao LONGA da folha
                # (ex: 6,00m do Q138) fica na vertical, cobrindo qualquer
                # parede de ate 6m com 1 unica folha, sem fragmentos.
                direcao_vertical = XYZ(0.0, 0.0, 1.0)

                fa = FabricArea.Create(
                    doc,
                    wall_ref,
                    curve_loops,
                    direcao_vertical,
                    origem,
                    fabric_area_type_id,
                    fabric_sheet_type_id
                )

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

# ── Resumo ────────────────────────────────────────────────────
altura_info = (
    "Automatica (por parede)" if altura_automatica
    else "{} cm".format(int(janela.altura_cm))
)

msg = (
    u"Tela de Canto aplicada!\n\n"
    u"Tipo       : {}\n"
    u"Folha      : {}\n"
    u"Largura    : {} cm\n"
    u"Altura     : {}\n"
    u"Cantos detectados : {}\n"
    u"Abas criadas      : {}/{}\n"
    u"Recobrimento      : 22 mm"
).format(
    fat_name,
    get_name(selected_fst),
    int(janela.largura_cm),
    altura_info,
    cantos_processados,
    criados,
    cantos_processados * 2
)

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Tela de Canto")
