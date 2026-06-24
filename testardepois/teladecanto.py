# -*- coding: utf-8 -*-
__title__   = "Tela de Canto"
__author__  = "Samuel"
__version__ = "Versao 1.2"

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
FT_TO_KG_PER_M2 = 4.88243  # 1 lb/ft² ≈ 4.88 kg/m²
FT2_TO_M2 = 0.092903

RECOBRIMENTO_FT = 22.0 * MM_TO_FT

SCHEDULE_NAME = "Telas de Canto - Resumo"


# ══════════════════════════════════════════════════════════════
#  UTILITÁRIOS
# ══════════════════════════════════════════════════════════════

def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)


def get_wall_base_z(wall):
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Min.Z
    base_level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
    base_level    = doc.GetElement(base_level_id)
    base_elev     = base_level.Elevation if base_level else 0.0
    offset_param  = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    base_offset   = offset_param.AsDouble() if offset_param else 0.0
    return base_elev + base_offset


def get_wall_top_z(wall):
    bb = wall.get_BoundingBox(None)
    if bb:
        return bb.Max.Z
    base_z  = get_wall_base_z(wall)
    h_param = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
    height  = h_param.AsDouble() if h_param else (2.7 / 0.3048)
    return base_z + height


def get_wall_height(wall):
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
    dir_a = get_wall_direction(wall_a)
    dir_b = get_wall_direction(wall_b)

    sinal_a = 1.0 if idx_a == 0 else -1.0
    sinal_b = 1.0 if idx_b == 0 else -1.0

    cx = ponto_canto.X
    cy = ponto_canto.Y

    base_z_a = get_wall_base_z(wall_a)
    base_z_b = get_wall_base_z(wall_b)

    top_z_a = base_z_a + altura_ft_a
    top_z_b = base_z_b + altura_ft_b

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


# ══════════════════════════════════════════════════════════════
#  SCHEDULE DE TELAS DE CANTO
# ══════════════════════════════════════════════════════════════

def get_level_name_for_fa(fa):
    """Retorna o nome do nível da parede hospedeira da FabricArea."""
    try:
        host = doc.GetElement(fa.GetHostIds()[0]) if fa.GetHostIds().Count > 0 else None
        if host and isinstance(host, Wall):
            level_id = host.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
            level = doc.GetElement(level_id)
            if level:
                return level.Name
    except Exception:
        pass
    # fallback: nível próprio da FabricArea
    lp = fa.get_Parameter(BuiltInParameter.FABRIC_AREA_LEVEL)
    if lp:
        lv = doc.GetElement(lp.AsElementId())
        if lv:
            return lv.Name
    return "—"


def get_fabric_area_type_name(fa):
    """Nome do FabricAreaType da área."""
    try:
        fat = doc.GetElement(fa.GetTypeId())
        return get_name(fat) if fat else "—"
    except Exception:
        return "—"


def get_fa_weight_kg(fa):
    """
    Tenta ler o peso total da FabricArea em kg.
    O Revit armazena em lb; converte para kg (÷ 2.20462).
    """
    # Parâmetros candidatos (em ordem de preferência)
    bips = [
        BuiltInParameter.FABRIC_AREA_TOTAL_WEIGHT,   # peso total da área
        BuiltInParameter.FABRIC_AREA_WEIGHT_PER_UNIT_AREA,  # kg/m²
    ]
    for bip in bips:
        p = fa.get_Parameter(bip)
        if p and p.StorageType == StorageType.Double:
            val_lb = p.AsDouble()
            return val_lb / 2.20462
    # fallback: busca por nome
    for name in [u"Peso Total", u"Total Weight", u"Peso"]:
        p = fa.LookupParameter(name)
        if p and p.StorageType == StorageType.Double:
            return p.AsDouble() / 2.20462
    return None


def coletar_dados_telas_canto():
    """
    Varre todos os Groups do projeto em busca de grupos com 2 FabricAreas
    (padrão do plugin Tela de Canto). Retorna lista de dicts com os dados.
    """
    # Coleta todos os groups
    todos_grupos = list(
        FilteredElementCollector(doc)
        .OfClass(Group)
        .ToElements()
    )

    # Coleta todas as FabricAreas do projeto indexadas por Id
    todas_fas = {
        fa.Id.IntegerValue: fa
        for fa in FilteredElementCollector(doc)
                    .OfClass(FabricArea)
                    .ToElements()
    }

    dados = []
    for grupo in todos_grupos:
        try:
            member_ids = grupo.GetMemberIds()
        except Exception:
            continue

        # Filtra apenas membros que são FabricArea
        fas_do_grupo = []
        for mid in member_ids:
            if mid.IntegerValue in todas_fas:
                fas_do_grupo.append(todas_fas[mid.IntegerValue])

        if len(fas_do_grupo) != 2:
            continue  # não é um grupo de Tela de Canto

        fa_a, fa_b = fas_do_grupo[0], fas_do_grupo[1]

        tipo_nome = get_fabric_area_type_name(fa_a)
        nivel_a   = get_level_name_for_fa(fa_a)
        nivel_b   = get_level_name_for_fa(fa_b)
        nivel     = nivel_a if nivel_a == nivel_b else u"{} / {}".format(nivel_a, nivel_b)

        peso_a_kg = get_fa_weight_kg(fa_a)
        peso_b_kg = get_fa_weight_kg(fa_b)

        if peso_a_kg is not None and peso_b_kg is not None:
            peso_total = peso_a_kg + peso_b_kg
            peso_str   = u"{:.2f} kg".format(peso_total)
        elif peso_a_kg is not None:
            peso_str   = u"{:.2f} kg (1 aba)".format(peso_a_kg)
        else:
            peso_str   = u"—"

        dados.append({
            "grupo_id"  : grupo.Id.IntegerValue,
            "tipo"      : tipo_nome,
            "nivel"     : nivel,
            "num_abas"  : len(fas_do_grupo),
            "peso"      : peso_str,
            "fa_ids"    : [fa_a.Id.IntegerValue, fa_b.Id.IntegerValue],
        })

    return dados


def gerar_schedule_telas_canto():
    """
    Cria (ou recria) um Schedule de FabricArea no projeto com as colunas:
    Tipo de Tela | Nível | Nº de Abas | Peso Total
    usando a API nativa de ScheduleView do Revit.
    """
    # ── 1. Remove schedule anterior com o mesmo nome ──────────
    existentes = list(
        FilteredElementCollector(doc)
        .OfClass(ViewSchedule)
        .ToElements()
    )
    for vs in existentes:
        if vs.Name == SCHEDULE_NAME:
            doc.Delete(vs.Id)
            break

    # ── 2. Cria novo Schedule de FabricArea ───────────────────
    cat_id = ElementId(BuiltInCategory.OST_FabricAreas)
    schedule = ViewSchedule.CreateSchedule(doc, cat_id)
    schedule.Name = SCHEDULE_NAME

    sd = schedule.Definition

    # ── 3. Campos disponíveis ─────────────────────────────────
    campos_disponiveis = {
        sf.GetSchedulableField().GetParameterId(doc).IntegerValue: sf
        for sf in sd.GetSchedulableFields()
        if sf.GetSchedulableField().FieldType == ScheduleFieldType.Instance
        or sf.GetSchedulableField().FieldType == ScheduleFieldType.ElementType
    }

    def add_field_by_bip(bip):
        eid = ElementId(bip)
        sf_entry = campos_disponiveis.get(eid.IntegerValue)
        if sf_entry:
            try:
                sd.AddField(sf_entry.GetSchedulableField())
                return True
            except Exception:
                pass
        return False

    def add_field_by_name(name):
        for sf in sd.GetSchedulableFields():
            lbl = sf.GetSchedulableField().GetName(doc)
            if lbl and lbl.lower() == name.lower():
                try:
                    sd.AddField(sf.GetSchedulableField())
                    return True
                except Exception:
                    pass
        return False

    # Tipo de Tela (nome do tipo)
    if not add_field_by_bip(BuiltInParameter.ALL_MODEL_TYPE_NAME):
        add_field_by_name("Type Name")

    # Nível
    add_field_by_bip(BuiltInParameter.FABRIC_AREA_LEVEL)

    # Área total (m²)
    add_field_by_bip(BuiltInParameter.FABRIC_AREA_AREA)

    # Peso total
    added_weight = add_field_by_bip(BuiltInParameter.FABRIC_AREA_TOTAL_WEIGHT)
    if not added_weight:
        add_field_by_name("Total Weight")

    # Peso por unidade de área
    add_field_by_bip(BuiltInParameter.FABRIC_AREA_WEIGHT_PER_UNIT_AREA)

    # ── 4. Agrupa por Tipo + Nível ────────────────────────────
    try:
        fields = sd.GetFieldOrder()
        if len(fields) >= 2:
            # Agrupa pelo campo Tipo (índice 0) e depois Nível (índice 1)
            grp0 = ScheduleGroup()
            grp0.FieldId = fields[0]
            sd.AddGroup(grp0)
    except Exception:
        pass  # agrupamento é cosmético, não crítico

    # ── 5. Ordena por Nível ───────────────────────────────────
    try:
        fields = sd.GetFieldOrder()
        if len(fields) >= 2:
            sort_field = ScheduleSortGroupField()
            sort_field.FieldId = fields[1]   # Nível
            sort_field.SortOrder = ScheduleSortOrder.Ascending
            sd.AddSortGroupField(sort_field)
    except Exception:
        pass

    # ── 6. Totais na última linha ─────────────────────────────
    sd.ShowGrandTotal      = True
    sd.ShowGrandTotalTitle = True
    sd.GrandTotalTitle     = u"TOTAL"

    return schedule


# ══════════════════════════════════════════════════════════════
#  INTERFACE GRÁFICA (WinForms)
# ══════════════════════════════════════════════════════════════

class JanelaTelaCanto(Form):

    def __init__(self, tipos_disponiveis):
        Form.__init__(self)
        self.Text            = "Tela de Canto"
        self.Size            = Size(360, 380)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        self.tipo_selecionado   = None
        self.largura_cm         = None
        self.altura_cm          = None
        self.altura_automatica  = False
        self.gerar_schedule     = False   # ← novo flag

        padding_x    = 20
        largura_ctrl = 300
        y            = 20

        # ── Tipo ──────────────────────────────────────────────
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

        # ── Largura ───────────────────────────────────────────
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

        # ── Altura ────────────────────────────────────────────
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

        # ── Altura automática ─────────────────────────────────
        self.chk_auto = CheckBox()
        self.chk_auto.Text     = "Altura automatica pela parede"
        self.chk_auto.Location = Point(padding_x, y)
        self.chk_auto.Size     = Size(largura_ctrl, 22)
        self.chk_auto.Checked  = False
        self.chk_auto.CheckedChanged += self.ao_mudar_checkbox
        self.Controls.Add(self.chk_auto)

        y += 34

        # ── Separador visual (linha) ──────────────────────────
        sep = Label()
        sep.Text      = u"─" * 44
        sep.Location  = Point(padding_x, y)
        sep.Size      = Size(largura_ctrl, 16)
        sep.ForeColor = System.Drawing.Color.Gray
        self.Controls.Add(sep)

        y += 18

        # ── Gerar Schedule ────────────────────────────────────
        self.chk_schedule = CheckBox()
        self.chk_schedule.Text     = u"Gerar Schedule de Telas de Canto"
        self.chk_schedule.Location = Point(padding_x, y)
        self.chk_schedule.Size     = Size(largura_ctrl, 22)
        self.chk_schedule.Checked  = False
        self.Controls.Add(self.chk_schedule)

        y += 36

        # ── Botões ────────────────────────────────────────────
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
        self.txt_altura.Enabled = not auto
        self.lbl_alt.Enabled    = not auto
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
        self.gerar_schedule    = self.chk_schedule.Checked   # ← lê checkbox
        self.DialogResult      = DialogResult.OK
        self.Close()

    def ao_clicar_cancelar(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ══════════════════════════════════════════════════════════════
#  FILTRO DE SELEÇÃO
# ══════════════════════════════════════════════════════════════

class WallFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, Wall)
    def AllowReference(self, ref, pt):
        return False


# ══════════════════════════════════════════════════════════════
#  FLUXO PRINCIPAL
# ══════════════════════════════════════════════════════════════

fat_map, fst_map = coletar_tipos_tela()

janela = JanelaTelaCanto(sorted(fat_map.keys()))
resultado_janela = janela.ShowDialog()

if resultado_janela != DialogResult.OK:
    script.exit()

fat_name          = janela.tipo_selecionado
largura_ft        = janela.largura_cm * CM_TO_FT
altura_automatica = janela.altura_automatica
altura_ft_manual  = janela.altura_cm * CM_TO_FT if not altura_automatica else None
criar_schedule    = janela.gerar_schedule

selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

selected_fst = resolver_sheet_type(fat_name, fst_map)
if not selected_fst:
    forms.alert(
        u"Nao foi possivel encontrar a folha automaticamente para '{}'.".format(fat_name),
        exitscript=True
    )
fabric_sheet_type_id = selected_fst.Id

# ── Seleção de paredes ────────────────────────────────────────
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
ids_grupos_criados = []   # ← rastreia grupos novos para o schedule

with revit.Transaction("Tela de Canto"):
    for wall_a, wall_b, ponto_canto, idx_a, idx_b in cantos_encontrados:
        cantos_processados += 1

        if altura_automatica:
            altura_ft_a = get_wall_height(wall_a)
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

        ids_grupo = List[ElementId]()

        for loop, wall_ref, direcao, origem in abas:
            try:
                curve_loops = List[CurveLoop]()
                curve_loops.Add(loop)

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

                ids_grupo.Add(fa.Id)
                criados += 1

            except Exception as e:
                erros.append(
                    u"Canto {} / parede {}: {}".format(
                        cantos_processados,
                        wall_ref.Id.IntegerValue,
                        str(e)
                    )
                )

        if ids_grupo.Count == 2:
            try:
                grp = doc.Create.NewGroup(ids_grupo)
                ids_grupos_criados.append(grp.Id)
            except Exception as e:
                erros.append(
                    u"Canto {}: nao foi possivel agrupar as abas: {}".format(
                        cantos_processados, str(e)
                    )
                )

# ── Schedule (transação separada) ─────────────────────────────
schedule_criado = False
schedule_id     = None

if criar_schedule:
    dados = coletar_dados_telas_canto()

    if not dados:
        erros.append(
            u"Schedule: nenhuma Tela de Canto encontrada no projeto "
            u"(grupos de 2 FabricAreas). Verifique se as telas foram criadas."
        )
    else:
        try:
            with revit.Transaction("Schedule - Telas de Canto"):
                sched = gerar_schedule_telas_canto()
                schedule_criado = True
                schedule_id     = sched.Id
        except Exception as e:
            erros.append(u"Erro ao gerar Schedule: {}".format(str(e)))

# ── Resumo ─────────────────────────────────────────────────────
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

if schedule_criado:
    msg += u"\n\nSchedule gerado: \"{}\"".format(SCHEDULE_NAME)
    msg += u"\n(aberto automaticamente no Revit)"
    # Abre o schedule na interface
    try:
        uidoc.ActiveView = doc.GetElement(schedule_id)
    except Exception:
        pass

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Tela de Canto")
