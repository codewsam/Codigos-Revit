# -*- coding: utf-8 -*-
__title__   = "Tela de Laje"
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
    Form, Label, ComboBox, TextBox, Button,
    DialogResult, FormBorderStyle, FormStartPosition,
    ComboBoxStyle, MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Size, Point
from pyrevit import forms, revit, script


doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

CM_TO_FT = 1.0 / 30.48
MM_TO_FT = 1.0 / 304.8
FT_TO_CM = 30.48

RECOBRIMENTO_FT = 22.0 * MM_TO_FT


# ── HELPERS (mesmo padrão do código-base) ────────────────────

def get_name(el):
    p = el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
    return p.AsString() if p else "Id_{}".format(el.Id.IntegerValue)


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


# ── LÓGICA DE LAJE ────────────────────────────────────────────

def coletar_lajes():
    """Retorna todas as lajes (Floor) do modelo."""
    lajes = list(
        FilteredElementCollector(doc)
        .OfClass(Floor)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return [l for l in lajes if l.IsValidObject]


def obter_loops_laje(laje):
    """
    Extrai os CurveLoops da face superior da laje.
    Retorna lista de CurveLoop prontos para uso no FabricArea.Create.
    """
    solid = None
    opts  = Options()
    opts.ComputeReferences = True
    opts.DetailLevel       = ViewDetailLevel.Fine

    geom_elem = laje.get_Geometry(opts)
    for geom_obj in geom_elem:
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 1e-9:
            solid = geom_obj
            break
        if isinstance(geom_obj, GeometryInstance):
            for g in geom_obj.GetInstanceGeometry():
                if isinstance(g, Solid) and g.Volume > 1e-9:
                    solid = g
                    break

    if solid is None:
        return None, None

    # Localiza a face horizontal superior (maior Z médio)
    face_superior = None
    z_max = -1e18
    for face in solid.Faces:
        normal = face.FaceNormal
        if abs(normal.Z - 1.0) < 0.1:          # face apontando para cima
            bbox = face.GetBoundingBox()
            z_centro = (bbox.Min.V + bbox.Max.V) * 0.5
            # Usa o Z real obtendo um ponto da face
            uv_centro = UV(
                (bbox.Min.U + bbox.Max.U) * 0.5,
                (bbox.Min.V + bbox.Max.V) * 0.5
            )
            pt = face.Evaluate(uv_centro)
            if pt.Z > z_max:
                z_max        = pt.Z
                face_superior = face

    if face_superior is None:
        return None, None

    loops = face_superior.GetEdgesAsCurveLoops()
    if not loops or len(list(loops)) == 0:
        return None, None

    # Ponto de origem: canto mínimo da face superior
    bbox    = face_superior.GetBoundingBox()
    uv_orig = UV(bbox.Min.U, bbox.Min.V)
    origem  = face_superior.Evaluate(uv_orig)

    curve_loops = List[CurveLoop]()
    for loop in face_superior.GetEdgesAsCurveLoops():
        curve_loops.Add(loop)

    return curve_loops, origem


# ── INTERFACE GRÁFICA (WinForms) ──────────────────────────────

class JanelaTelaLaje(Form):

    def __init__(self, tipos_disponiveis):
        Form.__init__(self)
        self.Text            = "Tela de Laje Automatica"
        self.Size            = Size(360, 230)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        self.tipo_selecionado  = None
        self.transpasse_min_cm = None
        self.transpasse_max_cm = None

        padding_x    = 20
        largura_ctrl = 300
        y            = 20

        # Tipo de Tela
        lbl_tipo = Label()
        lbl_tipo.Text     = "Tipo de Tela:"
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

        # Transpasse Mínimo
        lbl_min = Label()
        lbl_min.Text     = "Transpasse Minimo (cm):"
        lbl_min.Location = Point(padding_x, y)
        lbl_min.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(lbl_min)

        y += 22
        self.txt_min = TextBox()
        self.txt_min.Location = Point(padding_x, y)
        self.txt_min.Size     = Size(largura_ctrl, 24)
        self.txt_min.Text     = "20"
        self.Controls.Add(self.txt_min)

        y += 36

        # Transpasse Máximo
        lbl_max = Label()
        lbl_max.Text     = "Transpasse Maximo (cm):"
        lbl_max.Location = Point(padding_x, y)
        lbl_max.Size     = Size(largura_ctrl, 20)
        self.Controls.Add(lbl_max)

        y += 22
        self.txt_max = TextBox()
        self.txt_max.Location = Point(padding_x, y)
        self.txt_max.Size     = Size(largura_ctrl, 24)
        self.txt_max.Text     = "50"
        self.Controls.Add(self.txt_max)

        y += 36

        # Botões
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
        if self.cmb_tipo.SelectedIndex < 0:
            MessageBox.Show(
                "Selecione um tipo de Tela.",
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        try:
            t_min = float(self.txt_min.Text.replace(",", "."))
            if t_min < 0:
                raise ValueError()
        except Exception:
            MessageBox.Show(
                "Informe um transpasse minimo valido (numero >= 0 em cm).",
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        try:
            t_max = float(self.txt_max.Text.replace(",", "."))
            if t_max < t_min:
                raise ValueError()
        except Exception:
            MessageBox.Show(
                "Informe um transpasse maximo valido (>= transpasse minimo, em cm).",
                "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        self.tipo_selecionado  = self.cmb_tipo.SelectedItem
        self.transpasse_min_cm = t_min
        self.transpasse_max_cm = t_max
        self.DialogResult      = DialogResult.OK
        self.Close()

    def ao_clicar_cancelar(self, sender, e):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ── FLUXO PRINCIPAL ───────────────────────────────────────────

fat_map, fst_map = coletar_tipos_tela()

janela = JanelaTelaLaje(sorted(fat_map.keys()))
resultado_janela = janela.ShowDialog()

if resultado_janela != DialogResult.OK:
    script.exit()

fat_name           = janela.tipo_selecionado
transpasse_min_ft  = janela.transpasse_min_cm * CM_TO_FT
transpasse_max_ft  = janela.transpasse_max_cm * CM_TO_FT

selected_fat        = fat_map[fat_name]
fabric_area_type_id = selected_fat.Id

selected_fst = resolver_sheet_type(fat_name, fst_map)
if not selected_fst:
    forms.alert(
        u"Nao foi possivel encontrar a folha automaticamente para '{}'.".format(fat_name),
        exitscript=True
    )
fabric_sheet_type_id = selected_fst.Id

# Coleta todas as lajes automaticamente — sem seleção manual
lajes = coletar_lajes()

if not lajes:
    forms.alert("Nenhuma laje encontrada no modelo.", exitscript=True)

criados       = 0
ignorados     = 0
erros         = []

with revit.Transaction("Tela de Laje Automatica"):
    for laje in lajes:
        try:
            curve_loops, origem = obter_loops_laje(laje)

            if curve_loops is None or curve_loops.Count == 0:
                ignorados += 1
                erros.append(
                    u"Laje {}: nao foi possivel extrair geometria.".format(
                        laje.Id.IntegerValue
                    )
                )
                continue

            # Direção horizontal X para a malha da tela de laje
            direcao_horizontal = XYZ(1.0, 0.0, 0.0)

            fa = FabricArea.Create(
                doc,
                laje,
                curve_loops,
                direcao_horizontal,
                origem,
                fabric_area_type_id,
                fabric_sheet_type_id
            )

            # Aplica transpasse mínimo se o parâmetro existir
            p_min = fa.LookupParameter(u"Transpasse minimo")
            if p_min and not p_min.IsReadOnly:
                p_min.Set(transpasse_min_ft)

            # Aplica transpasse máximo se o parâmetro existir
            p_max = fa.LookupParameter(u"Transpasse maximo")
            if p_max and not p_max.IsReadOnly:
                p_max.Set(transpasse_max_ft)

            # Recobrimento padrão (mesmo padrão do código-base: 22 mm)
            p_recob = fa.LookupParameter(u"Deslocamento adicional da recobrimento")
            if p_recob and not p_recob.IsReadOnly:
                p_recob.Set(RECOBRIMENTO_FT)

            criados += 1

        except Exception as ex:
            ignorados += 1
            erros.append(
                u"Laje {}: {}".format(laje.Id.IntegerValue, str(ex))
            )

# ── Resumo ────────────────────────────────────────────────────
msg = (
    u"Tela de Laje aplicada!\n\n"
    u"Tipo            : {}\n"
    u"Folha           : {}\n"
    u"Transpasse Min  : {} cm\n"
    u"Transpasse Max  : {} cm\n"
    u"Lajes no modelo : {}\n"
    u"Telas criadas   : {}\n"
    u"Ignoradas       : {}\n"
    u"Recobrimento    : 22 mm"
).format(
    fat_name,
    get_name(selected_fst),
    int(janela.transpasse_min_cm),
    int(janela.transpasse_max_cm),
    len(lajes),
    criados,
    ignorados
)

if erros:
    msg += u"\n\nErros:\n" + u"\n".join(erros)

forms.alert(msg, warn_icon=bool(erros), title="Tela de Laje Automatica")
