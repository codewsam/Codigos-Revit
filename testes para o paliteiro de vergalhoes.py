# -*- coding: utf-8 -*-
"""
PLUGIN: Paliteiro
VERSAO: 2.0
AUTOR: ChatGPT
COMPATIBILIDADE: Revit 2024+

PLUGIN PYREVIT PARA GERAR
ARMADURAS TIPO "PALITEIRO"
EM PAREDES ESTRUTURAIS
"""

# =========================================================
# IMPORTS
# =========================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.UI.Selection import *

from pyrevit import forms

import math


# =========================================================
# DOCUMENTO
# =========================================================
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# =========================================================
# UTILITARIOS
# =========================================================
def cm_to_feet(cm):
    return float(cm) / 30.48


# =========================================================
# TIPOS DE VERGALHÃO
# =========================================================
def get_rebar_types():

    collector = (
        FilteredElementCollector(doc)
        .OfClass(RebarBarType)
    )

    types = {}

    for r in collector:

        try:

            param = r.LookupParameter("Type Name")

            if param:
                name = param.AsString()
            else:
                name = "Vergalhao_{}".format(
                    r.Id.IntegerValue
                )

        except:

            name = "Vergalhao_{}".format(
                r.Id.IntegerValue
            )

        types[name] = r

    return types


# =========================================================
# FILTRO DE PAREDE
# =========================================================
class WallSelectionFilter(ISelectionFilter):

    def AllowElement(self, elem):

        if isinstance(elem, Wall):
            return True

        return False

    def AllowReference(self, ref, point):
        return True


# =========================================================
# INTERFACE
# =========================================================
def get_user_inputs():

    rebar_types = get_rebar_types()

    if not rebar_types:
        forms.alert(
            "Nenhum tipo de vergalhão encontrado."
        )
        return None

    # =====================================================
    # TIPO DE VERGALHÃO
    # =====================================================
    selected_rebar = forms.SelectFromList.show(
        sorted(rebar_types.keys()),
        title="Tipo de Vergalhão",
        multiselect=False
    )

    if not selected_rebar:
        return None

    # =====================================================
    # ESPAÇAMENTO
    # =====================================================
    espacamento = forms.ask_for_string(
        default="24",
        prompt="Espaçamento máximo (cm):",
        title="Paliteiro"
    )

    # =====================================================
    # ARRANQUE
    # =====================================================
    arranque = forms.alert(
        "Lançar Arranque?",
        yes=True,
        no=True
    )

    arranque_comp = "0"

    if arranque:

        arranque_comp = forms.ask_for_string(
            default="60",
            prompt="Comprimento do Arranque (cm):",
            title="Arranque"
        )

    # =====================================================
    # EMBASAMENTO
    # =====================================================
    embasamento = forms.alert(
        "Lançar Embasamento?",
        yes=True,
        no=True
    )

    embasamento_comp = "0"

    if embasamento:

        embasamento_comp = forms.ask_for_string(
            default="100",
            prompt="Comprimento do Embasamento (cm):",
            title="Embasamento"
        )

    # =====================================================
    # DOBRA
    # =====================================================
    dobra_comp = forms.ask_for_string(
        default="20",
        prompt="Comprimento da dobra (cm):",
        title="Dobra"
    )

    # =====================================================
    # TOPO
    # =====================================================
    topo_comp = forms.ask_for_string(
        default="0",
        prompt="Comprimento do topo (cm):",
        title="Topo"
    )

    # =====================================================
    # BASE
    # =====================================================
    base_comp = forms.ask_for_string(
        default="0",
        prompt="Comprimento da base (cm):",
        title="Base"
    )

    # =====================================================
    # COBRIMENTO
    # =====================================================
    cobrimento = forms.ask_for_string(
        default="3",
        prompt="Cobrimento do concreto (cm):",
        title="Cobrimento"
    )

    return {

        "rebar_type":
            rebar_types[selected_rebar],

        "arranque":
            arranque,

        "arranque_comp":
            float(arranque_comp),

        "embasamento":
            embasamento,

        "embasamento_comp":
            float(embasamento_comp),

        "dobra_comp":
            float(dobra_comp),

        "espacamento":
            float(espacamento),

        "cobrimento":
            float(cobrimento),

        "topo_comp":
            float(topo_comp),

        "base_comp":
            float(base_comp)
    }


# =========================================================
# DADOS DA PAREDE
# =========================================================
def get_wall_data(wall):

    loc_curve = wall.Location.Curve

    start = loc_curve.GetEndPoint(0)
    end = loc_curve.GetEndPoint(1)

    direction = (
        end - start
    ).Normalize()

    width = wall.Width

    bbox = wall.get_BoundingBox(None)

    base_z = bbox.Min.Z
    top_z = bbox.Max.Z

    height = top_z - base_z

    length = loc_curve.Length

    return {

        "start":
            start,

        "end":
            end,

        "direction":
            direction,

        "width":
            width,

        "height":
            height,

        "length":
            length,

        "base_z":
            base_z,

        "top_z":
            top_z
    }


# =========================================================
# GERAR ARMADURAS
# =========================================================
def create_wall_rebars(wall, config):

    wall_data = get_wall_data(wall)

    spacing = cm_to_feet(
        config["espacamento"]
    )

    cover = cm_to_feet(
        config["cobrimento"]
    )

    hook = cm_to_feet(
        config["dobra_comp"]
    )

    topo = cm_to_feet(
        config["topo_comp"]
    )

    base = cm_to_feet(
        config["base_comp"]
    )

    arranque = cm_to_feet(
        config["arranque_comp"]
    )

    emb = cm_to_feet(
        config["embasamento_comp"]
    )

    rebar_type = config["rebar_type"]

    direction = wall_data["direction"]

    normal = XYZ.BasisZ.CrossProduct(
        direction
    )

    usable_length = (
        wall_data["length"] - (cover * 2)
    )

    qty = int(
        usable_length / spacing
    ) + 1

    rebars = []

    for i in range(qty):

        dist = cover + (i * spacing)

        if dist > usable_length:
            break

        point = (
            wall_data["start"] +
            (direction * dist)
        )

        x = point.X
        y = point.Y

        z1 = wall_data["base_z"]
        z2 = wall_data["top_z"]

        # =================================================
        # ARRANQUE
        # =================================================
        if config["arranque"]:
            z1 -= arranque

        # =================================================
        # TOPO
        # =================================================
        z2 += topo

        # =================================================
        # PONTOS
        # =================================================
        p1 = XYZ(x, y, z1)
        p2 = XYZ(x, y, z2)

        curves = []

        # =================================================
        # EMBASAMENTO
        # =================================================
        if config["embasamento"]:

            emb_point = (
                p1 - (direction * emb)
            )

            emb_curve = Line.CreateBound(
                emb_point,
                p1
            )

            curves.append(emb_curve)

        # =================================================
        # LINHA VERTICAL
        # =================================================
        vertical_curve = Line.CreateBound(
            p1,
            p2
        )

        curves.append(vertical_curve)

        # =================================================
        # DOBRA SUPERIOR
        # =================================================
        if hook > 0:

            hook_point = (
                p2 + (direction * hook)
            )

            hook_curve = Line.CreateBound(
                p2,
                hook_point
            )

            curves.append(hook_curve)

        # =================================================
        # BASE
        # =================================================
        if base > 0:

            base_point = (
                p1 - (direction * base)
            )

            base_curve = Line.CreateBound(
                base_point,
                p1
            )

            curves.insert(0, base_curve)

        # =================================================
        # CRIAR VERGALHÃO
        # =================================================
        try:

            rebar = Rebar.CreateFromCurves(
                doc,
                RebarStyle.Standard,
                rebar_type,
                None,
                None,
                wall,
                normal,
                curves,
                RebarHookOrientation.Left,
                RebarHookOrientation.Right,
                True,
                True
            )

            rebars.append(rebar)

        except Exception as ex:

            print(
                "Erro ao criar vergalhão:"
            )

            print(ex)

    return rebars


# =========================================================
# PREVIEW
# =========================================================
def show_preview_message(config):

    msg = []

    msg.append(
        "CONFIGURAÇÕES DO PALITEIRO"
    )

    msg.append("")

    msg.append(
        "Espaçamento: {} cm".format(
            config["espacamento"]
        )
    )

    msg.append(
        "Dobra: {} cm".format(
            config["dobra_comp"]
        )
    )

    if config["arranque"]:

        msg.append(
            "Arranque: {} cm".format(
                config["arranque_comp"]
            )
        )

    if config["embasamento"]:

        msg.append(
            "Embasamento: {} cm".format(
                config["embasamento_comp"]
            )
        )

    forms.alert(
        "\n".join(msg)
    )


# =========================================================
# MAIN
# =========================================================
def main():

    view = doc.ActiveView

    if (
        not isinstance(view, ViewPlan)
        and
        not isinstance(view, View3D)
    ):

        forms.alert(
            "Execute em planta ou 3D.",
            exitscript=True
        )

    config = get_user_inputs()

    if not config:
        return

    show_preview_message(config)

    # =====================================================
    # SELEÇÃO
    # =====================================================
    try:

        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallSelectionFilter(),
            "Selecione as paredes"
        )

    except:
        return

    walls = [
        doc.GetElement(r.ElementId)
        for r in refs
    ]

    if not walls:

        forms.alert(
            "Nenhuma parede selecionada."
        )

        return

    created = 0

    # =====================================================
    # TRANSACTION
    # =====================================================
    t = Transaction(
        doc,
        "Gerar Paliteiro"
    )

    t.Start()

    try:

        for wall in walls:

            rebars = create_wall_rebars(
                wall,
                config
            )

            created += len(rebars)

        t.Commit()

    except Exception as ex:

        print(ex)

        t.RollBack()

        forms.alert(
            "Erro durante a criação."
        )

        return

    # =====================================================
    # FINAL
    # =====================================================
    forms.alert(
        "Paliteiro gerado!\n\n"
        "{} vergalhões criados.".format(
            created
        )
    )


# =========================================================
# EXECUÇÃO
# =========================================================
if __name__ == "__main__":
    main()
