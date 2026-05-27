# -*- coding: utf-8 -*-
"""
PLUGIN: Paliteiro
VERSAO: 2.1
COMPATIBILIDADE: Revit 2024+

CORREÇÃO v2.1:
- Vergalhões NÃO são gerados dentro de aberturas
  de portas e janelas.
- A lógica projeta cada abertura na direção da parede
  e exclui posições que caem dentro do intervalo da abertura.
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
                name = "Vergalhao_{}".format(r.Id.IntegerValue)
        except:
            name = "Vergalhao_{}".format(r.Id.IntegerValue)

        types[name] = r

    return types


# =========================================================
# FILTRO DE PAREDE
# =========================================================
class WallSelectionFilter(ISelectionFilter):

    def AllowElement(self, elem):
        return isinstance(elem, Wall)

    def AllowReference(self, ref, point):
        return True


# =========================================================
# ABERTURAS DA PAREDE
# Retorna lista de intervalos [inicio, fim] em feet
# projetados na direção da parede (distância a partir do start).
# Cada intervalo representa uma zona proibida para vergalhão.
# =========================================================
def get_opening_zones(wall, wall_start, wall_direction, tolerance_cm=5.0):
    """
    Coleta todas as portas e janelas hospedadas nesta parede
    e calcula, para cada uma, o intervalo [proj_min, proj_max]
    ao longo do eixo da parede.

    tolerance_cm: margem extra em cada lado da abertura (cm).
    Vergalhões dentro de [proj_min - tol, proj_max + tol] são ignorados.
    """

    tolerance = cm_to_feet(tolerance_cm)
    zones = []

    # Categorias que representam aberturas
    opening_categories = [
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
    ]

    wall_id = wall.Id

    for category in opening_categories:

        openings = (
            FilteredElementCollector(doc)
            .OfCategory(category)
            .OfClass(FamilyInstance)
            .ToElements()
        )

        for opening in openings:

            # Verifica se está hospedado nesta parede
            host = opening.Host
            if host is None or host.Id != wall_id:
                continue

            bb = opening.get_BoundingBox(None)
            if bb is None:
                continue

            # Projeta os 4 cantos da bounding box (X/Y) na direção da parede
            # Usamos os 4 cantos do retângulo horizontal do bb
            corners = [
                XYZ(bb.Min.X, bb.Min.Y, 0),
                XYZ(bb.Max.X, bb.Min.Y, 0),
                XYZ(bb.Min.X, bb.Max.Y, 0),
                XYZ(bb.Max.X, bb.Max.Y, 0),
            ]

            projections = [
                (corner - wall_start).DotProduct(wall_direction)
                for corner in corners
            ]

            proj_min = min(projections) - tolerance
            proj_max = max(projections) + tolerance

            zones.append((proj_min, proj_max))

    return zones


def is_inside_opening(position, opening_zones):
    """
    Retorna True se a posição (distância ao longo da parede em feet)
    cair dentro de qualquer zona de abertura.
    """
    for (zone_min, zone_max) in opening_zones:
        if zone_min <= position <= zone_max:
            return True
    return False


# =========================================================
# INTERFACE
# =========================================================
def get_user_inputs():

    rebar_types = get_rebar_types()

    if not rebar_types:
        forms.alert("Nenhum tipo de vergalhão encontrado.")
        return None

    selected_rebar = forms.SelectFromList.show(
        sorted(rebar_types.keys()),
        title="Tipo de Vergalhão",
        multiselect=False
    )

    if not selected_rebar:
        return None

    espacamento = forms.ask_for_string(
        default="24",
        prompt="Espaçamento máximo (cm):",
        title="Paliteiro"
    )

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

    dobra_comp = forms.ask_for_string(
        default="20",
        prompt="Comprimento da dobra (cm):",
        title="Dobra"
    )

    topo_comp = forms.ask_for_string(
        default="0",
        prompt="Comprimento do topo (cm):",
        title="Topo"
    )

    base_comp = forms.ask_for_string(
        default="0",
        prompt="Comprimento da base (cm):",
        title="Base"
    )

    cobrimento = forms.ask_for_string(
        default="3",
        prompt="Cobrimento do concreto (cm):",
        title="Cobrimento"
    )

    return {
        "rebar_type":       rebar_types[selected_rebar],
        "arranque":         arranque,
        "arranque_comp":    float(arranque_comp),
        "embasamento":      embasamento,
        "embasamento_comp": float(embasamento_comp),
        "dobra_comp":       float(dobra_comp),
        "espacamento":      float(espacamento),
        "cobrimento":       float(cobrimento),
        "topo_comp":        float(topo_comp),
        "base_comp":        float(base_comp),
    }


# =========================================================
# DADOS DA PAREDE
# =========================================================
def get_wall_data(wall):

    loc_curve = wall.Location.Curve

    start = loc_curve.GetEndPoint(0)
    end   = loc_curve.GetEndPoint(1)

    direction = (end - start).Normalize()
    width     = wall.Width
    bbox      = wall.get_BoundingBox(None)
    base_z    = bbox.Min.Z
    top_z     = bbox.Max.Z
    height    = top_z - base_z
    length    = loc_curve.Length

    return {
        "start":     start,
        "end":       end,
        "direction": direction,
        "width":     width,
        "height":    height,
        "length":    length,
        "base_z":    base_z,
        "top_z":     top_z,
    }


# =========================================================
# GERAR ARMADURAS
# =========================================================
def create_wall_rebars(wall, config):

    wall_data = get_wall_data(wall)

    spacing  = cm_to_feet(config["espacamento"])
    cover    = cm_to_feet(config["cobrimento"])
    hook     = cm_to_feet(config["dobra_comp"])
    topo     = cm_to_feet(config["topo_comp"])
    base     = cm_to_feet(config["base_comp"])
    arranque = cm_to_feet(config["arranque_comp"])
    emb      = cm_to_feet(config["embasamento_comp"])

    rebar_type = config["rebar_type"]
    direction  = wall_data["direction"]

    normal = XYZ.BasisZ.CrossProduct(direction)

    usable_length = wall_data["length"] - (cover * 2)

    qty = int(usable_length / spacing) + 1

    # ---------------------------------------------------------
    # PRÉ-CALCULAR zonas de abertura (portas + janelas)
    # Uma única chamada por parede, antes do loop de vergalhões.
    # ---------------------------------------------------------
    opening_zones = get_opening_zones(
        wall,
        wall_data["start"],
        direction,
        tolerance_cm=2.0      # margem de 2 cm em cada lado da abertura
    )

    rebars   = []
    skipped  = 0

    for i in range(qty):

        dist = cover + (i * spacing)

        if dist > usable_length:
            break

        # ---------------------------------------------------------
        # VERIFICAÇÃO: esta posição cai dentro de uma abertura?
        # Se sim, pula sem criar o vergalhão.
        # ---------------------------------------------------------
        if is_inside_opening(dist, opening_zones):
            skipped += 1
            continue

        point = wall_data["start"] + (direction * dist)

        x = point.X
        y = point.Y

        z1 = wall_data["base_z"]
        z2 = wall_data["top_z"]

        if config["arranque"]:
            z1 -= arranque

        z2 += topo

        p1 = XYZ(x, y, z1)
        p2 = XYZ(x, y, z2)

        curves = []

        # Embasamento (extensão horizontal na base)
        if config["embasamento"]:
            emb_point = p1 - (direction * emb)
            curves.append(Line.CreateBound(emb_point, p1))

        # Linha vertical principal
        curves.append(Line.CreateBound(p1, p2))

        # Dobra no topo
        if hook > 0:
            hook_point = p2 + (direction * hook)
            curves.append(Line.CreateBound(p2, hook_point))

        # Extensão na base
        if base > 0:
            base_point = p1 - (direction * base)
            curves.insert(0, Line.CreateBound(base_point, p1))

        # Criar vergalhão
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
            print("Erro ao criar vergalhão na posição {:.2f} ft: {}".format(dist, ex))

    return rebars, skipped


# =========================================================
# PREVIEW
# =========================================================
def show_preview_message(config):

    msg = [
        "CONFIGURAÇÕES DO PALITEIRO",
        "",
        "Espaçamento: {} cm".format(config["espacamento"]),
        "Dobra: {} cm".format(config["dobra_comp"]),
    ]

    if config["arranque"]:
        msg.append("Arranque: {} cm".format(config["arranque_comp"]))

    if config["embasamento"]:
        msg.append("Embasamento: {} cm".format(config["embasamento_comp"]))

    msg.append("")
   

    forms.alert("\n".join(msg))


# =========================================================
# MAIN
# =========================================================
def main():

    view = doc.ActiveView

    if not isinstance(view, ViewPlan) and not isinstance(view, View3D):
        forms.alert("Execute em planta ou 3D.", exitscript=True)

    config = get_user_inputs()
    if not config:
        return

    show_preview_message(config)

    # Seleção de paredes
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            WallSelectionFilter(),
            "Selecione as paredes (ESC para cancelar)"
        )
    except:
        return

    walls = [doc.GetElement(r.ElementId) for r in refs]

    if not walls:
        forms.alert("Nenhuma parede selecionada.")
        return

    total_created = 0
    total_skipped = 0

    t = Transaction(doc, "Gerar Paliteiro")
    t.Start()

    try:
        for wall in walls:
            rebars, skipped = create_wall_rebars(wall, config)
            total_created += len(rebars)
            total_skipped += skipped

        t.Commit()

    except Exception as ex:
        print(ex)
        t.RollBack()
        forms.alert("Erro durante a criação:\n{}".format(ex))
        return

    # Resultado final
    msg = (
        "gerado com sucesso!\n\n"
        "{} vergalhões criados.\n"
    ).format(total_created, total_skipped)

    forms.alert(msg)


# =========================================================
# EXECUÇÃO
# =========================================================
if __name__ == "__main__":
    main()
