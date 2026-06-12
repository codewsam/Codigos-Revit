# -*- coding: utf-8 -*-
__title__ = "Vergalhoes"
__version__ = "2.4"
__doc__ = "Coloca varios vergalhões nas paredes selecionadas."

"""
PLUGIN: Paliteiro
VERSAO: 2.4
COMPATIBILIDADE: Revit 2024+
"""

# =========================================================
# IMPORTS
# =========================================================
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, FamilyInstance, Line, XYZ,
    BuiltInCategory, Transaction, ViewPlan, View3D
)
from Autodesk.Revit.DB.Structure import (
    RebarBarType, Rebar, RebarStyle, RebarHookOrientation
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from pyrevit import forms

from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
import System.Windows


# =========================================================
# DOCUMENTO
# =========================================================
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document


# =========================================================
# UTILITARIOS
# =========================================================
def cm_to_feet(cm):
    return float(cm) / 30.48


# =========================================================
# TIPOS DE VERGALHÃO
# =========================================================
def get_rebar_types():
    types = {}
    for r in FilteredElementCollector(doc).OfClass(RebarBarType):
        try:
            p    = r.LookupParameter("Type Name")
            name = p.AsString() if p else "Vergalhao_{}".format(r.Id.IntegerValue)
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
# ZONAS DE ABERTURA (portas + janelas)
# =========================================================
# =========================================================
# ZONAS DE ABERTURA (portas + janelas)
# =========================================================
def get_opening_zones(wall, wall_start, wall_direction, tolerance_cm=2.0):
    tolerance = cm_to_feet(tolerance_cm)
    zones     = []
    wall_id   = wall.Id

    for category in [BuiltInCategory.OST_Doors, BuiltInCategory.OST_Windows]:
        openings = (
            FilteredElementCollector(doc)
            .OfCategory(category)
            .OfClass(FamilyInstance)
            .ToElements()
        )
        for opening in openings:
            host = opening.Host
            if host is None or host.Id != wall_id:
                continue
            bb = opening.get_BoundingBox(None)
            if bb is None:
                continue
            corners = [
                XYZ(bb.Min.X, bb.Min.Y, 0),
                XYZ(bb.Max.X, bb.Min.Y, 0),
                XYZ(bb.Min.X, bb.Max.Y, 0),
                XYZ(bb.Max.X, bb.Max.Y, 0),
            ]
            projs = [(c - wall_start).DotProduct(wall_direction) for c in corners]
            zones.append((
                min(projs) - tolerance,  # x_min
                max(projs) + tolerance,  # x_max
                bb.Min.Z,                # z_min da abertura
                bb.Max.Z                 # z_max da abertura
            ))

    return zones


def is_inside_opening(position, rebar_z1, rebar_z2, opening_zones):
    for (x_min, x_max, z_min, z_max) in opening_zones:
        # Conflito horizontal E conflito vertical
        if x_min <= position <= x_max and rebar_z1 < z_max and rebar_z2 > z_min:
            return True
    return False


# =========================================================
# DADOS DA PAREDE
# =========================================================
def get_wall_data(wall):
    loc_curve = wall.Location.Curve
    start     = loc_curve.GetEndPoint(0)
    end       = loc_curve.GetEndPoint(1)
    direction = (end - start).Normalize()
    bbox      = wall.get_BoundingBox(None)
    return {
        "start":     start,
        "end":       end,
        "direction": direction,
        "width":     wall.Width,
        "length":    loc_curve.Length,
        "base_z":    bbox.Min.Z,
        "top_z":     bbox.Max.Z,
    }


# =========================================================
# GERAR ARMADURAS
# =========================================================
def create_wall_rebars(wall, config):
    wd          = get_wall_data(wall)
    spacing     = cm_to_feet(config["espacamento"])
    cover       = cm_to_feet(config["cobrimento"])
    hook        = cm_to_feet(config["dobra_comp"])
    base_ext    = cm_to_feet(config["base_comp"])
    altura      = cm_to_feet(config["altura"])
    embutimento = cm_to_feet(config["embutimento"])

    direction = wd["direction"]
    normal    = XYZ.BasisZ.CrossProduct(direction)
    usable    = wd["length"] - (cover * 2)
    qty       = int(usable / spacing) + 1

    opening_zones = get_opening_zones(wall, wd["start"], direction)

    rebars  = []
    skipped = 0

    for i in range(qty):
        dist = cover + (i * spacing)
        if dist > usable:
            break

        pt   = wd["start"] + (direction * dist)
        x, y = pt.X, pt.Y

        # z1 e z2 PRIMEIRO
        z1 = wd["base_z"] - embutimento
        z2 = wd["base_z"] + altura

        # DEPOIS a checagem
        if is_inside_opening(dist, z1, z2, opening_zones):
            skipped += 1
            continue

        p1 = XYZ(x, y, z1)
        p2 = XYZ(x, y, z2)

        curves = [Line.CreateBound(p1, p2)]
    # ... resto

        if hook > 0:
            curves.append(Line.CreateBound(p2, p2 + (direction * hook)))

        if base_ext > 0:
            curves.insert(0, Line.CreateBound(p1 - (direction * base_ext), p1))

        try:
            rebar = Rebar.CreateFromCurves(
                doc,
                RebarStyle.Standard,
                config["rebar_type"],
                None, None,
                wall,
                normal,
                curves,
                RebarHookOrientation.Left,
                RebarHookOrientation.Right,
                True, True
            )
            rebars.append(rebar)
        except Exception as ex:
            print("Erro vergalhao pos {:.2f}: {}".format(dist, ex))

    return rebars, skipped


# =========================================================
# XAML
# =========================================================
XAML = u"""
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Paliteiro - Inserir Armadura"
    Width="400" SizeToContent="Height"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    ShowInTaskbar="False"
    Background="#F4F4F4"
    FontFamily="Segoe UI"
    FontSize="12">

  <Border Padding="20">
    <StackPanel>

      <!-- VERGALHAO -->
      <TextBlock Text="VERGALHAO"
                 FontWeight="SemiBold" FontSize="11"
                 Foreground="#555" Margin="0,0,0,4"/>
      <ComboBox x:Name="cmbRebar"
                Height="26" Background="White" BorderBrush="#BBBBBB"/>

      <!-- ESPACAMENTO -->
      <TextBlock Text="Espacamento maximo (cm)"
                 Foreground="#222" Margin="0,10,0,2"/>
      <TextBox x:Name="txtEspacamento" Text="24"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <!-- COBRIMENTO -->
      <TextBlock Text="Cobrimento (cm)"
                 Foreground="#222" Margin="0,8,0,2"/>
      <TextBox x:Name="txtCobrimento" Text="3"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <!-- ALTURA DO VERGALHAO -->
      <TextBlock Text="ALTURA"
                 FontWeight="SemiBold" FontSize="11"
                 Foreground="#555" Margin="0,14,0,4"/>

      <TextBlock Text="Altura do vergalhao acima da laje (cm)"
                 Foreground="#222" Margin="0,0,0,2"/>
      <TextBox x:Name="txtAltura" Text="60"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <TextBlock Text="Embutimento na laje / abaixo da base (cm)"
                 Foreground="#222" Margin="0,8,0,2"/>
      <TextBox x:Name="txtEmbutimento" Text="5"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <!-- EXTENSOES -->
      <TextBlock Text="EXTENSOES"
                 FontWeight="SemiBold" FontSize="11"
                 Foreground="#555" Margin="0,14,0,4"/>

      <TextBlock Text="Dobra no topo (cm)" Foreground="#222" Margin="0,0,0,2"/>
      <TextBox x:Name="txtDobra" Text="20"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <TextBlock Text="Extensao na base (cm)" Foreground="#222" Margin="0,8,0,2"/>
      <TextBox x:Name="txtBase" Text="0"
               Height="26" Padding="5,3"
               BorderBrush="#BBBBBB" Background="White"
               VerticalContentAlignment="Center"/>

      <!-- AVISO ABERTURAS -->
      <Border Background="#E3F2FD" CornerRadius="4"
              Padding="10,8" Margin="0,16,0,4">
        <TextBlock TextWrapping="Wrap"
                   Foreground="#1565C0" FontSize="11"
                   Text="Vergalhoes sobre portas e janelas serao ignorados automaticamente."/>
      </Border>

      <!-- BOTOES -->
      <Grid Margin="0,12,0,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="10"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="btnOK"
                Grid.Column="0"
                Content="Selecionar Paredes e Gerar"
                Background="#1565C0" Foreground="White"
                FontWeight="SemiBold" BorderThickness="0"
                Height="32"/>
        <Button x:Name="btnCancel"
                Grid.Column="2"
                Content="Cancelar"
                Background="#E0E0E0" Foreground="#333"
                BorderThickness="0" Height="32"/>
      </Grid>

    </StackPanel>
  </Border>
</Window>
"""


# =========================================================
# JANELA WPF
# =========================================================
class PaliteiroForm(forms.WPFWindow):

    def __init__(self, rebar_names):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)

        self.result = None

        for name in rebar_names:
            self.cmbRebar.Items.Add(name)
        if rebar_names:
            self.cmbRebar.SelectedIndex = 0

        self.btnOK.Click     += self._on_ok
        self.btnCancel.Click += self._on_cancel

    def _float(self, textbox, default=0.0):
        try:
            return float(textbox.Text.replace(",", "."))
        except:
            return default

    def _on_ok(self, sender, args):
        if self.cmbRebar.SelectedItem is None:
            MessageBox.Show(
                "Selecione um tipo de vergalhao.",
                "Paliteiro",
                MessageBoxButton.OK,
                MessageBoxImage.Warning)
            return

        self.result = {
            "rebar_name":   self.cmbRebar.SelectedItem,
            "espacamento":  self._float(self.txtEspacamento, 24.0),
            "cobrimento":   self._float(self.txtCobrimento,  3.0),
            "altura":       self._float(self.txtAltura,      60.0),
            "embutimento":  self._float(self.txtEmbutimento,  5.0),
            "dobra_comp":   self._float(self.txtDobra,       20.0),
            "base_comp":    self._float(self.txtBase,         0.0),
        }
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# =========================================================
# MAIN
# =========================================================
def main():

    view = doc.ActiveView
    if not isinstance(view, ViewPlan) and not isinstance(view, View3D):
        forms.alert("Execute em planta ou 3D.", exitscript=True)

    rebar_map = get_rebar_types()
    if not rebar_map:
        forms.alert("Nenhum tipo de vergalhao encontrado no projeto.")
        return

    form = PaliteiroForm(sorted(rebar_map.keys()))
    form.show_dialog()

    if form.result is None:
        return

    config = form.result
    config["rebar_type"] = rebar_map[config["rebar_name"]]

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
            total_created  += len(rebars)
            total_skipped  += skipped
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert("Erro durante a criacao:\n{}".format(ex))
        return

    forms.alert(
        "Paliteiro gerado!\n\n"
        "{} vergalhoes criados.\n"
        "{} posicoes ignoradas (portas/janelas).".format(
            total_created, total_skipped)
    )


# =========================================================
# EXECUCAO
# =========================================================
if __name__ == "__main__":
    main()
