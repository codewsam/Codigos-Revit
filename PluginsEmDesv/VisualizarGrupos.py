# -*- coding: utf-8 -*-
"""
Plugin: Visualizar Grupos de Paredes Espelhadas
Extensão: Samuel PLUGIN
Versão: 1.0.0
Autor: Samuel

Descrição:
    Lista os grupos salvos, o usuário escolhe um,
    e as paredes do grupo são isoladas temporariamente
    na view ativa (Temporary Hide/Isolate nativo do Revit).
    Edições feitas durante o isolamento persistem no modelo.
    Botão "Resetar View" encerra o isolamento.
"""

# ─────────────────────────────────────────────
#  IMPORTS
# ─────────────────────────────────────────────
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, HorizontalAlignment,
    VerticalAlignment, MessageBox, MessageBoxButton, MessageBoxImage
)
from System.Windows.Controls import (
    StackPanel, ListBox, ListBoxItem, Button,
    Label, ScrollViewer, Grid, ColumnDefinition,
    RowDefinition, TextBlock, Separator
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import GridLength, GridUnitType

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB.ExtensibleStorage import Schema, Entity, DataStorage

from pyrevit import forms

# ─────────────────────────────────────────────
#  CONSTANTES — idênticas ao plugin de criação
# ─────────────────────────────────────────────
SCHEMA_GUID      = System.Guid("C7E3A912-4F5B-4D8E-9A1C-2B6D0F3E8C5A")
DATASTORAGE_NAME = "SamuelPlugin_GruposEspelhados"

FIELD_GRUPO_ID = "grupo_id"
FIELD_WALL_IDS = "wall_ids"
FIELD_DATA     = "data_criacao"
FIELD_NOME     = "nome_grupo"
FIELD_VERSAO   = "versao_schema"


# ─────────────────────────────────────────────
#  LEITURA DOS GRUPOS (Extensible Storage)
# ─────────────────────────────────────────────

def deserializar_wall_ids(wall_ids_str):
    """Converte string CSV de volta para lista de ElementId."""
    if not wall_ids_str:
        return []
    return [
        DB.ElementId(int(id_str))
        for id_str in wall_ids_str.split(",")
        if id_str.strip()
    ]


def carregar_grupos(doc):
    """
    Lê todos os grupos de paredes espelhadas salvos no documento.

    Retorna:
        list[dict]: cada dict tem grupo_id, wall_ids, data_criacao, nome_grupo
    """
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        return []

    collector = (
        DB.FilteredElementCollector(doc)
        .OfClass(DataStorage)
        .ToElements()
    )

    grupos = []
    for ds in collector:
        if not ds.Name.startswith(DATASTORAGE_NAME + "__"):
            continue

        entity = ds.GetEntity(schema)
        if not entity.IsValid():
            continue

        grupos.append({
            "grupo_id":     entity.Get[System.String](FIELD_GRUPO_ID),
            "wall_ids":     deserializar_wall_ids(entity.Get[System.String](FIELD_WALL_IDS)),
            "data_criacao": entity.Get[System.String](FIELD_DATA),
            "nome_grupo":   entity.Get[System.String](FIELD_NOME),
        })

    return grupos


def filtrar_paredes_validas(doc, wall_ids):
    """
    Remove da lista ElementIds que não existem mais no modelo.
    Evita erro ao tentar isolar elemento deletado.

    Retorna:
        list[DB.ElementId]: apenas os ids válidos
    """
    validos = []
    for wid in wall_ids:
        el = doc.GetElement(wid)
        if el is not None and isinstance(el, DB.Wall):
            validos.append(wid)
    return validos


# ─────────────────────────────────────────────
#  ISOLAMENTO TEMPORÁRIO
# ─────────────────────────────────────────────

def isolar_paredes(doc, view, wall_ids):
    """
    Aplica Temporary Hide/Isolate nas paredes do grupo.

    Usa IsolateElementsTemporary — o modo nativo do Revit
    que exibe a barra azul no rodapé da view.
    Tudo que não está na lista fica oculto temporariamente.
    Nenhuma modificação permanente é feita na view ou no modelo.

    Args:
        doc:      Document ativo
        view:     View ativa (deve suportar isolamento)
        wall_ids: list[DB.ElementId] das paredes a isolar
    """
    ids_para_isolar = List[DB.ElementId](wall_ids)

    t = DB.Transaction(doc, "Isolar Grupo Espelhado")
    t.Start()
    try:
        view.IsolateElementsTemporary(ids_para_isolar)
        t.Commit()
    except Exception as ex:
        t.RollbackIfOpen()
        raise ex


def resetar_isolamento(doc, view):
    """
    Encerra o Temporary Hide/Isolate da view ativa.
    Equivale a clicar em "Reset Temporary Hide/Isolate"
    na barra azul do Revit.

    Args:
        doc:  Document ativo
        view: View ativa
    """
    t = DB.Transaction(doc, "Resetar Isolamento")
    t.Start()
    try:
        view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
        t.Commit()
    except Exception as ex:
        t.RollbackIfOpen()
        raise ex


# ─────────────────────────────────────────────
#  INTERFACE WPF
# ─────────────────────────────────────────────

class JanelaVisualizarGrupos(Window):
    """
    Janela WPF para listagem e seleção de grupos espelhados.

    Layout:
    ┌─────────────────────────────────────┐
    │  Grupos de Paredes Espelhadas       │
    ├─────────────────────────────────────┤
    │  [ListBox com grupos]               │
    │   • Nome do grupo — X paredes       │
    │   • Nome do grupo — X paredes       │
    ├─────────────────────────────────────┤
    │  [Isolar Grupo]   [Resetar View]    │
    └─────────────────────────────────────┘
    """

    def __init__(self, doc, view, grupos):
        self.doc    = doc
        self.view   = view
        self.grupos = grupos

        self._construir_ui()

    def _construir_ui(self):
        # ── Janela principal
        self.Title          = "Grupos de Paredes Espelhadas"
        self.Width          = 420
        self.Height         = 480
        self.ResizeMode     = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background     = SolidColorBrush(Color.FromRgb(245, 245, 245))

        # ── Layout raiz
        root = StackPanel()
        root.Margin = Thickness(16)
        self.Content = root

        # ── Título
        titulo = TextBlock()
        titulo.Text       = "Selecione um grupo para isolar na view"
        titulo.FontSize   = 13
        titulo.FontWeight = System.Windows.FontWeights.SemiBold
        titulo.Margin     = Thickness(0, 0, 0, 10)
        root.Children.Add(titulo)

        # ── ListBox de grupos
        scroll = ScrollViewer()
        scroll.Height           = 280
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.Background       = SolidColorBrush(Color.FromRgb(255, 255, 255))

        self.listbox = ListBox()
        self.listbox.Margin = Thickness(0)

        if not self.grupos:
            item = ListBoxItem()
            item.Content   = "Nenhum grupo encontrado no modelo."
            item.IsEnabled = False
            self.listbox.Items.Add(item)
        else:
            for grupo in self.grupos:
                item          = ListBoxItem()
                item.Tag      = grupo
                item.Padding  = Thickness(8, 6, 8, 6)

                painel = StackPanel()

                nome_tb           = TextBlock()
                nome_tb.Text      = grupo["nome_grupo"] or ("Grupo_" + grupo["grupo_id"][:8])
                nome_tb.FontSize  = 12
                nome_tb.FontWeight = System.Windows.FontWeights.Medium

                info_tb          = TextBlock()
                info_tb.Text     = "{} paredes  •  criado em {}".format(
                    len(grupo["wall_ids"]),
                    grupo["data_criacao"] or "—"
                )
                info_tb.FontSize = 10
                info_tb.Foreground = SolidColorBrush(Color.FromRgb(120, 120, 120))
                info_tb.Margin   = Thickness(0, 2, 0, 0)

                painel.Children.Add(nome_tb)
                painel.Children.Add(info_tb)
                item.Content = painel
                self.listbox.Items.Add(item)

        scroll.Content = self.listbox
        root.Children.Add(scroll)

        # ── Separador
        sep = Separator()
        sep.Margin = Thickness(0, 14, 0, 14)
        root.Children.Add(sep)

        # ── Botões
        painel_botoes = StackPanel()
        painel_botoes.Orientation = System.Windows.Controls.Orientation.Horizontal
        painel_botoes.HorizontalAlignment = HorizontalAlignment.Stretch

        btn_isolar          = Button()
        btn_isolar.Content  = "🔍  Isolar Grupo"
        btn_isolar.Width    = 170
        btn_isolar.Height   = 36
        btn_isolar.FontSize = 12
        btn_isolar.Margin   = Thickness(0, 0, 12, 0)
        btn_isolar.Background  = SolidColorBrush(Color.FromRgb(0, 120, 212))
        btn_isolar.Foreground  = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_isolar.Click    += self._ao_isolar

        btn_reset           = Button()
        btn_reset.Content   = "↩  Resetar View"
        btn_reset.Width     = 170
        btn_reset.Height    = 36
        btn_reset.FontSize  = 12
        btn_reset.Background   = SolidColorBrush(Color.FromRgb(232, 17, 35))
        btn_reset.Foreground   = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_reset.Click     += self._ao_resetar

        painel_botoes.Children.Add(btn_isolar)
        painel_botoes.Children.Add(btn_reset)
        root.Children.Add(painel_botoes)

        # ── Rodapé informativo
        rodape          = TextBlock()
        rodape.Text     = "As edições feitas no isolamento são salvas no projeto normalmente."
        rodape.FontSize = 10
        rodape.Foreground = SolidColorBrush(Color.FromRgb(140, 140, 140))
        rodape.TextWrapping = System.Windows.TextWrapping.Wrap
        rodape.Margin   = Thickness(0, 12, 0, 0)
        root.Children.Add(rodape)

    # ── Handlers ──────────────────────────────

    def _ao_isolar(self, sender, args):
        item_selecionado = self.listbox.SelectedItem
        if item_selecionado is None:
            MessageBox.Show(
                "Selecione um grupo na lista antes de isolar.",
                "Aviso",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )
            return

        grupo    = item_selecionado.Tag
        wall_ids = filtrar_paredes_validas(self.doc, grupo["wall_ids"])

        if not wall_ids:
            MessageBox.Show(
                "Nenhuma parede deste grupo existe mais no modelo.\n"
                "O grupo pode estar desatualizado.",
                "Grupo Inválido",
                MessageBoxButton.OK,
                MessageBoxImage.Warning
            )
            return

        removidas = len(grupo["wall_ids"]) - len(wall_ids)
        aviso_removidas = ""
        if removidas > 0:
            aviso_removidas = "\n\nAtenção: {} parede(s) foram deletadas do modelo e serão ignoradas.".format(removidas)

        try:
            isolar_paredes(self.doc, self.view, wall_ids)
        except Exception as ex:
            MessageBox.Show(
                "Erro ao isolar paredes:\n\n{}".format(str(ex)),
                "Erro",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            )
            return

        nome = grupo["nome_grupo"] or ("Grupo_" + grupo["grupo_id"][:8])
        MessageBox.Show(
            "{} parede(s) isoladas.\nGrupo: {}{}".format(
                len(wall_ids), nome, aviso_removidas
            ),
            "Isolamento Ativo",
            MessageBoxButton.OK,
            MessageBoxImage.Information
        )
        self.Close()

    def _ao_resetar(self, sender, args):
        try:
            resetar_isolamento(self.doc, self.view)
            MessageBox.Show(
                "View restaurada. Todas as paredes estão visíveis novamente.",
                "View Resetada",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
        except Exception as ex:
            MessageBox.Show(
                "Erro ao resetar a view:\n\n{}".format(str(ex)),
                "Erro",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            )
        self.Close()


# ─────────────────────────────────────────────
#  PONTO DE ENTRADA
# ─────────────────────────────────────────────

def main():
    doc   = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    view  = uidoc.ActiveView

    # Verificar se a view suporta isolamento temporário
    if view.ViewType not in [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.Elevation,
        DB.ViewType.Section,
        DB.ViewType.ThreeD,
        DB.ViewType.Detail,
    ]:
        forms.alert(
            "A view ativa não suporta isolamento temporário.\n"
            "Abra uma planta, corte, elevação ou vista 3D.",
            title="Visualizar Grupos"
        )
        return

    grupos = carregar_grupos(doc)

    janela = JanelaVisualizarGrupos(doc, view, grupos)
    janela.ShowDialog()


if __name__ == '__main__':
    main()
