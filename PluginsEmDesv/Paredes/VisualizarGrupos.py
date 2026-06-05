# -*- coding: utf-8 -*-
__title__   = "Visualizar Grupos"
__author__  = "Samuel"
__version__ = "Versao 1.2"
"""
Plugin: Visualizar Grupos de Paredes Espelhadas
Extensão: Samuel PLUGIN
Versão: 1.2.0

Novidades v1.2:
    - Ao selecionar um grupo, lista automaticamente todas as
      aberturas (portas e janelas) hospedadas nas paredes do grupo
    - Permite editar largura, altura e deslocamento do peitoril
      de cada abertura diretamente pelo plugin
    - Ao isolar, inclui as aberturas no isolamento temporário
    - Aba "Aberturas" separada da aba de isolamento na UI
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Collections.Generic import List
from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    MessageBox, MessageBoxButton, MessageBoxImage, GridLength, GridUnitType
)
from System.Windows.Controls import (
    StackPanel, ListBox, ListBoxItem, Button, ScrollViewer,
    TextBlock, Separator, Grid, RowDefinition, ColumnDefinition,
    TabControl, TabItem, TextBox, Label, ComboBox, ComboBoxItem,
    GroupBox
)
from System.Windows.Media import SolidColorBrush, Color

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB.ExtensibleStorage import Schema, Entity, DataStorage

from pyrevit import forms, script

# ─────────────────────────────────────────────
#  CONSTANTES
# ─────────────────────────────────────────────
SCHEMA_GUID      = System.Guid("C7E3A912-4F5B-4D8E-9A1C-2B6D0F3E8C5A")
DATASTORAGE_NAME = "SamuelPlugin_GruposEspelhados"

FIELD_GRUPO_ID = "grupo_id"
FIELD_WALL_IDS = "wall_ids"
FIELD_DATA     = "data_criacao"
FIELD_NOME     = "nome_grupo"
FIELD_VERSAO   = "versao_schema"

FT_TO_CM = 30.48
CM_TO_FT = 1.0 / 30.48


# ─────────────────────────────────────────────
#  LEITURA DOS GRUPOS
# ─────────────────────────────────────────────

def deserializar_wall_ids(wall_ids_str):
    if not wall_ids_str:
        return []
    return [
        DB.ElementId(int(id_str))
        for id_str in wall_ids_str.split(",")
        if id_str.strip()
    ]


def carregar_grupos(doc):
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        return []

    collector = list(
        DB.FilteredElementCollector(doc)
        .OfClass(DataStorage)
        .ToElements()
    )

    grupos  = []
    prefixo = DATASTORAGE_NAME + "__"

    for ds in collector:
        if not ds.Name.startswith(prefixo):
            continue

        entity = ds.GetEntity(schema)
        if not entity.IsValid():
            continue

        try:
            grupo = {
                "grupo_id":     entity.Get[System.String](FIELD_GRUPO_ID),
                "wall_ids":     deserializar_wall_ids(entity.Get[System.String](FIELD_WALL_IDS)),
                "data_criacao": entity.Get[System.String](FIELD_DATA),
                "nome_grupo":   entity.Get[System.String](FIELD_NOME),
            }
            grupos.append(grupo)
        except Exception:
            continue

    return grupos


def filtrar_paredes_validas(doc, wall_ids):
    return [
        wid for wid in wall_ids
        if doc.GetElement(wid) is not None and isinstance(doc.GetElement(wid), DB.Wall)
    ]


# ─────────────────────────────────────────────
#  ABERTURAS (PORTAS / JANELAS)
# ─────────────────────────────────────────────

def coletar_aberturas_do_grupo(doc, wall_ids):
    ids_paredes = set(wid.IntegerValue for wid in wall_ids)
    aberturas   = []

    cats_alvo = [
        DB.BuiltInCategory.OST_Doors,
        DB.BuiltInCategory.OST_Windows,
    ]

    for bic in cats_alvo:
        elementos = (
            DB.FilteredElementCollector(doc)
            .OfCategory(bic)
            .OfClass(DB.FamilyInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        for el in elementos:
            host = el.Host
            if host is None:
                continue
            if host.Id.IntegerValue not in ids_paredes:
                continue

            categoria = "Porta" if bic == DB.BuiltInCategory.OST_Doors else "Janela"

            tipo_el = doc.GetElement(el.GetTypeId())
            familia = ""
            tipo_nome = ""
            if tipo_el:
                p_fam  = tipo_el.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                p_tipo = tipo_el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                familia   = p_fam.AsString()  if p_fam  else ""
                tipo_nome = p_tipo.AsString() if p_tipo else ""

            nome_exib = "{}: {} - {}".format(categoria, familia, tipo_nome)

            largura_cm = None
            p_larg = tipo_el.get_Parameter(DB.BuiltInParameter.DOOR_WIDTH) if tipo_el else None
            if p_larg is None and tipo_el:
                p_larg = tipo_el.LookupParameter("Width") or tipo_el.LookupParameter("Largura")
            if p_larg and p_larg.StorageType == DB.StorageType.Double:
                largura_cm = round(p_larg.AsDouble() * FT_TO_CM, 1)

            altura_cm = None
            p_alt = tipo_el.get_Parameter(DB.BuiltInParameter.DOOR_HEIGHT) if tipo_el else None
            if p_alt is None and tipo_el:
                p_alt = tipo_el.LookupParameter("Height") or tipo_el.LookupParameter("Altura")
            if p_alt and p_alt.StorageType == DB.StorageType.Double:
                altura_cm = round(p_alt.AsDouble() * FT_TO_CM, 1)

            peitoril_cm = None
            p_sill = el.get_Parameter(DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
            if p_sill is None:
                p_sill = el.LookupParameter("Sill Height") or el.LookupParameter("Altura do Peitoril")
            if p_sill and p_sill.StorageType == DB.StorageType.Double:
                peitoril_cm = round(p_sill.AsDouble() * FT_TO_CM, 1)

            aberturas.append({
                "element":     el,
                "id":          el.Id,
                "nome":        nome_exib,
                "categoria":   categoria,
                "wall_id":     host.Id,
                "largura_cm":  largura_cm,
                "altura_cm":   altura_cm,
                "peitoril_cm": peitoril_cm,
                "tipo_el":     tipo_el,
            })

    return aberturas


def editar_abertura(doc, abertura, nova_largura_cm, nova_altura_cm, novo_peitoril_cm):
    el      = abertura["element"]
    tipo_el = abertura["tipo_el"]
    erros   = []

    t = DB.Transaction(doc, "Editar Abertura")
    t.Start()
    try:
        # Peitoril — parametro de instancia
        if novo_peitoril_cm is not None:
            p_sill = el.get_Parameter(DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM)
            if p_sill is None:
                p_sill = el.LookupParameter("Sill Height") or el.LookupParameter("Altura do Peitoril")
            if p_sill and not p_sill.IsReadOnly:
                p_sill.Set(novo_peitoril_cm * CM_TO_FT)
            else:
                erros.append("Peitoril: parametro nao encontrado ou somente leitura.")

        # Largura e Altura — duplica o tipo para nao afetar outras instancias
        if tipo_el and (nova_largura_cm is not None or nova_altura_cm is not None):
            larg_atual = abertura["largura_cm"]
            alt_atual  = abertura["altura_cm"]
            larg_mudou = (nova_largura_cm is not None and nova_largura_cm != larg_atual)
            alt_mudou  = (nova_altura_cm  is not None and nova_altura_cm  != alt_atual)

            if larg_mudou or alt_mudou:
                p_nome_tipo = tipo_el.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                nome_base   = p_nome_tipo.AsString() if p_nome_tipo else "Tipo"
                novo_tipo   = tipo_el.Duplicate(
                    "{}_Plugin_L{}_H{}".format(
                        nome_base,
                        int(nova_largura_cm or larg_atual or 0),
                        int(nova_altura_cm  or alt_atual  or 0)
                    )
                )

                if larg_mudou:
                    p_larg = novo_tipo.get_Parameter(DB.BuiltInParameter.DOOR_WIDTH)
                    if p_larg is None:
                        p_larg = novo_tipo.LookupParameter("Width") or novo_tipo.LookupParameter("Largura")
                    if p_larg and not p_larg.IsReadOnly:
                        p_larg.Set(nova_largura_cm * CM_TO_FT)
                    else:
                        erros.append("Largura: parametro nao encontrado ou somente leitura.")

                if alt_mudou:
                    p_alt = novo_tipo.get_Parameter(DB.BuiltInParameter.DOOR_HEIGHT)
                    if p_alt is None:
                        p_alt = novo_tipo.LookupParameter("Height") or novo_tipo.LookupParameter("Altura")
                    if p_alt and not p_alt.IsReadOnly:
                        p_alt.Set(nova_altura_cm * CM_TO_FT)
                    else:
                        erros.append("Altura: parametro nao encontrado ou somente leitura.")

                el.ChangeTypeId(novo_tipo.Id)

        t.Commit()
        if erros:
            return True, "Salvo com avisos:\n" + "\n".join(erros)
        return True, "Abertura atualizada com sucesso."

    except Exception as ex:
        t.RollbackIfOpen()
        return False, "Erro: {}".format(str(ex))


# ─────────────────────────────────────────────
#  ISOLAMENTO TEMPORARIO
# ─────────────────────────────────────────────

def isolar_grupo(doc, view, wall_ids, aberturas):
    ids = list(wall_ids)
    for ab in aberturas:
        ids.append(ab["id"])
    ids_lista = List[DB.ElementId](ids)

    t = DB.Transaction(doc, "Isolar Grupo Espelhado")
    t.Start()
    try:
        if view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
            view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
        view.IsolateElementsTemporary(ids_lista)
        t.Commit()
    except Exception as ex:
        t.RollbackIfOpen()
        raise ex


def resetar_isolamento(doc, view):
    if not view.IsInTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate):
        MessageBox.Show(
            "A view nao esta em modo de isolamento temporario.",
            "Informacao", MessageBoxButton.OK, MessageBoxImage.Information
        )
        return False

    t = DB.Transaction(doc, "Resetar Isolamento")
    t.Start()
    try:
        view.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)
        t.Commit()
        return True
    except Exception as ex:
        t.RollbackIfOpen()
        raise ex


# ─────────────────────────────────────────────
#  JANELA DE EDICAO DE ABERTURA
# ─────────────────────────────────────────────

class JanelaEditarAbertura(Window):

    def __init__(self, doc, abertura):
        self.doc      = doc
        self.abertura = abertura
        self.salvo    = False
        self._construir_ui()

    def _construir_ui(self):
        self.Title    = "Editar Abertura"
        self.Width    = 340
        self.Height   = 310
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(245, 245, 245))

        root        = StackPanel()
        root.Margin = Thickness(16)
        self.Content = root

        titulo            = TextBlock()
        titulo.Text       = self.abertura["nome"]
        titulo.FontSize   = 12
        titulo.FontWeight = System.Windows.FontWeights.SemiBold
        titulo.TextWrapping = System.Windows.TextWrapping.Wrap
        titulo.Margin     = Thickness(0, 0, 0, 14)
        root.Children.Add(titulo)

        root.Children.Add(self._label("Largura (cm):"))
        self.txt_largura           = TextBox()
        self.txt_largura.Height    = 28
        self.txt_largura.FontSize  = 12
        self.txt_largura.Text      = str(self.abertura["largura_cm"]) if self.abertura["largura_cm"] is not None else ""
        self.txt_largura.IsEnabled = self.abertura["largura_cm"] is not None
        self.txt_largura.Margin    = Thickness(0, 2, 0, 10)
        root.Children.Add(self.txt_largura)

        root.Children.Add(self._label("Altura (cm):"))
        self.txt_altura           = TextBox()
        self.txt_altura.Height    = 28
        self.txt_altura.FontSize  = 12
        self.txt_altura.Text      = str(self.abertura["altura_cm"]) if self.abertura["altura_cm"] is not None else ""
        self.txt_altura.IsEnabled = self.abertura["altura_cm"] is not None
        self.txt_altura.Margin    = Thickness(0, 2, 0, 10)
        root.Children.Add(self.txt_altura)

        root.Children.Add(self._label("Altura do Peitoril (cm):"))
        self.txt_peitoril           = TextBox()
        self.txt_peitoril.Height    = 28
        self.txt_peitoril.FontSize  = 12
        self.txt_peitoril.Text      = str(self.abertura["peitoril_cm"]) if self.abertura["peitoril_cm"] is not None else ""
        self.txt_peitoril.IsEnabled = self.abertura["peitoril_cm"] is not None
        self.txt_peitoril.Margin    = Thickness(0, 2, 0, 14)
        root.Children.Add(self.txt_peitoril)


        painel_btns             = StackPanel()
        painel_btns.Orientation = System.Windows.Controls.Orientation.Horizontal

        btn_salvar              = Button()
        btn_salvar.Content      = "Salvar"
        btn_salvar.Width        = 120
        btn_salvar.Height       = 32
        btn_salvar.FontSize     = 12
        btn_salvar.Margin       = Thickness(0, 0, 12, 0)
        btn_salvar.Background   = SolidColorBrush(Color.FromRgb(0, 120, 212))
        btn_salvar.Foreground   = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_salvar.Click       += self._ao_salvar

        btn_cancelar            = Button()
        btn_cancelar.Content    = "Cancelar"
        btn_cancelar.Width      = 120
        btn_cancelar.Height     = 32
        btn_cancelar.FontSize   = 12
        btn_cancelar.Click     += self._ao_cancelar

        painel_btns.Children.Add(btn_salvar)
        painel_btns.Children.Add(btn_cancelar)
        root.Children.Add(painel_btns)

    def _label(self, texto):
        lbl            = TextBlock()
        lbl.Text       = texto
        lbl.FontSize   = 11
        lbl.Foreground = SolidColorBrush(Color.FromRgb(60, 60, 60))
        return lbl

    def _parse_float(self, txt_box):
        txt = txt_box.Text.strip().replace(",", ".")
        if not txt:
            return None
        try:
            return float(txt)
        except Exception:
            return "INVALIDO"

    def _ao_salvar(self, sender, args):
        nova_largura  = self._parse_float(self.txt_largura)
        nova_altura   = self._parse_float(self.txt_altura)
        novo_peitoril = self._parse_float(self.txt_peitoril)

        for nome, val in [("Largura", nova_largura), ("Altura", nova_altura), ("Peitoril", novo_peitoril)]:
            if val == "INVALIDO":
                MessageBox.Show(
                    "{}: valor invalido. Use apenas numeros (ex: 90 ou 90.5).".format(nome),
                    "Erro de Validacao", MessageBoxButton.OK, MessageBoxImage.Warning
                )
                return
            if val is not None and val <= 0:
                MessageBox.Show(
                    "{}: o valor deve ser maior que zero.".format(nome),
                    "Erro de Validacao", MessageBoxButton.OK, MessageBoxImage.Warning
                )
                return

        sucesso, mensagem = editar_abertura(
            self.doc, self.abertura,
            nova_largura, nova_altura, novo_peitoril
        )

        icone = MessageBoxImage.Information if sucesso else MessageBoxImage.Error
        MessageBox.Show(mensagem, "Resultado", MessageBoxButton.OK, icone)

        if sucesso:
            self.salvo = True
            self.Close()

    def _ao_cancelar(self, sender, args):
        self.Close()


# ─────────────────────────────────────────────
#  JANELA PRINCIPAL
# ─────────────────────────────────────────────

class JanelaVisualizarGrupos(Window):

    def __init__(self, doc, view, grupos):
        self.doc              = doc
        self.view             = view
        self.grupos           = grupos
        self.aberturas_atuais = []
        self._construir_ui()

    def _construir_ui(self):
        self.Title    = "Grupos de Paredes Espelhadas"
        self.Width    = 460
        self.Height   = 560
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(245, 245, 245))

        root        = StackPanel()
        root.Margin = Thickness(12)
        self.Content = root

        self.tabs        = TabControl()
        self.tabs.Height = 500
        root.Children.Add(self.tabs)

        # ── ABA 1: GRUPOS ─────────────────────────────────────
        tab_grupos        = TabItem()
        tab_grupos.Header = "  Grupos  "
        tab_grupos.FontSize = 12

        painel_g        = StackPanel()
        painel_g.Margin = Thickness(10)

        titulo_g            = TextBlock()
        titulo_g.Text       = "Selecione um grupo:"
        titulo_g.FontSize   = 12
        titulo_g.FontWeight = System.Windows.FontWeights.SemiBold
        titulo_g.Margin     = Thickness(0, 6, 0, 8)
        painel_g.Children.Add(titulo_g)

        scroll_g = ScrollViewer()
        scroll_g.Height = 320
        scroll_g.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll_g.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))

        self.listbox_grupos                    = ListBox()
        self.listbox_grupos.Margin             = Thickness(0)
        self.listbox_grupos.SelectionChanged  += self._ao_mudar_grupo

        if not self.grupos:
            item           = ListBoxItem()
            item.Content   = "Nenhum grupo encontrado. Crie com 'Filtrar Paredes'."
            item.IsEnabled = False
            self.listbox_grupos.Items.Add(item)
        else:
            for idx, grupo in enumerate(self.grupos):
                item         = ListBoxItem()
                item.Tag     = idx
                item.Padding = Thickness(8, 6, 8, 6)

                painel_item = StackPanel()

                nome_tb             = TextBlock()
                nome_tb.Text        = grupo["nome_grupo"] or ("Grupo_" + grupo["grupo_id"][:8])
                nome_tb.FontSize    = 12
                nome_tb.FontWeight  = System.Windows.FontWeights.Medium

                info_tb             = TextBlock()
                info_tb.Text        = "{} paredes  -  criado em {}".format(
                    len(grupo["wall_ids"]),
                    grupo["data_criacao"] or "-"
                )
                info_tb.FontSize    = 10
                info_tb.Foreground  = SolidColorBrush(Color.FromRgb(120, 120, 120))
                info_tb.Margin      = Thickness(0, 2, 0, 0)

                painel_item.Children.Add(nome_tb)
                painel_item.Children.Add(info_tb)
                item.Content = painel_item
                self.listbox_grupos.Items.Add(item)

        scroll_g.Content = self.listbox_grupos
        painel_g.Children.Add(scroll_g)

        sep_g        = Separator()
        sep_g.Margin = Thickness(0, 12, 0, 12)
        painel_g.Children.Add(sep_g)

        painel_btn_g             = StackPanel()
        painel_btn_g.Orientation = System.Windows.Controls.Orientation.Horizontal

        btn_isolar               = Button()
        btn_isolar.Content       = "Isolar Grupo"
        btn_isolar.Width         = 160
        btn_isolar.Height        = 36
        btn_isolar.FontSize      = 12
        btn_isolar.Margin        = Thickness(0, 0, 12, 0)
        btn_isolar.Background    = SolidColorBrush(Color.FromRgb(0, 120, 212))
        btn_isolar.Foreground    = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_isolar.Click        += self._ao_isolar

        btn_reset                = Button()
        btn_reset.Content        = "Resetar View"
        btn_reset.Width          = 160
        btn_reset.Height         = 36
        btn_reset.FontSize       = 12
        btn_reset.Background     = SolidColorBrush(Color.FromRgb(232, 17, 35))
        btn_reset.Foreground     = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_reset.Click         += self._ao_resetar

        painel_btn_g.Children.Add(btn_isolar)
        painel_btn_g.Children.Add(btn_reset)
        painel_g.Children.Add(painel_btn_g)

        tab_grupos.Content = painel_g
        self.tabs.Items.Add(tab_grupos)

        # ── ABA 2: ABERTURAS ──────────────────────────────────
        tab_ab         = TabItem()
        tab_ab.Header  = "  Aberturas  "
        tab_ab.FontSize = 12

        painel_ab        = StackPanel()
        painel_ab.Margin = Thickness(10)

        self.lbl_grupo_ab              = TextBlock()
        self.lbl_grupo_ab.Text         = "Selecione um grupo na aba Grupos."
        self.lbl_grupo_ab.FontSize     = 12
        self.lbl_grupo_ab.FontWeight   = System.Windows.FontWeights.SemiBold
        self.lbl_grupo_ab.Margin       = Thickness(0, 6, 0, 8)
        self.lbl_grupo_ab.TextWrapping = System.Windows.TextWrapping.Wrap
        painel_ab.Children.Add(self.lbl_grupo_ab)

        scroll_ab = ScrollViewer()
        scroll_ab.Height = 390
        scroll_ab.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll_ab.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))

        self.painel_aberturas        = StackPanel()
        self.painel_aberturas.Margin = Thickness(4)

        placeholder            = TextBlock()
        placeholder.Text       = "Nenhuma abertura carregada."
        placeholder.FontSize   = 11
        placeholder.Foreground = SolidColorBrush(Color.FromRgb(150, 150, 150))
        placeholder.Margin     = Thickness(6)
        self.painel_aberturas.Children.Add(placeholder)

        scroll_ab.Content = self.painel_aberturas
        painel_ab.Children.Add(scroll_ab)

        tab_ab.Content = painel_ab
        self.tabs.Items.Add(tab_ab)

    # ── Helpers ───────────────────────────────

    def _grupo_selecionado(self):
        item = self.listbox_grupos.SelectedItem
        if item is None or item.Tag is None:
            return None
        return self.grupos[item.Tag]

    def _reconstruir_lista_aberturas(self):
        self.painel_aberturas.Children.Clear()

        if not self.aberturas_atuais:
            txt              = TextBlock()
            txt.Text         = "Nenhuma porta ou janela encontrada nas paredes deste grupo."
            txt.FontSize     = 11
            txt.Foreground   = SolidColorBrush(Color.FromRgb(150, 150, 150))
            txt.Margin       = Thickness(6)
            txt.TextWrapping = System.Windows.TextWrapping.Wrap
            self.painel_aberturas.Children.Add(txt)
            return

        for idx, ab in enumerate(self.aberturas_atuais):
            grid            = Grid()
            grid.Background = SolidColorBrush(Color.FromRgb(250, 250, 250))
            grid.Margin     = Thickness(0, 0, 0, 6)

            col1       = ColumnDefinition()
            col1.Width = GridLength(1, GridUnitType.Star)
            col2       = ColumnDefinition()
            col2.Width = GridLength(90)
            grid.ColumnDefinitions.Add(col1)
            grid.ColumnDefinitions.Add(col2)

            painel_info        = StackPanel()
            painel_info.Margin = Thickness(8, 6, 4, 6)

            nome_tb              = TextBlock()
            nome_tb.Text         = ab["nome"]
            nome_tb.FontSize     = 11
            nome_tb.FontWeight   = System.Windows.FontWeights.Medium
            nome_tb.TextWrapping = System.Windows.TextWrapping.Wrap

            dim_parts = []
            if ab["largura_cm"]  is not None: dim_parts.append("L: {} cm".format(ab["largura_cm"]))
            if ab["altura_cm"]   is not None: dim_parts.append("H: {} cm".format(ab["altura_cm"]))
            if ab["peitoril_cm"] is not None: dim_parts.append("P: {} cm".format(ab["peitoril_cm"]))

            dim_tb            = TextBlock()
            dim_tb.Text       = "  ".join(dim_parts) if dim_parts else "Dimensoes nao disponiveis"
            dim_tb.FontSize   = 10
            dim_tb.Foreground = SolidColorBrush(Color.FromRgb(100, 100, 100))
            dim_tb.Margin     = Thickness(0, 2, 0, 0)

            painel_info.Children.Add(nome_tb)
            painel_info.Children.Add(dim_tb)
            Grid.SetColumn(painel_info, 0)
            grid.Children.Add(painel_info)

            btn_editar              = Button()
            btn_editar.Content      = "Editar"
            btn_editar.Width        = 74
            btn_editar.Height       = 30
            btn_editar.FontSize     = 11
            btn_editar.Margin       = Thickness(4, 8, 4, 8)
            btn_editar.Background   = SolidColorBrush(Color.FromRgb(0, 150, 80))
            btn_editar.Foreground   = SolidColorBrush(Color.FromRgb(255, 255, 255))
            btn_editar.VerticalAlignment = VerticalAlignment.Center
            btn_editar.Tag          = idx
            btn_editar.Click       += self._ao_editar_abertura
            Grid.SetColumn(btn_editar, 1)
            grid.Children.Add(btn_editar)

            self.painel_aberturas.Children.Add(grid)

            sep        = Separator()
            sep.Margin = Thickness(0)
            self.painel_aberturas.Children.Add(sep)

    # ── Handlers ──────────────────────────────

    def _ao_mudar_grupo(self, sender, args):
        grupo = self._grupo_selecionado()
        if grupo is None:
            return

        nome = grupo["nome_grupo"] or ("Grupo_" + grupo["grupo_id"][:8])
        self.lbl_grupo_ab.Text = "Aberturas do grupo: {}".format(nome)

        wall_ids_validos      = filtrar_paredes_validas(self.doc, grupo["wall_ids"])
        self.aberturas_atuais = coletar_aberturas_do_grupo(self.doc, wall_ids_validos)
        self._reconstruir_lista_aberturas()

    def _ao_editar_abertura(self, sender, args):
        idx = sender.Tag
        ab  = self.aberturas_atuais[idx]

        janela_edicao = JanelaEditarAbertura(self.doc, ab)
        janela_edicao.ShowDialog()

        if janela_edicao.salvo:
            grupo = self._grupo_selecionado()
            if grupo:
                wall_ids_validos      = filtrar_paredes_validas(self.doc, grupo["wall_ids"])
                self.aberturas_atuais = coletar_aberturas_do_grupo(self.doc, wall_ids_validos)
                self._reconstruir_lista_aberturas()

    def _ao_isolar(self, sender, args):
        grupo = self._grupo_selecionado()
        if grupo is None:
            MessageBox.Show(
                "Selecione um grupo na lista antes de isolar.",
                "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning
            )
            return

        wall_ids = filtrar_paredes_validas(self.doc, grupo["wall_ids"])
        if not wall_ids:
            MessageBox.Show(
                "Nenhuma parede deste grupo existe mais no modelo.",
                "Grupo Invalido", MessageBoxButton.OK, MessageBoxImage.Warning
            )
            return

        if not self.aberturas_atuais:
            self.aberturas_atuais = coletar_aberturas_do_grupo(self.doc, wall_ids)

        try:
            isolar_grupo(self.doc, self.view, wall_ids, self.aberturas_atuais)
        except Exception as ex:
            MessageBox.Show(
                "Erro ao isolar:\n\n{}".format(str(ex)),
                "Erro", MessageBoxButton.OK, MessageBoxImage.Error
            )
            return

        nome = grupo["nome_grupo"] or ("Grupo_" + grupo["grupo_id"][:8])
        MessageBox.Show(
            "{} parede(s) e {} abertura(s) isoladas.\nGrupo: {}".format(
                len(wall_ids), len(self.aberturas_atuais), nome
            ),
            "Isolamento Ativo", MessageBoxButton.OK, MessageBoxImage.Information
        )
        self.Close()

    def _ao_resetar(self, sender, args):
        try:
            resetou = resetar_isolamento(self.doc, self.view)
            if resetou:
                MessageBox.Show(
                    "View restaurada. Todos os elementos estao visiveis novamente.",
                    "View Resetada", MessageBoxButton.OK, MessageBoxImage.Information
                )
        except Exception as ex:
            MessageBox.Show(
                "Erro ao resetar a view:\n\n{}".format(str(ex)),
                "Erro", MessageBoxButton.OK, MessageBoxImage.Error
            )
        self.Close()


# ─────────────────────────────────────────────
#  PONTO DE ENTRADA
# ─────────────────────────────────────────────

def main():
    doc   = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    view  = uidoc.ActiveView

    views_suportadas = [
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.Elevation,
        DB.ViewType.Section,
        DB.ViewType.ThreeD,
        DB.ViewType.Detail,
    ]
    if view.ViewType not in views_suportadas:
        forms.alert(
            "A view ativa nao suporta isolamento temporario.\n"
            "Abra uma planta, corte, elevacao ou vista 3D.",
            title="Visualizar Grupos"
        )
        return

    grupos = carregar_grupos(doc)
    janela = JanelaVisualizarGrupos(doc, view, grupos)
    janela.ShowDialog()


if __name__ == '__main__':
    main()
