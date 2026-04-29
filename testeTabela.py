# -*- coding: utf-8 -*-
__title__ = 'Criar\nTabelas'
__author__ = 'Samuel PLUGIN'

import clr
import os
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule,
    ScheduleSortGroupField, ScheduleSortOrder,
    ScheduleFilter, ScheduleFilterType,
    ElementId, Transaction, ScheduleSheetInstance, UV,
    ScheduleFieldType, Category, BuiltInCategory,
    ExternalDefinitionCreationOptions, InstanceBinding,
    StorageType,
)
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.UI import TaskDialog
import System.Windows as SW
import System.Windows.Controls as SWC
import System.Windows.Media as SWM
import System

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application

# =====================================================================
# DEFINIÇÃO DOS CAMPOS
# =====================================================================
CAMPOS_VERGALHAO = [
    (u"LOCAL",               u"Partição",                   True,  False),
    (u"POSIÇÃO (P)",         u"Número do vergalhão",        True,  False),
    (u"QUANTIDADE",          u"Contagem",                   True,  True),
    (u"DIÂMETRO (mm)",       u"Diâmetro da barra",          True,  True),
    (u"LARGURA (cm)",        u"A",                          True,  True),
    (u"COMPRIMENTO (cm)",    u"Comprimento da barra",       True,  True),
    (u"COMP. TOTAL (m)",     u"Comprimento total da barra", True,  True),
    (u"PESO (kg)",           u"Peso barra",                 True,  True),
]

CAMPOS_TELA = [
    (u"LOCAL",               u"Marca do hospedeiro",        True,  False),
    (u"N",                   u"Número da folha",            True,  False),
    (u"QUANT.",              u"Contagem",                   True,  True),
    (u"TELA",                u"Marca de tipo",              True,  False),
    (u"LARGURA",             u"Largura total do corte",     True,  True),
    (u"COMPRIMENTO",         u"Comprimento total do corte", True,  True),
    (u"PESO (Kgf)",          u"Massa da folha de corte",    True,  True),
]

CAT_VERGALHAO = -2009000
CAT_TELA      = -2009016

# Nome do parâmetro auxiliar gravado nos elementos
PARAM_AUX_PREFIX = u"_CALC_"

# =====================================================================
# HELPERS GERAIS
# =====================================================================
def cor(r, g, b):
    return SWM.SolidColorBrush(SWM.Color.FromRgb(r, g, b))

def get_nome_unico(nome_base):
    existing = set(
        s.Name for s in FilteredElementCollector(doc)
                           .OfClass(ViewSchedule).ToElements()
    )
    nome = nome_base
    i = 1
    while nome in existing:
        nome = u"{} ({})".format(nome_base, i)
        i += 1
    return nome

def get_field_by_name(sched, keyword):
    sd = sched.Definition
    for sf in sd.GetSchedulableFields():
        try:
            if sf.GetName(doc).lower() == keyword.lower():
                return sf
        except:
            pass
    for sf in sd.GetSchedulableFields():
        try:
            if keyword.lower() in sf.GetName(doc).lower():
                return sf
        except:
            pass
    return None

# =====================================================================
# PARÂMETRO COMPARTILHADO AUXILIAR
# =====================================================================
def garantir_parametro_aux(nome_param, cat_id_int):
    """
    Garante que o parâmetro compartilhado 'nome_param' existe
    e está vinculado à categoria. Compatível com Revit 2022-2025.
    """
    cat = None
    try:
        cat = doc.Settings.Categories.get_Item(
            System.Enum.ToObject(BuiltInCategory, cat_id_int)
        )
    except:
        pass
    if cat is None:
        return False

    # Verifica se já existe
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    while it.MoveNext():
        if it.Key.Name == nome_param:
            return True

    # Cria arquivo shared param temporário
    tmp_path = os.path.join(
        os.environ.get("TEMP", "C:\\Temp"),
        "samuel_plugin_shared_params.txt"
    )
    if not os.path.exists(tmp_path):
        with open(tmp_path, "w") as f:
            f.write("# This is a Revit shared parameter file.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("GROUP\t1\tSamuelPlugin\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\n")

    old_file = None
    try:
        old_file = app.SharedParametersFilename
        app.SharedParametersFilename = tmp_path
        spf = app.OpenSharedParameterFile()
        grp = spf.Groups.get_Item("SamuelPlugin")
        if grp is None:
            grp = spf.Groups.Create("SamuelPlugin")

        # Busca definição existente no arquivo
        ext_def = None
        for d in grp.Definitions:
            if d.Name == nome_param:
                ext_def = d
                break

        if ext_def is None:
            # Tenta criar com SpecTypeId (Revit 2022+)
            try:
                from Autodesk.Revit.DB import SpecTypeId
                opts = ExternalDefinitionCreationOptions(nome_param, SpecTypeId.Number)
                ext_def = grp.Definitions.Create(opts)
            except:
                pass

            # Fallback para Revit mais antigo
            if ext_def is None:
                try:
                    from Autodesk.Revit.DB import UnitType
                    opts = ExternalDefinitionCreationOptions(nome_param, UnitType.UT_Number)
                    ext_def = grp.Definitions.Create(opts)
                except:
                    pass

            # Último fallback — só nome
            if ext_def is None:
                try:
                    ext_def = grp.Definitions.Create(nome_param)
                except:
                    pass

        if ext_def is None:
            return False

        # Vincula à categoria
        cat_set = app.Create.NewCategorySet()
        cat_set.Insert(cat)
        binding = app.Create.NewInstanceBinding(cat_set)

        # Tenta inserir com GroupTypeId (Revit 2023+)
        inserted = False
        try:
            from Autodesk.Revit.DB import GroupTypeId
            doc.ParameterBindings.Insert(ext_def, binding, GroupTypeId.Data)
            inserted = True
        except:
            pass

        # Fallback BuiltInParameterGroup para versões antigas
        if not inserted:
            try:
                from Autodesk.Revit.DB import BuiltInParameterGroup
                doc.ParameterBindings.Insert(
                    ext_def, binding, BuiltInParameterGroup.PG_DATA
                )
                inserted = True
            except:
                pass

        # Último fallback sem grupo
        if not inserted:
            try:
                doc.ParameterBindings.Insert(ext_def, binding)
            except:
                pass

        return True

    except Exception as ex:
        return False
    finally:
        if old_file is not None:
            try:
                app.SharedParametersFilename = old_file
            except:
                pass

def gravar_parametro_aux(elementos, nome_param, valores_por_id):
    """
    Grava o valor calculado em nome_param para cada elemento.
    valores_por_id: dict { ElementId.IntegerValue → valor_float }
    """
    for el in elementos:
        val = valores_por_id.get(el.Id.IntegerValue)
        if val is None:
            continue
        try:
            p = el.LookupParameter(nome_param)
            if p and not p.IsReadOnly:
                p.Set(float(val))
        except:
            pass


# =====================================================================
# LEITURA DE LOCAIS DO MODELO
# =====================================================================
def get_locais_vergalhao():
    """Retorna lista de valores únicos do parâmetro Partição nos vergalhões."""
    locais = set()
    cat_id = ElementId(CAT_VERGALHAO)
    els = FilteredElementCollector(doc)\
            .OfCategoryId(cat_id)\
            .WhereElementIsNotElementType()\
            .ToElements()
    for el in els:
        try:
            p = el.LookupParameter(u"Parti\xe7\xe3o")
            if p and p.HasValue:
                v = p.AsString()
                if v and v.strip():
                    locais.add(v.strip())
        except:
            pass
    return sorted(locais)


# =====================================================================
# CRIAÇÃO DA TABELA
# =====================================================================
def criar_tabela(nome, cat_id_int, campos_sel,
                 params_aux=None,
                 filtro_texto=None, campo_filtro_kw=None,
                 diametros=None, larguras=None):
    """
    params_aux: lista de nomes de parâmetros auxiliares a incluir como colunas
    """
    cat_id     = ElementId(cat_id_int)
    nome_final = get_nome_unico(nome)
    sched      = ViewSchedule.CreateSchedule(doc, cat_id)
    sched.Name = nome_final
    sd         = sched.Definition

    sd.ShowHeaders = True
    sd.IsItemized  = False

    primeiro_campo_id = None
    campo_posicao_id  = None
    campo_diametro_id = None
    campo_largura_id  = None

    # ── Campos normais ────────────────────────────────────────────────
    for header, kw, _, _ in campos_sel:
        sf = get_field_by_name(sched, kw)
        if not sf:
            continue
        campo = sd.AddField(sf)
        campo.ColumnHeading = header

        if primeiro_campo_id is None:
            primeiro_campo_id = campo.FieldId

        nome_lower = (header + kw).lower()
        if u"posição" in nome_lower or u"vergalhão" in nome_lower:
            campo_posicao_id = campo.FieldId
        if u"diâmetro" in nome_lower:
            campo_diametro_id = campo.FieldId
        if kw == u"A":
            campo_largura_id = campo.FieldId

    # ── Colunas auxiliares calculadas (parâmetros gravados no modelo) ─
    if params_aux:
        for nome_aux in params_aux:
            sf = get_field_by_name(sched, nome_aux)
            if sf:
                campo_aux = sd.AddField(sf)
                campo_aux.ColumnHeading = nome_aux

    # ── Ordenação e agrupamento ───────────────────────────────────────
    if primeiro_campo_id:
        sgf_local = ScheduleSortGroupField(
            primeiro_campo_id, ScheduleSortOrder.Ascending
        )
        sgf_local.ShowHeader = True
        sd.AddSortGroupField(sgf_local)

    if campo_posicao_id:
        sgf_pos = ScheduleSortGroupField(
            campo_posicao_id, ScheduleSortOrder.Ascending
        )
        sgf_pos.ShowHeader      = True
        sgf_pos.ShowFooter      = True
        sgf_pos.ShowFooterTitle = False
        sgf_pos.ShowFooterCount = False
        sgf_pos.ShowBlankLine   = False
        sd.AddSortGroupField(sgf_pos)

    # ── Filtro LOCAL ──────────────────────────────────────────────────
    if filtro_texto and campo_filtro_kw:
        for i in range(sd.GetFieldCount()):
            f = sd.GetField(i)
            try:
                if campo_filtro_kw.lower() in f.GetName().lower():
                    sd.AddFilter(ScheduleFilter(
                        f.FieldId,
                        ScheduleFilterType.Contains,
                        filtro_texto
                    ))
                    break
            except:
                pass

    # ── Filtro diâmetro ───────────────────────────────────────────────
    if diametros and len(diametros) == 1 and campo_diametro_id:
        try:
            sd.AddFilter(ScheduleFilter(
                campo_diametro_id,
                ScheduleFilterType.Equal,
                float(diametros[0]) / 304.8
            ))
        except:
            pass

    # ── Filtro largura ────────────────────────────────────────────────
    if larguras and len(larguras) == 1 and campo_largura_id:
        try:
            sd.AddFilter(ScheduleFilter(
                campo_largura_id,
                ScheduleFilterType.Equal,
                float(larguras[0]) / 30.48
            ))
        except:
            pass

    return sched, nome_final


def inserir_na_folha(sched):
    vista = uidoc.ActiveView
    if not isinstance(vista, ViewSheet):
        return False
    try:
        ScheduleSheetInstance.Create(doc, vista.Id, sched.Id, UV(0.15, 0.15))
        return True
    except:
        return False


# =====================================================================
# CÁLCULO E GRAVAÇÃO DOS VALORES AUXILIARES
# =====================================================================
def aplicar_calculos_por_local(regras, campos_sel_marcados):
    """
    regras: lista de dicts {
        'local': str,
        'campo_kw': str,       # keyword do parâmetro base
        'operacao': '*' ou '/',
        'valor': float,
        'nome_aux': str        # nome do parâmetro auxiliar a gravar
    }
    Retorna lista de nomes de parâmetros auxiliares criados.
    """
    if not regras:
        return []

    cat_id_int = CAT_VERGALHAO
    els = FilteredElementCollector(doc)\
            .OfCategoryId(ElementId(cat_id_int))\
            .WhereElementIsNotElementType()\
            .ToElements()

    # Agrupa elementos por (local, posição) para somar contagem
    # e calcular o valor auxiliar corretamente
    params_criados = []

    # Garante parâmetros auxiliares no modelo
    nomes_aux_unicos = list(set(r['nome_aux'] for r in regras))
    for nome_aux in nomes_aux_unicos:
        ok = garantir_parametro_aux(nome_aux, cat_id_int)
        if ok:
            params_criados.append(nome_aux)

    if not params_criados:
        return []

    # Para cada elemento, calcula e grava
    valores_por_id = {}  # { (nome_aux, el_id_int) → valor }

    for el in els:
        # Lê LOCAL
        try:
            p_local = el.LookupParameter(u"Parti\xe7\xe3o")
            local_val = p_local.AsString().strip() if (p_local and p_local.HasValue and p_local.AsString()) else u""
        except:
            local_val = u""

        for regra in regras:
            if regra['local'].strip().lower() != local_val.lower():
                continue

            nome_aux  = regra['nome_aux']
            campo_kw  = regra['campo_kw']
            operacao  = regra['operacao']
            fator     = regra['valor']

            # Lê o valor base do elemento
            try:
                p_base = el.LookupParameter(campo_kw)
                if p_base is None:
                    # tenta busca parcial
                    for p_nome in [campo_kw]:
                        p_base = el.LookupParameter(p_nome)
                        if p_base:
                            break

                if p_base and p_base.HasValue:
                    if p_base.StorageType == StorageType.Double:
                        base_val = p_base.AsDouble()
                    elif p_base.StorageType == StorageType.Integer:
                        base_val = float(p_base.AsInteger())
                    else:
                        continue
                else:
                    continue
            except:
                continue

            # Calcula
            try:
                if operacao == u"*":
                    resultado = base_val * fator
                else:
                    resultado = base_val / fator if fator != 0 else base_val
            except:
                continue

            key = (nome_aux, el.Id.IntegerValue)
            valores_por_id[key] = resultado

    # Grava nos elementos
    for el in els:
        for nome_aux in nomes_aux_unicos:
            key = (nome_aux, el.Id.IntegerValue)
            val = valores_por_id.get(key)
            if val is None:
                val = 0.0
            try:
                p = el.LookupParameter(nome_aux)
                if p and not p.IsReadOnly:
                    p.Set(float(val))
            except:
                pass

    return params_criados


# =====================================================================
# INTERFACE WPF — WIDGETS AUXILIARES
# =====================================================================
def label(txt, negrito=False, tamanho=10, cor_txt=(100, 100, 100)):
    t = SWC.TextBlock()
    t.Text       = txt
    t.FontSize   = tamanho
    t.Foreground = cor(*cor_txt)
    t.Margin     = SW.Thickness(0, 2, 0, 2)
    if negrito:
        t.FontWeight = SW.FontWeights.Bold
    return t

def criar_cb(texto, marcado=True, tamanho=11):
    cb = SWC.CheckBox()
    cb.Content   = texto
    cb.IsChecked = marcado
    cb.FontSize  = tamanho
    cb.Margin    = SW.Thickness(6, 3, 6, 3)
    return cb

def secao(titulo_txt):
    borda = SWC.Border()
    borda.Background   = cor(220, 230, 245)
    borda.CornerRadius = SW.CornerRadius(3)
    borda.Padding      = SW.Thickness(8, 4, 8, 4)
    borda.Margin       = SW.Thickness(0, 12, 0, 4)
    lbl = SWC.TextBlock()
    lbl.Text       = titulo_txt
    lbl.FontWeight = SW.FontWeights.Bold
    lbl.FontSize   = 11
    lbl.Foreground = cor(30, 70, 150)
    borda.Child    = lbl
    return borda

def campo_texto(texto_inicial=u""):
    tb = SWC.TextBox()
    tb.Text            = texto_inicial
    tb.FontSize        = 11
    tb.Padding         = SW.Thickness(6, 4, 6, 4)
    tb.Margin          = SW.Thickness(0, 2, 0, 6)
    tb.BorderBrush     = cor(180, 190, 210)
    tb.BorderThickness = SW.Thickness(1)
    return tb


def build_campos(parent, campos_lista):
    """Campos simples com checkbox — sem linha de cálculo."""
    itens = []
    for header, kw, default, permite_calc in campos_lista:
        cb = SWC.CheckBox()
        cb.Content   = u"{}  —  {}".format(header, kw)
        cb.IsChecked = default
        cb.FontSize  = 11
        cb.Margin    = SW.Thickness(6, 3, 6, 3)
        parent.Children.Add(cb)
        itens.append({"cb": cb, "header": header, "kw": kw,
                      "permite_calc": permite_calc})
    return itens


def build_painel_multiplicadores(parent, locais, campos_numericos):
    """
    Monta a seção de multiplicadores por LOCAL.
    Para cada LOCAL selecionado o usuário escolhe:
      - campo base (combo)
      - operação × ou ÷
      - valor

    Retorna função get_regras() que coleta os dados preenchidos.
    """
    linhas = []   # lista de dicts com os widgets de cada linha

    container = SWC.StackPanel()
    container.Margin = SW.Thickness(0, 4, 0, 4)
    parent.Children.Add(container)

    def adicionar_linha(local_fixo=None):
        borda = SWC.Border()
        borda.BorderBrush     = cor(200, 210, 230)
        borda.BorderThickness = SW.Thickness(1)
        borda.CornerRadius    = SW.CornerRadius(4)
        borda.Padding         = SW.Thickness(8, 6, 8, 6)
        borda.Margin          = SW.Thickness(0, 4, 0, 0)

        row = SWC.StackPanel()
        row.Orientation = SWC.Orientation.Vertical
        borda.Child = row

        # Linha 1: LOCAL + botão remover
        l1 = SWC.StackPanel()
        l1.Orientation = SWC.Orientation.Horizontal

        lbl_local = SWC.TextBlock()
        lbl_local.Text     = u"LOCAL:"
        lbl_local.FontSize = 10
        lbl_local.Margin   = SW.Thickness(0, 0, 6, 0)
        lbl_local.VerticalAlignment = SW.VerticalAlignment.Center

        combo_local = SWC.ComboBox()
        combo_local.FontSize = 10
        combo_local.MinWidth = 140
        combo_local.Margin   = SW.Thickness(0, 0, 12, 0)
        for loc in locais:
            item = SWC.ComboBoxItem()
            item.Content = loc
            combo_local.Items.Add(item)
        if locais:
            combo_local.SelectedIndex = 0

        btn_rem = SWC.Button()
        btn_rem.Content    = u"✕"
        btn_rem.FontSize   = 10
        btn_rem.Width      = 24
        btn_rem.Height     = 24
        btn_rem.Background = cor(220, 80, 80)
        btn_rem.Foreground = SWM.Brushes.White
        btn_rem.Margin     = SW.Thickness(8, 0, 0, 0)

        l1.Children.Add(lbl_local)
        l1.Children.Add(combo_local)
        l1.Children.Add(btn_rem)
        row.Children.Add(l1)

        # Linha 2: CAMPO + operação + valor + nome auxiliar
        l2 = SWC.StackPanel()
        l2.Orientation = SWC.Orientation.Horizontal
        l2.Margin      = SW.Thickness(0, 6, 0, 0)

        lbl_campo = SWC.TextBlock()
        lbl_campo.Text     = u"Campo:"
        lbl_campo.FontSize = 10
        lbl_campo.Margin   = SW.Thickness(0, 0, 6, 0)
        lbl_campo.VerticalAlignment = SW.VerticalAlignment.Center

        combo_campo = SWC.ComboBox()
        combo_campo.FontSize = 10
        combo_campo.MinWidth = 160
        combo_campo.Margin   = SW.Thickness(0, 0, 8, 0)
        for header, kw in campos_numericos:
            item = SWC.ComboBoxItem()
            item.Content = u"{} [{}]".format(header, kw)
            item.Tag     = kw
            combo_campo.Items.Add(item)
        if campos_numericos:
            combo_campo.SelectedIndex = 0

        combo_op = SWC.ComboBox()
        combo_op.FontSize = 10
        combo_op.Width    = 44
        combo_op.Margin   = SW.Thickness(0, 0, 6, 0)
        for op in [u"×", u"÷"]:
            oi = SWC.ComboBoxItem()
            oi.Content = op
            combo_op.Items.Add(oi)
        combo_op.SelectedIndex = 0

        txt_val = SWC.TextBox()
        txt_val.Text      = u"1"
        txt_val.FontSize  = 10
        txt_val.Width     = 56
        txt_val.Padding   = SW.Thickness(4, 2, 4, 2)
        txt_val.Margin    = SW.Thickness(0, 0, 10, 0)
        txt_val.BorderBrush     = cor(180, 190, 210)
        txt_val.BorderThickness = SW.Thickness(1)

        lbl_aux = SWC.TextBlock()
        lbl_aux.Text       = u"→ col. extra:"
        lbl_aux.FontSize   = 9
        lbl_aux.Foreground = cor(100, 100, 180)
        lbl_aux.Margin     = SW.Thickness(0, 0, 4, 0)
        lbl_aux.VerticalAlignment = SW.VerticalAlignment.Center

        txt_nome_aux = SWC.TextBox()
        txt_nome_aux.FontSize  = 10
        txt_nome_aux.Width     = 100
        txt_nome_aux.Padding   = SW.Thickness(4, 2, 4, 2)
        txt_nome_aux.BorderBrush     = cor(180, 190, 210)
        txt_nome_aux.BorderThickness = SW.Thickness(1)
        txt_nome_aux.Text = u"QTD CALC"

        l2.Children.Add(lbl_campo)
        l2.Children.Add(combo_campo)
        l2.Children.Add(combo_op)
        l2.Children.Add(txt_val)
        l2.Children.Add(lbl_aux)
        l2.Children.Add(txt_nome_aux)
        row.Children.Add(l2)

        container.Children.Add(borda)

        entrada = {
            "borda":        borda,
            "combo_local":  combo_local,
            "combo_campo":  combo_campo,
            "combo_op":     combo_op,
            "txt_val":      txt_val,
            "txt_nome_aux": txt_nome_aux,
        }
        linhas.append(entrada)

        def on_remover(s, e, b=borda, en=entrada):
            try:
                container.Children.Remove(b)
                linhas.remove(en)
            except:
                pass

        btn_rem.Click += on_remover

    # Botão adicionar linha
    btn_add = SWC.Button()
    btn_add.Content    = u"＋  Adicionar multiplicador por LOCAL"
    btn_add.FontSize   = 10
    btn_add.Height     = 28
    btn_add.Background = cor(0, 130, 80)
    btn_add.Foreground = SWM.Brushes.White
    btn_add.Margin     = SW.Thickness(0, 6, 0, 2)
    btn_add.Click     += lambda s, e: adicionar_linha()
    parent.Children.Add(btn_add)

    def get_regras():
        regras = []
        for en in linhas:
            try:
                local = en["combo_local"].SelectedItem.Content
            except:
                continue

            idx_campo = en["combo_campo"].SelectedIndex
            if idx_campo < 0 or idx_campo >= len(campos_numericos):
                continue
            campo_header, campo_kw = campos_numericos[idx_campo]

            op_idx   = en["combo_op"].SelectedIndex
            operacao = u"*" if op_idx == 0 else u"/"

            try:
                valor = float(en["txt_val"].Text.replace(u",", u"."))
            except:
                continue

            nome_aux = en["txt_nome_aux"].Text.strip() or u"QTD CALC"

            regras.append({
                "local":     local,
                "campo_kw":  campo_kw,
                "operacao":  operacao,
                "valor":     valor,
                "nome_aux":  nome_aux,
            })
        return regras

    return get_regras


# =====================================================================
# JANELA PRINCIPAL
# =====================================================================
def mostrar_janela():
    # Lê locais do modelo antes de abrir a janela
    locais = get_locais_vergalhao()
    campos_numericos_v = [
        (h, k) for h, k, _, permite in CAMPOS_VERGALHAO if permite
    ]

    w = SW.Window()
    w.Title    = u"Criar Tabelas  —  Samuel PLUGIN"
    w.Width    = 600
    w.Height   = 920
    w.ResizeMode = SW.ResizeMode.NoResize
    w.WindowStartupLocation = SW.WindowStartupLocation.CenterScreen
    w.Background = cor(245, 247, 252)

    scroll = SWC.ScrollViewer()
    scroll.VerticalScrollBarVisibility = SWC.ScrollBarVisibility.Auto

    main = SWC.StackPanel()
    main.Margin = SW.Thickness(22, 18, 22, 18)

    # Cabeçalho
    t1 = SWC.TextBlock()
    t1.Text       = u"Criador de Tabelas"
    t1.FontSize   = 18
    t1.FontWeight = SW.FontWeights.Bold
    t1.Foreground = cor(20, 70, 160)
    t1.Margin     = SW.Thickness(0, 0, 0, 2)
    main.Children.Add(t1)

    t2 = SWC.TextBlock()
    t2.Text       = u"Armação de Aberturas e Tela Soldada"
    t2.FontSize   = 10
    t2.Foreground = cor(130, 130, 130)
    t2.Margin     = SW.Thickness(0, 0, 0, 10)
    main.Children.Add(t2)

    # ── TIPO ──────────────────────────────────────────────────────────
    main.Children.Add(secao(u"  TIPO DE TABELA"))
    cb_verg = criar_cb(u"Armação de Aberturas  (Vergalhões)", True, 12)
    cb_tela = criar_cb(u"Armadura de Tela Soldada  (Fabric Sheets)", True, 12)
    main.Children.Add(cb_verg)
    main.Children.Add(cb_tela)

    # ── CAMPOS VERGALHÃO ──────────────────────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  —  VERGALHÕES"))
    painel_v = SWC.StackPanel()
    itens_v  = build_campos(painel_v, CAMPOS_VERGALHAO)
    main.Children.Add(painel_v)

    # ── CAMPOS TELA ───────────────────────────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  —  TELA SOLDADA"))
    painel_t = SWC.StackPanel()
    itens_t  = build_campos(painel_t, CAMPOS_TELA)
    main.Children.Add(painel_t)

    # ── MULTIPLICADORES POR LOCAL ─────────────────────────────────────
    main.Children.Add(secao(u"  MULTIPLICADORES POR LOCAL  (Vergalhões)"))
    main.Children.Add(label(
        u"Escolha o LOCAL, o campo, a operação e o fator. "
        u"Uma coluna extra será criada na tabela com o resultado.",
        cor_txt=(80, 80, 80)
    ))

    if not locais:
        main.Children.Add(label(
            u"⚠  Nenhum LOCAL (Partição) encontrado nos vergalhões do modelo.",
            cor_txt=(180, 80, 0)
        ))
        get_regras = lambda: []
    else:
        get_regras = build_painel_multiplicadores(
            main, locais, campos_numericos_v
        )

    # ── FILTROS ───────────────────────────────────────────────────────
    main.Children.Add(secao(u"  FILTROS  (opcionais)"))
    main.Children.Add(label(u"Filtrar por LOCAL / Partição  (vazio = todos):"))
    txt_filtro = campo_texto()
    main.Children.Add(txt_filtro)

    main.Children.Add(label(u"Filtrar por DIÂMETRO (mm)  ex: 6.3, 8.0:"))
    txt_diam = campo_texto()
    main.Children.Add(txt_diam)

    main.Children.Add(label(u"Filtrar por LARGURA A (cm)  ex: 20, 30:"))
    txt_larg = campo_texto()
    main.Children.Add(txt_larg)

    # ── OPÇÕES ────────────────────────────────────────────────────────
    main.Children.Add(secao(u"  OPÇÕES"))
    cb_folha = criar_cb(
        u"Inserir tabela na folha ativa  (só se for uma prancha)", False
    )
    main.Children.Add(cb_folha)

    # Aviso
    aviso = SWC.Border()
    aviso.Background   = cor(255, 243, 205)
    aviso.CornerRadius = SW.CornerRadius(3)
    aviso.Padding      = SW.Thickness(10, 6, 10, 6)
    aviso.Margin       = SW.Thickness(0, 10, 0, 4)
    aviso_txt = SWC.TextBlock()
    aviso_txt.Text = (
        u"⚠  Os multiplicadores gravam o resultado em parâmetros auxiliares\n"
        u"nos próprios elementos (criados automaticamente se não existirem).\n"
        u"QUANTIDADE usa 'Contagem' — barras reais no modelo."
    )
    aviso_txt.FontSize     = 10
    aviso_txt.Foreground   = cor(120, 80, 0)
    aviso_txt.TextWrapping = SW.TextWrapping.Wrap
    aviso.Child = aviso_txt
    main.Children.Add(aviso)

    # ── BOTÕES ────────────────────────────────────────────────────────
    btns = SWC.StackPanel()
    btns.Orientation         = SWC.Orientation.Horizontal
    btns.HorizontalAlignment = SW.HorizontalAlignment.Right
    btns.Margin              = SW.Thickness(0, 16, 0, 4)

    btn_ok = SWC.Button()
    btn_ok.Content    = u"✔  Criar Tabelas"
    btn_ok.Width      = 150
    btn_ok.Height     = 38
    btn_ok.FontSize   = 12
    btn_ok.FontWeight = SW.FontWeights.Bold
    btn_ok.Background = cor(0, 110, 200)
    btn_ok.Foreground = SWM.Brushes.White
    btn_ok.Margin     = SW.Thickness(0, 0, 10, 0)

    btn_cancel = SWC.Button()
    btn_cancel.Content  = u"Cancelar"
    btn_cancel.Width    = 90
    btn_cancel.Height   = 38
    btn_cancel.FontSize = 11

    resultado = [None]

    def parse_lista(texto):
        if not texto or not texto.strip():
            return None
        return [v.strip() for v in texto.split(u",") if v.strip()]

    def on_ok(s, e):
        campos_v = [(it["header"], it["kw"], True, it["permite_calc"])
                    for it in itens_v if it["cb"].IsChecked]
        campos_t = [(it["header"], it["kw"], True, it["permite_calc"])
                    for it in itens_t if it["cb"].IsChecked]

        resultado[0] = {
            'verg':      cb_verg.IsChecked and len(campos_v) > 0,
            'tela':      cb_tela.IsChecked and len(campos_t) > 0,
            'campos_v':  campos_v,
            'campos_t':  campos_t,
            'regras':    get_regras(),
            'filtro':    (txt_filtro.Text or u"").strip() or None,
            'diametros': parse_lista(txt_diam.Text),
            'larguras':  parse_lista(txt_larg.Text),
            'na_folha':  cb_folha.IsChecked,
        }
        w.DialogResult = True
        w.Close()

    def on_cancel(s, e):
        w.DialogResult = False
        w.Close()

    btn_ok.Click     += on_ok
    btn_cancel.Click += on_cancel
    btns.Children.Add(btn_ok)
    btns.Children.Add(btn_cancel)
    main.Children.Add(btns)

    scroll.Content = main
    w.Content      = scroll
    ok = w.ShowDialog()
    return resultado[0] if ok else None


# =====================================================================
# MAIN
# =====================================================================
def main():
    opcoes = mostrar_janela()
    if opcoes is None:
        return

    criadas = []
    erros   = []
    params_aux_criados = []

    t = Transaction(doc, u"Criar Tabelas - Samuel PLUGIN")
    t.Start()
    try:
        # 1. Aplica multiplicadores e grava parâmetros auxiliares
        if opcoes['regras'] and opcoes['verg']:
            params_aux_criados = aplicar_calculos_por_local(
                opcoes['regras'],
                opcoes['campos_v']
            )

        # 2. Tabela vergalhões
        if opcoes['verg']:
            try:
                sched, nome = criar_tabela(
                    u"TABELA DE ARMAÇÃO ABERTURAS",
                    CAT_VERGALHAO,
                    opcoes['campos_v'],
                    params_aux=params_aux_criados,
                    filtro_texto=opcoes['filtro'],
                    campo_filtro_kw=u"parti",
                    diametros=opcoes['diametros'],
                    larguras=opcoes['larguras'],
                )
                criadas.append(u"VERGALHÕES: {}".format(nome))
                if opcoes['na_folha']:
                    inserir_na_folha(sched)
            except Exception as ex:
                erros.append(u"Vergalhões: {}".format(str(ex)))

        # 3. Tabela tela soldada
        if opcoes['tela']:
            try:
                sched, nome = criar_tabela(
                    u"TABELA DE ARMADURA DE TELA SOLDADA",
                    CAT_TELA,
                    opcoes['campos_t'],
                    filtro_texto=opcoes['filtro'],
                    campo_filtro_kw=u"hospedeiro",
                )
                criadas.append(u"TELA SOLDADA: {}".format(nome))
                if opcoes['na_folha']:
                    inserir_na_folha(sched)
            except Exception as ex:
                erros.append(u"Tela Soldada: {}".format(str(ex)))

        t.Commit()
    except Exception as ex:
        t.RollBack()
        TaskDialog.Show(u"Erro Crítico", str(ex))
        return

    partes = []
    if criadas:
        partes.append(u"Tabelas criadas:\n" + u"\n".join(u"  - " + c for c in criadas))
    if params_aux_criados:
        partes.append(u"Parâmetros auxiliares gravados:\n" +
                      u"\n".join(u"  - " + p for p in params_aux_criados))
    if erros:
        partes.append(u"Atenção — erros:\n" + u"\n".join(u"  - " + e for e in erros))
    if not partes:
        partes.append(u"Nenhuma tabela criada.")

    TaskDialog.Show(u"Criar Tabelas", u"\n\n".join(partes))

main()
