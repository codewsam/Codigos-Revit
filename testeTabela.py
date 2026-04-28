# -*- coding: utf-8 -*-
__title__ = 'Criar\nTabelas'
__author__ = 'Samuel PLUGIN'

import clr
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
    ScheduleFieldType,
)
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.UI import TaskDialog
import System.Windows as SW
import System.Windows.Controls as SWC
import System.Windows.Media as SWM

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# =====================================================================
# DEFINICAO DOS CAMPOS
# =====================================================================
# FIX: Usa "Contagem" ao invés de "Quantidade por conjunto de vergalhões"
# "Quantidade por conjunto" multiplica pela geometria do detalhe do ferro
# "Contagem" representa a quantidade real de barras no modelo

CAMPOS_VERGALHAO = [
    (u"LOCAL",               u"Partição",                    True),
    (u"POSIÇÃO (P)",         u"Número do vergalhão",         True),
    (u"QUANTIDADE",          u"Contagem",                    True),   # ← CORRIGIDO
    (u"DIÂMETRO (mm)",       u"Diâmetro da barra",           True),
    (u"LARGURA (cm)",        u"A",                           True),   # ← NOVO
    (u"COMPRIMENTO (cm)",    u"Comprimento da barra",        True),
    (u"COMP. TOTAL (m)",     u"Comprimento total da barra",  True),
    (u"PESO (kg)",           u"Peso barra",                  True),
]

CAMPOS_TELA = [
    (u"LOCAL",               u"Marca do hospedeiro",         True),
    (u"N",                   u"Número da folha",             True),
    (u"QUANT.",              u"Contagem",                    True),
    (u"TELA",                u"Marca de tipo",               True),
    (u"LARGURA",             u"Largura total do corte",      True),
    (u"COMPRIMENTO",         u"Comprimento total do corte",  True),
    (u"PESO (Kgf)",          u"Massa da folha de corte",     True),
]

CAT_VERGALHAO = -2009000
CAT_TELA      = -2009016

# =====================================================================
# HELPERS
# =====================================================================
def cor(r, g, b):
    return SWM.SolidColorBrush(SWM.Color.FromRgb(r, g, b))

def get_nome_unico(nome_base):
    existing = set(s.Name for s in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements())
    nome = nome_base
    i = 1
    while nome in existing:
        nome = u"{} ({})".format(nome_base, i)
        i += 1
    return nome

def get_field_by_name(sched, keyword):
    sd = sched.Definition
    # Busca exata primeiro
    for sf in sd.GetSchedulableFields():
        try:
            n = sf.GetName(doc)
            if n.lower() == keyword.lower():
                return sf
        except:
            pass
    # Busca parcial como fallback
    for sf in sd.GetSchedulableFields():
        try:
            n = sf.GetName(doc)
            if keyword.lower() in n.lower():
                return sf
        except:
            pass
    return None

# =====================================================================
# CRIACAO DAS TABELAS
# =====================================================================
def criar_tabela(nome, cat_id_int, campos_sel, filtro_texto=None, campo_filtro_kw=None,
                 diametros=None, larguras=None):
    cat_id     = ElementId(cat_id_int)
    nome_final = get_nome_unico(nome)
    sched      = ViewSchedule.CreateSchedule(doc, cat_id)
    sched.Name = nome_final
    sd         = sched.Definition
    sd.ShowHeaders = True

    primeiro_campo_id = None
    campo_diametro_id = None
    campo_largura_id  = None

    for header, kw, _ in campos_sel:
        sf = get_field_by_name(sched, kw)
        if sf:
            campo = sd.AddField(sf)
            campo.ColumnHeading = header
            if primeiro_campo_id is None:
                primeiro_campo_id = campo.FieldId
            if u"diâmetro" in kw.lower() or u"diameter" in kw.lower():
                campo_diametro_id = campo.FieldId
            if kw == u"A":
                campo_largura_id = campo.FieldId

    # Ordenar por LOCAL → POSIÇÃO
    if primeiro_campo_id:
        sgf = ScheduleSortGroupField(primeiro_campo_id, ScheduleSortOrder.Ascending)
        sd.AddSortGroupField(sgf)

    # Filtro de texto (LOCAL/Partição)
    if filtro_texto and campo_filtro_kw:
        for i in range(sd.GetFieldCount()):
            f = sd.GetField(i)
            try:
                n = f.GetName()
                if campo_filtro_kw.lower() in n.lower():
                    sf_filter = ScheduleFilter(f.FieldId, ScheduleFilterType.Contains, filtro_texto)
                    sd.AddFilter(sf_filter)
                    break
            except:
                pass

    # Filtro por diâmetro (múltiplos valores = OR via múltiplos filtros separados)
    # Revit não suporta OR nativamente — criamos uma tabela por diâmetro se necessário
    # Por ora, filtra apenas se exatamente 1 diâmetro selecionado
    if diametros and len(diametros) == 1 and campo_diametro_id:
        try:
            val = float(diametros[0]) / 304.8  # mm → pés
            sf_d = ScheduleFilter(campo_diametro_id, ScheduleFilterType.Equal, val)
            sd.AddFilter(sf_d)
        except:
            pass

    # Filtro por largura
    if larguras and len(larguras) == 1 and campo_largura_id:
        try:
            val = float(larguras[0]) / 30.48  # cm → pés
            sf_l = ScheduleFilter(campo_largura_id, ScheduleFilterType.Equal, val)
            sd.AddFilter(sf_l)
        except:
            pass

    return sched, nome_final

def inserir_na_folha(sched):
    vista = uidoc.ActiveView
    if not isinstance(vista, ViewSheet):
        return False
    try:
        ponto = UV(0.15, 0.15)
        ScheduleSheetInstance.Create(doc, vista.Id, sched.Id, ponto)
        return True
    except:
        return False

# =====================================================================
# INTERFACE WPF
# =====================================================================
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

def campo_texto(placeholder, texto_inicial=u""):
    tb = SWC.TextBox()
    tb.Text            = texto_inicial
    tb.FontSize        = 11
    tb.Padding         = SW.Thickness(6, 4, 6, 4)
    tb.Margin          = SW.Thickness(0, 2, 0, 6)
    tb.BorderBrush     = cor(180, 190, 210)
    tb.BorderThickness = SW.Thickness(1)
    return tb

def label(txt, negrito=False, tamanho=10, cor_txt=(100,100,100)):
    t = SWC.TextBlock()
    t.Text       = txt
    t.FontSize   = tamanho
    t.Foreground = cor(*cor_txt)
    t.Margin     = SW.Thickness(0, 2, 0, 2)
    if negrito:
        t.FontWeight = SW.FontWeights.Bold
    return t

def linha_horizontal():
    sep = SWC.Separator()
    sep.Margin = SW.Thickness(0, 8, 0, 4)
    return sep

def mostrar_janela():
    w = SW.Window()
    w.Title    = u"Criar Tabelas  -  Samuel PLUGIN"
    w.Width    = 520
    w.Height   = 800
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

    # ── TIPO ──────────────────────────────────────────
    main.Children.Add(secao(u"  TIPO DE TABELA"))
    cb_verg = criar_cb(u"Armação de Aberturas  (Vergalhões / Rebar)", True, 12)
    cb_tela = criar_cb(u"Armadura de Tela Soldada  (Fabric Sheets)",  True, 12)
    main.Children.Add(cb_verg)
    main.Children.Add(cb_tela)

    # ── NOMES ─────────────────────────────────────────
    main.Children.Add(secao(u"  NOME DAS TABELAS"))
    main.Children.Add(label(u"Nome da tabela de Vergalhões:"))
    txt_nv = campo_texto(u"", u"TABELA DE ARMAÇÃO ABERTURAS")
    main.Children.Add(txt_nv)
    main.Children.Add(label(u"Nome da tabela de Tela Soldada:"))
    txt_nt = campo_texto(u"", u"TABELA DE ARMADURA DE TELA SOLDADA")
    main.Children.Add(txt_nt)

    # ── CAMPOS VERGALHAO ──────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  -  VERGALHÕES"))
    main.Children.Add(label(u"Marque os campos que deseja incluir:"))
    cbs_v = []
    for header, kw, default in CAMPOS_VERGALHAO:
        cb = criar_cb(u"{}".format(header), default)
        cbs_v.append((cb, header, kw))
        main.Children.Add(cb)

    # ── CAMPOS TELA ───────────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  -  TELA SOLDADA"))
    main.Children.Add(label(u"Marque os campos que deseja incluir:"))
    cbs_t = []
    for header, kw, default in CAMPOS_TELA:
        cb = criar_cb(u"{}".format(header), default)
        cbs_t.append((cb, header, kw))
        main.Children.Add(cb)

    # ── FILTROS ───────────────────────────────────────
    main.Children.Add(secao(u"  FILTROS  (opcionais)"))

    main.Children.Add(label(u"Filtrar por LOCAL / Partição:"))
    main.Children.Add(label(u"(Deixe vazio para todos)"))
    txt_filtro = campo_texto(u"ex: Terreo, P1, Cobertura...")
    main.Children.Add(txt_filtro)

    main.Children.Add(label(u"Filtrar por DIÂMETRO (mm) — separe por vírgula:"))
    main.Children.Add(label(u"(ex: 6.3, 8.0 — deixe vazio para todos)"))
    txt_diam = campo_texto(u"ex: 6.3, 8.0")
    main.Children.Add(txt_diam)

    main.Children.Add(label(u"Filtrar por LARGURA A (cm) — separe por vírgula:"))
    main.Children.Add(label(u"(ex: 20, 30 — deixe vazio para todos)"))
    txt_larg = campo_texto(u"ex: 20, 30")
    main.Children.Add(txt_larg)

    # ── OPCOES ────────────────────────────────────────
    main.Children.Add(secao(u"  OPÇÕES"))
    cb_folha = criar_cb(u"Inserir tabela na folha ativa  (apenas se for uma prancha)", False)
    main.Children.Add(cb_folha)

    # Aviso sobre quantidade
    aviso = SWC.Border()
    aviso.Background   = cor(255, 243, 205)
    aviso.CornerRadius = SW.CornerRadius(3)
    aviso.Padding      = SW.Thickness(10, 6, 10, 6)
    aviso.Margin       = SW.Thickness(0, 10, 0, 4)
    aviso_txt = SWC.TextBlock()
    aviso_txt.Text       = u"⚠ QUANTIDADE: usa o campo 'Contagem' (barras reais no modelo),\nnão 'Quantidade por conjunto' que multiplica pelo detalhe."
    aviso_txt.FontSize   = 10
    aviso_txt.Foreground = cor(120, 80, 0)
    aviso_txt.TextWrapping = SW.TextWrapping.Wrap
    aviso.Child = aviso_txt
    main.Children.Add(aviso)

    # ── BOTOES ────────────────────────────────────────
    btns = SWC.StackPanel()
    btns.Orientation          = SWC.Orientation.Horizontal
    btns.HorizontalAlignment  = SW.HorizontalAlignment.Right
    btns.Margin               = SW.Thickness(0, 16, 0, 4)

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
        campos_v  = [(h, k, True) for cb, h, k in cbs_v if cb.IsChecked]
        campos_t  = [(h, k, True) for cb, h, k in cbs_t if cb.IsChecked]
        filtro    = (txt_filtro.Text or u"").strip() or None
        diametros = parse_lista(txt_diam.Text)
        larguras  = parse_lista(txt_larg.Text)

        resultado[0] = {
            'verg':      cb_verg.IsChecked and len(campos_v) > 0,
            'tela':      cb_tela.IsChecked and len(campos_t) > 0,
            'nome_v':    txt_nv.Text.strip() or u"TABELA DE ARMAÇÃO ABERTURAS",
            'nome_t':    txt_nt.Text.strip() or u"TABELA DE ARMADURA DE TELA SOLDADA",
            'campos_v':  campos_v,
            'campos_t':  campos_t,
            'filtro':    filtro,
            'diametros': diametros,
            'larguras':  larguras,
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

    t = Transaction(doc, u"Criar Tabelas - Samuel PLUGIN")
    t.Start()
    try:
        if opcoes['verg']:
            try:
                sched, nome = criar_tabela(
                    opcoes['nome_v'], CAT_VERGALHAO,
                    opcoes['campos_v'],
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

        if opcoes['tela']:
            try:
                sched, nome = criar_tabela(
                    opcoes['nome_t'], CAT_TELA,
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
        partes.append(u"Tabelas criadas com sucesso:\n" + u"\n".join(u"  - " + c for c in criadas))
    if erros:
        partes.append(u"Atenção - erros:\n" + u"\n".join(u"  - " + e for e in erros))
    if not partes:
        partes.append(u"Nenhuma tabela criada.\nMarque pelo menos um tipo e pelo menos um campo.")

    TaskDialog.Show(u"Criar Tabelas", u"\n\n".join(partes))

main()
