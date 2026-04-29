
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
    ScheduleField,
    ElementId, Transaction, ScheduleSheetInstance, UV,
)
from Autodesk.Revit.DB import ViewSheet
from Autodesk.Revit.UI import TaskDialog
import System.Windows as SW
import System.Windows.Controls as SWC
import System.Windows.Media as SWM
import unicodedata
 
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
 
# =====================================================================
# DEFINICAO DOS CAMPOS DISPONIVEIS
# =====================================================================
# (cabecalho_exibido, nome_do_campo_no_revit, marcado_por_padrao)
 
CAMPOS_VERGALHAO = [
    (u"LOCAL",                 u"Parti",                               True),
    (u"POSICAO (P)",           u"Numero do vergalhao",                 True),
    (u"QUANTIDADE",            u"Quantidade por conjunto",             True),
    (u"DIAMETRO (mm)",         u"Diametro da barra",                   True),
    (u"COMPRIMENTO (cm)",      u"Comprimento da barra",                True),
    (u"COMP. TOTAL (m)",       u"Comprimento total da barra",          True),
    (u"MASSA",                 u"Massa",                               False),
    (u"PESO (kg)",             u"Peso",                                True),
]
 
CAMPOS_TELA = [
    (u"LOCAL",                 u"Marca do hospedeiro",                 True),
    (u"N",                     u"Numero da folha de tela",             True),
    (u"QUANT.",                u"Contagem",                            True),
    (u"TELA",                  u"Marca de tipo",                       True),
    (u"LARGURA",               u"Largura total do corte",              True),
    (u"COMPRIMENTO",           u"Comprimento total do corte",          True),
    (u"PESO (Kgf)",            u"Massa da folha de corte",             True),
]
 
CAT_VERGALHAO = -2009000   # OST_Rebar
CAT_TELA      = -2009016   # OST_FabricSheets
 
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
    """Busca campo disponivel pela keyword (case-insensitive, match parcial)."""
    sd = sched.Definition
    def _strip_accents(s):
        try:
            return u"".join(c for c in unicodedata.normalize('NFKD', unicode(s)) if not unicodedata.combining(c)).lower()
        except:
            try:
                return unicode(s).lower()
            except:
                return (s or u"").lower()

    kw_norm = _strip_accents(keyword)
    for sf in sd.GetSchedulableFields():
        try:
            n = sf.GetName(doc)
            if kw_norm in _strip_accents(n):
                return sf
        except:
            pass
    return None

def _strip_accents(s):
    try:
        return u"".join(c for c in unicodedata.normalize('NFKD', unicode(s)) if not unicodedata.combining(c)).lower()
    except:
        try:
            return unicode(s).lower()
        except:
            return (s or u"").lower()

def schedule_has_schedulable(sd, schedulable_sf):
    try:
        target = _strip_accents(schedulable_sf.GetName(doc))
    except:
        return False
    for i in range(sd.GetFieldCount()):
        try:
            f = sd.GetField(i)
            sf = f.GetSchedulableField()
            if _strip_accents(sf.GetName(doc)) == target:
                return True
        except:
            pass
    return False
 
# =====================================================================
# CRIACAO DAS TABELAS
# =====================================================================
def criar_tabela(nome, cat_id_int, campos_sel, filtro_texto=None, campo_filtro_kw=None, is_vergalhao=False):
    cat_id     = ElementId(cat_id_int)
    nome_final = get_nome_unico(nome)
    sched      = ViewSchedule.CreateSchedule(doc, cat_id)
    sched.Name = nome_final
    sd         = sched.Definition
    sd.ShowHeaders = True
 
    primeiro_campo_id = None
    for header, kw, _ in campos_sel:
        sf = get_field_by_name(sched, kw)
        if sf:
            campo = sd.AddField(sf)
            campo.ColumnHeading = header
            if primeiro_campo_id is None:
                primeiro_campo_id = campo.FieldId
 

 
    # Filtro opcional
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

    # Agrupamento e agregacao para vergalhao
    if is_vergalhao:
        # tentar localizar o campo de POSIÇAO (várias variações)
        pos_keywords = [u"numero do vergalhao", u"numero do vergalhão", u"numero", u"posicao", u"posição", u"pos", u"posicao (p)", u"posição (p)"]
        pos_schedulable = None
        for pk in pos_keywords:
            pos_schedulable = get_field_by_name(sched, pk)
            if pos_schedulable:
                break

        if pos_schedulable:
            # adicionar o campo se ainda não existir na definição
            try:
                if not schedule_has_schedulable(sd, pos_schedulable):
                    campo_pos = sd.AddField(pos_schedulable)
                    try:
                        campo_pos.ColumnHeading = u"POSIÇÃO (P)"
                    except:
                        pass
            except:
                pass

            # localizar FieldId do campo de posicao
            pos_field_id = None
            for i in range(sd.GetFieldCount()):
                try:
                    f = sd.GetField(i)
                    if _strip_accents(f.GetSchedulableField().GetName(doc)) == _strip_accents(pos_schedulable.GetName(doc)):
                        pos_field_id = f.FieldId
                        break
                except:
                    pass

            if pos_field_id:
                # agrupar por posicao (reduz linhas repetidas por posição)
                sd.AddSortGroupField(ScheduleSortGroupField(pos_field_id, ScheduleSortOrder.Ascending))

                # agregar (somar) campos numéricos quando agrupados
                for i in range(sd.GetFieldCount()):
                    try:
                        f = sd.GetField(i)
                        fname = _strip_accents(f.GetSchedulableField().GetName(doc))
                        if any(k in fname for k in [u"quantidade", u"comprimento", u"peso", u"massa", u"comp total", u"comp. total"]):
                            f.SetAggregate(1)  # 1 = Sum
                    except:
                        pass

    return sched, nome_final
 
def inserir_na_folha(sched):
    """Insere a tabela na folha ativa se for uma ViewSheet."""
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
    borda.Background    = cor(220, 230, 245)
    borda.CornerRadius  = SW.CornerRadius(3)
    borda.Padding       = SW.Thickness(8, 4, 8, 4)
    borda.Margin        = SW.Thickness(0, 12, 0, 4)
    lbl = SWC.TextBlock()
    lbl.Text       = titulo_txt
    lbl.FontWeight = SW.FontWeights.Bold
    lbl.FontSize   = 11
    lbl.Foreground = cor(30, 70, 150)
    borda.Child    = lbl
    return borda
 
def campo_texto(placeholder, texto_inicial=u""):
    tb = SWC.TextBox()
    tb.Text    = texto_inicial
    tb.FontSize = 11
    tb.Padding = SW.Thickness(6, 4, 6, 4)
    tb.Margin  = SW.Thickness(0, 2, 0, 6)
    tb.BorderBrush     = cor(180, 190, 210)
    tb.BorderThickness = SW.Thickness(1)
    return tb
 
def label(txt, negrito=False, tamanho=10, cor_txt=(100,100,100)):
    t = SWC.TextBlock()
    t.Text      = txt
    t.FontSize  = tamanho
    t.Foreground = cor(*cor_txt)
    t.Margin    = SW.Thickness(0, 2, 0, 2)
    if negrito:
        t.FontWeight = SW.FontWeights.Bold
    return t
 
def mostrar_janela():
    w = SW.Window()
    w.Title    = u"Criar Tabelas  -  Samuel PLUGIN"
    w.Width    = 500
    w.Height   = 720
    w.ResizeMode = SW.ResizeMode.NoResize
    w.WindowStartupLocation = SW.WindowStartupLocation.CenterScreen
    w.Background = cor(245, 247, 252)
 
    scroll = SWC.ScrollViewer()
    scroll.VerticalScrollBarVisibility = SWC.ScrollBarVisibility.Auto
 
    main = SWC.StackPanel()
    main.Margin = SW.Thickness(22, 18, 22, 18)
 
    # Cabecalho
    t1 = SWC.TextBlock()
    t1.Text      = u"Criador de Tabelas"
    t1.FontSize  = 18
    t1.FontWeight= SW.FontWeights.Bold
    t1.Foreground= cor(20, 70, 160)
    t1.Margin    = SW.Thickness(0, 0, 0, 2)
    main.Children.Add(t1)
 
    t2 = SWC.TextBlock()
    t2.Text     = u"Armacao de Aberturas e Tela Soldada"
    t2.FontSize = 10
    t2.Foreground = cor(130, 130, 130)
    t2.Margin   = SW.Thickness(0, 0, 0, 10)
    main.Children.Add(t2)
 
    # ── TIPO ──────────────────────────────────────────
    main.Children.Add(secao(u"  TIPO DE TABELA"))
    cb_verg = criar_cb(u"Armacao de Aberturas  (Vergalhoes / Rebar)", True, 12)
    cb_tela = criar_cb(u"Armadura de Tela Soldada  (Fabric Sheets)",  True, 12)
    main.Children.Add(cb_verg)
    main.Children.Add(cb_tela)
 
    # ── NOMES ─────────────────────────────────────────
    main.Children.Add(secao(u"  NOME DAS TABELAS"))
    main.Children.Add(label(u"Nome da tabela de Vergalhoes:"))
    txt_nv = campo_texto(u"", u"TABELA DE ARMACAO ABERTURAS")
    main.Children.Add(txt_nv)
    main.Children.Add(label(u"Nome da tabela de Tela Soldada:"))
    txt_nt = campo_texto(u"", u"TABELA DE ARMADURA DE TELA SOLDADA")
    main.Children.Add(txt_nt)
 
    # ── CAMPOS VERGALHAO ──────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  -  VERGALHOES"))
    main.Children.Add(label(u"Marque os campos que deseja incluir na tabela:"))
    cbs_v = []
    for header, kw, default in CAMPOS_VERGALHAO:
        cb = criar_cb(u"{}".format(header), default)
        cbs_v.append((cb, header, kw))
        main.Children.Add(cb)
 
    # ── CAMPOS TELA ───────────────────────────────────
    main.Children.Add(secao(u"  CAMPOS  -  TELA SOLDADA"))
    main.Children.Add(label(u"Marque os campos que deseja incluir na tabela:"))
    cbs_t = []
    for header, kw, default in CAMPOS_TELA:
        cb = criar_cb(u"{}".format(header), default)
        cbs_t.append((cb, header, kw))
        main.Children.Add(cb)
 
    # ── FILTRO ────────────────────────────────────────
    main.Children.Add(secao(u"  FILTRO  (opcional)"))
    main.Children.Add(label(u"Filtrar por texto no campo LOCAL / Particao:"))
    main.Children.Add(label(u"(Deixe vazio para mostrar todos os itens)"))
    txt_filtro = campo_texto(u"ex: Terreo, P1, Cobertura...")
    main.Children.Add(txt_filtro)
 
    # ── OPCOES ────────────────────────────────────────
    main.Children.Add(secao(u"  OPCOES"))
    cb_folha = criar_cb(u"Inserir tabela na folha ativa  (apenas se for uma prancha)", False)
    main.Children.Add(cb_folha)
 
    # ── BOTOES ────────────────────────────────────────
    btns = SWC.StackPanel()
    btns.Orientation = SWC.Orientation.Horizontal
    btns.HorizontalAlignment = SW.HorizontalAlignment.Right
    btns.Margin = SW.Thickness(0, 16, 0, 4)
 
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
 
    def on_ok(s, e):
        campos_v = [(h, k, True) for cb, h, k in cbs_v if cb.IsChecked]
        campos_t = [(h, k, True) for cb, h, k in cbs_t if cb.IsChecked]
        filtro   = (txt_filtro.Text or u"").strip()
        resultado[0] = {
            'verg':       cb_verg.IsChecked and len(campos_v) > 0,
            'tela':       cb_tela.IsChecked and len(campos_t) > 0,
            'nome_v':     txt_nv.Text.strip() or u"TABELA DE ARMACAO ABERTURAS",
            'nome_t':     txt_nt.Text.strip() or u"TABELA DE ARMADURA DE TELA SOLDADA",
            'campos_v':   campos_v,
            'campos_t':   campos_t,
            'filtro':     filtro if filtro else None,
            'na_folha':   cb_folha.IsChecked,
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
                    is_vergalhao=True,
                )
                criadas.append(u"VERGALHOES: {}".format(nome))
                if opcoes['na_folha']:
                    inserir_na_folha(sched)
            except Exception as ex:
                erros.append(u"Vergalhoes: {}".format(str(ex)))
 
        if opcoes['tela']:
            try:
                sched, nome = criar_tabela(
                    opcoes['nome_t'], CAT_TELA,
                    opcoes['campos_t'],
                    filtro_texto=opcoes['filtro'],
                    campo_filtro_kw=u"hospedeiro",
                    is_vergalhao=False,
                )
                criadas.append(u"TELA SOLDADA: {}".format(nome))
                if opcoes['na_folha']:
                    inserir_na_folha(sched)
            except Exception as ex:
                erros.append(u"Tela Soldada: {}".format(str(ex)))
 
        t.Commit()
    except Exception as ex:
        t.RollBack()
        TaskDialog.Show(u"Erro Critico", str(ex))
        return
 
    partes = []
    if criadas:
        partes.append(u"Tabelas criadas com sucesso:\n" + u"\n".join(u"  - " + c for c in criadas))
    if erros:
        partes.append(u"Atencao - erros:\n" + u"\n".join(u"  - " + e for e in erros))
    if not partes:
        partes.append(u"Nenhuma tabela criada.\nMarque pelo menos um tipo e pelo menos um campo.")
 
    TaskDialog.Show(u"Criar Tabelas", u"\n\n".join(partes))
 
main()
