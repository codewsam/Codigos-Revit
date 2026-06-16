# -*- coding: utf-8 -*-
"""
__title__   = "Renomear Paredes"
__author__  = "Samuel"
__version__ = "Versao 1.3"

Descricao:
    Renomeia automaticamente o parametro "Marca" (Mark) das paredes
    selecionadas pelo usuario, seguindo o padrao PR01_NIVEL, PR02_NIVEL etc.,
    onde NIVEL e o nome do nivel (Base Constraint) da parede, normalizado
    (sem acentos, espacos ou caracteres especiais). Isso permite que a
    mesma numeracao (ex: PR01) se repita em niveis diferentes sem gerar
    Marca duplicada (ex: "PR01_TERREO" e "PR01_PLATIBANDA").

    As paredes sao agrupadas por NIVEL (pavimento, ex: TERREO, 1 PAVTO,
    LAJE, PLATIBANDA, etc.) e, dentro de cada nivel, ordenadas seguindo
    leitura tipo planta/bussola: de Oeste para Leste (menor X primeiro) e,
    dentro de cada coluna, de Norte para Sul (maior Y primeiro).

    A numeracao REINICIA no numero inicial informado pelo usuario para
    CADA nivel encontrado.

Fluxo:
    1. Script abre o formulario.
    2. Usuario define o numero inicial.
    3. Usuario clica em OK -> janela some -> usuario seleciona as paredes no modelo.
    4. Usuario finaliza a selecao pressionando ENTER ou clicando com botao direito.
    5. Script agrupa por nivel, ordena espacialmente, monta a Marca com o
       sufixo do nivel, renomeia e exibe o resumo.
"""

# ==============================================================================
# IMPORTS
# ==============================================================================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from Autodesk.Revit.DB import (
    Transaction,
    BuiltInCategory,
    BuiltInParameter,
    ElementId
)
from Autodesk.Revit.UI.Selection import (
    ObjectType,
    ISelectionFilter
)

import System
from System.Windows.Forms import (
    Form,
    Label,
    TextBox,
    Button,
    DialogResult,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
    FormBorderStyle,
    FormStartPosition,
    BorderStyle,
    Panel,
    FlatStyle
)
from System.Drawing import (
    Point,
    Size,
    Font,
    FontStyle,
    Color
)

# ==============================================================================
# VARIAVEIS GLOBAIS DO REVIT
# ==============================================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument



TOLERANCIA_COLUNA_FT = 3.28084



class FiltroParedes(ISelectionFilter):
    """
    Filtro para o PickObjects: permite selecionar apenas elementos
    da categoria Walls, ignorando qualquer outro tipo de elemento.
    """

    def AllowElement(self, elemento):
        """Retorna True somente se o elemento for uma parede."""
        if elemento is None:
            return False
        if elemento.Category is None:
            return False
        return elemento.Category.Id == ElementId(BuiltInCategory.OST_Walls)

    def AllowReference(self, ref, ponto):
        """Permite a referencia de qualquer elemento que passe pelo AllowElement."""
        return True


# ==============================================================================
# CLASSE: FORMULARIO WINDOWS FORMS
# ==============================================================================
class RenomearParedesForm(Form):
    """
    Janela de configuracao do script.
    O usuario define o numero inicial ANTES de selecionar as paredes.
    """

    def __init__(self):
        Form.__init__(self)
        self.numero_inicial = None  # Preenchido ao confirmar
        self._inicializar_componentes()

    def _inicializar_componentes(self):
        """Configura todos os componentes visuais do formulario."""

        # ------------------------------------------------------------------
        # jANELA PRINCIPLA
        # ------------------------------------------------------------------
        self.Text            = "Renomear Paredes - Marca (Mark)"
        self.Size            = Size(440, 340)
        self.MinimumSize     = Size(440, 340)
        self.MaximizeBox     = False
        self.MinimizeBox     = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.BackColor       = Color.FromArgb(245, 245, 245)

        # ------------------------------------------------------------------
        # CABECALHO AZUL
        # ------------------------------------------------------------------
        painel_header           = Panel()
        painel_header.Size      = Size(440, 60)
        painel_header.Location  = Point(0, 0)
        painel_header.BackColor = Color.FromArgb(41, 128, 185)

        lbl_titulo           = Label()
        lbl_titulo.Text      = "  Renomear Paredes"
        lbl_titulo.Font      = Font("Segoe UI", 13, FontStyle.Bold)
        lbl_titulo.ForeColor = Color.White
        lbl_titulo.Size      = Size(420, 35)
        lbl_titulo.Location  = Point(10, 12)
        painel_header.Controls.Add(lbl_titulo)
        self.Controls.Add(painel_header)

        # ------------------------------------------------------------------
        # INSTRUCOES DE USO
        # ------------------------------------------------------------------
        lbl_instrucao           = Label()
        lbl_instrucao.Text      = (
            "1. Informe o numero inicial abaixo.\n"
            "2. Clique em OK.\n"
            "3. Selecione as paredes no modelo (ENTER para finalizar).\n"
            "As paredes serao agrupadas por NIVEL e numeradas reiniciando\n"
            "a contagem em cada nivel. O nome do nivel (Base Constraint)\n"
            "e adicionado como sufixo na Marca, ex: PR01_TERREO."
        )
        lbl_instrucao.Font      = Font("Segoe UI", 9, FontStyle.Regular)
        lbl_instrucao.ForeColor = Color.FromArgb(70, 70, 70)
        lbl_instrucao.Size      = Size(400, 90)
        lbl_instrucao.Location  = Point(20, 72)
        self.Controls.Add(lbl_instrucao)

        # ------------------------------------------------------------------
        # LABEL: PADRAO
        # ------------------------------------------------------------------
        lbl_padrao           = Label()
        lbl_padrao.Text      = "Padrao gerado: PR01_NIVEL, PR02_NIVEL ... (reinicia por nivel)"
        lbl_padrao.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        lbl_padrao.ForeColor = Color.FromArgb(120, 120, 120)
        lbl_padrao.Size      = Size(400, 20)
        lbl_padrao.Location  = Point(20, 165)
        self.Controls.Add(lbl_padrao)

        # ------------------------------------------------------------------
        # LABEL: NUMERO INICIAL
        # ------------------------------------------------------------------
        lbl_numero           = Label()
        lbl_numero.Text      = "Numero inicial da sequencia (em cada nivel):"
        lbl_numero.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        lbl_numero.ForeColor = Color.FromArgb(50, 50, 50)
        lbl_numero.Size      = Size(400, 22)
        lbl_numero.Location  = Point(20, 190)
        self.Controls.Add(lbl_numero)

        # ------------------------------------------------------------------
        # CAMPO DE TEXTO
        # ------------------------------------------------------------------
        self.txt_numero             = TextBox()
        self.txt_numero.Font        = Font("Segoe UI", 12, FontStyle.Regular)
        self.txt_numero.Size        = Size(100, 30)
        self.txt_numero.Location    = Point(20, 216)
        self.txt_numero.Text        = "1"
        self.txt_numero.TabIndex    = 0
        self.txt_numero.BorderStyle = BorderStyle.FixedSingle
        self.txt_numero.BackColor   = Color.White
        self.Controls.Add(self.txt_numero)

        # ------------------------------------------------------------------
        # LABEL: PREVIEW DINAMICO
        # ------------------------------------------------------------------
        self.lbl_preview           = Label()
        self.lbl_preview.Text      = "Preview: PR01_<NIVEL>, PR02_<NIVEL> ..."
        self.lbl_preview.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        self.lbl_preview.ForeColor = Color.FromArgb(41, 128, 185)
        self.lbl_preview.Size      = Size(300, 22)
        self.lbl_preview.Location  = Point(130, 221)
        self.Controls.Add(self.lbl_preview)

        # Evento: atualiza preview ao digitar
        self.txt_numero.TextChanged += self._atualizar_preview

        # ------------------------------------------------------------------
        # BOTAO OK
        # ------------------------------------------------------------------
        btn_ok           = Button()
        btn_ok.Text      = "Selecionar Paredes"
        btn_ok.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        btn_ok.Size      = Size(200, 36)
        btn_ok.Location  = Point(20, 265)
        btn_ok.BackColor = Color.FromArgb(41, 128, 185)
        btn_ok.ForeColor = Color.White
        btn_ok.FlatStyle = FlatStyle.Flat
        btn_ok.FlatAppearance.BorderSize = 0
        btn_ok.TabIndex  = 1
        btn_ok.Click    += self._btn_ok_click
        self.Controls.Add(btn_ok)

        # ------------------------------------------------------------------
        # BOTAO CANCELAR
        # ------------------------------------------------------------------
        btn_cancelar           = Button()
        btn_cancelar.Text      = "Cancelar"
        btn_cancelar.Font      = Font("Segoe UI", 10, FontStyle.Regular)
        btn_cancelar.Size      = Size(110, 36)
        btn_cancelar.Location  = Point(235, 265)
        btn_cancelar.BackColor = Color.FromArgb(200, 200, 200)
        btn_cancelar.ForeColor = Color.FromArgb(50, 50, 50)
        btn_cancelar.FlatStyle = FlatStyle.Flat
        btn_cancelar.FlatAppearance.BorderSize = 0
        btn_cancelar.TabIndex  = 2
        btn_cancelar.Click    += self._btn_cancelar_click
        self.Controls.Add(btn_cancelar)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancelar

        self.txt_numero.Select()
        self.txt_numero.SelectAll()

    # ------------------------------------------------------------------
    # EVENTOS
    # ------------------------------------------------------------------

    def _atualizar_preview(self, sender, e):
        """Atualiza o preview conforme o usuario digita."""
        try:
            valor = int(self.txt_numero.Text.strip())
            if valor < 0:
                self.lbl_preview.Text = "Numero invalido"
                return
            p1 = "PR{0:02d}_<NIVEL>".format(valor)
            p2 = "PR{0:02d}_<NIVEL>".format(valor + 1)
            self.lbl_preview.Text = "Preview: {0}, {1} ...".format(p1, p2)
        except (ValueError, System.FormatException):
            self.lbl_preview.Text = "Digite um numero valido"

    def _btn_ok_click(self, sender, e):
        """Valida a entrada e confirma."""
        texto = self.txt_numero.Text.strip()

        if not texto:
            MessageBox.Show(
                "Por favor, informe o numero inicial.",
                "Campo obrigatorio",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            self.txt_numero.Focus()
            return

        try:
            valor = int(texto)
        except (ValueError, System.FormatException):
            MessageBox.Show(
                "O valor informado nao e um numero inteiro valido.\nExemplo: 1, 5, 12",
                "Valor invalido",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            self.txt_numero.Focus()
            self.txt_numero.SelectAll()
            return

        if valor < 0:
            MessageBox.Show(
                "O numero inicial deve ser maior ou igual a zero.",
                "Valor invalido",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            self.txt_numero.Focus()
            self.txt_numero.SelectAll()
            return

        self.numero_inicial = valor
        self.DialogResult   = DialogResult.OK
        self.Close()

    def _btn_cancelar_click(self, sender, e):
        """Cancela a operacao."""
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ==============================================================================
# FUNCOES AUXILIARES
# ==============================================================================

def selecionar_paredes_no_modelo(numero_inicial):
    """
    Abre o modo de selecao interativa no Revit, permitindo que o usuario
    clique nas paredes diretamente no modelo.

    Args:
        numero_inicial (int): Usado apenas para exibir a dica na barra de status.

    Returns:
        list: Lista de elementos Wall selecionados. Vazia se cancelado.
    """
    filtro = FiltroParedes()
    dica   = (
        "Clique nas paredes para selecionar (numeracao reinicia em PR{0:02d} "
        "para cada nivel, com sufixo do nivel). Pressione ENTER ou clique "
        "com botao direito para finalizar."
    ).format(numero_inicial)

    try:
        # PickObjects retorna uma colecao de References
        referencias = uidoc.Selection.PickObjects(
            ObjectType.Element,
            filtro,
            dica
        )

        paredes = []
        for ref in referencias:
            elemento = doc.GetElement(ref.ElementId)
            if elemento is not None:
                paredes.append(elemento)

        return paredes

    except System.OperationCanceledException:
        # Usuario pressionou ESC — cancelamento limpo
        return []
    except Exception:
        # Qualquer outro erro na selecao
        return []


def obter_nome_nivel(parede):
    """
    Retorna o nome do nivel (pavimento) ao qual a parede pertence,
    usando o parametro de Base Constraint (Restricao da Base) da parede.
    Ex: "TERREO", "1 PAVTO", "LAJE COBERTURA", "PLATIBANDA", etc.

    Args:
        parede: Elemento Wall do Revit.

    Returns:
        str: Nome do nivel, ou "SEM NIVEL" se nao for possivel determinar.
    """
    try:
        level_param = parede.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        if level_param is None:
            return "SEM NIVEL"

        level_id = level_param.AsElementId()
        if level_id is None or level_id == ElementId.InvalidElementId:
            return "SEM NIVEL"

        level = doc.GetElement(level_id)
        if level is None:
            return "SEM NIVEL"

        return level.Name

    except Exception:
        return "SEM NIVEL"


def obter_elevacao_nivel(parede):
    """
    Retorna a elevacao (Z) do nivel base da parede, usada para
    ordenar os grupos de niveis do mais baixo para o mais alto.

    Args:
        parede: Elemento Wall do Revit.

    Returns:
        float: Elevacao em pes. 0.0 se nao for possivel determinar.
    """
    try:
        level_param = parede.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
        if level_param is None:
            return 0.0

        level_id = level_param.AsElementId()
        if level_id is None or level_id == ElementId.InvalidElementId:
            return 0.0

        level = doc.GetElement(level_id)
        if level is None:
            return 0.0

        return level.Elevation

    except Exception:
        return 0.0


def agrupar_paredes_por_nivel(paredes):
    """
    Agrupa as paredes por nivel (pavimento/laje), preservando a ordem
    de elevacao (do nivel mais baixo para o mais alto).

    Paredes sem nivel identificavel sao agrupadas em "SEM NIVEL" e
    posicionadas por ultimo.

    Args:
        paredes (list): Elementos Wall.

    Returns:
        list: Lista de tuplas (nome_nivel, elevacao, [paredes]),
              ordenada por elevacao crescente (SEM NIVEL ao final).
    """
    grupos = {}

    for parede in paredes:
        nome_nivel = obter_nome_nivel(parede)
        elevacao   = obter_elevacao_nivel(parede)

        if nome_nivel not in grupos:
            grupos[nome_nivel] = {"elevacao": elevacao, "itens": []}

        grupos[nome_nivel]["itens"].append(parede)

    resultado = []
    for nome_nivel, dados in grupos.items():
        resultado.append((nome_nivel, dados["elevacao"], dados["itens"]))

    # Ordena por elevacao crescente; "SEM NIVEL" sempre por ultimo
    def chave_ordenacao(grupo):
        nome, elevacao, _itens = grupo
        if nome == "SEM NIVEL":
            return (1, 0.0)
        return (0, elevacao)

    resultado.sort(key=chave_ordenacao)

    return resultado


def obter_ponto_referencia(parede):
    """
    Retorna o ponto medio (X, Y) da linha de localizacao da parede,
    usado para ordenacao espacial.

    Args:
        parede: Elemento Wall do Revit.

    Returns:
        tuple: (x, y) em pes. (0.0, 0.0) se a parede nao tiver Location.Curve.
    """
    loc = parede.Location
    if hasattr(loc, "Curve"):
        curve = loc.Curve
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        x = (p0.X + p1.X) / 2.0
        y = (p0.Y + p1.Y) / 2.0
        return (x, y)
    return (0.0, 0.0)


def ordenar_paredes_espacialmente(paredes, tolerancia_coluna_ft=TOLERANCIA_COLUNA_FT):
    """
    Ordena as paredes seguindo leitura tipo planta/bussola:
    agrupadas em colunas de Oeste para Leste (menor X primeiro)
    e, dentro de cada coluna, de Norte para Sul (maior Y primeiro).

    Args:
        paredes (list): Elementos Wall.
        tolerancia_coluna_ft (float): Diferenca de X (em pes) para
            considerar duas paredes na mesma "coluna". Default ~1 metro.

    Returns:
        list: Paredes ordenadas.
    """
    pontos = [(p, obter_ponto_referencia(p)) for p in paredes]

    # Ordena por X crescente (Oeste primeiro)
    pontos.sort(key=lambda item: item[1][0])

    colunas = []
    for parede, (x, y) in pontos:
        encaixou = False
        for coluna in colunas:
            if abs(coluna["x_ref"] - x) <= tolerancia_coluna_ft:
                coluna["itens"].append((parede, x, y))
                encaixou = True
                break
        if not encaixou:
            colunas.append({"x_ref": x, "itens": [(parede, x, y)]})

    resultado = []
    for coluna in colunas:
        # Dentro da coluna: Y decrescente (Norte para Sul)
        itens_ordenados = sorted(coluna["itens"], key=lambda t: -t[2])
        resultado.extend([item[0] for item in itens_ordenados])

    return resultado


def remover_acentos(texto):
    """
    Remove acentos comuns do portugues, convertendo para caracteres
    ASCII equivalentes. Implementado sem dependencia de modulos
    externos (ex: unicodedata) para maior compatibilidade com IronPython 2.

    Args:
        texto (str): Texto de entrada, possivelmente acentuado.

    Returns:
        unicode: Texto sem acentos.
    """
    if not texto:
        return u""

    mapa_acentos = {
        u"\u00e1": u"a", u"\u00e0": u"a", u"\u00e3": u"a", u"\u00e2": u"a", u"\u00e4": u"a",
        u"\u00e9": u"e", u"\u00e8": u"e", u"\u00ea": u"e", u"\u00eb": u"e",
        u"\u00ed": u"i", u"\u00ec": u"i", u"\u00ee": u"i", u"\u00ef": u"i",
        u"\u00f3": u"o", u"\u00f2": u"o", u"\u00f5": u"o", u"\u00f4": u"o", u"\u00f6": u"o",
        u"\u00fa": u"u", u"\u00f9": u"u", u"\u00fb": u"u", u"\u00fc": u"u",
        u"\u00e7": u"c", u"\u00f1": u"n",
        u"\u00c1": u"A", u"\u00c0": u"A", u"\u00c3": u"A", u"\u00c2": u"A", u"\u00c4": u"A",
        u"\u00c9": u"E", u"\u00c8": u"E", u"\u00ca": u"E", u"\u00cb": u"E",
        u"\u00cd": u"I", u"\u00cc": u"I", u"\u00ce": u"I", u"\u00cf": u"I",
        u"\u00d3": u"O", u"\u00d2": u"O", u"\u00d5": u"O", u"\u00d4": u"O", u"\u00d6": u"O",
        u"\u00da": u"U", u"\u00d9": u"U", u"\u00db": u"U", u"\u00dc": u"U",
        u"\u00c7": u"C", u"\u00d1": u"N",
    }

    resultado = u""
    for caractere in texto:
        resultado += mapa_acentos.get(caractere, caractere)

    return resultado


def sanitizar_sufixo_nivel(nome_nivel):
    """
    Converte o nome do nivel num sufixo limpo para compor a Marca,
    removendo acentos, espacos e caracteres especiais e deixando
    tudo em caixa alta.

    Ex: "1 PAVTO" -> "1PAVTO" ; "Laje Cobertura" -> "LAJECOBERTURA"
        "Platibanda" -> "PLATIBANDA"

    Args:
        nome_nivel (str): Nome do nivel (Base Constraint) da parede.

    Returns:
        str: Sufixo normalizado, sem espacos/acentos/caracteres especiais.
    """
    if not nome_nivel:
        return ""

    texto = remover_acentos(nome_nivel)
    texto = texto.upper()

    sufixo = ""
    for caractere in texto:
        if caractere.isalnum():
            sufixo += caractere
        # demais caracteres (espacos, hifens, underscores, etc.) sao descartados

    return sufixo


def formatar_marca(numero, nome_nivel=""):
    """
    Formata o numero seguindo o padrao PR##_NIVEL.

    O sufixo do nivel permite que a mesma sequencia numerica (ex: PR01)
    se repita em niveis diferentes sem gerar Marca duplicada no modelo,
    pois o valor final ja inclui o nivel de origem da parede.
    Ex: "PR01_TERREO", "PR01_PLATIBANDA", "PR12_1PAVTO".

    Args:
        numero (int): Numero sequencial da parede.
        nome_nivel (str): Nome do nivel (Base Constraint) da parede.

    Returns:
        str: Marca formatada.
    """
    base   = "PR{0:02d}".format(numero)
    sufixo = sanitizar_sufixo_nivel(nome_nivel)

    if sufixo:
        return "{0}_{1}".format(base, sufixo)
    return base


def definir_parametro_mark(elemento, valor):
    """
    Define o valor do parametro Mark (Marca) de um elemento Revit.

    Args:
        elemento: Elemento Wall do Revit.
        valor (str): Valor a ser atribuido.

    Returns:
        bool: True se bem-sucedido, False caso contrario.
    """
    try:
        param = elemento.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)

        # Fallback por nome (pt / en)
        if param is None:
            param = elemento.LookupParameter("Marca")
        if param is None:
            param = elemento.LookupParameter("Mark")

        if param is None or param.IsReadOnly:
            return False

        param.Set(valor)
        return True

    except Exception:
        return False


def renomear_paredes(grupos_por_nivel, numero_inicial):
    """
    Executa a renomeacao dentro de uma Transaction.
    A numeracao reinicia em "numero_inicial" para CADA nivel, e a Marca
    final de cada parede recebe o sufixo do respectivo nivel.

    Args:
        grupos_por_nivel (list): Lista de tuplas
            (nome_nivel, elevacao, [paredes_ordenadas]).
        numero_inicial (int): Numero inicial da sequencia em cada nivel.

    Returns:
        tuple: (renomeadas, erros, sem_param, resumo_por_nivel)
            resumo_por_nivel (list): Lista de tuplas
                (nome_nivel, qtd_paredes, marca_inicial, marca_final).
    """
    renomeadas = 0
    erros      = 0
    sem_param  = 0
    resumo_por_nivel = []

    t = Transaction(doc, "Renomear Marca das Paredes")
    t.Start()

    try:
        for nome_nivel, _elevacao, paredes in grupos_por_nivel:

            if not paredes:
                continue

            for i, parede in enumerate(paredes):
                numero = numero_inicial + i
                marca  = formatar_marca(numero, nome_nivel)

                try:
                    sucesso = definir_parametro_mark(parede, marca)
                    if sucesso:
                        renomeadas += 1
                    else:
                        sem_param += 1
                except Exception:
                    erros += 1

            marca_inicial = formatar_marca(numero_inicial, nome_nivel)
            marca_final   = formatar_marca(numero_inicial + len(paredes) - 1, nome_nivel)
            resumo_por_nivel.append((nome_nivel, len(paredes), marca_inicial, marca_final))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        raise ex

    return renomeadas, erros, sem_param, resumo_por_nivel


def exibir_resultado(renomeadas, erros, sem_param, total, resumo_por_nivel):
    """
    Exibe o resumo final da operacao, detalhando a sequencia
    gerada em cada nivel.
    """
    linhas = [
        "Operacao concluida!",
        "",
        "Paredes processadas : {0}".format(total),
        "Renomeadas          : {0}".format(renomeadas),
    ]

    if sem_param > 0:
        linhas.append("Sem parametro Mark  : {0}".format(sem_param))
    if erros > 0:
        linhas.append("Com erro            : {0}".format(erros))

    linhas.append("")
    linhas.append("Detalhamento por nivel:")
    for nome_nivel, qtd, marca_ini, marca_fim in resumo_por_nivel:
        linhas.append("  {0}: {1} parede(s)  ->  {2} a {3}".format(
            nome_nivel, qtd, marca_ini, marca_fim
        ))

    icone = MessageBoxIcon.Information if erros == 0 else MessageBoxIcon.Warning

    MessageBox.Show(
        "\n".join(linhas),
        "Renomear Paredes - Concluido",
        MessageBoxButtons.OK,
        icone
    )


# ==============================================================================
# FUNCAO PRINCIPAL
# ==============================================================================

def main():
    """
    Fluxo principal:
      1. Abre o formulario para o usuario definir o numero inicial.
      2. Fecha o formulario e ativa a selecao interativa no modelo.
      3. Agrupa as paredes por NIVEL (pavimento, laje, etc.).
      4. Dentro de cada nivel, ordena espacialmente
         (Oeste->Leste, Norte->Sul).
      5. Renomeia as paredes, reiniciando a numeracao em cada nivel e
         compondo a Marca com o sufixo do nivel (ex: PR01_TERREO).
      6. Exibe o resumo detalhado por nivel.
    """

    # --------------------------------------------------------------------------
    # ETAPA 1: Formulario de configuracao
    # --------------------------------------------------------------------------
    formulario = RenomearParedesForm()
    resultado  = formulario.ShowDialog()

    if resultado != DialogResult.OK or formulario.numero_inicial is None:
        # Usuario cancelou — encerra silenciosamente
        return

    numero_inicial = formulario.numero_inicial

    # --------------------------------------------------------------------------
    # ETAPA 2: Selecao interativa das paredes no modelo
    # --------------------------------------------------------------------------
    # O formulario ja foi fechado; o Revit retoma o foco automaticamente
    paredes = selecionar_paredes_no_modelo(numero_inicial)

    if not paredes:
        MessageBox.Show(
            "Nenhuma parede foi selecionada.\nOperacao cancelada.",
            "Selecao vazia",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
        return

    # --------------------------------------------------------------------------
    # ETAPA 3: Agrupamento por nivel (pavimento, laje, etc.)
    # --------------------------------------------------------------------------
    grupos_por_nivel = agrupar_paredes_por_nivel(paredes)

    # --------------------------------------------------------------------------
    # ETAPA 4: Ordenacao espacial dentro de cada nivel
    #          (Oeste -> Leste, Norte -> Sul)
    # --------------------------------------------------------------------------
    grupos_ordenados = []
    for nome_nivel, elevacao, paredes_do_nivel in grupos_por_nivel:
        paredes_ordenadas = ordenar_paredes_espacialmente(paredes_do_nivel)
        grupos_ordenados.append((nome_nivel, elevacao, paredes_ordenadas))

    # --------------------------------------------------------------------------
    # ETAPA 5: Confirmacao rapida antes de aplicar
    # --------------------------------------------------------------------------
    linhas_preview = [
        "Confirma a renomeacao de {0} parede(s)?".format(len(paredes)),
        "",
        "A numeracao reinicia em {0:02d} para cada nivel, com sufixo do nivel.".format(numero_inicial),
        "Ordem: Oeste -> Leste, Norte -> Sul",
        "",
        "Niveis encontrados:",
    ]
    for nome_nivel, _elevacao, itens in grupos_ordenados:
        marca_ini_nivel = formatar_marca(numero_inicial, nome_nivel)
        marca_fim_nivel = formatar_marca(numero_inicial + len(itens) - 1, nome_nivel)
        linhas_preview.append("  - {0}: {1} parede(s) ({2} a {3})".format(
            nome_nivel,
            len(itens),
            marca_ini_nivel,
            marca_fim_nivel
        ))

    confirmacao = MessageBox.Show(
        "\n".join(linhas_preview),
        "Confirmar Renomeacao",
        MessageBoxButtons.OKCancel,
        MessageBoxIcon.Question
    )

    if confirmacao != DialogResult.OK:
        return

    # --------------------------------------------------------------------------
    # ETAPA 6: Execucao da renomeacao
    # --------------------------------------------------------------------------
    try:
        renomeadas, erros, sem_param, resumo_por_nivel = renomear_paredes(
            grupos_ordenados, numero_inicial
        )
        exibir_resultado(renomeadas, erros, sem_param, len(paredes), resumo_por_nivel)

    except Exception as ex:
        MessageBox.Show(
            "Erro critico durante a execucao:\n\n{0}\n\n"
            "A operacao foi cancelada (RollBack aplicado).".format(str(ex)),
            "Erro critico",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )


# ==============================================================================
# ==============================================================================
if __name__ == "__main__":
    main()
