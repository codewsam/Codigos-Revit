# -*- coding: utf-8 -*-
"""
__title__   = "Renomear Paredes"
__author__  = "Samuel"
__version__ = "Versao 1.0"

Descricao:
    Renomeia automaticamente o parametro "Marca" (Mark) das paredes
    selecionadas pelo usuario, seguindo o padrao PR01, PR02, etc.

Fluxo:
    1. Script abre o formulario.
    2. Usuario define o numero inicial.
    3. Usuario clica em OK -> janela some -> usuario seleciona as paredes no modelo.
    4. Usuario finaliza a selecao pressionando ENTER ou clicando com botao direito.
    5. Script renomeia e exibe o resumo.
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
# VARIAVEIS GLOBAIS DO REVIT (injetadas automaticamente pelo pyRevit)
# ==============================================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# ==============================================================================
# FILTRO DE SELECAO: apenas paredes
# ==============================================================================
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
        # JANELA PRINCIPAL
        # ------------------------------------------------------------------
        self.Text            = "Renomear Paredes - Marca (Mark)"
        self.Size            = Size(420, 310)
        self.MinimumSize     = Size(420, 310)
        self.MaximizeBox     = False
        self.MinimizeBox     = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = FormStartPosition.CenterScreen
        self.BackColor       = Color.FromArgb(245, 245, 245)

        # ------------------------------------------------------------------
        # CABECALHO AZUL
        # ------------------------------------------------------------------
        painel_header           = Panel()
        painel_header.Size      = Size(420, 60)
        painel_header.Location  = Point(0, 0)
        painel_header.BackColor = Color.FromArgb(41, 128, 185)

        lbl_titulo           = Label()
        lbl_titulo.Text      = "  Renomear Paredes"
        lbl_titulo.Font      = Font("Segoe UI", 13, FontStyle.Bold)
        lbl_titulo.ForeColor = Color.White
        lbl_titulo.Size      = Size(400, 35)
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
            "3. Selecione as paredes no modelo (ENTER para finalizar)."
        )
        lbl_instrucao.Font      = Font("Segoe UI", 9, FontStyle.Regular)
        lbl_instrucao.ForeColor = Color.FromArgb(70, 70, 70)
        lbl_instrucao.Size      = Size(380, 60)
        lbl_instrucao.Location  = Point(20, 72)
        self.Controls.Add(lbl_instrucao)

        # ------------------------------------------------------------------
        # LABEL: PADRAO
        # ------------------------------------------------------------------
        lbl_padrao           = Label()
        lbl_padrao.Text      = "Padrao gerado: PR01, PR02, PR03 ..."
        lbl_padrao.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        lbl_padrao.ForeColor = Color.FromArgb(120, 120, 120)
        lbl_padrao.Size      = Size(380, 20)
        lbl_padrao.Location  = Point(20, 138)
        self.Controls.Add(lbl_padrao)

        # ------------------------------------------------------------------
        # LABEL: NUMERO INICIAL
        # ------------------------------------------------------------------
        lbl_numero           = Label()
        lbl_numero.Text      = "Numero inicial da sequencia:"
        lbl_numero.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        lbl_numero.ForeColor = Color.FromArgb(50, 50, 50)
        lbl_numero.Size      = Size(380, 22)
        lbl_numero.Location  = Point(20, 165)
        self.Controls.Add(lbl_numero)

        # ------------------------------------------------------------------
        # CAMPO DE TEXTO
        # ------------------------------------------------------------------
        self.txt_numero             = TextBox()
        self.txt_numero.Font        = Font("Segoe UI", 12, FontStyle.Regular)
        self.txt_numero.Size        = Size(100, 30)
        self.txt_numero.Location    = Point(20, 191)
        self.txt_numero.Text        = "1"
        self.txt_numero.TabIndex    = 0
        self.txt_numero.BorderStyle = BorderStyle.FixedSingle
        self.txt_numero.BackColor   = Color.White
        self.Controls.Add(self.txt_numero)

        # ------------------------------------------------------------------
        # LABEL: PREVIEW DINAMICO
        # ------------------------------------------------------------------
        self.lbl_preview           = Label()
        self.lbl_preview.Text      = "Preview: PR01, PR02, PR03 ..."
        self.lbl_preview.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        self.lbl_preview.ForeColor = Color.FromArgb(41, 128, 185)
        self.lbl_preview.Size      = Size(270, 22)
        self.lbl_preview.Location  = Point(130, 196)
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
        btn_ok.Location  = Point(20, 240)
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
        btn_cancelar.Location  = Point(235, 240)
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
            p1 = "PR{0:02d}".format(valor)
            p2 = "PR{0:02d}".format(valor + 1)
            p3 = "PR{0:02d}".format(valor + 2)
            self.lbl_preview.Text = "Preview: {0}, {1}, {2} ...".format(p1, p2, p3)
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
        "Clique nas paredes para selecionar (comecando em PR{0:02d}). "
        "Pressione ENTER ou clique com botao direito para finalizar."
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


def formatar_marca(numero):
    """
    Formata o numero seguindo o padrao PR##.

    Args:
        numero (int): Numero sequencial da parede.

    Returns:
        str: Ex: "PR01", "PR12", "PR100".
    """
    return "PR{0:02d}".format(numero)


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


def renomear_paredes(paredes, numero_inicial):
    """
    Executa a renomeacao dentro de uma Transaction.

    Args:
        paredes (list): Elementos Wall a renomear.
        numero_inicial (int): Numero inicial da sequencia.

    Returns:
        tuple: (renomeadas, erros, sem_param)
    """
    renomeadas = 0
    erros      = 0
    sem_param  = 0

    t = Transaction(doc, "Renomear Marca das Paredes")
    t.Start()

    try:
        for i, parede in enumerate(paredes):
            numero = numero_inicial + i
            marca  = formatar_marca(numero)

            try:
                sucesso = definir_parametro_mark(parede, marca)
                if sucesso:
                    renomeadas += 1
                else:
                    sem_param += 1
            except Exception:
                erros += 1

        t.Commit()

    except Exception as ex:
        t.RollBack()
        raise ex

    return renomeadas, erros, sem_param


def exibir_resultado(renomeadas, erros, sem_param, total):
    """
    Exibe o resumo final da operacao ao usuario.
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
      3. Renomeia as paredes selecionadas.
      4. Exibe o resumo.
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
    # ETAPA 3: Confirmacao rapida antes de aplicar
    # --------------------------------------------------------------------------
    p_inicio = formatar_marca(numero_inicial)
    p_fim    = formatar_marca(numero_inicial + len(paredes) - 1)

    confirmacao = MessageBox.Show(
        "Confirma a renomeacao de {0} parede(s)?\n\n"
        "Sequencia: {1}  ate  {2}".format(len(paredes), p_inicio, p_fim),
        "Confirmar Renomeacao",
        MessageBoxButtons.OKCancel,
        MessageBoxIcon.Question
    )

    if confirmacao != DialogResult.OK:
        return

    # --------------------------------------------------------------------------
    # ETAPA 4: Execucao da renomeacao
    # --------------------------------------------------------------------------
    try:
        renomeadas, erros, sem_param = renomear_paredes(paredes, numero_inicial)
        exibir_resultado(renomeadas, erros, sem_param, len(paredes))

    except Exception as ex:
        MessageBox.Show(
            "Erro critico durante a execucao:\n\n{0}\n\n"
            "A operacao foi cancelada (RollBack aplicado).".format(str(ex)),
            "Erro critico",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )


# ==============================================================================
# EXECUCAO
# ==============================================================================
if __name__ == "__main__":
    main()
