# -*- coding: utf-8 -*-

__title__   = "nomear Paredes"
__author__  = "Samuel"
__version__ = "Versao 1.0"



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
    ElementId
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

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
    AnchorStyles,
    BorderStyle,
    Panel,
    FlatStyle
)
from System.Drawing import (
    Point,
    Size,
    Font,
    FontStyle,
    Color,
    ContentAlignment
)

# ==============================================================================
# VARIAVEIS GLOBAIS DO REVIT (injetadas automaticamente pelo pyRevit)
# ==============================================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# ==============================================================================
# CLASSE: FORMULARIO WINDOWS FORMS
# ==============================================================================
class RenomearParedesForm(Form):
    """
    Janela de entrada de dados para o script de renomeacao de paredes.
    Solicita ao usuario o numero inicial da sequencia.
    """

    def __init__(self, total_paredes):
        """
        Inicializa o formulario com o numero total de paredes selecionadas.

        Args:
            total_paredes (int): Quantidade de paredes filtradas na selecao.
        """
        Form.__init__(self)
        self.numero_inicial = None  # Resultado retornado apos confirmacao
        self._total_paredes = total_paredes
        self._inicializar_componentes()

    def _inicializar_componentes(self):
        """Configura todos os componentes visuais do formulario."""

        # ------------------------------------------------------------------
        # CONFIGURACAO DA JANELA PRINCIPAL
        # ------------------------------------------------------------------
        self.Text             = "Renomear Paredes - Marca (Mark)"
        self.Size             = Size(400, 280)
        self.MinimumSize      = Size(400, 280)
        self.MaximizeBox      = False
        self.MinimizeBox      = False
        self.FormBorderStyle  = FormBorderStyle.FixedDialog
        self.StartPosition    = FormStartPosition.CenterScreen
        self.BackColor        = Color.FromArgb(245, 245, 245)

        # ------------------------------------------------------------------
        # PAINEL DE CABECALHO
        # ------------------------------------------------------------------
        painel_header = Panel()
        painel_header.Size      = Size(400, 60)
        painel_header.Location  = Point(0, 0)
        painel_header.BackColor = Color.FromArgb(41, 128, 185)

        lbl_titulo = Label()
        lbl_titulo.Text      = "  Renomear Paredes"
        lbl_titulo.Font      = Font("Segoe UI", 13, FontStyle.Bold)
        lbl_titulo.ForeColor = Color.White
        lbl_titulo.Size      = Size(380, 35)
        lbl_titulo.Location  = Point(10, 12)
        painel_header.Controls.Add(lbl_titulo)

        self.Controls.Add(painel_header)

        # ------------------------------------------------------------------
        # LABEL: INFORMACAO DE SELECAO
        # ------------------------------------------------------------------
        lbl_info = Label()
        lbl_info.Text      = "Paredes selecionadas: {0}".format(self._total_paredes)
        lbl_info.Font      = Font("Segoe UI", 9, FontStyle.Regular)
        lbl_info.ForeColor = Color.FromArgb(80, 80, 80)
        lbl_info.Size      = Size(360, 22)
        lbl_info.Location  = Point(20, 75)
        self.Controls.Add(lbl_info)

        # ------------------------------------------------------------------
        # LABEL: PADRAO DE NOMENCLATURA
        # ------------------------------------------------------------------
        lbl_padrao = Label()
        lbl_padrao.Text      = "Padrao gerado: PR01, PR02, PR03 ..."
        lbl_padrao.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        lbl_padrao.ForeColor = Color.FromArgb(120, 120, 120)
        lbl_padrao.Size      = Size(360, 20)
        lbl_padrao.Location  = Point(20, 100)
        self.Controls.Add(lbl_padrao)

        # ------------------------------------------------------------------
        # LABEL: CAMPO NUMERO INICIAL
        # ------------------------------------------------------------------
        lbl_numero = Label()
        lbl_numero.Text      = "Numero inicial da sequencia:"
        lbl_numero.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        lbl_numero.ForeColor = Color.FromArgb(50, 50, 50)
        lbl_numero.Size      = Size(360, 22)
        lbl_numero.Location  = Point(20, 135)
        self.Controls.Add(lbl_numero)

        # ------------------------------------------------------------------
        # CAMPO DE TEXTO: ENTRADA DO NUMERO INICIAL
        # ------------------------------------------------------------------
        self.txt_numero = TextBox()
        self.txt_numero.Font      = Font("Segoe UI", 12, FontStyle.Regular)
        self.txt_numero.Size      = Size(100, 30)
        self.txt_numero.Location  = Point(20, 160)
        self.txt_numero.Text      = "1"
        self.txt_numero.TabIndex  = 0
        self.txt_numero.BorderStyle = BorderStyle.FixedSingle
        self.txt_numero.BackColor = Color.White
        self.Controls.Add(self.txt_numero)

        # ------------------------------------------------------------------
        # LABEL: PREVIEW DO RESULTADO
        # ------------------------------------------------------------------
        self.lbl_preview = Label()
        self.lbl_preview.Text      = "Preview: PR01, PR02 ..."
        self.lbl_preview.Font      = Font("Segoe UI", 9, FontStyle.Italic)
        self.lbl_preview.ForeColor = Color.FromArgb(41, 128, 185)
        self.lbl_preview.Size      = Size(230, 22)
        self.lbl_preview.Location  = Point(130, 165)
        self.Controls.Add(self.lbl_preview)

        # Atualiza preview conforme o usuario digita
        self.txt_numero.TextChanged += self._atualizar_preview

        # ------------------------------------------------------------------
        # BOTAO: OK
        # ------------------------------------------------------------------
        btn_ok = Button()
        btn_ok.Text      = "OK"
        btn_ok.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        btn_ok.Size      = Size(100, 35)
        btn_ok.Location  = Point(170, 205)
        btn_ok.BackColor = Color.FromArgb(41, 128, 185)
        btn_ok.ForeColor = Color.White
        btn_ok.FlatStyle = FlatStyle.Flat
        btn_ok.FlatAppearance.BorderSize = 0
        btn_ok.TabIndex  = 1
        btn_ok.Click    += self._btn_ok_click
        self.Controls.Add(btn_ok)

        # ------------------------------------------------------------------
        # BOTAO: CANCELAR
        # ------------------------------------------------------------------
        btn_cancelar = Button()
        btn_cancelar.Text      = "Cancelar"
        btn_cancelar.Font      = Font("Segoe UI", 10, FontStyle.Regular)
        btn_cancelar.Size      = Size(100, 35)
        btn_cancelar.Location  = Point(280, 205)
        btn_cancelar.BackColor = Color.FromArgb(200, 200, 200)
        btn_cancelar.ForeColor = Color.FromArgb(50, 50, 50)
        btn_cancelar.FlatStyle = FlatStyle.Flat
        btn_cancelar.FlatAppearance.BorderSize = 0
        btn_cancelar.TabIndex  = 2
        btn_cancelar.Click    += self._btn_cancelar_click
        self.Controls.Add(btn_cancelar)

        # Define botao padrao (Enter) e cancelamento (Escape)
        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancelar

        # Foco inicial no campo de numero
        self.txt_numero.Select()
        self.txt_numero.SelectAll()

    def _atualizar_preview(self, sender, e):
        """
        Atualiza o label de preview conforme o usuario digita o numero inicial.
        """
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
        """
        Valida a entrada e confirma o formulario ao clicar em OK.
        """
        texto = self.txt_numero.Text.strip()

        # Validacao: campo vazio
        if not texto:
            MessageBox.Show(
                "Por favor, informe o numero inicial da sequencia.",
                "Campo obrigatorio",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            self.txt_numero.Focus()
            return

        # Validacao: deve ser numero inteiro
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

        # Validacao: numero deve ser positivo
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

        # Armazena resultado e fecha
        self.numero_inicial = valor
        self.DialogResult   = DialogResult.OK
        self.Close()

    def _btn_cancelar_click(self, sender, e):
        """
        Cancela a operacao ao clicar em Cancelar.
        """
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ==============================================================================
# FUNCOES AUXILIARES
# ==============================================================================

def obter_paredes_selecionadas():
    """
    Obtem e filtra as paredes presentes na selecao atual do usuario.

    Returns:
        list: Lista de elementos Wall selecionados, ou lista vazia.
    """
    selecao   = uidoc.Selection.GetElementIds()
    paredes   = []

    for eid in selecao:
        elemento = doc.GetElement(eid)
        if elemento is None:
            continue
        # Filtra apenas elementos da categoria "Walls"
        if elemento.Category is not None:
            if elemento.Category.Id == ElementId(BuiltInCategory.OST_Walls):
                paredes.append(elemento)

    return paredes


def formatar_marca(numero):
    """
    Formata o numero da parede seguindo o padrao PR##.

    Args:
        numero (int): Numero da parede na sequencia.

    Returns:
        str: String formatada, ex: "PR01", "PR12", "PR100".
    """
    return "PR{0:02d}".format(numero)


def definir_parametro_mark(elemento, valor):
    """
    Define o valor do parametro Mark (Marca) de um elemento.

    Args:
        elemento: Elemento Revit (Wall).
        valor (str): Valor a ser atribuido ao parametro Mark.

    Returns:
        bool: True se bem-sucedido, False caso contrario.
    """
    try:
        param = elemento.get_Parameter(
            Autodesk.Revit.DB.BuiltInParameter.ALL_MODEL_MARK
        )
        if param is None:
            # Tenta buscar por nome como fallback
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
    Executa a renomeacao de todas as paredes dentro de uma Transaction.

    Args:
        paredes (list): Lista de elementos Wall.
        numero_inicial (int): Numero inicial da sequencia.

    Returns:
        tuple: (renomeadas, erros) — contagens de sucesso e falha.
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
    Exibe uma mensagem ao usuario com o resumo da operacao.

    Args:
        renomeadas (int): Paredes renomeadas com sucesso.
        erros      (int): Paredes com erros durante o processo.
        sem_param  (int): Paredes sem o parametro Mark disponivel.
        total      (int): Total de paredes processadas.
    """
    linhas = []
    linhas.append("Operacao concluida com sucesso!")
    linhas.append("")
    linhas.append("Paredes processadas : {0}".format(total))
    linhas.append("Renomeadas          : {0}".format(renomeadas))

    if sem_param > 0:
        linhas.append("Sem parametro Mark  : {0}".format(sem_param))
    if erros > 0:
        linhas.append("Com erro            : {0}".format(erros))

    mensagem = "\n".join(linhas)

    icone = MessageBoxIcon.Information if erros == 0 else MessageBoxIcon.Warning

    MessageBox.Show(
        mensagem,
        "Renomear Paredes - Concluido",
        MessageBoxButtons.OK,
        icone
    )


# ==============================================================================
# FUNCAO PRINCIPAL
# ==============================================================================

def main():
    """
    Ponto de entrada principal do script.
    Coordena validacao de selecao, abertura do formulario e execucao.
    """

    # --------------------------------------------------------------------------
    # 1. OBTER PAREDES SELECIONADAS
    # --------------------------------------------------------------------------
    paredes = obter_paredes_selecionadas()

    if not paredes:
        MessageBox.Show(
            "Nenhuma parede foi encontrada na selecao atual.\n\n"
            "Por favor, selecione uma ou mais paredes no modelo\n"
            "e execute o script novamente.",
            "Selecao vazia",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
        return

    # --------------------------------------------------------------------------
    # 2. EXIBIR FORMULARIO DE ENTRADA
    # --------------------------------------------------------------------------
    formulario = RenomearParedesForm(len(paredes))
    resultado  = formulario.ShowDialog()

    if resultado != DialogResult.OK or formulario.numero_inicial is None:
        # Usuario cancelou
        return

    numero_inicial = formulario.numero_inicial

    # --------------------------------------------------------------------------
    # 3. CONFIRMAR OPERACAO COM O USUARIO
    # --------------------------------------------------------------------------
    p_inicio = formatar_marca(numero_inicial)
    p_fim    = formatar_marca(numero_inicial + len(paredes) - 1)

    confirmacao = MessageBox.Show(
        "Confirma a renomeacao de {0} parede(s)?\n\n"
        "Sequencia: {1} ate {2}".format(len(paredes), p_inicio, p_fim),
        "Confirmar Renomeacao",
        MessageBoxButtons.OKCancel,
        MessageBoxIcon.Question
    )

    if confirmacao != DialogResult.OK:
        return

    # --------------------------------------------------------------------------
    # 4. EXECUTAR RENOMEACAO
    # --------------------------------------------------------------------------
    try:
        renomeadas, erros, sem_param = renomear_paredes(paredes, numero_inicial)
        exibir_resultado(renomeadas, erros, sem_param, len(paredes))

    except Exception as ex:
        MessageBox.Show(
            "Ocorreu um erro critico durante a execucao:\n\n{0}\n\n"
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
