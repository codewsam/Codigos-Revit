# -*- coding: utf-8 -*-
__title__ = "Colorir Famílias\npor Tipo"
__version__ = "2.0"
__doc__ = "Colore os elementos da vista ativa por tipo de família, permitindo ao usuário escolher a cor de cada tipo."

from Autodesk.Revit.DB import *
from pyrevit import forms

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ─────────────────────────────────────────────
# ETAPA 1 — Mapeamento de categorias
# ─────────────────────────────────────────────

CATEGORIAS = {
    "Paredes"      : BuiltInCategory.OST_Walls,
    "Portas"       : BuiltInCategory.OST_Doors,
    "Janelas"      : BuiltInCategory.OST_Windows,
    "Pisos"        : BuiltInCategory.OST_Floors,
    "Pilares"      : BuiltInCategory.OST_StructuralColumns,
    "Vigas"        : BuiltInCategory.OST_StructuralFraming,
    "Tetos"        : BuiltInCategory.OST_Ceilings,
    "Móveis"       : BuiltInCategory.OST_Furniture,
    "Equipamentos" : BuiltInCategory.OST_MechanicalEquipment,
    "Luminárias"   : BuiltInCategory.OST_LightingFixtures,
    "Dutos"        : BuiltInCategory.OST_DuctCurves,
    "Tubulações"   : BuiltInCategory.OST_PipeCurves,
    "Eletrodutos"  : BuiltInCategory.OST_Conduit,
}

# ─────────────────────────────────────────────
# ETAPA 2 — Selecionar categoria
# ─────────────────────────────────────────────

categoria_nome = forms.SelectFromList.show(
    sorted(CATEGORIAS.keys()),
    title="Colorir Famílias por Tipo",
    prompt="Selecione a categoria a colorir:",
    multiselect=False
)

if not categoria_nome:
    forms.alert("Nenhuma categoria selecionada.", exitscript=True)

bic = CATEGORIAS[categoria_nome]

# ─────────────────────────────────────────────
# ETAPA 3 — Buscar elementos na vista ativa
# ─────────────────────────────────────────────

elementos = (
    FilteredElementCollector(doc, doc.ActiveView.Id)
    .OfCategory(bic)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not elementos:
    forms.alert("Nenhum elemento encontrado para '{}'.".format(categoria_nome), exitscript=True)

# ─────────────────────────────────────────────
# ETAPA 4 — Descobrir tipos únicos
# ─────────────────────────────────────────────

tipos_unicos = {}
for el in elementos:
    tipo_id = el.GetTypeId()
    if tipo_id not in tipos_unicos:
        tipo = doc.GetElement(tipo_id)
        nome_tipo = tipo.Name if tipo and hasattr(tipo, "Name") else "Tipo {}".format(tipo_id.IntegerValue)
        tipos_unicos[tipo_id] = nome_tipo

# ─────────────────────────────────────────────
# ETAPA 5 — Paleta padrão (fallback)
# ─────────────────────────────────────────────

PALETA_PADRAO = [
    (52,  152, 219),
    (231, 76,  60 ),
    (46,  204, 113),
    (230, 126, 34 ),
    (155, 89,  182),
    (26,  188, 156),
    (241, 196, 15 ),
    (236, 240, 241),
    (52,  73,  94 ),
    (211, 84,  0  ),
    (39,  174, 96 ),
    (142, 68,  173),
]

# Nomes amigáveis das cores para o usuário escolher
CORES_NOMES = {
    "Azul"            : (52,  152, 219),
    "Vermelho"        : (231, 76,  60 ),
    "Verde"           : (46,  204, 113),
    "Laranja"         : (230, 126, 34 ),
    "Roxo"            : (155, 89,  182),
    "Turquesa"        : (26,  188, 156),
    "Amarelo"         : (241, 196, 15 ),
    "Branco"          : (236, 240, 241),
    "Azul Escuro"     : (52,  73,  94 ),
    "Laranja Escuro"  : (211, 84,  0  ),
    "Verde Escuro"    : (39,  174, 96 ),
    "Roxo Escuro"     : (142, 68,  173),
    "Rosa"            : (255, 105, 180),
    "Ciano"           : (0,   188, 212),
    "Lima"            : (205, 220, 57 ),
    "Marrom"          : (121, 85,  72 ),
    "Cinza"           : (158, 158, 158),
    "Índigo"          : (63,  81,  181),
    "Teal"            : (0,   150, 136),
    "Âmbar"           : (255, 193, 7  ),
}

lista_cores = sorted(CORES_NOMES.keys())

# ─────────────────────────────────────────────
# ETAPA 6 — Usuário escolhe cor para cada tipo
# ─────────────────────────────────────────────

cor_por_tipo = {}

for i, (tipo_id, nome_tipo) in enumerate(tipos_unicos.items()):
    # Sugerir uma cor padrão da paleta
    r_pad, g_pad, b_pad = PALETA_PADRAO[i % len(PALETA_PADRAO)]
    # Encontrar o nome mais próximo da paleta padrão para pré-selecionar
    cor_sugerida = None
    for nome_c, rgb in CORES_NOMES.items():
        if rgb == (r_pad, g_pad, b_pad):
            cor_sugerida = nome_c
            break

    escolha = forms.SelectFromList.show(
        lista_cores,
        title="Tipo: {}".format(nome_tipo),
        prompt="Escolha a cor para o tipo:\n\"{}\"".format(nome_tipo),
        multiselect=False
    )

    if not escolha:
        # Se o usuário cancelar, usa a cor padrão da paleta
        cor_por_tipo[tipo_id] = Color(r_pad, g_pad, b_pad)
    else:
        r, g, b = CORES_NOMES[escolha]
        cor_por_tipo[tipo_id] = Color(r, g, b)

# ─────────────────────────────────────────────
# ETAPA 7 — Buscar padrão sólido
# ─────────────────────────────────────────────

padrao_solido = None
for fp in FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements():
    if fp.GetFillPattern().IsSolidFill:
        padrao_solido = fp
        break

if not padrao_solido:
    forms.alert("Padrão de preenchimento sólido não encontrado.", exitscript=True)

# ─────────────────────────────────────────────
# ETAPA 8 — Aplicar cores
# ─────────────────────────────────────────────

t = Transaction(doc, "Colorir Famílias por Tipo (cores personalizadas)")
try:
    t.Start()

    for el in elementos:
        tipo_id = el.GetTypeId()
        cor = cor_por_tipo.get(tipo_id)
        if not cor:
            continue

        ogs = OverrideGraphicSettings()
        ogs.SetSurfaceForegroundPatternId(padrao_solido.Id)
        ogs.SetSurfaceForegroundPatternColor(cor)
        doc.ActiveView.SetElementOverrides(el.Id, ogs)

    t.Commit()

except Exception as e:
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    forms.alert("Erro ao aplicar cores:\n{}".format(str(e)), exitscript=True)

# ─────────────────────────────────────────────
# ETAPA 9 — Conclusão
# ─────────────────────────────────────────────

resumo = "\n".join(
    "• {} → {}".format(nome, next((n for n, rgb in CORES_NOMES.items() if Color(*rgb).Red == cor_por_tipo[tid].Red), "Personalizada"))
    for tid, nome in tipos_unicos.items()
)

forms.alert(
    "Cores aplicadas com sucesso!\n\nTipos coloridos: {}\n\n{}".format(len(tipos_unicos), resumo),
    warn_icon=False
)