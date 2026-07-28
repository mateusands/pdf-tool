"""Tokens de estilo — a ÚNICA fonte de cor, fonte e espaçamento do app.

Nenhum hex literal fora daqui. Faltou uma cor? Acrescente neste arquivo.

## Elevação tonal

No escuro não se usa sombra para dar profundidade: **quanto mais alto o
elemento, mais clara a superfície**. São cinco níveis, subindo ~6% de luminância
cada, e nenhum deles é preto puro (preto puro achata a percepção de profundidade
e faz a cor do conteúdo "vazar"):

    BG          janela, o fundo de tudo
    SURFACE_1   sidebar e barras de cromo
    SURFACE_2   o card de conteúdo
    SURFACE_3   elementos dentro do card (campo, dropzone, grade)
    SURFACE_4   estado hover / item selecionado

## Cor de acento

`ACCENT` é para preenchimento (botão primário, com texto branco por cima).
`ACCENT_TEXT` é a mesma cor clareada, para quando o azul é o TEXTO sobre fundo
escuro — o tom de preenchimento não tem contraste suficiente nesse papel.
"""

from .core import fonts

# ── Superfícies (elevação tonal) ──────────────────────────────────────────────
BG              = "#0B0F16"
SURFACE_1       = "#111722"
SURFACE_2       = "#161D2A"
SURFACE_3       = "#1D2634"
SURFACE_4       = "#252F42"

# Aliases de intenção
SURFACE         = SURFACE_2
SURFACE_HOVER   = SURFACE_4
SIDEBAR_BG      = SURFACE_1
CARD_BG         = SURFACE_2
FIELD_BG        = SURFACE_3
GRID_BG         = "#10151F"
GRID_CELL_BG    = SURFACE_3

# ── Bordas ────────────────────────────────────────────────────────────────────
BORDER          = "#222B3A"
BORDER_LIGHT    = "#2E3950"
BORDER_FOCUS    = "#3B82F6"

# ── Acento ────────────────────────────────────────────────────────────────────
# O tom de preenchimento é escolhido pelo contraste com o texto BRANCO por cima:
# não dá para deixar a letra mais clara que branco, então quem tem de escurecer
# é o fundo. `#3B82F6` dava só 3,68:1 (abaixo do mínimo AA de 4,5:1) e a leitura
# sofria; `#2563EB` dá 5,17:1. O hover clareia de volta para o tom anterior.
ACCENT          = "#2563EB"   # branco por cima: 5,17:1
ACCENT_HOVER    = "#3B82F6"
ACCENT_ACTIVE   = "#1D4ED8"
ACCENT_TEXT     = "#7EB0FF"   # azul como TEXTO sobre o card: 7,67:1
ACCENT_SOFT     = "#16243D"   # fundo tênue do item de navegação ativo

# Nomes históricos, mantidos para não espalhar renomeação pelas abas
PRIMARY         = ACCENT
PRIMARY_HOVER   = ACCENT_HOVER
PRIMARY_DARK    = ACCENT_ACTIVE

# ── Semânticas ────────────────────────────────────────────────────────────────
# Mesma regra do acento: SUCCESS e DANGER são PREENCHIMENTO com texto branco por
# cima, então precisam ser escuros o bastante. Os tons claros (`*_TEXT`) são para
# quando a cor é o próprio texto, sobre fundo escuro — papéis opostos.
SUCCESS         = "#1A7F37"   # branco por cima: 5,08:1
SUCCESS_HOVER   = "#238636"
SUCCESS_TEXT    = "#56D364"
DANGER          = "#C93C37"   # branco por cima: 5,02:1
DANGER_HOVER    = "#A82B26"
DANGER_TEXT     = "#FF7B72"
WARNING         = "#D29922"
WARNING_TEXT    = "#E3B341"

# ── Texto ─────────────────────────────────────────────────────────────────────
FG              = "#E6EDF3"
FG_SECONDARY    = "#A9B4C4"
MUTED           = "#6E7A8C"
ON_ACCENT       = "#FFFFFF"

# Desabilitado. O default do CustomTkinter para texto desabilitado é `gray60`
# (#999999) SOBRE a cor de preenchimento — em cima do azul isso fica ilegível, e
# o botão continua parecendo ativo. Aqui o estado apaga o preenchimento também:
# fica claro que está inativo E o texto continua legível.
DISABLED_BG     = SURFACE_3
DISABLED_FG     = "#78849A"

# ── Navegação ─────────────────────────────────────────────────────────────────
NAV_ITEM_FG         = FG_SECONDARY
NAV_ITEM_HOVER_BG   = SURFACE_3
NAV_ITEM_ACTIVE_BG  = ACCENT_SOFT
NAV_ITEM_ACTIVE_FG  = FG
NAV_GROUP_FG        = MUTED
NAV_INDICATOR       = ACCENT

# ── Drop zone ─────────────────────────────────────────────────────────────────
DROPZONE_BG         = SURFACE_3
DROPZONE_HOVER_BG   = "#182337"
DROPZONE_BORDER     = BORDER_LIGHT
DROPZONE_HOVER_BD   = ACCENT

# ── Arrastar e soltar (aba Organizar) ─────────────────────────────────────────
COLOR_DRAG_SRC  = WARNING
COLOR_DRAG_TGT  = SUCCESS

# ── Tipografia ────────────────────────────────────────────────────────────────
# `FONT_FAMILY` e as tuplas abaixo são recalculadas por `configurar_fontes(root)`
# assim que a janela existe — antes disso não dá para perguntar ao Tk quais
# famílias o sistema tem. Os valores aqui são só o ponto de partida.
FONT_FAMILY = fonts.ALTERNATIVA

_ESCALA = {
    "FONT_DISPLAY":    (14, "bold"),    # nome do app na sidebar
    "FONT_HEADER":     (15, "bold"),
    "FONT_TITLE":      (15, "bold"),    # título da ferramenta
    "FONT_SUBTITLE":   (11, "normal"),
    "FONT_BODY":       (12, "normal"),
    "FONT_LABEL":      (11, "normal"),
    "FONT_SMALL":      (10, "normal"),
    "FONT_TINY":       (9,  "normal"),
    "FONT_BUTTON":     (12, "bold"),
    "FONT_NAV":        (12, "normal"),  # item da sidebar
    "FONT_NAV_ACTIVE": (12, "bold"),
    "FONT_GROUP":      (9,  "bold"),    # rótulo de grupo da sidebar
    "FONT_PILL":       (11, "normal"),
}


def _aplicar_escala(familia: str) -> None:
    globals()["FONT_FAMILY"] = familia
    for nome, (tamanho, peso) in _ESCALA.items():
        globals()[nome] = (familia, tamanho) if peso == "normal" else (familia, tamanho, peso)


def configurar_fontes(root) -> str:
    """Escolhe a melhor família disponível no sistema e recalcula as tuplas.

    Precisa rodar logo depois de criar a janela e ANTES de montar qualquer
    widget — as abas leem `T.FONT_BODY` na hora em que se desenham.
    """
    import tkinter.font as tkfont

    familia = fonts.escolher(
        tkfont.families(root), fonts.preferidas_da_plataforma(), fonts.ALTERNATIVA
    )
    _aplicar_escala(familia)
    return familia


_aplicar_escala(FONT_FAMILY)

# ── Espaçamento ───────────────────────────────────────────────────────────────
PAD_XS  = 4
PAD_S   = 8
PAD_M   = 12
PAD_L   = 16
PAD_XL  = 24
PAD_XXL = 32

# ── Raios ─────────────────────────────────────────────────────────────────────
RADIUS_SM = 6
RADIUS    = 10
RADIUS_LG = 14

# ── Medidas ───────────────────────────────────────────────────────────────────
SIDEBAR_W   = 216
TOPBAR_H    = 66
STATUSBAR_H = 30
NAV_ITEM_H  = 34
ICON_SM     = 14
ICON_MD     = 17
ICON_LG     = 20
ICON_XL     = 30
BUTTON_H    = 42

# ── Grade de miniaturas ───────────────────────────────────────────────────────
THUMB_W   = 88
THUMB_H   = 124
GRID_COLS = 5
