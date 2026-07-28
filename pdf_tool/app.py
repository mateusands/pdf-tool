"""Janela principal: sidebar de ferramentas, cabeçalho de contexto e status bar."""

import tkinter as tk
from typing import NamedTuple

import customtkinter as ctk

from . import theme as T
from .core import icons
from .tabs.tab_compress import TabCompress
from .tabs.tab_convert import TabConvert
from .tabs.tab_image_to_pdf import TabImageToPdf
from .tabs.tab_merge import TabMerge
from .tabs.tab_organize import TabOrganize
from .tabs.tab_pdf_to_image import TabPdfToImage
from .tabs.tab_protect import TabProtect
from .tabs.tab_rotate import TabRotate
from .tabs.tab_split import TabSplit
from .tabs.tab_unlock import TabUnlock
from .widgets import icone


class Ferramenta(NamedTuple):
    """Uma entrada da sidebar e a aba que ela abre."""
    chave: str
    rotulo: str        # nome curto, na sidebar
    icone: str         # nome no catálogo de `core.icons`
    titulo: str        # título completo, no cabeçalho de contexto
    descricao: str     # subtítulo explicativo, no cabeçalho de contexto
    classe: type


# ── As 10 ferramentas, agrupadas por finalidade ───────────────────────────────
# A ordem aqui é a ordem da sidebar. Ferramenta nova entra nesta lista e em
# lugar nenhum mais: a navegação, as páginas e o cabeçalho saem daqui.

GRUPOS = [
    ("Organizar", [
        Ferramenta("split", "Dividir", "scissors", "Dividir PDF",
                   "Escolha as páginas que deseja extrair para um novo arquivo.", TabSplit),
        Ferramenta("merge", "Juntar", "combine", "Juntar PDFs",
                   "Combine vários PDFs num só, na ordem que você definir.", TabMerge),
        Ferramenta("organize", "Organizar páginas", "layout-grid", "Organizar páginas",
                   "Reordene as páginas arrastando as miniaturas.", TabOrganize),
        Ferramenta("rotate", "Girar", "rotate-cw", "Girar páginas",
                   "Gire todas as páginas ou apenas as que você selecionar.", TabRotate),
    ]),
    ("Converter", [
        Ferramenta("convert", "PDF e Word", "arrow-right-left", "Converter PDF e Word",
                   "Converta PDF para Word editável, ou Word para PDF.", TabConvert),
        Ferramenta("pdf2img", "PDF para imagem", "image-down", "PDF para imagem",
                   "Exporte cada página como PNG ou JPG.", TabPdfToImage),
        Ferramenta("img2pdf", "Imagem para PDF", "file-image", "Imagem para PDF",
                   "Junte várias imagens num único PDF.", TabImageToPdf),
    ]),
    ("Otimizar e proteger", [
        Ferramenta("compress", "Compactar", "shrink", "Compactar PDF",
                   "Reduza o tamanho do arquivo escolhendo o nível de compactação.", TabCompress),
        Ferramenta("protect", "Proteger", "lock", "Proteger com senha",
                   "Gere uma cópia cifrada com AES-256, exigindo senha para abrir.", TabProtect),
        Ferramenta("unlock", "Desbloquear", "lock-open", "Remover senha",
                   "Gere uma cópia sem proteção, informando a senha atual.", TabUnlock),
    ]),
]

FERRAMENTAS = {f.chave: f for _, grupo in GRUPOS for f in grupo}

#: Estados da status bar: ícone + cor.
_ESTADOS = {
    "info":   ("info", T.MUTED),
    "ok":     ("circle-check", T.SUCCESS_TEXT),
    "erro":   ("circle-alert", T.DANGER_TEXT),
    "ocupado": ("loader-circle", T.ACCENT_TEXT),
}


class PDFToolApp:
    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Gerenciador de PDF e Word")
        self.root.geometry("1060x740")
        self.root.minsize(940, 620)
        self.root.configure(fg_color=T.BG)
        self._definir_icone_da_janela()

        # Antes de qualquer widget: as abas leem T.FONT_* ao se desenharem.
        T.configurar_fontes(self.root)

        self._ativa = "split"
        self._itens: dict = {}

        # sidebar(col 0) ocupa as duas primeiras linhas; a status bar cruza tudo.
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self._montar_sidebar()
        self._montar_cabecalho()
        self._montar_conteudo()
        self._montar_statusbar()

        self._selecionar("split")

    def _definir_icone_da_janela(self):
        """Ícone da barra de título, da barra de tarefas e do alt-tab.

        Vários tamanhos porque cada contexto do sistema pede um: 16 na barra de
        título, 48 no alt-tab, 128 no lançador. Sem isto o app aparece com o
        losango genérico do Tk.
        """
        self._icones_da_janela = [
            tk.PhotoImage(data=icons.renderizar_icone_do_app(t, T.ACCENT, T.ON_ACCENT))
            for t in (16, 32, 48, 64, 128)
        ]
        try:
            self.root.iconphoto(True, *self._icones_da_janela)
        except tk.TclError:
            pass   # gerenciador de janelas sem suporte: só perde o ícone

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def _montar_sidebar(self):
        barra = ctk.CTkFrame(self.root, fg_color=T.SIDEBAR_BG, corner_radius=0,
                             width=T.SIDEBAR_W)
        barra.grid(row=0, column=0, rowspan=2, sticky="nsw")
        barra.grid_propagate(False)

        # Marca
        marca = ctk.CTkFrame(barra, fg_color="transparent", height=T.TOPBAR_H)
        marca.pack(fill="x", padx=T.PAD_L, pady=(T.PAD_L, T.PAD_M))
        tk.Label(marca, image=icone("file-text", 22, T.ACCENT),
                 bg=T.SIDEBAR_BG, bd=0).pack(side="left")
        textos = ctk.CTkFrame(marca, fg_color="transparent")
        textos.pack(side="left", padx=(T.PAD_S, 0))
        ctk.CTkLabel(textos, text="PDF e Word", font=T.FONT_DISPLAY,
                     text_color=T.FG, anchor="w").pack(anchor="w")
        ctk.CTkLabel(textos, text="Gerenciador", font=T.FONT_TINY,
                     text_color=T.MUTED, anchor="w").pack(anchor="w")

        ctk.CTkFrame(barra, fg_color=T.BORDER, height=1).pack(
            fill="x", padx=T.PAD_M, pady=(0, T.PAD_S))

        for titulo_do_grupo, ferramentas in GRUPOS:
            ctk.CTkLabel(
                barra, text=titulo_do_grupo.upper(), font=T.FONT_GROUP,
                text_color=T.NAV_GROUP_FG, anchor="w",
            ).pack(fill="x", padx=T.PAD_L, pady=(T.PAD_M, T.PAD_XS))

            for ferramenta in ferramentas:
                self._itens[ferramenta.chave] = self._montar_item(barra, ferramenta)

    def _montar_item(self, parent, ferramenta: Ferramenta):
        """Item da sidebar: filete de seleção + botão com ícone e rótulo."""
        linha = tk.Frame(parent, bg=T.SIDEBAR_BG)
        linha.pack(fill="x", padx=(T.PAD_S, T.PAD_M), pady=1)

        filete = tk.Frame(linha, bg=T.SIDEBAR_BG, width=3, height=T.NAV_ITEM_H)
        filete.pack(side="left", fill="y")
        filete.pack_propagate(False)

        botao = ctk.CTkButton(
            linha, text=f"  {ferramenta.rotulo}",
            image=icone(ferramenta.icone, T.ICON_MD, T.NAV_ITEM_FG),
            compound="left", anchor="w",
            fg_color="transparent", hover_color=T.NAV_ITEM_HOVER_BG,
            text_color=T.NAV_ITEM_FG, font=T.FONT_NAV,
            corner_radius=T.RADIUS_SM, height=T.NAV_ITEM_H, cursor="hand2",
            command=lambda c=ferramenta.chave: self._selecionar(c),
        )
        botao.pack(side="left", fill="x", expand=True, padx=(T.PAD_XS, 0))
        return botao, filete

    # ── Cabeçalho de contexto ─────────────────────────────────────────────────

    def _montar_cabecalho(self):
        cabecalho = ctk.CTkFrame(self.root, fg_color=T.BG, corner_radius=0,
                                 height=T.TOPBAR_H)
        cabecalho.grid(row=0, column=1, sticky="ew")
        cabecalho.grid_propagate(False)

        caixa = ctk.CTkFrame(cabecalho, fg_color="transparent")
        caixa.pack(side="left", padx=T.PAD_XL, pady=T.PAD_M)

        self._icone_do_titulo = tk.Label(caixa, bg=T.BG, bd=0)
        self._icone_do_titulo.pack(side="left", padx=(0, T.PAD_M))

        textos = ctk.CTkFrame(caixa, fg_color="transparent")
        textos.pack(side="left")
        self._titulo = ctk.CTkLabel(textos, text="", font=T.FONT_TITLE,
                                    text_color=T.FG, anchor="w")
        self._titulo.pack(anchor="w")
        self._descricao = ctk.CTkLabel(textos, text="", font=T.FONT_SUBTITLE,
                                       text_color=T.MUTED, anchor="w")
        self._descricao.pack(anchor="w")

    # ── Conteúdo ──────────────────────────────────────────────────────────────

    def _montar_conteudo(self):
        area = ctk.CTkFrame(self.root, fg_color=T.BG, corner_radius=0)
        area.grid(row=1, column=1, sticky="nsew")

        self._paginas: dict = {}
        for chave, ferramenta in FERRAMENTAS.items():
            pagina = ctk.CTkFrame(area, fg_color=T.BG)
            pagina.place(relx=0, rely=0, relwidth=1, relheight=1)

            card = ctk.CTkFrame(pagina, fg_color=T.CARD_BG, corner_radius=T.RADIUS,
                                border_width=1, border_color=T.BORDER)
            card.pack(fill="both", expand=True, padx=(0, T.PAD_XL),
                      pady=(0, T.PAD_L))

            ferramenta.classe(card, set_status=self._set_status, root=self.root)
            self._paginas[chave] = pagina

    # ── Status bar ────────────────────────────────────────────────────────────

    def _montar_statusbar(self):
        barra = ctk.CTkFrame(self.root, fg_color=T.SURFACE_1, corner_radius=0,
                             height=T.STATUSBAR_H)
        barra.grid(row=2, column=0, columnspan=2, sticky="ew")
        barra.grid_propagate(False)

        self._icone_de_status = tk.Label(barra, bg=T.SURFACE_1, bd=0)
        self._icone_de_status.pack(side="left", padx=(T.PAD_L, T.PAD_S))

        self._texto_de_status = ctk.CTkLabel(barra, text="Pronto", font=T.FONT_SMALL,
                                             text_color=T.MUTED)
        self._texto_de_status.pack(side="left")
        self._set_status("Pronto")

    # ── Troca de ferramenta ───────────────────────────────────────────────────

    def _selecionar(self, chave: str):
        self._ativa = chave
        ferramenta = FERRAMENTAS[chave]

        for outra, (botao, filete) in self._itens.items():
            ativo = outra == chave
            cor = T.NAV_ITEM_ACTIVE_FG if ativo else T.NAV_ITEM_FG
            botao.configure(
                fg_color=T.NAV_ITEM_ACTIVE_BG if ativo else "transparent",
                text_color=cor,
                font=T.FONT_NAV_ACTIVE if ativo else T.FONT_NAV,
                image=icone(FERRAMENTAS[outra].icone, T.ICON_MD,
                            T.ACCENT_TEXT if ativo else T.NAV_ITEM_FG),
            )
            filete.config(bg=T.NAV_INDICATOR if ativo else T.SIDEBAR_BG)

        self._icone_do_titulo.config(image=icone(ferramenta.icone, 22, T.ACCENT_TEXT))
        self._titulo.configure(text=ferramenta.titulo)
        self._descricao.configure(text=ferramenta.descricao)
        self._paginas[chave].tkraise()

    def _set_status(self, msg: str, estado: str = "info"):
        nome_do_icone, cor = _ESTADOS.get(estado, _ESTADOS["info"])
        self._icone_de_status.config(image=icone(nome_do_icone, T.ICON_SM, cor))
        self._texto_de_status.configure(text=msg, text_color=cor)
        self.root.update_idletasks()
