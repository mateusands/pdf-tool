"""Componentes reutilizáveis da interface.

Tudo que aparece em mais de uma aba mora aqui: ícone, título de seção, botão,
grupo de pills, campo de senha, área de arrastar arquivo e grade de miniaturas.
Aba não redesenha componente na mão — se falta um, ele nasce neste arquivo.
"""

import tkinter as tk
import warnings

import customtkinter as ctk

from . import theme as T
from .core import icons, pdf_io

# Os ícones são rasterizados no tamanho exato em pixels por `icone()`, então não
# há o que o CustomTkinter escalar — o aviso de HiDPI dele não se aplica aqui.
# (Usar CTkImage exigiria Pillow, uma dependência a mais só para isto.)
warnings.filterwarnings("ignore", message=".*Given image is not CTkImage.*")


# ── Ícones ────────────────────────────────────────────────────────────────────

_CACHE_DE_ICONES: dict = {}


def icone(nome: str, tamanho: int = T.ICON_MD, cor: str = T.FG_SECONDARY) -> tk.PhotoImage:
    """`tk.PhotoImage` do ícone, com cache.

    O cache não é só velocidade: o Tk descarta a imagem assim que ninguém mais
    a referencia, e o widget sozinho não segura essa referência — sem o cache o
    ícone simplesmente some da tela.
    """
    chave = (nome, tamanho, cor)
    if chave not in _CACHE_DE_ICONES:
        _CACHE_DE_ICONES[chave] = tk.PhotoImage(
            data=icons.renderizar_png(nome, tamanho, cor)
        )
    return _CACHE_DE_ICONES[chave]


# ── Rolagem com a roda do mouse ───────────────────────────────────────────────

_EVENTOS_DE_RODA = ("<MouseWheel>", "<Button-4>", "<Button-5>")


def passos_de_rolagem(delta: int, num=None) -> int:
    """Unidades de `yview_scroll` para um evento de roda. Ver `tests/test_scroll.py`."""
    if num == 4:
        return -1
    if num == 5:
        return 1
    if not delta:
        return 0
    if abs(delta) >= 120:          # Windows: múltiplos de 120, acumulam
        return -int(delta / 120)
    return -1 if delta > 0 else 1  # macOS: delta pequeno


def _ponteiro_sobre(widget) -> bool:
    """True se o ponteiro está sobre o widget ou sobre um filho dele."""
    try:
        alvo = widget.winfo_containing(*widget.winfo_pointerxy())
    except Exception:
        return False
    while alvo is not None:
        if alvo is widget:
            return True
        alvo = getattr(alvo, "master", None)
    return False


def ligar_rolagem(canvas) -> None:
    """Faz a roda do mouse rolar `canvas` em qualquer plataforma.

    Usa `bind_all` enquanto o ponteiro está sobre a área: os cartões e rótulos
    desenhados dentro do canvas são widgets próprios, e evento de mouse no Tk não
    sobe para o pai — sem isso a roda só funcionaria sobre o fundo vazio.
    """
    def _rolar(evento):
        passos = passos_de_rolagem(getattr(evento, "delta", 0),
                                   getattr(evento, "num", None))
        if passos:
            canvas.yview_scroll(passos, "units")
        return "break"

    def _ativar(_evento=None):
        for nome in _EVENTOS_DE_RODA:
            canvas.bind_all(nome, _rolar)

    def _desativar(_evento=None):
        # <Leave> também dispara ao entrar num filho — aí a rolagem continua valendo.
        if _ponteiro_sobre(canvas):
            return
        for nome in _EVENTOS_DE_RODA:
            canvas.unbind_all(nome)

    canvas.bind("<Enter>", _ativar, add="+")
    canvas.bind("<Leave>", _desativar, add="+")


# ── Área rolável ──────────────────────────────────────────────────────────────

def criar_area_rolavel(parent, **pack_kw):
    """Painel com borda + canvas rolável. Devolve `(canvas, conteudo)`.

    O frame interno acompanha a largura do canvas — sem isso ele encolhe até o
    tamanho do conteúdo e qualquer coisa centralizada dentro dele (o estado
    vazio, por exemplo) aparece grudada na esquerda.
    """
    painel = ctk.CTkFrame(parent, fg_color=T.GRID_BG, corner_radius=T.RADIUS_SM,
                          border_width=1, border_color=T.BORDER)
    painel.pack(**pack_kw)

    canvas = tk.Canvas(painel, bg=T.GRID_BG, highlightthickness=0)
    barra = ctk.CTkScrollbar(painel, command=canvas.yview,
                             button_color=T.BORDER_LIGHT, button_hover_color=T.MUTED)
    canvas.configure(yscrollcommand=barra.set)
    barra.pack(side="right", fill="y", padx=(0, 3), pady=3)
    canvas.pack(side="left", fill="both", expand=True, padx=3, pady=3)

    conteudo = tk.Frame(canvas, bg=T.GRID_BG)
    janela = canvas.create_window((0, 0), window=conteudo, anchor="nw")
    conteudo.bind("<Configure>",
                  lambda _: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",
                lambda e: canvas.itemconfigure(janela, width=e.width), add="+")
    ligar_rolagem(canvas)
    return canvas, conteudo


def estado_vazio(parent, nome_do_icone: str, texto: str):
    """Aviso centralizado de 'nada aqui ainda', com ícone apagado."""
    quadro = tk.Frame(parent, bg=T.GRID_BG, pady=52)
    quadro.pack(expand=True)
    tk.Label(quadro, image=icone(nome_do_icone, T.ICON_XL, T.BORDER_LIGHT),
             bg=T.GRID_BG, bd=0).pack()
    tk.Label(quadro, text=texto, bg=T.GRID_BG, fg=T.MUTED,
             font=T.FONT_BODY, pady=T.PAD_M).pack()
    return quadro


# ── Cabeçalho de seção ────────────────────────────────────────────────────────

def titulo_secao(parent, titulo: str, subtitulo: str = "", nome_do_icone: str = None):
    """Cabeçalho padrão de toda aba: ícone + título, subtítulo e um filete."""
    linha = ctk.CTkFrame(parent, fg_color="transparent")
    linha.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XL, 0))

    if nome_do_icone:
        tk.Label(
            linha, image=icone(nome_do_icone, T.ICON_LG, T.ACCENT_TEXT),
            bg=T.CARD_BG, bd=0,
        ).pack(side="left", padx=(0, T.PAD_M))

    ctk.CTkLabel(
        linha, text=titulo, font=T.FONT_TITLE, text_color=T.FG, anchor="w",
    ).pack(side="left")

    if subtitulo:
        ctk.CTkLabel(
            parent, text=subtitulo, font=T.FONT_SUBTITLE, text_color=T.MUTED, anchor="w",
        ).pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XS, 0))

    ctk.CTkFrame(parent, fg_color=T.BORDER, height=1).pack(
        fill="x", padx=T.PAD_XL, pady=(T.PAD_L, T.PAD_L),
    )


# Nome antigo, mantido para não quebrar quem ainda chamar assim.
def section_title(parent, title, subtitle=""):
    titulo_secao(parent, title, subtitle)


# ── Botões ────────────────────────────────────────────────────────────────────

_VARIANTES = {
    "primario":   (T.ACCENT, T.ACCENT_HOVER, T.ON_ACCENT, None),
    "sucesso":    (T.SUCCESS, T.SUCCESS_HOVER, T.ON_ACCENT, None),
    "perigo":     (T.DANGER, T.DANGER_HOVER, T.ON_ACCENT, None),
    "secundario": ("transparent", T.SURFACE_4, T.FG_SECONDARY, T.BORDER_LIGHT),
    "fantasma":   ("transparent", T.SURFACE_3, T.MUTED, None),
}


class Botao(ctk.CTkButton):
    """Botão do app. A variante decide a cor; o ícone é opcional e vai à esquerda.

    Desabilitar repinta **fundo, texto e ícone juntos**. O CustomTkinter sozinho
    só escurece o texto (`text_color_disabled`, que é `gray60`) e deixa o
    preenchimento e a imagem intactos: em cima do azul isso dá um texto cinza
    ilegível num botão que continua com cara de ativo.
    """

    def __init__(self, parent, texto: str, command=None, variante: str = "primario",
                 nome_do_icone: str = None, altura: int = T.BUTTON_H,
                 largura: int = 0, state: str = "normal"):
        self._variante = variante
        self._nome_do_icone = nome_do_icone
        fundo, hover, cor_do_texto, borda = self._cores(state)

        extras = {}
        if nome_do_icone:
            extras["image"] = icone(nome_do_icone, T.ICON_MD, cor_do_texto)
            extras["compound"] = "left"

        super().__init__(
            parent, text=f"  {texto}" if nome_do_icone else texto,
            command=command, state=state,
            fg_color=fundo, hover_color=hover, text_color=cor_do_texto,
            text_color_disabled=T.DISABLED_FG,
            # `border_color` não aceita "transparent" — quando não há borda, a
            # largura zero é o que a esconde.
            border_width=1 if borda else 0, border_color=borda or T.BORDER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=T.RADIUS_SM,
            height=altura, width=largura, **extras,
        )

    def _cores(self, state: str):
        fundo, hover, cor_do_texto, borda = _VARIANTES[self._variante]
        if state != "disabled":
            return fundo, hover, cor_do_texto, borda
        if fundo == "transparent":          # secundário/fantasma: só apaga o texto
            return fundo, hover, T.DISABLED_FG, borda
        return T.DISABLED_BG, T.DISABLED_BG, T.DISABLED_FG, T.BORDER

    def configure(self, **kwargs):
        if "state" in kwargs:
            fundo, hover, cor_do_texto, borda = self._cores(kwargs["state"])
            kwargs.setdefault("fg_color", fundo)
            kwargs.setdefault("hover_color", hover)
            kwargs.setdefault("text_color", cor_do_texto)
            kwargs.setdefault("border_color", borda or T.BORDER)
            if self._nome_do_icone:
                kwargs.setdefault("image", icone(self._nome_do_icone, T.ICON_MD,
                                                 cor_do_texto))
        super().configure(**kwargs)


def botao(parent, texto: str, command=None, variante: str = "primario",
          nome_do_icone: str = None, altura: int = T.BUTTON_H, largura: int = 0,
          state: str = "normal"):
    return Botao(parent, texto, command=command, variante=variante,
                 nome_do_icone=nome_do_icone, altura=altura, largura=largura,
                 state=state)


# ── Grupo de pills (escolha única) ────────────────────────────────────────────

class GrupoPills:
    """Linha de botões-pílula onde só um fica ativo.

    Substitui a mesma montagem repetida em Compactar (nível), PDF → Imagem
    (formato e DPI) e Girar (ângulo).
    """

    def __init__(self, parent, opcoes, valor_inicial=None, ao_mudar=None,
                 largura: int = 104):
        self._ao_mudar = ao_mudar
        self._valor = valor_inicial if valor_inicial is not None else opcoes[0][0]
        self._botoes: dict = {}

        linha = ctk.CTkFrame(parent, fg_color="transparent")
        linha.pack(anchor="w")

        for opcao in opcoes:
            valor, rotulo = opcao[0], opcao[1]
            nome_do_icone = opcao[2] if len(opcao) > 2 else None
            extras = {}
            if nome_do_icone:
                extras["image"] = icone(nome_do_icone, T.ICON_SM, T.FG_SECONDARY)
                extras["compound"] = "left"
                rotulo = f"  {rotulo}"

            btn = ctk.CTkButton(
                linha, text=rotulo, width=largura, height=34,
                corner_radius=T.RADIUS_SM, cursor="hand2", font=T.FONT_PILL,
                border_width=1, border_color=T.BORDER,
                command=lambda v=valor: self.definir(v), **extras,
            )
            btn.pack(side="left", padx=(0, T.PAD_S))
            self._botoes[valor] = (btn, nome_do_icone)

        self._pintar()

    def _pintar(self):
        for valor, (btn, nome_do_icone) in self._botoes.items():
            ativo = valor == self._valor
            cor_do_texto = T.ON_ACCENT if ativo else T.FG_SECONDARY
            btn.configure(
                fg_color=T.ACCENT if ativo else T.SURFACE_3,
                hover_color=T.ACCENT_HOVER if ativo else T.SURFACE_4,
                text_color=cor_do_texto,
                border_color=T.ACCENT if ativo else T.BORDER,
            )
            if nome_do_icone:
                btn.configure(image=icone(nome_do_icone, T.ICON_SM, cor_do_texto))

    def definir(self, valor):
        self._valor = valor
        self._pintar()
        if self._ao_mudar:
            self._ao_mudar(valor)

    def valor(self):
        return self._valor


# ── Campo de senha ────────────────────────────────────────────────────────────

class CampoSenha:
    """Campo de senha com botão de mostrar/ocultar."""

    def __init__(self, parent, rotulo: str):
        ctk.CTkLabel(parent, text=rotulo, font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(anchor="w", pady=(0, T.PAD_XS))

        linha = ctk.CTkFrame(parent, fg_color="transparent")
        linha.pack(anchor="w", pady=(0, T.PAD_M))

        self._visivel = False
        self._entrada = ctk.CTkEntry(
            linha, show="•", width=300, height=38,
            fg_color=T.FIELD_BG, border_color=T.BORDER, text_color=T.FG,
            font=T.FONT_BODY, corner_radius=T.RADIUS_SM,
        )
        self._entrada.pack(side="left")

        self._olho = ctk.CTkButton(
            linha, text="", width=38, height=38,
            image=icone("eye", T.ICON_MD, T.MUTED),
            fg_color=T.FIELD_BG, hover_color=T.SURFACE_4,
            border_width=1, border_color=T.BORDER, corner_radius=T.RADIUS_SM,
            cursor="hand2", command=self._alternar,
        )
        self._olho.pack(side="left", padx=(T.PAD_S, 0))

    def _alternar(self):
        self._visivel = not self._visivel
        self._entrada.configure(show="" if self._visivel else "•")
        self._olho.configure(
            image=icone("eye-off" if self._visivel else "eye", T.ICON_MD, T.MUTED)
        )

    def valor(self) -> str:
        return self._entrada.get()

    def limpar(self):
        self._entrada.delete(0, tk.END)


# ── Área de seleção de arquivo ────────────────────────────────────────────────

class DropZone(tk.Canvas):
    """Área clicável de seleção de arquivo, com borda tracejada e botão de limpar."""

    def __init__(self, parent, icon="upload", text="Selecione um arquivo",
                 subtitle="ou clique para escolher", command=None,
                 on_clear=None, height=132):
        super().__init__(parent, bg=T.CARD_BG, highlightthickness=0,
                         cursor="hand2", height=height)
        self._cmd = command
        self._on_clear = on_clear
        self._hover = False
        self._hover_x = False
        self._has_file = False
        self._nome_do_arquivo = ""
        self._detalhe = ""
        self._icone = icon
        self._text = text
        self._subtitle = subtitle

        self.bind("<Configure>", self._draw)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", lambda _: self._set_hover(True))
        self.bind("<Leave>", lambda _: self._set_hover(False))
        self.bind("<Motion>", self._on_motion)
        self.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XL, T.PAD_L))

    # -- estado --

    def _set_hover(self, val):
        self._hover = val
        if not val:
            self._hover_x = False
        self._draw()

    def set_file(self, name, size_str=""):
        self._has_file = True
        self._nome_do_arquivo = name
        self._detalhe = size_str
        self._draw()

    def clear_file(self):
        self._has_file = False
        self._nome_do_arquivo = ""
        self._detalhe = ""
        self._draw()

    # -- botão de limpar --

    def _x_bbox(self):
        w = self.winfo_width()
        x2 = w - 18
        x1 = x2 - 26
        y1 = 16
        return x1, y1, x2, y1 + 26

    def _in_x(self, mx, my):
        if not self._has_file:
            return False
        x1, y1, x2, y2 = self._x_bbox()
        return x1 <= mx <= x2 and y1 <= my <= y2

    def _on_motion(self, event):
        antes = self._hover_x
        self._hover_x = self._in_x(event.x, event.y)
        if antes != self._hover_x:
            self._draw()

    def _on_click(self, event):
        if self._has_file and self._in_x(event.x, event.y):
            self.clear_file()
            if self._on_clear:
                self._on_clear()
        elif self._cmd:
            self._cmd()

    # -- desenho --

    def _draw(self, _event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1:
            return

        pad = 2
        cor_da_borda = T.DROPZONE_HOVER_BD if self._hover else T.DROPZONE_BORDER
        cor_do_fundo = T.DROPZONE_HOVER_BG if self._hover else T.DROPZONE_BG

        self.create_rectangle(pad, pad, w - pad, h - pad, fill=cor_do_fundo, outline="")
        self.create_rectangle(pad, pad, w - pad, h - pad,
                              outline=cor_da_borda, dash=(6, 4), width=1)

        cx, cy = w / 2, h / 2
        if self._has_file:
            self.create_image(cx - 108, cy, image=icone("file-text", T.ICON_LG, T.ACCENT_TEXT))
            self.create_text(cx - 88, cy - 8, text=self._nome_do_arquivo,
                             fill=T.FG, font=T.FONT_BODY, anchor="w")
            legenda = self._detalhe or "Clique para trocar de arquivo"
            self.create_text(cx - 88, cy + 12, text=legenda,
                             fill=T.MUTED, font=T.FONT_SMALL, anchor="w")

            x1, y1, x2, y2 = self._x_bbox()
            self.create_oval(x1, y1, x2, y2,
                             fill=T.DANGER if self._hover_x else T.SURFACE_4, outline="")
            self.create_image((x1 + x2) / 2, (y1 + y2) / 2,
                              image=icone("x", T.ICON_SM,
                                          T.ON_ACCENT if self._hover_x else T.FG_SECONDARY))
        else:
            self.create_image(cx, cy - 22,
                              image=icone(self._icone, T.ICON_XL,
                                          T.ACCENT if self._hover else T.MUTED))
            self.create_text(cx, cy + 14, text=self._text, fill=T.FG, font=T.FONT_BODY)
            self.create_text(cx, cy + 36, text=self._subtitle,
                             fill=T.MUTED, font=T.FONT_SMALL)


# ── Grade de miniaturas ───────────────────────────────────────────────────────

class ThumbnailGrid:
    """Grade rolável de miniaturas do PDF, com seleção por clique."""

    def __init__(self, parent, cols=T.GRID_COLS, thumb_w=T.THUMB_W, thumb_h=T.THUMB_H):
        self.cols = cols
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h
        self._images: list = []
        self._cells: dict = {}
        self._selected: set = set()
        self._on_change_cb = None

        self._canvas, self._inner = criar_area_rolavel(
            parent, fill="both", expand=True, padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._mostrar_vazio()

    def _mostrar_vazio(self, texto="Nenhum PDF carregado"):
        estado_vazio(self._inner, "file-text", texto)

    def on_change(self, callback):
        self._on_change_cb = callback

    def clear(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()
        self._cells.clear()
        self._selected.clear()
        self._mostrar_vazio()
        self._canvas.configure(scrollregion=(0, 0, 0, 0))
        if self._on_change_cb:
            self._on_change_cb(set())

    def load_pdf(self, path: str) -> int:
        """Renderiza as páginas do PDF na grade. Levanta `pdf_io.PdfError`."""
        # Renderiza ANTES de limpar a tela: se o PDF estiver protegido, a grade
        # atual continua no lugar em vez de virar uma área vazia sem explicação.
        miniaturas = pdf_io.miniaturas_ppm(path, self.thumb_w, self.thumb_h)

        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()
        self._cells.clear()
        self._selected.clear()

        for i, ppm in enumerate(miniaturas):
            photo = tk.PhotoImage(data=ppm)
            self._images.append(photo)

            row, col = divmod(i, self.cols)
            cell = tk.Frame(self._inner, bg=T.GRID_BG, padx=6, pady=6)
            border = tk.Frame(cell, bg=T.BORDER, padx=2, pady=2)
            img_lbl = tk.Label(border, image=photo, cursor="hand2", bg=T.BORDER)
            img_lbl.pack()
            border.pack()
            num_lbl = tk.Label(
                cell, text=str(i + 1), bg=T.GRID_BG, fg=T.MUTED, font=T.FONT_TINY,
            )
            num_lbl.pack(pady=(4, 0))
            cell.grid(row=row, column=col, padx=4, pady=4)
            self._cells[i] = (cell, border, num_lbl)

            for w in (cell, border, img_lbl, num_lbl):
                w.bind("<Button-1>", lambda _, idx=i: self.toggle(idx))

            if i % 3 == 2:
                self._canvas.update_idletasks()

        self._canvas.yview_moveto(0)
        return len(self._cells)

    def toggle(self, idx: int):
        if idx in self._selected:
            self._selected.discard(idx)
        else:
            self._selected.add(idx)
        self._refresh()

    def select_all(self):
        self._selected = set(self._cells.keys())
        self._refresh()

    def deselect_all(self):
        self._selected.clear()
        self._refresh()

    def get_selected(self) -> set:
        return set(self._selected)

    def _refresh(self):
        for idx, (_, border, num_lbl) in self._cells.items():
            selecionada = idx in self._selected
            border.config(bg=T.ACCENT if selecionada else T.BORDER)
            num_lbl.config(
                fg=T.ACCENT_TEXT if selecionada else T.MUTED,
                font=T.FONT_TINY,
            )
        if self._on_change_cb:
            self._on_change_cb(self._selected)
