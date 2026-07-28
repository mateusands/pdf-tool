import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pdf2docx import Converter

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..core.docx_convert import docx_to_pdf
from ..widgets import DropZone, botao, icone


class TabConvert:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._mode = None
        self._build(parent)

    def _build(self, p):
        # Os dois sentidos possíveis; o ativo acende conforme o arquivo escolhido
        cartoes = ctk.CTkFrame(p, fg_color="transparent")
        cartoes.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XL, T.PAD_M))
        cartoes.grid_columnconfigure(0, weight=1)
        cartoes.grid_columnconfigure(1, weight=1)

        self._card_pdf = self._montar_cartao(
            cartoes, "file-text", "PDF para Word",
            "Gera um .docx editável a partir do PDF")
        self._card_pdf.grid(row=0, column=0, padx=(0, T.PAD_S), sticky="nsew")

        self._card_word = self._montar_cartao(
            cartoes, "file-image", "Word para PDF",
            "Gera um PDF a partir do documento .docx")
        self._card_word.grid(row=0, column=1, padx=(T.PAD_S, 0), sticky="nsew")

        self._drop = DropZone(p, icon="upload",
                              text="Selecione um arquivo PDF ou DOCX",
                              subtitle="o sentido da conversão é detectado sozinho",
                              command=self._select, on_clear=self._clear)

        # Selo da ação detectada
        self._selo = ctk.CTkFrame(p, fg_color=T.SURFACE_3, corner_radius=T.RADIUS_SM)
        self._selo.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_L))
        self._icone_do_selo = tk.Label(self._selo, bg=T.SURFACE_3, bd=0)
        self._icone_do_selo.pack(side="left", padx=(T.PAD_M, T.PAD_S), pady=T.PAD_S)
        self._action_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._rotulo_do_selo = ctk.CTkLabel(
            self._selo, textvariable=self._action_var, font=T.FONT_SMALL,
            text_color=T.MUTED)
        self._rotulo_do_selo.pack(side="left", padx=(0, T.PAD_M))
        self._pintar_selo(None)

        self._btn = botao(p, "Converter agora", nome_do_icone="arrow-right-left",
                          command=self._convert, state="disabled")
        self._btn.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_XL))

    def _montar_cartao(self, parent, nome_do_icone, titulo, descricao):
        card = ctk.CTkFrame(parent, fg_color=T.SURFACE_3, corner_radius=T.RADIUS_SM,
                            border_width=1, border_color=T.BORDER)
        tk.Label(card, image=icone(nome_do_icone, T.ICON_LG, T.FG_SECONDARY),
                 bg=T.SURFACE_3, bd=0).pack(pady=(T.PAD_L, T.PAD_S))
        ctk.CTkLabel(card, text=titulo, font=T.FONT_BUTTON,
                     text_color=T.FG).pack()
        ctk.CTkLabel(card, text=descricao, font=T.FONT_SMALL,
                     text_color=T.MUTED, wraplength=220).pack(pady=(T.PAD_XS, T.PAD_L))
        return card

    def _pintar_selo(self, modo):
        if modo is None:
            self._icone_do_selo.config(image=icone("info", T.ICON_SM, T.MUTED))
            self._rotulo_do_selo.configure(text_color=T.MUTED)
        else:
            self._icone_do_selo.config(
                image=icone("arrow-right-left", T.ICON_SM, T.SUCCESS_TEXT))
            self._rotulo_do_selo.configure(text_color=T.SUCCESS_TEXT)

    def _select(self):
        path = filedialog.askopenfilename(
            filetypes=[("Documentos", "*.pdf *.docx")])
        if not path:
            return
        self._path = path
        self._drop.set_file(os.path.basename(path))
        ext = os.path.splitext(path)[1].lower()

        if ext == ".pdf":
            self._mode = "pdf2word"
            self._action_var.set("Será convertido para Word (.docx)")
            self._card_pdf.configure(border_color=T.ACCENT)
            self._card_word.configure(border_color=T.BORDER)
        else:
            self._mode = "word2pdf"
            self._action_var.set("Será convertido para PDF")
            self._card_word.configure(border_color=T.ACCENT)
            self._card_pdf.configure(border_color=T.BORDER)

        self._pintar_selo(self._mode)
        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._mode = None
        self._action_var.set("Nenhum arquivo selecionado")
        self._pintar_selo(None)
        self._btn.configure(state="disabled")
        self._card_pdf.configure(border_color=T.BORDER)
        self._card_word.configure(border_color=T.BORDER)
        self.set_status("Pronto")

    def _convert(self):
        if self._mode == "pdf2word":
            save = filedialog.asksaveasfilename(
                defaultextension=".docx", filetypes=[("Word", "*.docx")])
        else:
            save = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        self._btn.configure(state="disabled")
        mode, path = self._mode, self._path
        self.set_status("Convertendo…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: self._converter(mode, path, save),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    @staticmethod
    def _converter(mode, path, save):
        pdf_io.validar_destino(save, path)
        if mode != "pdf2word":
            docx_to_pdf(path, save)
            return
        cv = Converter(path)
        try:
            cv.convert(save, start=0, end=None)
        finally:
            cv.close()   # sem o finally, uma falha na conversão vazava o arquivo aberto

    def _done(self, save):
        self._btn.configure(state="normal")
        self.set_status(f"Salvo: {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")

    def _error(self, msg):
        self._btn.configure(state="normal")
        self.set_status("Erro na conversão", "erro")
        messagebox.showerror("Erro", msg)
