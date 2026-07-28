import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import DropZone, GrupoPills, botao

LEVELS = {
    "Baixa": {"garbage": 1, "deflate": False},
    "Média": {"garbage": 3, "deflate": True},
    "Alta":  {"garbage": 4, "deflate": True, "clean": True},
}

LEVEL_DESC = {
    "Baixa": "Mais rápido, arquivo maior",
    "Média": "Equilíbrio entre tamanho e tempo",
    "Alta":  "Mais lento, arquivo menor",
}


def _fmt(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024**2:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1024**2:.1f} MB"


class TabCompress:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF para compactar",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

        bloco = ctk.CTkFrame(p, fg_color="transparent")
        bloco.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_M))

        ctk.CTkLabel(bloco, text="Nível de compactação", font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(anchor="w", pady=(0, T.PAD_S))

        self._nivel = GrupoPills(
            bloco, [(n, n) for n in LEVELS], valor_inicial="Média",
            ao_mudar=lambda v: self._desc_var.set(LEVEL_DESC[v]), largura=110,
        )

        self._desc_var = tk.StringVar(value=LEVEL_DESC["Média"])
        ctk.CTkLabel(bloco, textvariable=self._desc_var, font=T.FONT_SMALL,
                     text_color=T.MUTED).pack(anchor="w", pady=(T.PAD_S, 0))

        self._size_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self._size_var, font=T.FONT_BODY,
                     text_color=T.FG_SECONDARY).pack(anchor="w", padx=T.PAD_XL,
                                                     pady=(0, T.PAD_M))

        self._btn = botao(p, "Compactar e salvar", nome_do_icone="shrink",
                          command=self._run, state="disabled")
        self._btn.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_XL))

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        size = os.path.getsize(path)
        self._drop.set_file(os.path.basename(path), _fmt(size))
        self._size_var.set(f"Tamanho original: {_fmt(size)}")
        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._size_var.set("")
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _run(self):
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        opts = LEVELS[self._nivel.valor()]
        path, orig = self._path, os.path.getsize(self._path)
        self._btn.configure(state="disabled")
        self.set_status("Compactando…", "ocupado")

        executar_em_thread(
            self.root,
            lambda: pdf_io.comprimir_pdf(path, save, **opts),
            ao_terminar=lambda final: self._done(orig, final, save),
            ao_falhar=self._error,
        )

    def _done(self, orig: int, final: int, save: str):
        self._btn.configure(state="normal")
        ratio = (1 - final / orig) * 100 if orig else 0
        self._size_var.set(
            f"Original: {_fmt(orig)}   →   Final: {_fmt(final)}   ({ratio:.1f}% menor)")
        self.set_status(f"Compactado: {os.path.basename(save)}", "ok")
        messagebox.showinfo(
            "Sucesso",
            f"PDF compactado com sucesso!\n\n"
            f"Original: {_fmt(orig)}\nFinal: {_fmt(final)}\nRedução: {ratio:.1f}%",
        )

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("Erro ao compactar", "erro")
        messagebox.showerror("Erro", msg)
