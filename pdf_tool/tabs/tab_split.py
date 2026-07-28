import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import DropZone, ThumbnailGrid, botao


class TabSplit:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

        # Barra: contagem à esquerda, ações de seleção à direita
        bar = ctk.CTkFrame(p, fg_color="transparent")
        bar.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_S))

        self._count_var = tk.StringVar(value="")
        ctk.CTkLabel(bar, textvariable=self._count_var,
                     font=T.FONT_LABEL, text_color=T.ACCENT_TEXT).pack(side="left")

        self._btn_none = botao(bar, "Limpar", variante="fantasma",
                               altura=30, largura=78,
                               command=lambda: self._grid.deselect_all())
        self._btn_all = botao(bar, "Selecionar tudo", variante="secundario",
                              altura=30, largura=132,
                              command=lambda: self._grid.select_all())

        self._btn_save = botao(p, "Salvar páginas selecionadas", nome_do_icone="save",
                               command=self._save, state="disabled")
        self._btn_save.pack(side="bottom", anchor="w", padx=T.PAD_XL,
                            pady=(T.PAD_S, T.PAD_XL))

        self._grid = ThumbnailGrid(p)
        self._grid.on_change(self._on_selection_change)

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self.set_status("Carregando páginas…", "ocupado")
        try:
            n = self._grid.load_pdf(path)
        except Exception as e:
            # PDF protegido ou corrompido: a grade anterior fica intacta.
            self.set_status("Não foi possível abrir o PDF", "erro")
            messagebox.showerror("Erro ao abrir", str(e))
            return

        self._path = path
        self._drop.set_file(os.path.basename(path), f"{n} página(s)")
        self._btn_all.pack(side="right")
        self._btn_none.pack(side="right", padx=(0, T.PAD_S))
        self.set_status(f"{n} página(s) carregadas — clique para selecionar")

    def _clear(self):
        self._path = None
        self._count_var.set("")
        self._btn_all.pack_forget()
        self._btn_none.pack_forget()
        self._btn_save.configure(state="disabled")
        self._grid.clear()
        self.set_status("Pronto")

    def _on_selection_change(self, selected):
        n = len(selected)
        self._count_var.set(f"{n} página(s) selecionada(s)" if n else "")
        self._btn_save.configure(state="normal" if n else "disabled")

    def _save(self):
        selected = self._grid.get_selected()
        if not selected:
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        path = self._path
        self._btn_save.configure(state="disabled")
        self.set_status("Salvando páginas…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.dividir_pdf(path, save, selected),
            ao_terminar=lambda n: self._done(n, save),
            ao_falhar=self._error,
        )

    def _done(self, n: int, save: str):
        self._btn_save.configure(state="normal")
        self.set_status(f"{n} página(s) salvas em {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", f"{n} página(s) salvas com sucesso!")

    def _error(self, msg: str):
        self._btn_save.configure(state="normal")
        self.set_status("Erro ao salvar", "erro")
        messagebox.showerror("Erro", msg)
