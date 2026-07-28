import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import DropZone, GrupoPills, ThumbnailGrid, botao

ANGULOS = [
    (90,  "90° à direita", "rotate-cw"),
    (-90, "90° à esquerda", "rotate-ccw"),
    (180, "180°", "flip-vertical"),
]


class TabRotate:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

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

        # Controles no rodapé — empacotados antes da grade, que é elástica
        rodape = ctk.CTkFrame(p, fg_color="transparent")
        rodape.pack(side="bottom", fill="x", padx=T.PAD_XL, pady=(T.PAD_S, T.PAD_XL))

        ctk.CTkLabel(rodape, text="Rotação", font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(anchor="w", pady=(0, T.PAD_XS))
        self._angulo = GrupoPills(rodape, ANGULOS, valor_inicial=90, largura=140)

        self._btn_save = botao(rodape, "Girar e salvar", nome_do_icone="save",
                               command=self._save, state="disabled")
        self._btn_save.pack(anchor="w", pady=(T.PAD_M, 0))

        self._grid = ThumbnailGrid(p)
        self._grid.on_change(self._on_change)

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self.set_status("Carregando páginas…", "ocupado")
        try:
            n = self._grid.load_pdf(path)
        except Exception as e:
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

    def _on_change(self, selected):
        n = len(selected)
        self._count_var.set(f"{n} página(s) selecionada(s)" if n else "")
        self._btn_save.configure(state="normal" if n else "disabled")

    def _save(self):
        selected = self._grid.get_selected()
        if not selected:
            return
        angle = self._angulo.valor()
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        path = self._path
        self._btn_save.configure(state="disabled")
        self.set_status("Girando páginas…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.girar_pdf(path, save, selected, angle),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    def _done(self, save: str):
        self._btn_save.configure(state="normal")
        self.set_status(f"Salvo: {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", "PDF girado e salvo com sucesso!")

    def _error(self, msg: str):
        self._btn_save.configure(state="normal")
        self.set_status("Erro ao girar", "erro")
        messagebox.showerror("Erro", msg)
