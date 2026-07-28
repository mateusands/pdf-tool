import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import DropZone, GrupoPills, botao

FORMATOS = [("png", "PNG"), ("jpg", "JPG")]
RESOLUCOES = [("72", "72 · Baixa"), ("150", "150 · Média"), ("300", "300 · Alta")]


class TabPdfToImage:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

        opcoes = ctk.CTkFrame(p, fg_color="transparent")
        opcoes.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_M))

        ctk.CTkLabel(opcoes, text="Formato de saída", font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(anchor="w", pady=(0, T.PAD_S))
        self._formato = GrupoPills(opcoes, FORMATOS, valor_inicial="png", largura=96)

        ctk.CTkLabel(opcoes, text="Resolução (DPI)", font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(anchor="w",
                                                     pady=(T.PAD_L, T.PAD_S))
        self._dpi = GrupoPills(opcoes, RESOLUCOES, valor_inicial="150", largura=124)

        # Progresso
        progresso = ctk.CTkFrame(p, fg_color="transparent")
        progresso.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_M, T.PAD_M))

        self._progress = ctk.CTkProgressBar(progresso, fg_color=T.SURFACE_3,
                                            progress_color=T.ACCENT, height=6,
                                            corner_radius=3)
        self._progress.pack(fill="x")
        self._progress.set(0)

        self._prog_var = tk.StringVar(value="")
        ctk.CTkLabel(progresso, textvariable=self._prog_var, font=T.FONT_SMALL,
                     text_color=T.MUTED).pack(anchor="w", pady=(T.PAD_XS, 0))

        self._btn = botao(p, "Exportar imagens", nome_do_icone="image-down",
                          command=self._run, state="disabled")
        self._btn.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_XL))

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._drop.set_file(os.path.basename(path))
        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._btn.configure(state="disabled")
        self._progress.set(0)
        self._prog_var.set("")
        self.set_status("Pronto")

    def _run(self):
        folder = filedialog.askdirectory(title="Escolha a pasta de saída")
        if not folder:
            return
        fmt = self._formato.valor()
        dpi = int(self._dpi.valor())
        path = self._path
        self._btn.configure(state="disabled")
        self._progress.set(0)
        self.set_status("Exportando páginas…", "ocupado")

        def update(current, total):
            pct = current / total
            self.root.after(0, lambda c=current, t=total, v=pct: (
                self._progress.set(v),
                self._prog_var.set(f"Página {c} de {t}"),
            ))

        def exportar():
            with pdf_io.abrir_documento(path) as doc:
                n = len(doc)
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(matrix=mat)
                    pix.save(os.path.join(folder, f"pagina_{i + 1}.{fmt}"))
                    update(i + 1, n)
                return n

        executar_em_thread(
            self.root, exportar,
            ao_terminar=lambda n: self._done(n, folder),
            ao_falhar=self._error,
        )

    def _done(self, n: int, folder: str):
        self._btn.configure(state="normal")
        self._progress.set(1)
        self._prog_var.set(f"{n} imagem(ns) salva(s)")
        self.set_status(f"{n} imagem(ns) salva(s) em {os.path.basename(folder)}", "ok")
        messagebox.showinfo("Sucesso", f"{n} página(s) convertida(s)!\n\nPasta: {folder}")

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("Erro na exportação", "erro")
        messagebox.showerror("Erro", msg)
