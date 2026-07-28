import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz

from .. import theme as T
from ..core import pdf_io, reorder
from ..core.background import executar_em_thread
from ..widgets import botao, criar_area_rolavel, estado_vazio, icone


class TabImageToPdf:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list[str] = []
        self._build(parent)

    def _build(self, p):
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XL, T.PAD_M))

        botao(ctrl, "Adicionar imagens", nome_do_icone="plus", variante="secundario",
              altura=36, command=self._add).pack(side="left")
        botao(ctrl, "Limpar lista", nome_do_icone="trash-2", variante="fantasma",
              altura=36, command=self._clear_all).pack(side="right")

        self._canvas, self._inner = criar_area_rolavel(
            p, fill="both", expand=True, padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._rebuild()

        self._btn_generate = botao(p, "Gerar PDF", nome_do_icone="file-image",
                                   variante="sucesso", command=self._generate)
        self._btn_generate.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_XL))

    # ── Lista ─────────────────────────────────────────────────────────────────

    def _rebuild(self):
        for w in self._inner.winfo_children():
            w.destroy()

        if not self._files:
            estado_vazio(self._inner, "file-image", "Nenhuma imagem na lista")
            return

        for i, path in enumerate(self._files):
            self._add_row(i, path)

    def _add_row(self, i: int, path: str):
        row = tk.Frame(self._inner, bg=T.GRID_BG)
        row.pack(fill="x", padx=T.PAD_S, pady=1)

        tk.Label(row, text=f"{i + 1}", bg=T.GRID_BG, fg=T.MUTED,
                 font=T.FONT_SMALL, width=3, anchor="e").pack(side="left",
                                                              padx=(T.PAD_S, 0))
        tk.Label(row, image=icone("file-image", T.ICON_SM, T.ACCENT_TEXT),
                 bg=T.GRID_BG, bd=0).pack(side="left", padx=T.PAD_S)
        tk.Label(row, text=os.path.basename(path), bg=T.GRID_BG, fg=T.FG,
                 font=T.FONT_BODY, anchor="w").pack(side="left", fill="x",
                                                    expand=True, pady=T.PAD_S)

        remover = tk.Label(row, image=icone("x", T.ICON_SM, T.MUTED),
                           bg=T.GRID_BG, bd=0, cursor="hand2")
        remover.pack(side="right", padx=T.PAD_M)
        remover.bind("<Button-1>", lambda _, idx=i: self._remove(idx))

        for nome_do_icone, direcao in (("chevron-down", 1), ("chevron-up", -1)):
            seta = tk.Label(row, image=icone(nome_do_icone, T.ICON_SM, T.MUTED),
                            bg=T.GRID_BG, bd=0, cursor="hand2")
            seta.pack(side="right", padx=T.PAD_XS)
            seta.bind("<Button-1>", lambda _, idx=i, d=direcao: self._move(idx, d))

        tk.Frame(self._inner, bg=T.BORDER, height=1).pack(fill="x", padx=T.PAD_S)

    def _add(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if not files:
            return
        self._files.extend(files)
        self._rebuild()
        self.set_status(f"{len(files)} imagem(ns) adicionada(s) — total: {len(self._files)}")

    def _remove(self, idx: int):
        if 0 <= idx < len(self._files):
            self._files.pop(idx)
            self._rebuild()
            self.set_status(
                f"{len(self._files)} imagem(ns) na lista" if self._files else "Lista vazia")

    def _clear_all(self):
        self._files.clear()
        self._rebuild()
        self.set_status("Lista limpa")

    def _move(self, idx: int, direction: int):
        self._files, _ = reorder.mover(
            self._files, idx, "left" if direction < 0 else "right")
        self._rebuild()

    # ── Gerar ─────────────────────────────────────────────────────────────────

    def _generate(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos uma imagem.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        files = list(self._files)
        self._btn_generate.configure(state="disabled")
        self.set_status("Gerando PDF…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: self._montar_pdf(files, save),
            ao_terminar=lambda _: self._done(save, len(files)),
            ao_falhar=self._error,
        )

    @staticmethod
    def _montar_pdf(files, save):
        pdf_io.validar_destino(save, *files)
        doc = fitz.open()
        try:
            for img_path in files:
                img = fitz.open(img_path)
                try:
                    pdfbytes = img.convert_to_pdf()
                finally:
                    img.close()
                imgpdf = fitz.open("pdf", pdfbytes)
                try:
                    doc.insert_pdf(imgpdf)
                finally:
                    imgpdf.close()
            doc.save(save)
        finally:
            doc.close()

    def _done(self, save: str, n: int):
        self._btn_generate.configure(state="normal")
        self.set_status(f"PDF gerado: {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", f"PDF criado com {n} página(s)!")

    def _error(self, msg: str):
        self._btn_generate.configure(state="normal")
        self.set_status("Erro ao gerar PDF", "erro")
        messagebox.showerror("Erro", msg)
