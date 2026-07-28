import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io, reorder
from ..core.background import executar_em_thread
from ..widgets import botao, criar_area_rolavel, estado_vazio, icone


class TabMerge:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list[str] = []
        self._rows: list[tk.Frame] = []
        self._selected_idx: int | None = None
        self._build(parent)

    def _build(self, p):
        # Ações
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.pack(fill="x", padx=T.PAD_XL, pady=(T.PAD_XL, T.PAD_M))

        botao(ctrl, "Adicionar PDFs", nome_do_icone="plus", variante="secundario",
              altura=36, command=self._add).pack(side="left", padx=(0, T.PAD_S))
        botao(ctrl, "Subir", nome_do_icone="chevron-up", variante="fantasma",
              altura=36, largura=96, command=self._move_up).pack(side="left",
                                                                 padx=(0, T.PAD_XS))
        botao(ctrl, "Descer", nome_do_icone="chevron-down", variante="fantasma",
              altura=36, largura=104, command=self._move_down).pack(side="left")
        botao(ctrl, "Limpar lista", nome_do_icone="trash-2", variante="fantasma",
              altura=36, command=self._clear_all).pack(side="right")

        self._canvas, self._inner = criar_area_rolavel(
            p, fill="both", expand=True, padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._rebuild_list()

        self._btn_merge = botao(p, "Juntar e salvar PDF", nome_do_icone="combine",
                                variante="sucesso", command=self._merge)
        self._btn_merge.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_XL))

    # ── Lista ─────────────────────────────────────────────────────────────────

    def _rebuild_list(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._rows = []

        if not self._files:
            estado_vazio(self._inner, "combine", "Nenhum arquivo na lista")
            self._aplicar_selecao()
            return

        for i, path in enumerate(self._files):
            row = tk.Frame(self._inner, bg=T.GRID_BG)
            row.pack(fill="x", padx=T.PAD_S, pady=1)
            self._rows.append(row)

            tk.Label(row, text=f"{i + 1}", bg=T.GRID_BG, fg=T.MUTED,
                     font=T.FONT_SMALL, width=3, anchor="e").pack(side="left",
                                                                  padx=(T.PAD_S, 0))
            tk.Label(row, image=icone("file-text", T.ICON_SM, T.ACCENT_TEXT),
                     bg=T.GRID_BG, bd=0).pack(side="left", padx=T.PAD_S)
            tk.Label(row, text=os.path.basename(path), bg=T.GRID_BG, fg=T.FG,
                     font=T.FONT_BODY, anchor="w").pack(side="left", fill="x",
                                                        expand=True, pady=T.PAD_S)

            remover = tk.Label(row, image=icone("x", T.ICON_SM, T.MUTED),
                               bg=T.GRID_BG, bd=0, cursor="hand2")
            remover.pack(side="right", padx=T.PAD_M)
            remover.bind("<Button-1>", lambda _, idx=i: self._remove(idx))

            for w in [row] + list(row.winfo_children()):
                if w is not remover:
                    w.bind("<Button-1>", lambda _, idx=i: self._select_item(idx))

            tk.Frame(self._inner, bg=T.BORDER, height=1).pack(fill="x", padx=T.PAD_S)

        # A seleção sobrevive ao redesenho — sem isso "Subir" só funcionava uma vez.
        self._aplicar_selecao()

    def _aplicar_selecao(self):
        if self._selected_idx is not None and not 0 <= self._selected_idx < len(self._files):
            self._selected_idx = None
        for i, row in enumerate(self._rows):
            bg = T.SURFACE_4 if i == self._selected_idx else T.GRID_BG
            row.config(bg=bg)
            for w in row.winfo_children():
                w.config(bg=bg)

    def _select_item(self, idx):
        self._selected_idx = idx
        self._aplicar_selecao()

    def _add(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if not files:
            return
        self._files.extend(files)
        self._rebuild_list()
        self.set_status(f"{len(files)} arquivo(s) adicionados — total: {len(self._files)}")

    def _remove(self, idx: int):
        if 0 <= idx < len(self._files):
            self._files.pop(idx)
            self._selected_idx = None
            self._rebuild_list()
            self.set_status(
                f"{len(self._files)} arquivo(s) na lista" if self._files else "Lista vazia")

    def _clear_all(self):
        self._files.clear()
        self._selected_idx = None
        self._rebuild_list()
        self.set_status("Lista limpa")

    def _move_up(self):
        self._files, self._selected_idx = reorder.mover(
            self._files, self._selected_idx, "left")
        self._rebuild_list()

    def _move_down(self):
        self._files, self._selected_idx = reorder.mover(
            self._files, self._selected_idx, "right")
        self._rebuild_list()

    # ── Juntar ────────────────────────────────────────────────────────────────

    def _merge(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos um arquivo PDF.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        files = list(self._files)
        self._btn_merge.configure(state="disabled")
        self.set_status("Juntando PDFs…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.juntar_pdfs(files, save),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    def _done(self, save: str):
        self._btn_merge.configure(state="normal")
        self.set_status(f"Salvo: {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", "PDFs juntados com sucesso!")

    def _error(self, msg: str):
        self._btn_merge.configure(state="normal")
        self.set_status("Erro ao juntar PDFs", "erro")
        messagebox.showerror("Erro", msg)
