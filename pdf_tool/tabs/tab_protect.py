import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import CampoSenha, DropZone, botao, icone


class TabProtect:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF para proteger",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

        form = ctk.CTkFrame(p, fg_color="transparent")
        form.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._senha = CampoSenha(form, "Nova senha")
        self._confirmacao = CampoSenha(form, "Confirmar senha")

        # Selo de criptografia — o usuário merece saber o que está sendo aplicado
        selo = ctk.CTkFrame(p, fg_color=T.SURFACE_3, corner_radius=T.RADIUS_SM)
        selo.pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_L))
        tk.Label(selo, image=icone("lock", T.ICON_SM, T.SUCCESS_TEXT),
                 bg=T.SURFACE_3, bd=0).pack(side="left", padx=(T.PAD_M, T.PAD_S),
                                            pady=T.PAD_S)
        ctk.CTkLabel(selo, text="Criptografia AES-256", font=T.FONT_SMALL,
                     text_color=T.FG_SECONDARY).pack(side="left", padx=(0, T.PAD_M))

        self._btn = botao(p, "Proteger e salvar", nome_do_icone="lock",
                          command=self._protect, state="disabled")
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
        self._senha.limpar()
        self._confirmacao.limpar()
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _protect(self):
        senha = self._senha.valor()
        if not senha:
            messagebox.showwarning("Aviso", "Digite uma senha.")
            return
        if senha != self._confirmacao.valor():
            messagebox.showwarning("Aviso", "As senhas não coincidem.")
            return
        save = filedialog.asksaveasfilename(
            initialfile=f"protegido_{os.path.basename(self._path)}",
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return

        path = self._path
        self._btn.configure(state="disabled")
        self.set_status("Protegendo PDF…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.proteger_pdf(path, save, senha),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    def _done(self, save: str):
        self._btn.configure(state="normal")
        self.set_status(f"PDF protegido: {os.path.basename(save)}", "ok")
        messagebox.showinfo(
            "Sucesso",
            f"PDF protegido com AES-256!\n\nSalvo em: {os.path.basename(save)}")

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("Erro ao proteger PDF", "erro")
        messagebox.showerror("Erro", msg)
