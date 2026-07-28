import os
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io
from ..core.background import executar_em_thread
from ..widgets import CampoSenha, DropZone, botao


class TabUnlock:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        self._drop = DropZone(p, icon="lock", text="Selecione um PDF protegido",
                              subtitle="ou clique para escolher no computador",
                              command=self._select, on_clear=self._clear)

        form = ctk.CTkFrame(p, fg_color="transparent")
        form.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._senha = CampoSenha(form, "Senha atual do arquivo")

        ctk.CTkLabel(
            p, text="O arquivo original não é alterado — uma cópia sem senha é criada.",
            font=T.FONT_SMALL, text_color=T.MUTED,
        ).pack(anchor="w", padx=T.PAD_XL, pady=(0, T.PAD_L))

        self._btn = botao(p, "Remover senha e salvar", nome_do_icone="lock-open",
                          command=self._unlock, state="disabled")
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
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _unlock(self):
        senha = self._senha.valor()
        if not senha:
            messagebox.showwarning("Aviso", "Digite a senha atual do PDF.")
            return
        save = filedialog.asksaveasfilename(
            initialfile=f"desbloqueado_{os.path.basename(self._path)}",
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return

        path = self._path
        self._btn.configure(state="disabled")
        self.set_status("Removendo senha…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.desbloquear_pdf(path, save, senha),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    def _done(self, save: str):
        self._btn.configure(state="normal")
        self.set_status(f"PDF desbloqueado: {os.path.basename(save)}", "ok")
        messagebox.showinfo(
            "Sucesso", f"Senha removida com sucesso!\n\nSalvo em: {os.path.basename(save)}")

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("Não foi possível remover a senha", "erro")
        messagebox.showerror("Erro", msg)
