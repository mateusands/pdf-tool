import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pypdf import PdfReader, PdfWriter

import theme as T
from widgets import DropZone, section_title


class TabUnlock:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "🔓  Remover Senha do PDF",
                      "Gere uma cópia do PDF sem proteção por senha.")

        self._drop = DropZone(p, icon="🔒", text="Selecione um PDF protegido",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Password field
        form = ctk.CTkFrame(p, fg_color="transparent")
        form.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_L))

        ctk.CTkLabel(form, text="Senha atual:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_XS))

        pwd_row = ctk.CTkFrame(form, fg_color="transparent")
        pwd_row.pack(fill="x")
        self._pwd = ctk.CTkEntry(
            pwd_row, show="●", width=280, height=38,
            fg_color=T.BG_SECONDARY, border_color=T.BORDER,
            text_color=T.FG, font=T.FONT_BODY, corner_radius=8,
        )
        self._pwd.pack(side="left")
        self._show_pwd = False
        ctk.CTkButton(
            pwd_row, text="👁", width=38, height=38,
            fg_color=T.BG_SECONDARY, hover_color=T.SURFACE_HOVER,
            border_width=1, border_color=T.BORDER, corner_radius=8,
            cursor="hand2", font=T.FONT_BODY,
            command=self._toggle_show,
        ).pack(side="left", padx=(6, 0))

        # Unlock button
        self._btn = ctk.CTkButton(
            p, text="🔓  Remover senha e salvar",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._unlock, state="disabled",
        )
        self._btn.pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    def _toggle_show(self):
        self._show_pwd = not self._show_pwd
        self._pwd.configure(show="" if self._show_pwd else "●")

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
        self._pwd.delete(0, tk.END)
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _unlock(self):
        pwd = self._pwd.get()
        if not pwd:
            messagebox.showwarning("Aviso", "Digite a senha atual do PDF.")
            return
        save = filedialog.asksaveasfilename(
            initialfile=f"desbloqueado_{os.path.basename(self._path)}",
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        try:
            self.set_status("⏳  Removendo senha…")
            reader = PdfReader(self._path)
            if reader.is_encrypted:
                result = reader.decrypt(pwd)
                if not result:
                    messagebox.showerror("Senha incorreta",
                                         "A senha informada está incorreta.")
                    self.set_status("✗  Senha incorreta")
                    return
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"✓  PDF desbloqueado: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso",
                                f"Senha removida com sucesso!\nSalvo em: {os.path.basename(save)}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("✗  Erro ao remover senha")
