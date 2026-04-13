import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pypdf import PdfReader, PdfWriter

import theme as T
from widgets import DropZone, section_title


class TabProtect:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "🔒  Proteger PDF com Senha",
                      "Gere uma cópia do PDF protegida por senha.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF para proteger",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Password fields
        form = ctk.CTkFrame(p, fg_color="transparent")
        form.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_L))

        ctk.CTkLabel(form, text="Nova senha:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_XS))

        pwd1_row = ctk.CTkFrame(form, fg_color="transparent")
        pwd1_row.pack(fill="x", pady=(0, T.PAD_M))
        self._pwd1 = ctk.CTkEntry(
            pwd1_row, show="●", width=280, height=38,
            fg_color=T.BG_SECONDARY, border_color=T.BORDER,
            text_color=T.FG, font=T.FONT_BODY, corner_radius=8,
        )
        self._pwd1.pack(side="left")
        self._show1 = False
        ctk.CTkButton(
            pwd1_row, text="👁", width=38, height=38,
            fg_color=T.BG_SECONDARY, hover_color=T.SURFACE_HOVER,
            border_width=1, border_color=T.BORDER, corner_radius=8,
            cursor="hand2", font=T.FONT_BODY,
            command=lambda: self._toggle_show(1),
        ).pack(side="left", padx=(6, 0))

        ctk.CTkLabel(form, text="Confirmar senha:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_XS))

        pwd2_row = ctk.CTkFrame(form, fg_color="transparent")
        pwd2_row.pack(fill="x")
        self._pwd2 = ctk.CTkEntry(
            pwd2_row, show="●", width=280, height=38,
            fg_color=T.BG_SECONDARY, border_color=T.BORDER,
            text_color=T.FG, font=T.FONT_BODY, corner_radius=8,
        )
        self._pwd2.pack(side="left")
        self._show2 = False
        ctk.CTkButton(
            pwd2_row, text="👁", width=38, height=38,
            fg_color=T.BG_SECONDARY, hover_color=T.SURFACE_HOVER,
            border_width=1, border_color=T.BORDER, corner_radius=8,
            cursor="hand2", font=T.FONT_BODY,
            command=lambda: self._toggle_show(2),
        ).pack(side="left", padx=(6, 0))

        # Protect button
        self._btn = ctk.CTkButton(
            p, text="🔒  Proteger e salvar",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._protect, state="disabled",
        )
        self._btn.pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    def _toggle_show(self, field: int):
        if field == 1:
            self._show1 = not self._show1
            self._pwd1.configure(show="" if self._show1 else "●")
        else:
            self._show2 = not self._show2
            self._pwd2.configure(show="" if self._show2 else "●")

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
        self._pwd1.delete(0, tk.END)
        self._pwd2.delete(0, tk.END)
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _protect(self):
        pwd1 = self._pwd1.get()
        pwd2 = self._pwd2.get()
        if not pwd1:
            messagebox.showwarning("Aviso", "Digite uma senha.")
            return
        if pwd1 != pwd2:
            messagebox.showwarning("Aviso", "As senhas não coincidem.")
            return
        save = filedialog.asksaveasfilename(
            initialfile=f"protegido_{os.path.basename(self._path)}",
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        try:
            self.set_status("⏳  Protegendo PDF…")
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(pwd1)
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"✓  PDF protegido: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso",
                                f"PDF protegido com senha!\nSalvo em: {os.path.basename(save)}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("✗  Erro ao proteger PDF")
