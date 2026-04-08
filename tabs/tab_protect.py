import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter

from widgets import file_row, section_title


class TabProtect:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(
            p,
            "Proteger PDF com Senha",
            "Gere uma cópia do PDF protegida por senha.",
        )

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        ttk.Button(
            p, text="Selecionar PDF", style="Ghost.TButton", command=self._select
        ).pack(anchor="w", pady=(0, 20))

        ttk.Label(p, text="Nova senha:", style="Sub.TLabel").pack(anchor="w")
        self._pwd1 = ttk.Entry(p, show="●", width=30)
        self._pwd1.pack(anchor="w", pady=(4, 12))

        ttk.Label(p, text="Confirmar senha:", style="Sub.TLabel").pack(anchor="w")
        self._pwd2 = ttk.Entry(p, show="●", width=30)
        self._pwd2.pack(anchor="w", pady=(4, 20))

        self._btn = ttk.Button(
            p,
            text="Proteger e salvar",
            style="Primary.TButton",
            command=self._protect,
            state="disabled",
        )
        self._btn.pack(anchor="w")

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._file_var.set(os.path.basename(path))
        self._btn.config(state="normal")
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._file_var.set("Nenhum arquivo selecionado")
        self._pwd1.delete(0, tk.END)
        self._pwd2.delete(0, tk.END)
        self._btn.config(state="disabled")
        self._btn_clear.pack_forget()
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
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        try:
            self.set_status("Protegendo PDF…")
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.encrypt(pwd1)
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"PDF protegido: {os.path.basename(save)}")
            messagebox.showinfo(
                "Sucesso",
                f"PDF protegido com senha!\nSalvo em: {os.path.basename(save)}",
            )
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("Erro ao proteger PDF")
