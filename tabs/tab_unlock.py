import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter

from widgets import file_row, section_title


class TabUnlock:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(
            p,
            "Remover Senha do PDF",
            "Gere uma cópia do PDF sem proteção por senha.",
        )

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        ttk.Button(
            p,
            text="Selecionar PDF protegido",
            style="Ghost.TButton",
            command=self._select,
        ).pack(anchor="w", pady=(0, 20))

        ttk.Label(p, text="Senha atual:", style="Sub.TLabel").pack(anchor="w")
        self._pwd = ttk.Entry(p, show="●", width=30)
        self._pwd.pack(anchor="w", pady=(4, 20))

        self._btn = ttk.Button(
            p,
            text="Remover senha e salvar",
            style="Primary.TButton",
            command=self._unlock,
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
        self._pwd.delete(0, tk.END)
        self._btn.config(state="disabled")
        self._btn_clear.pack_forget()
        self.set_status("Pronto")

    def _unlock(self):
        pwd = self._pwd.get()
        if not pwd:
            messagebox.showwarning("Aviso", "Digite a senha atual do PDF.")
            return
        save = filedialog.asksaveasfilename(
            initialfile=f"desbloqueado_{os.path.basename(self._path)}",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
        )
        if not save:
            return
        try:
            self.set_status("Removendo senha…")
            reader = PdfReader(self._path)
            if reader.is_encrypted:
                result = reader.decrypt(pwd)
                if not result:
                    messagebox.showerror(
                        "Senha incorreta", "A senha informada está incorreta."
                    )
                    self.set_status("Senha incorreta")
                    return
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"PDF desbloqueado: {os.path.basename(save)}")
            messagebox.showinfo(
                "Sucesso",
                f"Senha removida com sucesso!\nSalvo em: {os.path.basename(save)}",
            )
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("Erro ao remover senha")
