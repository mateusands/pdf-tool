import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz

from widgets import file_row, section_title

LEVELS = {
    "Baixa  (mais rápido, arquivo maior)": {"garbage": 1, "deflate": False},
    "Média  (equilibrado)": {"garbage": 3, "deflate": True},
    "Alta   (mais lento, arquivo menor)": {"garbage": 4, "deflate": True, "clean": True},
}


def _fmt(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024**2:
        return f"{n/1024:.1f} KB"
    return f"{n/1024**2:.1f} MB"


class TabCompress:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "Compactar PDF", "Reduza o tamanho de um arquivo PDF.")

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        ttk.Button(
            p, text="Selecionar PDF", style="Ghost.TButton", command=self._select
        ).pack(anchor="w", pady=(0, 16))

        ttk.Label(p, text="Nível de compactação:", style="Sub.TLabel").pack(anchor="w")
        self._level_var = tk.StringVar(value=list(LEVELS.keys())[1])
        for level in LEVELS:
            ttk.Radiobutton(p, text=level, variable=self._level_var, value=level).pack(
                anchor="w", pady=2
            )

        self._size_var = tk.StringVar(value="")
        ttk.Label(p, textvariable=self._size_var, style="Muted.TLabel").pack(
            anchor="w", pady=12
        )

        self._btn = ttk.Button(
            p,
            text="Compactar e salvar",
            style="Primary.TButton",
            command=self._run,
            state="disabled",
        )
        self._btn.pack(anchor="w")

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._file_var.set(os.path.basename(path))
        self._size_var.set(f"Tamanho original: {_fmt(os.path.getsize(path))}")
        self._btn.config(state="normal")
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._file_var.set("Nenhum arquivo selecionado")
        self._size_var.set("")
        self._btn.config(state="disabled")
        self._btn_clear.pack_forget()
        self.set_status("Pronto")

    def _run(self):
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
        )
        if not save:
            return
        opts = LEVELS[self._level_var.get()]
        path, orig = self._path, os.path.getsize(self._path)
        self._btn.config(state="disabled")
        self.set_status("Compactando…")

        def task():
            try:
                doc = fitz.open(path)
                doc.save(save, **opts)
                doc.close()
                final = os.path.getsize(save)
                self.root.after(0, lambda: self._done(orig, final, save))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, orig: int, final: int, save: str):
        self._btn.config(state="normal")
        ratio = (1 - final / orig) * 100 if orig else 0
        self._size_var.set(
            f"Original: {_fmt(orig)}  →  Final: {_fmt(final)}  ({ratio:.1f}% menor)"
        )
        self.set_status(f"Compactado: {os.path.basename(save)}")
        messagebox.showinfo(
            "Sucesso",
            f"PDF compactado com sucesso!\n"
            f"Original: {_fmt(orig)}\n"
            f"Final: {_fmt(final)}\n"
            f"Redução: {ratio:.1f}%",
        )

    def _error(self, msg: str):
        self._btn.config(state="normal")
        self.set_status("Erro ao compactar")
        messagebox.showerror("Erro", msg)
