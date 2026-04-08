import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz

from widgets import file_row, section_title


class TabPdfToImage:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(
            p, "PDF → Imagem", "Converta cada página do PDF em uma imagem."
        )

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        ttk.Button(
            p, text="Selecionar PDF", style="Ghost.TButton", command=self._select
        ).pack(anchor="w", pady=(0, 16))

        ttk.Label(p, text="Formato de saída:", style="Sub.TLabel").pack(anchor="w")
        self._fmt_var = tk.StringVar(value="png")
        fmt_row = ttk.Frame(p, style="Card.TFrame")
        fmt_row.pack(fill="x", pady=(4, 12))
        ttk.Radiobutton(fmt_row, text="PNG", variable=self._fmt_var, value="png").pack(
            side="left", padx=(0, 16)
        )
        ttk.Radiobutton(fmt_row, text="JPG", variable=self._fmt_var, value="jpg").pack(
            side="left"
        )

        ttk.Label(p, text="Resolução (DPI):", style="Sub.TLabel").pack(anchor="w")
        self._dpi_var = tk.StringVar(value="150")
        for dpi, label in [
            ("72", "72 – Baixa"),
            ("150", "150 – Média"),
            ("300", "300 – Alta"),
        ]:
            ttk.Radiobutton(
                p, text=label, variable=self._dpi_var, value=dpi
            ).pack(anchor="w", pady=2)

        self._progress = ttk.Progressbar(p, mode="determinate", maximum=100)
        self._progress.pack(fill="x", pady=(12, 4))
        self._prog_lbl = ttk.Label(p, text="", style="Muted.TLabel")
        self._prog_lbl.pack(anchor="w", pady=(0, 8))

        self._btn = ttk.Button(
            p,
            text="Converter e salvar imagens",
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
        self._btn.config(state="normal")
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._file_var.set("Nenhum arquivo selecionado")
        self._btn.config(state="disabled")
        self._btn_clear.pack_forget()
        self._progress["value"] = 0
        self._prog_lbl.config(text="")
        self.set_status("Pronto")

    def _run(self):
        folder = filedialog.askdirectory(title="Escolha a pasta de saída")
        if not folder:
            return
        fmt = self._fmt_var.get()
        dpi = int(self._dpi_var.get())
        path = self._path
        self._btn.config(state="disabled")
        self._progress["value"] = 0
        self.set_status("Convertendo…")

        def update(current, total):
            pct = int(current / total * 100)
            self.root.after(
                0,
                lambda c=current, t=total, p=pct: (
                    self._progress.configure(value=p),
                    self._prog_lbl.configure(text=f"Página {c} de {t}"),
                ),
            )

        def task():
            try:
                doc = fitz.open(path)
                n = len(doc)
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(matrix=mat)
                    pix.save(os.path.join(folder, f"pagina_{i+1}.{fmt}"))
                    update(i + 1, n)
                doc.close()
                self.root.after(0, lambda: self._done(n, folder))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, n: int, folder: str):
        self._btn.config(state="normal")
        self._progress["value"] = 100
        self._prog_lbl.config(text=f"{n} imagem(ns) salva(s)")
        self.set_status(f"{n} imagem(ns) salva(s) em {os.path.basename(folder)}")
        messagebox.showinfo(
            "Sucesso", f"{n} página(s) convertida(s)!\nPasta: {folder}"
        )

    def _error(self, msg: str):
        self._btn.config(state="normal")
        self.set_status("Erro na conversão")
        messagebox.showerror("Erro", msg)
