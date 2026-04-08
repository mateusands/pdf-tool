import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter

from constants import PRIMARY
from widgets import ThumbnailGrid, file_row, section_title


class TabSplit:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(
            p,
            "Dividir PDF",
            "Selecione o PDF e clique nas páginas que deseja extrair.",
        )

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        top = ttk.Frame(p, style="Card.TFrame")
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(
            top, text="Selecionar PDF", style="Ghost.TButton", command=self._select
        ).pack(side="left")
        self._count_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self._count_var, style="Count.TLabel").pack(
            side="left", padx=14
        )
        self._btn_all = ttk.Button(
            top,
            text="Tudo",
            style="Small.TButton",
            command=lambda: self._grid.select_all(),
        )
        self._btn_none = ttk.Button(
            top,
            text="Limpar",
            style="Small.TButton",
            command=lambda: self._grid.deselect_all(),
        )

        self._btn_save = ttk.Button(
            p,
            text="Salvar páginas selecionadas",
            style="Primary.TButton",
            command=self._save,
            state="disabled",
        )
        self._btn_save.pack(side="bottom", anchor="w", pady=(8, 0))

        self._grid = ThumbnailGrid(p)
        self._grid.on_change(self._on_selection_change)

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._file_var.set(os.path.basename(path))
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status("Carregando páginas…")
        self._btn_all.pack(side="right", padx=(0, 4))
        self._btn_none.pack(side="right")
        n = self._grid.load_pdf(path)
        self.set_status(f"{n} página(s) carregadas — clique para selecionar")

    def _clear(self):
        self._path = None
        self._file_var.set("Nenhum arquivo selecionado")
        self._count_var.set("")
        self._btn_clear.pack_forget()
        self._btn_all.pack_forget()
        self._btn_none.pack_forget()
        self._btn_save.config(state="disabled")
        self._grid.clear()
        self.set_status("Pronto")

    def _on_selection_change(self, selected):
        n = len(selected)
        self._count_var.set(f"{n} página(s) selecionada(s)" if n else "")
        self._btn_save.config(state="normal" if n else "disabled")

    def _save(self):
        selected = self._grid.get_selected()
        if not selected:
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
        )
        if not save:
            return
        try:
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for i in sorted(selected):
                writer.add_page(reader.pages[i])
            with open(save, "wb") as f:
                writer.write(f)
            n = len(selected)
            self.set_status(f"{n} página(s) salvas em {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", f"{n} página(s) salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
