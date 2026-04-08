import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from docx2pdf import convert as docx2pdf
from pdf2docx import Converter

from constants import CARD, SUCCESS
from widgets import file_row, section_title


class TabConvert:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._mode = None
        self._build(parent)

    def _build(self, p):
        section_title(
            p, "Converter PDF ↔ Word", "Converta entre PDF e DOCX automaticamente."
        )

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._action_var = tk.StringVar(value="")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        ttk.Button(
            p,
            text="Selecionar arquivo (PDF ou DOCX)",
            style="Ghost.TButton",
            command=self._select,
        ).pack(anchor="w")

        badge = ttk.Frame(p, style="Card.TFrame")
        badge.pack(fill="x", pady=16)
        ttk.Label(badge, text="Ação detectada: ", style="Muted.TLabel").pack(
            side="left"
        )
        tk.Label(
            badge,
            textvariable=self._action_var,
            fg=SUCCESS,
            bg=CARD,
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")

        self._btn = ttk.Button(
            p,
            text="Converter agora",
            style="Primary.TButton",
            command=self._convert,
            state="disabled",
        )
        self._btn.pack(anchor="w")

    def _select(self):
        path = filedialog.askopenfilename(
            filetypes=[("Documentos", "*.pdf *.docx")]
        )
        if not path:
            return
        self._path = path
        self._file_var.set(os.path.basename(path))
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            self._mode = "pdf2word"
            self._action_var.set("PDF → Word (.docx)")
        else:
            self._mode = "word2pdf"
            self._action_var.set("Word (.docx) → PDF")
        self._btn.config(state="normal")
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._mode = None
        self._file_var.set("Nenhum arquivo selecionado")
        self._action_var.set("")
        self._btn.config(state="disabled")
        self._btn_clear.pack_forget()
        self.set_status("Pronto")

    def _convert(self):
        if self._mode == "pdf2word":
            save = filedialog.asksaveasfilename(
                defaultextension=".docx", filetypes=[("Word", "*.docx")]
            )
        else:
            save = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
            )
        if not save:
            return

        self._btn.config(state="disabled")
        mode, path = self._mode, self._path
        self.set_status("Convertendo…")

        def task():
            try:
                if mode == "pdf2word":
                    cv = Converter(path)
                    cv.convert(save, start=0, end=None)
                    cv.close()
                else:
                    docx2pdf(path, save)
                self.root.after(0, lambda: self._done(save))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, save):
        self._btn.config(state="normal")
        self.set_status(f"Salvo: {os.path.basename(save)}")
        messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")

    def _error(self, msg):
        self._btn.config(state="normal")
        self.set_status("Erro na conversão")
        messagebox.showerror("Erro", msg)
