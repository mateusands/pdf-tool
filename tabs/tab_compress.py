import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz

import theme as T
from widgets import DropZone, section_title

LEVELS = {
    "Baixa": {"garbage": 1, "deflate": False},
    "Média": {"garbage": 3, "deflate": True},
    "Alta":  {"garbage": 4, "deflate": True, "clean": True},
}

LEVEL_DESC = {
    "Baixa": "Mais rápido, arquivo maior",
    "Média": "Equilibrado",
    "Alta":  "Mais lento, arquivo menor",
}


def _fmt(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024**2:
        return f"{n / 1024:.1f} KB"
    return f"{n / 1024**2:.1f} MB"


class TabCompress:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "📦  Compactar PDF",
                      "Reduza o tamanho de um arquivo PDF.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF para compactar",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Compression level — pill toggle
        lvl_frame = ctk.CTkFrame(p, fg_color="transparent")
        lvl_frame.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_M))

        ctk.CTkLabel(lvl_frame, text="Nível de compactação:",
                     font=T.FONT_BODY, text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_S))

        self._level_var = tk.StringVar(value="Média")
        pill_row = ctk.CTkFrame(lvl_frame, fg_color="transparent")
        pill_row.pack(anchor="w")

        self._level_btns: dict[str, ctk.CTkButton] = {}
        for level in LEVELS:
            btn = ctk.CTkButton(
                pill_row, text=level, width=100, height=36,
                corner_radius=8, cursor="hand2", font=T.FONT_PILL,
                fg_color=T.PRIMARY if level == "Média" else T.BG_SECONDARY,
                hover_color=T.PRIMARY_HOVER if level == "Média" else T.SURFACE_HOVER,
                text_color="white" if level == "Média" else T.FG_SECONDARY,
                border_width=1, border_color=T.BORDER,
                command=lambda l=level: self._set_level(l),
            )
            btn.pack(side="left", padx=(0, 6))
            self._level_btns[level] = btn

        # Description
        self._desc_var = tk.StringVar(value=LEVEL_DESC["Média"])
        ctk.CTkLabel(lvl_frame, textvariable=self._desc_var,
                     font=T.FONT_SMALL, text_color=T.MUTED).pack(anchor="w", pady=(T.PAD_S, 0))

        # Size info
        self._size_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self._size_var, font=T.FONT_BODY,
                     text_color=T.FG_SECONDARY).pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_M))

        # Compress button
        self._btn = ctk.CTkButton(
            p, text="📦  Compactar e salvar",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._run, state="disabled",
        )
        self._btn.pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    def _set_level(self, level: str):
        self._level_var.set(level)
        self._desc_var.set(LEVEL_DESC[level])
        for l, btn in self._level_btns.items():
            if l == level:
                btn.configure(fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
                              text_color="white")
            else:
                btn.configure(fg_color=T.BG_SECONDARY, hover_color=T.SURFACE_HOVER,
                              text_color=T.FG_SECONDARY)

    def _select(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        size = os.path.getsize(path)
        self._drop.set_file(os.path.basename(path), _fmt(size))
        self._size_var.set(f"Tamanho original: {_fmt(size)}")
        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._size_var.set("")
        self._btn.configure(state="disabled")
        self.set_status("Pronto")

    def _run(self):
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        opts = LEVELS[self._level_var.get()]
        path, orig = self._path, os.path.getsize(self._path)
        self._btn.configure(state="disabled")
        self.set_status("⏳  Compactando…")

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
        self._btn.configure(state="normal")
        ratio = (1 - final / orig) * 100 if orig else 0
        self._size_var.set(
            f"Original: {_fmt(orig)}  →  Final: {_fmt(final)}  ({ratio:.1f}% menor)")
        self.set_status(f"✓  Compactado: {os.path.basename(save)}")
        messagebox.showinfo(
            "Sucesso",
            f"PDF compactado com sucesso!\n"
            f"Original: {_fmt(orig)}\nFinal: {_fmt(final)}\nRedução: {ratio:.1f}%",
        )

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("✗  Erro ao compactar")
        messagebox.showerror("Erro", msg)
