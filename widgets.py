import os
import tempfile
import tkinter as tk
from tkinter import ttk

import fitz

from constants import BORDER, CARD, COLS, MUTED, PRIMARY, THUMB_H, THUMB_W


def section_title(parent, title, subtitle=""):
    ttk.Label(parent, text=title, style="Title.TLabel").pack(anchor="w")
    if subtitle:
        ttk.Label(parent, text=subtitle, style="Sub.TLabel").pack(
            anchor="w", pady=(2, 0)
        )
    ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=12)


def file_row(parent, label_var):
    """File indicator row. Returns the hidden X button — caller shows/hides it."""
    row = ttk.Frame(parent, style="Card.TFrame")
    row.pack(fill="x", pady=(0, 10))
    tk.Frame(row, bg=PRIMARY, width=4).pack(side="left", fill="y")
    ttk.Label(row, textvariable=label_var, style="File.TLabel", padding=[10, 5]).pack(
        side="left"
    )
    btn_clear = tk.Button(
        row,
        text="✕",
        bg="#e53935",
        fg="white",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        bd=0,
        padx=8,
        pady=3,
        cursor="hand2",
        activebackground="#c62828",
        activeforeground="white",
    )
    # Not packed yet — shown by the tab after a file is selected
    return btn_clear


class ThumbnailGrid:
    """Scrollable grid of PDF page thumbnails with click-to-select."""

    def __init__(self, parent, cols=COLS, thumb_w=THUMB_W, thumb_h=THUMB_H):
        self.cols = cols
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h
        self._images: list = []
        self._cells: dict = {}
        self._selected: set = set()
        self._on_change_cb = None

        wrap = tk.Frame(parent, bg=BORDER, bd=1, relief="solid")
        wrap.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrap, bg="#F0F2F5", highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg="#F0F2F5")
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind(
            "<Configure>",
            lambda _: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")
            ),
        )
        self._canvas.bind(
            "<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"),
        )

        self._placeholder = tk.Label(
            self._inner,
            text="Selecione um PDF para visualizar as páginas",
            bg="#F0F2F5",
            fg=MUTED,
            font=("Segoe UI", 10),
            pady=60,
        )
        self._placeholder.pack()

    def on_change(self, callback):
        self._on_change_cb = callback

    def clear(self):
        """Reset grid to empty state with placeholder."""
        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()
        self._cells.clear()
        self._selected.clear()
        tk.Label(
            self._inner,
            text="Selecione um PDF para visualizar as páginas",
            bg="#F0F2F5",
            fg=MUTED,
            font=("Segoe UI", 10),
            pady=60,
        ).pack()
        self._canvas.configure(scrollregion=(0, 0, 0, 0))
        if self._on_change_cb:
            self._on_change_cb(set())

    def load_pdf(self, path: str) -> int:
        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()
        self._cells.clear()
        self._selected.clear()

        doc = fitz.open(path)
        for i, page in enumerate(doc):
            mat = fitz.Matrix(
                self.thumb_w / page.rect.width, self.thumb_h / page.rect.height
            )
            pix = page.get_pixmap(matrix=mat)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                tmp = f.name
            pix.save(tmp)
            photo = tk.PhotoImage(file=tmp)
            os.unlink(tmp)
            self._images.append(photo)

            row, col = divmod(i, self.cols)
            cell = tk.Frame(self._inner, bg="#F0F2F5", padx=6, pady=6)
            border = tk.Frame(cell, bg=BORDER, padx=2, pady=2)
            img_lbl = tk.Label(border, image=photo, cursor="hand2", bg=BORDER)
            img_lbl.pack()
            border.pack()
            num_lbl = tk.Label(
                cell, text=str(i + 1), bg="#F0F2F5", fg=MUTED, font=("Segoe UI", 8)
            )
            num_lbl.pack(pady=(3, 0))
            cell.grid(row=row, column=col, padx=4, pady=4)
            self._cells[i] = (cell, border, num_lbl)

            for w in (cell, border, img_lbl, num_lbl):
                w.bind("<Button-1>", lambda _, idx=i: self.toggle(idx))

            if i % 3 == 2:
                self._canvas.update_idletasks()

        doc.close()
        self._canvas.yview_moveto(0)
        return len(self._cells)

    def toggle(self, idx: int):
        if idx in self._selected:
            self._selected.discard(idx)
        else:
            self._selected.add(idx)
        self._refresh()

    def select_all(self):
        self._selected = set(self._cells.keys())
        self._refresh()

    def deselect_all(self):
        self._selected.clear()
        self._refresh()

    def get_selected(self) -> set:
        return set(self._selected)

    def _refresh(self):
        for idx, (_, border, num_lbl) in self._cells.items():
            if idx in self._selected:
                border.config(bg=PRIMARY)
                num_lbl.config(fg=PRIMARY, font=("Segoe UI", 8, "bold"))
            else:
                border.config(bg=BORDER)
                num_lbl.config(fg=MUTED, font=("Segoe UI", 8))
        if self._on_change_cb:
            self._on_change_cb(self._selected)
