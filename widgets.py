"""Reusable dark-themed widgets for the PDF Tool."""

import os
import tempfile
import tkinter as tk

import customtkinter as ctk
import fitz

import theme as T


# ── Section header ────────────────────────────────────────────────────────────

def section_title(parent, title, subtitle=""):
    """Dark-themed section header with optional subtitle and separator."""
    ctk.CTkLabel(
        parent, text=title, font=T.FONT_TITLE, text_color=T.FG, anchor="w",
    ).pack(fill="x", padx=T.PAD_L, pady=(T.PAD_L, 0))
    if subtitle:
        ctk.CTkLabel(
            parent, text=subtitle, font=T.FONT_SMALL, text_color=T.MUTED, anchor="w",
        ).pack(fill="x", padx=T.PAD_L, pady=(2, 0))
    ctk.CTkFrame(parent, fg_color=T.BORDER, height=1).pack(
        fill="x", padx=T.PAD_L, pady=(T.PAD_M, T.PAD_L),
    )


# ── Drop Zone ────────────────────────────────────────────────────────────────

class DropZone(tk.Canvas):
    """Clickable file-selection area with dashed border, hover, and clear button."""

    def __init__(self, parent, icon="📁", text="Arraste o arquivo aqui",
                 subtitle="ou clique para selecionar", command=None,
                 on_clear=None, height=130):
        super().__init__(parent, bg=T.SURFACE, highlightthickness=0,
                         cursor="hand2", height=height)
        self._cmd = command
        self._on_clear = on_clear
        self._hover = False
        self._hover_x = False  # hovering the X button
        self._has_file = False
        self._file_label = ""
        self._icon = icon
        self._text = text
        self._subtitle = subtitle

        self.bind("<Configure>", self._draw)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", lambda _: self._set_hover(True))
        self.bind("<Leave>", lambda _: self._set_hover(False))
        self.bind("<Motion>", self._on_motion)
        self.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_M))

    # -- state helpers --

    def _set_hover(self, val):
        self._hover = val
        if not val:
            self._hover_x = False
        self._draw()

    def set_file(self, name, size_str=""):
        self._has_file = True
        self._file_label = f"📄  {name}" + (f"  •  {size_str}" if size_str else "")
        self._draw()

    def clear_file(self):
        self._has_file = False
        self._file_label = ""
        self._draw()

    # -- X button hit detection --

    def _x_bbox(self):
        """Return (x1, y1, x2, y2) of the X close button area."""
        w = self.winfo_width()
        # top-right corner, padded from edge
        x2 = w - 16
        x1 = x2 - 28
        y1 = 14
        y2 = y1 + 28
        return x1, y1, x2, y2

    def _in_x(self, mx, my):
        if not self._has_file:
            return False
        x1, y1, x2, y2 = self._x_bbox()
        return x1 <= mx <= x2 and y1 <= my <= y2

    def _on_motion(self, event):
        was = self._hover_x
        self._hover_x = self._in_x(event.x, event.y)
        if was != self._hover_x:
            self._draw()

    def _on_click(self, event):
        if self._has_file and self._in_x(event.x, event.y):
            self.clear_file()
            if self._on_clear:
                self._on_clear()
        elif self._cmd:
            self._cmd()

    # -- drawing --

    def _draw(self, _event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 1:
            return

        pad = 6
        border_color = T.DROPZONE_HOVER_BD if self._hover else T.DROPZONE_BORDER
        fill_color = T.DROPZONE_HOVER_BG if self._hover else T.SURFACE

        self.create_rectangle(pad, pad, w - pad, h - pad,
                              fill=fill_color, outline="")
        self.create_rectangle(pad, pad, w - pad, h - pad,
                              outline=border_color, dash=(8, 4), width=2)

        cx, cy = w / 2, h / 2
        if self._has_file:
            # File info centered
            self.create_text(cx, cy - 6, text=self._file_label,
                             fill=T.FG, font=T.FONT_BODY)
            self.create_text(cx, cy + 18, text="Clique para trocar arquivo",
                             fill=T.MUTED, font=T.FONT_SMALL)

            # X clear button — top right
            x1, y1, x2, y2 = self._x_bbox()
            xc = (x1 + x2) / 2
            yc = (y1 + y2) / 2
            bg = T.DANGER if self._hover_x else T.BORDER_LIGHT
            self.create_oval(x1, y1, x2, y2, fill=bg, outline="")
            self.create_text(xc, yc, text="✕", fill="white",
                             font=(T.FONT_FAMILY, 10, "bold"))
        else:
            icon_color = T.PRIMARY if self._hover else T.MUTED
            self.create_text(cx, cy - 20, text=self._icon,
                             fill=icon_color, font=T.FONT_ICON_LG)
            self.create_text(cx, cy + 14, text=self._text,
                             fill=T.FG, font=T.FONT_BODY)
            self.create_text(cx, cy + 36, text=self._subtitle,
                             fill=T.MUTED, font=T.FONT_SMALL)


# ── Thumbnail Grid ────────────────────────────────────────────────────────────

class ThumbnailGrid:
    """Scrollable grid of PDF page thumbnails with click-to-select."""

    def __init__(self, parent, cols=T.GRID_COLS, thumb_w=T.THUMB_W, thumb_h=T.THUMB_H):
        self.cols = cols
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h
        self._images: list = []
        self._cells: dict = {}
        self._selected: set = set()
        self._on_change_cb = None

        wrap = ctk.CTkFrame(parent, fg_color=T.GRID_BG, corner_radius=T.RADIUS_SM,
                            border_width=1, border_color=T.BORDER)
        wrap.pack(fill="both", expand=True, padx=T.PAD_L, pady=(0, T.PAD_S))

        self._canvas = tk.Canvas(wrap, bg=T.GRID_BG, highlightthickness=0)
        vsb = ctk.CTkScrollbar(wrap, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._canvas.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._inner = tk.Frame(self._canvas, bg=T.GRID_BG)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind(
            "<Configure>",
            lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.bind(
            "<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"),
        )

        self._placeholder = tk.Label(
            self._inner, text="Selecione um PDF para visualizar as páginas",
            bg=T.GRID_BG, fg=T.MUTED, font=T.FONT_BODY, pady=60,
        )
        self._placeholder.pack()

    def on_change(self, callback):
        self._on_change_cb = callback

    def clear(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()
        self._cells.clear()
        self._selected.clear()
        tk.Label(
            self._inner, text="Selecione um PDF para visualizar as páginas",
            bg=T.GRID_BG, fg=T.MUTED, font=T.FONT_BODY, pady=60,
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
            mat = fitz.Matrix(self.thumb_w / page.rect.width,
                              self.thumb_h / page.rect.height)
            pix = page.get_pixmap(matrix=mat)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                tmp = f.name
            pix.save(tmp)
            photo = tk.PhotoImage(file=tmp)
            os.unlink(tmp)
            self._images.append(photo)

            row, col = divmod(i, self.cols)
            cell = tk.Frame(self._inner, bg=T.GRID_BG, padx=6, pady=6)
            border = tk.Frame(cell, bg=T.BORDER, padx=2, pady=2)
            img_lbl = tk.Label(border, image=photo, cursor="hand2", bg=T.BORDER)
            img_lbl.pack()
            border.pack()
            num_lbl = tk.Label(
                cell, text=str(i + 1), bg=T.GRID_BG, fg=T.MUTED, font=T.FONT_TINY,
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
                border.config(bg=T.PRIMARY)
                num_lbl.config(fg=T.PRIMARY, font=(T.FONT_FAMILY, 9, "bold"))
            else:
                border.config(bg=T.BORDER)
                num_lbl.config(fg=T.MUTED, font=T.FONT_TINY)
        if self._on_change_cb:
            self._on_change_cb(self._selected)
