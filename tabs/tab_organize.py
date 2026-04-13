import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz
from pypdf import PdfReader, PdfWriter

import theme as T
from widgets import DropZone, section_title

_THUMB_W = T.THUMB_W
_THUMB_H = T.THUMB_H
_COLS    = T.GRID_COLS


class TabOrganize:
    def __init__(self, parent, set_status, root):
        self.set_status  = set_status
        self.root        = root
        self._path: str | None                       = None
        self._pages: list[tuple[int, tk.PhotoImage]] = []
        self._selected: int | None                   = None
        self._drag_src: int | None                   = None
        self._drag_tgt: int | None                   = None
        self._cells: dict                            = {}
        self._build(parent)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self, p):
        section_title(p, "⚙️  Organizar Páginas",
                      "Reordene as páginas arrastando ou usando os botões.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF para organizar",
                              subtitle="ou clique para selecionar",
                              command=self._select_file, on_clear=self._clear)

        # Info
        top = ctk.CTkFrame(p, fg_color="transparent")
        top.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_S))
        self._info_var = tk.StringVar(value="")
        ctk.CTkLabel(top, textvariable=self._info_var,
                     font=T.FONT_BODY, text_color=T.MUTED).pack(side="left")

        # Save button — packed BEFORE grid so it stays at bottom
        self._btn_save = ctk.CTkButton(
            p, text="💾  Salvar PDF reorganizado",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._save, state="disabled",
        )
        self._btn_save.pack(side="bottom", anchor="w", padx=T.PAD_L,
                            pady=(T.PAD_S, T.PAD_L))

        # Movement bar — also before grid
        move_bar = ctk.CTkFrame(p, fg_color="transparent")
        move_bar.pack(side="bottom", fill="x", padx=T.PAD_L, pady=(T.PAD_XS, 0))

        ctk.CTkLabel(move_bar, text="Mover:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(side="left", padx=(0, 8))

        self._move_btns: dict[str, ctk.CTkButton] = {}
        for text, action in [
            ("⏮ Início",   "start"),
            ("◀ Esquerda", "left"),
            ("Direita ▶",  "right"),
            ("Fim ⏭",      "end"),
        ]:
            btn = ctk.CTkButton(
                move_bar, text=text, fg_color="transparent",
                border_width=1, border_color=T.BORDER,
                text_color=T.FG_SECONDARY, hover_color=T.SURFACE_HOVER,
                font=T.FONT_BODY, cursor="hand2", corner_radius=6,
                height=32, command=lambda a=action: self._move(a),
                state="disabled",
            )
            btn.pack(side="left", padx=(0, 4))
            self._move_btns[action] = btn

        ctk.CTkLabel(move_bar, text="  |  Arraste para reordenar",
                     font=T.FONT_SMALL, text_color=T.MUTED).pack(side="left", padx=(8, 0))

        # Canvas with thumbnails
        wrap = ctk.CTkFrame(p, fg_color=T.GRID_BG, corner_radius=T.RADIUS_SM,
                            border_width=1, border_color=T.BORDER)
        wrap.pack(fill="both", expand=True, padx=T.PAD_L, pady=(0, T.PAD_S))

        self._canvas = tk.Canvas(wrap, bg=T.GRID_BG, highlightthickness=0)
        vsb = ctk.CTkScrollbar(wrap, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._canvas.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._inner = tk.Frame(self._canvas, bg=T.GRID_BG)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._show_placeholder()

    # ── File selection ────────────────────────────────────────────────────────

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._drop.set_file(os.path.basename(path))
        self.set_status("Carregando páginas…")
        self._load_pages(path)

    def _clear(self):
        self._path = None
        self._pages = []
        self._selected = None
        self._drag_src = None
        self._drag_tgt = None
        self._cells = {}
        self._info_var.set("")
        self._btn_save.configure(state="disabled")
        for btn in self._move_btns.values():
            btn.configure(state="disabled")
        self._show_placeholder()
        self.set_status("Pronto")

    def _load_pages(self, path: str):
        doc = fitz.open(path)
        self._pages = []
        self._selected = None
        for i, page in enumerate(doc):
            mat = fitz.Matrix(_THUMB_W / page.rect.width,
                              _THUMB_H / page.rect.height)
            pix = page.get_pixmap(matrix=mat)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                tmp = f.name
            pix.save(tmp)
            photo = tk.PhotoImage(file=tmp)
            os.unlink(tmp)
            self._pages.append((i, photo))
            if i % 3 == 2:
                self._canvas.update_idletasks()
        doc.close()
        n = len(self._pages)
        self._info_var.set(f"{n} página(s) — clique para selecionar, arraste para mover")
        self._btn_save.configure(state="normal")
        self._rebuild_grid()
        self.set_status(f"✓  {n} página(s) carregadas")

    # ── Grid ──────────────────────────────────────────────────────────────────

    def _show_placeholder(self):
        for w in self._inner.winfo_children():
            w.destroy()
        tk.Label(self._inner, text="Selecione um PDF para organizar as páginas",
                 bg=T.GRID_BG, fg=T.MUTED, font=T.FONT_BODY, pady=60).pack()

    def _rebuild_grid(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._cells = {}

        for i, (_, photo) in enumerate(self._pages):
            row, col = divmod(i, _COLS)
            cell = tk.Frame(self._inner, bg=T.GRID_BG, padx=5, pady=5)
            bc = self._border_color(i)
            border = tk.Frame(cell, bg=bc, padx=2, pady=2)
            img_lbl = tk.Label(border, image=photo, cursor="fleur", bg=bc)
            img_lbl.pack()
            border.pack()

            is_sel = (i == self._selected)
            num_lbl = tk.Label(
                cell, text=str(i + 1), bg=T.GRID_BG,
                fg=T.PRIMARY if is_sel else T.MUTED,
                font=(T.FONT_FAMILY, 8, "bold") if is_sel else T.FONT_TINY,
            )
            num_lbl.pack(pady=(3, 0))
            cell.grid(row=row, column=col, padx=4, pady=4)
            self._cells[i] = (cell, border, img_lbl, num_lbl)

            for w in (cell, border, img_lbl, num_lbl):
                w.bind("<ButtonPress-1>",   lambda _, idx=i: self._on_press(idx))
                w.bind("<B1-Motion>",       lambda e, idx=i: self._on_motion(e))
                w.bind("<ButtonRelease-1>", lambda _: self._on_release())

        self._canvas.yview_moveto(0)
        self._update_move_btns()

    def _border_color(self, idx: int) -> str:
        if idx == self._drag_src:
            return T.COLOR_DRAG_SRC
        if idx == self._drag_tgt:
            return T.COLOR_DRAG_TGT
        if idx == self._selected:
            return T.PRIMARY
        return T.BORDER

    def _refresh_visuals(self):
        for idx, (_, border, img_lbl, num_lbl) in self._cells.items():
            bc = self._border_color(idx)
            border.config(bg=bc)
            img_lbl.config(bg=bc)
            is_sel = (idx == self._selected)
            num_lbl.config(
                fg=T.PRIMARY if is_sel else T.MUTED,
                font=(T.FONT_FAMILY, 8, "bold") if is_sel else T.FONT_TINY,
            )
        self._update_move_btns()

    # ── Drag & drop ───────────────────────────────────────────────────────────

    def _on_press(self, idx: int):
        self._selected = idx
        self._drag_src = idx
        self._drag_tgt = None
        self._refresh_visuals()

    def _on_motion(self, event):
        if self._drag_src is None:
            return
        ax = event.widget.winfo_rootx() + event.x
        ay = event.widget.winfo_rooty() + event.y
        tgt = self._pos_to_idx(ax, ay)
        if tgt is not None and tgt != self._drag_src and tgt != self._drag_tgt:
            self._drag_tgt = tgt
            self._refresh_visuals()

    def _on_release(self):
        src = self._drag_src
        tgt = self._drag_tgt
        self._drag_src = None
        self._drag_tgt = None

        if src is not None and tgt is not None and src != tgt:
            page = self._pages.pop(src)
            insert_at = tgt - 1 if tgt > src else tgt
            self._pages.insert(insert_at, page)
            self._selected = insert_at
            self._rebuild_grid()
        else:
            self._refresh_visuals()

    def _pos_to_idx(self, ax: int, ay: int) -> int | None:
        for idx, (cell, *_) in self._cells.items():
            cx = cell.winfo_rootx()
            cy = cell.winfo_rooty()
            if cx <= ax <= cx + cell.winfo_width() and \
               cy <= ay <= cy + cell.winfo_height():
                return idx
        return None

    def _update_move_btns(self):
        has = self._selected is not None
        n   = len(self._pages)
        sel = self._selected
        self._move_btns["start"].configure(
            state="normal" if has and sel > 0 else "disabled")
        self._move_btns["left"].configure(
            state="normal" if has and sel > 0 else "disabled")
        self._move_btns["right"].configure(
            state="normal" if has and sel < n - 1 else "disabled")
        self._move_btns["end"].configure(
            state="normal" if has and sel < n - 1 else "disabled")

    # ── Move buttons ──────────────────────────────────────────────────────────

    def _move(self, action: str):
        i = self._selected
        if i is None:
            return
        n = len(self._pages)

        if   action == "left"  and i > 0:
            self._pages[i], self._pages[i - 1] = self._pages[i - 1], self._pages[i]
            self._selected = i - 1
        elif action == "right" and i < n - 1:
            self._pages[i], self._pages[i + 1] = self._pages[i + 1], self._pages[i]
            self._selected = i + 1
        elif action == "start" and i > 0:
            self._pages.insert(0, self._pages.pop(i))
            self._selected = 0
        elif action == "end"   and i < n - 1:
            self._pages.append(self._pages.pop(i))
            self._selected = n - 1

        self._rebuild_grid()

    # ── Save ──────────────────────────────────────────────────────────────────

    def _save(self):
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        try:
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for orig_idx, _ in self._pages:
                writer.add_page(reader.pages[orig_idx])
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"✓  Salvo: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", "PDF salvo com a nova ordem das páginas!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("✗  Erro ao salvar")
