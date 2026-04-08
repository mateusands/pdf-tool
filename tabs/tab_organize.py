import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
from pypdf import PdfReader, PdfWriter

from constants import BORDER, MUTED, PRIMARY
from widgets import file_row, section_title

_THUMB_W = 88
_THUMB_H = 124
_COLS    = 5

_COLOR_DRAG_SRC = "#F59E0B"   # âmbar  — página sendo arrastada
_COLOR_DRAG_TGT = "#10B981"   # verde  — destino do drop


class TabOrganize:
    def __init__(self, parent, set_status, root):
        self.set_status  = set_status
        self.root        = root
        self._path: str | None                        = None
        self._pages: list[tuple[int, tk.PhotoImage]]  = []
        self._selected: int | None                    = None
        self._drag_src: int | None                    = None
        self._drag_tgt: int | None                    = None
        self._cells: dict                             = {}
        self._build(parent)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build(self, p):
        section_title(p, "Organizar Páginas",
                      "Reordene as páginas de um PDF arrastando ou usando os botões.")

        self._file_var = tk.StringVar(value="Nenhum arquivo selecionado")
        self._btn_clear = file_row(p, self._file_var)
        self._btn_clear.config(command=self._clear)

        # Barra superior: botão selecionar + contagem
        top = ttk.Frame(p, style="Card.TFrame")
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="Selecionar PDF", style="Ghost.TButton",
                   command=self._select_file).pack(side="left")
        self._info_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self._info_var,
                  style="Muted.TLabel").pack(side="left", padx=14)

        # Botão salvar — empacotado ANTES do canvas para não ser empurrado
        self._btn_save = ttk.Button(
            p, text="Salvar PDF reorganizado",
            style="Primary.TButton",
            command=self._save,
            state="disabled",
        )
        self._btn_save.pack(side="bottom", anchor="w", pady=(8, 0))

        # Barra de movimentação — também antes do canvas
        move_bar = ttk.Frame(p, style="Card.TFrame")
        move_bar.pack(side="bottom", fill="x", pady=(4, 0))

        ttk.Label(move_bar, text="Mover selecionada:",
                  style="Sub.TLabel").pack(side="left", padx=(0, 8))

        self._move_btns: dict[str, ttk.Button] = {}
        for text, action, tip in [
            ("⏮ Início",    "start", ""),
            ("◀ Esquerda",  "left",  ""),
            ("Direita ▶",   "right", ""),
            ("Fim ⏭",       "end",   ""),
        ]:
            btn = ttk.Button(move_bar, text=text, style="Ghost.TButton",
                             command=lambda a=action: self._move(a))
            btn.pack(side="left", padx=(0, 4))
            btn.config(state="disabled")
            self._move_btns[action] = btn

        ttk.Label(move_bar, text="  |  Arraste para reordenar",
                  style="Muted.TLabel").pack(side="left", padx=(8, 0))

        # Canvas com thumbnails
        wrap = tk.Frame(p, bg=BORDER, bd=1, relief="solid")
        wrap.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrap, bg="#F0F2F5", highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg="#F0F2F5")
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>", lambda _: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(
                              -1 * (e.delta // 120), "units"))

        self._show_placeholder()

    # ── Seleção de arquivo ────────────────────────────────────────────────────

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self._path = path
        self._file_var.set(os.path.basename(path))
        self._btn_clear.pack(side="right", padx=(4, 6))
        self.set_status("Carregando páginas…")
        self._load_pages(path)

    def _clear(self):
        self._path     = None
        self._pages    = []
        self._selected = None
        self._drag_src = None
        self._drag_tgt = None
        self._cells    = {}
        self._file_var.set("Nenhum arquivo selecionado")
        self._info_var.set("")
        self._btn_clear.pack_forget()
        self._btn_save.config(state="disabled")
        for btn in self._move_btns.values():
            btn.config(state="disabled")
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
        self._info_var.set(f"{n} página(s) — clique para selecionar")
        self._btn_save.config(state="normal")
        self._rebuild_grid()
        self.set_status(f"{n} página(s) carregadas")

    # ── Grid ──────────────────────────────────────────────────────────────────

    def _show_placeholder(self):
        for w in self._inner.winfo_children():
            w.destroy()
        tk.Label(self._inner,
                 text="Selecione um PDF para organizar as páginas",
                 bg="#F0F2F5", fg=MUTED,
                 font=("Segoe UI", 10), pady=60).pack()

    def _rebuild_grid(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._cells = {}

        for i, (_, photo) in enumerate(self._pages):
            row, col = divmod(i, _COLS)

            cell    = tk.Frame(self._inner, bg="#F0F2F5", padx=5, pady=5)
            bc      = self._border_color(i)
            border  = tk.Frame(cell, bg=bc, padx=2, pady=2)
            img_lbl = tk.Label(border, image=photo, cursor="fleur", bg=bc)
            img_lbl.pack()
            border.pack()

            is_sel  = (i == self._selected)
            num_lbl = tk.Label(cell, text=str(i + 1), bg="#F0F2F5",
                               fg=PRIMARY if is_sel else MUTED,
                               font=("Segoe UI", 8, "bold") if is_sel
                               else ("Segoe UI", 8))
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
        if idx == self._drag_src:  return _COLOR_DRAG_SRC
        if idx == self._drag_tgt:  return _COLOR_DRAG_TGT
        if idx == self._selected:  return PRIMARY
        return BORDER

    def _refresh_visuals(self):
        for idx, (_, border, img_lbl, num_lbl) in self._cells.items():
            bc = self._border_color(idx)
            border.config(bg=bc)
            img_lbl.config(bg=bc)
            is_sel = (idx == self._selected)
            num_lbl.config(
                fg=PRIMARY if is_sel else MUTED,
                font=("Segoe UI", 8, "bold") if is_sel else ("Segoe UI", 8),
            )
        self._update_move_btns()

    # ── Interação: clique e drag & drop ───────────────────────────────────────

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
        """Retorna o índice da célula que contém as coordenadas absolutas (ax, ay)."""
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
        self._move_btns["start"].config(
            state="normal" if has and sel > 0      else "disabled")
        self._move_btns["left"].config(
            state="normal" if has and sel > 0      else "disabled")
        self._move_btns["right"].config(
            state="normal" if has and sel < n - 1  else "disabled")
        self._move_btns["end"].config(
            state="normal" if has and sel < n - 1  else "disabled")

    # ── Botões de mover ───────────────────────────────────────────────────────

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

    # ── Salvar ────────────────────────────────────────────────────────────────

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
            self.set_status(f"Salvo: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", "PDF salvo com a nova ordem das páginas!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("Erro ao salvar")
