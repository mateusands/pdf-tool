import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from .. import theme as T
from ..core import pdf_io, reorder
from ..core.background import executar_em_thread
from ..widgets import DropZone, botao, criar_area_rolavel, estado_vazio

_THUMB_W = T.THUMB_W
_THUMB_H = T.THUMB_H
_COLS    = T.GRID_COLS

MOVIMENTOS = [
    ("start", "Início", "chevrons-left"),
    ("left",  "Esquerda", "arrow-left"),
    ("right", "Direita", "arrow-right"),
    ("end",   "Fim", "chevrons-right"),
]


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
        self._drop = DropZone(p, icon="upload", text="Selecione um PDF para organizar",
                              subtitle="ou clique para escolher no computador",
                              command=self._select_file, on_clear=self._clear)

        info = ctk.CTkFrame(p, fg_color="transparent")
        info.pack(fill="x", padx=T.PAD_XL, pady=(0, T.PAD_S))
        self._info_var = tk.StringVar(value="")
        ctk.CTkLabel(info, textvariable=self._info_var,
                     font=T.FONT_LABEL, text_color=T.ACCENT_TEXT).pack(side="left")

        # Rodapé — empacotado antes da grade, que é elástica
        rodape = ctk.CTkFrame(p, fg_color="transparent")
        rodape.pack(side="bottom", fill="x", padx=T.PAD_XL, pady=(T.PAD_S, T.PAD_XL))

        barra = ctk.CTkFrame(rodape, fg_color="transparent")
        barra.pack(fill="x", pady=(0, T.PAD_M))

        ctk.CTkLabel(barra, text="Mover página", font=T.FONT_LABEL,
                     text_color=T.FG_SECONDARY).pack(side="left", padx=(0, T.PAD_M))

        self._move_btns: dict = {}
        for acao, rotulo, nome_do_icone in MOVIMENTOS:
            btn = botao(barra, rotulo, nome_do_icone=nome_do_icone,
                        variante="secundario", altura=32, largura=112,
                        command=lambda a=acao: self._move(a), state="disabled")
            btn.pack(side="left", padx=(0, T.PAD_XS))
            self._move_btns[acao] = btn

        ctk.CTkLabel(barra, text="ou arraste as miniaturas", font=T.FONT_SMALL,
                     text_color=T.MUTED).pack(side="left", padx=(T.PAD_M, 0))

        self._btn_save = botao(rodape, "Salvar nova ordem", nome_do_icone="save",
                               command=self._save, state="disabled")
        self._btn_save.pack(anchor="w")

        # Grade
        self._canvas, self._inner = criar_area_rolavel(
            p, fill="both", expand=True, padx=T.PAD_XL, pady=(0, T.PAD_M))

        self._show_placeholder()

    # ── Seleção de arquivo ────────────────────────────────────────────────────

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        self.set_status("Carregando páginas…", "ocupado")
        try:
            self._load_pages(path)
        except Exception as e:
            self.set_status("Não foi possível abrir o PDF", "erro")
            messagebox.showerror("Erro ao abrir", str(e))
            return

        self._path = path
        self._drop.set_file(os.path.basename(path), f"{len(self._pages)} página(s)")

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
        # Renderiza antes de mexer no estado: PDF protegido levanta aqui e a
        # grade anterior continua no lugar.
        miniaturas = pdf_io.miniaturas_ppm(path, _THUMB_W, _THUMB_H)

        self._pages = []
        self._selected = None
        for i, ppm in enumerate(miniaturas):
            self._pages.append((i, tk.PhotoImage(data=ppm)))
            if i % 3 == 2:
                self._canvas.update_idletasks()
        n = len(self._pages)
        self._info_var.set(f"{n} página(s) — clique para selecionar, arraste para mover")
        self._btn_save.configure(state="normal")
        self._rebuild_grid()
        self.set_status(f"{n} página(s) carregadas")

    # ── Grade ─────────────────────────────────────────────────────────────────

    def _show_placeholder(self):
        for w in self._inner.winfo_children():
            w.destroy()
        estado_vazio(self._inner, "layout-grid", "Nenhum PDF carregado")

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
                fg=T.ACCENT_TEXT if is_sel else T.MUTED, font=T.FONT_TINY,
            )
            num_lbl.pack(pady=(4, 0))
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
            return T.ACCENT
        return T.BORDER

    def _refresh_visuals(self):
        for idx, (_, border, img_lbl, num_lbl) in self._cells.items():
            bc = self._border_color(idx)
            border.config(bg=bc)
            img_lbl.config(bg=bc)
            is_sel = (idx == self._selected)
            num_lbl.config(fg=T.ACCENT_TEXT if is_sel else T.MUTED)
        self._update_move_btns()

    # ── Arrastar e soltar ─────────────────────────────────────────────────────

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
            insert_at = reorder.indice_apos_arrastar(src, tgt)
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
        tem = self._selected is not None
        n   = len(self._pages)
        sel = self._selected
        for acao in ("start", "left"):
            self._move_btns[acao].configure(
                state="normal" if tem and sel > 0 else "disabled")
        for acao in ("right", "end"):
            self._move_btns[acao].configure(
                state="normal" if tem and sel < n - 1 else "disabled")

    def _move(self, action: str):
        self._pages, self._selected = reorder.mover(self._pages, self._selected, action)
        self._rebuild_grid()

    # ── Salvar ────────────────────────────────────────────────────────────────

    def _save(self):
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        ordem = [orig_idx for orig_idx, _ in self._pages]
        path = self._path
        self._btn_save.configure(state="disabled")
        self.set_status("Salvando nova ordem…", "ocupado")
        executar_em_thread(
            self.root,
            lambda: pdf_io.reorganizar_pdf(path, save, ordem),
            ao_terminar=lambda _: self._done(save),
            ao_falhar=self._error,
        )

    def _done(self, save: str):
        self._btn_save.configure(state="normal")
        self.set_status(f"Salvo: {os.path.basename(save)}", "ok")
        messagebox.showinfo("Sucesso", "PDF salvo com a nova ordem das páginas!")

    def _error(self, msg: str):
        self._btn_save.configure(state="normal")
        self.set_status("Erro ao salvar", "erro")
        messagebox.showerror("Erro", msg)
