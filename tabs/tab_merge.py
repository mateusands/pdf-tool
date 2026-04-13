import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pypdf import PdfWriter

import theme as T
from widgets import section_title


class TabMerge:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list[str] = []
        self._build(parent)

    def _build(self, p):
        section_title(p, "🔗  Juntar PDFs",
                      "Combine múltiplos PDFs em um único arquivo.")

        # Scrollable file list
        list_frame = ctk.CTkFrame(p, fg_color=T.BG_SECONDARY, corner_radius=T.RADIUS_SM,
                                  border_width=1, border_color=T.BORDER)
        list_frame.pack(fill="both", expand=True, padx=T.PAD_L, pady=(0, T.PAD_S))

        self._canvas = tk.Canvas(list_frame, bg=T.BG_SECONDARY, highlightthickness=0)
        vsb = ctk.CTkScrollbar(list_frame, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._canvas.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._inner = tk.Frame(self._canvas, bg=T.BG_SECONDARY)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._rebuild_list()

        # Controls
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.pack(fill="x", padx=T.PAD_L, pady=T.PAD_S)

        ctk.CTkButton(
            ctrl, text="＋  Adicionar PDFs", fg_color="transparent",
            border_width=1, border_color=T.PRIMARY, text_color=T.PRIMARY,
            hover_color=T.SURFACE_HOVER, font=T.FONT_BODY, cursor="hand2",
            corner_radius=8, height=38, command=self._add,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ctrl, text="🗑  Limpar tudo", fg_color=T.DANGER, hover_color=T.DANGER_HOVER,
            font=T.FONT_BODY, cursor="hand2", corner_radius=8, height=38,
            command=self._clear_all,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ctrl, text="↑ Subir", fg_color="transparent",
            border_width=1, border_color=T.BORDER, text_color=T.FG_SECONDARY,
            hover_color=T.SURFACE_HOVER, font=T.FONT_BODY, cursor="hand2",
            corner_radius=8, height=38, command=self._move_up,
        ).pack(side="left", padx=(0, 4))

        ctk.CTkButton(
            ctrl, text="↓ Descer", fg_color="transparent",
            border_width=1, border_color=T.BORDER, text_color=T.FG_SECONDARY,
            hover_color=T.SURFACE_HOVER, font=T.FONT_BODY, cursor="hand2",
            corner_radius=8, height=38, command=self._move_down,
        ).pack(side="left")

        ctk.CTkButton(
            p, text="🔗  Juntar e salvar PDF",
            fg_color=T.SUCCESS, hover_color=T.SUCCESS_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._merge,
        ).pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    # ── List management ───────────────────────────────────────────────────────

    def _rebuild_list(self):
        for w in self._inner.winfo_children():
            w.destroy()

        if not self._files:
            tk.Label(self._inner,
                     text="Nenhum arquivo adicionado.\nClique em '＋ Adicionar PDFs' para começar.",
                     bg=T.BG_SECONDARY, fg=T.MUTED, font=T.FONT_BODY,
                     justify="center", pady=40).pack()
            return

        for i, path in enumerate(self._files):
            row = tk.Frame(self._inner, bg=T.BG_SECONDARY)
            row.pack(fill="x", padx=8, pady=2)

            tk.Label(row, text=f"{i + 1}.", bg=T.BG_SECONDARY, fg=T.MUTED,
                     font=T.FONT_BODY, width=3, anchor="e").pack(side="left")

            tk.Label(row, text=f"  📄 {os.path.basename(path)}",
                     bg=T.BG_SECONDARY, fg=T.FG, font=T.FONT_BODY,
                     anchor="w").pack(side="left", fill="x", expand=True)

            # Per-row X remove button
            btn_x = tk.Button(
                row, text="✕", bg=T.DANGER, fg="white",
                activebackground=T.DANGER_HOVER, activeforeground="white",
                font=(T.FONT_FAMILY, 9, "bold"), relief="flat", bd=0,
                padx=7, pady=3, cursor="hand2",
                command=lambda idx=i: self._remove(idx),
            )
            btn_x.pack(side="right", padx=(4, 0))

            # Separator
            tk.Frame(self._inner, bg=T.BORDER, height=1).pack(fill="x", padx=8)

        # Bind row selection
        self._selected_idx = None
        for i, row in enumerate(self._inner.winfo_children()):
            if isinstance(row, tk.Frame) and row.cget("bg") == T.BG_SECONDARY:
                for w in [row] + list(row.winfo_children()):
                    if not isinstance(w, tk.Button):
                        w.bind("<Button-1>", lambda _, idx=i // 2: self._select_item(idx))

    def _select_item(self, idx):
        if idx >= len(self._files):
            return
        self._selected_idx = idx
        for i, child in enumerate(self._inner.winfo_children()):
            if isinstance(child, tk.Frame) and child.cget("height") != 1:
                bg = T.SURFACE_HOVER if i // 2 == idx else T.BG_SECONDARY
                child.config(bg=bg)
                for w in child.winfo_children():
                    if not isinstance(w, tk.Button):
                        w.config(bg=bg)

    def _add(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if not files:
            return
        self._files.extend(files)
        self._rebuild_list()
        self.set_status(f"✓  {len(files)} arquivo(s) adicionados — total: {len(self._files)}")

    def _remove(self, idx: int):
        if 0 <= idx < len(self._files):
            self._files.pop(idx)
            self._selected_idx = None
            self._rebuild_list()
            self.set_status(
                f"{len(self._files)} arquivo(s) na lista" if self._files else "Lista vazia")

    def _clear_all(self):
        self._files.clear()
        self._selected_idx = None
        self._rebuild_list()
        self.set_status("Lista limpa")

    def _move_up(self):
        i = getattr(self, "_selected_idx", None)
        if i is not None and i > 0:
            self._files[i], self._files[i - 1] = self._files[i - 1], self._files[i]
            self._selected_idx = i - 1
            self._rebuild_list()

    def _move_down(self):
        i = getattr(self, "_selected_idx", None)
        if i is not None and i < len(self._files) - 1:
            self._files[i], self._files[i + 1] = self._files[i + 1], self._files[i]
            self._selected_idx = i + 1
            self._rebuild_list()

    def _merge(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos um arquivo PDF.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        try:
            self.set_status("Juntando PDFs…")
            writer = PdfWriter()
            for f in self._files:
                writer.append(f)
            writer.write(save)
            writer.close()
            self.set_status(f"✓  Salvo: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", "PDFs juntados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.set_status("✗  Erro ao juntar PDFs")
