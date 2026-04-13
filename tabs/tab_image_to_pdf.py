import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz

import theme as T
from widgets import section_title


class TabImageToPdf:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list[str] = []
        self._build(parent)

    def _build(self, p):
        section_title(p, "📄  Imagem → PDF",
                      "Combine imagens PNG/JPG em um único arquivo PDF.")

        # Scrollable image list
        wrap = ctk.CTkFrame(p, fg_color=T.BG_SECONDARY, corner_radius=T.RADIUS_SM,
                            border_width=1, border_color=T.BORDER)
        wrap.pack(fill="both", expand=True, padx=T.PAD_L, pady=(0, T.PAD_S))

        self._canvas = tk.Canvas(wrap, bg=T.BG_SECONDARY, highlightthickness=0)
        vsb = ctk.CTkScrollbar(wrap, command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y", padx=(0, 2), pady=2)
        self._canvas.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        self._inner = tk.Frame(self._canvas, bg=T.BG_SECONDARY)
        self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda _: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<MouseWheel>",
                          lambda e: self._canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._rebuild()

        # Controls
        ctrl = ctk.CTkFrame(p, fg_color="transparent")
        ctrl.pack(fill="x", padx=T.PAD_L, pady=T.PAD_S)

        ctk.CTkButton(
            ctrl, text="＋  Adicionar imagens",
            fg_color="transparent", border_width=1, border_color=T.PRIMARY,
            text_color=T.PRIMARY, hover_color=T.SURFACE_HOVER,
            font=T.FONT_BODY, cursor="hand2", corner_radius=8, height=38,
            command=self._add,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            ctrl, text="🗑  Limpar tudo",
            fg_color=T.DANGER, hover_color=T.DANGER_HOVER,
            font=T.FONT_BODY, cursor="hand2", corner_radius=8, height=38,
            command=self._clear_all,
        ).pack(side="left")

        ctk.CTkButton(
            p, text="📄  Gerar PDF",
            fg_color=T.SUCCESS, hover_color=T.SUCCESS_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._generate,
        ).pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    # ── List management ───────────────────────────────────────────────────────

    def _rebuild(self):
        for w in self._inner.winfo_children():
            w.destroy()

        if not self._files:
            tk.Label(self._inner,
                     text="Nenhuma imagem adicionada.\nClique em '＋ Adicionar imagens'.",
                     bg=T.BG_SECONDARY, fg=T.MUTED, font=T.FONT_BODY,
                     justify="center", pady=40).pack()
            return

        for i, path in enumerate(self._files):
            self._add_row(i, path)

    def _add_row(self, i: int, path: str):
        row = tk.Frame(self._inner, bg=T.BG_SECONDARY)
        row.pack(fill="x", padx=8, pady=2)

        tk.Label(row, text=f"{i + 1}.", bg=T.BG_SECONDARY, fg=T.MUTED,
                 font=T.FONT_BODY, width=3, anchor="e").pack(side="left")

        tk.Label(row, text=f"  🖼️ {os.path.basename(path)}",
                 bg=T.BG_SECONDARY, fg=T.FG, font=T.FONT_BODY,
                 anchor="w").pack(side="left", fill="x", expand=True, padx=(6, 0))

        # Per-row X remove button
        btn_x = tk.Button(
            row, text="✕", bg=T.DANGER, fg="white",
            activebackground=T.DANGER_HOVER, activeforeground="white",
            font=(T.FONT_FAMILY, 9, "bold"), relief="flat", bd=0,
            padx=7, pady=3, cursor="hand2",
            command=lambda idx=i: self._remove(idx),
        )
        btn_x.pack(side="right", padx=(4, 0))

        # Move buttons
        nav = tk.Frame(row, bg=T.BG_SECONDARY)
        nav.pack(side="right", padx=(4, 0))

        for text, direction in [("↑", -1), ("↓", 1)]:
            btn = tk.Button(
                nav, text=text, bg=T.SURFACE, fg=T.MUTED,
                activebackground=T.SURFACE_HOVER, activeforeground=T.FG,
                relief="flat", font=T.FONT_BODY, padx=6, cursor="hand2",
                command=lambda idx=i, d=direction: self._move(idx, d),
            )
            btn.pack(side="left", padx=1)

        # Separator
        tk.Frame(self._inner, bg=T.BORDER, height=1).pack(fill="x", padx=8)

    def _add(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if not files:
            return
        self._files.extend(files)
        self._rebuild()
        self.set_status(f"✓  {len(files)} imagem(ns) adicionada(s) — total: {len(self._files)}")

    def _remove(self, idx: int):
        self._files.pop(idx)
        self._rebuild()
        self.set_status(
            f"{len(self._files)} imagem(ns) na lista" if self._files else "Lista vazia")

    def _clear_all(self):
        self._files.clear()
        self._rebuild()
        self.set_status("Lista limpa")

    def _move(self, idx: int, direction: int):
        new_idx = idx + direction
        if 0 <= new_idx < len(self._files):
            self._files[idx], self._files[new_idx] = (
                self._files[new_idx], self._files[idx])
            self._rebuild()

    # ── Generate PDF ──────────────────────────────────────────────────────────

    def _generate(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos uma imagem.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        files = list(self._files)
        self.set_status("⏳  Gerando PDF…")

        def task():
            try:
                doc = fitz.open()
                for img_path in files:
                    img = fitz.open(img_path)
                    pdfbytes = img.convert_to_pdf()
                    img.close()
                    imgpdf = fitz.open("pdf", pdfbytes)
                    doc.insert_pdf(imgpdf)
                    imgpdf.close()
                doc.save(save)
                doc.close()
                self.root.after(0, lambda: self._done(save, len(files)))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, save: str, n: int):
        self.set_status(f"✓  PDF gerado: {os.path.basename(save)}")
        messagebox.showinfo("Sucesso", f"PDF criado com {n} página(s)!")

    def _error(self, msg: str):
        self.set_status("✗  Erro ao gerar PDF")
        messagebox.showerror("Erro", msg)
