import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz

from constants import BORDER, CARD, FG, MUTED
from widgets import section_title

# Colors for the per-row X button
_X_BG     = "#e53935"
_X_HOVER  = "#c62828"


class TabImageToPdf:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list[str] = []
        self._build(parent)

    # ── UI build ──────────────────────────────────────────────────────────────

    def _build(self, p):
        section_title(
            p, "Imagem → PDF", "Combine imagens PNG/JPG em um único arquivo PDF."
        )

        # Scrollable image list
        wrap = tk.Frame(p, bg=BORDER, bd=1, relief="solid")
        wrap.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrap, bg=CARD, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=CARD)
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

        self._rebuild()  # show placeholder

        # Controls
        ctrl = ttk.Frame(p, style="Card.TFrame")
        ctrl.pack(fill="x", pady=10)
        ttk.Button(
            ctrl,
            text="+ Adicionar imagens",
            style="Ghost.TButton",
            command=self._add,
        ).pack(side="left")

        ttk.Button(
            p, text="Gerar PDF", style="Success.TButton", command=self._generate
        ).pack(anchor="w")

    # ── List management ───────────────────────────────────────────────────────

    def _rebuild(self):
        """Redraw the entire image list from self._files."""
        for w in self._inner.winfo_children():
            w.destroy()

        if not self._files:
            tk.Label(
                self._inner,
                text="Nenhuma imagem adicionada.\nClique em '+ Adicionar imagens' para começar.",
                bg=CARD,
                fg=MUTED,
                font=("Segoe UI", 9),
                justify="center",
                pady=40,
            ).pack()
            return

        for i, path in enumerate(self._files):
            self._add_row(i, path)

    def _add_row(self, i: int, path: str):
        row = tk.Frame(self._inner, bg=CARD)
        row.pack(fill="x", padx=6, pady=2)

        # Index number
        tk.Label(
            row,
            text=f"{i + 1}.",
            bg=CARD,
            fg=MUTED,
            font=("Segoe UI", 9),
            width=3,
            anchor="e",
        ).pack(side="left")

        # Filename
        tk.Label(
            row,
            text=os.path.basename(path),
            bg=CARD,
            fg=FG,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(6, 0))

        # Move buttons
        nav = tk.Frame(row, bg=CARD)
        nav.pack(side="right", padx=(4, 0))
        tk.Button(
            nav,
            text="↑",
            command=lambda idx=i: self._move(idx, -1),
            bg=CARD,
            fg=MUTED,
            relief="flat",
            font=("Segoe UI", 9),
            padx=4,
            cursor="hand2",
        ).pack(side="left")
        tk.Button(
            nav,
            text="↓",
            command=lambda idx=i: self._move(idx, 1),
            bg=CARD,
            fg=MUTED,
            relief="flat",
            font=("Segoe UI", 9),
            padx=4,
            cursor="hand2",
        ).pack(side="left")

        # X remove button
        btn_x = tk.Button(
            row,
            text="✕",
            command=lambda idx=i: self._remove(idx),
            bg=_X_BG,
            fg="white",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2",
            activebackground=_X_HOVER,
            activeforeground="white",
        )
        btn_x.pack(side="right", padx=(4, 0))

        # Separator
        tk.Frame(self._inner, bg=BORDER, height=1).pack(fill="x", padx=6)

    def _add(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Imagens", "*.png *.jpg *.jpeg")]
        )
        if not files:
            return
        self._files.extend(files)
        self._rebuild()
        self.set_status(f"{len(files)} imagem(ns) adicionada(s) — total: {len(self._files)}")

    def _remove(self, idx: int):
        self._files.pop(idx)
        self._rebuild()
        self.set_status(
            f"{len(self._files)} imagem(ns) na lista" if self._files else "Lista vazia"
        )

    def _move(self, idx: int, direction: int):
        new_idx = idx + direction
        if 0 <= new_idx < len(self._files):
            self._files[idx], self._files[new_idx] = (
                self._files[new_idx],
                self._files[idx],
            )
            self._rebuild()

    # ── Generate PDF ──────────────────────────────────────────────────────────

    def _generate(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos uma imagem.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
        )
        if not save:
            return
        files = list(self._files)
        self.set_status("Gerando PDF…")

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
        self.set_status(f"PDF gerado: {os.path.basename(save)}")
        messagebox.showinfo("Sucesso", f"PDF criado com {n} página(s)!")

    def _error(self, msg: str):
        self.set_status("Erro ao gerar PDF")
        messagebox.showerror("Erro", msg)
