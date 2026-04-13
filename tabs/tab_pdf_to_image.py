import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import fitz

import theme as T
from widgets import DropZone, section_title


class TabPdfToImage:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "🖼️  PDF → Imagem",
                      "Converta cada página do PDF em uma imagem.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Format toggle
        fmt_frame = ctk.CTkFrame(p, fg_color="transparent")
        fmt_frame.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_M))

        ctk.CTkLabel(fmt_frame, text="Formato de saída:",
                     font=T.FONT_BODY, text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_S))

        self._fmt_var = tk.StringVar(value="png")
        fmt_row = ctk.CTkFrame(fmt_frame, fg_color="transparent")
        fmt_row.pack(anchor="w")

        self._fmt_btns: dict[str, ctk.CTkButton] = {}
        for fmt, label in [("png", "PNG"), ("jpg", "JPG")]:
            btn = ctk.CTkButton(
                fmt_row, text=label, width=90, height=36,
                corner_radius=8, cursor="hand2", font=T.FONT_PILL,
                fg_color=T.PRIMARY if fmt == "png" else T.BG_SECONDARY,
                hover_color=T.PRIMARY_HOVER if fmt == "png" else T.SURFACE_HOVER,
                text_color="white" if fmt == "png" else T.FG_SECONDARY,
                border_width=1, border_color=T.BORDER,
                command=lambda f=fmt: self._set_fmt(f),
            )
            btn.pack(side="left", padx=(0, 6))
            self._fmt_btns[fmt] = btn

        # DPI toggle
        dpi_frame = ctk.CTkFrame(p, fg_color="transparent")
        dpi_frame.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_M))

        ctk.CTkLabel(dpi_frame, text="Resolução (DPI):",
                     font=T.FONT_BODY, text_color=T.MUTED).pack(anchor="w", pady=(0, T.PAD_S))

        self._dpi_var = tk.StringVar(value="150")
        dpi_row = ctk.CTkFrame(dpi_frame, fg_color="transparent")
        dpi_row.pack(anchor="w")

        self._dpi_btns: dict[str, ctk.CTkButton] = {}
        for dpi, label in [("72", "72 — Baixa"), ("150", "150 — Média"), ("300", "300 — Alta")]:
            btn = ctk.CTkButton(
                dpi_row, text=label, width=120, height=36,
                corner_radius=8, cursor="hand2", font=T.FONT_PILL,
                fg_color=T.PRIMARY if dpi == "150" else T.BG_SECONDARY,
                hover_color=T.PRIMARY_HOVER if dpi == "150" else T.SURFACE_HOVER,
                text_color="white" if dpi == "150" else T.FG_SECONDARY,
                border_width=1, border_color=T.BORDER,
                command=lambda d=dpi: self._set_dpi(d),
            )
            btn.pack(side="left", padx=(0, 6))
            self._dpi_btns[dpi] = btn

        # Progress bar
        self._progress = ctk.CTkProgressBar(p, fg_color=T.BORDER,
                                            progress_color=T.PRIMARY, height=6,
                                            corner_radius=3)
        self._progress.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_XS))
        self._progress.set(0)

        self._prog_var = tk.StringVar(value="")
        ctk.CTkLabel(p, textvariable=self._prog_var, font=T.FONT_SMALL,
                     text_color=T.MUTED).pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_M))

        # Export button
        self._btn = ctk.CTkButton(
            p, text="🖼️  Exportar imagens",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._run, state="disabled",
        )
        self._btn.pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    def _set_fmt(self, fmt):
        self._fmt_var.set(fmt)
        for f, btn in self._fmt_btns.items():
            if f == fmt:
                btn.configure(fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
                              text_color="white")
            else:
                btn.configure(fg_color=T.BG_SECONDARY, hover_color=T.SURFACE_HOVER,
                              text_color=T.FG_SECONDARY)

    def _set_dpi(self, dpi):
        self._dpi_var.set(dpi)
        for d, btn in self._dpi_btns.items():
            if d == dpi:
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
        self._drop.set_file(os.path.basename(path))
        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._btn.configure(state="disabled")
        self._progress.set(0)
        self._prog_var.set("")
        self.set_status("Pronto")

    def _run(self):
        folder = filedialog.askdirectory(title="Escolha a pasta de saída")
        if not folder:
            return
        fmt = self._fmt_var.get()
        dpi = int(self._dpi_var.get())
        path = self._path
        self._btn.configure(state="disabled")
        self._progress.set(0)
        self.set_status("⏳  Convertendo…")

        def update(current, total):
            pct = current / total
            self.root.after(0, lambda c=current, t=total, v=pct: (
                self._progress.set(v),
                self._prog_var.set(f"Página {c} de {t}"),
            ))

        def task():
            try:
                doc = fitz.open(path)
                n = len(doc)
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                for i, page in enumerate(doc):
                    pix = page.get_pixmap(matrix=mat)
                    pix.save(os.path.join(folder, f"pagina_{i + 1}.{fmt}"))
                    update(i + 1, n)
                doc.close()
                self.root.after(0, lambda: self._done(n, folder))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, n: int, folder: str):
        self._btn.configure(state="normal")
        self._progress.set(1)
        self._prog_var.set(f"✓  {n} imagem(ns) salva(s)")
        self.set_status(f"✓  {n} imagem(ns) salva(s) em {os.path.basename(folder)}")
        messagebox.showinfo("Sucesso", f"{n} página(s) convertida(s)!\nPasta: {folder}")

    def _error(self, msg: str):
        self._btn.configure(state="normal")
        self.set_status("✗  Erro na conversão")
        messagebox.showerror("Erro", msg)
