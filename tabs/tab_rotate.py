import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pypdf import PdfReader, PdfWriter

import theme as T
from widgets import DropZone, ThumbnailGrid, section_title


class TabRotate:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "🔃  Girar PDF",
                      "Gire todas ou páginas específicas de um PDF.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Toolbar
        bar = ctk.CTkFrame(p, fg_color="transparent")
        bar.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_S))

        self._count_var = tk.StringVar(value="")
        ctk.CTkLabel(bar, textvariable=self._count_var,
                     font=T.FONT_BODY, text_color=T.PRIMARY).pack(side="left")

        self._btn_none = ctk.CTkButton(
            bar, text="Limpar", fg_color="transparent", border_width=1,
            border_color=T.BORDER, text_color=T.MUTED, font=T.FONT_BODY,
            width=80, height=32, corner_radius=6, cursor="hand2",
            command=lambda: self._grid.deselect_all(),
        )
        self._btn_all = ctk.CTkButton(
            bar, text="Tudo", fg_color="transparent", border_width=1,
            border_color=T.BORDER, text_color=T.MUTED, font=T.FONT_BODY,
            width=80, height=32, corner_radius=6, cursor="hand2",
            command=lambda: self._grid.select_all(),
        )

        # Bottom controls — packed before the expanding grid
        bottom = ctk.CTkFrame(p, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=T.PAD_L, pady=(T.PAD_S, T.PAD_L))

        # Save button
        self._btn_save = ctk.CTkButton(
            bottom, text="🔃  Girar e salvar",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._save, state="disabled",
        )
        self._btn_save.pack(anchor="w", pady=(T.PAD_S, 0))

        # Rotation angle — visual buttons
        angle_frame = ctk.CTkFrame(bottom, fg_color="transparent")
        angle_frame.pack(anchor="w", pady=(0, T.PAD_XS))

        ctk.CTkLabel(angle_frame, text="Rotação:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(side="left", padx=(0, T.PAD_M))

        self._angle_var = tk.IntVar(value=90)
        self._angle_btns: dict[int, ctk.CTkButton] = {}

        for angle, label in [(90, "↻ 90°"), (-90, "↺ 90°"), (180, "↕ 180°")]:
            btn = ctk.CTkButton(
                angle_frame, text=label, width=100, height=38,
                corner_radius=8, cursor="hand2", font=T.FONT_PILL,
                fg_color=T.PRIMARY if angle == 90 else T.BG_SECONDARY,
                hover_color=T.PRIMARY_HOVER if angle == 90 else T.SURFACE_HOVER,
                text_color="white" if angle == 90 else T.FG_SECONDARY,
                border_width=1, border_color=T.BORDER,
                command=lambda a=angle: self._set_angle(a),
            )
            btn.pack(side="left", padx=(0, 6))
            self._angle_btns[angle] = btn

        # Thumbnail grid
        self._grid = ThumbnailGrid(p)
        self._grid.on_change(self._on_change)

    def _set_angle(self, angle: int):
        self._angle_var.set(angle)
        for a, btn in self._angle_btns.items():
            if a == angle:
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
        self.set_status("Carregando páginas…")
        self._btn_all.pack(side="right", padx=(4, 0))
        self._btn_none.pack(side="right")
        n = self._grid.load_pdf(path)
        self.set_status(f"✓  {n} página(s) carregadas — clique para selecionar")

    def _clear(self):
        self._path = None
        self._count_var.set("")
        self._btn_all.pack_forget()
        self._btn_none.pack_forget()
        self._btn_save.configure(state="disabled")
        self._grid.clear()
        self.set_status("Pronto")

    def _on_change(self, selected):
        n = len(selected)
        self._count_var.set(f"{n} página(s) selecionada(s)" if n else "")
        self._btn_save.configure(state="normal" if n else "disabled")

    def _save(self):
        selected = self._grid.get_selected()
        if not selected:
            return
        angle = self._angle_var.get()
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        try:
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for i, page in enumerate(reader.pages):
                if i in selected:
                    page.rotate(angle)
                writer.add_page(page)
            with open(save, "wb") as f:
                writer.write(f)
            self.set_status(f"✓  Salvo: {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", "PDF girado e salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
