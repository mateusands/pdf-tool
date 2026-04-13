import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pypdf import PdfReader, PdfWriter

import theme as T
from widgets import DropZone, ThumbnailGrid, section_title


class TabSplit:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "✂️  Dividir PDF",
                      "Selecione o PDF e clique nas páginas que deseja extrair.")

        self._drop = DropZone(p, icon="📄", text="Selecione um PDF",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Toolbar: select-all / clear / count
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

        # Save button (bottom)
        self._btn_save = ctk.CTkButton(
            p, text="💾  Salvar páginas selecionadas",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._save, state="disabled",
        )
        self._btn_save.pack(side="bottom", anchor="w", padx=T.PAD_L, pady=(T.PAD_S, T.PAD_L))

        self._grid = ThumbnailGrid(p)
        self._grid.on_change(self._on_selection_change)

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
        self.set_status(f"{n} página(s) carregadas — clique para selecionar")

    def _clear(self):
        self._path = None
        self._count_var.set("")
        self._btn_all.pack_forget()
        self._btn_none.pack_forget()
        self._btn_save.configure(state="disabled")
        self._grid.clear()
        self.set_status("Pronto")

    def _on_selection_change(self, selected):
        n = len(selected)
        self._count_var.set(f"{n} página(s) selecionada(s)" if n else "")
        self._btn_save.configure(state="normal" if n else "disabled")

    def _save(self):
        selected = self._grid.get_selected()
        if not selected:
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return
        try:
            reader = PdfReader(self._path)
            writer = PdfWriter()
            for i in sorted(selected):
                writer.add_page(reader.pages[i])
            with open(save, "wb") as f:
                writer.write(f)
            n = len(selected)
            self.set_status(f"✓  {n} página(s) salvas em {os.path.basename(save)}")
            messagebox.showinfo("Sucesso", f"{n} página(s) salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", str(e))
