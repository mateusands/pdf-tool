import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from pdf2docx import Converter

import theme as T
from docx_convert import docx_to_pdf
from widgets import DropZone, section_title


class TabConvert:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._path = None
        self._mode = None
        self._build(parent)

    def _build(self, p):
        section_title(p, "🔄  Converter PDF ↔ Word",
                      "Converta entre PDF e DOCX automaticamente.")

        # Two option cards side by side
        cards = ctk.CTkFrame(p, fg_color="transparent")
        cards.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_L))
        cards.grid_columnconfigure(0, weight=1)
        cards.grid_columnconfigure(1, weight=1)

        self._card_pdf = self._make_card(cards, "📄 → 📝", "PDF → Word",
                                          "Converta PDF para documento Word editável")
        self._card_pdf.grid(row=0, column=0, padx=(0, 6), sticky="nsew")

        self._card_word = self._make_card(cards, "📝 → 📄", "Word → PDF",
                                           "Converta documento Word para PDF")
        self._card_word.grid(row=0, column=1, padx=(6, 0), sticky="nsew")

        # Drop zone
        self._drop = DropZone(p, icon="📁", text="Selecione um arquivo PDF ou DOCX",
                              subtitle="ou clique para selecionar",
                              command=self._select, on_clear=self._clear)

        # Detected action badge
        badge = ctk.CTkFrame(p, fg_color="transparent")
        badge.pack(fill="x", padx=T.PAD_L, pady=(0, T.PAD_M))
        ctk.CTkLabel(badge, text="Ação detectada:", font=T.FONT_BODY,
                     text_color=T.MUTED).pack(side="left")
        self._action_var = tk.StringVar(value="—")
        ctk.CTkLabel(badge, textvariable=self._action_var, font=(T.FONT_FAMILY, 12, "bold"),
                     text_color=T.SUCCESS).pack(side="left", padx=(8, 0))

        # Convert button
        self._btn = ctk.CTkButton(
            p, text="🔄  Converter agora",
            fg_color=T.PRIMARY, hover_color=T.PRIMARY_HOVER,
            font=T.FONT_BUTTON, cursor="hand2", corner_radius=8, width=0, height=48,
            command=self._convert, state="disabled",
        )
        self._btn.pack(anchor="w", padx=T.PAD_L, pady=(0, T.PAD_L))

    def _make_card(self, parent, icon, title, desc):
        card = ctk.CTkFrame(parent, fg_color=T.BG_SECONDARY,
                            corner_radius=T.RADIUS, border_width=1,
                            border_color=T.BORDER)
        ctk.CTkLabel(card, text=icon, font=(T.FONT_FAMILY, 24),
                     text_color=T.FG).pack(pady=(T.PAD_L, T.PAD_S))
        ctk.CTkLabel(card, text=title, font=T.FONT_BUTTON,
                     text_color=T.FG).pack()
        ctk.CTkLabel(card, text=desc, font=T.FONT_SMALL,
                     text_color=T.MUTED, wraplength=200).pack(pady=(4, T.PAD_L))
        return card

    def _select(self):
        path = filedialog.askopenfilename(
            filetypes=[("Documentos", "*.pdf *.docx")])
        if not path:
            return
        self._path = path
        self._drop.set_file(os.path.basename(path))
        ext = os.path.splitext(path)[1].lower()

        if ext == ".pdf":
            self._mode = "pdf2word"
            self._action_var.set("PDF → Word (.docx)")
            self._card_pdf.configure(border_color=T.PRIMARY)
            self._card_word.configure(border_color=T.BORDER)
        else:
            self._mode = "word2pdf"
            self._action_var.set("Word (.docx) → PDF")
            self._card_word.configure(border_color=T.PRIMARY)
            self._card_pdf.configure(border_color=T.BORDER)

        self._btn.configure(state="normal")
        self.set_status(f"Selecionado: {os.path.basename(path)}")

    def _clear(self):
        self._path = None
        self._mode = None
        self._action_var.set("—")
        self._btn.configure(state="disabled")
        self._card_pdf.configure(border_color=T.BORDER)
        self._card_word.configure(border_color=T.BORDER)
        self.set_status("Pronto")

    def _convert(self):
        if self._mode == "pdf2word":
            save = filedialog.asksaveasfilename(
                defaultextension=".docx", filetypes=[("Word", "*.docx")])
        else:
            save = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not save:
            return

        self._btn.configure(state="disabled")
        mode, path = self._mode, self._path
        self.set_status("⏳  Convertendo…")

        def task():
            try:
                if mode == "pdf2word":
                    cv = Converter(path)
                    cv.convert(save, start=0, end=None)
                    cv.close()
                else:
                    docx_to_pdf(path, save)
                self.root.after(0, lambda: self._done(save))
            except Exception as e:
                self.root.after(0, lambda: self._error(str(e)))

        threading.Thread(target=task, daemon=True).start()

    def _done(self, save):
        self._btn.configure(state="normal")
        self.set_status(f"✓  Salvo: {os.path.basename(save)}")
        messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")

    def _error(self, msg):
        self._btn.configure(state="normal")
        self.set_status("✗  Erro na conversão")
        messagebox.showerror("Erro", msg)
