"""Gerenciador de PDF e Word — main application with dark-themed custom tab bar."""

import tkinter as tk

import customtkinter as ctk

import theme as T
from tabs.tab_compress import TabCompress
from tabs.tab_convert import TabConvert
from tabs.tab_image_to_pdf import TabImageToPdf
from tabs.tab_merge import TabMerge
from tabs.tab_organize import TabOrganize
from tabs.tab_pdf_to_image import TabPdfToImage
from tabs.tab_protect import TabProtect
from tabs.tab_rotate import TabRotate
from tabs.tab_split import TabSplit
from tabs.tab_unlock import TabUnlock

# ── Tab definitions ───────────────────────────────────────────────────────────

MAIN_TABS = [
    ("split",    "✂️  Dividir"),
    ("merge",    "🔗  Juntar"),
    ("convert",  "🔄  Converter"),
    ("organize", "⚙️  Organizar"),
    ("more",     "🧰  Mais Ferramentas"),
]

SUB_TABS = [
    ("compress", "📦 Compactar"),
    ("pdf2img",  "🖼️ PDF → Imagem"),
    ("img2pdf",  "📄 Imagem → PDF"),
    ("protect",  "🔒 Proteger"),
    ("unlock",   "🔓 Desbloquear"),
    ("rotate",   "🔃 Girar"),
]

TAB_CLASSES = {
    "split":    TabSplit,
    "merge":    TabMerge,
    "convert":  TabConvert,
    "organize": TabOrganize,
    "compress": TabCompress,
    "pdf2img":  TabPdfToImage,
    "img2pdf":  TabImageToPdf,
    "protect":  TabProtect,
    "unlock":   TabUnlock,
    "rotate":   TabRotate,
}


class PDFToolApp:
    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Gerenciador de PDF e Word")
        self.root.geometry("980x720")
        self.root.minsize(820, 600)
        self.root.configure(fg_color=T.BG)

        self._active_main = "split"
        self._active_sub = "compress"

        # Layout uses grid: header(0) | main_bar(1) | sub_bar(2) | content(3) | status(4)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_main_tab_bar()
        self._build_sub_tab_bar()
        self._build_content()
        self._build_statusbar()

        self._switch_main("split")

    # ── Header ────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = ctk.CTkFrame(self.root, fg_color=T.PRIMARY_DARK, height=52,
                           corner_radius=0)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.grid_propagate(False)

        ctk.CTkLabel(
            hdr, text="📄  Gerenciador de PDF e Word",
            font=T.FONT_HEADER, text_color="white",
        ).pack(side="left", padx=20, pady=10)

    # ── Main tab bar ──────────────────────────────────────────────────────────

    def _build_main_tab_bar(self):
        bar = ctk.CTkFrame(self.root, fg_color=T.BG_SECONDARY, height=62,
                           corner_radius=0)
        bar.grid(row=1, column=0, sticky="ew")
        bar.grid_propagate(False)

        # Use a plain tk.Frame so child widget widths propagate correctly
        inner = tk.Frame(bar, bg=T.BG_SECONDARY)
        inner.pack(side="left", padx=8, fill="y")

        self._main_btns: dict[str, ctk.CTkButton] = {}
        self._main_indicators: dict[str, tk.Frame] = {}

        import tkinter.font as tkfont
        measure_font = tkfont.Font(family=T.FONT_FAMILY, size=13)

        def _measure(text):
            measured = measure_font.measure(text)
            estimate = len(text) * 11
            return max(measured, estimate)

        for key, text in MAIN_TABS:
            col = tk.Frame(inner, bg=T.BG_SECONDARY)
            col.pack(side="left", padx=2, fill="y")

            btn_width = _measure(text) + 28

            btn = ctk.CTkButton(
                col, text=text, fg_color="transparent",
                hover_color=T.TAB_HOVER, text_color=T.MUTED,
                font=T.FONT_TAB, cursor="hand2", corner_radius=6,
                width=btn_width, height=42, anchor="center",
                command=lambda k=key: self._switch_main(k),
            )
            btn.pack(padx=4, pady=(8, 0))

            indicator = tk.Frame(col, bg=T.BG_SECONDARY, height=3)
            indicator.pack(fill="x", padx=8, pady=(4, 0))

            self._main_btns[key] = btn
            self._main_indicators[key] = indicator

    # ── Sub-tab bar (for "Mais Ferramentas") ──────────────────────────────────

    def _build_sub_tab_bar(self):
        self._sub_bar = ctk.CTkFrame(self.root, fg_color=T.SURFACE, height=52,
                                     corner_radius=0)
        self._sub_bar.grid(row=2, column=0, sticky="ew")
        self._sub_bar.grid_propagate(False)
        self._sub_bar.grid_remove()  # hidden initially

        inner = tk.Frame(self._sub_bar, bg=T.SURFACE)
        inner.pack(side="left", padx=12, fill="y")

        import tkinter.font as tkfont
        sub_font = tkfont.Font(family=T.FONT_FAMILY, size=12)

        def _measure_sub(text):
            measured = sub_font.measure(text)
            estimate = len(text) * 10
            return max(measured, estimate)

        self._sub_btns: dict[str, ctk.CTkButton] = {}
        for key, text in SUB_TABS:
            btn_width = _measure_sub(text) + 24

            btn = ctk.CTkButton(
                inner, text=text, fg_color="transparent",
                hover_color=T.TAB_HOVER, text_color=T.MUTED,
                font=T.FONT_SUB_TAB, cursor="hand2", corner_radius=14,
                width=btn_width, height=36, anchor="center",
                command=lambda k=key: self._switch_sub(k),
            )
            btn.pack(side="left", padx=3, pady=8)
            self._sub_btns[key] = btn

    # ── Content area ──────────────────────────────────────────────────────────

    def _build_content(self):
        container = ctk.CTkFrame(self.root, fg_color=T.BG, corner_radius=0)
        container.grid(row=3, column=0, sticky="nsew")

        self._pages: dict[str, ctk.CTkFrame] = {}

        for key, TabClass in TAB_CLASSES.items():
            page = ctk.CTkFrame(container, fg_color=T.BG)
            page.place(relx=0, rely=0, relwidth=1, relheight=1)

            card = ctk.CTkFrame(page, fg_color=T.SURFACE, corner_radius=T.RADIUS)
            card.pack(fill="both", expand=True, padx=20, pady=(12, 16))

            TabClass(card, set_status=self._set_status, root=self.root)
            self._pages[key] = page

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_statusbar(self):
        self._status_var = tk.StringVar(value="✓  Pronto")
        bar = ctk.CTkFrame(self.root, fg_color=T.BG_SECONDARY, height=28,
                           corner_radius=0)
        bar.grid(row=4, column=0, sticky="ew")
        bar.grid_propagate(False)
        ctk.CTkLabel(
            bar, textvariable=self._status_var,
            font=T.FONT_SMALL, text_color=T.MUTED,
        ).pack(side="left", padx=16, pady=4)

    # ── Tab switching ─────────────────────────────────────────────────────────

    def _switch_main(self, key: str):
        self._active_main = key

        for k, btn in self._main_btns.items():
            if k == key:
                btn.configure(text_color=T.FG, fg_color=T.TAB_ACTIVE)
                self._main_indicators[k].config(bg=T.TAB_INDICATOR)
            else:
                btn.configure(text_color=T.MUTED, fg_color="transparent")
                self._main_indicators[k].config(bg=T.BG_SECONDARY)

        if key == "more":
            self._sub_bar.grid()
            self._switch_sub(self._active_sub)
        else:
            self._sub_bar.grid_remove()
            self._pages[key].tkraise()

    def _switch_sub(self, key: str):
        self._active_sub = key
        for k, btn in self._sub_btns.items():
            if k == key:
                btn.configure(text_color="white", fg_color=T.PRIMARY,
                              hover_color=T.PRIMARY_HOVER)
            else:
                btn.configure(text_color=T.MUTED, fg_color="transparent",
                              hover_color=T.TAB_HOVER)
        self._pages[key].tkraise()

    def _set_status(self, msg: str):
        self._status_var.set(msg)
        self.root.update_idletasks()


if __name__ == "__main__":
    app = PDFToolApp()
    app.root.mainloop()
