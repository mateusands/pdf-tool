import tkinter as tk
from tkinter import ttk

from constants import BG, BORDER, CARD, DANGER, FG, MUTED, PRIMARY, SUCCESS
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


class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de PDF e Word")
        self.root.geometry("760x620")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self._setup_styles()
        self._build_header()
        self._build_notebook()
        self._build_statusbar()

    def _setup_styles(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")

        s.configure(".", background=BG, foreground=FG, font=("Segoe UI", 10))
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure(
            "TNotebook.Tab",
            background=BORDER,
            foreground=MUTED,
            padding=[14, 8],
            font=("Segoe UI", 10),
        )
        s.map(
            "TNotebook.Tab",
            background=[("selected", CARD)],
            foreground=[("selected", PRIMARY)],
        )

        s.configure("Card.TFrame", background=CARD, relief="flat")
        s.configure("TFrame", background=BG)
        s.configure("TSeparator", background=BORDER)

        s.configure(
            "Title.TLabel",
            background=CARD,
            foreground=FG,
            font=("Segoe UI", 13, "bold"),
        )
        s.configure(
            "Sub.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9)
        )
        s.configure(
            "File.TLabel",
            background=CARD,
            foreground=PRIMARY,
            font=("Segoe UI", 9, "bold"),
        )
        s.configure(
            "Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9)
        )
        s.configure(
            "Count.TLabel",
            background=CARD,
            foreground=PRIMARY,
            font=("Segoe UI", 9, "bold"),
        )
        s.configure(
            "Status.TLabel",
            background=BORDER,
            foreground=MUTED,
            font=("Segoe UI", 8),
            padding=[8, 4],
        )

        for name, bg, fg, abg in [
            ("Primary.TButton", PRIMARY, "white", "#1D4ED8"),
            ("Success.TButton", SUCCESS, "white", "#15803D"),
            ("Danger.TButton", DANGER, "white", "#B91C1C"),
        ]:
            s.configure(
                name,
                background=bg,
                foreground=fg,
                borderwidth=0,
                focusthickness=0,
                font=("Segoe UI", 9, "bold"),
                padding=[14, 7],
            )
            s.map(name, background=[("active", abg)])

        s.configure(
            "Ghost.TButton",
            background=CARD,
            foreground=PRIMARY,
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 9),
            padding=[12, 6],
        )
        s.map("Ghost.TButton", background=[("active", BG)])

        s.configure(
            "Small.TButton",
            background=CARD,
            foreground=MUTED,
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 8),
            padding=[8, 4],
        )
        s.map("Small.TButton", background=[("active", BG)])

        s.configure(
            "TRadiobutton", background=CARD, foreground=FG, font=("Segoe UI", 9)
        )
        s.configure(
            "TProgressbar",
            troughcolor=BORDER,
            background=PRIMARY,
            borderwidth=0,
            thickness=8,
        )
        s.configure(
            "TEntry",
            fieldbackground=CARD,
            bordercolor=BORDER,
            relief="solid",
            padding=[8, 6],
            font=("Segoe UI", 9),
        )

        # Sub-notebook (inside "Mais Ferramentas")
        s.configure("Sub.TNotebook", background=CARD, borderwidth=0)
        s.configure(
            "Sub.TNotebook.Tab",
            background=BORDER,
            foreground=MUTED,
            padding=[12, 6],
            font=("Segoe UI", 9),
        )
        s.map(
            "Sub.TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", PRIMARY)],
        )

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=PRIMARY, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr,
            text="  Gerenciador de PDF e Word",
            bg=PRIMARY,
            fg="white",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left", padx=16)

    def _build_notebook(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=16, pady=12)

        for label, TabClass in [
            ("  Dividir",   TabSplit),
            ("  Juntar",    TabMerge),
            ("  Converter", TabConvert),
            ("  Organizar", TabOrganize),
        ]:
            f = ttk.Frame(nb, style="TFrame")
            nb.add(f, text=label)
            card = ttk.Frame(f, style="Card.TFrame", padding=20)
            card.pack(fill="both", expand=True)
            TabClass(card, set_status=self._set_status, root=self.root)

        # "Mais Ferramentas" with sub-notebook
        mais = ttk.Frame(nb, style="TFrame")
        nb.add(mais, text="  Mais Ferramentas")

        sub = ttk.Notebook(mais, style="Sub.TNotebook")
        sub.pack(fill="both", expand=True, padx=8, pady=8)

        for label, TabClass in [
            ("Compactar",    TabCompress),
            ("PDF → Imagem", TabPdfToImage),
            ("Imagem → PDF", TabImageToPdf),
            ("Proteger",     TabProtect),
            ("Desbloquear",  TabUnlock),
            ("Girar",        TabRotate),
        ]:
            f = ttk.Frame(sub, style="TFrame")
            sub.add(f, text=f"  {label}")
            card = ttk.Frame(f, style="Card.TFrame", padding=20)
            card.pack(fill="both", expand=True)
            TabClass(card, set_status=self._set_status, root=self.root)

    def _build_statusbar(self):
        self._status_var = tk.StringVar(value="Pronto")
        ttk.Label(
            self.root, textvariable=self._status_var, style="Status.TLabel"
        ).pack(fill="x", side="bottom")

    def _set_status(self, msg: str):
        self._status_var.set(msg)
        self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolApp(root)
    root.mainloop()
