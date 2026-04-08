import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfWriter

from constants import BORDER, CARD, FG, PRIMARY
from widgets import section_title


class TabMerge:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status
        self.root = root
        self._files: list = []
        self._build(parent)

    def _build(self, p):
        section_title(p, "Juntar PDFs", "Combine múltiplos PDFs em um único arquivo.")

        list_frame = tk.Frame(p, bg=BORDER, bd=1, relief="solid")
        list_frame.pack(fill="both", expand=True)
        vsb = ttk.Scrollbar(list_frame, orient="vertical")
        self._list = tk.Listbox(
            list_frame,
            yscrollcommand=vsb.set,
            selectmode=tk.SINGLE,
            bd=0,
            relief="flat",
            bg=CARD,
            fg=FG,
            selectbackground=PRIMARY,
            selectforeground="white",
            font=("Segoe UI", 9),
            activestyle="none",
            height=8,
        )
        vsb.config(command=self._list.yview)
        self._list.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        vsb.pack(side="right", fill="y")

        ctrl = ttk.Frame(p, style="Card.TFrame")
        ctrl.pack(fill="x", pady=10)
        ttk.Button(
            ctrl, text="+ Adicionar", style="Ghost.TButton", command=self._add
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            ctrl, text="Remover", style="Danger.TButton", command=self._remove
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            ctrl, text="↑ Subir", style="Ghost.TButton", command=self._move_up
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            ctrl, text="↓ Descer", style="Ghost.TButton", command=self._move_down
        ).pack(side="left")

        ttk.Button(
            p, text="Juntar e salvar PDF", style="Success.TButton", command=self._merge
        ).pack(anchor="w")

    def _add(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for f in files:
            self._files.append(f)
            self._list.insert(tk.END, f"  {os.path.basename(f)}")
        if files:
            self.set_status(f"{len(files)} arquivo(s) adicionados")

    def _remove(self):
        sel = self._list.curselection()
        if sel:
            i = sel[0]
            self._list.delete(i)
            self._files.pop(i)

    def _move_up(self):
        sel = self._list.curselection()
        if sel and sel[0] > 0:
            i = sel[0]
            t = self._list.get(i)
            self._list.delete(i)
            self._list.insert(i - 1, t)
            self._list.selection_set(i - 1)
            self._files[i], self._files[i - 1] = self._files[i - 1], self._files[i]

    def _move_down(self):
        sel = self._list.curselection()
        if sel and sel[0] < self._list.size() - 1:
            i = sel[0]
            t = self._list.get(i)
            self._list.delete(i)
            self._list.insert(i + 1, t)
            self._list.selection_set(i + 1)
            self._files[i], self._files[i + 1] = self._files[i + 1], self._files[i]

    def _merge(self):
        if not self._files:
            messagebox.showwarning("Aviso", "Adicione pelo menos um arquivo PDF.")
            return
        save = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
        )
        if save:
            try:
                self.set_status("Juntando PDFs…")
                writer = PdfWriter()
                for f in self._files:
                    writer.append(f)
                writer.write(save)
                writer.close()
                self.set_status(f"Salvo: {os.path.basename(save)}")
                messagebox.showinfo("Sucesso", "PDFs juntados com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                self.set_status("Erro ao juntar PDFs")
