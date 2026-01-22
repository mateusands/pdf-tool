import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf_convert

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerenciador de PDF e Word")
        self.root.geometry("600x500")

        # Configuração das Abas
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill="both")

        # Criar os frames para cada aba
        self.frame_split = ttk.Frame(self.notebook)
        self.frame_merge = ttk.Frame(self.notebook)
        self.frame_convert = ttk.Frame(self.notebook)

        self.notebook.add(self.frame_split, text="Dividir PDF")
        self.notebook.add(self.frame_merge, text="Juntar PDFs")
        self.notebook.add(self.frame_convert, text="Converter (PDF <-> Word)")

        # Inicializar as funcionalidades
        self.setup_split_tab()
        self.setup_merge_tab()
        self.setup_convert_tab()

    # ==========================
    # ABA 1: DIVIDIR PDF
    # ==========================
    def setup_split_tab(self):
        lbl = ttk.Label(self.frame_split, text="Selecione um PDF para extrair páginas", font=("Arial", 12))
        lbl.pack(pady=10)

        self.btn_select_split = ttk.Button(self.frame_split, text="Selecionar PDF", command=self.select_pdf_to_split)
        self.btn_select_split.pack(pady=5)

        self.lbl_split_file = ttk.Label(self.frame_split, text="Nenhum arquivo selecionado", foreground="gray")
        self.lbl_split_file.pack(pady=5)

        lbl_instr = ttk.Label(self.frame_split, text="Digite as páginas (ex: 1,3,5 ou 1-3):")
        lbl_instr.pack(pady=(20, 5))

        self.entry_pages = ttk.Entry(self.frame_split, width=30)
        self.entry_pages.pack(pady=5)

        self.btn_split_save = ttk.Button(self.frame_split, text="Salvar Páginas Selecionadas", command=self.split_pdf, state="disabled")
        self.btn_split_save.pack(pady=20)

    def select_pdf_to_split(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.split_file_path = file_path
            self.lbl_split_file.config(text=os.path.basename(file_path))
            self.btn_split_save.config(state="normal")

    def split_pdf(self):
        pages_str = self.entry_pages.get()
        if not pages_str:
            messagebox.showwarning("Aviso", "Digite quais páginas deseja extrair.")
            return

        try:
            reader = PdfReader(self.split_file_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            
            # Lógica simples para processar strings como "1,2,5-7"
            selected_indices = set()
            parts = pages_str.split(',')
            for part in parts:
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    # Ajuste para índice 0 e inclusivo
                    selected_indices.update(range(start-1, end))
                else:
                    selected_indices.add(int(part) - 1)

            # Adicionar páginas selecionadas
            for idx in sorted(selected_indices):
                if 0 <= idx < total_pages:
                    writer.add_page(reader.pages[idx])

            save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if save_path:
                with open(save_path, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

    # ==========================
    # ABA 2: JUNTAR PDFS
    # ==========================
    def setup_merge_tab(self):
        lbl = ttk.Label(self.frame_merge, text="Adicione PDFs e organize a ordem", font=("Arial", 12))
        lbl.pack(pady=10)

        # Listbox para mostrar arquivos
        self.list_merge = tk.Listbox(self.frame_merge, selectmode=tk.SINGLE, width=50, height=10)
        self.list_merge.pack(pady=5, padx=10)

        # Botões de controle
        frame_btns = ttk.Frame(self.frame_merge)
        frame_btns.pack(pady=5)

        ttk.Button(frame_btns, text="Adicionar Arquivos", command=self.add_pdfs_to_merge).grid(row=0, column=0, padx=5)
        ttk.Button(frame_btns, text="Remover", command=self.remove_pdf_from_merge).grid(row=0, column=1, padx=5)
        
        frame_order = ttk.Frame(self.frame_merge)
        frame_order.pack(pady=5)
        ttk.Button(frame_order, text="Mover Para Cima", command=self.move_up).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_order, text="Mover Para Baixo", command=self.move_down).pack(side=tk.LEFT, padx=5)

        ttk.Button(self.frame_merge, text="Juntar e Salvar PDF", command=self.merge_pdfs).pack(pady=20)
        
        self.merge_files = [] # Lista para guardar caminhos completos

    def add_pdfs_to_merge(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for f in files:
            self.merge_files.append(f)
            self.list_merge.insert(tk.END, os.path.basename(f))

    def remove_pdf_from_merge(self):
        sel = self.list_merge.curselection()
        if sel:
            idx = sel[0]
            self.list_merge.delete(idx)
            self.merge_files.pop(idx)

    def move_up(self):
        sel = self.list_merge.curselection()
        if sel:
            idx = sel[0]
            if idx > 0:
                text = self.list_merge.get(idx)
                self.list_merge.delete(idx)
                self.list_merge.insert(idx-1, text)
                self.list_merge.selection_set(idx-1)
                # Atualizar lista interna
                self.merge_files[idx], self.merge_files[idx-1] = self.merge_files[idx-1], self.merge_files[idx]

    def move_down(self):
        sel = self.list_merge.curselection()
        if sel:
            idx = sel[0]
            if idx < self.list_merge.size() - 1:
                text = self.list_merge.get(idx)
                self.list_merge.delete(idx)
                self.list_merge.insert(idx+1, text)
                self.list_merge.selection_set(idx+1)
                # Atualizar lista interna
                self.merge_files[idx], self.merge_files[idx+1] = self.merge_files[idx+1], self.merge_files[idx]

    def merge_pdfs(self):
        if not self.merge_files:
            messagebox.showwarning("Aviso", "Adicione pelo menos um arquivo PDF.")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if save_path:
            try:
                merger = PdfWriter()
                for pdf in self.merge_files:
                    merger.append(pdf)
                merger.write(save_path)
                merger.close()
                messagebox.showinfo("Sucesso", "PDFs juntados com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao juntar: {str(e)}")

    # ==========================
    # ABA 3: CONVERTER
    # ==========================
    def setup_convert_tab(self):
        lbl = ttk.Label(self.frame_convert, text="Converta entre PDF e Word", font=("Arial", 12))
        lbl.pack(pady=10)

        self.btn_select_conv = ttk.Button(self.frame_convert, text="Selecionar Arquivo (PDF ou DOCX)", command=self.select_file_convert)
        self.btn_select_conv.pack(pady=5)

        self.lbl_conv_file = ttk.Label(self.frame_convert, text="Nenhum arquivo selecionado", foreground="gray")
        self.lbl_conv_file.pack(pady=5)

        self.lbl_action = ttk.Label(self.frame_convert, text="Ação detectada: ...")
        self.lbl_action.pack(pady=10)

        self.btn_run_convert = ttk.Button(self.frame_convert, text="Converter Agora", command=self.run_conversion, state="disabled")
        self.btn_run_convert.pack(pady=20)

    def select_file_convert(self):
        file_path = filedialog.askopenfilename(filetypes=[("Documents", "*.pdf *.docx")])
        if file_path:
            self.convert_file_path = file_path
            self.lbl_conv_file.config(text=os.path.basename(file_path))
            
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".pdf":
                self.lbl_action.config(text="Ação: Converter PDF para Word")
                self.conversion_mode = "pdf2word"
            elif ext == ".docx":
                self.lbl_action.config(text="Ação: Converter Word para PDF")
                self.conversion_mode = "word2pdf"
            
            self.btn_run_convert.config(state="normal")

    def run_conversion(self):
        try:
            input_file = self.convert_file_path
            
            if self.conversion_mode == "pdf2word":
                save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word File", "*.docx")])
                if save_path:
                    # Usando pdf2docx
                    cv = Converter(input_file)
                    cv.convert(save_path, start=0, end=None)
                    cv.close()
                    messagebox.showinfo("Sucesso", "Convertido para Word com sucesso!")
            
            elif self.conversion_mode == "word2pdf":
                save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF File", "*.pdf")])
                if save_path:
                    # Usando docx2pdf
                    docx_to_pdf_convert(input_file, save_path)
                    messagebox.showinfo("Sucesso", "Convertido para PDF com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro na conversão: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolApp(root)
    root.mainloop()