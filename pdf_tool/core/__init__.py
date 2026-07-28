"""Lógica pura do app — recebe caminhos e parâmetros, não toca em widget.

É a camada que a suíte de testes cobre, e por isso nada aqui pode importar
Tkinter, CustomTkinter ou qualquer módulo de UI.

    pdf_io.py        toda a escrita de PDF: destino ≠ origem, escrita atômica
    docx_convert.py  Word → PDF (Word no Windows/macOS, LibreOffice no Linux)
    background.py    trabalho pesado em thread, com retorno seguro para a UI
    reorder.py       reordenação de listas (páginas e arquivos)
"""
