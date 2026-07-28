#!/usr/bin/env python3
"""Gerenciador de PDF e Word — ponto de entrada.

    python main.py

Todo o código vive no pacote `pdf_tool/`:

    pdf_tool/app.py     a janela, a barra de abas e a status bar
    pdf_tool/tabs/      uma classe por ferramenta (as 10 abas)
    pdf_tool/core/      lógica pura, sem UI — é o que a suíte de testes cobre
    pdf_tool/widgets.py componentes reutilizáveis
    pdf_tool/theme.py   única fonte de cor, fonte e espaçamento
"""

from pdf_tool.app import PDFToolApp


def main():
    PDFToolApp().root.mainloop()


if __name__ == "__main__":
    main()
