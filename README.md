# Gerenciador de PDF e Word

Aplicação desktop em Python com interface gráfica dark mode profissional (CustomTkinter) para manipulação de arquivos PDF e Word.

## Screenshots

**Interface dark mode** com abas customizadas, drop zones interativas e botões estilizados.

<img width="981" height="752" alt="image" src="https://github.com/user-attachments/assets/23557529-b72b-477a-82f7-e1d3bf033c4e" />

<img width="978" height="748" alt="image" src="https://github.com/user-attachments/assets/c478ffff-4562-4947-8a72-f7a83ff3b0f2" />

## Funcionalidades

### Ferramentas principais
| Aba | Descrição |
|-----|-----------|
| ✂️ **Dividir PDF** | Visualize as páginas em miniaturas e selecione quais deseja extrair |
| 🔗 **Juntar PDFs** | Combine múltiplos PDFs em um só, organizando a ordem livremente |
| 🔄 **Converter** | Converta PDF → Word (.docx) ou Word → PDF automaticamente |
| ⚙️ **Organizar Páginas** | Reordene páginas arrastando miniaturas ou usando botões de mover |

### Mais Ferramentas
| Aba | Descrição |
|-----|-----------|
| 📦 **Compactar** | Reduza o tamanho de um PDF com níveis baixo, médio ou alto |
| 🖼️ **PDF → Imagem** | Exporte cada página como PNG ou JPG em DPI configurável (72/150/300) |
| 📄 **Imagem → PDF** | Combine imagens PNG/JPG em um único PDF, reordenando como quiser |
| 🔒 **Proteger** | Adicione senha a um PDF gerando uma cópia protegida |
| 🔓 **Desbloquear** | Remova a senha de um PDF protegido gerando uma cópia livre |
| 🔃 **Girar** | Gire páginas específicas ou todas em 90°, -90° ou 180° |

## Interface

- **Dark mode profissional** com acentos em azul elétrico (#2979FF)
- **Abas customizadas** com ícones, indicador de aba ativa e sub-abas em pills
- **Drop zones** com borda tracejada, hover animado e botão ✕ para remover arquivo
- **Botões pill** para seleção de opções (compressão, formato, DPI, rotação)
- **Grid de miniaturas** com seleção por clique e drag-and-drop para reordenar
- **Campos de senha** com toggle mostrar/ocultar
- **Feedback visual** com ícones de status (✓ sucesso, ✗ erro, ⏳ processando)

## Requisitos

- Python 3.10+
- Windows, macOS ou Linux

A conversão **Word → PDF** precisa de um motor externo, que varia por sistema:

| Sistema | Motor | Como obter |
|---------|-------|------------|
| Windows / macOS | Microsoft Word (via `docx2pdf`) | Já instalado com o Office |
| Linux | LibreOffice headless | `sudo pacman -S libreoffice-fresh` ou `sudo apt install libreoffice-writer` |

As demais abas não dependem disso e funcionam em qualquer sistema.

## Instalação

```bash
# Clone o repositório
git clone https://github.com/mateusands/pdf-tool.git
cd pdf-tool

# Crie e ative o ambiente virtual
python -m venv .venv
source .venv/bin/activate     # Windows: .venv\Scripts\activate

# Instale as dependências
pip install -r requirements.txt
```

## Como usar

```bash
python pdf_tool.py
```

## Estrutura do projeto

```
pdfpython/
├── pdf_tool.py              # Ponto de entrada — app principal com tab bar customizada
├── theme.py                 # Paleta dark mode, fontes e espaçamentos
├── widgets.py               # Componentes reutilizáveis (DropZone, ThumbnailGrid)
├── docx_convert.py          # Word → PDF: Word no Windows/macOS, LibreOffice no Linux
├── requirements.txt         # Dependências do projeto
└── tabs/
    ├── tab_split.py         # Dividir PDF
    ├── tab_merge.py         # Juntar PDFs
    ├── tab_convert.py       # Converter PDF ↔ Word
    ├── tab_organize.py      # Organizar páginas (drag-and-drop)
    ├── tab_compress.py      # Compactar PDF
    ├── tab_pdf_to_image.py  # PDF → Imagem
    ├── tab_image_to_pdf.py  # Imagem → PDF
    ├── tab_protect.py       # Proteger com senha
    ├── tab_unlock.py        # Remover senha
    └── tab_rotate.py        # Girar páginas
```

## Dependências

| Pacote | Uso |
|--------|-----|
| `customtkinter` | Interface dark mode com widgets modernos e cantos arredondados |
| `pypdf` | Leitura, escrita, divisão, rotação e criptografia de PDFs |
| `pymupdf` | Renderização de miniaturas, compactação e conversão PDF↔imagem |
| `pdf2docx` | Conversão de PDF para Word |
| `docx2pdf` | Conversão de Word para PDF no Windows/macOS (no Linux usa-se o LibreOffice) |
| `cryptography` | Suporte a PDFs criptografados com AES |
