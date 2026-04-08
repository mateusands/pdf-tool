# Gerenciador de PDF e Word

Aplicação desktop em Python com interface gráfica (Tkinter) para manipulação de arquivos PDF e Word.

## Screenshots

<img width="594" height="526" alt="Screenshot_10" src="https://github.com/user-attachments/assets/13061a6c-7bd4-4cc0-8287-11aef3285437" />

<img width="591" height="521" alt="Screenshot_11" src="https://github.com/user-attachments/assets/cedf6103-73c3-4db5-a83f-e5bc02c14aa1" />

<img width="591" height="519" alt="Screenshot_12" src="https://github.com/user-attachments/assets/82373684-f2e8-4833-a118-4d1c522ba7e8" />

## Funcionalidades

### Ferramentas principais
| Aba | Descrição |
|-----|-----------|
| **Dividir PDF** | Visualize as páginas em miniaturas e selecione quais deseja extrair |
| **Juntar PDFs** | Combine múltiplos PDFs em um só, organizando a ordem livremente |
| **Converter** | Converta PDF → Word (.docx) ou Word → PDF automaticamente |

### Mais Ferramentas
| Aba | Descrição |
|-----|-----------|
| **Compactar** | Reduza o tamanho de um PDF com níveis baixo, médio ou alto |
| **PDF → Imagem** | Exporte cada página como PNG ou JPG em DPI configurável |
| **Imagem → PDF** | Combine imagens PNG/JPG em um único PDF, reordenando como quiser |
| **Proteger** | Adicione senha a um PDF gerando uma cópia protegida |
| **Desbloquear** | Remova a senha de um PDF protegido gerando uma cópia livre |
| **Girar** | Gire páginas específicas ou todas as páginas em 90°, -90° ou 180° |

## Requisitos

- Python 3.10+
- Windows (necessário para `docx2pdf`, que usa o Microsoft Word instalado)

## Instalação

```bash
# Clone o repositório
git clone https://github.com/mateusands/pdf-tool.git
cd pdf-tool

# Crie e ative o ambiente virtual (recomendado: uv)
uv venv
.venv\Scripts\activate

# Instale as dependências
uv pip install -r requeriments.txt
```

Ou com pip tradicional:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requeriments.txt
```

## Como usar

```bash
python pdf_toll.py
```

## Estrutura do projeto

```
pdfpython/
├── pdf_toll.py              # Ponto de entrada e shell da aplicação
├── constants.py             # Paleta de cores e constantes visuais
├── widgets.py               # Componentes reutilizáveis (FileRow, ThumbnailGrid)
├── requeriments.txt         # Dependências do projeto
└── tabs/
    ├── tab_split.py         # Dividir PDF
    ├── tab_merge.py         # Juntar PDFs
    ├── tab_convert.py       # Converter PDF ↔ Word
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
| `pypdf` | Leitura, escrita, divisão, rotação e criptografia de PDFs |
| `pymupdf` | Renderização de miniaturas, compactação e conversão PDF↔imagem |
| `pdf2docx` | Conversão de PDF para Word |
| `docx2pdf` | Conversão de Word para PDF (requer Microsoft Word no Windows) |
| `cryptography` | Suporte a PDFs criptografados com AES |
