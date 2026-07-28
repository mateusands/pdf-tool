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
| 🔒 **Proteger** | Adicione senha a um PDF gerando uma cópia protegida com AES-256 |
| 🔓 **Desbloquear** | Remova a senha de um PDF protegido gerando uma cópia livre |
| 🔃 **Girar** | Gire páginas específicas ou todas em 90°, -90° ou 180° |

## Interface

- **Sidebar única** com as 10 ferramentas agrupadas por finalidade — nenhuma
  fica escondida atrás de um clique extra
- **Ícones vetoriais** ([Lucide](https://lucide.dev), ISC) rasterizados em tempo
  de execução no tamanho e na cor de cada estado. Nada de emoji: emoji muda de
  desenho conforme o sistema, tem cor fixa e no Linux depende de fonte instalada
- **Elevação tonal** — no escuro a profundidade vem da luminância da superfície,
  não de sombra: cinco níveis subindo ~6% cada, nenhum preto puro
- **Contraste verificado, não estimado** — cada par de cor do tema é medido pelo
  WCAG e checado em teste (`pytest -k contraste`); todos passam no AA (4,5:1)
- **Tipografia nativa** — a família é escolhida entre as instaladas no sistema
  (Segoe UI, SF Pro, Inter/Cantarell/Noto), em vez de cair num fallback sorteado
- **Cabeçalho de contexto** com o nome e a descrição da ferramenta ativa
- **Status bar** com ícone de estado (pronto, processando, sucesso, erro)
- **Grade de miniaturas** com seleção por clique e arrastar para reordenar

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
python main.py
```

## Testes

A lógica de arquivo vive fora da interface (`pdf_io.py`, `docx_convert.py`,
`background.py`, `reorder.py`), então a suíte roda sem abrir janela:

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest                       # suíte completa
pytest -k destino            # um recorte
```

Cada arquivo de teste começa com a especificação do contrato que ele cobre: o que
a função promete, por que existe e qual regra de negócio está sendo protegida.

## Segurança dos seus arquivos

- **Nenhuma ferramenta sobrescreve o arquivo de entrada.** Se o destino escolhido
  for uma das origens, a operação é recusada com aviso em vez de executada.
- **A gravação é atômica**: o PDF é escrito num arquivo temporário e só então
  assume o nome final. Falha no meio do caminho não deixa arquivo truncado.
- **Proteger usa AES-256**, não o RC4 legado.

## Estrutura do projeto

```
pdf-tool/
├── main.py                      ← EXECUTÁVEL: python main.py
│
├── pdf_tool/                    # o aplicativo
│   ├── app.py                   # janela, barra de abas customizada, status bar
│   ├── theme.py                 # única fonte de cor, fonte e espaçamento
│   ├── widgets.py               # DropZone, ThumbnailGrid, rolagem multiplataforma
│   │
│   ├── core/                    # lógica pura, sem UI — é o que os testes cobrem
│   │   ├── pdf_io.py            # toda escrita de PDF: destino ≠ origem, gravação atômica
│   │   ├── docx_convert.py      # Word → PDF: Word no Windows/macOS, LibreOffice no Linux
│   │   ├── background.py        # trabalho pesado em thread + retorno seguro para a UI
│   │   └── reorder.py           # reordenação de listas (páginas e arquivos)
│   │
│   └── tabs/                    # uma classe por ferramenta
│       ├── tab_split.py         # Dividir PDF
│       ├── tab_merge.py         # Juntar PDFs
│       ├── tab_convert.py       # Converter PDF ↔ Word
│       ├── tab_organize.py      # Organizar páginas (drag-and-drop)
│       ├── tab_compress.py      # Compactar PDF
│       ├── tab_pdf_to_image.py  # PDF → Imagem
│       ├── tab_image_to_pdf.py  # Imagem → PDF
│       ├── tab_protect.py       # Proteger com senha
│       ├── tab_unlock.py        # Remover senha
│       └── tab_rotate.py        # Girar páginas
│
├── tests/                       # suíte da lógica pura — roda sem abrir janela
│   ├── test_pdf_io.py           # integridade do arquivo, senha, operações
│   ├── test_docx_convert.py     # escolha de motor por plataforma
│   ├── test_background.py       # erro em thread chega mesmo à tela
│   ├── test_reorder.py          # mover e arrastar
│   └── test_scroll.py           # roda do mouse por plataforma
│
├── requirements.txt             # dependências para rodar o app
├── requirements-dev.txt         # + pytest (desenvolvimento)
└── pytest.ini
```

A separação que importa: **`pdf_tool/core/` não importa Tkinter**. É por isso que a suíte roda sem
abrir janela e sem display.

## Dependências

| Pacote | Uso |
|--------|-----|
| `customtkinter` | Interface dark mode com widgets modernos e cantos arredondados |
| `pypdf` | Leitura, escrita, divisão, rotação e criptografia de PDFs |
| `pymupdf` | Renderização de miniaturas, compactação e conversão PDF↔imagem |
| `pdf2docx` | Conversão de PDF para Word |
| `docx2pdf` | Conversão de Word para PDF no Windows/macOS (no Linux usa-se o LibreOffice) |
| `cryptography` | Suporte a PDFs criptografados com AES |

Os ícones são do conjunto [Lucide](https://lucide.dev) v1.27.0, sob licença ISC.
Os desenhos ficam embutidos em `pdf_tool/core/icons.py` como caminhos SVG e são
rasterizados pelo `pymupdf` — não há pacote nem fonte a instalar por causa deles.
