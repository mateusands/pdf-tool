---
name: rodar-local
description: Rodar o Gerenciador de PDF e Word localmente (venv, dependências por plataforma, LibreOffice para Word→PDF no Linux) e o roteiro de teste manual das 10 abas. Use ao rodar, testar manualmente ou debugar o ambiente.
---

# Rodar o Gerenciador de PDF e Word localmente

## Setup

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

O `.venv/` já está no `.gitignore`.

## Pré-requisitos do sistema

| Requisito | Por quê | Como instalar |
|---|---|---|
| **Tk** | `tkinter` é stdlib mas depende da lib Tk do SO | `sudo pacman -S tk` · `sudo apt install python3-tk` |
| **LibreOffice** | só para **Word → PDF no Linux** | `sudo pacman -S libreoffice-fresh` · `sudo apt install libreoffice-writer` |
| **Microsoft Word** | só para **Word → PDF no Windows/macOS** | vem com o Office |

Verificar o Tk:

```bash
python -c "import tkinter; print('Tk', tkinter.TkVersion)"
```

Sem ele: `ImportError: libtk8.6.so: cannot open shared object file`.

## Pegadinhas

- **`docx2pdf` não é instalado no Linux** — o `requirements.txt` tem o marker `sys_platform != "linux"`.
  Isso é intencional. O pip informa: `Ignoring docx2pdf: markers ... don't match your environment`.

- **Word → PDF falha com mensagem pedindo LibreOffice** → é o `pdf_tool/core/docx_convert.py` funcionando como deveria.
  Instale o LibreOffice; o resto do app não depende dele.

- **LibreOffice converte "em silêncio" e nada aparece** → normalmente é conflito de perfil, quando já há
  uma janela do LibreOffice aberta. O `pdf_tool/core/docx_convert.py` já contorna isso com `-env:UserInstallation`
  apontando para um perfil temporário. Se voltar a acontecer, é aí que se olha.

- **A UI parece diferente do Windows** → a fonte `Segoe UI` (em `pdf_tool/theme.py`) não existe no Linux e o Tk faz
  fallback silencioso. Esperado, não é bug.

- **A janela congela ao processar arquivo grande** → só 4 das 10 abas usam thread. `merge`, `split`,
  `protect`, `unlock`, `rotate` e `organize` rodam na UI thread. Não é travamento de verdade: espere.

- **PDF protegido por senha** faz `pypdf`/`pymupdf` levantarem exceção na abertura, não no processamento.
  Teste isso de propósito — várias abas não tratam esse caso.

- **Sem servidor gráfico não roda** (SSH, container): `no display name and no $DISPLAY`.

---

## Roteiro de teste manual

A suíte (`pytest`) cobre só a lógica extraída; nenhuma aba é coberta de ponta a ponta. **Este roteiro
cobre o resto.** Rode `pytest` antes de abrir o app. Prepare antes: um PDF de 1 página, um PDF de
50+ páginas, um PDF protegido por senha, um `.docx` e duas imagens PNG/JPG.

### Abas principais
| Aba | O que verificar |
|---|---|
| ✂️ **Dividir** | miniaturas renderizam; seleção de páginas; o PDF gerado tem exatamente as escolhidas |
| 🔗 **Juntar** | ordem respeitada; total de páginas = soma; com arquivo grande, observe se a UI trava |
| 🔄 **Converter** | PDF→DOCX abre no Word/LibreOffice; DOCX→PDF (no Linux exige LibreOffice) |
| ⚙️ **Organizar** | arrastar reordena; a ordem da tela é a do arquivo salvo |

### Mais Ferramentas
| Aba | O que verificar |
|---|---|
| 📦 **Compactar** | os 3 níveis geram tamanhos diferentes; o PDF continua legível |
| 🖼️ **PDF → Imagem** | PNG e JPG; DPIs 72/150/300 mudam a resolução de fato |
| 📄 **Imagem → PDF** | ordem das imagens; múltiplas páginas |
| 🔒 **Proteger** | o PDF gerado **pede senha** ao abrir; o original continua sem senha |
| 🔓 **Desbloquear** | abre sem senha; senha errada dá erro claro, não trava |
| 🔃 **Girar** | 90°, −90°, 180°; página específica e todas |

### Transversal (o que mais pega bug)
- **Cancelar o diálogo de salvar** no meio → o botão volta a ficar habilitado, sem estado preso
- **Arquivo protegido por senha** em cada aba → erro tratado, não stack trace no terminal
- **Status bar** acompanha: ícone de processando durante, de sucesso ou erro no fim — nada de status
  parado em "Processando…" para sempre
- **O arquivo de entrada não foi alterado** em nenhuma operação
- **Trocar de aba durante um processamento** não quebra nada

---

## Empacotar (opcional)

```bash
pyinstaller --onefile --windowed main.py
```

Atenção: `pymupdf` e `cryptography` trazem binários nativos e costumam exigir ajuste de hidden imports.
Empacotamento não é parte do projeto hoje — trate como decisão a ser pedida.
