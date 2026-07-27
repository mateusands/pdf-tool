# CLAUDE.md — Gerenciador de PDF e Word

## Propósito do projeto

Aplicação desktop em Python com interface dark mode (CustomTkinter) para manipular PDFs e documentos Word.
Dez ferramentas em abas: dividir, juntar, converter PDF↔Word, organizar páginas, compactar, PDF→imagem,
imagem→PDF, proteger com senha, desbloquear e girar.

Projeto de portfólio. Nasceu **Windows-first**; hoje roda também em Linux (ver "Diferenças por plataforma").

---

## Fonte da verdade

**O estado real do sistema é o código.** Nunca presuma que uma aba, um widget ou uma dependência existe —
verifique `tabs/`, `widgets.py` e `requirements.txt`.

---

## Stack

- **Python 3.10+** (o ambiente local está em 3.14)
- **CustomTkinter** — widgets dark mode sobre Tkinter
- **pypdf** — leitura, escrita, divisão, rotação, criptografia
- **pymupdf** — renderização de miniaturas, compactação, PDF↔imagem
- **pdf2docx** — PDF → Word
- **docx2pdf** — Word → PDF **só no Windows/macOS** (automatiza o Microsoft Word); no Linux o backend é o
  LibreOffice headless
- **cryptography** — PDFs com AES
- **pytest** — suíte da lógica pura, sem abrir janela

### Estrutura

```
pdf-tool/
├── pdf_tool.py          # ponto de entrada: classe PDFToolApp, tab bar customizada, status bar
├── theme.py             # ÚNICA fonte de cor, fonte e espaçamento (paleta dark)
├── widgets.py           # componentes reutilizáveis: section_title(), DropZone, ThumbnailGrid
├── docx_convert.py      # Word → PDF multiplataforma (Word no Windows/macOS, LibreOffice no Linux)
├── constants.py         # ⚠️ CÓDIGO MORTO — paleta clara, não importada por ninguém
├── requirements.txt     # dependências para rodar o app
├── requirements-dev.txt # + pytest (desenvolvimento)
├── pytest.ini           # pythonpath=., testpaths=tests
├── tests/
│   └── test_docx_convert.py   # escolha de motor por plataforma, erro acionável, mover o PDF gerado
└── tabs/
    ├── tab_split.py         tab_merge.py        tab_convert.py     tab_organize.py
    ├── tab_compress.py      tab_pdf_to_image.py tab_image_to_pdf.py
    └── tab_protect.py       tab_unlock.py       tab_rotate.py
```

---

## Arquitetura da UI

`PDFToolApp` monta um grid de 5 linhas: header(0) · barra de abas principal(1) · barra de sub-abas(2) ·
conteúdo(3, expansível) · status bar(4).

**Todas as 10 abas são instanciadas no boot**, em `_build_content()`. Cada página fica empilhada com
`place()` ocupando 100% da área, e a troca de aba é `tkraise()` — não há criação sob demanda.

> **Consequência:** trabalho pesado no `__init__` de uma aba atrasa a abertura do app inteiro. Construtor
> de aba deve só montar widgets.

**Contrato de toda aba** — a classe recebe sempre os mesmos três argumentos:

```python
TabClass(card, set_status=self._set_status, root=self.root)
```

- `parent` — o frame-card onde a aba desenha
- `set_status` — callback para escrever na status bar (`self.set_status("⏳ Processando…")`)
- `root` — a janela, necessária para `root.after(...)` ao voltar de uma thread

Aba nova precisa ser registrada em **dois lugares**: `TAB_CLASSES` e a lista `MAIN_TABS` ou `SUB_TABS`.
Só em `TAB_CLASSES` = a página é construída mas nenhum botão a alcança.

---

## Trabalho pesado e a thread da UI

Tkinter é single-threaded: qualquer operação longa no callback **congela a janela**.

O padrão do repo (ver `tabs/tab_convert.py`):

```python
def task():
    try:
        ...  # trabalho pesado
        self.root.after(0, lambda: self._done(save))
    except Exception as e:
        self.root.after(0, lambda: self._error(str(e)))

threading.Thread(target=task, daemon=True).start()
```

**Nunca toque em widget dentro da thread** — volte pela `root.after(0, ...)`.

⚠️ **Inconsistência atual:** só 4 das 10 abas usam thread (`convert`, `compress`, `pdf_to_image`,
`image_to_pdf`). As outras (`merge`, `split`, `protect`, `unlock`, `rotate`, `organize`) rodam na UI
thread — com arquivo grande, a janela trava até terminar. Não é bug reportado, mas é a origem provável de
qualquer "travou" que aparecer.

---

## Estilo

`theme.py` é a **única** fonte de cores, fontes e espaçamentos. Nunca escreva hex literal numa aba —
importe `theme as T` e use a constante. Se falta uma cor, acrescente lá.

⚠️ Dois detalhes conhecidos:

1. **`constants.py` é código morto.** Define uma paleta **clara** (`BG = "#F4F6F9"`) e não é importado por
   nenhum arquivo. Ignore-o; se for mexer, o candidato certo é apagá-lo.
2. **`FONT_FAMILY = "Segoe UI"` não existe no Linux.** O Tk cai num fallback silencioso, então o app é
   funcional mas visualmente diferente do Windows. Não é erro — é esperado.

---

## Diferenças por plataforma

| Recurso | Windows / macOS | Linux |
|---|---|---|
| Word → PDF | `docx2pdf` (precisa do Word instalado) | LibreOffice headless (`sudo pacman -S libreoffice-fresh`) |
| Fonte da UI | Segoe UI | fallback do sistema |
| Instalação da dep | `docx2pdf` instalado | pulado pelo marker `sys_platform != "linux"` |

A escolha do backend é automática em `docx_convert.py` (`docx_to_pdf()`), que levanta `ConversionError`
com mensagem legível quando o motor da plataforma não está disponível. **Não importe `docx2pdf`
diretamente numa aba** — sempre passe por `docx_convert`.

---

## Regras de desenvolvimento

- **Não assuma que uma aba existe** — verifique `TAB_CLASSES` em `pdf_tool.py`.
- **Não escreva cor/fonte literal** — tudo em `theme.py`.
- **Operação longa vai para thread** + `root.after(0, ...)` para voltar à UI.
- **Nunca deixe exceção subir sem feedback.** O usuário precisa ver `messagebox` ou status bar; falha
  silenciosa numa ferramenta de arquivo é péssima (ele acha que salvou).
- **Não sobrescreva arquivo de entrada.** As ferramentas geram cópia; manter isso é regra de segurança
  de dados do usuário.
- **Dependência nova** → adicionar ao `requirements.txt` com versão fixada, e conferir se funciona nas
  três plataformas (foi exatamente o problema do `docx2pdf`).

---

## Regra inegociável: SDD + BDD + TDD

Nenhum código de produção é escrito sem spec (SDD) → comportamento (BDD) → teste vermelho (TDD).
Sem exceções, mesmo em mudança pequena.

### 1. SDD — a spec mora no topo do arquivo de teste

Cabeçalho explicando **qual é o contrato**, **por que existe** (o bug ou a decisão que o originou) e
**o que é regra de negócio**. Modelo: `tests/test_docx_convert.py`.

### 2. BDD — comportamento, não implementação

`class Test<CenárioDeNegócio>` → `def test_deve_<resultado>_quando_<condição>`, em português, na
linguagem da operação (converter, proteger, dividir).

### 3. TDD — Red → Green → Refactor

Escreva o teste, rode e **veja falhar**; só então escreva o mínimo para passar.

### O que testar (por prioridade)

| Prioridade | Alvo | Por quê |
|---|---|---|
| 🔴 Alta | Lógica de arquivo pura | `docx_convert` (escolha de motor, mover o PDF gerado, erro acionável) |
| 🔴 Alta | Integridade do arquivo | destino ≠ origem, escrita parcial, sobrescrita |
| 🟡 Média | Regras extraídas das abas | validação de faixa de páginas, ordem de junção, nível de compactação |
| 🟢 Baixa | Widget | Tkinter não é testável sem display — valida-se à mão |

⚠️ **O desafio deste repo:** as abas hoje misturam UI e processamento, então a lógica delas **não é
testável como está**. Ao mexer numa aba, o caminho é **extrair a regra para função pura** (recebe
caminhos e parâmetros, devolve resultado, sem tocar em widget) e testar essa função. Foi assim que o
`docx_convert.py` nasceu. Não é refatoração grande — é fazer a extração da parte que você já ia mexer.

**Mocks só para o externo** (LibreOffice/`subprocess`, sistema de arquivos quando necessário). Use
`tmp_path` do pytest para arquivo de verdade em vez de mockar `open`.

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest                       # suíte completa
pytest -k libreoffice        # um recorte
```

Depois de verde, **abra o app** — a suíte não cobre nenhuma aba de ponta a ponta.

---

### Convenção de commits

Conventional Commits: `feat(tabs): adiciona aba de assinatura`, `fix(convert): trata PDF protegido`,
`test(convert): cobre ausência do LibreOffice`, `refactor(theme): consolida paleta`.

---

## Regras gerais

- **O código é a fonte da verdade.** Se algo aqui parecer inconsistente com o código, o código vence —
  e atualize este arquivo.
- Decisão técnica não-óbvia deve ser documentada (no commit e/ou aqui).
- **Não commite nem faça push sem ordem explícita.**
