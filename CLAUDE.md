# CLAUDE.md — Gerenciador de PDF e Word

## Propósito do projeto

Aplicação desktop em Python com interface dark mode (CustomTkinter) para manipular PDFs e documentos Word.
Dez ferramentas numa sidebar agrupada: dividir, juntar, organizar páginas, girar, converter PDF↔Word,
PDF→imagem, imagem→PDF, compactar, proteger com senha e desbloquear.

Projeto de portfólio. Nasceu **Windows-first**; hoje roda também em Linux (ver "Diferenças por plataforma").

---

## Fonte da verdade

**O estado real do sistema é o código.** Nunca presuma que uma ferramenta, um widget, um ícone ou uma
dependência existe — verifique `pdf_tool/app.py` (lista `GRUPOS`), `pdf_tool/widgets.py`,
`pdf_tool/core/icons.py` e `requirements.txt`.

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
├── main.py              # ponto de entrada — só isto é executável (`python main.py`)
├── pdf_tool/            # o app
│   ├── app.py           # classe PDFToolApp: tab bar customizada, status bar
│   ├── theme.py         # ÚNICA fonte de cor, fonte e espaçamento (paleta dark)
│   ├── widgets.py       # icone(), botao(), GrupoPills, CampoSenha, DropZone, grade…
│   ├── core/            # ⚠️ camada pura: NÃO pode importar Tkinter/CustomTkinter
│   │   ├── pdf_io.py         # TODA escrita de PDF do app
│   │   ├── docx_convert.py   # Word → PDF (Word no Windows/macOS, LibreOffice no Linux)
│   │   ├── background.py     # ÚNICO lugar que cria thread; volta pela root.after
│   │   ├── icons.py          # catálogo Lucide (ISC), rasterização SVG, ícone da janela
│   │   ├── fonts.py          # escolha da família de fonte por plataforma
│   │   ├── contrast.py       # razão de contraste WCAG — decide cor com número
│   │   └── reorder.py        # reordenação de listas (mover item, soltar no alvo)
│   └── tabs/
│       ├── tab_split.py       tab_merge.py        tab_convert.py     tab_organize.py
│       ├── tab_compress.py    tab_pdf_to_image.py tab_image_to_pdf.py
│       └── tab_protect.py     tab_unlock.py       tab_rotate.py
├── requirements.txt     # dependências para rodar o app (todas fixadas)
├── requirements-dev.txt # + pytest (desenvolvimento)
├── pytest.ini           # pythonpath=., testpaths=tests
└── tests/
    ├── test_pdf_io.py         # destino ≠ origem, escrita atômica, AES-256, senha, operações
    ├── test_docx_convert.py   # escolha de motor por plataforma, erro acionável, mover o PDF gerado
    ├── test_background.py     # a mensagem de erro da thread chega mesmo à tela
    ├── test_reorder.py        # mover mantendo a seleção, soltar no cartão certo
    ├── test_icons.py          # render, ícone da janela, e o varredor de emoji
    ├── test_contrast.py       # cada par de cor do tema contra o AA (4,5:1)
    ├── test_fonts.py          # escolha da família por plataforma
    └── test_scroll.py         # roda do mouse no X11, Windows e macOS
```

**A regra de dependência é uma só e vale mais que o desenho da pasta:** `tabs/` e `widgets.py` importam
de `core/`; `core/` não importa de ninguém acima dele e não conhece Tkinter. É isso que faz a suíte rodar
sem display. Import dentro do pacote é relativo (`from ..core import pdf_io`).

---

## Arquitetura da UI

`PDFToolApp` monta um grid de 2 colunas: sidebar(col 0, ocupa as duas linhas) · cabeçalho de
contexto(col 1, linha 0) · conteúdo(col 1, linha 1, expansível) · status bar(linha 2, cruza tudo).

**Todas as 10 abas são instanciadas no boot**, em `_montar_conteudo()`. Cada página fica empilhada com
`place()` ocupando 100% da área, e a troca é `tkraise()` — não há criação sob demanda.

> **Consequência:** trabalho pesado no `__init__` de uma aba atrasa a abertura do app inteiro. Construtor
> de aba deve só montar widgets.

**Registro de uma ferramenta nova: um lugar só.** A lista `GRUPOS` em `pdf_tool/app.py` é a fonte única —
dela saem a sidebar, as páginas e o texto do cabeçalho:

```python
Ferramenta("split", "Dividir", "scissors", "Dividir PDF",
           "Escolha as páginas que deseja extrair…", TabSplit)
#          chave    rótulo     ícone       título     descrição              classe
```

O `icone` precisa existir no catálogo de `core/icons.py` — nome desconhecido levanta `ValueError` na hora.

**Contrato de toda aba** — a classe recebe sempre os mesmos três argumentos:

```python
TabClass(card, set_status=self._set_status, root=self.root)
```

- `parent` — o frame-card onde a aba desenha
- `set_status` — callback da status bar: `set_status(mensagem, estado)`, com `estado` em
  `"info"` (padrão), `"ocupado"`, `"ok"` ou `"erro"`. É o estado que escolhe o ícone e a cor —
  **não escreva o símbolo no texto**
- `root` — a janela, necessária para `root.after(...)` ao voltar de uma thread

A aba **não desenha seu próprio título**: quem mostra o nome e a descrição é o cabeçalho de contexto,
a partir da `Ferramenta`. A aba começa direto pelo conteúdo.

---

## Trabalho pesado e a thread da UI

Tkinter é single-threaded: qualquer operação longa no callback **congela a janela**.

O padrão do repo é **um só**, e mora em `background.py`. Nenhuma aba cria thread na mão:

```python
self._btn.configure(state="disabled")          # trava: sem disparo duplo
self.set_status("Salvando…", "ocupado")        # o estado escolhe ícone e cor
executar_em_thread(
    self.root,
    lambda: pdf_io.dividir_pdf(path, save, paginas),   # trabalho pesado
    ao_terminar=lambda n: self._done(n, save),         # volta na thread da UI
    ao_falhar=self._error,                             # recebe a MENSAGEM, já pronta
)
```

**Nunca toque em widget dentro da thread** — o `executar_em_thread` já volta pela `root.after(0, ...)`.

⚠️ **Nunca escreva `lambda: self._error(str(e))` dentro de um `except`.** O Python apaga o nome `e` ao
sair do bloco, e o lambda só roda depois, no `after` — aí `str(e)` levanta `NameError` e o erro morre
dentro do callback do Tk, sem messagebox e sem status. Foi exatamente esse o bug que fez as quatro abas
com thread (`convert`, `compress`, `pdf_to_image`, `image_to_pdf`) nunca mostrarem erro nenhum. O
`background.py` extrai a mensagem dentro do `except` e passa por valor — por isso ele existe.

Hoje as 10 abas rodam o trabalho pesado fora da UI thread.

---

## Estilo

`theme.py` é a **única** fonte de cores, fontes e espaçamentos. Nunca escreva hex literal numa aba —
importe `theme as T` e use a constante. Se falta uma cor, acrescente lá.

**Elevação tonal:** no escuro a profundidade vem da luminância, não de sombra. `SURFACE_1..4` sobem
~6% cada; use o nível que corresponde à altura real do elemento (sidebar < card < campo dentro do card).

**Preenchimento e texto são papéis opostos, com tons diferentes.** `ACCENT`/`SUCCESS`/`DANGER` são
fundo, julgados contra o texto **branco** por cima. `ACCENT_TEXT`/`*_TEXT` são o próprio texto, julgados
contra a superfície escura. Trocar um pelo outro é o erro clássico aqui.

⚠️ **Não existe letra mais branca que branco.** Quando faltar contraste num botão, quem escurece é o
**preenchimento**, não o texto. Foi assim que um botão azul ficou ilegível: `#3B82F6` dava só 3,68:1 com
branco (o mínimo AA é 4,5:1), e o `#2563EB` de hoje dá 5,17:1. `tests/test_contrast.py` verifica cada par
do tema — mexeu em cor, rode `pytest -k contraste` antes de olhar na tela.

⚠️ **Estado desabilitado apaga o preenchimento, não só o texto.** O default do CustomTkinter escurece
apenas a letra (`text_color_disabled` = `gray60`) e mantém o fundo colorido: em cima do azul isso dá
**1,29:1** — invisível, num botão que ainda parece clicável. A classe `widgets.Botao` repinta fundo,
texto e ícone juntos ao receber `configure(state=…)`; use-a em vez de `ctk.CTkButton` cru.

**Fontes são resolvidas em tempo de execução.** `T.configurar_fontes(root)` roda logo depois de criar a
janela e escolhe a primeira família instalada da lista da plataforma (`core/fonts.py`). Por isso as
tuplas `T.FONT_*` só valem **depois** disso — não as leia em argumento default de função, que é avaliado
na importação.

⚠️ **Nada de emoji na interface.** Emoji muda de desenho conforme o sistema, tem cor fixa (não acompanha
hover, foco nem estado desabilitado) e no Linux depende de fonte instalada. Use `widgets.icone(nome, …)`,
do catálogo Lucide. Há um teste que varre o pacote e falha se algum emoji voltar
(`tests/test_icons.py::TestInterfaceSemEmoji`).

**Componente novo em mais de uma aba vai para `widgets.py`**, nunca duplicado: já moram lá `icone()`,
`botao()`, `GrupoPills`, `CampoSenha`, `DropZone`, `ThumbnailGrid`, `criar_area_rolavel()` e
`estado_vazio()`.

---

## Diferenças por plataforma

| Recurso | Windows / macOS | Linux |
|---|---|---|
| Word → PDF | `docx2pdf` (precisa do Word instalado) | LibreOffice headless (`sudo pacman -S libreoffice-fresh`) |
| Fonte da UI | Segoe UI | fallback do sistema |
| Instalação da dep | `docx2pdf` instalado | pulado pelo marker `sys_platform != "linux"` |
| Roda do mouse | evento `<MouseWheel>` com `delta` | X11/Tk 8.6 manda `<Button-4>` / `<Button-5>` |

⚠️ **Área rolável nova usa `widgets.ligar_rolagem(canvas)`.** Escutar só `<MouseWheel>` deixa a roda do
mouse morta no Linux — foi o que aconteceu com as quatro áreas roláveis do app. E não basta ligar no
canvas: os cartões desenhados dentro dele são widgets próprios, e evento de mouse no Tk não sobe para o
pai, então a roda só funcionaria sobre o fundo vazio.

A escolha do backend é automática em `docx_convert.py` (`docx_to_pdf()`), que levanta `ConversionError`
com mensagem legível quando o motor da plataforma não está disponível. **Não importe `docx2pdf`
diretamente numa aba** — sempre passe por `docx_convert`.

---

## Regras de desenvolvimento

- **Não assuma que uma ferramenta existe** — verifique a lista `GRUPOS` em `pdf_tool/app.py`.
- **Não escreva cor/fonte literal** — tudo em `theme.py`.
- **Aba não escreve PDF.** Toda gravação passa por `pdf_io`, que é onde vivem a checagem de destino e a
  escrita atômica. Aba que chama `PdfWriter`/`doc.save()` direto fura as duas garantias.
- **Operação longa vai por `background.executar_em_thread`** — não crie `threading.Thread` na aba.
- **Nunca deixe exceção subir sem feedback.** O usuário precisa ver `messagebox` ou status bar; falha
  silenciosa numa ferramenta de arquivo é péssima (ele acha que salvou).
- **Não sobrescreva arquivo de entrada.** Garantido por `pdf_io.validar_destino()`: o destino é comparado
  com as origens por `realpath`, então caminho relativo e link simbólico também são pegos. É regra de
  segurança de dados do usuário, não detalhe de implementação.
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
| 🔴 Alta | Integridade do arquivo | destino ≠ origem, escrita parcial, sobrescrita (`pdf_io`) |
| 🔴 Alta | Lógica de arquivo pura | `docx_convert` (escolha de motor, erro acionável), senha, operações |
| 🔴 Alta | Caminho de erro | a mensagem CHEGA na tela? (`background`) — falha silenciosa é o pior defeito daqui |
| 🟡 Média | Regras extraídas das abas | reordenação, faixa de páginas, nível de compactação |
| 🟢 Baixa | Widget | Tkinter não é testável sem display — valida-se à mão |

⚠️ **O desafio deste repo:** as abas nasceram misturando UI e processamento. `pdf_io`, `background` e
`reorder` já extraíram a maior parte, mas ainda sobra regra dentro de aba (montagem do PDF a partir de
imagens, DPI e formato do PDF → imagem). Ao mexer numa aba, o caminho continua o mesmo: **extraia a
regra para função pura** (recebe caminhos e parâmetros, devolve resultado, sem tocar em widget) e teste
essa função. Não é refatoração grande — é fazer a extração da parte que você já ia mexer.

Depois de verde, **abra o app**: a suíte não cobre nenhuma aba de ponta a ponta. Vale dirigir os fluxos
com os diálogos instrumentados (`filedialog`/`messagebox` trocados por dublês) para confirmar que a caixa
certa aparece — foi assim que se confirmou que as sete falhas silenciosas passaram a avisar.

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
