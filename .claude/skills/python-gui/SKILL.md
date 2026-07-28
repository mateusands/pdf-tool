---
name: python-gui
description: Desenvolvimento do Gerenciador de PDF e Word (CustomTkinter + pypdf/pymupdf). Codifica as convenções do repo — contrato das abas, registro da ferramenta em GRUPOS, ícones sem emoji, thread + root.after para não travar a UI, theme.py como única fonte de estilo, backend de conversão por plataforma. Use ao criar ou alterar aba, widget ou fluxo de arquivo.
---

# Python GUI — Gerenciador de PDF e Word

Guia para qualquer mexida no app. Segue o `CLAUDE.md`: **não escreva estilo literal**, **operação longa
vai para thread**, **sem commit/push sem ordem**.

---

## Criar ou alterar uma aba

### Contrato (idêntico nas 10 abas)

```python
class TabAlgumaCoisa:
    def __init__(self, parent, set_status, root):
        self.set_status = set_status   # escreve na status bar
        self.root = root               # necessário para root.after(...)
        self._path = None
        self._build(parent)
```

- `set_status(mensagem, estado)` é o feedback contínuo; `messagebox` é o desfecho. O `estado`
  (`"info"`, `"ocupado"`, `"ok"`, `"erro"`) escolhe o ícone e a cor — **não escreva o símbolo no texto**.
- `root` existe para **voltar da thread**, não para virar janela nova.
- O `__init__` deve **só montar widgets**. Todas as abas são instanciadas no boot do app —
  trabalho pesado no construtor atrasa a abertura inteira.

### Registrar a ferramenta: um lugar só

A lista `GRUPOS` em `pdf_tool/app.py` é a fonte única — dela saem a sidebar, as páginas e o cabeçalho:

```python
Ferramenta("minha", "Rótulo curto", "scissors", "Título completo",
           "Frase que explica o que a ferramenta faz.", TabMinhaAba)
```

O terceiro campo é o nome do ícone no catálogo de `core/icons.py`; nome desconhecido levanta `ValueError`
na hora, em vez de desenhar um vazio silencioso. Precisa de um ícone que não existe? Baixe o SVG do
[Lucide](https://lucide.dev) e acrescente ao `DESENHOS`.

**A aba não desenha o próprio título** — quem mostra nome e descrição é o cabeçalho de contexto, a partir
da `Ferramenta`. Comece direto pelo conteúdo.

### Layout

As páginas são empilhadas com `place(relx=0, rely=0, relwidth=1, relheight=1)` e a troca é `tkraise()`.
Dentro da sua aba, use `pack`/`grid` normalmente — mas **não misture os dois no mesmo container**, que é
como se congela layout no Tkinter.

---

## Não travar a interface

Tkinter é single-threaded. Qualquer trabalho longo (ler PDF grande, renderizar miniatura, converter,
compactar) **congela a janela** se rodar no callback do botão.

Padrão do repo — **nenhuma aba cria thread na mão**. Copie de `pdf_tool/tabs/tab_convert.py`:

```python
self._btn.configure(state="disabled")          # trava: sem disparo duplo
self.set_status("Convertendo…", "ocupado")
executar_em_thread(
    self.root,
    lambda: pdf_io.dividir_pdf(path, save, paginas),   # trabalho pesado, SEM tocar em widget
    ao_terminar=lambda n: self._done(n, save),         # já volta na thread da UI
    ao_falhar=self._error,                             # recebe a MENSAGEM, pronta
)
```

Regras:

- **Nenhuma chamada de widget dentro da thread.** O `executar_em_thread` já volta pelo `root.after(0, …)`.
- **Desabilite o botão antes**, reabilite no `_done`/`_error` — senão o usuário dispara duas operações
  concorrentes sobre o mesmo arquivo.
- Cuidado com `lambda` capturando variável de laço: use argumento default (`lambda p=path: ...`).

⚠️ **Nunca escreva `lambda: self._error(str(e))` dentro de um `except`.** O Python apaga o nome `e` ao sair
do bloco e o lambda só roda depois, no `after` — aí `str(e)` levanta `NameError`, o erro morre dentro do
callback do Tk e o usuário **não vê aviso nenhum**. Foi esse o bug que fez as quatro abas com thread nunca
mostrarem erro. É exatamente por isso que o `background.py` existe: ele extrai a mensagem dentro do
`except` e passa por valor. Use-o.

---

## Estilo: `pdf_tool/theme.py` e nada mais

```python
from .. import theme as T          # dentro de pdf_tool/tabs/
ctk.CTkLabel(parent, text="...", font=T.FONT_BODY, text_color=T.MUTED)
```

- **Nunca escreva hex literal** numa aba. Faltou cor? Acrescente em `pdf_tool/theme.py`.
- Espaçamento também vem de lá: `T.PAD_S`, `T.PAD_M`, `T.PAD_L`, `T.RADIUS`.
- **Elevação tonal:** no escuro a profundidade vem da luminância, não de sombra. `SURFACE_1..4` sobem
  ~6% cada — use o nível da altura real do elemento (sidebar < card < campo dentro do card).
- **`ACCENT` preenche, `ACCENT_TEXT` escreve.** O azul de preenchimento não tem contraste suficiente
  quando é o próprio texto sobre fundo escuro. Vale igual para `SUCCESS`/`SUCCESS_TEXT` e
  `DANGER`/`DANGER_TEXT`.
- ⚠️ **Não existe letra mais branca que branco.** Faltou contraste num botão? Escureça o
  **preenchimento**. `pytest -k contraste` julga cada par do tema pelo AA (4,5:1) — rode antes de
  confiar no olho.
- **Botão vem de `widgets.botao()`**, não de `ctk.CTkButton` cru: o CustomTkinter desabilitado escurece
  só a letra e mantém o fundo colorido (1,29:1 sobre o azul, ilegível). A classe `Botao` repinta fundo,
  texto e ícone juntos.
- **As tuplas `T.FONT_*` só valem depois de `T.configurar_fontes(root)`**, que roda no boot e escolhe a
  família instalada da plataforma. Não as leia em argumento default de função — isso é avaliado na
  importação, antes da janela existir.

## Widgets reutilizáveis (`pdf_tool/widgets.py`)

- `icone(nome, tamanho, cor)` — ícone Lucide como `PhotoImage`. **Nada de emoji na interface**; há teste
  que varre o pacote e falha se algum voltar.
- `botao(parent, texto, command=…, variante=…, nome_do_icone=…)` — variantes `primario`, `sucesso`,
  `perigo`, `secundario`, `fantasma`.
- `GrupoPills(parent, opcoes, valor_inicial=…, ao_mudar=…)` — escolha única (nível, formato, DPI, ângulo).
- `CampoSenha(parent, rotulo)` — campo com mostrar/ocultar; `.valor()` e `.limpar()`.
- `criar_area_rolavel(parent, **pack_kw)` — devolve `(canvas, conteudo)` com rolagem já ligada.
- `estado_vazio(parent, icone, texto)` — o "nada aqui ainda" centralizado.
- `DropZone(parent, icon=…, text=…, subtitle=…, command=…, on_clear=…)` — área de seleção de arquivo,
  com `set_file(nome)` para refletir a escolha.
- `ThumbnailGrid(parent, cols, thumb_w, thumb_h)` — grade de miniaturas com seleção.
- `ligar_rolagem(canvas)` — **obrigatório** em área rolável nova. Escutar só `<MouseWheel>` deixa a roda
  do mouse morta no Linux (X11 manda `<Button-4>`/`<Button-5>`).

Precisa de um componente novo em mais de uma aba? Ele vai para `pdf_tool/widgets.py`, não duplicado.

---

## Conversão Word → PDF: sempre por `docx_convert`

```python
from ..core.docx_convert import docx_to_pdf, ConversionError   # dentro de pdf_tool/tabs/
```

- **Não importe `docx2pdf` diretamente.** Ele só funciona no Windows/macOS (automatiza o Word) e nem
  chega a ser instalado no Linux — o `requirements.txt` tem o marker `sys_platform != "linux"`.
- `docx_to_pdf(src, dest)` escolhe o backend pela plataforma: Word no Windows/macOS, **LibreOffice
  headless** no Linux.
- Ele levanta `ConversionError` com **mensagem já pronta para o usuário** (inclusive o comando de
  instalação do LibreOffice). Deixe essa mensagem chegar ao `messagebox` — não a substitua por um texto
  genérico.

---

## Segurança de dados do usuário

Esta ferramenta mexe em arquivos que a pessoa não tem cópia:

- **Toda escrita de PDF passa por `pdf_tool/core/pdf_io.py`.** É lá que moram as duas garantias que
  protegem o usuário: destino ≠ origem (comparado por `realpath`, então caminho relativo e link
  simbólico também são pegos) e gravação atômica (`.part` → `os.replace`, então falha no meio não deixa
  arquivo truncado). Aba que chama `PdfWriter`/`doc.save()` direto fura as duas.
- **Operação que escreve fora do `pdf_io`** (montar PDF a partir de imagens, exportar para pasta) precisa
  chamar `pdf_io.validar_destino(destino, *origens)` na mão.
- **Nunca falhe em silêncio.** Exceção engolida faz o usuário achar que salvou. Todo caminho de erro
  termina em `messagebox.showerror` **e** status bar com estado `"erro"`.
- **Senha de PDF** (`tab_protect`/`tab_unlock`) não vai para log, título de janela nem status bar.

---

## SDD + BDD + TDD (obrigatório) + validar verde

**Ordem: spec → comportamento → teste falhando → código.** Detalhe completo no `CLAUDE.md`.

- **SDD:** cabeçalho do teste explica contrato e porquê. Modelo: `tests/test_docx_convert.py`.
- **BDD:** `class Test<Cenário>` → `def test_deve_<resultado>_quando_<condição>`, em português.
- **TDD:** Red (roda e **falha**) → Green → Refactor.

### O desafio deste repo: a aba não é testável como está

A escrita de arquivo, a thread, a reordenação, os ícones e a escolha de fonte já saíram para
`pdf_tool/core/`. Mas ainda sobra regra dentro de aba (montagem do PDF a partir de imagens, DPI e
formato do PDF → imagem), e isso não é testável sem display. **O caminho não é heroico — é extrair
enquanto você já está mexendo:**

```python
# ❌ regra presa dentro da aba, intestável
def _convert(self):
    save = filedialog.asksaveasfilename(...)
    paginas = [int(p) for p in self._entry.get().split(",")]   # ← a regra está aqui
    ...

# ✅ regra extraída, testável direto
# em um módulo próprio (ex.: paginas.py)
def parse_intervalo(texto: str, total: int) -> list[int]:
    """Interpreta "1,3-5" como páginas. Levanta ValueError com msg ao usuário."""
    ...
```

Foi exatamente assim que o `pdf_tool/core/docx_convert.py` nasceu, e ele é o único pedaço deste projeto hoje coberto
por teste. Ao mexer numa aba, extraia a regra que você ia tocar — não o arquivo inteiro.

**Mocks só para o externo** (LibreOffice via `subprocess`). Para arquivo, use `tmp_path` do pytest e
escreva de verdade — mockar `open` num projeto que manipula PDF esconde justamente o que importa.

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest                   # suíte completa
pytest -k libreoffice    # um recorte
```

**Verde não basta.** Nenhuma aba é coberta de ponta a ponta. Depois de passar, abra o app e exercite —
roteiro em `/rodar-local`, com **PDF real** (arquivo grande é o que revela travamento de UI).

Seja honesto no relato: se testou dividir e não testou juntar, diga isso.
