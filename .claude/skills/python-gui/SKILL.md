---
name: python-gui
description: Desenvolvimento do Gerenciador de PDF e Word (CustomTkinter + pypdf/pymupdf). Codifica as convenções do repo — contrato das abas, registro em dois lugares, thread + root.after para não travar a UI, theme.py como única fonte de estilo, backend de conversão por plataforma. Use ao criar ou alterar aba, widget ou fluxo de arquivo.
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

- `set_status("⏳  Processando…")` é o feedback contínuo; `messagebox` é o desfecho.
- `root` existe para **voltar da thread**, não para virar janela nova.
- O `__init__` deve **só montar widgets**. Todas as abas são instanciadas no boot do app —
  trabalho pesado no construtor atrasa a abertura inteira.

### Registrar em DOIS lugares (senão a aba não aparece)

Em `pdf_tool.py`:

1. `TAB_CLASSES["minha_aba"] = TabMinhaAba` — constrói a página
2. `MAIN_TABS` **ou** `SUB_TABS` — cria o botão que a alcança

Só no `TAB_CLASSES`: a página é construída, ocupa memória e **nenhum botão a alcança**. Sem erro nenhum —
o sintoma é "a aba não existe". É a pegadinha número um deste arquivo.

### Layout

As páginas são empilhadas com `place(relx=0, rely=0, relwidth=1, relheight=1)` e a troca é `tkraise()`.
Dentro da sua aba, use `pack`/`grid` normalmente — mas **não misture os dois no mesmo container**, que é
como se congela layout no Tkinter.

---

## Não travar a interface

Tkinter é single-threaded. Qualquer trabalho longo (ler PDF grande, renderizar miniatura, converter,
compactar) **congela a janela** se rodar no callback do botão.

Padrão do repo — copie de `tabs/tab_convert.py`:

```python
self._btn.configure(state="disabled")
self.set_status("⏳  Convertendo…")

def task():
    try:
        ...  # trabalho pesado, SEM tocar em widget
        self.root.after(0, lambda: self._done(save))
    except Exception as e:
        self.root.after(0, lambda: self._error(str(e)))

threading.Thread(target=task, daemon=True).start()
```

Regras:

- **Nenhuma chamada de widget dentro da thread.** `configure`, `insert`, `set` só depois do `root.after(0, ...)`.
  Violar isso gera travamento intermitente, difícil de reproduzir.
- **Desabilite o botão antes**, reabilite no `_done`/`_error` — senão o usuário dispara duas operações
  concorrentes sobre o mesmo arquivo.
- **`daemon=True`** para a thread não segurar o processo ao fechar a janela.
- Cuidado com `lambda` capturando variável de laço: use argumento default (`lambda p=path: ...`).

⚠️ **Só 4 das 10 abas fazem isso hoje** (`convert`, `compress`, `pdf_to_image`, `image_to_pdf`). As demais
rodam na UI thread. Se você mexer numa aba sem thread e ela ficar lenta, **migre para o padrão acima** em
vez de aceitar o congelamento.

---

## Estilo: `theme.py` e nada mais

```python
import theme as T
ctk.CTkLabel(parent, text="...", font=T.FONT_BODY, text_color=T.MUTED)
```

- **Nunca escreva hex literal** numa aba. Faltou cor? Acrescente em `theme.py`.
- Espaçamento também vem de lá: `T.PAD_S`, `T.PAD_M`, `T.PAD_L`, `T.RADIUS`.
- **`constants.py` é código morto** — paleta clara, não importada por ninguém. Não importe dele por engano;
  o nome parece certo e o conteúdo está errado.
- **`FONT_FAMILY = "Segoe UI"` não existe no Linux** — o Tk faz fallback silencioso. Diferença visual
  entre plataformas é esperada, não é bug a caçar.

## Widgets reutilizáveis (`widgets.py`)

- `section_title(parent, title, subtitle="")` — cabeçalho padrão de toda aba. Use, não recrie.
- `DropZone(parent, icon=…, text=…, subtitle=…, command=…, on_clear=…)` — área de seleção de arquivo,
  com `set_file(nome)` para refletir a escolha.
- `ThumbnailGrid(parent, cols, thumb_w, thumb_h)` — grade de miniaturas com seleção.

Precisa de um componente novo em mais de uma aba? Ele vai para `widgets.py`, não duplicado.

---

## Conversão Word → PDF: sempre por `docx_convert`

```python
from docx_convert import docx_to_pdf, ConversionError
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

- **Nunca sobrescreva o arquivo de entrada.** Todas as operações geram destino novo, escolhido por
  `asksaveasfilename`. Mantenha isso.
- **Nunca falhe em silêncio.** Exceção engolida faz o usuário achar que salvou. Todo caminho de erro
  termina em `messagebox.showerror` **e** status bar.
- **Senha de PDF** (`tab_protect`/`tab_unlock`) não vai para log, título de janela nem status bar.
- **Escreva primeiro, confirme depois:** se a operação falhar no meio, não deixe arquivo truncado no
  lugar de um válido.

---

## SDD + BDD + TDD (obrigatório) + validar verde

**Ordem: spec → comportamento → teste falhando → código.** Detalhe completo no `CLAUDE.md`.

- **SDD:** cabeçalho do teste explica contrato e porquê. Modelo: `tests/test_docx_convert.py`.
- **BDD:** `class Test<Cenário>` → `def test_deve_<resultado>_quando_<condição>`, em português.
- **TDD:** Red (roda e **falha**) → Green → Refactor.

### O desafio deste repo: a aba não é testável como está

As 10 abas misturam seleção de arquivo, processamento e widget num método só. Não dá para testar isso
sem display. **O caminho não é heroico — é extrair enquanto você já está mexendo:**

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

Foi exatamente assim que o `docx_convert.py` nasceu, e ele é o único pedaço deste projeto hoje coberto
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
