---
name: codereview
description: Code review sênior das últimas mudanças do Gerenciador de PDF e Word, focado em integridade dos arquivos do usuário, thread da UI, contrato das abas e manutenibilidade. Apenas reporta problemas com arquivo/linha e a refatoração sugerida — não aplica correções.
---

# Code Review Sênior — Gerenciador de PDF e Word

Atue como Engenheiro Sênior e revise criticamente as **últimas mudanças deste repositório**.

## Como identificar o que revisar (nesta ordem)

1. Working tree: `git status` + `git diff` + arquivos novos relevantes.
2. Se limpo, os últimos commits da branch (`git log` + `git show`).
3. **Leia o arquivo inteiro** quando o diff não der contexto. Aba nova = leia uma aba existente do mesmo
   tipo para comparar o padrão.

## Pilar 1 — Integridade dos dados do usuário (prioridade máxima aqui)

Esta ferramenta escreve em arquivos que a pessoa pode não ter backup. Um bug aqui destrói trabalho alheio.

- **Sobrescrita do arquivo de entrada:** o diff grava no mesmo caminho que leu? Toda operação deve gerar
  destino novo via `asksaveasfilename`. Reporte no topo se encontrar.
- **Escrita parcial:** se a operação falhar no meio, sobra arquivo truncado no lugar de um válido? O
  padrão seguro é escrever em temporário e mover no fim (é o que `pdf_tool/core/docx_convert.py` faz).
- **Falha silenciosa:** `except` que só faz `pass`/`print` sem `messagebox` nem status bar. O usuário
  conclui que salvou. Isso é mais grave aqui do que numa app comum.
- **Senha vazando:** senha de PDF em log, `print`, título de janela, status bar ou mensagem de erro.
- **Caminho vindo do usuário** concatenado sem `os.path`/`pathlib`; nome de arquivo com `..` ou separador
  ao montar destino em lote.

## Pilar 2 — Thread da UI

- **Operação longa no callback do botão** (ler PDF grande, renderizar miniatura, converter, compactar)
  sem `threading.Thread`? A janela congela. Compare com `pdf_tool/tabs/tab_convert.py`.
- **Widget tocado dentro da thread** — `configure`, `insert`, `set` fora do `root.after(0, ...)`. Isso é
  travamento intermitente, o pior tipo de bug para reproduzir. Verifique linha a linha o corpo de `task()`.
- **Botão não desabilitado** durante o processamento → duas operações concorrentes no mesmo arquivo.
- **`daemon=True` ausente** → a thread segura o processo ao fechar a janela.
- **`lambda` capturando variável de laço** sem argumento default.
- **Estado preso:** o caminho de erro reabilita o botão e reseta a status bar? Cancelar o diálogo de
  salvar no meio deixa a aba utilizável?

## Pilar 3 — Contrato das abas

- **Ferramenta nova na lista `GRUPOS`** de `pdf_tool/app.py`? É a fonte única da sidebar e do cabeçalho.
- **Emoji na interface?** Proibido — `widgets.icone()` resolve. O teste varre o pacote, mas revise mesmo assim.
- **Símbolo escrito no texto do status** (`"✓ Salvo"`)? O ícone e a cor saem do `estado`:
  `set_status("Salvo", "ok")`.
- **Assinatura correta:** `__init__(self, parent, set_status, root)`. Divergir quebra o `_montar_conteudo()`.
- **`__init__` faz trabalho pesado?** Todas as abas são instanciadas no boot — isso atrasa a abertura do
  app inteiro. Construtor só monta widgets.
- **`pack` e `grid` misturados no mesmo container** — congela o layout no Tkinter.

## Pilar 4 — Estilo e reuso

- **Hex literal ou fonte literal** numa aba? Tudo vem de `pdf_tool/theme.py` (`T.PRIMARY`, `T.FONT_BODY`, `T.PAD_M`).
- **Aba escrevendo PDF na mão?** `PdfWriter`/`doc.save()` dentro de `pdf_tool/tabs/` fura a checagem de
  destino e a gravação atômica. Toda escrita passa por `pdf_tool/core/pdf_io.py` — reporte.
- **Componente duplicado** que já existe em `pdf_tool/widgets.py` (`botao`, `GrupoPills`, `CampoSenha`,
  `DropZone`, `ThumbnailGrid`, `criar_area_rolavel`, `estado_vazio`).
- **`docx2pdf` importado diretamente** numa aba? Tem que passar por `docx_convert.docx_to_pdf` — o import
  direto quebra no Linux, onde o pacote nem é instalado.

## Pilar 5 — Robustez com arquivos reais

- **PDF protegido por senha** faz `pypdf`/`pymupdf` levantarem na **abertura**. O caminho novo trata?
- **PDF corrompido / não-PDF renomeado** — erro tratado ou stack trace no terminal?
- **PDF de 500 páginas** — a aba tenta renderizar tudo de uma vez? Miniatura é o gargalo clássico.
- **Arquivo sem permissão de escrita** no destino.
- **Dependência nova** no `requirements.txt`: versão fixada? funciona nas três plataformas? precisa de
  marker de ambiente (foi o caso do `docx2pdf`)? Reporte pacote + licença.

## Pilar 6 — TDD (obrigatório neste repo)

- **Código de produção novo sem teste correspondente?** Viola a regra inegociável do `CLAUDE.md`.
  Reporte — é achado de review, não detalhe de estilo.
- **Regra de negócio nova enterrada dentro da aba**, misturada com widget, quando poderia ter sido
  extraída para função pura e testada. É a dívida estrutural deste repo: aponte a extração possível.
- **O teste tem cabeçalho SDD** explicando contrato e porquê, ou é assert solto sem contexto?
- **O nome descreve comportamento** (`test_deve_<resultado>_quando_<condição>`) ou detalhe interno?
- **Mock demais:** teste que mocka `open`/`Path` num projeto que manipula arquivo esconde justamente o
  que importa. `tmp_path` com arquivo real é melhor.
- **Teste que não falharia** se a implementação fosse removida.

## Pilar 7 — Manutenibilidade

- Clean Code: função que faz seleção + processamento + UI ao mesmo tempo; complexidade desnecessária.
- Nomes e comentários seguindo o arquivo (o repo mistura inglês e português — siga o vizinho, não
  padronize por conta própria).
- Duplicação entre abas que pede extração para `pdf_tool/widgets.py`.

## Formato da resposta

- Nada de micro-otimização irrelevante.
- Para cada problema: **arquivo e linha**, impacto, e o código refatorado. Ordene por severidade —
  integridade de arquivo do usuário primeiro, depois thread da UI.
- **Apenas revise e reporte. Não aplique as correções** sem ordem explícita.
