---
name: finalizar-sessao
description: Encerra a sessão de trabalho no Gerenciador de PDF e Word — gera o relatório da sessão em .claude/sessions/ e atualiza o CLAUDE.md se algo que ele afirma mudou. Use ao final de cada sessão.
---

# Encerramento de Sessão — Gerenciador de PDF e Word

O objetivo agora **não é codar**, e sim consolidar o que a sessão mudou.

## 1. Relatório da sessão

- Crie `.claude/sessions/YYYY-MM-DD.md` (data de hoje). Se já existir arquivo com a data de hoje,
  **acrescente** uma seção em vez de sobrescrever.
- Conteúdo exigido:
  - **O que foi feito** — qual aba, qual widget, qual fluxo de arquivo.
  - **Decisões técnicas não-óbvias** — e o porquê.
  - **Validação manual** — quais abas você realmente exercitou e **com que tipo de arquivo** (1 página?
    50 páginas? protegido por senha?). Testar só com PDF de uma página não valida quase nada aqui.
  - **Pendências** — explícitas o bastante para retomar sem contexto.
  - **Estado do git** — branch, se ficou coisa não commitada.

> `.claude/sessions/` é **gitignorado** — caderno de bordo local, não documentação do repo.

## 2. Atualização do CLAUDE.md

Avalie se a sessão mudou algo que o `CLAUDE.md` afirma. Gatilhos:

- **Aba nova ou removida** — a estrutura de diretórios e a lista de ferramentas estão documentadas.
- **Aba migrada para thread** — o arquivo afirma hoje que **4 das 10** usam thread e lista quais. Se você
  migrou mais uma, **atualize o número e a lista**. É a afirmação mais fácil de ficar desatualizada.
- **`constants.py` apagado** — o arquivo o descreve como código morto em dois lugares.
- **Dependência nova ou removida** — a tabela de stack e a de diferenças por plataforma.
- **Mudança no contrato das abas** (assinatura do `__init__`, forma de registro).
- **Armadilha nova descoberta** — acrescente; é o conteúdo mais valioso do documento.

## 3. Validação final

**Primeiro a suíte** — e relate o resultado real:

```bash
pytest
```

Se houve código de produção nesta sessão, houve teste vermelho antes? Se não, a regra do `CLAUDE.md`
foi quebrada — registre no relatório em vez de esconder.

**Depois o app**, porque nenhuma aba é coberta de ponta a ponta:

```bash
.venv/bin/python pdf_tool.py
```

Use o roteiro de `/rodar-local`. Para qualquer mudança em fluxo de arquivo, confirme os três pontos
inegociáveis:

1. O **arquivo de entrada não foi alterado**
2. O erro chega ao usuário (messagebox **e** status bar), sem stack trace só no terminal
3. Cancelar o diálogo de salvar no meio não deixa a aba com botão travado

**Relate o que de fato testou.** Se exercitou 2 das 10 abas, diga 2.

## O que responder ao usuário

1. Caminho do relatório gerado.
2. Se o `CLAUDE.md` foi atualizado, e o que mudou (ou que nada foi necessário).
3. O que foi validado manualmente, com que arquivos, e o que ficou de fora.
4. **Não commite nem faça push** — só quando o dono mandar.
