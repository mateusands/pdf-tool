---
name: iniciar-sessao
description: Inicializa a sessão de trabalho no Gerenciador de PDF e Word — lê o CLAUDE.md, o estado do git e as pendências da última sessão, em modo somente leitura, e confirma o alinhamento de escopo antes de qualquer código. Use no começo de cada sessão.
---

# Inicialização de Sessão — Gerenciador de PDF e Word

App desktop em **Python + CustomTkinter**, 10 abas de manipulação de PDF/Word, com suíte `pytest` da
lógica pura (as abas em si ainda não são cobertas — ver o gate de TDD abaixo).
A fonte da verdade é **o código**.

Antes de qualquer ação, execute os passos de leitura abaixo:

1. **Leia o `CLAUDE.md` da raiz** — arquitetura da UI, contrato das abas, padrão de thread, e os dois
   detalhes que enganam: `constants.py` é código morto e `Segoe UI` não existe no Linux.

2. **Leia a última sessão**, se houver: `.claude/sessions/` (arquivo mais recente).

3. **Levante o estado real do git** (somente leitura):
   ```bash
   git status --short && git branch --show-current && git log --oneline -10
   ```

4. **Leia só o que a tarefa exige.** São ~1.000 linhas em `tabs/` — não carregue as 10 abas de uma vez.
   O caminho eficiente é: `pdf_tool.py` (como tudo se conecta) → `theme.py` → a aba alvo → uma aba
   vizinha do mesmo tipo, para pegar o padrão.

5. **MODO SOMENTE LEITURA:** é proibido alterar código, criar ou apagar arquivo nesta etapa.

## Gates que valem nesta sessão

Confirme explicitamente que estão ativos:

- **Integridade dos arquivos do usuário.** Nunca sobrescrever o arquivo de entrada; nunca falhar em
  silêncio. Esta ferramenta mexe em documento que a pessoa pode não ter backup.
- **Operação longa vai para thread** + `root.after(0, ...)` para voltar à UI. Nenhuma chamada de widget
  dentro da thread.
- **Estilo só de `theme.py`.** Nada de hex literal; `constants.py` é código morto e não deve ser importado.
- **Word → PDF só por `docx_convert.docx_to_pdf`.** Import direto de `docx2pdf` quebra no Linux.
- **Aba nova se registra em dois lugares** (`TAB_CLASSES` + `MAIN_TABS`/`SUB_TABS`).
- **SDD + BDD + TDD obrigatório** — spec no topo do teste → `test_deve_<resultado>_quando_<condição>` →
  teste vermelho → código. A suíte roda com `pytest`. Aba não é testável como está: ao mexer numa,
  **extraia a regra para função pura** e teste ela (foi assim que o `docx_convert.py` nasceu).
- **Verde não basta** — nenhuma aba é coberta de ponta a ponta. Validação real é abrir o app, pelo
  roteiro de `/rodar-local`, com arquivo real.
- **Sem commit/push sem ordem explícita.**

## O que responder ao usuário

Retorno **curto**: branch atual, se o working tree está limpo, qual aba/arquivo vamos tocar, e se havia
pendência da sessão anterior. Confirme numa frase que os gates acima estão ativos.
