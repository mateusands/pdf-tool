"""
SDD — Especificação: trabalho pesado fora da thread da UI

CONTRATO
  `executar_em_thread(root, trabalho, ao_terminar, ao_falhar)` roda `trabalho()`
  numa thread daemon e volta para a thread da UI por `root.after(0, …)`:
  `ao_terminar(resultado)` no caminho feliz, `ao_falhar(mensagem)` no erro.
  Nenhum callback é chamado de dentro da thread — Tkinter é single-threaded.

POR QUE EXISTE
  As quatro abas que já usavam thread repetiam este trecho:

      except Exception as e:
          self.root.after(0, lambda: self._error(str(e)))

  Isso NUNCA mostrou erro nenhum ao usuário. O Python apaga o nome `e` ao sair do
  bloco `except`, e o lambda só roda depois, no `after` — aí `str(e)` levanta
  `NameError: cannot access free variable 'e'`. A exceção morria dentro do
  callback do Tk: sem messagebox, sem status, só um traceback no terminal.

  O caso mais caro era a conversão Word → PDF: o `docx_convert` monta uma
  mensagem com o comando de instalação do LibreOffice, e ela nunca chegava na
  tela. O usuário via a barra de status parada em "⏳ Convertendo…" para sempre.

REGRA DE NEGÓCIO
  - A mensagem do erro é extraída DENTRO do `except` e passada por valor. É o que
    garante que o texto sobrevive até o `after` disparar.
  - Erro em `trabalho` nunca escapa silenciosamente: sempre vira `ao_falhar`.
"""

import threading

import pytest

from pdf_tool.core.background import executar_em_thread


class RootFalso:
    """Dublê da janela: guarda o que o `after` receberia e executa na hora do teste."""

    def __init__(self):
        self.agendado = []

    def after(self, _atraso, funcao):
        self.agendado.append((threading.current_thread().name, funcao))

    def rodar_agendados(self):
        for _, funcao in self.agendado:
            funcao()


class TestCaminhoFeliz:
    def test_deve_entregar_o_resultado_do_trabalho(self):
        root, recebido = RootFalso(), []

        executar_em_thread(
            root, lambda: 42, ao_terminar=recebido.append, ao_falhar=lambda m: None
        ).join(timeout=5)
        root.rodar_agendados()

        assert recebido == [42]

    def test_nao_deve_tocar_na_ui_de_dentro_da_thread(self):
        root = RootFalso()

        executar_em_thread(
            root, lambda: "ok", ao_terminar=lambda r: None, ao_falhar=lambda m: None
        ).join(timeout=5)

        assert len(root.agendado) == 1, "o retorno tem que passar pelo root.after"


class TestCaminhoDeErro:
    def test_deve_entregar_a_mensagem_do_erro(self):
        # Regressão: com `lambda: str(e)` isto levantava NameError e o usuário
        # não via aviso nenhum.
        root, falhas = RootFalso(), []

        def trabalho():
            raise RuntimeError("O LibreOffice não gerou o PDF.")

        executar_em_thread(
            root, trabalho, ao_terminar=lambda r: None, ao_falhar=falhas.append
        ).join(timeout=5)
        root.rodar_agendados()

        assert falhas == ["O LibreOffice não gerou o PDF."]

    def test_deve_preservar_a_mensagem_acionavel_inteira(self):
        root, falhas = RootFalso(), []
        mensagem = (
            "LibreOffice não encontrado — ele é necessário para converter "
            "Word → PDF no Linux.\n\nArch/CachyOS:  sudo pacman -S libreoffice-fresh"
        )

        def trabalho():
            raise RuntimeError(mensagem)

        executar_em_thread(
            root, trabalho, ao_terminar=lambda r: None, ao_falhar=falhas.append
        ).join(timeout=5)
        root.rodar_agendados()

        assert falhas[0] == mensagem, "o comando de instalação precisa chegar à tela"

    def test_nao_deve_chamar_ao_terminar_quando_o_trabalho_falha(self):
        root, sucessos, falhas = RootFalso(), [], []

        def trabalho():
            raise ValueError("boom")

        executar_em_thread(
            root, trabalho, ao_terminar=sucessos.append, ao_falhar=falhas.append
        ).join(timeout=5)
        root.rodar_agendados()

        assert sucessos == []
        assert falhas == ["boom"]

    def test_deve_avisar_mesmo_quando_a_excecao_nao_tem_mensagem(self):
        root, falhas = RootFalso(), []

        def trabalho():
            raise MemoryError()

        executar_em_thread(
            root, trabalho, ao_terminar=lambda r: None, ao_falhar=falhas.append
        ).join(timeout=5)
        root.rodar_agendados()

        assert falhas and falhas[0].strip(), "não pode chegar mensagem vazia ao usuário"


class TestExecucaoForaDaThreadPrincipal:
    def test_deve_rodar_o_trabalho_numa_thread_separada(self):
        root = RootFalso()
        onde = []

        executar_em_thread(
            root,
            lambda: onde.append(threading.current_thread()),
            ao_terminar=lambda r: None,
            ao_falhar=lambda m: None,
        ).join(timeout=5)

        assert onde[0] is not threading.main_thread()


@pytest.mark.parametrize("callback", ["ao_terminar", "ao_falhar"])
def test_deve_agendar_o_callback_pela_thread_de_trabalho(callback):
    # O `after` do Tk é seguro de chamar de outra thread; o callback em si não.
    root = RootFalso()
    trabalho = (lambda: 1) if callback == "ao_terminar" else _falha

    executar_em_thread(
        root, trabalho, ao_terminar=lambda r: None, ao_falhar=lambda m: None
    ).join(timeout=5)

    nome_da_thread = root.agendado[0][0]
    assert nome_da_thread != threading.main_thread().name


def _falha():
    raise RuntimeError("x")
