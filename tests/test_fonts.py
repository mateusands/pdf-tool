"""
SDD — Especificação: escolha da família de fonte

CONTRATO
  `escolher(disponiveis, preferidas, alternativa)` devolve a primeira família de
  `preferidas` que exista em `disponiveis`, ignorando maiúsculas/minúsculas. Se
  nenhuma existir, devolve `alternativa`.

POR QUE EXISTE
  O tema fixava `FONT_FAMILY = "Segoe UI"`, que só existe no Windows. No Linux o
  Tk fazia um fallback silencioso para a fonte padrão do sistema — o app
  funcionava, mas com a tipografia que sobrasse, e o `CLAUDE.md` registrava isso
  como "esperado".

  Não é preciso aceitar: cada plataforma tem uma fonte de interface boa e
  previsível (Segoe UI no Windows, SF Pro no macOS, Inter/Cantarell/Noto no
  Linux). Escolher a primeira que existe dá tipografia consistente em vez de
  sorteada.

REGRA DE NEGÓCIO
  - A ordem de `preferidas` é a ordem de preferência real, não alfabética.
  - A comparação ignora caixa: o Tk reporta "DejaVu Sans", listas trazem
    "dejavu sans".
  - Sempre devolve algo utilizável — nunca `None`. Interface sem fonte não abre.
"""

import pytest

from pdf_tool.core.fonts import escolher


class TestEscolhaDaFamilia:
    def test_deve_pegar_a_primeira_preferida_que_existe(self):
        disponiveis = ["DejaVu Sans", "Cantarell", "Noto Sans"]

        assert escolher(disponiveis, ["Inter", "Cantarell", "Noto Sans"], "TkDefaultFont") == "Cantarell"

    def test_deve_respeitar_a_ordem_de_preferencia_e_nao_a_alfabetica(self):
        disponiveis = ["Arial", "Inter"]

        assert escolher(disponiveis, ["Inter", "Arial"], "TkDefaultFont") == "Inter"

    def test_deve_ignorar_maiusculas_e_minusculas(self):
        # O Tk reporta "DejaVu Sans"; a lista de preferência pode vir em caixa baixa.
        assert escolher(["dejavu sans"], ["DejaVu Sans"], "TkDefaultFont") == "dejavu sans"

    def test_deve_cair_na_alternativa_quando_nenhuma_existe(self):
        assert escolher(["Comic Sans MS"], ["Inter", "Segoe UI"], "TkDefaultFont") == "TkDefaultFont"

    def test_deve_devolver_a_alternativa_quando_o_sistema_nao_lista_nada(self):
        assert escolher([], ["Inter"], "TkDefaultFont") == "TkDefaultFont"

    @pytest.mark.parametrize("alternativa", ["TkDefaultFont", "Helvetica"])
    def test_nunca_deve_devolver_nulo(self, alternativa):
        # Interface sem família de fonte não abre.
        assert escolher([], [], alternativa) == alternativa
