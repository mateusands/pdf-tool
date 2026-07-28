"""
SDD — Especificação: contraste de texto do tema

CONTRATO
  `razao_de_contraste(cor_a, cor_b)` devolve a razão de contraste WCAG 2.1 entre
  duas cores hexadecimais, de 1:1 (idênticas) a 21:1 (branco sobre preto). A
  ordem dos argumentos não importa.

  O critério AA para texto normal é **4,5:1**.

POR QUE EXISTE
  A interface é dark mode e a paleta é escolhida a olho, o que engana: um azul
  bonito pode ter contraste péssimo com o branco por cima. Foi o que aconteceu.

  O botão primário usava `#3B82F6`, que dá só 3,68:1 com texto branco. E quando
  o botão ficava desabilitado, o CustomTkinter escurecia só o TEXTO para `gray60`
  mantendo o preenchimento azul: **1,29:1**, praticamente invisível — foi assim
  que o defeito apareceu, num botão "Exportar imagens" ilegível.

  A lição que este arquivo protege: **não existe letra mais branca que branco**.
  Quando falta contraste num botão, quem tem de escurecer é o preenchimento.

REGRA DE NEGÓCIO
  - Cor de PREENCHIMENTO (`ACCENT`, `SUCCESS`, `DANGER`) é julgada contra o texto
    branco que fica por cima dela.
  - Cor de TEXTO (`ACCENT_TEXT`, `*_TEXT`, `FG`, `MUTED`) é julgada contra a
    superfície escura em que ela é escrita. São papéis opostos, e é por isso que
    o tema tem dois tons de cada cor.
  - Texto desabilitado é dispensado do AA pela norma, mas ainda precisa ser
    legível — o piso aqui é 3:1.
"""

import pytest

from pdf_tool import theme as T
from pdf_tool.core.contrast import razao_de_contraste

AA = 4.5


class TestCalculoDaRazao:
    def test_branco_sobre_preto_e_o_maximo(self):
        assert razao_de_contraste("#FFFFFF", "#000000") == pytest.approx(21, abs=0.01)

    def test_cores_identicas_nao_tem_contraste(self):
        assert razao_de_contraste("#2563EB", "#2563EB") == pytest.approx(1, abs=0.01)

    def test_a_ordem_dos_argumentos_nao_importa(self):
        a = razao_de_contraste("#FFFFFF", "#2563EB")
        b = razao_de_contraste("#2563EB", "#FFFFFF")

        assert a == b

    def test_deve_aceitar_hexadecimal_sem_cerquilha(self):
        assert razao_de_contraste("FFFFFF", "000000") == pytest.approx(21, abs=0.01)


class TestPreenchimentoComTextoBranco:
    """Botões preenchidos: o texto branco por cima precisa passar no AA."""

    @pytest.mark.parametrize("token", ["ACCENT", "ACCENT_ACTIVE", "SUCCESS",
                                       "SUCCESS_HOVER", "DANGER", "DANGER_HOVER"])
    def test_deve_passar_no_aa_com_texto_branco(self, token):
        razao = razao_de_contraste(T.ON_ACCENT, getattr(T, token))

        assert razao >= AA, f"T.{token} dá {razao:.2f}:1 — escureça o preenchimento"


class TestTextoSobreSuperficie:
    """Cores usadas como texto: julgadas contra a superfície onde são escritas."""

    @pytest.mark.parametrize("token, fundo", [
        ("FG", "CARD_BG"), ("FG_SECONDARY", "CARD_BG"), ("ACCENT_TEXT", "CARD_BG"),
        ("FG", "SIDEBAR_BG"), ("FG_SECONDARY", "SIDEBAR_BG"),
        ("SUCCESS_TEXT", "SURFACE_1"), ("DANGER_TEXT", "SURFACE_1"),
        ("WARNING_TEXT", "SURFACE_1"), ("ACCENT_TEXT", "SURFACE_1"),
    ])
    def test_deve_passar_no_aa(self, token, fundo):
        razao = razao_de_contraste(getattr(T, token), getattr(T, fundo))

        assert razao >= AA, f"T.{token} sobre T.{fundo} dá {razao:.2f}:1"


class TestTextoSecundario:
    def test_texto_apagado_deve_ser_legivel_mesmo_dispensado_do_aa(self):
        # `MUTED` é rótulo de apoio (grupo da sidebar, legenda). O piso é 3:1.
        for fundo in (T.CARD_BG, T.SIDEBAR_BG, T.GRID_BG):
            assert razao_de_contraste(T.MUTED, fundo) >= 3.0

    def test_botao_desabilitado_deve_continuar_legivel(self):
        # Foi exatamente aqui que estava o defeito: 1,29:1 com o azul mantido.
        razao = razao_de_contraste(T.DISABLED_FG, T.DISABLED_BG)

        assert razao >= 3.0, f"texto desabilitado dá {razao:.2f}:1"

    def test_o_estado_desabilitado_nao_pode_usar_o_preenchimento_ativo(self):
        # Se o fundo continuasse azul, o botão pareceria clicável — e foi assim
        # que o texto apagado virou ilegível.
        assert T.DISABLED_BG != T.ACCENT
