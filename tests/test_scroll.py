"""
SDD — Especificação: rolagem com a roda do mouse, multiplataforma

CONTRATO
  `passos_de_rolagem(delta, num)` traduz um evento de roda do mouse no número de
  unidades para `yview_scroll`. Negativo sobe, positivo desce, zero não faz nada.

POR QUE EXISTE
  As quatro áreas roláveis do app (grade de miniaturas, lista do Juntar, grade do
  Organizar, lista do Imagem → PDF) escutavam só `<MouseWheel>`. No Linux com
  Tk 8.6 em X11 esse evento NUNCA é emitido: a roda vira `<Button-4>` (cima) e
  `<Button-5>` (baixo). Resultado: a roda do mouse não rolava nada no Linux, e o
  usuário só conseguia usar a barra lateral.

  Confirmado no ambiente do projeto (`tk.windowingsystem == "x11"`,
  `TkVersion 8.6`): disparando `<Button-4>`, só o handler de `Button-4` recebeu.

REGRA DE NEGÓCIO
  - X11 manda `num` 4/5 e `delta` 0 — é o `num` que decide.
  - Windows manda `delta` em múltiplos de 120; girar rápido acumula (360 = três
    passos), e isso precisa ser respeitado ou a rolagem fica lenta.
  - macOS manda `delta` pequeno (±1, ±2) e sem múltiplo de 120 — dividir por 120
    ali daria zero e a roda não faria nada.
  - Em toda plataforma, delta positivo = rolar para cima = unidades negativas.
"""

import pytest

from pdf_tool.widgets import passos_de_rolagem


class TestRolagemNoLinuxX11:
    def test_deve_subir_com_o_botao_4(self):
        assert passos_de_rolagem(delta=0, num=4) == -1

    def test_deve_descer_com_o_botao_5(self):
        assert passos_de_rolagem(delta=0, num=5) == 1


class TestRolagemNoWindows:
    def test_deve_subir_um_passo_com_delta_120(self):
        assert passos_de_rolagem(delta=120, num=None) == -1

    def test_deve_descer_um_passo_com_delta_negativo(self):
        assert passos_de_rolagem(delta=-120, num=None) == 1

    def test_deve_acumular_quando_a_roda_gira_rapido(self):
        assert passos_de_rolagem(delta=360, num=None) == -3

    def test_deve_ignorar_o_num_invalido_que_o_tk_manda_no_mousewheel(self):
        # No Windows o Tk preenche `num` com "??" nos eventos de roda.
        assert passos_de_rolagem(delta=120, num="??") == -1


class TestRolagemNoMacos:
    @pytest.mark.parametrize("delta, esperado", [(1, -1), (2, -1), (-1, 1), (-3, 1)])
    def test_deve_rolar_mesmo_com_delta_pequeno(self, delta, esperado):
        # Dividir por 120 aqui daria zero e a roda não faria nada.
        assert passos_de_rolagem(delta=delta, num=None) == esperado


class TestEventoSemMovimento:
    def test_deve_devolver_zero_quando_nao_ha_delta_nem_botao_de_roda(self):
        assert passos_de_rolagem(delta=0, num=1) == 0
