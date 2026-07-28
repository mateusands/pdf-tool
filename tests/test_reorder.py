"""
SDD — Especificação: reordenação de listas (páginas e arquivos)

CONTRATO
  `mover(itens, indice, acao)` devolve `(nova_lista, novo_indice)`. O índice
  devolvido é a NOVA posição do item movido — quem chama usa esse valor para
  manter a seleção depois de redesenhar a tela.

  `indice_apos_arrastar(origem, alvo)` devolve em que posição inserir o item
  depois de removê-lo da lista, para que ele termine exatamente onde foi solto.

POR QUE EXISTE
  Três abas reimplementavam reordenação (Juntar, Organizar, Imagem → PDF) e duas
  tinham defeito:

  - **Juntar**: "↑ Subir" só funcionava uma vez. O método guardava a nova posição
    e em seguida redesenhava a lista, e o redesenho zerava a seleção. O usuário
    tinha que clicar de novo na linha a cada movimento.
  - **Organizar**: o arrastar soltava o item uma posição ANTES do cartão onde o
    usuário largou (`insert_at = tgt - 1 if tgt > src`).

  Devolver a nova posição junto com a lista faz o chamador ficar correto por
  construção — não há como redesenhar e esquecer a seleção.

REGRA DE NEGÓCIO
  - Movimento inválido (subir o primeiro, descer o último) não é erro: não faz
    nada e mantém a seleção onde está.
  - A lista original nunca é mutada — a função devolve uma nova.
  - Soltar sobre a própria origem não muda nada.
"""

import pytest

from pdf_tool.core.reorder import indice_apos_arrastar, mover


class TestMoverItemDaLista:
    @pytest.mark.parametrize(
        "acao, indice, esperado_lista, esperado_indice",
        [
            ("left",  2, ["a", "c", "b", "d"], 1),
            ("right", 1, ["a", "c", "b", "d"], 2),
            ("start", 2, ["c", "a", "b", "d"], 0),
            ("end",   1, ["a", "c", "d", "b"], 3),
        ],
    )
    def test_deve_mover_e_devolver_a_nova_posicao(
        self, acao, indice, esperado_lista, esperado_indice
    ):
        lista, novo = mover(["a", "b", "c", "d"], indice, acao)

        assert lista == esperado_lista
        assert novo == esperado_indice

    @pytest.mark.parametrize(
        "acao, indice",
        [("left", 0), ("start", 0), ("right", 3), ("end", 3)],
    )
    def test_deve_ignorar_movimento_impossivel_sem_perder_a_selecao(self, acao, indice):
        lista, novo = mover(["a", "b", "c", "d"], indice, acao)

        assert lista == ["a", "b", "c", "d"]
        assert novo == indice, "a seleção precisa continuar onde estava"

    def test_deve_ignorar_quando_nada_esta_selecionado(self):
        assert mover(["a", "b"], None, "left") == (["a", "b"], None)

    def test_nao_deve_mutar_a_lista_original(self):
        original = ["a", "b", "c"]

        mover(original, 0, "end")

        assert original == ["a", "b", "c"]

    def test_deve_recusar_uma_acao_desconhecida(self):
        with pytest.raises(ValueError):
            mover(["a", "b"], 0, "diagonal")


class TestArrastarParaReordenar:
    def test_deve_soltar_o_item_exatamente_onde_o_usuario_largou(self):
        paginas = ["a", "b", "c", "d"]
        origem, alvo = 0, 2

        item = paginas.pop(origem)
        paginas.insert(indice_apos_arrastar(origem, alvo), item)

        assert paginas == ["b", "c", "a", "d"]
        assert paginas.index("a") == alvo, "tem que parar no cartão onde foi solto"

    def test_deve_funcionar_arrastando_da_direita_para_a_esquerda(self):
        paginas = ["a", "b", "c", "d"]
        origem, alvo = 3, 1

        item = paginas.pop(origem)
        paginas.insert(indice_apos_arrastar(origem, alvo), item)

        assert paginas == ["a", "d", "b", "c"]
        assert paginas.index("d") == alvo
