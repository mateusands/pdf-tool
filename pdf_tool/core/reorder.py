"""Reordenação de listas — camada pura, compartilhada por Juntar, Organizar e Imagem → PDF.

Ver a spec em `tests/test_reorder.py`.
"""

_ACOES = ("left", "right", "start", "end")


def mover(itens, indice, acao: str):
    """Move `itens[indice]` e devolve `(nova_lista, nova_posicao)`.

    A nova posição vem junto de propósito: quem chama redesenha a tela e precisa
    manter a seleção no item que acabou de mover.
    """
    if acao not in _ACOES:
        raise ValueError(f"Ação de movimento desconhecida: {acao!r}")

    lista = list(itens)
    if indice is None:
        return lista, None

    ultimo = len(lista) - 1
    if acao in ("left", "start") and indice <= 0:
        return lista, indice
    if acao in ("right", "end") and indice >= ultimo:
        return lista, indice

    if acao == "left":
        destino = indice - 1
    elif acao == "right":
        destino = indice + 1
    elif acao == "start":
        destino = 0
    else:
        destino = ultimo

    lista.insert(destino, lista.pop(indice))
    return lista, destino


def indice_apos_arrastar(origem: int, alvo: int) -> int:
    """Posição de inserção para o item cair exatamente onde foi solto.

    Como o item já foi removido da lista, inserir em `alvo` o coloca no índice
    `alvo` da lista final — que é o cartão sobre o qual o usuário largou.
    """
    return alvo
