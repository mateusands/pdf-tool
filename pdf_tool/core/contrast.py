"""Contraste de cor pelo WCAG 2.1 — camada pura, sem UI.

Serve para decidir cor no tema com número em vez de opinião. Ver a spec em
`tests/test_contrast.py`.
"""


def _canais(cor: str):
    cor = cor.lstrip("#")
    if len(cor) == 3:                      # forma curta: #abc
        cor = "".join(c * 2 for c in cor)
    if len(cor) != 6:
        raise ValueError(f"Cor hexadecimal inválida: {cor!r}")
    return tuple(int(cor[i:i + 2], 16) / 255 for i in (0, 2, 4))


def luminancia_relativa(cor: str) -> float:
    """Luminância relativa da cor, de 0 (preto) a 1 (branco)."""
    def linear(canal):
        return canal / 12.92 if canal <= 0.03928 else ((canal + 0.055) / 1.055) ** 2.4

    vermelho, verde, azul = (linear(c) for c in _canais(cor))
    return 0.2126 * vermelho + 0.7152 * verde + 0.0722 * azul


def razao_de_contraste(cor_a: str, cor_b: str) -> float:
    """Razão de contraste entre duas cores, de 1 (idênticas) a 21 (branco/preto).

    O critério AA para texto normal é 4,5; para texto grande, 3.
    """
    clara = luminancia_relativa(cor_a) + 0.05
    escura = luminancia_relativa(cor_b) + 0.05
    return max(clara, escura) / min(clara, escura)
