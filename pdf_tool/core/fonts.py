"""Escolha da família de fonte da interface — camada pura, sem UI.

Ver a spec em `tests/test_fonts.py`.
"""

import sys

#: Ordem de preferência por plataforma. A primeira que existir no sistema vence.
PREFERIDAS = {
    "win32": ["Segoe UI Variable Text", "Segoe UI", "Tahoma"],
    "darwin": ["SF Pro Text", "Helvetica Neue", "Lucida Grande"],
    "linux": ["Inter", "Cantarell", "Ubuntu", "Noto Sans", "DejaVu Sans", "Liberation Sans"],
}

ALTERNATIVA = "TkDefaultFont"


def preferidas_da_plataforma(plataforma: str = None) -> list:
    """Lista de preferência para a plataforma atual (ou a informada)."""
    plataforma = plataforma if plataforma is not None else sys.platform
    if plataforma.startswith("linux"):
        return list(PREFERIDAS["linux"])
    return list(PREFERIDAS.get(plataforma, PREFERIDAS["linux"]))


def escolher(disponiveis, preferidas, alternativa: str = ALTERNATIVA) -> str:
    """Primeira de `preferidas` presente em `disponiveis`, ou `alternativa`."""
    indice = {str(familia).casefold(): str(familia) for familia in disponiveis}
    for familia in preferidas:
        encontrada = indice.get(str(familia).casefold())
        if encontrada:
            return encontrada
    return alternativa
