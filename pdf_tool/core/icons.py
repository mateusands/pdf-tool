"""Ícones vetoriais da interface — camada pura, sem UI.

Conjunto **Lucide** v1.27.0 (https://lucide.dev), licença ISC:

    Copyright (c) 2026 Lucide Icons and Contributors

    Permission to use, copy, modify, and/or distribute this software for any
    purpose with or without fee is hereby granted, provided that the above
    copyright notice and this permission notice appear in all copies.

Os desenhos são guardados como o conteúdo interno do `<svg>` de cada ícone, numa
grade 24×24 com traço de 2. A moldura é montada em tempo de execução com o
tamanho e a cor pedidos — é isso que permite o mesmo ícone aparecer branco no
item ativo e cinza no inativo.

O PyMuPDF rasteriza o SVG; nenhuma dependência nova e nenhuma fonte instalada.
Ver a spec em `tests/test_icons.py`.
"""

from functools import lru_cache

import fitz

_MOLDE = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="{tamanho}" height="{tamanho}" '
    'viewBox="0 0 24 24" fill="none" stroke="{cor}" stroke-width="{espessura}" '
    'stroke-linecap="round" stroke-linejoin="round">{corpo}</svg>'
)

# nome -> conteúdo interno do <svg> (Lucide v1.27.0, ISC)
DESENHOS = {
    "arrow-left": "<path d='m12 19-7-7 7-7' /> <path d='M19 12H5' />",
    "arrow-right-left": (
        "<path d='m16 3 4 4-4 4' /> <path d='M20 7H4' /> <path d='m8 21-4-4 4-4' /> <path "
        "d='M4 17h16' />"
    ),
    "arrow-right": "<path d='M5 12h14' /> <path d='m12 5 7 7-7 7' />",
    "check": "<path d='M20 6 9 17l-5-5' />",
    "chevron-down": "<path d='m6 9 6 6 6-6' />",
    "chevron-up": "<path d='m18 15-6-6-6 6' />",
    "chevrons-left": "<path d='m11 17-5-5 5-5' /> <path d='m18 17-5-5 5-5' />",
    "chevrons-right": "<path d='m6 17 5-5-5-5' /> <path d='m13 17 5-5-5-5' />",
    "circle-alert": (
        "<circle cx='12' cy='12' r='10' /> <line x1='12' x2='12' y1='8' y2='12' /> <line "
        "x1='12' x2='12.01' y1='16' y2='16' />"
    ),
    "circle-check": "<circle cx='12' cy='12' r='10' /> <path d='m9 12 2 2 4-4' />",
    "combine": (
        "<path d='M14 3a1 1 0 0 1 1 1v5a1 1 0 0 1-1 1' /> <path d='M19 3a1 1 0 0 1 1 1v5a1 1 "
        "0 0 1-1 1' /> <path d='m7 15 3 3' /> <path d='m7 21 3-3H5a2 2 0 0 1-2-2v-2' /> <rect "
        "x='14' y='14' width='7' height='7' rx='1' /> <rect x='3' y='3' width='7' height='7' "
        "rx='1' />"
    ),
    "eye-off": (
        "<path d='M10.733 5.076a10.744 10.744 0 0 1 11.205 6.575 1 1 0 0 1 0 .696 10.747 "
        "10.747 0 0 1-1.444 2.49' /> <path d='M14.084 14.158a3 3 0 0 1-4.242-4.242' /> <path "
        "d='M17.479 17.499a10.75 10.75 0 0 1-15.417-5.151 1 1 0 0 1 0-.696 10.75 10.75 0 0 1 "
        "4.446-5.143' /> <path d='m2 2 20 20' />"
    ),
    "eye": (
        "<path d='M2.062 12.348a1 1 0 0 1 0-.696 10.75 10.75 0 0 1 19.876 0 1 1 0 0 1 0 .696 "
        "10.75 10.75 0 0 1-19.876 0' /> <circle cx='12' cy='12' r='3' />"
    ),
    "file-image": (
        "<path d='M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 "
        "3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z' /> <path d='M14 2v5a1 1 0 0 0 1 1h5' /> "
        "<circle cx='10' cy='12' r='2' /> <path d='m20 17-1.296-1.296a2.41 2.41 0 0 0-3.408 "
        "0L9 22' />"
    ),
    "file-text": (
        "<path d='M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 "
        "3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z' /> <path d='M14 2v5a1 1 0 0 0 1 1h5' /> "
        "<path d='M10 9H8' /> <path d='M16 13H8' /> <path d='M16 17H8' />"
    ),
    "flip-vertical": (
        "<path d='M21 8V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v3' /> <path d='M21 16v3a2 2 0 0 1-2 "
        "2H5a2 2 0 0 1-2-2v-3' /> <path d='M4 12H2' /> <path d='M10 12H8' /> <path d='M16 "
        "12h-2' /> <path d='M22 12h-2' />"
    ),
    "folder-open": (
        "<path d='m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 "
        "1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 "
        "1.67.9H18a2 2 0 0 1 2 2v2' />"
    ),
    "image-down": (
        "<path d='M10.3 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v10l-3.1-3.1a2 2 0 "
        "0 0-2.814.014L6 21' /> <path d='m14 19 3 3v-5.5' /> <path d='m17 22 3-3' /> <circle "
        "cx='9' cy='9' r='2' />"
    ),
    "info": "<circle cx='12' cy='12' r='10' /> <path d='M12 16v-4' /> <path d='M12 8h.01' />",
    "layout-grid": (
        "<rect width='7' height='7' x='3' y='3' rx='1' /> <rect width='7' height='7' x='14' "
        "y='3' rx='1' /> <rect width='7' height='7' x='14' y='14' rx='1' /> <rect width='7' "
        "height='7' x='3' y='14' rx='1' />"
    ),
    "loader-circle": "<path d='M21 12a9 9 0 1 1-6.219-8.56' />",
    "lock-open": (
        "<rect width='18' height='11' x='3' y='11' rx='2' ry='2' /> <path d='M7 11V7a5 5 0 0 "
        "1 9.9-1' />"
    ),
    "lock": (
        "<rect width='18' height='11' x='3' y='11' rx='2' ry='2' /> <path d='M7 11V7a5 5 0 0 "
        "1 10 0v4' />"
    ),
    "panel-left-close": (
        "<rect width='18' height='18' x='3' y='3' rx='2' /> <path d='M9 3v18' /> <path d='m16 "
        "15-3-3 3-3' />"
    ),
    "plus": "<path d='M5 12h14' /> <path d='M12 5v14' />",
    "rotate-ccw": "<path d='M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8' /> <path d='M3 3v5h5' />",
    "rotate-cw": "<path d='M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8' /> <path d='M21 3v5h-5' />",
    "save": (
        "<path d='M15.2 3a2 2 0 0 1 1.4.6l3.8 3.8a2 2 0 0 1 .6 1.4V19a2 2 0 0 1-2 2H5a2 2 0 0 "
        "1-2-2V5a2 2 0 0 1 2-2z' /> <path d='M17 21v-7a1 1 0 0 0-1-1H8a1 1 0 0 0-1 1v7' /> "
        "<path d='M7 3v4a1 1 0 0 0 1 1h7' />"
    ),
    "scissors": (
        "<circle cx='6' cy='6' r='3' /> <path d='M8.12 8.12 12 12' /> <path d='M20 4 8.12 "
        "15.88' /> <circle cx='6' cy='18' r='3' /> <path d='M14.8 14.8 20 20' />"
    ),
    "shrink": (
        "<path d='m15 15 6 6m-6-6v4.8m0-4.8h4.8' /> <path d='M9 19.8V15m0 0H4.2M9 15l-6 6' /> "
        "<path d='M15 4.2V9m0 0h4.8M15 9l6-6' /> <path d='M9 4.2V9m0 0H4.2M9 9 3 3' />"
    ),
    "trash-2": (
        "<path d='M10 11v6' /> <path d='M14 11v6' /> <path d='M19 6v14a2 2 0 0 1-2 2H7a2 2 0 "
        "0 1-2-2V6' /> <path d='M3 6h18' /> <path d='M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2' "
        "/>"
    ),
    "upload": (
        "<path d='M12 3v12' /> <path d='m17 8-5-5-5 5' /> <path d='M21 15v4a2 2 0 0 1-2 2H5a2 "
        "2 0 0 1-2-2v-4' />"
    ),
    "x": "<path d='M18 6 6 18' /> <path d='m6 6 12 12' />",
}


def nomes() -> list:
    """Nomes dos ícones disponíveis, em ordem."""
    return sorted(DESENHOS)


@lru_cache(maxsize=512)
def renderizar_png(nome: str, tamanho: int, cor: str, espessura: float = 2) -> bytes:
    """Bytes de um PNG `tamanho`×`tamanho`, fundo transparente, traço em `cor`.

    Pronto para `tk.PhotoImage(data=…)`. O resultado fica em cache: o mesmo
    ícone é pedido a cada redesenho de aba.
    """
    corpo = DESENHOS.get(nome)
    if corpo is None:
        raise ValueError(
            f"Ícone desconhecido: {nome!r}. Disponíveis: {', '.join(nomes())}"
        )

    svg = _MOLDE.format(tamanho=tamanho, cor=cor, espessura=espessura, corpo=corpo)
    return _rasterizar(svg)


def _rasterizar(svg: str) -> bytes:
    documento = fitz.open(stream=svg.encode("utf-8"), filetype="svg")
    try:
        return documento[0].get_pixmap(alpha=True).tobytes("png")
    finally:
        documento.close()


# ── Ícone da janela ───────────────────────────────────────────────────────────

_MOLDE_DO_APP = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="{tamanho}" height="{tamanho}" '
    'viewBox="0 0 64 64">'
    '<rect x="0" y="0" width="64" height="64" rx="14" ry="14" fill="{fundo}"/>'
    '<g transform="translate(16 16) scale(1.3333)" fill="none" stroke="{traco}" '
    'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{corpo}</g>'
    "</svg>"
)


@lru_cache(maxsize=8)
def renderizar_icone_do_app(tamanho: int, fundo: str, traco: str) -> bytes:
    """Ícone da janela: quadrado arredondado com a cor da marca e o glifo de documento.

    É o que o sistema mostra na barra de título, na barra de tarefas e no
    alt-tab. Sem ele o app aparece com o losango genérico do Tk.
    """
    return _rasterizar(
        _MOLDE_DO_APP.format(tamanho=tamanho, fundo=fundo, traco=traco,
                             corpo=DESENHOS["file-text"])
    )
