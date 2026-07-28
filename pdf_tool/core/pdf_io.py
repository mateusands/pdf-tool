"""Operações de arquivo PDF — camada pura, sem UI.

Toda escrita de PDF do app passa por aqui. As funções recebem caminhos e
parâmetros, não tocam em widget e levantam `PdfError` com a mensagem já pronta
para o `messagebox`.

Duas garantias valem para todas elas:

1. **O destino nunca é a origem.** `validar_destino()` roda antes de qualquer
   abertura de arquivo de saída — ver a spec em `tests/test_pdf_io.py`.
2. **A escrita é atômica.** Grava num `.part` ao lado do destino e só então faz
   `os.replace`. Se algo falhar no meio, o destino não é criado (nem um destino
   antigo é destruído).
"""

import os
from contextlib import contextmanager

import fitz
from pypdf import PdfReader, PdfWriter


class PdfError(RuntimeError):
    """Erro com mensagem já pronta para exibir ao usuário."""


class DestinoInvalidoError(PdfError):
    """O arquivo de saída escolhido sobrescreveria uma das origens."""


class PdfProtegidoError(PdfError):
    """O PDF exige senha para ser lido."""


class SenhaIncorretaError(PdfError):
    """A senha informada não abre o PDF."""


# ── Destino ───────────────────────────────────────────────────────────────────

def _caminho_real(caminho: str) -> str:
    return os.path.normcase(os.path.realpath(caminho))


def validar_destino(destino: str, *origens: str) -> None:
    """Recusa salvar por cima de qualquer arquivo de entrada.

    Compara por `realpath`, então caminho relativo e link simbólico também são
    detectados.
    """
    alvo = _caminho_real(destino)
    for origem in origens:
        if _caminho_real(origem) == alvo:
            nome = os.path.basename(origem)
            raise DestinoInvalidoError(
                f"O destino escolhido é o próprio arquivo de origem ({nome}).\n\n"
                "Escolha outro nome — esta ferramenta nunca sobrescreve o "
                "arquivo de entrada."
            )


# ── Escrita atômica ───────────────────────────────────────────────────────────

@contextmanager
def _escrita_atomica(destino: str):
    """Abre um `.part` ao lado do destino e só promove no fim, sem erro."""
    parcial = destino + ".part"
    try:
        with open(parcial, "wb") as arquivo:
            yield arquivo
    except BaseException:
        if os.path.exists(parcial):
            os.unlink(parcial)
        raise
    os.replace(parcial, destino)


# ── Leitura ───────────────────────────────────────────────────────────────────

def _ler(caminho: str) -> PdfReader:
    """Abre o PDF para leitura, traduzindo os erros do motor."""
    try:
        leitor = PdfReader(caminho)
    except Exception as erro:
        raise PdfError(
            f"Não foi possível ler «{os.path.basename(caminho)}».\n\n"
            f"O arquivo pode estar corrompido ou não ser um PDF.\n\n{erro}"
        ) from None

    if leitor.is_encrypted:
        raise PdfProtegidoError(_MSG_PROTEGIDO.format(nome=os.path.basename(caminho)))
    return leitor


_MSG_PROTEGIDO = (
    "«{nome}» está protegido por senha e não pode ser aberto aqui.\n\n"
    "Use a ferramenta Desbloquear para gerar uma cópia sem senha e trabalhe nela."
)


@contextmanager
def abrir_documento(caminho: str):
    """Abre o PDF com o PyMuPDF (miniaturas, compactação) e fecha no fim.

    O `fitz` levanta "document closed or encrypted" para PDF com senha, o que não
    diz nada ao usuário — aqui isso vira `PdfProtegidoError`.
    """
    try:
        documento = fitz.open(caminho)
    except Exception as erro:
        raise PdfError(
            f"Não foi possível abrir «{os.path.basename(caminho)}».\n\n"
            f"O arquivo pode estar corrompido ou não ser um PDF.\n\n{erro}"
        ) from None

    if documento.needs_pass:
        documento.close()
        raise PdfProtegidoError(_MSG_PROTEGIDO.format(nome=os.path.basename(caminho)))

    try:
        yield documento
    finally:
        documento.close()


def miniaturas_ppm(caminho: str, largura: int, altura: int) -> list:
    """Renderiza cada página como bytes PPM, prontos para `tk.PhotoImage(data=…)`.

    Devolve uma lista (e não um gerador) de propósito: assim um PDF protegido
    falha na chamada, dentro do `try` da aba, e não lá na frente durante o laço.
    """
    with abrir_documento(caminho) as documento:
        return [
            pagina.get_pixmap(
                matrix=fitz.Matrix(largura / pagina.rect.width,
                                   altura / pagina.rect.height)
            ).tobytes("ppm")
            for pagina in documento
        ]


# ── Operações ─────────────────────────────────────────────────────────────────

def dividir_pdf(origem: str, destino: str, paginas) -> int:
    """Grava em `destino` só as `paginas` (índices 0-based) de `origem`."""
    validar_destino(destino, origem)
    indices = sorted(set(paginas))
    if not indices:
        raise PdfError("Nenhuma página foi selecionada.")

    leitor = _ler(origem)
    total = len(leitor.pages)
    fora = [i for i in indices if not 0 <= i < total]
    if fora:
        raise PdfError(
            f"O PDF tem {total} página(s); a seleção aponta para páginas que não existem."
        )

    escritor = PdfWriter()
    for i in indices:
        escritor.add_page(leitor.pages[i])
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)
    return len(indices)


def juntar_pdfs(origens, destino: str) -> None:
    """Concatena `origens` na ordem da lista."""
    origens = list(origens)
    if not origens:
        raise PdfError("Adicione pelo menos um arquivo PDF.")
    validar_destino(destino, *origens)

    escritor = PdfWriter()
    for caminho in origens:
        try:
            escritor.append(caminho)
        except Exception as erro:
            raise PdfError(
                f"Não foi possível juntar «{os.path.basename(caminho)}».\n\n"
                f"Ele pode estar protegido por senha ou corrompido.\n\n{erro}"
            ) from None
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)


def girar_pdf(origem: str, destino: str, paginas, angulo: int) -> None:
    """Gira as `paginas` escolhidas em `angulo` graus; as demais ficam como estão."""
    validar_destino(destino, origem)
    alvo = set(paginas)
    if not alvo:
        raise PdfError("Nenhuma página foi selecionada.")

    leitor = _ler(origem)
    escritor = PdfWriter()
    for i, pagina in enumerate(leitor.pages):
        if i in alvo:
            pagina.rotate(angulo)
        escritor.add_page(pagina)
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)


def reorganizar_pdf(origem: str, destino: str, ordem) -> None:
    """Grava as páginas de `origem` na sequência de índices `ordem`."""
    validar_destino(destino, origem)
    ordem = list(ordem)
    if not ordem:
        raise PdfError("Nenhuma página para salvar.")

    leitor = _ler(origem)
    total = len(leitor.pages)
    if any(not 0 <= i < total for i in ordem):
        raise PdfError(
            f"O PDF tem {total} página(s); a ordem aponta para páginas que não existem."
        )

    escritor = PdfWriter()
    for i in ordem:
        escritor.add_page(leitor.pages[i])
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)


def comprimir_pdf(origem: str, destino: str, **opcoes) -> int:
    """Salva uma versão compactada de `origem`; devolve o tamanho final em bytes."""
    validar_destino(destino, origem)
    with abrir_documento(origem) as documento:
        parcial = destino + ".part"
        try:
            documento.save(parcial, **opcoes)
        except BaseException:
            if os.path.exists(parcial):
                os.unlink(parcial)
            raise
    os.replace(parcial, destino)
    return os.path.getsize(destino)


# ── Senha ─────────────────────────────────────────────────────────────────────

def proteger_pdf(origem: str, destino: str, senha: str) -> None:
    """Gera uma cópia de `origem` cifrada com AES-256.

    O default do pypdf é RC4-128, que é criptografia quebrada — numa ferramenta
    cujo propósito é proteger o arquivo, isso seria falsa sensação de segurança.
    """
    validar_destino(destino, origem)
    if not senha:
        raise PdfError("Digite uma senha.")

    leitor = _ler(origem)
    escritor = PdfWriter()
    for pagina in leitor.pages:
        escritor.add_page(pagina)
    escritor.encrypt(senha, algorithm="AES-256")
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)


def desbloquear_pdf(origem: str, destino: str, senha: str) -> None:
    """Gera uma cópia de `origem` sem proteção por senha."""
    validar_destino(destino, origem)
    if not senha:
        raise PdfError("Digite a senha atual do PDF.")

    try:
        leitor = PdfReader(origem)
    except Exception as erro:
        raise PdfError(
            f"Não foi possível ler «{os.path.basename(origem)}».\n\n{erro}"
        ) from None

    if not leitor.is_encrypted:
        raise PdfError(
            f"«{os.path.basename(origem)}» não está protegido por senha — "
            "não há nada para remover."
        )
    if not leitor.decrypt(senha):
        raise SenhaIncorretaError("A senha informada está incorreta.")

    escritor = PdfWriter()
    for pagina in leitor.pages:
        escritor.add_page(pagina)
    with _escrita_atomica(destino) as arquivo:
        escritor.write(arquivo)
