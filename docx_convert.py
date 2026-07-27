"""Conversão Word → PDF multiplataforma.

No Windows e no macOS usa `docx2pdf`, que automatiza o Microsoft Word instalado.
No Linux o Word não existe, então usa o LibreOffice em modo headless.
"""

import os
import shutil
import subprocess
import sys
import tempfile

TIMEOUT = 180


class ConversionError(RuntimeError):
    """Erro de conversão com mensagem já pronta para exibir ao usuário."""


def _find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _convert_with_libreoffice(src, dest):
    soffice = _find_soffice()
    if not soffice:
        raise ConversionError(
            "LibreOffice não encontrado — ele é necessário para converter "
            "Word → PDF no Linux.\n\n"
            "Arch/CachyOS:  sudo pacman -S libreoffice-fresh\n"
            "Debian/Ubuntu: sudo apt install libreoffice-writer"
        )

    with tempfile.TemporaryDirectory() as tmp:
        # Perfil próprio: sem isso a conversão falha em silêncio se já houver
        # uma janela do LibreOffice aberta usando o perfil padrão.
        profile = os.path.join(tmp, "profile")
        try:
            proc = subprocess.run(
                [
                    soffice,
                    f"-env:UserInstallation=file://{profile}",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", tmp,
                    src,
                ],
                capture_output=True,
                text=True,
                timeout=TIMEOUT,
            )
        except subprocess.TimeoutExpired:
            raise ConversionError(
                f"O LibreOffice não respondeu em {TIMEOUT}s e a conversão foi cancelada."
            ) from None

        # O LibreOffice ignora o nome de destino: grava <basename>.pdf no --outdir.
        produced = os.path.join(
            tmp, os.path.splitext(os.path.basename(src))[0] + ".pdf"
        )
        if not os.path.exists(produced):
            detail = (proc.stderr or proc.stdout or "").strip()
            raise ConversionError(
                f"O LibreOffice não gerou o PDF.\n\n{detail}"
                if detail
                else "O LibreOffice não gerou o PDF."
            )

        shutil.move(produced, dest)


def _convert_with_word(src, dest):
    try:
        from docx2pdf import convert
    except ImportError:
        raise ConversionError(
            "Pacote `docx2pdf` não instalado. Rode:\n\n"
            "    pip install -r requirements.txt"
        ) from None

    convert(src, dest)


def docx_to_pdf(src, dest):
    """Converte o .docx em `src` para um PDF em `dest`.

    Levanta ConversionError com uma mensagem legível quando o backend da
    plataforma não está disponível.
    """
    if sys.platform.startswith("linux"):
        _convert_with_libreoffice(src, dest)
    else:
        _convert_with_word(src, dest)
