"""
SDD — Especificação: conversão Word → PDF multiplataforma

CONTRATO
  `docx_to_pdf(src, dest)` escolhe o motor de conversão pela plataforma:
  Microsoft Word (via `docx2pdf`) no Windows/macOS, LibreOffice headless no
  Linux. Quando o motor não está disponível, levanta `ConversionError` com uma
  mensagem JÁ PRONTA para exibir ao usuário — as abas repassam esse texto direto
  ao `messagebox`.

POR QUE EXISTE
  O projeto nasceu Windows-first e importava `docx2pdf` direto na aba de
  conversão. `docx2pdf` automatiza o Word por COM/AppleScript e NÃO funciona no
  Linux — instalava e falhava em runtime com erro incompreensível. Por isso o
  `requirements.txt` tem o marker `sys_platform != "linux"` e a escolha de motor
  virou este módulo.

REGRA DE NEGÓCIO
  - O LibreOffice ignora o nome de destino: grava `<basename>.pdf` no `--outdir`.
    O módulo é responsável por mover o resultado para o caminho pedido.
  - A conversão usa um perfil temporário próprio (`-env:UserInstallation`); sem
    isso ela falha em silêncio quando já há uma janela do LibreOffice aberta.
  - Mensagem de erro precisa conter o comando de instalação — o usuário final
    não sabe o que é "soffice".
"""

import subprocess
from pathlib import Path

import pytest

import docx_convert
from docx_convert import ConversionError


class TestLocalizacaoDoLibreOffice:
    def test_deve_encontrar_pelo_nome_soffice(self, monkeypatch):
        monkeypatch.setattr(
            docx_convert.shutil, "which", lambda n: "/usr/bin/soffice" if n == "soffice" else None
        )
        assert docx_convert._find_soffice() == "/usr/bin/soffice"

    def test_deve_cair_para_libreoffice_quando_soffice_nao_existe(self, monkeypatch):
        monkeypatch.setattr(
            docx_convert.shutil,
            "which",
            lambda n: "/usr/bin/libreoffice" if n == "libreoffice" else None,
        )
        assert docx_convert._find_soffice() == "/usr/bin/libreoffice"

    def test_deve_devolver_nulo_quando_nenhum_dos_dois_existe(self, monkeypatch):
        monkeypatch.setattr(docx_convert.shutil, "which", lambda _: None)
        assert docx_convert._find_soffice() is None


class TestConversaoPeloLibreOffice:
    def test_deve_orientar_a_instalacao_quando_o_libreoffice_falta(self, monkeypatch, tmp_path):
        monkeypatch.setattr(docx_convert, "_find_soffice", lambda: None)

        with pytest.raises(ConversionError) as erro:
            docx_convert._convert_with_libreoffice(str(tmp_path / "a.docx"), str(tmp_path / "a.pdf"))

        mensagem = str(erro.value)
        assert "LibreOffice" in mensagem
        assert "pacman" in mensagem and "apt" in mensagem, "precisa citar como instalar"

    def test_deve_mover_o_pdf_gerado_para_o_destino_pedido(self, monkeypatch, tmp_path):
        origem = tmp_path / "contrato.docx"
        origem.write_text("conteúdo")
        destino = tmp_path / "saida" / "resultado.pdf"
        destino.parent.mkdir()

        monkeypatch.setattr(docx_convert, "_find_soffice", lambda: "/usr/bin/soffice")

        def falso_run(cmd, **kwargs):
            # O LibreOffice grava <basename>.pdf no --outdir, ignorando o destino.
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            (outdir / "contrato.pdf").write_text("%PDF falso")
            return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

        monkeypatch.setattr(docx_convert.subprocess, "run", falso_run)

        docx_convert._convert_with_libreoffice(str(origem), str(destino))

        assert destino.exists(), "o PDF deveria ter sido movido para o caminho pedido"
        assert destino.read_text() == "%PDF falso"

    def test_deve_usar_perfil_temporario_proprio(self, monkeypatch, tmp_path):
        # Sem -env:UserInstallation a conversão falha em silêncio quando há uma
        # janela do LibreOffice aberta usando o perfil padrão.
        origem = tmp_path / "a.docx"
        origem.write_text("x")
        capturado = {}

        monkeypatch.setattr(docx_convert, "_find_soffice", lambda: "/usr/bin/soffice")

        def falso_run(cmd, **kwargs):
            capturado["cmd"] = cmd
            outdir = Path(cmd[cmd.index("--outdir") + 1])
            (outdir / "a.pdf").write_text("pdf")
            return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

        monkeypatch.setattr(docx_convert.subprocess, "run", falso_run)
        docx_convert._convert_with_libreoffice(str(origem), str(tmp_path / "out.pdf"))

        assert any(a.startswith("-env:UserInstallation=") for a in capturado["cmd"])
        assert "--headless" in capturado["cmd"]

    def test_deve_avisar_quando_o_libreoffice_nao_gerou_o_pdf(self, monkeypatch, tmp_path):
        origem = tmp_path / "a.docx"
        origem.write_text("x")

        monkeypatch.setattr(docx_convert, "_find_soffice", lambda: "/usr/bin/soffice")
        monkeypatch.setattr(
            docx_convert.subprocess,
            "run",
            lambda cmd, **kw: subprocess.CompletedProcess(cmd, 1, stdout="", stderr="erro do soffice"),
        )

        with pytest.raises(ConversionError) as erro:
            docx_convert._convert_with_libreoffice(str(origem), str(tmp_path / "out.pdf"))

        assert "erro do soffice" in str(erro.value), "o detalhe do motor deve chegar ao usuário"

    def test_deve_avisar_quando_a_conversao_estoura_o_tempo(self, monkeypatch, tmp_path):
        origem = tmp_path / "a.docx"
        origem.write_text("x")

        monkeypatch.setattr(docx_convert, "_find_soffice", lambda: "/usr/bin/soffice")

        def estoura(cmd, **kwargs):
            raise subprocess.TimeoutExpired(cmd, docx_convert.TIMEOUT)

        monkeypatch.setattr(docx_convert.subprocess, "run", estoura)

        with pytest.raises(ConversionError) as erro:
            docx_convert._convert_with_libreoffice(str(origem), str(tmp_path / "out.pdf"))

        assert str(docx_convert.TIMEOUT) in str(erro.value)


class TestEscolhaDoMotor:
    def test_deve_usar_libreoffice_no_linux(self, monkeypatch):
        chamados = []
        monkeypatch.setattr(docx_convert.sys, "platform", "linux")
        monkeypatch.setattr(
            docx_convert, "_convert_with_libreoffice", lambda s, d: chamados.append("libreoffice")
        )
        monkeypatch.setattr(docx_convert, "_convert_with_word", lambda s, d: chamados.append("word"))

        docx_convert.docx_to_pdf("a.docx", "a.pdf")

        assert chamados == ["libreoffice"]

    @pytest.mark.parametrize("plataforma", ["win32", "darwin"])
    def test_deve_usar_o_word_fora_do_linux(self, monkeypatch, plataforma):
        chamados = []
        monkeypatch.setattr(docx_convert.sys, "platform", plataforma)
        monkeypatch.setattr(
            docx_convert, "_convert_with_libreoffice", lambda s, d: chamados.append("libreoffice")
        )
        monkeypatch.setattr(docx_convert, "_convert_with_word", lambda s, d: chamados.append("word"))

        docx_convert.docx_to_pdf("a.docx", "a.pdf")

        assert chamados == ["word"]

    def test_deve_orientar_a_instalar_dependencia_quando_docx2pdf_falta(self, monkeypatch):
        # Cenário: Windows/macOS sem o pacote instalado.
        import builtins

        real_import = builtins.__import__

        def sem_docx2pdf(nome, *args, **kwargs):
            if nome == "docx2pdf":
                raise ImportError("No module named 'docx2pdf'")
            return real_import(nome, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", sem_docx2pdf)

        with pytest.raises(ConversionError) as erro:
            docx_convert._convert_with_word("a.docx", "a.pdf")

        assert "requirements.txt" in str(erro.value)
