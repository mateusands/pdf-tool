"""
SDD — Especificação: operações de arquivo PDF (camada pura, sem UI)

CONTRATO
  `pdf_io` concentra TODA escrita de PDF do app. Cada operação recebe caminhos e
  parâmetros, devolve nada e levanta `PdfError` com uma mensagem JÁ PRONTA para
  exibir ao usuário — as abas repassam esse texto direto ao `messagebox`.

  Toda operação que escreve chama `validar_destino()` ANTES de abrir o arquivo de
  saída. Esse é o único ponto de checagem: nenhuma aba pode escrever sem passar
  por aqui.

POR QUE EXISTE
  As abas misturavam UI e processamento, então a regra de arquivo não era
  testável e cada aba reimplementava a sua. O resultado foi um bug de perda de
  dado: nenhuma das seis abas que escrevem comparava o destino com a origem.

  Como o `PdfReader` do pypdf carrega o arquivo inteiro em memória antes de
  escrever, salvar por cima da origem NÃO dá erro — grava em silêncio. Um PDF de
  3 páginas dividido em cima de si mesmo virava um PDF de 1 página, e o original
  sumia. O `asksaveasfilename` abre na pasta da origem com o arquivo listado, ou
  seja, clicar nele é um gesto natural do usuário.

REGRA DE NEGÓCIO
  - O destino NUNCA pode ser uma das origens. A comparação é por caminho real
    (`realpath`), então caminho relativo e link simbólico também são pegos.
  - PDF protegido por senha precisa de aviso próprio, não do erro cru do motor.
    `fitz` levanta "document closed or encrypted", que não diz nada ao usuário.
  - Proteger com senha usa AES-256. O default do pypdf é RC4-128, que é
    criptografia quebrada — dá falsa sensação de segurança numa ferramenta cujo
    propósito É proteger o arquivo.
  - Desbloquear um PDF que não tem senha é engano do usuário, não sucesso: avisa
    em vez de gerar uma cópia idêntica silenciosamente.
"""

import fitz
import pytest
from pypdf import PdfReader

from pdf_tool.core import pdf_io
from pdf_tool.core.pdf_io import (
    DestinoInvalidoError,
    PdfProtegidoError,
    SenhaIncorretaError,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────

def _criar_pdf(caminho, n_paginas=3):
    doc = fitz.open()
    for i in range(n_paginas):
        pagina = doc.new_page()
        pagina.insert_text((72, 72), f"pagina {i + 1}")
    doc.save(str(caminho))
    doc.close()
    return str(caminho)


@pytest.fixture
def pdf_de_3_paginas(tmp_path):
    return _criar_pdf(tmp_path / "contrato.pdf", 3)


@pytest.fixture
def pdf_protegido(tmp_path, pdf_de_3_paginas):
    destino = str(tmp_path / "protegido.pdf")
    pdf_io.proteger_pdf(pdf_de_3_paginas, destino, "segredo")
    return destino


# ── Destino ≠ origem ──────────────────────────────────────────────────────────

class TestDestinoDiferenteDaOrigem:
    def test_deve_aceitar_quando_o_destino_e_outro_arquivo(self, tmp_path, pdf_de_3_paginas):
        pdf_io.validar_destino(str(tmp_path / "recorte.pdf"), pdf_de_3_paginas)

    def test_deve_recusar_quando_o_destino_e_a_propria_origem(self, pdf_de_3_paginas):
        with pytest.raises(DestinoInvalidoError):
            pdf_io.validar_destino(pdf_de_3_paginas, pdf_de_3_paginas)

    def test_deve_recusar_quando_o_destino_e_a_origem_por_caminho_relativo(
        self, tmp_path, pdf_de_3_paginas, monkeypatch
    ):
        # O usuário digitou "./contrato.pdf" no diálogo de salvar.
        monkeypatch.chdir(tmp_path)

        with pytest.raises(DestinoInvalidoError):
            pdf_io.validar_destino("./contrato.pdf", pdf_de_3_paginas)

    def test_deve_recusar_quando_o_destino_e_a_origem_por_link_simbolico(
        self, tmp_path, pdf_de_3_paginas
    ):
        atalho = tmp_path / "atalho.pdf"
        atalho.symlink_to(pdf_de_3_paginas)

        with pytest.raises(DestinoInvalidoError):
            pdf_io.validar_destino(str(atalho), pdf_de_3_paginas)

    def test_deve_recusar_quando_o_destino_e_uma_das_varias_origens(self, tmp_path):
        # Cenário do Juntar: o destino coincide com o 2º da lista.
        a = _criar_pdf(tmp_path / "a.pdf", 1)
        b = _criar_pdf(tmp_path / "b.pdf", 1)

        with pytest.raises(DestinoInvalidoError):
            pdf_io.validar_destino(b, a, b)

    def test_deve_citar_o_nome_do_arquivo_na_mensagem(self, pdf_de_3_paginas):
        with pytest.raises(DestinoInvalidoError) as erro:
            pdf_io.validar_destino(pdf_de_3_paginas, pdf_de_3_paginas)

        assert "contrato.pdf" in str(erro.value), "o usuário precisa saber qual arquivo"

    def test_deve_aceitar_quando_o_destino_ainda_nao_existe(self, tmp_path, pdf_de_3_paginas):
        pdf_io.validar_destino(str(tmp_path / "nova" / "saida.pdf"), pdf_de_3_paginas)


# ── Integridade: a origem sobrevive a toda operação ───────────────────────────

class TestAOrigemNuncaEDestruida:
    @pytest.mark.parametrize(
        "operacao",
        [
            lambda src, dest: pdf_io.dividir_pdf(src, dest, [0]),
            lambda src, dest: pdf_io.girar_pdf(src, dest, [0], 90),
            lambda src, dest: pdf_io.reorganizar_pdf(src, dest, [2, 1, 0]),
            lambda src, dest: pdf_io.proteger_pdf(src, dest, "x"),
            lambda src, dest: pdf_io.comprimir_pdf(src, dest, garbage=3, deflate=True),
            lambda src, dest: pdf_io.juntar_pdfs([src], dest),
        ],
    )
    def test_deve_recusar_a_operacao_quando_o_destino_e_a_origem(
        self, pdf_de_3_paginas, operacao
    ):
        antes = open(pdf_de_3_paginas, "rb").read()

        with pytest.raises(DestinoInvalidoError):
            operacao(pdf_de_3_paginas, pdf_de_3_paginas)

        assert open(pdf_de_3_paginas, "rb").read() == antes, "a origem foi alterada"

    def test_nao_deve_deixar_arquivo_pela_metade_quando_a_escrita_falha(
        self, tmp_path, pdf_de_3_paginas, monkeypatch
    ):
        # Disco cheio no meio da gravação: melhor nenhum arquivo do que um PDF
        # truncado que o usuário acha que salvou.
        destino = tmp_path / "recorte.pdf"

        def escrita_que_falha(self, stream):
            stream.write(b"%PDF-1.7 truncado")
            raise OSError("No space left on device")

        monkeypatch.setattr(pdf_io.PdfWriter, "write", escrita_que_falha)

        with pytest.raises(OSError):
            pdf_io.dividir_pdf(pdf_de_3_paginas, str(destino), [0])

        assert not destino.exists(), "não pode sobrar PDF truncado no destino"
        assert list(tmp_path.glob("*.part")) == [], "não pode sobrar arquivo temporário"


# ── Operações ─────────────────────────────────────────────────────────────────

class TestDividirPdf:
    def test_deve_salvar_apenas_as_paginas_pedidas(self, tmp_path, pdf_de_3_paginas):
        destino = str(tmp_path / "recorte.pdf")

        pdf_io.dividir_pdf(pdf_de_3_paginas, destino, [0, 2])

        assert len(PdfReader(destino).pages) == 2

    def test_deve_respeitar_a_ordem_crescente_das_paginas(self, tmp_path, pdf_de_3_paginas):
        destino = str(tmp_path / "recorte.pdf")

        pdf_io.dividir_pdf(pdf_de_3_paginas, destino, [2, 0])

        texto = [p.extract_text().strip() for p in PdfReader(destino).pages]
        assert texto == ["pagina 1", "pagina 3"]

    def test_deve_avisar_quando_nenhuma_pagina_foi_escolhida(self, tmp_path, pdf_de_3_paginas):
        with pytest.raises(pdf_io.PdfError):
            pdf_io.dividir_pdf(pdf_de_3_paginas, str(tmp_path / "x.pdf"), [])


class TestJuntarPdfs:
    def test_deve_somar_as_paginas_na_ordem_da_lista(self, tmp_path):
        a = _criar_pdf(tmp_path / "a.pdf", 2)
        b = _criar_pdf(tmp_path / "b.pdf", 1)
        destino = str(tmp_path / "junto.pdf")

        pdf_io.juntar_pdfs([a, b], destino)

        assert len(PdfReader(destino).pages) == 3

    def test_deve_avisar_quando_a_lista_esta_vazia(self, tmp_path):
        with pytest.raises(pdf_io.PdfError):
            pdf_io.juntar_pdfs([], str(tmp_path / "x.pdf"))


class TestGirarPdf:
    def test_deve_girar_somente_as_paginas_escolhidas(self, tmp_path, pdf_de_3_paginas):
        destino = str(tmp_path / "girado.pdf")

        pdf_io.girar_pdf(pdf_de_3_paginas, destino, [1], 90)

        rotacoes = [p.rotation for p in PdfReader(destino).pages]
        assert rotacoes == [0, 90, 0]


class TestReorganizarPdf:
    def test_deve_gravar_as_paginas_na_ordem_pedida(self, tmp_path, pdf_de_3_paginas):
        destino = str(tmp_path / "nova_ordem.pdf")

        pdf_io.reorganizar_pdf(pdf_de_3_paginas, destino, [2, 0, 1])

        texto = [p.extract_text().strip() for p in PdfReader(destino).pages]
        assert texto == ["pagina 3", "pagina 1", "pagina 2"]


# ── Senha ─────────────────────────────────────────────────────────────────────

class TestProtegerComSenha:
    def test_deve_usar_aes_256_e_nao_rc4(self, pdf_protegido):
        # RC4-128 (o default do pypdf) grava V=2/R=3 e é criptografia quebrada.
        encrypt = PdfReader(pdf_protegido).trailer["/Encrypt"].get_object()

        assert encrypt.get("/V") == 5 and encrypt.get("/R") == 6, "deveria ser AES-256"

    def test_deve_exigir_a_senha_para_abrir(self, pdf_protegido):
        assert PdfReader(pdf_protegido).is_encrypted
        assert len(PdfReader(pdf_protegido, password="segredo").pages) == 3

    def test_deve_recusar_senha_vazia(self, tmp_path, pdf_de_3_paginas):
        with pytest.raises(pdf_io.PdfError):
            pdf_io.proteger_pdf(pdf_de_3_paginas, str(tmp_path / "x.pdf"), "")


class TestDesbloquearPdf:
    def test_deve_gerar_uma_copia_sem_senha(self, tmp_path, pdf_protegido):
        destino = str(tmp_path / "livre.pdf")

        pdf_io.desbloquear_pdf(pdf_protegido, destino, "segredo")

        assert not PdfReader(destino).is_encrypted
        assert len(PdfReader(destino).pages) == 3

    def test_deve_avisar_quando_a_senha_esta_errada(self, tmp_path, pdf_protegido):
        with pytest.raises(SenhaIncorretaError):
            pdf_io.desbloquear_pdf(pdf_protegido, str(tmp_path / "x.pdf"), "errada")

    def test_deve_avisar_quando_o_pdf_nao_tem_senha(self, tmp_path, pdf_de_3_paginas):
        # Engano do usuário — gerar uma cópia idêntica em silêncio esconde o erro.
        with pytest.raises(pdf_io.PdfError) as erro:
            pdf_io.desbloquear_pdf(pdf_de_3_paginas, str(tmp_path / "x.pdf"), "qualquer")

        assert "senha" in str(erro.value).lower()

    def test_nao_deve_deixar_arquivo_pela_metade_quando_a_senha_falha(
        self, tmp_path, pdf_protegido
    ):
        destino = tmp_path / "x.pdf"

        with pytest.raises(SenhaIncorretaError):
            pdf_io.desbloquear_pdf(pdf_protegido, str(destino), "errada")

        assert not destino.exists(), "não pode sobrar arquivo vazio no disco"


# ── Leitura para miniaturas ───────────────────────────────────────────────────

class TestAberturaParaMiniaturas:
    def test_deve_abrir_um_pdf_normal(self, pdf_de_3_paginas):
        with pdf_io.abrir_documento(pdf_de_3_paginas) as doc:
            assert len(doc) == 3

    def test_deve_avisar_que_o_pdf_esta_protegido(self, pdf_protegido):
        # `fitz` levanta "document closed or encrypted", que não diz nada ao usuário.
        with pytest.raises(PdfProtegidoError) as erro:
            with pdf_io.abrir_documento(pdf_protegido):
                pass

        mensagem = str(erro.value).lower()
        assert "senha" in mensagem
        assert "desbloquear" in mensagem, "precisa apontar a aba que resolve"

    def test_deve_avisar_quando_o_arquivo_nao_e_um_pdf(self, tmp_path):
        falso = tmp_path / "foto.pdf"
        falso.write_text("isto nao e um pdf")

        with pytest.raises(pdf_io.PdfError):
            with pdf_io.abrir_documento(str(falso)):
                pass


class TestMiniaturas:
    def test_deve_gerar_uma_miniatura_por_pagina(self, pdf_de_3_paginas):
        miniaturas = pdf_io.miniaturas_ppm(pdf_de_3_paginas, 88, 124)

        assert len(miniaturas) == 3
        assert all(m.startswith(b"P6") for m in miniaturas), "deveria ser PPM binário"

    def test_deve_falhar_de_imediato_com_pdf_protegido(self, pdf_protegido):
        # Eager de propósito: se fosse gerador, o erro só apareceria ao iterar —
        # longe do `try` da aba, e o usuário não veria aviso nenhum.
        with pytest.raises(PdfProtegidoError):
            pdf_io.miniaturas_ppm(pdf_protegido, 88, 124)
