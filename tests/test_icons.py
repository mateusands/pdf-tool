"""
SDD — Especificação: ícones vetoriais da interface

CONTRATO
  `renderizar_png(nome, tamanho, cor)` devolve os bytes de um PNG quadrado de
  `tamanho`×`tamanho` pixels, com fundo transparente e o traço na `cor` pedida.
  O resultado vai direto para `tk.PhotoImage(data=…)`.

  `nomes()` lista os ícones disponíveis. Nome desconhecido é erro de programação,
  não do usuário: levanta `ValueError` na hora, em vez de desenhar um quadrado
  vazio que ninguém percebe.

POR QUE EXISTE
  A interface usava emoji como ícone ("✂️ Dividir", "🔒 Proteger"). Emoji não é
  ícone de produto: o desenho muda em cada sistema operacional, a cor é fixa
  (não acompanha estado de foco, hover ou desabilitado), o alinhamento vertical
  com o texto é imprevisível, e no Linux ele depende de uma fonte de emoji estar
  instalada — sem ela vira retângulo vazio.

  A saída é o conjunto Lucide (ISC), desenhado numa grade 24×24 com traço de 2.
  Como os SVGs usam `stroke="currentColor"`, dá para trocar a cor em tempo de
  execução. O PyMuPDF — que já é dependência do projeto para miniaturas — sabe
  rasterizar SVG, então isso não custa nenhum pacote novo nem fonte instalada.

REGRA DE NEGÓCIO
  - Fundo transparente é obrigatório: o mesmo ícone é desenhado sobre a sidebar,
    sobre o card e sobre o botão azul, que têm fundos diferentes.
  - A cor é parâmetro, não constante: é ela que faz o ícone do item ativo
    contrastar com o dos inativos.
  - Nenhum emoji sobra na interface — há um teste que varre o pacote.
"""

import re

import fitz
import pytest

from pdf_tool.core import icons


class TestRenderizacao:
    def test_deve_devolver_um_png(self):
        dados = icons.renderizar_png("scissors", 20, "#FFFFFF")

        assert dados[:8] == b"\x89PNG\r\n\x1a\n"

    @pytest.mark.parametrize("tamanho", [16, 20, 24, 32])
    def test_deve_respeitar_o_tamanho_pedido(self, tamanho):
        dados = icons.renderizar_png("lock", tamanho, "#FFFFFF")

        imagem = fitz.Pixmap(dados)
        assert (imagem.width, imagem.height) == (tamanho, tamanho)

    def test_deve_ter_fundo_transparente(self):
        # O mesmo ícone é desenhado sobre a sidebar, o card e o botão azul.
        imagem = fitz.Pixmap(icons.renderizar_png("lock", 24, "#FFFFFF"))

        assert imagem.alpha, "sem canal alfa o ícone ganharia um quadrado de fundo"

    def test_deve_aplicar_a_cor_pedida(self):
        imagem = fitz.Pixmap(icons.renderizar_png("check", 24, "#2979FF"))

        pixels = imagem.samples
        largura_do_pixel = imagem.n
        cores = {
            tuple(pixels[i:i + 3])
            for i in range(0, len(pixels), largura_do_pixel)
            if pixels[i + 3] > 200          # só o traço, ignorando a borda suavizada
        }
        assert (0x29, 0x79, 0xFF) in cores

    def test_deve_desenhar_cores_diferentes_para_estados_diferentes(self):
        ativo = icons.renderizar_png("scissors", 20, "#FFFFFF")
        inativo = icons.renderizar_png("scissors", 20, "#8B949E")

        assert ativo != inativo, "o ícone do item ativo precisa contrastar com o inativo"


class TestCatalogo:
    def test_deve_listar_os_icones_disponiveis(self):
        assert "scissors" in icons.nomes()
        assert len(icons.nomes()) >= 20

    def test_deve_recusar_um_nome_desconhecido(self):
        # Erro de programação: melhor estourar do que desenhar um vazio silencioso.
        with pytest.raises(ValueError) as erro:
            icons.renderizar_png("nao-existe", 20, "#FFFFFF")

        assert "nao-existe" in str(erro.value)

    def test_todo_icone_do_catalogo_deve_renderizar(self):
        for nome in icons.nomes():
            assert icons.renderizar_png(nome, 20, "#FFFFFF")[:8] == b"\x89PNG\r\n\x1a\n", nome


class TestIconeDaJanela:
    """O ícone que o sistema operacional mostra na barra de título, na barra de
    tarefas e no alt-tab. Sem ele o app aparece com o losango genérico do Tk."""

    def test_deve_devolver_um_png_quadrado_no_tamanho_pedido(self):
        imagem = fitz.Pixmap(icons.renderizar_icone_do_app(64, "#3B82F6", "#FFFFFF"))

        assert (imagem.width, imagem.height) == (64, 64)

    def test_deve_pintar_o_fundo_com_a_cor_da_marca(self):
        imagem = fitz.Pixmap(icons.renderizar_icone_do_app(64, "#3B82F6", "#FFFFFF"))

        # Um ponto perto da borda esquerda, na altura do meio: é fundo, não glifo.
        pixel = imagem.pixel(6, 32)
        assert pixel[:3] == (0x3B, 0x82, 0xF6)

    def test_deve_ter_cantos_arredondados_transparentes(self):
        # Canto quadrado num ícone de app fica com aparência de recorte errado.
        imagem = fitz.Pixmap(icons.renderizar_icone_do_app(64, "#3B82F6", "#FFFFFF"))

        assert imagem.alpha
        assert imagem.pixel(0, 0)[3] == 0, "o canto deveria ser transparente"

    @pytest.mark.parametrize("tamanho", [16, 32, 64, 128])
    def test_deve_servir_nos_tamanhos_que_o_sistema_pede(self, tamanho):
        dados = icons.renderizar_icone_do_app(tamanho, "#3B82F6", "#FFFFFF")

        assert fitz.Pixmap(dados).width == tamanho


class TestInterfaceSemEmoji:
    def test_nao_deve_sobrar_emoji_no_pacote(self):
        """Emoji não é ícone de produto: muda de desenho por sistema, tem cor fixa
        e no Linux depende de fonte instalada. Este teste é o guarda-corpo."""
        import pathlib

        # Faixas de emoji e símbolos pictográficos do Unicode.
        emoji = re.compile(
            "[\U0001F300-\U0001FAFF"   # símbolos e pictogramas
            "\U00002600-\U000027BF"    # dingbats e símbolos diversos
            "\U0001F000-\U0001F02F"    # peças de mahjong
            "\U0000FE0F"               # seletor de variação (o "torna emoji")
            "]"
        )
        raiz = pathlib.Path(__file__).parent.parent / "pdf_tool"

        encontrados = []
        for arquivo in sorted(raiz.rglob("*.py")):
            for n, linha in enumerate(arquivo.read_text().splitlines(), 1):
                if emoji.search(linha):
                    encontrados.append(f"{arquivo.relative_to(raiz)}:{n}: {linha.strip()[:60]}")

        assert not encontrados, "emoji na interface:\n" + "\n".join(encontrados)
