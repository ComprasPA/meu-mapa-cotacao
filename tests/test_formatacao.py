"""Testes das funções puras de formatação/normalização de logica_cotacao.py."""
import math

import pandas as pd
import pytest

from logica_cotacao import (
    limpar_valor,
    normalizar_codigo,
    formatar_brl,
    formatar_qtd,
    formatar_pct,
    formatar_pct_com_seta,
    limpar_texto_pdf,
    _var_fill,
    GREEN,
    YELLOW,
    RED,
    GRAY,
)


# ------------------------------------------------------------------ limpar_valor
class TestLimparValor:
    def test_none_e_nan_viram_zero(self):
        assert limpar_valor(None) == 0.0
        assert limpar_valor(float('nan')) == 0.0
        assert limpar_valor(pd.NA) == 0.0

    def test_string_vazia_vira_zero(self):
        assert limpar_valor("") == 0.0
        assert limpar_valor("   ") == 0.0

    @pytest.mark.parametrize("lixo", [
        "Total", "TOTAL ITEM", "##########", "A VISTA", "25 dias",
        "Item", "Código", "Produto", "Descrição", "nan",
    ])
    def test_textos_de_cabecalho_e_rodape_viram_zero(self, lixo):
        assert limpar_valor(lixo) == 0.0

    def test_formato_br_milhar_ponto_decimal_virgula(self):
        assert limpar_valor("1.234,56") == 1234.56

    def test_formato_us_milhar_virgula_decimal_ponto(self):
        # vírgula antes do ponto -> tratado como separador de milhar (US)
        assert limpar_valor("1,234.56") == 1234.56

    def test_apenas_virgula_decimal_br(self):
        assert limpar_valor("123,45") == 123.45

    def test_multiplos_pontos_sem_virgula_vira_so_milhar(self):
        assert limpar_valor("1.234.567") == 1234567.0

    def test_com_prefixo_moeda_rs(self):
        assert limpar_valor("R$ 1.500,00") == 1500.0

    def test_numero_puro_float_ou_int(self):
        assert limpar_valor(42) == 42.0
        assert limpar_valor(42.5) == 42.5

    def test_texto_nao_numerico_vira_zero(self):
        assert limpar_valor("abc") == 0.0


# ------------------------------------------------------------------ normalizar_codigo
class TestNormalizarCodigo:
    def test_com_e_sem_zeros_a_esquerda_viram_o_mesmo_codigo(self):
        assert normalizar_codigo("1268") == normalizar_codigo("0000001268")
        assert normalizar_codigo("1268") == "0000001268"

    def test_extrai_apenas_digitos(self):
        assert normalizar_codigo("PROD-1268/A") == "0000001268"

    def test_nan_vira_none(self):
        assert normalizar_codigo(float('nan')) is None

    def test_sem_nenhum_digito_vira_none(self):
        assert normalizar_codigo("ABC") is None

    def test_codigo_maior_que_10_digitos_nao_trunca(self):
        assert normalizar_codigo("12345678901") == "12345678901"


# ------------------------------------------------------------------ formatar_brl
class TestFormatarBrl:
    def test_valor_vazio_none_ou_nan(self):
        assert formatar_brl("") == ""
        assert formatar_brl(None) == ""
        assert formatar_brl(float('nan')) == ""

    def test_valor_zero_ou_negativo_vira_vazio(self):
        assert formatar_brl(0) == ""
        assert formatar_brl(-10) == ""

    def test_formato_moeda_br(self):
        assert formatar_brl(1234.5) == "R$ 1.234,50"
        assert formatar_brl(10) == "R$ 10,00"

    def test_texto_nao_conversivel_vira_vazio(self):
        assert formatar_brl("abc") == ""


# ------------------------------------------------------------------ formatar_qtd
class TestFormatarQtd:
    def test_inteiro_sem_casas_decimais(self):
        assert formatar_qtd(1000) == "1.000"
        assert formatar_qtd(5.0) == "5"

    def test_fracionario_com_duas_casas(self):
        assert formatar_qtd(1234.5) == "1.234,50"

    def test_nao_numerico_vira_zero_formatado(self):
        assert formatar_qtd("abc") == "0"


# ------------------------------------------------------------------ formatar_pct
class TestFormatarPct:
    def test_vazio_none_nan(self):
        assert formatar_pct("") == ""
        assert formatar_pct(None) == ""
        assert formatar_pct(float('nan')) == ""

    def test_positivo_tem_sinal_de_mais(self):
        assert formatar_pct(5.5) == "+5,50%"

    def test_negativo_tem_sinal_de_menos(self):
        assert formatar_pct(-3.2) == "-3,20%"

    def test_zero(self):
        assert formatar_pct(0) == "+0,00%"


# ------------------------------------------------------------------ formatar_pct_com_seta
class TestFormatarPctComSeta:
    def test_vazio(self):
        assert formatar_pct_com_seta("") == ""

    def test_positivo_seta_vermelha_para_cima(self):
        html = formatar_pct_com_seta(7.5)
        assert "↑" in html
        assert "#c00000" in html
        assert "+7,50%" in html

    def test_negativo_seta_verde_para_baixo(self):
        html = formatar_pct_com_seta(-7.5)
        assert "↓" in html
        assert "#2ca02c" in html

    def test_zero_sem_seta(self):
        html = formatar_pct_com_seta(0)
        assert "↑" not in html
        assert "↓" not in html
        assert "#555555" in html


# ------------------------------------------------------------------ limpar_texto_pdf
class TestLimparTextoPdf:
    def test_remove_acentos(self):
        assert limpar_texto_pdf("Cotação") == "Cotacao"
        assert limpar_texto_pdf("Descrição") == "Descricao"

    def test_converte_nao_string_em_string(self):
        assert limpar_texto_pdf(123) == "123"

    def test_caracter_fora_do_latin1_vira_substituto_sem_quebrar(self):
        # '€' não existe em latin-1 puro; a função não deve lançar exceção.
        resultado = limpar_texto_pdf("preço: 10€")
        assert isinstance(resultado, str)


# ------------------------------------------------------------------ _var_fill
class TestVarFill:
    def test_vazio_ou_nan_vira_cinza(self):
        assert _var_fill("") == GRAY
        assert _var_fill(float('nan')) == GRAY

    def test_abaixo_de_menos_5_vira_verde(self):
        assert _var_fill(-5) == GREEN
        assert _var_fill(-20) == GREEN

    def test_acima_de_5_vira_vermelho(self):
        assert _var_fill(5) == RED
        assert _var_fill(20) == RED

    def test_dentro_da_faixa_vira_amarelo(self):
        assert _var_fill(0) == YELLOW
        assert _var_fill(4.9) == YELLOW
        assert _var_fill(-4.9) == YELLOW
