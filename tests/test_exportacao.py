"""Testes de gerar_pdf/gerar_excel (logica_cotacao.py): não validam o layout
pixel a pixel, só que os bytes produzidos são um PDF/XLSX válido e que os
valores essenciais (preço cotado, variação, observação) aparecem na
planilha exportada — o suficiente para pegar uma regressão que quebre a
geração do arquivo ou perca uma coluna inteira."""
import io

import pandas as pd
import openpyxl
import pytest

from logica_cotacao import gerar_pdf, gerar_excel


def _df_final_exemplo():
    return pd.DataFrame([{
        'Item': '1',
        'Código': '0000001268',
        'Descrição': 'Parafuso Sextavado 1/4',
        'Unidade': 'UN',
        'Qtd': 100,
        'Fornecedor Cotado': 'Fornecedor Barato',
        'Valor Cotado (R$)': 9.90,
        'Último Preço Pago (R$)': 10.0,
        'Data Última Compra': pd.Timestamp('2026-01-15'),
        'Fornecedor Última Compra': 'Fornecedor Antigo',
        'Preço Médio (R$)': 10.0,
        'Preço Mín. Histórico (R$)': 8.0,
        'Preço Máx. Histórico (R$)': 12.0,
        'Var. vs Último (%)': -1.0,
        'Var. vs Médio (%)': -1.0,
        'Observação': 'Próximo da média histórica',
        'Representatividade (%)': 66.44,
    }, {
        'Item': '2',
        'Código': '0000009999',
        'Descrição': 'Porca Sextavada 1/4',
        'Unidade': 'UN',
        'Qtd': 200,
        'Fornecedor Cotado': 'Fornecedor X',
        'Valor Cotado (R$)': 2.5,
        'Último Preço Pago (R$)': "",
        'Data Última Compra': "",
        'Fornecedor Última Compra': "",
        'Preço Médio (R$)': "",
        'Preço Mín. Histórico (R$)': "",
        'Preço Máx. Histórico (R$)': "",
        'Var. vs Último (%)': "",
        'Var. vs Médio (%)': "",
        'Observação': 'Sem histórico de compra',
        'Representatividade (%)': 33.56,
    }])


class TestGerarPdf:
    def test_gera_bytes_de_pdf_valido(self):
        pdf_bytes = gerar_pdf(_df_final_exemplo())
        assert isinstance(pdf_bytes, (bytes, bytearray))
        assert pdf_bytes[:5] == b"%PDF-"
        assert len(pdf_bytes) > 500

    def test_df_com_uma_linha_so(self):
        df = _df_final_exemplo().iloc[[0]]
        pdf_bytes = gerar_pdf(df)
        assert pdf_bytes[:5] == b"%PDF-"

    def test_item_sem_historico_nao_quebra_geracao(self):
        # Linha 2 do exemplo já cobre "" em todos os campos de histórico —
        # esse teste garante que sozinha (sem a linha com histórico
        # completo) ela também não derruba a geração do PDF.
        df = _df_final_exemplo().iloc[[1]]
        pdf_bytes = gerar_pdf(df)
        assert pdf_bytes[:5] == b"%PDF-"

    def test_com_numero_de_cotacao_continua_gerando_pdf_valido(self):
        pdf_bytes = gerar_pdf(_df_final_exemplo(), numero_cotacao='021132')
        assert pdf_bytes[:5] == b"%PDF-"
        assert len(pdf_bytes) > 500


class TestGerarExcel:
    def test_gera_xlsx_valido_e_recarregavel(self):
        xlsx_bytes = gerar_excel(_df_final_exemplo())
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Comparativo_Cotacao"]

        cabecalho = [c.value for c in ws[1]]
        assert cabecalho[0] == 'Item'
        assert 'Valor Cotado (R$)' in cabecalho
        assert 'Observação' in cabecalho

    def test_valores_das_linhas_batem_com_o_dataframe(self):
        xlsx_bytes = gerar_excel(_df_final_exemplo())
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Comparativo_Cotacao"]

        linha1 = [c.value for c in ws[2]]
        assert linha1[0] == '1'
        assert linha1[6] == pytest.approx(9.90)  # Valor Cotado (R$)

        linha2 = [c.value for c in ws[3]]
        assert linha2[6] == pytest.approx(2.5)
        assert linha2[7] is None  # Último Preço Pago (R$) vazio -> None na planilha

    def test_df_vazio_nao_quebra_geracao(self):
        colunas = ['Item', 'Código', 'Descrição', 'Unidade', 'Qtd',
                   'Fornecedor Cotado', 'Valor Cotado (R$)',
                   'Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
                   'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)',
                   'Var. vs Último (%)', 'Var. vs Médio (%)', 'Observação']
        df_vazio = pd.DataFrame(columns=colunas)
        xlsx_bytes = gerar_excel(df_vazio)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        assert wb["Comparativo_Cotacao"].max_row >= 1


class TestGerarExcelComNumeroCotacao:
    """Sem numero_cotacao o layout não muda (ver testes acima: cabeçalho
    sempre na linha 1) — só quando informado é que entra uma linha de
    título acima, empurrando cabeçalho/dados uma linha pra baixo."""

    def test_sem_numero_cotacao_cabecalho_continua_na_linha_1(self):
        xlsx_bytes = gerar_excel(_df_final_exemplo(), numero_cotacao=None)
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Comparativo_Cotacao"]
        assert ws.cell(row=1, column=1).value == 'Item'

    def test_com_numero_cotacao_linha_1_e_titulo_e_cabecalho_vai_pra_linha_2(self):
        xlsx_bytes = gerar_excel(_df_final_exemplo(), numero_cotacao='021132')
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        ws = wb["Comparativo_Cotacao"]

        assert ws.cell(row=1, column=1).value == 'Mapa de Cotação nº 021132'
        assert ws.cell(row=2, column=1).value == 'Item'

        linha1_dados = [c.value for c in ws[3]]
        assert linha1_dados[0] == '1'
        assert linha1_dados[6] == pytest.approx(9.90)  # Valor Cotado (R$)
