"""Testes de consolidar_base_precos (parte pura de construir_base_precos em
app.py): recebe o histórico de Pedidos já lido (colunas internas) e devolve
a base consolidada por item + o histórico linha a linha já normalizado.
"""
import pandas as pd
import pytest

from logica_cotacao import consolidar_base_precos


def _historico_bruto(linhas):
    """Monta um DataFrame de histórico com as colunas internas esperadas,
    igual ao que _ler_historico_bruto() devolveria depois de renomear as
    colunas da aba "Pedidos" (ver MAPA_COLUNAS_PEDIDOS em app.py)."""
    colunas = ['Produto', 'Descricao.1', 'Unidade', 'Prc Unitario',
               'Data Emissao', 'Nome Fornece', 'Quantidade', 'Status Aprov']
    return pd.DataFrame(linhas, columns=colunas)


class TestConsolidarBasePrecos:
    def test_historico_vazio_retorna_dataframes_vazios(self):
        df = _historico_bruto([])
        base_precos, df_f = consolidar_base_precos(df)
        assert base_precos.empty
        assert df_f.empty

    def test_um_unico_item_uma_unica_compra(self):
        df = _historico_bruto([
            ['1268', 'Parafuso Sextavado', 'UN', '10,50', '05/03/2026', 'Fornecedor A', '100', 'Aprovado'],
        ])
        base_precos, df_f = consolidar_base_precos(df)

        assert len(base_precos) == 1
        linha = base_precos.iloc[0]
        assert linha['Cod_Norm'] == '0000001268'
        assert linha['Ultimo_Preco'] == 10.5
        assert linha['Preco_Medio'] == 10.5
        assert linha['Preco_Minimo'] == 10.5
        assert linha['Preco_Maximo'] == 10.5
        assert linha['Qtd_Compras'] == 1
        assert linha['Fornecedores_Distintos'] == 1
        assert linha['Fornecedor_Ultima_Compra'] == 'Fornecedor A'

    def test_ultimo_preco_e_o_da_compra_mais_recente_nao_o_ultimo_da_planilha(self):
        # A ordem das linhas na planilha não segue necessariamente a ordem
        # cronológica — "último preço" deve ser sempre o de maior data.
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '12,00', '20/01/2026', 'Fornecedor B', '50', 'Aprovado'],
            ['1268', 'Parafuso', 'UN', '10,00', '05/03/2026', 'Fornecedor A', '100', 'Aprovado'],
            ['1268', 'Parafuso', 'UN', '11,00', '10/02/2026', 'Fornecedor C', '30', 'Aprovado'],
        ])
        base_precos, df_f = consolidar_base_precos(df)

        linha = base_precos.iloc[0]
        assert linha['Ultimo_Preco'] == 10.00
        assert linha['Fornecedor_Ultima_Compra'] == 'Fornecedor A'
        assert linha['Data_Ultima_Compra'] == pd.Timestamp('2026-03-05')
        assert linha['Preco_Medio'] == pytest.approx((12 + 10 + 11) / 3)
        assert linha['Preco_Minimo'] == 10.00
        assert linha['Preco_Maximo'] == 12.00
        assert linha['Qtd_Compras'] == 3
        assert linha['Fornecedores_Distintos'] == 3
        assert linha['Qtd_Total_Comprada'] == 180.0

    def test_data_dd_mm_aaaa_realista_e_interpretada_com_dayfirst(self):
        # Entrada real que circula pela planilha "Pedidos" é sempre string
        # "DD/MM/AAAA" (ver comentário de construir_base_precos em app.py).
        # "05/03/2026" tem que virar 5 de MARÇO, não 3 de maio.
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '10,00', '05/03/2026', 'Fornecedor A', '10', 'Aprovado'],
        ])
        _, df_f = consolidar_base_precos(df)
        data = df_f.iloc[0]['Data Emissao']
        assert data == pd.Timestamp('2026-03-05')
        assert data.day == 5
        assert data.month == 3

    def test_codigos_com_e_sem_zero_a_esquerda_sao_o_mesmo_item(self):
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '10,00', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
            ['0000001268', 'Parafuso', 'UN', '11,00', '02/01/2026', 'Fornecedor B', '10', 'Aprovado'],
        ])
        base_precos, _ = consolidar_base_precos(df)
        assert len(base_precos) == 1
        assert base_precos.iloc[0]['Qtd_Compras'] == 2

    def test_precos_zerados_ou_negativos_sao_descartados(self):
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '0', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
            ['1268', 'Parafuso', 'UN', '-5,00', '02/01/2026', 'Fornecedor A', '10', 'Aprovado'],
        ])
        base_precos, df_f = consolidar_base_precos(df)
        assert base_precos.empty
        assert df_f.empty

    def test_produto_vazio_e_descartado(self):
        df = _historico_bruto([
            ['', 'Sem produto', 'UN', '10,00', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
            ['1268', 'Parafuso', 'UN', '10,00', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
        ])
        base_precos, _ = consolidar_base_precos(df)
        assert len(base_precos) == 1
        assert base_precos.iloc[0]['Cod_Norm'] == '0000001268'

    def test_status_filtro_considera_so_o_status_pedido(self):
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '10,00', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
            ['1268', 'Parafuso', 'UN', '99,00', '02/01/2026', 'Fornecedor B', '10', 'Rejeitado'],
        ])
        base_precos_tudo, _ = consolidar_base_precos(df)
        assert base_precos_tudo.iloc[0]['Qtd_Compras'] == 2  # "HISTÓRICO É HISTÓRICO"

        base_precos_aprovado, _ = consolidar_base_precos(df, status_filtro='Aprovado')
        assert base_precos_aprovado.iloc[0]['Qtd_Compras'] == 1
        assert base_precos_aprovado.iloc[0]['Ultimo_Preco'] == 10.00

    def test_preco_numerico_ja_float_nao_precisa_de_limpar_valor(self):
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', 10.5, '01/01/2026', 'Fornecedor A', 10, 'Aprovado'],
        ])
        base_precos, _ = consolidar_base_precos(df)
        assert base_precos.iloc[0]['Ultimo_Preco'] == 10.5

    def test_multiplos_itens_distintos_geram_uma_linha_cada(self):
        df = _historico_bruto([
            ['1268', 'Parafuso', 'UN', '10,00', '01/01/2026', 'Fornecedor A', '10', 'Aprovado'],
            ['9999', 'Porca', 'UN', '2,50', '01/01/2026', 'Fornecedor A', '100', 'Aprovado'],
        ])
        base_precos, _ = consolidar_base_precos(df)
        assert len(base_precos) == 2
        assert set(base_precos['Cod_Norm']) == {'0000001268', '0000009999'}
