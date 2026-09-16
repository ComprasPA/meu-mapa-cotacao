"""Testes de processar_mapa_cotacao (parte pura do bloco de PROCESSAMENTO DO
MAPA DE COTAÇÃO em app.py): dado o mapa de cotação já extraído do arquivo de
upload + a base histórica consolidada, monta o comparativo por item —
melhor cotação (menor preço), variação vs. último preço/preço médio, e a
"Observação" (abaixo/acima/próximo da média, ou sem histórico)."""
import pandas as pd
import pytest

from logica_cotacao import processar_mapa_cotacao, extrair_numero_cotacao


def _base_precos(linhas):
    colunas = ['Cod_Norm', 'Ultimo_Preco', 'Data_Ultima_Compra',
               'Fornecedor_Ultima_Compra', 'Preco_Medio', 'Preco_Minimo', 'Preco_Maximo']
    return pd.DataFrame(linhas, columns=colunas)


class TestCotacaoVaziaOuSemItensValidos:
    def test_cotacao_vazia_retorna_df_vazio(self):
        df_final, aviso = processar_mapa_cotacao(pd.DataFrame(), _base_precos([]))
        assert df_final.empty
        assert aviso is False

    def test_cotacao_sem_coluna_de_codigo_reconhecivel_retorna_vazio(self):
        cotacao = pd.DataFrame({'Nada a ver': ['x', 'y']})
        df_final, aviso = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.empty

    def test_linha_so_de_cabecalho_repetido_e_ignorada(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'],
            'Descrição': ['Descrição'],  # repete o nome do cabeçalho -> deve ser ignorada
            'Valor Unitário': ['10,00'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.empty


class TestUmUnicoFornecedor:
    def test_item_unico_sem_historico(self):
        cotacao = pd.DataFrame({
            'Item': ['1'],
            'Código': ['1268'],
            'Descrição': ['Parafuso Sextavado'],
            'Unidade': ['UN'],
            'Qtd': ['100'],
            'Valor Unitário': ['12,50'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, aviso = processar_mapa_cotacao(cotacao, _base_precos([]))

        assert len(df_final) == 1
        linha = df_final.iloc[0]
        assert linha['Código'] == '0000001268'
        assert linha['Valor Cotado (R$)'] == 12.5
        assert linha['Fornecedor Cotado'] == 'Fornecedor A'
        assert linha['Observação'] == 'Sem histórico de compra'
        assert linha['Último Preço Pago (R$)'] == ""
        assert linha['Var. vs Último (%)'] == ""
        assert aviso is False

    def test_item_unico_com_historico_preco_igual_variacao_zero(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['10,00'], 'Fornecedor': ['Fornecedor A'],
        })
        base = _base_precos([
            ['0000001268', 10.0, pd.Timestamp('2026-01-01'), 'Fornecedor Antigo', 10.0, 10.0, 10.0],
        ])
        df_final, _ = processar_mapa_cotacao(cotacao, base)

        linha = df_final.iloc[0]
        assert linha['Var. vs Último (%)'] == 0.0
        assert linha['Var. vs Médio (%)'] == 0.0
        assert linha['Observação'] == 'Próximo da média histórica'


class TestMultiplosFornecedoresMesmoItem:
    def test_melhor_cotacao_e_o_menor_valor_entre_fornecedores(self):
        # Mapa "bruto": mesmo código em 3 linhas (uma por fornecedor cotando
        # o mesmo item) - só a primeira linha do grupo tem o código
        # preenchido (ffill cobre isso, ver Passo 2 da skill).
        cotacao = pd.DataFrame({
            'Código': ['1268', None, None],
            'Descrição': ['Parafuso', 'Parafuso', 'Parafuso'],
            'Qtd': ['10', '10', '10'],
            'Valor Unitário': ['15,00', '9,90', '11,00'],
            'Fornecedor': ['Fornecedor Caro', 'Fornecedor Barato', 'Fornecedor Médio'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))

        assert len(df_final) == 1
        linha = df_final.iloc[0]
        assert linha['Valor Cotado (R$)'] == 9.90
        assert linha['Fornecedor Cotado'] == 'Fornecedor Barato'

    def test_empate_de_preco_entre_fornecedores_fica_com_um_deles(self):
        cotacao = pd.DataFrame({
            'Código': ['1268', None],
            'Descrição': ['Parafuso', 'Parafuso'],
            'Qtd': ['10', '10'],
            'Valor Unitário': ['10,00', '10,00'],
            'Fornecedor': ['Fornecedor A', 'Fornecedor B'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))

        assert len(df_final) == 1
        assert df_final.iloc[0]['Valor Cotado (R$)'] == 10.0
        assert df_final.iloc[0]['Fornecedor Cotado'] in ('Fornecedor A', 'Fornecedor B')

    def test_status_vencedor_filtra_so_a_linha_vencedora(self):
        cotacao = pd.DataFrame({
            'Código': ['1268', None],
            'Descrição': ['Parafuso', 'Parafuso'],
            'Qtd': ['10', '10'],
            'Valor Unitário': ['15,00', '9,90'],
            'Fornecedor': ['Fornecedor Caro', 'Fornecedor Barato'],
            'Status': ['', 'Vencedor'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))

        assert len(df_final) == 1
        assert df_final.iloc[0]['Fornecedor Cotado'] == 'Fornecedor Barato'
        assert df_final.iloc[0]['Valor Cotado (R$)'] == 9.90


class TestVariacaoEObservacao:
    def test_preco_muito_acima_da_media_marca_acima_da_media(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['20,00'], 'Fornecedor': ['Fornecedor A'],
        })
        base = _base_precos([
            ['0000001268', 10.0, pd.Timestamp('2026-01-01'), 'Fornecedor X', 10.0, 8.0, 12.0],
        ])
        df_final, aviso = processar_mapa_cotacao(cotacao, base)
        linha = df_final.iloc[0]
        assert linha['Var. vs Médio (%)'] == 100.0
        assert linha['Var. vs Último (%)'] == 100.0
        assert linha['Observação'] == 'Acima da média histórica'

    def test_preco_muito_abaixo_da_media_marca_abaixo_da_media(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['5,00'], 'Fornecedor': ['Fornecedor A'],
        })
        base = _base_precos([
            ['0000001268', 10.0, pd.Timestamp('2026-01-01'), 'Fornecedor X', 10.0, 5.0, 12.0],
        ])
        df_final, _ = processar_mapa_cotacao(cotacao, base)
        linha = df_final.iloc[0]
        assert linha['Var. vs Médio (%)'] == -50.0
        assert linha['Observação'] == 'Abaixo da média histórica'

    def test_aviso_de_escala_dispara_quando_maioria_esta_fora_da_curva(self):
        # Cenário lote vs. unitário: mapa cotado em "R$ por lote" contra um
        # histórico de preço unitário -> razão >10x sistemática em >30% dos
        # itens deve disparar o aviso.
        cotacao = pd.DataFrame({
            'Código': ['1', '2', '3', '4'],
            'Descrição': ['Item 1', 'Item 2', 'Item 3', 'Item 4'],
            'Qtd': ['1', '1', '1', '1'],
            'Valor Unitário': ['1000,00', '2000,00', '1500,00', '10,00'],
            'Fornecedor': ['F1', 'F2', 'F3', 'F4'],
        })
        base = _base_precos([
            ['0000000001', 10.0, pd.Timestamp('2026-01-01'), 'X', 10.0, 8.0, 12.0],
            ['0000000002', 20.0, pd.Timestamp('2026-01-01'), 'X', 20.0, 18.0, 22.0],
            ['0000000003', 15.0, pd.Timestamp('2026-01-01'), 'X', 15.0, 13.0, 17.0],
            ['0000000004', 10.0, pd.Timestamp('2026-01-01'), 'X', 10.0, 8.0, 12.0],
        ])
        df_final, aviso = processar_mapa_cotacao(cotacao, base)
        assert aviso is True

    def test_sem_aviso_quando_precos_estao_na_mesma_escala(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['10,50'], 'Fornecedor': ['Fornecedor A'],
        })
        base = _base_precos([
            ['0000001268', 10.0, pd.Timestamp('2026-01-01'), 'Fornecedor X', 10.0, 8.0, 12.0],
        ])
        _, aviso = processar_mapa_cotacao(cotacao, base)
        assert aviso is False


class TestValorUnitarioComDesconto:
    """Quando a planilha traz TOTAL e Desconto separados do Vl.Unitário
    bruto (formato real do TOTVS, confirmado com arquivo do usuário
    021132-1.xlsx), o valor cotado tem que ser o unitário JÁ com o
    desconto aplicado: (TOTAL - Desconto) / Qtde — nunca o Vl.Unitário
    bruto. Fórmula e nomes de coluna conferidos com a planilha real do
    usuário (que já fazia essa conta manualmente numa coluna auxiliar)."""

    def test_usa_total_menos_desconto_dividido_pela_qtde(self):
        cotacao = pd.DataFrame({
            'Código': ['2892'], 'Descrição': ['Cinta Carga'], 'Qtde': ['2'],
            'Vl.Unitario': ['68,000'], 'TOTAL': ['123,05'], 'Desconto': ['6,80'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        # (123,05 - 6,80) / 2 = 58,125
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(58.125)

    def test_desconto_zero_valor_liquido_igual_ao_bruto(self):
        cotacao = pd.DataFrame({
            'Código': ['2892'], 'Descrição': ['Cinta Carga'], 'Qtde': ['2'],
            'Vl.Unitario': ['31,000'], 'TOTAL': ['62,00'], 'Desconto': ['0'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(31.0)

    def test_ignora_vl_unitario_bruto_quando_total_e_desconto_existem(self):
        # Vl.Unitario bruto (100) é bem maior que o líquido calculado (45) —
        # se a função usasse o bruto por engano, o teste pegaria.
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Item'], 'Qtde': ['1'],
            'Vl.Unitario': ['100,00'], 'TOTAL': ['100,00'], 'Desconto': ['55,00'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(45.0)

    def test_sem_coluna_total_ou_desconto_usa_vl_unitario_bruto_normalmente(self):
        # Formato antigo (sem TOTAL/Desconto) continua funcionando como
        # antes — não pode quebrar os arquivos que não têm essas colunas.
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['12,50'], 'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(12.5)

    def test_total_vazio_cai_pro_vl_unitario_bruto(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Item'], 'Qtde': ['1'],
            'Vl.Unitario': ['20,00'], 'TOTAL': [''], 'Desconto': ['5,00'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(20.0)

    def test_melhor_cotacao_considera_o_valor_liquido_nao_o_bruto(self):
        # Fornecedor Caro tem Vl.Unitario bruto MENOR mas desconto pequeno
        # (líquido 90); Fornecedor Barato tem bruto maior mas desconto
        # grande (líquido 45) — o vencedor tem que ser o de líquido menor.
        cotacao = pd.DataFrame({
            'Código': ['1268', None],
            'Descrição': ['Item', 'Item'],
            'Qtde': ['1', '1'],
            'Vl.Unitario': ['95,00', '100,00'],
            'TOTAL': ['95,00', '100,00'],
            'Desconto': ['5,00', '55,00'],
            'Fornecedor': ['Fornecedor Caro Aparente', 'Fornecedor Barato de Verdade'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert len(df_final) == 1
        assert df_final.iloc[0]['Fornecedor Cotado'] == 'Fornecedor Barato de Verdade'
        assert df_final.iloc[0]['Valor Cotado (R$)'] == pytest.approx(45.0)


class TestColunaValorAlternativa:
    def test_sem_coluna_de_valor_reconhecivel_tenta_outra_coluna_numerica(self):
        # Export "cru" às vezes não tem uma coluna óbvia de valor: a função
        # cai para a primeira coluna com número > 0 e diferente da
        # quantidade. NOTA: esse fallback varre TODAS as colunas da linha,
        # inclusive a própria coluna de código — por isso o código aqui
        # precisa ter um prefixo não numérico ("SKU-1268"), senão o valor
        # "1268" (parseado pelo limpar_valor da própria coluna Código) seria
        # escolhido antes da "Coluna Misteriosa" (ver observação no
        # relatório: comportamento real, não corrigido nesta tarefa).
        cotacao = pd.DataFrame({
            'Código': ['SKU-1268'],
            'Descrição': ['Parafuso'],
            'Qtd': ['10'],
            'Coluna Misteriosa': ['25,00'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert len(df_final) == 1
        assert df_final.iloc[0]['Valor Cotado (R$)'] == 25.0

    def test_sem_nenhum_valor_positivo_item_e_descartado(self):
        cotacao = pd.DataFrame({
            'Código': ['SKU-1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.empty

    def test_codigo_puramente_numerico_sem_coluna_de_valor_e_confundido_com_preco(self):
        # Comportamento real documentado (NÃO é o comportamento desejado, e
        # NÃO foi corrigido nesta tarefa — ver relatório): quando o código
        # do produto é só dígitos (o formato real do TOTVS, ex. "1268") e o
        # mapa não tem nenhuma coluna reconhecível de valor unitário, o
        # fallback de "primeira coluna numérica positiva" pega o PRÓPRIO
        # CÓDIGO como se fosse o preço, porque ele aparece antes de
        # qualquer outra coluna numérica na linha.
        cotacao = pd.DataFrame({
            'Código': ['1268'],
            'Descrição': ['Parafuso'],
            'Qtd': ['10'],
            'Coluna Misteriosa': ['25,00'],
            'Fornecedor': ['Fornecedor A'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert len(df_final) == 1
        assert df_final.iloc[0]['Valor Cotado (R$)'] == 1268.0


class TestPrioridadeRazaoSocial:
    def test_razao_social_tem_prioridade_sobre_fornecedor(self):
        cotacao = pd.DataFrame({
            'Código': ['1268'], 'Descrição': ['Parafuso'], 'Qtd': ['10'],
            'Valor Unitário': ['10,00'],
            'Fornecedor': ['000079'],
            'Razão Social': ['LJ GUERRA E CIA LTDA'],
        })
        df_final, _ = processar_mapa_cotacao(cotacao, _base_precos([]))
        assert df_final.iloc[0]['Fornecedor Cotado'] == 'LJ GUERRA E CIA LTDA'


class TestExtrairNumeroCotacao:
    """extrair_numero_cotacao acha o número da cotação (ex.: "021132") na
    coluna "COTACAO :" do export do TOTVS, repetida em toda linha — usado
    pra nomear os arquivos exportados (PDF/Excel) igual ao arquivo original."""

    def test_acha_o_numero_na_coluna_cotacao(self):
        cotacao = pd.DataFrame({
            'COTACAO :': ['021132', '021132'],
            'Produto': ['1268', '1930'],
        })
        assert extrair_numero_cotacao(cotacao) == '021132'

    def test_sem_coluna_de_cotacao_devolve_none(self):
        cotacao = pd.DataFrame({'Produto': ['1268'], 'Descrição': ['Parafuso']})
        assert extrair_numero_cotacao(cotacao) is None

    def test_cotacao_vazia_devolve_none(self):
        assert extrair_numero_cotacao(pd.DataFrame()) is None

    def test_coluna_de_cotacao_toda_vazia_devolve_none(self):
        cotacao = pd.DataFrame({'COTACAO :': ['', None], 'Produto': ['1268', '1930']})
        assert extrair_numero_cotacao(cotacao) is None

    def test_ignora_linhas_vazias_e_usa_o_primeiro_valor_valido(self):
        cotacao = pd.DataFrame({'COTACAO :': ['', '021132', '021132'], 'Produto': ['x', 'y', 'z']})
        assert extrair_numero_cotacao(cotacao) == '021132'
