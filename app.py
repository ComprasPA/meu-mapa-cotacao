import streamlit as pd_st
import pandas as pd
import numpy as np

# Configuração da Página
pd_st.set_page_config(
    page_title="Mapa de Cotação e Análise de Custos",
    page_icon="📊",
    layout="wide"
)

# Estilização CSS profissional
pd_st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    h1 { color: #1f2c34; font-family: 'Helvetica Neue', sans-serif; }
    .stAlert { font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

pd_st.title("📊 Gestão Estratégica de Suprimentos | Mapa de Cotação & Histórico")
pd_st.markdown("Plataforma analítica para homologação de preços, variação de custos e tomada de decisão comercial.")

# Sidebar para upload interativo de dados
pd_st.sidebar.header("📁 Painel de Dados")
uploaded_file = pd_st.sidebar.file_uploader("Carregar Mapa de Cotação (.csv ou .xlsx)", type=["csv", "xlsx"])
uploaded_history = pd_st.sidebar.file_uploader("Carregar Histórico de Compras (.csv ou .xlsx)", type=["csv", "xlsx"])

# Bases de Dados Padrão (Garantia de funcionamento imediato)
def carregar_dados_padrao():
    hist_data = pd.DataFrame({
        'Código': ['SKU001', 'SKU001', 'SKU002', 'SKU003', 'SKU003'],
        'Data': ['2025-10-15', '2026-02-10', '2026-01-20', '2025-11-05', '2026-03-12'],
        'Item': ['Luva de Raspa', 'Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular', 'Terminal Tubular'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [100, 150, 20, 50, 60],
        'Preço Unit. (R$)': [12.50, 13.20, 85.00, 22.00, 20.50],
        'Fornecedor': ['EPI Manaus Comércio', 'Industrial Sul Ltda', 'Papelaria & Cia', 'Conexões Amazônia', 'Global Elétrica SP']
    })
    
    cot_data = pd.DataFrame({
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [120, 25, 40],
        'Novo Preço Unit. (R$)': [14.00, 82.00, 23.50],
        'Fornecedor do Preço Novo': ['EPI Manaus Comércio', 'Papelaria & Cia', 'Metaltec Suprimentos']
    })
    return hist_data, cot_data

# Leitura dos arquivos enviados pelo usuário ou uso do padrão
try:
    if uploaded_history is not None:
        hist_data = pd.read_csv(uploaded_history) if uploaded_history.name.endswith('.csv') else pd.read_excel(uploaded_history)
    else:
        # Tenta ler do GitHub se existir, senão usa o padrão
        try:
            hist_data = pd.read_csv("historico_compras.csv")
        except:
            hist_data, _ = carregar_dados_padrao()

    if uploaded_file is not None:
        cot_data = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
    else:
        _, cot_data = carregar_dados_padrao()
except Exception as e:
    pd_st.error(f"Erro ao carregar os arquivos: {e}")
    hist_data, cot_data = carregar_dados_padrao()

# Normalização de nomes de colunas para aceitar variações de maiúsculas/minúsculas
def limpar_colunas(df):
    df.columns = df.columns.str.strip().str.title()
    return df

hist_data = limpar_colunas(hist_data)
cot_data = limpar_colunas(cot_data)

# Processamento do Comparativo Histórico
def processar_comparativo(historico, cotacao):
    resultado = []
    
    # Padronização interna de nomes comuns
    for col in historico.columns:
        if 'Cod' in col or 'Sku' in col: historico.rename(columns={col: 'Código'}, inplace=True)
        if 'Preco' in col or 'Preço' in col: historico.rename(columns={col: 'Preço Unit. (R$)'}, inplace=True)
        if 'Fornecedor' in col: historico.rename(columns={col: 'Fornecedor'}, inplace=True)

    for col in cotacao.columns:
        if 'Cod' in col or 'Sku' in col: cotacao.rename(columns={col: 'Código'}, inplace=True)
        if 'Novo Preco' in col or 'Novo Preço' in col: cotacao.rename(columns={col: 'Novo Preço Unit. (R$)'}, inplace=True)
        if 'Qtd' in col or 'Quantidade' in col: cotacao.rename(columns={col: 'Qtd'}, inplace=True)

    for _, row in cotacao.iterrows():
        codigo = str(row.get('Código', 'N/D'))
        item = str(row.get('Item', 'Item Genérico'))
        desc = str(row.get('Descrição Resumida', row.get('Descricao Resumida', '')))
        qtd_novo = float(row.get('Qtd', 0))
        preco_novo = float(row.get('Novo Preço Unit. (R$)', row.get('Preço Unit. (R$)', 0.0)))
        forn_novo = str(row.get('Fornecedor Do Preço Novo', row.get('Fornecedor', 'Não informado')))
        
        # Filtra histórico do item correspondente
        if 'Código' in historico.columns:
            hist_item = historico[historico['Código'].astype(str) == codigo]
        else:
            hist_item = pd.DataFrame()
        
        if not hist_item.empty and 'Preço Unit. (R$)' in hist_item.columns:
            ultimo_preco_hist = float(hist_item.iloc[0]['Preço Unit. (R$)'])
            forn_hist = str(hist_item.iloc[0].get('Fornecedor', 'Histórico Anterior'))
        else:
            ultimo_preco_hist = preco_novo
            forn_hist = "Sem Histórico"
            
        # Cálculo de Variação (Δ%)
        if ultimo_preco_hist > 0:
            variacao = ((preco_novo - ultimo_preco_hist) / ultimo_preco_hist) * 100
        else:
            variacao = 0.0
            
        # Determinar Tendência
        if variacao > 1.0:
            tendencia = "📈 Alta"
        elif variacao < -1.0:
            tendencia = "📉 Queda"
        else:
            tendencia = "➡️ Estabilidade"
            
        resultado.append({
            'Item': item,
            'Código': codigo,
            'Descrição Resumida': desc,
            'Qtd': qtd_novo,
            'Último Preço Hist. (R$)': round(ultimo_preco_hist, 2),
            'Fornecedor do Último Preço': forn_hist,
            'Novo Preço Unit. (R$)': round(preco_novo, 2),
            'Fornecedor do Preço Novo': forn_novo,
            'Variação (Δ%)': round(variacao, 2),
            'Tendência': tendencia
        })
        
    df_res = pd.DataFrame(resultado)
    
    # Ordem rigorosa exigida pelas diretrizes do especialista
    colunas_ordenadas = [
        'Item', 'Código', 'Descrição Resumida', 'Qtd', 
        'Último Preço Hist. (R$)', 'Fornecedor do Último Preço', 
        'Novo Preço Unit. (R$)', 'Fornecedor do Preço Novo', 
        'Variação (Δ%)', 'Tendência'
    ]
    
    return df_res[colunas_ordenadas]

df_resultado = processar_comparativo(hist_data, cot_data)

# Exibição da Tabela Consolidada
pd_st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

pd_st.dataframe(df_resultado.style.format({
    'Qtd': '{:,.0f}',
    'Último Preço Hist. (R$)': 'R$ {:.2f}',
    'Novo Preço Unit. (R$)': 'R$ {:.2f}',
    'Variação (Δ%)': '{:+.2f}%'
}), use_container_width=True)

# Bloco de Observações Técnicas (Logística e ZFM)
pd_st.markdown("---")
pd_st.subheader("🔍 Observações Logísticas e Fiscais (ZFM)")
pd_st.markdown("""
* **Impacto Logístico:** Avaliação do custo total de frete (modal aéreo/fluvial) para suprimentos oriundos de 'Fora do Estado' em comparação com fornecedores locais de Manaus.
* **Incentivos Fiscais:** Validação da aplicação correta dos benefícios tributários da Zona Franca de Manaus (ZFM) para preservação de margem.
* **Picos Fora da Curva:** Análise crítica obrigatória em variações superiores a +5%, verificando oscilações de matéria-prima e custos logísticos antes da emissão da O.C.
""")

# Seção Obrigatória: Insight do Especialista
pd_st.markdown("---")
pd_st.subheader("💡 Insight do Especialista")

itens_em_alta = df_resultado[df_resultado['Variação (Δ%)'] > 0]
if not itens_em_alta.empty:
    insight_texto = (
        f"Detectada pressão inflacionária em **{len(itens_em_alta)} item(ns)** da cotação atual. "
        "Recomenda-se a **renegociação imediata baseada no volume histórico de consumo** ou "
        "o acionamento de praças alternativas na região de Manaus para otimização do custo total (preço + frete)."
    )
else:
    insight_texto = (
        "Os preços encontram-se estáveis ou em tendência de deflação. "
        "Recomenda-se a **antecipação de compras estratégicas** para garantir o abastecimento "
        "e fixar condições comerciais favoráveis perante a sazonalidade."
    )

pd_st.success(insight_texto)
