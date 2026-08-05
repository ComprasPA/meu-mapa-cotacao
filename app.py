import streamlit as pd_st # Usando alias para evitar conflito com pandas
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

# Sidebar para upload de dados
pd_st.sidebar.header("📁 Fontes de Dados")
uploaded_file = pd_st.sidebar.file_uploader("Carregar Mapa de Cotação Atual (.csv ou .xlsx)", type=["csv", "xlsx"])
uploaded_history = pd_st.sidebar.file_uploader("Carregar Histórico de Compras (.csv ou .xlsx)", type=["csv", "xlsx"])

# Dados Mockados para demonstração imediata caso o usuário não envie arquivos
def carregar_dados_exemplo():
    # Histórico simulado
    hist_data = pd.DataFrame({
        'Código': ['SKU001', 'SKU001', 'SKU002', 'SKU003', 'SKU003'],
        'Data': ['2025-10-15', '2026-02-10', '2026-01-20', '2025-11-05', '2026-03-12'],
        'Item': ['Luva de Raspa', 'Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular', 'Terminal Tubular'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [100, 150, 20, 50, 60],
        'Preço Unit. (R$)': [12.50, 13.20, 85.00, 22.00, 20.50],
        'Fornecedor': ['EPI Manaus Comércio', 'Industrial Sul Ltda', 'Papelaria & Cia', 'Conexões Amazônia', 'Global Elétrica SP'],
        'Origem': ['Local (Manaus)', 'Fora do Estado', 'Local (Manaus)', 'Local (Manaus)', 'Fora do Estado']
    })
    
    # Cotação nova simulada
    cot_data = pd.DataFrame({
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [120, 25, 40],
        'Novo Preço Unit. (R$)': [14.00, 82.00, 23.50],
        'Fornecedor do Preço Novo': ['EPI Manaus Comércio', 'Papelaria & Cia', 'Metaltec Suprimentos'],
        'Origem Novo': ['Local (Manaus)', 'Local (Manaus)', 'Fora do Estado']
    })
    return hist_data, cot_data

# Carregamento efetivo
if uploaded_file is not None and uploaded_history is not None:
    if uploaded_file.name.endswith('.csv'):
        cot_data = pd.read_csv(uploaded_file)
    else:
        cot_data = pd.read_excel(uploaded_file)
        
    if uploaded_history.name.endswith('.csv'):
        hist_data = pd.read_csv(uploaded_history)
    else:
        hist_data = pd.read_excel(uploaded_history)
else:
    pd_st.sidebar.info("Exibindo dados de demonstração. Faça o upload dos seus arquivos para análise real.")
    hist_data, cot_data = carregar_dados_exemplo()

# Processamento do Comparativo Histórico
def processar_comparativo(historico, cotacao):
    resultado = []
    
    # Ordenar histórico por data decrescente para pegar sempre a compra anterior mais recente
    historico['Data'] = pd.to_datetime(historico['Data'])
    historico = historico.sort_values(by='Data', ascending=False)
    
    for _, row in cotacao.iterrows():
        codigo = row['Código']
        item = row['Item']
        desc = row['Descrição Resumida']
        qtd_novo = row['Qtd']
        preco_novo = row['Novo Preço Unit. (R$)']
        forn_novo = row['Fornecedor do Preço Novo']
        
        # Filtrar histórico do item
        hist_item = historico[historico['Código'] == codigo]
        
        if not hist_item.empty:
            ultimo_preco_hist = hist_item.iloc[0]['Preço Unit. (R$)']
            forn_hist = hist_item.iloc[0]['Fornecedor']
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
        
    return pd.DataFrame(resultado)

df_resultado = processar_comparativo(hist_data, cot_data)

# Exibição da Tabela Obrigatória
pd_st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

# Formatando visualmente a tabela no Streamlit
pd_st.dataframe(df_resultado.style.format({
    'Último Preço Hist. (R$)': 'R$ {:.2f}',
    'Novo Preço Unit. (R$)': 'R$ {:.2f}',
    'Variação (Δ%)': '{:+.2f}%'
}), use_container_width=True)

# Bloco de Observações Técnicas (Bullet Points)
pd_st.markdown("---")
pd_st.subheader("🔍 Observações Logísticas e Fiscais (ZFM)")
pd_st.markdown("""
* **Impacto Logístico:** Itens adquiridos de fornecedores 'Fora do Estado' devem ser reavaliados frente aos custos de frete aéreo/fluvial para Manaus e eventuais impactos no *lead time*.
* **Incentivos Fiscais:** Garantir a aplicação correta dos benefícios da Zona Franca de Manaus (ZFM) em aquisições de insumos produtivos para assegurar competitividade de margem.
* **Picos Fora da Curva:** Variações superiores a +5% exigem auditoria nas composições de custos dos insumos primários (matéria-prima e embalagem) antes da emissão da Ordem de Compra.
""")

# Seção Obrigatória: Insight do Especialista
pd_st.markdown("---")
pd_st.subheader("💡 Insight do Especialista")

# Análise automática rápida para o insight
itens_em_alta = df_resultado[df_resultado['Variação (Δ%)'] > 0]
if not itens_em_alta.empty:
    insight_texto = (
        f"Identificada pressão inflacionária em **{len(itens_em_alta)} item(ns)** do escopo atual. "
        "Recomenda-se **renegociação imediata com base no histórico de volume** ou consulta a "
        "fornecedores homologados alternativos na praça local (Manaus) para mitigar o impacto margem, "
        "priorizando fornecedores com entrega imediata para evitar ruptura de estoque."
    )
else:
    insight_texto = (
        "O cenário de preços apresenta estabilidade ou tendência de queda. "
        "Recomenda-se a **antecipação de compras estratégicas** para garantir o abastecimento "
        "aproveitando as condições atuais favoráveis de mercado."
    )

pd_st.success(insight_texto)
