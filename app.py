import streamlit as st
import pandas as pd
import numpy as np

# Configuração da Página
st.set_page_config(
    page_title="Mapa de Cotação e Análise de Custos",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Gestão Estratégica de Suprimentos | Mapa de Cotação & Histórico")
st.markdown("Plataforma analítica para homologação de preços, variação de custos e tomada de decisão comercial.")

# Barra lateral para upload opcional
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader("Mapa de Cotação Atual (.csv ou .xlsx)", type=["csv", "xlsx"])
uploaded_hist = st.sidebar.file_uploader("Histórico de Compras (.csv ou .xlsx)", type=["csv", "xlsx"])

# Função segura para carregar arquivos ou usar dados padrão de demonstração
@st.cache_data
def carregar_dados_padrao():
    df_hist = pd.DataFrame({
        'Código': ['SKU001', 'SKU001', 'SKU002', 'SKU003', 'SKU003'],
        'Data': ['2025-10-15', '2026-02-10', '2026-01-20', '2025-11-05', '2026-03-12'],
        'Item': ['Luva de Raspa', 'Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular', 'Terminal Tubular'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [100, 150, 20, 50, 60],
        'Preço Unit. (R$)': [12.50, 13.20, 85.00, 22.00, 20.50],
        'Fornecedor': ['EPI Manaus Comércio', 'Industrial Sul Ltda', 'Papelaria & Cia', 'Conexões Amazônia', 'Global Elétrica SP']
    })
    
    df_cot = pd.DataFrame({
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [120, 25, 40],
        'Novo Preço Unit. (R$)': [14.00, 82.00, 23.50],
        'Fornecedor do Preço Novo': ['EPI Manaus Comércio', 'Papelaria & Cia', 'Metaltec Suprimentos']
    })
    return df_hist, df_cot

# Leitura com proteção contra erros
hist_padrao, cot_padrao = carregar_dados_padrao()

try:
    if uploaded_hist is not None:
        historico = pd.read_csv(uploaded_hist) if uploaded_hist.name.endswith('.csv') else pd.read_excel(uploaded_hist)
    else:
        # Tenta ler do workspace local/GitHub, se falhar usa o padrão
        try:
            historico = pd.read_csv("historico_compras.csv")
        except:
            historico = hist_padrao
except Exception as e:
    st.sidebar.warning(f"Aviso no histórico: {e}. Usando dados padrão.")
    historico = hist_padrao

try:
    if uploaded_cot is not None:
        cotacao = pd.read_csv(uploaded_cot) if uploaded_cot.name.endswith('.csv') else pd.read_excel(uploaded_cot)
    else:
        cotacao = cot_padrao
except Exception as e:
    st.sidebar.warning(f"Aviso na cotação: {e}. Usando dados padrão.")
    cotacao = cot_padrao

# Limpeza e normalização básica de colunas
historico.columns = [str(c).strip() for c in historico.columns]
cotacao.columns = [str(c).strip() for c in cotacao.columns]

# Motor de processamento e comparativo histórico
resultados = []

for _, row_cot in cotacao.iterrows():
    # Identificação flexível de colunas
    codigo = str(row_cot.get('Código', row_cot.get('Codigo', row_cot.get('SKU', 'N/D'))))
    item = str(row_cot.get('Item', row_cot.get('Produto', 'Genérico')))
    desc = str(row_cot.get('Descrição Resumida', row_cot.get('Descricao Resumida', '')))
    qtd = float(row_cot.get('Qtd', row_cot.get('Quantidade', 0)))
    
    # Preço novo
    preco_novo = float(row_cot.get('Novo Preço Unit. (R$)', row_cot.get('Novo Preco Unit. (R$)', row_cot.get('Preço Unit. (R$)', 0.0))))
    forn_novo = str(row_cot.get('Fornecedor do Preço Novo', row_cot.get('Fornecedor', 'Não informado')))
    
    # Busca histórico do SKU/Código correspondente
    match_hist = historico[historico.astype(str).apply(lambda x: x.str.contains(codigo, case=False)).any(axis=1)] if not historico.empty else pd.DataFrame()
    
    if not match_hist.empty:
        # Pega o último registro disponível no histórico
        col_preco_hist = [c for c in match_hist.columns if 'preço' in c.lower() or 'preco' in c.lower() or 'unit' in c.lower()]
        col_forn_hist = [c for c in match_hist.columns if 'fornecedor' in c.lower()]
        
        if col_preco_hist:
            ultimo_preco = float(match_hist.iloc[-1][col_preco_hist[0]])
        else:
            ultimo_preco = preco_novo
            
        if col_forn_hist:
            forn_hist = str(match_hist.iloc[-1][col_forn_hist[0]])
        else:
            forn_hist = "Histórico Anterior"
    else:
        ultimo_preco = preco_novo
        forn_hist = "Sem Histórico"
        
    # Cálculos de Variação e Tendência
    if ultimo_preco > 0:
        variacao = ((preco_novo - ultimo_preco) / ultimo_preco) * 100
    else:
        variacao = 0.0
        
    if variacao > 0.5:
        tendencia = "📈 Alta"
    elif variacao < -0.5:
        tendencia = "📉 Queda"
    else:
        tendencia = "➡️ Estabilidade"
        
    resultados.append({
        'Item': item,
        'Código': codigo,
        'Descrição Resumida': desc,
        'Qtd': qtd,
        'Último Preço Hist. (R$)': round(ultimo_preco, 2),
        'Fornecedor do Último Preço': forn_hist,
        'Novo Preço Unit. (R$)': round(preco_novo, 2),
        'Fornecedor do Preço Novo': forn_novo,
        'Variação (Δ%)': round(variacao, 2),
        'Tendência': tendencia
    })

df_final = pd.DataFrame(resultados)

# Exibição da Tabela Consolidada na Ordem Exata Solicitada
st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

st.dataframe(df_final.style.format({
    'Qtd': '{:,.0f}',
    'Último Preço Hist. (R$)': 'R$ {:.2f}',
    'Novo Preço Unit. (R$)': 'R$ {:.2f}',
    'Variação (Δ%)': '{:+.2f}%'
}), use_container_width=True)

# Bloco de Observações Técnicas (Logística e ZFM)
st.markdown("---")
st.subheader("🔍 Observações Logísticas e Fiscais (ZFM)")
st.markdown("""
* **Impacto Logístico:** Avaliação do custo total de frete (modal aéreo/fluvial) para suprimentos oriundos de 'Fora do Estado' em comparação com fornecedores locais de Manaus.
* **Incentivos Fiscais:** Validação da aplicação correta dos benefícios tributários da Zona Franca de Manaus (ZFM) para preservação de margem.
* **Picos Fora da Curva:** Análise crítica obrigatória em variações superiores a +5%, verificando oscilações de matéria-prima e custos logísticos antes da emissão da O.C.
""")

# Seção Obrigatória: Insight do Especialista
st.markdown("---")
st.subheader("💡 Insight do Especialista")

itens_em_alta = df_final[df_final['Variação (Δ%)'] > 0]
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

st.success(insight_texto)
