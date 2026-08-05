import streamlit as pd_st
import pandas as pd
import numpy as np
import os

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

# Sidebar para upload de dados (Planilhas e Documentos estruturados)
pd_st.sidebar.header("📁 Fontes de Dados (Upload)")
uploaded_file = pd_st.sidebar.file_uploader("Carregar Mapa de Cotação Atual (.csv ou .xlsx)", type=["csv", "xlsx"])
uploaded_history = pd_st.sidebar.file_uploader("Carregar Histórico de Compras (.csv ou .xlsx)", type=["csv", "xlsx"])

# Função auxiliar para ler arquivos suportados
def ler_arquivo(arquivo_upload, caminho_padrao):
    try:
        if arquivo_upload is not None:
            nome = arquivo_upload.name.lower()
            if nome.endswith('.csv'):
                return pd.read_csv(arquivo_upload)
            elif nome.endswith(('.xlsx', '.xls')):
                return pd.read_excel(arquivo_upload)
        elif os.path.exists(caminho_padrao):
            if caminho_padrao.endswith('.csv'):
                return pd.read_csv(caminho_padrao)
            elif caminho_padrao.endswith(('.xlsx', '.xls')):
                return pd.read_excel(caminho_padrao)
    except Exception as e:
        pd_st.sidebar.error(f"Erro ao ler o arquivo: {e}")
    return None

# Função para padronizar colunas independentemente de como vierem no documento
def padronizar_colunas(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    renomear = {}
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ['codigo', 'código', 'cod', 'sku', 'item code']:
            renomear[col] = 'Código'
        elif col_lower in ['item', 'produto', 'material', 'descrição', 'descricao']:
            renomear[col] = 'Item'
        elif col_lower in ['descricao resumida', 'descrição resumida', 'descrição', 'detalhes']:
            renomear[col] = 'Descrição Resumida'
        elif col_lower in ['qtd', 'quantidade', 'quant', 'qtde']:
            renomear[col] = 'Qtd'
        elif col_lower in ['preco unit. (r$)', 'preço unit. (r$)', 'preco unitario', 'preço unitário', 'preco', 'preço', 'valor unitario', 'valor unitário']:
            renomear[col] = 'Preço Unit. (R$)'
        elif col_lower in ['novo preco unit. (r$)', 'novo preço unit. (r$)', 'novo preco', 'novo preço', 'preco cotado']:
            renomear[col] = 'Novo Preço Unit. (R$)'
        elif col_lower in ['fornecedor', 'forn', 'empresa']:
            renomear[col] = 'Fornecedor'
        elif col_lower in ['fornecedor do preço novo', 'fornecedor preco novo', 'novo fornecedor']:
            renomear[col] = 'Fornecedor do Preço Novo'
        elif col_lower in ['data', 'dt', 'data da compra']:
            renomear[col] = 'Data'
    return df.rename(columns=renomear)

# Mocks de segurança caso não haja arquivos enviados
def gerar_hist_mock():
    return pd.DataFrame({
        'Código': ['SKU001', 'SKU001', 'SKU002', 'SKU003', 'SKU003'],
        'Data': ['2025-10-15', '2026-02-10', '2026-01-20', '2025-11-05', '2026-03-12'],
        'Item': ['Luva de Raspa', 'Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular', 'Terminal Tubular'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [100, 150, 20, 50, 60],
        'Preço Unit. (R$)': [12.50, 13.20, 85.00, 22.00, 20.50],
        'Fornecedor': ['EPI Manaus Comércio', 'Industrial Sul Ltda', 'Papelaria & Cia', 'Conexões Amazônia', 'Global Elétrica SP']
    })

def gerar_cot_mock():
    return pd.DataFrame({
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [120, 25, 40],
        'Novo Preço Unit. (R$)': [14.00, 82.00, 23.50],
        'Fornecedor do Preço Novo': ['EPI Manaus Comércio', 'Papelaria & Cia', 'Metaltec Suprimentos']
    })

# Carregamento efetivo
hist_raw = ler_arquivo(uploaded_history, "historico_compras.csv")
if hist_raw is None or hist_raw.empty:
    hist_data = gerar_hist_mock()
else:
    hist_data = padronizar_colunas(hist_raw)

cot_raw = ler_arquivo(uploaded_file, "mapa_cotacao.csv")
if cot_raw is None or cot_raw.empty:
    cot_data = gerar_cot_mock()
else:
    cot_data = padronizar_colunas(cot_raw)

# Processamento do Comparativo Histórico na ordem exata solicitada
def processar_comparativo(historico, cotacao):
    resultado = []
    
    if 'Data' in historico.columns:
        historico['Data'] = pd.to_datetime(historico['Data'], errors='coerce')
        historico = historico.sort_values(by='Data', ascending=False)
    
    for _, row in cotacao.iterrows():
        codigo = str(row.get('Código', 'N/D'))
        item = str(row.get('Item', 'Item Genérico'))
        desc = str(row.get('Descrição Resumida', ''))
        qtd_novo = float(row.get('Qtd', 0))
        preco_novo = float(row.get('Novo Preço Unit. (R$)', row.get('Preço Unit. (R$)', 0.0)))
        forn_novo = str(row.get('Fornecedor do Preço Novo', row.get('Fornecedor', 'Não informado')))
        
        hist_item = historico[historico['Código'].astype(str) == codigo] if 'Código' in historico.columns else pd.DataFrame()
        
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
    
    # Ordem rigorosa exigida:
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
* **Impacto Logístico:** Avaliação do custo total de frete (modal aéreo/fluvial) para suprimentos oriundos de 'Fora do Estado' em comparação com fornecedores de Manaus.
* **Incentivos Fiscais:** Validação da aplicação correta dos benefícios tributários da Zona Franca de Manaus (ZFM) para preservar a margem operacional.
* **Picos Fora da Curva:** Análise crítica obrigatória em variações superiores a +5%, verificando oscilações de matéria-prima e custos de transporte antes da emissão do pedido.
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
