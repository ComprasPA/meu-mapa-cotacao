import streamlit as st
import pandas as pd
import numpy as np
import os
import docx

# Configuração da Página
st.set_page_config(
    page_title="Mapa de Cotação e Análise de Custos",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Gestão Estratégica de Suprimentos | Mapa de Cotação & Histórico")
st.markdown("Plataforma analítica para homologação de preços, variação de custos e tomada de decisão comercial.")

# Barra lateral para upload (agora aceitando CSV, XLSX e DOCX)
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader(
    "Carregar Mapa de Cotação Atual (.csv, .xlsx ou .docx)", 
    type=["csv", "xlsx", "docx"]
)

# 1. Leitura Automática do Histórico direto do GitHub (historico_compras.csv)
@st.cache_data
def carregar_historico_github():
    caminho = "historico_compras.csv"
    if os.path.exists(caminho):
        try:
            return pd.read_csv(caminho)
        except Exception as e:
            st.sidebar.error(f"Erro ao ler historico_compras.csv do GitHub: {e}")
    
    # Fallback caso o arquivo não seja encontrado no GitHub
    return pd.DataFrame({
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Data': ['2026-02-10', '2026-01-20', '2026-03-12'],
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [150, 20, 60],
        'Preço Unit. (R$)': [13.20, 85.00, 20.50],
        'Fornecedor': ['Industrial Sul Ltda', 'Papelaria & Cia', 'Global Elétrica SP']
    })

historico = carregar_historico_github()

# 2. Leitura da Cotação (Suporte a CSV, XLSX e extração de tabelas de DOCX)
def extrair_tabela_docx(arquivo_docx):
    doc = docx.Document(arquivo_docx)
    dados = []
    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            texto_linha = [celula.text.strip() for celula in linha.cells]
            dados.append(texto_linha)
    if len(dados) > 1:
        # Assume a primeira linha como cabeçalho
        df = pd.DataFrame(dados[1:], columns=dados[0])
        return df
    return pd.DataFrame()

@st.cache_data
def carregar_cotacao_padrao():
    return pd.DataFrame({
        'Item': ['Luva de Raspa', 'Bobina Térmica 80x40', 'Terminal Tubular'],
        'Código': ['SKU001', 'SKU002', 'SKU003'],
        'Descrição Resumida': ['Luva raspa couro cano curto', 'Bobina papel térmico cx c/ 30', 'Terminal tubular 2.5mm pct 100un'],
        'Qtd': [120, 25, 40],
        'Novo Preço Unit. (R$)': [14.00, 82.00, 23.50],
        'Fornecedor do Preço Novo': ['EPI Manaus Comércio', 'Papelaria & Cia', 'Metaltec Suprimentos']
    })

cotacao = pd.DataFrame()

if uploaded_cot is not None:
    nome_arquivo = uploaded_cot.name.lower()
    try:
        if nome_arquivo.endswith('.csv'):
            cotacao = pd.read_csv(uploaded_cot)
        elif nome_arquivo.endswith(('.xlsx', '.xls')):
            cotacao = pd.read_excel(uploaded_cot)
        elif nome_arquivo.endswith('.docx'):
            cotacao = extrair_tabela_docx(uploaded_cot)
            if cotacao.empty:
                st.sidebar.warning("Nenhuma tabela encontrada dentro do arquivo DOCX. Usando dados padrão.")
                cotacao = carregar_cotacao_padrao()
    except Exception as e:
        st.sidebar.error(f"Erro ao processar o arquivo enviado: {e}")
        cotacao = carregar_cotacao_padrao()
else:
    cotacao = carregar_cotacao_padrao()

# Limpeza e normalização de colunas
historico.columns = [str(c).strip() for c in historico.columns]
cotacao.columns = [str(c).strip() for c in cotacao.columns]

# Motor de processamento do comparativo
resultados = []

for _, row_cot in cotacao.iterrows():
    # Mapeamento flexível de colunas
    codigo = str(row_cot.get('Código', row_cot.get('Codigo', row_cot.get('SKU', row_cot.get('Cód.', 'N/D')))))
    item = str(row_cot.get('Item', row_cot.get('Produto', row_cot.get('Material', 'Genérico'))))
    desc = str(row_cot.get('Descrição Resumida', row_cot.get('Descricao Resumida', row_cot.get('Descrição', ''))))
    
    try:
        qtd = float(str(row_cot.get('Qtd', row_cot.get('Quantidade', row_cot.get('Qtde', 0)))).replace('R$', '').replace(',', '.'))
    except:
        qtd = 0.0

    try:
        preco_novo = float(str(row_cot.get('Novo Preço Unit. (R$)', row_cot.get('Novo Preco Unit. (R$)', row_cot.get('Preço Unit. (R$)', row_cot.get('Preço Unitário', row_cot.get('Preço', 0))))))
                           .replace('R$', '').replace('.', '').replace(',', '.'))
    except:
        preco_novo = 0.0

    forn_novo = str(row_cot.get('Fornecedor do Preço Novo', row_cot.get('Fornecedor', row_cot.get('Empresa', 'Não informado'))))
    
    # Busca no Histórico do GitHub
    match_hist = historico[historico.astype(str).apply(lambda x: x.str.contains(codigo, case=False)).any(axis=1)] if not historico.empty else pd.DataFrame()
    
    if not match_hist.empty:
        col_preco_hist = [c for c in match_hist.columns if 'preço' in c.lower() or 'preco' in c.lower() or 'unit' in c.lower()]
        col_forn_hist = [c for c in match_hist.columns if 'fornecedor' in c.lower()]
        
        if col_preco_hist:
            try:
                ultimo_preco = float(str(match_hist.iloc[-1][col_preco_hist[0]]).replace('R$', '').replace('.', '').replace(',', '.'))
            except:
                ultimo_preco = preco_novo
        else:
            ultimo_preco = preco_novo
            
        if col_forn_hist:
            forn_hist = str(match_hist.iloc[-1][col_forn_hist[0]])
        else:
            forn_hist = "Histórico Anterior"
    else:
        ultimo_preco = preco_novo
        forn_hist = "Sem Histórico"
        
    # Cálculos de Variação (Δ%) e Tendência
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

# Exibição da Tabela Consolidada na Ordem Exata Requerida
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
