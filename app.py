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

# Barra lateral para upload
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader(
    "Carregar Mapa de Cotação Atual (.csv, .xlsx ou .docx)", 
    type=["csv", "xlsx", "docx"]
)

# 1. Leitura Automática do Histórico direto do GitHub
@st.cache_data
def carregar_historico_github():
    caminho = "historico_compras.csv"
    if os.path.exists(caminho):
        try:
            return pd.read_csv(caminho)
        except Exception as e:
            st.sidebar.error(f"Erro ao ler historico_compras.csv: {e}")
    
    # Fallback compatível com os códigos de exemplo
    return pd.DataFrame({
        'Código': ['0000005177', '0000007519'],
        'Data': ['2026-02-10', '2026-01-20'],
        'Item': ['0001', '0002'],
        'Descrição Resumida': ['PAINEL DIVIS CEGO MAD AGLOM BG 1200X2110MM', 'FECHADURA CILINDRO INOX POLIDO PORTA DIVIS CHAV/BOTAO 90MM'],
        'Qtd': [15.0, 2.0],
        'Preço Unit. (R$)': [375.58, 83.36],
        'Fornecedor': ['CAA Com. e Ind. Amaz. de Alumínio Ltda.', 'CAA Comércio Amazonense de Alumínio Ltda.']
    })

historico = carregar_historico_github()

# 2. Leitura de Cotação (Suporte a DOCX, CSV e XLSX)
def extrair_tabela_docx(arquivo_docx):
    doc = docx.Document(arquivo_docx)
    dados = []
    for tabela in doc.tables:
        for linha in tabela.rows:
            texto_linha = [celula.text.strip() for celula in linha.cells]
            dados.append(texto_linha)
    if len(dados) > 1:
        df = pd.DataFrame(dados[1:], columns=dados[0])
        return df
    return pd.DataFrame()

@st.cache_data
def carregar_cotacao_padrao():
    return pd.DataFrame({
        'Item': ['0001', '0002'],
        'Código': ['0000005177', '0000007519'],
        'Descrição Resumida': [
            'PAINEL DIVIS CEGO MAD AGLOM BG 1200X2110MM', 
            'FECHADURA CILINDRO INOX POLIDO PORTA DIVIS CHAV/BOTAO 90MM'
        ],
        'Qtd': [15.00, 2.00],
        'Novo Preço Unit. (R$)': [183.62, 70.07],
        'Fornecedor do Preço Novo': [
            'CENTRO DO ALUMINIO INDUSTRIA E COMERCIO DE FERRAGENS, FERRAM', 
            'CENTRO DO ALUMINIO INDUSTRIA E COMERCIO DE FERRAGENS, FERRAM'
        ]
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
                cotacao = carregar_cotacao_padrao()
    except Exception as e:
        st.sidebar.error(f"Erro ao processar arquivo: {e}")
        cotacao = carregar_cotacao_padrao()
else:
    cotacao = carregar_cotacao_padrao()

# Padronização de nomes de colunas
historico.columns = [str(c).strip() for c in historico.columns]
cotacao.columns = [str(c).strip() for c in cotacao.columns]

# Motor de processamento idêntico ao layout de referência
resultados = []

for idx, row_cot in cotacao.iterrows():
    num_item = str(row_cot.get('Item', f"{idx+1:04d}"))
    if len(num_item) < 4:
        num_item = num_item.zfill(4)
        
    codigo = str(row_cot.get('Código', row_cot.get('Codigo', row_cot.get('SKU', 'N/D')))).replace('.', '').replace(' ', '')
    desc = str(row_cot.get('Descrição Resumida', row_cot.get('Descricao Resumida', row_cot.get('Descrição', ''))))
    
    try:
        qtd = float(str(row_cot.get('Qtd', row_cot.get('Quantidade', 0))).replace('R$', '').replace(',', '.'))
    except:
        qtd = 0.0

    try:
        preco_novo = float(str(row_cot.get('Novo Preço Unit. (R$)', row_cot.get('Novo Preco Unit. (R$)', row_cot.get('Preço Unit. (R$)', 0)))).replace('R$', '').replace('.', '').replace(',', '.'))
    except:
        preco_novo = 0.0

    forn_novo = str(row_cot.get('Fornecedor do Preço Novo', row_cot.get('Fornecedor', 'Não informado')))
    
    # Busca correspondente no histórico
    match_hist = historico[historico['Código'].astype(str).str.replace('.', '').str.replace(' ', '') == codigo] if not historico.empty else pd.DataFrame()
    
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
        
    # Cálculos de Variação (Δ%) e Padrão de Tendência Exato da Imagem
    if ultimo_preco > 0:
        variacao = ((preco_novo - ultimo_preco) / ultimo_preco) * 100
    else:
        variacao = 0.0
        
    if variacao < 0:
        tendencia = f"Queda (Favorável)"
    elif variacao > 0:
        tendencia = f"Alta (Desfavorável)"
    else:
        tendencia = f"Estabilidade"
        
    resultados.append({
        'Item': num_item,
        'Código': codigo,
        'Descrição Resumida': desc,
        'Qtd': qtd,
        'Último Preço Hist. (R$)': round(ultimo_preco, 2),
        'Fornecedor do Último Preço': forn_hist,
        'Novo Preço Unit. (R$)': round(preco_novo, 2),
        'Fornecedor do Preço Novo': forn_novo,
        'Variação (Δ%)': variacao,
        'Tendência': tendencia
    })

df_final = pd.DataFrame(resultados)

# Exibição da Tabela Consolidada rigorosamente na ordem solicitada
st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

st.dataframe(df_final.style.format({
    'Qtd': '{:,.2f}',
    'Último Preço Hist. (R$)': 'R$ {:,.2f}',
    'Novo Preço Unit. (R$)': 'R$ {:,.2f}',
    'Variação (Δ%)': '{:+,.2f}%'
}), use_container_width=True)

# Bloco de Observações Técnicas
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

itens_em_queda = df_final[df_final['Variação (Δ%)'] < 0]
if not itens_em_queda.empty:
    insight_texto = (
        f"Identificada oportunidade expressiva de economia em **{len(itens_em_queda)} item(ns)** com redução de custos favorável. "
        "Recomenda-se a **homologação imediata com o novo fornecedor** para captura dos ganhos de margem, "
        "certificando-se de que os prazos de entrega e condições logísticas para Manaus atendem ao cronograma operacional."
    )
else:
    insight_texto = (
        "Cenário de alta nos preços detectado. Recomenda-se a **renegociação com base no histórico de volume** "
        "ou busca por fornecedores alternativos na praça local para evitar impacto no orçamento."
    )

st.success(insight_texto)
