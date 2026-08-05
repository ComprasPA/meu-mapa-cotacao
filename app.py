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

# Estilização visual corporativa
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    h1 { color: #1f2c34; font-family: 'Helvetica Neue', sans-serif; }
    table { width: 100% !important; border-collapse: collapse !important; }
    th {
        background-color: #205081 !important;
        color: white !important;
        text-align: center !important;
        font-weight: bold !important;
        padding: 10px !important;
        border: 1px solid #dddddd !important;
        font-size: 14px !important;
    }
    td {
        padding: 10px !important;
        border: 1px solid #dddddd !important;
        color: #000000 !important;
        font-size: 13px !important;
        text-align: right;
    }
    td:nth-child(1), td:nth-child(2), td:nth-child(3), td:nth-child(6), td:nth-child(8), td:nth-child(10) {
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Gestão Estratégica de Suprimentos | Mapa de Cotação & Histórico")
st.markdown("Plataforma analítica para homologação de preços, variação de custos e tomada de decisão comercial.")

# Barra lateral para upload
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader(
    "Carregar Mapa de Cotação Atual (.csv, .xlsx ou .docx)", 
    type=["csv", "xlsx", "docx"]
)

# Função auxiliar para conversão rigorosa de valores monetários e numéricos (Tratamento de pontos e vírgulas)
def limpar_valor(valor):
    if pd.isna(valor):
        return 0.0
    val_str = str(valor).replace('R$', '').strip()
    if not val_str or val_str.lower() == 'nan':
        return 0.0
    
    # Se contiver tanto ponto quanto vírgula (ex: 18.362,00)
    if '.' in val_str and ',' in val_str:
        if val_str.find('.') < val_str.find(','):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        # Se contiver apenas vírgula (ex: 18362,00)
        val_str = val_str.replace('.', '').replace(',', '.')
    elif val_str.count('.') > 1:
        # Se contiver múltiplos pontos de milhar sem vírgula (ex: 18.362)
        val_str = val_str.replace('.', '')
    
    try:
        return float(val_str)
    except:
        return 0.0

# Formatação monetária padrão brasileiro (R$ X.XXX,XX)
def formatar_brl(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

# Formatação percentual padrão brasileiro
def formatar_pct(valor):
    return f"{valor:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

# 1. Leitura do histórico de compras do GitHub
@st.cache_data
def carregar_historico_github():
    caminho = "historico_compras.csv"
    if os.path.exists(caminho):
        try:
            df = pd.read_csv(caminho)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except:
            pass
    
    return pd.DataFrame({
        'Código': ['0000005177', '0000007519'],
        'Prc Unitario': [375.58, 83.36],
        'Nome Fornece': [
            'CAA Com. e Ind. Amaz. de Alumínio Ltda.', 
            'CAA Comércio Amazonense de Alumínio Ltda.'
        ]
    })

historico = carregar_historico_github()
historico.columns = [str(c).strip() for c in historico.columns]

# 2. Leitura e extração do arquivo DOCX da Cotação
def extrair_tabela_docx(arquivo_docx):
    doc = docx.Document(arquivo_docx)
    dados = []
    for tabela in doc.tables:
        for linha in tabela.rows:
            texto_linha = [celula.text.strip() for celula in linha.cells]
            dados.append(texto_linha)
    if len(dados) > 1:
        return pd.DataFrame(dados[1:], columns=dados[0])
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
        'Vlr. Unitário': [183.62, 70.07],
        'Fornecedor': [
            'CENTRO DO ALUMINIO INDUSTRIA E COMERCIO DE FERRAGENS, FERRAM', 
            'CENTRO DO ALUMINIO INDUSTRIA E COMERCIO DE FERRAGENS, FERRAM'
        ]
    })

cotacao = pd.DataFrame()

if uploaded_cot is not None:
    nome = uploaded_cot.name.lower()
    try:
        if nome.endswith('.csv'):
            cotacao = pd.read_csv(uploaded_cot)
        elif nome.endswith(('.xlsx', '.xls')):
            cotacao = pd.read_excel(uploaded_cot)
        elif nome.endswith('.docx'):
            cotacao = extrair_tabela_docx(uploaded_cot)
            if cotacao.empty:
                cotacao = carregar_cotacao_padrao()
    except:
        cotacao = carregar_cotacao_padrao()
else:
    cotacao = carregar_cotacao_padrao()

cotacao.columns = [str(c).strip() for c in cotacao.columns]

def achar_coluna(df, termos):
    for col in df.columns:
        if any(t.lower() in col.lower() for t in termos):
            return col
    return None

c_item = achar_coluna(cotacao, ['item'])
c_cod = achar_coluna(cotacao, ['código', 'codigo', 'sku', 'cod'])
c_desc = achar_coluna(cotacao, ['descri', 'resumida', 'detalhe'])
c_qtd = achar_coluna(cotacao, ['qtd', 'quantidade', 'quant'])
c_vlr = achar_coluna(cotacao, ['vlr. unitário', 'vlr unitario', 'preço', 'preco', 'unit'])
c_forn = achar_coluna(cotacao, ['fornecedor', 'empresa'])

resultados = []

for idx, row in cotacao.iterrows():
    num_item = str(row[c_item] if c_item and pd.notna(row[c_item]) else f"{idx+1:04d}").zfill(4)
    codigo = str(row[c_cod] if c_cod and pd.notna(row[c_cod]) else 'N/D').replace('.', '').replace(' ', '')
    desc = str(row[c_desc] if c_desc and pd.notna(row[c_desc]) else '')
    
    qtd = limpar_valor(row[c_qtd] if c_qtd and pd.notna(row[c_qtd]) else 0)
    preco_novo = limpar_valor(row[c_vlr] if c_vlr and pd.notna(row[c_vlr]) else 0)
    forn_novo = str(row[c_forn] if c_forn and pd.notna(row[c_forn]) else 'Não informado')
    
    # Correlação com o histórico de compras do GitHub
    match = pd.DataFrame()
    if not historico.empty:
        col_cod_hist = achar_coluna(historico, ['código', 'codigo', 'sku', 'cod'])
        if col_cod_hist:
            match = historico[historico[col_cod_hist].astype(str).str.replace('.', '').str.replace(' ', '') == codigo]
    
    if not match.empty:
        col_prc_hist = achar_coluna(match, ['prc unitario', 'preco', 'preço', 'unit'])
        col_forn_hist = achar_coluna(match, ['nome fornece', 'fornecedor'])
        
        ultimo_preco = limpar_valor(match.iloc[-1][col_prc_hist]) if col_prc_hist else preco_novo
        forn_hist = str(match.iloc[-1][col_forn_hist]) if col_forn_hist else "Histórico Anterior"
    else:
        ultimo_preco = preco_novo
        forn_hist = "Sem Histórico"
        
    # Variação (Δ%) e Tendência
    variacao = ((preco_novo - ultimo_preco) / ultimo_preco) * 100 if ultimo_preco > 0 else 0.0
    
    if variacao < 0:
        tendencia = "Queda (Favorável)"
    elif variacao > 0:
        tendencia = "Alta (Desfavorável)"
    else:
        tendencia = "Estabilidade"
        
    resultados.append({
        'Item': num_item,
        'Código': codigo,
        'Descrição Resumida': desc,
        'Qtd': qtd,
        'Último Preço Hist. (R$)': ultimo_preco,
        'Fornecedor do Último Preço': forn_hist,
        'Novo Preço Unit. (R$)': preco_novo,
        'Fornecedor do Preço Novo': forn_novo,
        'Variação (Δ%)': variacao,
        'Tendência': tendencia
    })

df_final = pd.DataFrame(resultados)

# Ordem exata solicitada pelo usuário
colunas_exatas = [
    'Item', 
    'Código', 
    'Descrição Resumida', 
    'Qtd', 
    'Último Preço Hist. (R$)', 
    'Fornecedor do Último Preço', 
    'Novo Preço Unit. (R$)', 
    'Fornecedor do Preço Novo', 
    'Variação (Δ%)', 
    'Tendência'
]

df_final = df_final[colunas_exatas]

# Criação de cópia formatada visualmente com padrão brasileiro correto (R$ X.XXX,XX)
df_display = df_final.copy()
df_display['Qtd'] = df_display['Qtd'].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
df_display['Último Preço Hist. (R$)'] = df_display['Último Preço Hist. (R$)'].apply(formatar_brl)
df_display['Novo Preço Unit. (R$)'] = df_display['Novo Preço Unit. (R$)'].apply(formatar_brl)
df_display['Variação (Δ%)'] = df_display['Variação (Δ%)'].apply(formatar_pct)

# Exibição do painel interativo formatado
st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

st.markdown(df_display.to_html(escape=False, index=False), unsafe_allow_html=True)

# Bloco de Observações Técnicas (ZFM e Logística)
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
