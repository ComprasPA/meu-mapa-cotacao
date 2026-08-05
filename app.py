import streamlit as st
import pandas as pd
import numpy as np
import os
import docx
from fpdf import FPDF
import unicodedata

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

# Gerenciamento de Estado para Limpar Histórico/Sessão
if 'limpar_cache' not in st.session_state:
    st.session_state.limpar_cache = False

# Barra lateral para upload e ações de controle
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader(
    "Carregar Mapa de Cotação Atual (.csv, .xlsx ou .docx)", 
    type=["csv", "xlsx", "docx"]
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Ações e Ferramentas")

if st.sidebar.button("🧹 Limpar Histórico / Cache"):
    st.cache_data.clear()
    st.session_state.clear()
    st.sidebar.success("Cache e histórico limpos com sucesso!")
    st.rerun()

# Funções de Conversão e Formatação Padrão Brasileiro Rigoroso (X.XXX,XX)
def limpar_valor(valor):
    if pd.isna(valor):
        return 0.0
    val_str = str(valor).replace('R$', '').strip()
    if not val_str or val_str.lower() in ['nan', 'total item', 'total', '##########', 'a vista', '25 dias']:
        return 0.0
    
    if '.' in val_str and ',' in val_str:
        if val_str.find('.') < val_str.find(','):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif val_str.count('.') > 1:
        val_str = val_str.replace('.', '')
    
    try:
        return float(val_str)
    except:
        return 0.0

def formatar_brl(valor):
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    return f"R$ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_qtd(valor):
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    return f"{val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_pct(valor):
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    return f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

def extrair_numeros(codigo):
    if pd.isna(codigo):
        return ""
    apenas_nums = ''.join(filter(str.isdigit, str(codigo)))
    return str(int(apenas_nums)) if apenas_nums.isdigit() else str(codigo).strip()

def limpar_texto_pdf(texto):
    """Remove acentos, caracteres especiais e converte de forma segura para o PDF"""
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    # Força a codificação para latin-1 substituindo qualquer caractere estranho remanescente
    return texto_sem_acento.encode('latin-1', 'replace').decode('latin-1')

# 1. Leitura do Histórico do GitHub
@st.cache_data
def carregar_historico_github():
    caminho = "historico_compras.csv"
    if os.path.exists(caminho):
        try:
            df = pd.read_csv(caminho, header=None, dtype=str)
            return df, "Conectado com sucesso ao GitHub (historico_compras.csv)"
        except Exception as e:
            return pd.DataFrame(), f"Erro ao ler historico_compras.csv: {e}"
    else:
        df = pd.DataFrame({
            4: ['CAA COMERCIO AMAZONENSE DE ALUMINIO LTDA.'], 
            6: ['0000005177'],                              
            10: ['375.58']                                   
        })
        return df, "historico_compras.csv não encontrado. Usando dados padrão."

historico, status_historico = carregar_historico_github()
st.sidebar.info(f"ℹ️ **Status:** {status_historico}")

# 2. Leitura inteligente de arquivos DOCX do TOTVS
def extrair_tabela_docx_inteligente(arquivo_docx):
    try:
        doc = docx.Document(arquivo_docx)
        todas_linhas = []
        for tabela in doc.tables:
            for linha in tabela.rows:
                texto_linha = [celula.text.strip().replace('\n', ' ') for celula in linha.cells]
                if any(texto_linha):
                    todas_linhas.append(texto_linha)
                    
        if len(todas_linhas) > 0:
            cabecalho_idx = 0
            for idx, linha in enumerate(todas_linhas[:5]):
                texto_unido = " ".join(linha).lower()
                if 'item' in texto_unido or 'codigo' in texto_unido or 'descrição' in texto_unido or 'unitário' in texto_unido:
                    cabecalho_idx = idx
                    break
                    
            headers = todas_linhas[cabecalho_idx]
            dados = todas_linhas[cabecalho_idx+1:]
            
            if not dados:
                return pd.DataFrame()
                
            df_temp = pd.DataFrame(dados)
            if len(headers) < len(df_temp.columns):
                headers = [f"Col_{i}" for i in range(len(df_temp.columns))]
            elif len(headers) > len(df_temp.columns):
                headers = headers[:len(df_temp.columns)]
                
            return pd.DataFrame(dados, columns=headers)
    except Exception as e:
        st.error(f"Erro ao processar o documento Word: {e}")
    return pd.DataFrame()

# O painel permanece limpo se nenhum arquivo for enviado
cotacao = pd.DataFrame()
if uploaded_cot is not None:
    nome = uploaded_cot.name.lower()
    try:
        if nome.endswith('.csv'):
            cotacao = pd.read_csv(uploaded_cot)
        elif nome.endswith(('.xlsx', '.xls')):
            cotacao = pd.read_excel(uploaded_cot)
        elif nome.endswith('.docx'):
            cotacao = extrair_tabela_docx_inteligente(uploaded_cot)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        cotacao = pd.DataFrame()

# Se nenhum arquivo foi carregado ou a tabela estiver vazia, exibe instrução inicial
if cotacao.empty:
    st.info("👈 Por favor, faça o upload do seu Mapa de Cotação Atual (.csv, .xlsx ou .docx) na barra lateral para iniciar a análise.")
else:
    cotacao.columns = [str(c).strip() for c in cotacao.columns]

    def identificar_colunas_docx(df):
        col_item, col_cod, col_desc, col_qtd, col_vlr, col_forn = None, None, None, None, None, None
        
        for col in df.columns:
            c_low = str(col).lower()
            if 'item' in c_low and not col_item: col_item = col
            elif any(k in c_low for k in ['cód', 'cod', 'sku']) and not col_cod: col_cod = col
            elif any(k in c_low for k in ['descri', 'produto', 'material']) and not col_desc: col_desc = col
            elif any(k in c_low for k in ['qtd', 'quant']) and not col_qtd: col_qtd = col
            elif any(k in c_low for k in ['vlr', 'unit', 'preço', 'preco']) and not col_vlr: col_vlr = col
            elif any(k in c_low for k in ['fornecedor', 'empresa', 'razão']) and not col_forn: col_forn = col

        if not col_cod or not col_vlr:
            for col in df.columns:
                valores_serie = df[col].dropna()
                if valores_serie.empty:
                    continue
                amostra = " ".join(valores_serie.astype(str).values).lower()
                if not col_cod and any(str(digitos).isdigit() and len(str(digitos)) >= 4 for digitos in valores_serie):
                    col_cod = col
                if not col_vlr and 'r$' in amostra:
                    col_vlr = col
                    
        return col_item, col_cod, col_desc, col_qtd, col_vlr, col_forn

    c_item, c_cod, c_desc, c_qtd, c_vlr, c_forn = identificar_colunas_docx(cotacao)

    resultados = []

    for idx, row in cotacao.iterrows():
        linha_texto = " ".join([str(v) for v in row.values]).lower()
        if 'total' in linha_texto and not c_cod:
            continue

        num_item = str(row[c_item] if c_item and pd.notna(row[c_item]) else f"{idx+1:04d}").zfill(4)
        codigo_original = str(row[c_cod] if c_cod and pd.notna(row[c_cod]) else f"SKU{idx+1}")
        codigo_busca = extrair_numeros(codigo_original)
        
        desc = str(row[c_desc] if c_desc and pd.notna(row[c_desc]) else 'Descrição não informada')
        qtd = limpar_valor(row[c_qtd] if c_qtd and pd.notna(row[c_qtd]) else 1)
        preco_novo = limpar_valor(row[c_vlr] if c_vlr and pd.notna(row[c_vlr]) else 0)
        forn_novo = str(row[c_forn] if c_forn and pd.notna(row[c_forn]) else 'Fornecedor não informado')
        
        if preco_novo == 0.0:
            for val in row.values:
                v_limpo = limpar_valor(val)
                if v_limpo > 0 and v_limpo != qtd:
                    preco_novo = v_limpo
                    break

        ultimo_preco = preco_novo
        forn_hist = "Sem Histórico"
        
        if not historico.empty:
            match_linhas = []
            for h_idx, h_row in historico.iterrows():
                for col_idx in h_row.index:
                    val_celula = str(h_row[col_idx])
                    if extrair_numeros(val_celula) == codigo_busca and codigo_busca != '':
                        match_linhas.append(h_row)
                        break
                        
            if match_linhas:
                ultima_ocorrencia = match_linhas[-1]
                try:
                    preco_hist_col = limpar_valor(ultima_ocorrencia.get(10, 0))
                    if preco_hist_col > 0:
                        ultimo_preco = preco_hist_col
                    else:
                        for val in ultima_ocorrencia.values:
                            v = limpar_valor(val)
                            if v > 1.0 and v != qtd:
                                ultimo_preco = v
                                break
                except:
                    pass
                    
                try:
                    forn_col = str(ultima_ocorrencia.get(4, ""))
                    if len(forn_col) > 3 and forn_col.lower() not in ['nan', 'a vista', '25 dias']:
                        forn_hist = forn_col
                    else:
                        for val in ultima_ocorrencia.values:
                            t = str(val)
                            if len(t) > 5 and not any(char.isdigit() for char in t[:3]) and t.lower() not in ['a vista', '25 dias']:
                                forn_hist = t
                                break
                except:
                    pass
            
        variacao = ((preco_novo - ultimo_preco) / ultimo_preco) * 100 if ultimo_preco > 0 else 0.0
        
        if variacao < 0:
            tendencia = "Queda (Favorável)"
        elif variacao > 0:
            tendencia = "Alta (Desfavorável)"
        else:
            tendencia = "Estabilidade"
            
        resultados.append({
            'Item': num_item,
            'Código': codigo_original,
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

    colunas_exatas = [
        'Item', 'Código', 'Descrição Resumida', 'Qtd', 
        'Último Preço Hist. (R$)', 'Fornecedor do Último Preço', 
        'Novo Preço Unit. (R$)', 'Fornecedor do Preço Novo', 
        'Variação (Δ%)', 'Tendência'
    ]

    df_final = df_final[colunas_exatas]

    df_display = df_final.copy()
    df_display['Qtd'] = df_display['Qtd'].apply(formatar_qtd)
    df_display['Último Preço Hist. (R$)'] = df_display['Último Preço Hist. (R$)'].apply(formatar_brl)
    df_display['Novo Preço Unit. (R$)'] = df_display['Novo Preço Unit. (R$)'].apply(formatar_brl)
    df_display['Variação (Δ%)'] = df_display['Variação (Δ%)'].apply(formatar_pct)

    # Exibição do painel interativo formatado
    st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

    st.markdown(df_display.to_html(escape=False, index=False), unsafe_allow_html=True)

    # Função para Gerar PDF do Relatório com tratamento blindado
    def gerar_pdf(df):
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font("helvetica", "B", 14)
        pdf.cell(0, 10, limpar_texto_pdf("Mapa de Cotacao & Comparativo Historico - Suprimentos"), 0, 1, "C")
        pdf.ln(5)
        
        pdf.set_font("helvetica", "B", 8)
        col_widths = [12, 22, 65, 15, 25, 45, 25, 45, 18, 20]
        headers = [
            "Item", "Codigo", "Descricao", "Qtd", 
            "Ult. Preco", "Forn. Ant.", "Novo Preco", 
            "Forn. Novo", "Var (%)", "Tendencia"
        ]
        
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 7, limpar_texto_pdf(h), 1, 0, "C")
        pdf.ln()
        
        pdf.set_font("helvetica", "", 7)
        for _, row in df.iterrows():
            pdf.cell(col_widths[0], 6, limpar_texto_pdf(str(row['Item'])), 1, 0, "C")
            pdf.cell(col_widths[1], 6, limpar_texto_pdf(str(row['Código'])), 1, 0, "C")
            pdf.cell(col_widths[2], 6, limpar_texto_pdf(str(row['Descrição Resumida'])[:35]), 1, 0, "L")
            pdf.cell(col_widths[3], 6, limpar_texto_pdf(str(row['Qtd'])), 1, 0, "R")
            pdf.cell(col_widths[4], 6, limpar_texto_pdf(formatar_brl(row['Último Preço Hist. (R$)'])), 1, 0, "R")
            pdf.cell(col_widths[5], 6, limpar_texto_pdf(str(row['Fornecedor do Último Preço'])[:25]), 1, 0, "L")
            pdf.cell(col_widths[6], 6, limpar_texto_pdf(formatar_brl(row['Novo Preço Unit. (R$)'])), 1, 0, "R")
            pdf.cell(col_widths[7], 6, limpar_texto_pdf(str(row['Fornecedor do Preço Novo'])[:25]), 1, 0, "L")
            pdf.cell(col_widths[8], 6, limpar_texto_pdf(formatar_pct(row['Variação (Δ%)'])), 1, 0, "R")
            pdf.cell(col_widths[9], 6, limpar_texto_pdf(str(row['Tendência'])), 1, 0, "C")
            pdf.ln()
            
        pdf_output = pdf.output(dest='S')
        if isinstance(pdf_output, str):
            return pdf_output.encode('latin1')
        return bytes(pdf_output)

    # Botão de Download em PDF na interface principal
    st.markdown("---")
    col_bt1, col_bt2 = st.columns([2, 8])
    with col_bt1:
        pdf_bytes = gerar_pdf(df_final)
        st.download_button(
            label="📥 Baixar Mapa em PDF",
            data=pdf_bytes,
            file_name="mapa_de_cotacao_suprimentos.pdf",
            mime="application/pdf"
        )

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
