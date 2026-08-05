import streamlit as st
import pandas as pd
import numpy as np
import os
import docx
from fpdf import FPDF
import unicodedata
import email
from bs4 import BeautifulSoup

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

# Gerenciamento de Estado
if 'limpar_cache' not in st.session_state:
    st.session_state.limpar_cache = False

# Barra lateral para upload e ações de controle
st.sidebar.header("📁 Fontes de Dados")
uploaded_cot = st.sidebar.file_uploader(
    "Carregar Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml)", 
    type=["csv", "xlsx", "docx", "mhtml", "html"]
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
    if not val_str or val_str.lower() in ['nan', 'total item', 'total', '##########', 'a vista', '25 dias', 'item', 'código', 'produto']:
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

def padronizar_codigo_10_digitos(codigo):
    """Padroniza rigorosamente o código do produto com 10 dígitos, completando com zeros à esquerda"""
    if pd.isna(codigo):
        return ""
    apenas_nums = ''.join(filter(str.isdigit, str(codigo)))
    if apenas_nums.isdigit():
        return apenas_nums.zfill(10)
    return str(codigo).strip()

def limpar_texto_pdf(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
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

# 2. Leitura inteligente de arquivos MHTML / HTML do TOTVS
def extrair_tabela_mhtml(arquivo_bytes):
    try:
        conteudo_str = arquivo_bytes.getvalue().decode('utf-8', errors='ignore')
        
        # Se for MHTML multipart
        if "MIME-Version:" in conteudo_str or "multipart/related" in conteudo_str:
            msg = email.message_from_string(conteudo_str)
            html_content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() in ['text/html', 'application/xhtml+xml']:
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                html_content += payload.decode('utf-8', errors='ignore')
                            except:
                                html_content += payload.decode('latin-1', errors='ignore')
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    html_content = payload.decode('utf-8', errors='ignore')
            
            if not html_content:
                html_content = conteudo_str
        else:
            html_content = conteudo_str

        # Extrai tabelas via BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Procura por elementos de tabela tradicionais ou grids do TOTVS
        tabelas = soup.find_all('table')
        if not tabelas:
            # Tenta procurar divs que simulam grids ou dataframes no webapp
            tabelas = soup.find_all(['div'], class_=lambda x: x and ('grid' in x or 'table' in x or 'browse' in x))
            
        dfs = []
        for tab in tabelas:
            try:
                df_list = pd.read_html(str(tab))
                for d in df_list:
                    if len(d) > 0 and len(d.columns) >= 2:
                        dfs.append(d)
            except:
                continue
                
        if dfs:
            # Retorna a maior tabela encontrada (geralmente o grid principal de cotação)
            df_principal = max(dfs, key=len)
            df_principal.columns = [str(c).strip() for c in df_principal.columns]
            return df_principal
            
    except Exception as e:
        st.error(f"Erro ao processar o arquivo MHTML: {e}")
    return pd.DataFrame()

def extrair_tabela_excel_inteligente(arquivo_excel):
    try:
        xls = pd.ExcelFile(arquivo_excel)
        sheet_name = xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
        
        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)]).lower()
            if 'código' in row_str or 'codigo' in row_str or 'descrição' in row_str or 'vlr' in row_str or 'item' in row_str:
                header_row_idx = idx
                break
                
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row_idx, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao processar o arquivo Excel: {e}")
        return pd.DataFrame()

def extrair_tabela_docx_limpa(arquivo_docx):
    try:
        doc = docx.Document(arquivo_docx)
        linhas_validas = []
        for tabela in doc.tables:
            for linha in tabela.rows:
                celulas = [str(cel.text).strip().replace('\n', ' ') for cel in linha.cells]
                if any(celulas):
                    linhas_validas.append(celulas)
        if linhas_validas:
            max_cols = max(len(l) for l in linhas_validas)
            headers = [f"Col_{i}" for i in range(max_cols)]
            dados_norm = [l + [''] * (max_cols - len(l)) for l in linhas_validas]
            return pd.DataFrame(dados_norm, columns=headers)
    except Exception as e:
        st.error(f"Erro ao processar o documento Word: {e}")
    return pd.DataFrame()

# O painel permanece limpo se nenhum arquivo for enviado
cotacao = pd.DataFrame()
if uploaded_cot is not None:
    nome = uploaded_cot.name.lower()
    try:
        if nome.endswith('.csv'):
            cotacao = pd.read_csv(uploaded_cot, dtype=str)
        elif nome.endswith(('.xlsx', '.xls')):
            cotacao = extrair_tabela_excel_inteligente(uploaded_cot)
        elif nome.endswith('.docx'):
            cotacao = extrair_tabela_docx_limpa(uploaded_cot)
        elif nome.endswith(('.mhtml', '.html', '.mht')):
            cotacao = extrair_tabela_mhtml(uploaded_cot)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        cotacao = pd.DataFrame()

# Se nenhum arquivo foi carregado ou a tabela estiver vazia, exibe instrução inicial
if cotacao.empty:
    st.info("👈 Por favor, faça o upload do seu Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml) na barra lateral para iniciar a análise.")
else:
    cotacao.columns = [str(c).strip() for c in cotacao.columns]

    def achar_coluna(df, termos):
        for col in df.columns:
            c_low = str(col).lower()
            if any(t in c_low for t in termos):
                return col
        return None

    c_item = achar_coluna(cotacao, ['item'])
    c_cod = achar_coluna(cotacao, ['código', 'codigo', 'produto', 'sku'])
    c_desc = achar_coluna(cotacao, ['descrição', 'descricao'])
    c_qtd = achar_coluna(cotacao, ['qtd', 'quantidade'])
    c_vlr = achar_coluna(cotacao, ['vlr. unitário', 'vlr unitario', 'unitario', 'preço', 'preco', 'vlr'])
    c_forn = achar_coluna(cotacao, ['fornecedor', 'empresa'])
    c_status = achar_coluna(cotacao, ['status'])

    if c_status and not cotacao.empty:
        df_vencedores = cotacao[cotacao[c_status].astype(str).str.contains('vencedor|melhor preço', case=False, na=False)]
        if not df_vencedores.empty:
            cotacao = df_vencedores

    if c_cod and not cotacao.empty:
        cotacao = cotacao.drop_duplicates(subset=[c_cod])

    resultados = []
    item_contador = 1

    for idx, row in cotacao.iterrows():
        num_item = str(row[c_item] if c_item and pd.notna(row[c_item]) else f"{item_contador:04d}").zfill(4)
        
        raw_cod = str(row[c_cod] if c_cod and pd.notna(row[c_cod]) else f"SKU{item_contador}")
        codigo_original = padronizar_codigo_10_digitos(raw_cod)
        codigo_busca = codigo_original
        
        desc = str(row[c_desc] if c_desc and pd.notna(row[c_desc]) else 'Descrição não informada')
        qtd = limpar_valor(row[c_qtd] if c_qtd and pd.notna(row[c_qtd]) else 1)
        preco_novo = limpar_valor(row[c_vlr] if c_vlr and pd.notna(row[c_vlr]) else 0)
        forn_novo = str(row[c_forn] if c_forn and pd.notna(row[c_forn]) else 'Fornecedor não informado')
        
        if codigo_original.lower().replace('0', '') in ['código', 'codigo', 'item', 'produto', 'nan']:
            continue

        item_contador += 1

        ultimo_preco = preco_novo
        forn_hist = "Sem Histórico"
        
        if not historico.empty:
            match_linhas = []
            for h_idx, h_row in historico.iterrows():
                for col_idx in h_row.index:
                    val_celula = str(h_row[col_idx])
                    if padronizar_codigo_10_digitos(val_celula) == codigo_busca and codigo_busca != '0000000000':
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

    if df_final.empty:
        st.warning("⚠️ Nenhum item válido encontrado no arquivo MHTML. Certifique-se de salvar a página completa do TOTVS.")
    else:
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

        # Função para Gerar PDF do Relatório
        def gerar_pdf(df):
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("helvetica", "B", 14)
            pdf.cell(0, 10, limpar_texto_pdf("Mapa de Cotacao & Comparativo Historico - Suprimentos"), 0, 1, "C")
            pdf.ln(5)
            
            pdf.set_font("helvetica", "B", 8)
            col_widths = [12, 25, 62, 15, 25, 45, 25, 45, 18, 20]
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

        # Seção Obrigatória: Insight Rápido do Especialista Local
        st.markdown("---")
        st.subheader("💡 Insight Rápido do Especialista")

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
