import streamlit as st
import pandas as pd
import numpy as np
import os
import docx
from fpdf import FPDF
import unicodedata
import email
from bs4 import BeautifulSoup
import datetime
import time

# Configuração da Página
st.set_page_config(
    page_title="Gestão Estratégica de Compras | Mapa de Cotação",
    page_icon="📊",
    layout="wide"
)

# Estilização visual corporativa estilo Dados Bancários e ocultação de elementos padrão
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    h1 { color: #1f2c34; font-family: 'Helvetica Neue', sans-serif; margin-bottom: 5px; }
    
    .block-container {
        padding-top: 1rem !important;
    }

    /* Oculta completamente a barra superior padrão do Streamlit (Share, GitHub, Menu, etc.) */
    header { visibility: hidden !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }

    /* Anula e oculta qualquer pop-up ou modal de Clear Caches nativo do Streamlit */
    div[data-baseweb="modal"], div.stDialog, div[role="dialog"] {
        display: none !important;
    }

    /* Cabeçalho alinhado com status à direita */
    .status-badge {
        background-color: #e8f0fe;
        color: #1967d2;
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 600;
        font-family: 'Helvetica Neue', sans-serif;
        border: 1px solid #d2e3fc;
    }

    /* Expander fixo no topo de ponta a ponta */
    div[data-testid="stExpander"] {
        border: 1px solid #d9d9d9 !important;
        background-color: #ffffff !important;
        border-radius: 6px !important;
        box-shadow: none !important;
        margin-bottom: 20px !important;
    }
    .streamlit-expanderHeader {
        padding-top: 8px !important;
        padding-bottom: 8px !important;
        min-height: 40px !important;
        font-size: 14px !important;
        background-color: #f8f9fa !important;
        border-radius: 6px !important;
    }
    .streamlit-expanderContent {
        padding: 15px !important;
        background-color: #ffffff !important;
    }

    /* Estilização da Tabela no Estilo Dados Bancários */
    .dataframe {
        width: 100% !important;
        border-collapse: collapse !important;
        font-family: 'Helvetica Neue', sans-serif !important;
    }
    .dataframe th {
        background-color: #2f5597 !important;
        color: white !important;
        text-align: center !important;
        font-weight: bold !important;
        padding: 10px !important;
        border: 1px solid #b4c6e7 !important;
        font-size: 13px !important;
        white-space: nowrap !important;
    }
    .dataframe td {
        padding: 9px 10px !important;
        border: 1px solid #d9d9d9 !important;
        color: #000000 !important;
        font-size: 13px !important;
        text-align: right;
    }
    .dataframe td:nth-child(10), .dataframe th:nth-child(10) {
        white-space: nowrap !important;
        text-align: center !important;
    }
    .dataframe tr:nth-child(even) {
        background-color: #f2f5f9 !important;
    }
    .dataframe tr:nth-child(odd) {
        background-color: #ffffff !important;
    }
    .dataframe td:nth-child(1), .dataframe td:nth-child(2), .dataframe td:nth-child(3), .dataframe td:nth-child(6), .dataframe td:nth-child(8), .dataframe td:nth-child(11) {
        text-align: left;
    }
    </style>
""", unsafe_allow_html=True)

# 1. Leitura do Histórico do GitHub
@st.cache_data
def carregar_historico_github():
    caminho = "historico_compras.csv"
    if os.path.exists(caminho):
        try:
            df = pd.read_csv(caminho, header=None, dtype=str)
            mod_time = os.path.getmtime(caminho)
            data_atualizacao = datetime.datetime.fromtimestamp(mod_time).strftime('%d/%m/%Y')
            return df, f"Base de dados atualizada em: {data_atualizacao}"
        except Exception as e:
            return pd.DataFrame(), f"Erro ao ler historico_compras.csv: {e}"
    else:
        df = pd.DataFrame({
            2: ['174781'],
            4: ['CAA COMERCIO AMAZONENSE DE ALUMINIO LTDA.'], 
            6: ['0000005177'],                              
            10: ['375.58'],
            12: ['05/01/2026']
        })
        return df, "Base de dados atualizada em: Indisponível"

historico, status_historico = carregar_historico_github()

# CAIXA DE CONFIGURAÇÕES (ABRE/FECHA FIXA NO TOPO)
with st.expander("⚙️ Abrir / Fechar Configurações (Upload e Exportação PDF)", expanded=False):
    col_exp1, col_exp2 = st.columns([2, 1])
    with col_exp1:
        st.markdown("### 📁 Upload de Arquivo")
        uploaded_cot = st.file_uploader(
            "Carregar Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml)", 
            type=["csv", "xlsx", "docx", "mhtml", "html"]
        )
    with col_exp2:
        st.markdown("### 📥 Exportar Relatório")
        placeholder_pdf = st.empty()

# Topo do App: Título e Status no canto superior direito
col_title, col_status = st.columns([7, 3])
with col_title:
    st.title("📊 Gestão Estratégica de Compras | Mapa de Cotação")
with col_status:
    st.markdown(f"<div style='text-align: right; margin-top: 15px;'><span class='status-badge'>ℹ️ {status_historico}</span></div>", unsafe_allow_html=True)

st.markdown("---")

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
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except:
        return ""
    if val_float <= 0:
        return ""
    return f"R$ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_qtd(valor):
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    return f"{val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_pct(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    return f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_pct_com_seta(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except:
        val_float = 0.0
    
    val_fmt = f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')
    
    if val_float > 0:
        return f"<span style='color: #c00000; font-size: 15px; font-weight: 900; white-space: nowrap;'>↑ {val_fmt}</span>"
    elif val_float < 0:
        return f"<span style='color: #2ca02c; font-size: 15px; font-weight: 900; white-space: nowrap;'>↓ {val_fmt}</span>"
    else:
        return f"<span style='color: #555555; font-size: 13px; font-weight: bold; white-space: nowrap;'>{val_fmt}</span>"

def padronizar_codigo_10_digitos(codigo):
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

# Leitura de arquivos MHTML / HTML
def extrair_tabela_mhtml(arquivo_bytes):
    try:
        conteudo_str = arquivo_bytes.getvalue().decode('utf-8', errors='ignore')
        html_contents = []
        
        if "MIME-Version:" in conteudo_str or "multipart/related" in conteudo_str:
            msg = email.message_from_string(conteudo_str)
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() in ['text/html', 'application/xhtml+xml']:
                        payload = part.get_payload(decode=True)
                        if payload:
                            try:
                                html_contents.append(payload.decode('utf-8', errors='ignore'))
                            except:
                                html_contents.append(payload.decode('latin-1', errors='ignore'))
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    html_contents.append(payload.decode('utf-8', errors='ignore'))
        else:
            html_contents.append(conteudo_str)

        dfs = []
        for html_content in html_contents:
            soup = BeautifulSoup(html_content, 'html.parser')
            tabelas = soup.find_all('table')
            for tab in tabelas:
                try:
                    df_list = pd.read_html(str(tab))
                    for d in df_list:
                        if len(d) > 0 and len(d.columns) >= 2:
                            dfs.append(d)
                except:
                    continue
                    
        if dfs:
            df_principal = max(dfs, key=lambda x: len(x) * len(x.columns))
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

# Função para Gerar PDF do Relatório com colunas ajustadas para 11 colunas somando exatamente 284mm
def gerar_pdf(df):
    class PDFProfissional(FPDF):
        def __init__(self):
            super().__init__(orientation='L', unit='mm', format='A4')
            self.set_margins(left=6.4, top=19.1, right=6.4)
            self.set_auto_page_break(auto=True, margin=19.1)

        def header(self):
            self.set_fill_color(47, 85, 151)
            self.rect(6.4, 8, 284.2, 20, 'F')
            
            self.set_font("helvetica", "B", 14)
            self.set_text_color(255, 255, 255)
            self.set_xy(6.4, 10)
            self.cell(284.2, 6, limpar_texto_pdf("Mapa de Cotacao & Comparativo Historico"), 0, 1, "C")
            
            self.set_font("helvetica", "", 9)
            self.set_xy(6.4, 16)
            self.cell(284.2, 5, limpar_texto_pdf("Gestao Estratégica de Compras | Parente Andrade"), 0, 1, "C")
            self.ln(10)

        def footer(self):
            self.set_y(-12)
            self.set_font("helvetica", "I", 8)
            self.set_text_color(128, 128, 128)
            data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            self.cell(0, 8, limpar_texto_pdf(f"Gerado em {data_hora} | Pagina {self.page_no()}"), 0, 0, "C")

    pdf = PDFProfissional()
    pdf.add_page()
    
    col_widths = [9, 20, 52, 9, 18, 40, 18, 40, 18, 20, 20]
    headers = [
        "Item", "Codigo", "Descricao", "Qtd", 
        "Novo Preco", "Forn. Novo", "Ult. Preco", 
        "Forn. Ant.", "Preco Med.", "Var(%)", "Tendencia"
    ]
    
    pdf.set_fill_color(47, 85, 151)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 7)
    
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 7, limpar_texto_pdf(h), border=1, fill=True, align="C")
    pdf.ln()
    
    pdf.set_font("helvetica", "", 6.5)
    
    fill = False
    for _, row in df.iterrows():
        if fill:
            pdf.set_fill_color(242, 245, 249)
        else:
            pdf.set_fill_color(255, 255, 255)
            
        pdf.set_text_color(0, 0, 0)
        var_val = row['Variação (Δ%)']
        tendencia = str(row['Tendência'])
        
        ult_preco_val = row['Último Preço Hist. (R$)']
        ult_preco_str = formatar_brl(ult_preco_val) if ult_preco_val != "" else ""
        forn_ant_str = str(row['Fornecedor do Último Preço']) if ult_preco_val != "" else ""
        
        preco_med_val = row['Preço Médio (R$)']
        preco_med_str = formatar_brl(preco_med_val) if preco_med_val != "" else ""
        
        var_str = formatar_pct(var_val) if var_val != "" else ""
        
        pdf.cell(col_widths[0], 6, limpar_texto_pdf(str(row['Item'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[1], 6, limpar_texto_pdf(str(row['Código'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[2], 6, limpar_texto_pdf(str(row['Descrição Resumida'])[:35]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[3], 6, limpar_texto_pdf(str(row['Qtd'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[4], 6, limpar_texto_pdf(formatar_brl(row['Novo Preço Unit. (R$)'])), border=1, fill=fill, align="R")
        pdf.cell(col_widths[5], 6, limpar_texto_pdf(str(row['Fornecedor do Preço Novo'])[:22]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[6], 6, limpar_texto_pdf(ult_preco_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[7], 6, limpar_texto_pdf(forn_ant_str[:22]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[8], 6, limpar_texto_pdf(preco_med_str), border=1, fill=fill, align="R")
        
        if var_val != "":
            if var_val < 0:
                pdf.set_text_color(44, 160, 44)
            elif var_val > 0:
                pdf.set_text_color(192, 0, 0)
            else:
                pdf.set_text_color(0, 0, 0)
        
        pdf.cell(col_widths[9], 6, limpar_texto_pdf(var_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[10], 6, limpar_texto_pdf(tendencia), border=1, fill=fill, align="C")
        
        pdf.ln()
        fill = not fill
        
    pdf_output = pdf.output(dest='S')
    if isinstance(pdf_output, str):
        return pdf_output.encode('latin1')
    return bytes(pdf_output)

cotacao = pd.DataFrame()
if uploaded_cot is not None:
    bar = st.progress(0)
    st.text("Processando dados...")
    for i in range(100):
        time.sleep(0.01)
        bar.progress(i + 1)
        
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
    bar.empty()

df_final = pd.DataFrame()

if not cotacao.empty:
    cotacao.columns = [str(c).strip() for c in cotacao.columns]

    def achar_coluna(df, termos):
        for col in df.columns:
            c_low = str(col).lower()
            c_low_norm = "".join([c for c in unicodedata.normalize('NFKD', c_low) if not unicodedata.combining(c)])
            for t in termos:
                t_norm = "".join([c for c in unicodedata.normalize('NFKD', t) if not unicodedata.combining(c)])
                if t_norm in c_low_norm:
                    return col
        return None

    c_item = achar_coluna(cotacao, ['item'])
    c_cod = achar_coluna(cotacao, ['código', 'codigo', 'produto', 'sku'])
    c_desc = achar_coluna(cotacao, ['descrição', 'descricao'])
    c_qtd = achar_coluna(cotacao, ['qtd', 'quantidade'])
    c_vlr = achar_coluna(cotacao, ['valor unit', 'vlr. unit', 'unitario', 'preço unit', 'preco unit', 'vlr', 'preço', 'preco'])
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

        ultimo_preco = 0.0
        preco_medio = 0.0
        forn_hist = ""
        tem_historico = False
        
        if not historico.empty:
            match_linhas = []
            precos_encontrados = []
            for h_idx, h_row in historico.iterrows():
                for col_idx in h_row.index:
                    val_celula = str(h_row[col_idx])
                    if padronizar_codigo_10_digitos(val_celula) == codigo_busca and codigo_busca != '0000000000':
                        match_linhas.append(h_row)
                        try:
                            p_val = limpar_valor(h_row.get(10, 0))
                            if p_val > 0:
                                precos_encontrados.append(p_val)
                            else:
                                for val in h_row.values:
                                    v = limpar_valor(val)
                                    if v > 1.0 and v != qtd:
                                        precos_encontrados.append(v)
                                        break
                        except:
                            pass
                        break
                        
            if match_linhas:
                tem_historico = True
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
                
                if precos_encontrados:
                    preco_medio = sum(precos_encontrados) / len(precos_encontrados)
                elif ultimo_preco > 0:
                    preco_medio = ultimo_preco
                    
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
            
        if tem_historico and preco_medio > 0:
            variacao = ((preco_novo - preco_medio) / preco_medio) * 100
            if variacao < 0:
                tendencia = "Queda"
            elif variacao > 0:
                tendencia = "Alta"
            else:
                tendencia = "Estabilidade"
        else:
            ultimo_preco = ""
            preco_medio = ""
            forn_hist = ""
            variacao = ""
            tendencia = "Sem Histórico"
            
        resultados.append({
            'Item': num_item,
            'Código': codigo_original,
            'Descrição Resumida': desc,
            'Qtd': qtd,
            'Novo Preço Unit. (R$)': preco_novo,
            'Fornecedor do Preço Novo': forn_novo,
            'Último Preço Hist. (R$)': ultimo_preco,
            'Fornecedor do Último Preço': forn_hist,
            'Preço Médio (R$)': preco_medio,
            'Variação (Δ%)': variacao,
            'Tendência': tendencia
        })

    df_final = pd.DataFrame(resultados)

if cotacao.empty:
    st.info("👆 Clique na caixa **⚙️ Abrir / Fechar Configurações** acima e faça o upload do seu Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml) para iniciar a análise.")
elif df_final.empty:
    st.warning("⚠️ Nenhum item válido encontrado. Verifique o arquivo carregado.")
else:
    colunas_exatas = [
        'Item', 'Código', 'Descrição Resumida', 'Qtd', 
        'Novo Preço Unit. (R$)', 'Fornecedor do Preço Novo',
        'Último Preço Hist. (R$)', 'Fornecedor do Último Preço',
        'Preço Médio (R$)', 'Variação (Δ%)', 'Tendência'
    ]

    df_final = df_final[colunas_exatas]

    df_display = df_final.copy()
    df_display['Qtd'] = df_display['Qtd'].apply(formatar_qtd)
    df_display['Novo Preço Unit. (R$)'] = df_display['Novo Preço Unit. (R$)'].apply(formatar_brl)
    df_display['Último Preço Hist. (R$)'] = df_display['Último Preço Hist. (R$)'].apply(formatar_brl)
    df_display['Preço Médio (R$)'] = df_display['Preço Médio (R$)'].apply(formatar_brl)
    df_display['Variação (Δ%)'] = df_final['Variação (Δ%)'].apply(formatar_pct_com_seta)

    # Exibição do painel interativo formatado com classe CSS 'dataframe' (Estilo Dados Bancários)
    st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

    html_tabela = df_display.to_html(escape=False, index=False, classes='dataframe')
    st.markdown(html_tabela, unsafe_allow_html=True)

    # Preenche o botão de download de PDF dentro da caixa superior aberta
    pdf_bytes = gerar_pdf(df_final)
    placeholder_pdf.download_button(
        label="📥 Baixar Mapa em PDF",
        data=pdf_bytes,
        file_name="mapa_de_cotacao_suprimentos.pdf",
        mime="application/pdf",
        key="btn_pdf_top"
    )

    # Bloco de Pesquisa por Código do Item posicionado abaixo do mapa com Gráfico ajustado
    st.markdown("---")
    st.subheader("🔍 Consulta de Histórico por Código do Item")
    
    col_search1, col_search2 = st.columns([3, 7])
    with col_search1:
        codigo_pesquisa = st.text_input("Digite ou cole o código do item (10 dígitos ou parcial):", "")

    if codigo_pesquisa:
        cod_limpo = padronizar_codigo_10_digitos(codigo_pesquisa)
        compras_encontradas = []
        dados_grafico = []
        
        if not historico.empty:
            for h_idx, h_row in historico.iterrows():
                encontrou = False
                for col_idx in h_row.index:
                    val_celula = str(h_row[col_idx])
                    if cod_limpo in padronizar_codigo_10_digitos(val_celula) and codigo_busca != '0000000000':
                        encontrou = True
                        break
                
                if encontrou:
                    forn_val = "Não identificado"
                    preco_val = 0.0
                    
                    # Coluna C (Número do Pedido) = Índice 2
                    num_pedido = str(h_row.get(2, "N/D"))
                    
                    # Coluna M (Data Emissão) = Índice 12
                    data_raw = str(h_row.get(12, ""))
                    try:
                        data_dt = pd.to_datetime(data_raw, dayfirst=True)
                        data_val = data_dt.strftime('%d/%m/%Y')
                    except:
                        data_val = data_raw
                    
                    try:
                        f_col = str(h_row.get(4, ""))
                        if len(f_col) > 3 and f_col.lower() not in ['nan', 'a vista', '25 dias']:
                            forn_val = f_col
                        else:
                            for val in h_row.values:
                                t = str(val)
                                if len(t) > 5 and not any(char.isdigit() for char in t[:3]) and t.lower() not in ['a vista', '25 dias', '30 dias']:
                                    forn_val = t
                                    break
                    except:
                        pass
                        
                    # Coluna K (Preço Unitário) = Índice 10
                    try:
                        p_col = limpar_valor(h_row.get(10, 0))
                        if p_col > 0:
                            preco_val = p_col
                    except:
                        pass
                        
                    compras_encontradas.append({
                        'Nº do Pedido': num_pedido,
                        'Data Emissao PC': data_val,
                        'Fornecedor da Compra': forn_val,
                        'Prc Unitario': formatar_brl(preco_val) if preco_val > 0 else "R$ 0,00"
                    })
                    
                    if preco_val > 0 and pd.notna(pd.to_datetime(data_raw, dayfirst=True, errors='coerce')):
                        dados_grafico.append({"Data Emissao": pd.to_datetime(data_raw, dayfirst=True), "Prc Unitario": preco_val})

        if compras_encontradas:
            st.success(f"Foram encontradas **{len(compras_encontradas)}** ocorrência(s) de compra para o código pesquisado:")
            df_historico_item = pd.DataFrame(compras_encontradas)
            st.table(df_historico_item)

            # GRÁFICO DE LINHA NATIVO ORDENADO CRONOLOGICAMENTE
            if dados_grafico:
                st.markdown("#### 📈 Evolução do Preço Histórico")
                df_chart = pd.DataFrame(dados_grafico).sort_values("Data Emissao")
                
                df_chart_display = df_chart.copy()
                df_chart_display["Data Emissao"] = df_chart_display["Data Emissao"].dt.strftime('%d/%m/%Y')
                
                st.line_chart(df_chart_display.set_index("Data Emissao")["Prc Unitario"])
        else:
            st.warning("⚠️ Nenhuma compra anterior encontrada no histórico para este código.")
