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

# Estilização visual corporativa estilo Dados Bancários e layout limpo sem barra lateral
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    h1 { color: #1f2c34; font-family: 'Helvetica Neue', sans-serif; margin-bottom: 5px; }
    
    /* Cabeçalho alinhado com status à direita */
    .top-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
    }
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
    .dataframe td:nth-child(9), .dataframe th:nth-child(9) {
        white-space: nowrap !important;
        text-align: center !important;
    }
    .dataframe tr:nth-child(even) {
        background-color: #f2f5f9 !important;
    }
    .dataframe tr:nth-child(odd) {
        background-color: #ffffff !important;
    }
    .dataframe td:nth-child(1), .dataframe td:nth-child(2), .dataframe td:nth-child(3), .dataframe td:nth-child(6), .dataframe td:nth-child(8), .dataframe td:nth-child(10) {
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
            4: ['CAA COMERCIO AMAZONENSE DE ALUMINIO LTDA.'], 
            6: ['0000005177'],                              
            10: ['375.58']                                   
        })
        return df, "Base de dados atualizada em: Indisponível"

historico, status_historico = carregar_historico_github()

# Topo do App: Título e Status no canto superior direito
col_title, col_status = st.columns([7, 3])
with col_title:
    st.title("📊 Gestão Estratégica de Compras | Mapa de Cotação")
with col_status:
    st.markdown(f"<div style='text-align: right; margin-top: 15px;'><span class='status-badge'>ℹ️ {status_historico}</span></div>", unsafe_allow_html=True)

st.markdown("---")

# Caixa Oculta (Abre / Fecha) para Upload, Data Base e Temas
with st.expander("⚙️ Abrir / Fechar Configurações (Upload, Data Base e Tema)", expanded=False):
    col_exp1, col_exp2 = st.columns([2, 1])
    
    with col_exp1:
        st.markdown("### 📁 Upload de Arquivo")
        uploaded_cot = st.file_uploader(
            "Carregar Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml)", 
            type=["csv", "xlsx", "docx", "mhtml", "html"]
        )
        
    with col_exp2:
        st.markdown("### 🎨 Tema do Sistema")
        tema_selecionado = st.selectbox(
            "Selecione o Estilo Visual",
            ["Dados Bancários (Padrão)", "Corporativo Azul Limpo", "Modo Compacto"]
        )
        st.info(f"Tema ativo: **{tema_selecionado}**")

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
        return ""
    return f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_pct_com_seta(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except:
        return ""
    
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

cotacao = pd.DataFrame()
if uploaded_cot is not None:
    with st.spinner("Loading... Processando mapa de cotação e cruzando com o histórico..."):
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

if cotacao.empty:
    st.info("👆 Abra a caixa **⚙️ Abrir / Fechar Configurações** acima e faça o upload do seu Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml) para iniciar a análise.")
else:
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
        forn_hist = ""
        tem_historico = False
        
        if not historico.empty:
            match_linhas = []
            for h_idx, h_row in historico.iterrows():
                for col_idx in h_row.index:
                    val_celula = str(h_row[col_idx])
                    if padronizar_codigo_10_digitos(val_celula) == codigo_busca and codigo_busca != '0000000000':
                        match_linhas.append(h_row)
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
            
        if tem_historico and ultimo_preco > 0:
            variacao = ((preco_novo - ultimo_preco) / ultimo_preco) * 100
            if variacao < 0:
                tendencia = "Queda"
            elif variacao > 0:
                tendencia = "Alta"
            else:
                tendencia = "Estabilidade"
        else:
            ultimo_preco = ""
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
            'Variação (Δ%)': variacao,
            'Tendência': tendencia
        })

    df_final = pd.DataFrame(resultados)

    if df_final.empty:
        st.warning("⚠️ Nenhum item válido encontrado. Verifique o arquivo carregado.")
    else:
        colunas_exatas = [
            'Item', 'Código', 'Descrição Resumida', 'Qtd', 
            'Novo Preço Unit. (R$)', 'Fornecedor do Preço Novo',
            'Último Preço Hist. (R$)', 'Fornecedor do Último Preço',  
            'Variação (Δ%)', 'Tendência'
        ]

        df_final = df_final[colunas_exatas]

        df_display = df_final.copy()
        df_display['Qtd'] = df_display['Qtd'].apply(formatar_qtd)
        df_display['Novo Preço Unit. (R$)'] = df_display['Novo Preço Unit. (R$)'].apply(formatar_brl)
        df_display['Último Preço Hist. (R$)'] = df_display['Último Preço Hist. (R$)'].apply(formatar_brl)
        df_display['Variação (Δ%)'] = df_final['Variação (Δ%)'].apply(formatar_pct_com_seta)

        # Exibição do painel interativo formatado com classe CSS 'dataframe' (Estilo Dados Bancários)
        st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

        html_tabela = df_display.to_html(escape=False, index=False, classes='dataframe')
        st.markdown(html_tabela, unsafe_allow_html=True)

        # Configuração da Classe do PDF com Margens Estreitas (Esq/Dir: 0.64cm = 6.4mm, Sup/Inf: 1.91cm = 19.1mm)
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

        # Função para Gerar PDF do Relatório com colunas somando exatamente 284mm
        def gerar_pdf(df):
            pdf = PDFProfissional()
            pdf.add_page()
            
            col_widths = [10, 22, 60, 10, 20, 50, 20, 50, 22, 20]
            headers = [
                "Item", "Codigo", "Descricao", "Qtd", 
                "Novo Preco", "Forn. Novo", "Ult. Preco", 
                "Forn. Ant.", "Var(%)", "Tendencia"
            ]
            
            pdf.set_fill_color(47, 85, 151)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("helvetica", "B", 8)
            
            for i, h in enumerate(headers):
                pdf.cell(col_widths[i], 7, limpar_texto_pdf(h), border=1, fill=True, align="C")
            pdf.ln()
            
            pdf.set_font("helvetica", "", 7)
            
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
                var_str = formatar_pct(var_val) if var_val != "" else ""
                
                pdf.cell(col_widths[0], 6, limpar_texto_pdf(str(row['Item'])), border=1, fill=fill, align="C")
                pdf.cell(col_widths[1], 6, limpar_texto_pdf(str(row['Código'])), border=1, fill=fill, align="C")
                pdf.cell(col_widths[2], 6, limpar_texto_pdf(str(row['Descrição Resumida'])[:40]), border=1, fill=fill, align="L")
                pdf.cell(col_widths[3], 6, limpar_texto_pdf(str(row['Qtd'])), border=1, fill=fill, align="C")
                pdf.cell(col_widths[4], 6, limpar_texto_pdf(formatar_brl(row['Novo Preço Unit. (R$)'])), border=1, fill=fill, align="R")
                pdf.cell(col_widths[5], 6, limpar_texto_pdf(str(row['Fornecedor do Preço Novo'])[:28]), border=1, fill=fill, align="L")
                pdf.cell(col_widths[6], 6, limpar_texto_pdf(ult_preco_str), border=1, fill=fill, align="R")
                pdf.cell(col_widths[7], 6, limpar_texto_pdf(forn_ant_str[:28]), border=1, fill=fill, align="L")
                
                if var_val != "":
                    if var_val < 0:
                        pdf.set_text_color(44, 160, 44)
                    elif var_val > 0:
                        pdf.set_text_color(192, 0, 0)
                    else:
                        pdf.set_text_color(0, 0, 0)
                
                pdf.cell(col_widths[8], 6, limpar_texto_pdf(var_str), border=1, fill=fill, align="R")
                pdf.cell(col_widths[9], 6, limpar_texto_pdf(tendencia), border=1, fill=fill, align="C")
                
                pdf.ln()
                fill = not fill
                
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

        # Seção Obrigatória: Insight Rápido do Especialista
        st.markdown("---")
        st.subheader("💡 Insight Rápido do Especialista")

        itens_em_queda = df_final[(df_final['Variação (Δ%)'] != "") & (pd.to_numeric(df_final['Variação (Δ%)']) < 0)]
        if not itens_em_queda.empty:
            insight_texto = (
                f"Identificada oportunidade expressiva de economia em **{len(itens_em_queda)} item(ns)** com redução de custos favorável. "
                "Recomenda-se a **homologação imediata com o novo fornecedor** para captura dos ganhos de margem, "
                "certificando-se de que os prazos de entrega e condições logísticas para Manaus atendem ao cronograma operacional."
            )
        else:
            insight_texto = (
                "Cenário de alta nos preços detectado ou itens sem histórico prévio. Recomenda-se a **renegociação com base no histórico de volume** "
                "ou busca por fornecedores alternativos na praça local para evitar impacto no orçamento."
            )

        st.success(insight_texto)
