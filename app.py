import streamlit as st
import pandas as pd
import numpy as np
import os
import docx
from fpdf import FPDF
import unicodedata
from google import genai

# Configuração da Página
st.set_page_config(
    page_title="Mapa de Cotação e Análise de Custos",
    page_icon="📊",
    layout="wide"
)

# Chave do Gemini pré-configurada
GEMINI_API_KEY = "AQ.Ab8RN6I0XGpR9RtaXtPUpM1gvgYbVRMTtwkRDXYvMTKxfDEgtQ"

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

def extrair_numeros(codigo):
    if pd.isna(codigo):
        return ""
    apenas_nums = ''.join(filter(str.isdigit, str(codigo)))
    return str(int(apenas_nums)) if apenas_nums.isdigit() else str(codigo).strip()

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

# 2. Leitura inteligente e limpa de arquivos DOCX do TOTVS
def extrair_tabela_docx_limpa(arquivo_docx):
    try:
        doc = docx.Document(arquivo_docx)
        linhas_validas = []
        
        for tabela in doc.tables:
            for linha in tabela.rows:
                celulas = [cel.text.strip().replace('\n', ' ') for cel in linha.cells]
                # Filtra apenas linhas que contêm conteúdo real e descarta cabeçalhos institucionais
                texto_linha_unido = " ".join(celulas).lower()
                if any(celulas) and not any(ignorar in texto_linha_unido for ignorar in ['parente andrade', 'departamento de suprimentos', 'mapa de cotacao']):
                    # Verifica se a linha tem colunas suficientes para ser um item de produto
                    if len(celulas) >= 3:
                        linhas_validas.append(celulas)
                        
        if len(linhas_validas) > 0:
            # Tenta achar a linha de cabeçalho real dos itens
            inicio_dados = 0
            for idx, l in enumerate(linhas_validas[:5]):
                unido = " ".join(l).lower()
                if 'código' in unido or 'codigo' in unido or 'descrição' in unido or 'unitário' in unido:
                    inicio_dados = idx + 1
                    break
                    
            dados = linhas_validas[inicio_dados:] if inicio_dados < len(linhas_validas) else linhas_validas
            if not dados:
                dados = linhas_validas
                
            max_cols = max(len(l) for l in dados)
            headers = [f"Col_{i}" for i in range(max_cols)]
            
            # Normaliza o tamanho das linhas
            dados_norm = [l + [''] * (max_cols - len(l)) for l in dados]
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
            cotacao = pd.read_csv(uploaded_cot)
        elif nome.endswith(('.xlsx', '.xls')):
            cotacao = pd.read_excel(uploaded_cot)
        elif nome.endswith('.docx'):
            cotacao = extrair_tabela_docx_limpa(uploaded_cot)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")
        cotacao = pd.DataFrame()

# Se nenhum arquivo foi carregado ou a tabela estiver vazia, exibe instrução inicial
if cotacao.empty:
    st.info("👈 Por favor, faça o upload do seu Mapa de Cotação Atual (.csv, .xlsx ou .docx) na barra lateral para iniciar a análise.")
else:
    cotacao.columns = [str(c).strip() for c in cotacao.columns]

    resultados = []
    item_contador = 1

    for idx, row in cotacao.iterrows():
        # Varre as células da linha para identificar dinamicamente: Código, Descrição, Qtd, Preço Novo e Fornecedor
        celulas_str = [str(val).strip() for val in row.values if pd.notna(val) and str(val).strip() != '']
        
        if not celulas_str:
            continue
            
        # Tenta identificar código (geralmente sequências numéricas longas ou SKUs)
        codigo_original = ""
        desc = "Descrição não informada"
        qtd = 1.0
        preco_novo = 0.0
        forn_novo = "Fornecedor não informado"
        
        candidatos_numericos = []
        candidatos_texto = []
        candidatos_monetarios = []
        
        for val in celulas_str:
            v_limpo = limpar_valor(val)
            if v_limpo > 0 and ('r$' in val.lower() or ',' in val or '.' in val):
                candidatos_monetarios.append(v_limpo)
            elif val.isdigit() and len(val) >= 4:
                codigo_original = val
            elif len(val) > 4 and not val.isdigit():
                candidatos_texto.append(val)
                
        if not codigo_original and celulas_str:
            codigo_original = celulas_str[0]
            
        codigo_busca = extrair_numeros(codigo_original)
        
        if candidatos_texto:
            desc = candidatos_texto[0]
            if len(candidatos_texto) > 1:
                forn_novo = candidatos_texto[-1] # Pega o último texto longo como fornecedor provável
                
        if candidatos_monetarios:
            preco_novo = candidatos_monetarios[-1] # O último valor monetário costuma ser o preço unitário ou total
            
        # Ignora linhas que sejam cabeçalhos repetidos
        if codigo_original.lower() in ['código', 'codigo', 'item', 'produto'] or len(codigo_original) == 0:
            continue

        num_item = f"{item_contador:04d}"
        item_contador += 1

        # Busca rigorosa da última ocorrência do código no histórico
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

    if df_final.empty:
        st.warning("⚠️ Nenhum item de produto válido foi encontrado no arquivo carregado. Verifique se o arquivo está no formato correto.")
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

        # Seção de Painel Analítico & Estratégico com Gemini AI
        st.markdown("---")
        st.subheader("🤖 Painel Analítico & Executivo (Gerado por Gemini AI)")
        
        if st.button("✨ Executar Análise Comparativa e Montar Painel Executivo"):
            with st.spinner("A Inteligência Artificial está processando as comparações de preços, fornecedores e montando o painel executivo..."):
                try:
                    client = genai.Client(api_key=GEMINI_API_KEY)
                    
                    resumo_dados = df_final[['Código', 'Descrição Resumida', 'Qtd', 'Último Preço Hist. (R$)', 'Fornecedor do Último Preço', 'Novo Preço Unit. (R$)', 'Fornecedor do Preço Novo', 'Variação (Δ%)', 'Tendência']].to_string()
                    
                    prompt = (
                        "Atue como um Especialista Sênior em Supply Chain, Gestão de Compras e Negociação Corporativa. "
                        "Com base no mapa de cotação atual comparado estritamente ao histórico de compras abaixo:\n\n"
                        f"{resumo_dados}\n\n"
                        "Monte um **Painel Comparativo e Executivo** formatado em blocos claros com as seguintes seções:\n"
                        "1. **Resumo Executivo do Cenário:** Visão geral da cotação frente ao histórico consolidado.\n"
                        "2. **Quadro de Oportunidades (Ganhos de Margem / Reduções):** Destaque dos itens com variação favorável (queda) e os fornecedores envolvidos.\n"
                        "3. **Matriz de Alertas e Riscos (Aumentos de Custo):** Análise crítica dos itens que apresentaram alta e diretrizes para renegociação urgente.\n"
                        "4. **Recomendações Estratégicas Finais para Homologação:** Considerações logísticas, fiscais (ZFM) e fechamento de ordens de compra."
                    )
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                    )
                    
                    st.success("Painel Comparativo Estratégico gerado com sucesso!")
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"Erro ao comunicar com a API do Gemini: {e}")

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
