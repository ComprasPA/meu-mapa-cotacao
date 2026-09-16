import streamlit as st
import pandas as pd
import numpy as np
import os
import base64
import docx
import unicodedata
import email
from bs4 import BeautifulSoup
import datetime
import time
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import gspread
from google.oauth2.service_account import Credentials

from logica_cotacao import (
    _achar_coluna_normalizada,
    normalizar_codigo,
    consolidar_base_precos,
    processar_mapa_cotacao,
    extrair_numero_cotacao,
    formatar_brl,
    gerar_pdf,
    gerar_excel,
)

# ==============================================================================
# Configuração da Página
# ==============================================================================
st.set_page_config(
    page_title="Gestão Estratégica de Compras | Mapa de Cotação",
    page_icon="📊",
    layout="wide"
)

# ==============================================================================
# Tema visual (Claro / Escuro) — selecionável em ⚙️ Configurações.
# A escolha fica em st.session_state['tema'] e é lida no topo do script a
# cada execução, então a troca feita no seletor (mais abaixo, dentro do
# expander) já se aplica no mesmo rerun.
# ==============================================================================
if 'tema' not in st.session_state:
    st.session_state['tema'] = 'Claro'

TEMAS = {
    'Claro': dict(
        main_bg='#ffffff', text='#1f2c34', body_text='#000000',
        badge_bg='#e8f0fe', badge_text='#1967d2', badge_border='#d2e3fc',
        expander_border='#d9d9d9', expander_bg='#ffffff', expander_header_bg='#f8f9fa',
        footer_bg='#f8f9fa', footer_border='#2f5597', footer_text='#1f2c34',
        th_bg='#2f5597', th_text='#ffffff', th_border='#b4c6e7',
        td_border='#d9d9d9', td_text='#000000', row_bg='#ffffff',
        # Cores da marca Parente Andrade — mesma paleta do Portal Gestão de
        # Compras (comum.py: TEMAS['claro']), usadas no cabeçalho e nas
        # caixas/botões (input_bg, mist, verde_soft).
        verde='#3E8E41', verde_deep='#2E6B31', verde_soft='#E7F3E6',
        laranja='#F2861D', laranja_deep='#CE6E10',
        input_bg='#F1F2EE', mist='#E4E7E0',
    ),
    'Escuro': dict(
        main_bg='#0e1117', text='#f5f5f5', body_text='#e6e6e6',
        badge_bg='#1c2e4a', badge_text='#8ab4f8', badge_border='#2c4770',
        expander_border='#333a45', expander_bg='#161a20', expander_header_bg='#1c212a',
        footer_bg='#161a20', footer_border='#4f7cff', footer_text='#f5f5f5',
        th_bg='#1f2a3f', th_text='#e8eef7', th_border='#3a4a66',
        td_border='#2a2f38', td_text='#e6e6e6', row_bg='#161a20',
        verde='#4FA653', verde_deep='#3E8E41', verde_soft='#1F3B22',
        laranja='#F2951D', laranja_deep='#CE6E10',
        input_bg='#2A342E', mist='#3B4A42',
    ),
}


def gerar_css(tema: str) -> str:
    t = TEMAS.get(tema, TEMAS['Claro'])
    return f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Public+Sans:wght@400;500;600;700&display=swap');
    div.st-key-header_card {{
        background: {t['expander_bg']};
        padding: 16px 28px;
        border-radius: 14px;
        margin-bottom: 16px;
        box-shadow: 0 1px 2px rgba(28,36,32,.04), 0 10px 28px -14px rgba(28,36,32,.14);
        position: relative;
        overflow: hidden;
    }}
    div.st-key-header_card::before {{
        content: "";
        position: absolute;
        left: 0; top: 0; bottom: 0;
        width: 6px;
        background: linear-gradient(180deg, {t['verde']}, {t['laranja']});
    }}
    div.st-key-header_card > div {{ align-items: center; }}
    .brand-text-block {{ display: flex; flex-direction: column; align-items: center; line-height: 1.35; }}
    .brand-eyebrow {{
        font-family: 'Public Sans', sans-serif; font-weight: 700; font-size: 15.75px;
        letter-spacing: .12em; text-transform: uppercase; color: {t['laranja_deep']} !important;
        margin: 0; display: block; white-space: nowrap;
    }}
    .brand-subtitle {{
        font-family: 'Sora', sans-serif; font-weight: 700; font-size: 18px;
        color: {t['text']} !important; display: block; white-space: nowrap;
    }}
    .main {{ background-color: {t['main_bg']}; }}
    [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
        background-color: {t['main_bg']} !important;
    }}
    body, p, span, label, .stMarkdown {{ color: {t['body_text']} !important; }}
    h1 {{ color: {t['text']} !important; font-family: 'Helvetica Neue', sans-serif; margin-bottom: 5px; }}
    h2, h3, h4, h5, h6 {{ color: {t['text']} !important; }}

    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 150px !important;
        max-width: 100% !important;
    }}
    div[data-testid="stMarkdownContainer"] hr {{
        margin: 4px 0 !important;
    }}
    header {{ visibility: hidden !important; }}
    #MainMenu {{ visibility: hidden !important; }}
    footer {{ visibility: hidden !important; }}
    div[data-baseweb="modal"], div.stDialog, div[role="dialog"] {{
        display: none !important;
    }}
    button[data-testid^="stBaseButton"] {{
        background-color: {t['expander_header_bg']} !important;
        border: 1px solid {t['expander_border']} !important;
        border-radius: 7px !important;
        color: {t['text']} !important;
        font-weight: 600 !important;
        box-shadow: 0 1px 2px rgba(28,36,32,.08) !important;
    }}
    button[data-testid^="stBaseButton"]:hover {{
        border-color: {t['verde']} !important;
        color: {t['verde']} !important;
    }}
    button[data-testid^="stBaseButton"]:disabled {{
        opacity: 0.45 !important;
        box-shadow: none !important;
    }}
    div[data-testid="stTextInput"] input {{
        background-color: {t['input_bg']} !important;
        border: none !important;
        border-radius: 9px !important;
        box-shadow: none !important;
        color: {t['text']} !important;
        transition: background-color 0.2s;
    }}
    div[data-testid="stTextInput"] input:focus {{
        background-color: {t['verde_soft']} !important;
    }}
    .status-badge {{
        background-color: {t['badge_bg']};
        color: {t['badge_text']};
        padding: 6px 12px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 600;
        font-family: 'Helvetica Neue', sans-serif;
        border: 1px solid {t['badge_border']};
    }}
    div[data-testid="stExpander"] {{
        border: 1px solid {t['mist']} !important;
        background-color: {t['expander_bg']} !important;
        border-radius: 16px !important;
        box-shadow: 0 1px 2px rgba(28,36,32,.04), 0 10px 28px -14px rgba(28,36,32,.14) !important;
        margin-bottom: 4px !important;
    }}
    .streamlit-expanderHeader,
    div[data-testid="stExpander"] summary {{
        padding-top: 8px !important;
        padding-bottom: 8px !important;
        min-height: 40px !important;
        font-size: 14px !important;
        background-color: {t['expander_header_bg']} !important;
        border-radius: 16px 16px 0 0 !important;
        color: {t['text']} !important;
    }}
    .streamlit-expanderContent,
    div[data-testid="stExpander"] summary ~ div {{
        padding: 15px !important;
        background-color: {t['expander_bg']} !important;
    }}
    .footer-pesquisa {{
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: {t['footer_bg']};
        border-top: 2px solid {t['footer_border']};
        padding: 8px 20px;
        z-index: 99999;
        box-shadow: 0px -4px 10px rgba(0, 0, 0, 0.1);
        color: {t['footer_text']};
    }}

    div[data-testid="stTextInput"] {{
        max-width: 400px !important;
    }}
    .dataframe {{
        width: 100% !important;
        table-layout: auto !important;
        border-collapse: collapse !important;
        font-family: 'Helvetica Neue', sans-serif !important;
        font-size: 12px !important;
    }}
    .dataframe th {{
        background-color: {t['th_bg']} !important;
        color: {t['th_text']} !important;
        text-align: center !important;
        font-weight: bold !important;
        padding: 8px 6px !important;
        border: 1px solid {t['th_border']} !important;
        font-size: 12px !important;
        white-space: nowrap !important;
    }}
    .dataframe td {{
        padding: 7px 6px !important;
        border: 1px solid {t['td_border']} !important;
        color: {t['td_text']} !important;
        font-size: 12px !important;
        text-align: right;
    }}
    .dataframe tr:nth-child(even),
    .dataframe tr:nth-child(odd) {{
        background-color: {t['row_bg']} !important;
    }}
    .dataframe td:nth-child(1), .dataframe th:nth-child(1),
    .dataframe td:nth-child(2), .dataframe th:nth-child(2) {{
        white-space: nowrap !important;
        text-align: center !important;
    }}
    .dataframe td:nth-child(3), .dataframe td:nth-child(6), .dataframe td:nth-child(8) {{
        text-align: left;
    }}
    .dataframe td:nth-child(11), .dataframe th:nth-child(11) {{
        text-align: center !important;
        white-space: nowrap !important;
    }}
    </style>
    """


@st.cache_data(ttl=86400)
def get_base64_logo(tema: str):
    # No tema Escuro usa a variante com o texto "PARENTE ANDRADE" em
    # branco (logo_dark.png) — o "PA" verde/laranja é o mesmo nas duas.
    nome_arquivo = "logo_dark.png" if tema == "Escuro" else "logo.png"
    caminho_logo = os.path.join(os.path.dirname(os.path.abspath(__file__)), nome_arquivo)
    try:
        with open(caminho_logo, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        return None


st.markdown(gerar_css(st.session_state['tema']), unsafe_allow_html=True)

# ==============================================================================
# 1. BASE HISTÓRICA — mesma planilha do Portal Gestão de Compras
#
# Antes, o histórico vinha de um historico_compras.csv versionado no
# repositório (precisava ser atualizado manualmente). Agora lê ao vivo a
# mesma aba "Pedidos" da planilha Google Sheets usada pelo Portal Gestão de
# Compras (consulta-parente-andrade/comum.py: obter_client_gspread,
# carregar_dados_seguros) — mesma conta de serviço, mesmo FILE_ID. Cada
# linha do Pedidos já É uma compra histórica; não há upload/edição pelo
# painel — para corrigir um dado, corrija na planilha.
# ==============================================================================

FILE_ID = "1e7pQ512ge5XMnXxsRODEO7V48KgWo6FpKeITFqBSg1o"

HIST_REQUIRED_COLS = ['Produto', 'Descricao.1', 'Unidade', 'Prc Unitario',
                       'Data Emissao', 'Nome Fornece', 'Quantidade', 'Status Aprov']

# Mapeia o nome interno (usado em todo o resto do arquivo, herdado do antigo
# export TOTVS) para os termos aceitos no cabeçalho real da aba "Pedidos"
# (comparados sem acento/caixa — ver _achar_coluna_normalizada).
MAPA_COLUNAS_PEDIDOS = {
    'Produto': ['produto'],
    'Descricao.1': ['descricao'],
    'Unidade': ['um'],
    'Prc Unitario': ['preco unitario'],
    'Data Emissao': ['data pedido'],
    'Nome Fornece': ['fornecedor'],
    'Quantidade': ['qtd'],
    'Status Aprov': ['status'],
    'Pedido': ['pedido'],
}


def obter_client_gspread():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scope)
    return gspread.authorize(creds)


def _ler_historico_bruto() -> pd.DataFrame:
    """Lê a aba "Pedidos" da mesma planilha Google Sheets usada pelo Portal
    Gestão de Compras (todas as colunas como texto — igual a
    carregar_dados_seguros() em consulta-parente-andrade/main.py) e renomeia
    para os nomes internos usados no resto deste arquivo.
    """
    client = obter_client_gspread()
    spreadsheet = client.open_by_key(FILE_ID)
    try:
        worksheet = spreadsheet.worksheet("Pedidos")
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.get_worksheet(0)

    dados = worksheet.get_all_values()
    if not dados:
        return pd.DataFrame(columns=HIST_REQUIRED_COLS)

    cabecalho = [str(c).strip() for c in dados[0]]
    linhas = dados[1:]
    linhas_normalizadas = []
    for linha in linhas:
        linha = list(linha) + [""] * (len(cabecalho) - len(linha))
        linhas_normalizadas.append(linha[:len(cabecalho)])

    df = pd.DataFrame(linhas_normalizadas, columns=cabecalho, dtype=str).fillna('')

    # Pedido marcado "EXCLUÍDO DO TOTVS" (sumiu do relatorio do Totvs, ver
    # detectar_pedidos_excluidos_import no Portal Gestão de Compras) fica na
    # planilha mas nao deve entrar em nenhum calculo/consulta daqui - decisao
    # explicita do usuario, mesmo com a regra "HISTÓRICO É HISTÓRICO" abaixo
    # (aquela regra e sobre Status Aprov normal, essa exclusao e diferente:
    # o pedido pode nem ter existido de verdade no Totvs).
    col_status_bruto = next((c for c in df.columns if c.upper().strip() == "STATUS"), None)
    if col_status_bruto:
        df = df[df[col_status_bruto].astype(str).str.strip().str.upper() != "EXCLUÍDO DO TOTVS"].reset_index(drop=True)

    renomeio = {}
    for nome_interno, termos in MAPA_COLUNAS_PEDIDOS.items():
        col_real = _achar_coluna_normalizada(df.columns, termos)
        if col_real:
            renomeio[col_real] = nome_interno
    return df.rename(columns=renomeio)


@st.cache_data(ttl=60, show_spinner="Consolidando base histórica de preços...")
def construir_base_precos(status_filtro: str = None):
    """
    Constrói a base histórica consolidada de preços a partir da planilha ao
    vivo (aba "Pedidos", mesma base do Portal Gestão de Compras). Resultado
    fica em cache por 60s (mesmo TTL de carregar_dados_seguros() no Portal)
    para não bater na API do Google Sheets a cada rerun.

    Regra de negócio (decisão explícita do usuário — "HISTÓRICO É HISTÓRICO"):
    por padrão TODAS as linhas do histórico entram na consolidação,
    independente do Status Aprov (Aprovado, Bloqueado, Rejeitado, Pendente).
    Passe status_filtro='Aprovado' apenas se quiser considerar só compras
    efetivadas em um caso específico.

    Retorna:
        base_precos: DataFrame consolidado por item (1 linha por Cod_Norm)
        historico_bruto: DataFrame do histórico filtrado, linha a linha
                          (usado na consulta rápida por código)
        status_msg: texto para o badge de status no topo da página
    """
    try:
        df = _ler_historico_bruto()
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), f"Erro ao ler a planilha de Pedidos: {e}"

    faltando = [c for c in HIST_REQUIRED_COLS if c not in df.columns]
    if faltando:
        return pd.DataFrame(), pd.DataFrame(), (
            f"Planilha de Pedidos inválida — colunas não encontradas: {faltando}. "
            "Confira os cabeçalhos da aba \"Pedidos\"."
        )

    base_precos, df_f = consolidar_base_precos(df, status_filtro)

    status_msg = f"Base atualizada em: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}"
    return base_precos, df_f, status_msg


base_precos, historico_bruto, status_historico = construir_base_precos()

# ==============================================================================
# Cabeçalho + Configurações (upload do mapa de cotação e exportação)
#
# A base histórica não é mais editável pelo painel — ela vem exclusivamente
# do historico_compras.csv versionado no repositório do GitHub. Para
# atualizar os preços históricos, atualize esse arquivo no repositório.
# ==============================================================================
# Cabeçalho com marca — mesmo padrão do Portal Gestão de Compras
# (comum.py: renderizar_cabecalho): logo + "Coordenação de Suprimentos" +
# título, com o alternador de tema no canto direito do cartão.
base64_logo = get_base64_logo(st.session_state['tema'])
with st.container(key="header_card"):
    c1, c2, c3 = st.columns([1.5, 6.0, 1.5])
    with c1:
        if base64_logo:
            st.markdown(
                f'<img src="data:image/png;base64,{base64_logo}" style="width:130px; display:block;">',
                unsafe_allow_html=True,
            )
    with c2:
        st.markdown(
            '''
            <div class="brand-text-block">
                <span class="brand-eyebrow">Coordenação de Suprimentos</span>
                <span class="brand-subtitle">Gestão Estratégica de Compras | Mapa de Cotação</span>
            </div>
            ''',
            unsafe_allow_html=True,
        )
    with c3:
        modo_escuro = st.toggle(
            "☀️ / 🌙",
            value=(st.session_state['tema'] == 'Escuro'),
            key="seletor_tema",
            help="Alternar entre tema claro e escuro",
        )
        tema_escolhido = "Escuro" if modo_escuro else "Claro"
        if tema_escolhido != st.session_state['tema']:
            st.session_state['tema'] = tema_escolhido
            st.rerun()

with st.expander("⚙️ Abrir / Fechar Configurações (Upload e Exportação)", expanded=False):
    col_exp1, col_exp2 = st.columns([2, 1])

    with col_exp1:
        st.markdown("##### 📁 Upload do Mapa de Cotação")
        uploaded_cot = st.file_uploader(
            "Carregar Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml)",
            type=["csv", "xlsx", "docx", "mhtml", "html"]
        )
    with col_exp2:
        st.markdown("##### 📥 Exportar")
        col_pdf, col_xlsx = st.columns(2)
        placeholder_pdf = col_pdf.empty()
        placeholder_xlsx = col_xlsx.empty()
        # Placeholders ficam vazios até o mapa ser processado — sem essa
        # mensagem parece que os botões de PDF/Excel sumiram ou quebraram.
        placeholder_pdf.button("📥 PDF", disabled=True, key="btn_pdf_desabilitado", use_container_width=True)
        placeholder_xlsx.button("📊 Excel", disabled=True, key="btn_xlsx_desabilitado", use_container_width=True)

st.markdown("---")


# ==============================================================================
# Funções de Conversão e Formatação, geração de PDF/Excel e o comparativo
# cotação x histórico agora vivem em logica_cotacao.py (nenhuma delas
# depende de st.*, então foram extraídas para módulo puro e testável — ver
# tests/). Seguem importadas no topo deste arquivo.
# ==============================================================================


# ==============================================================================
# Extração de tabela a partir de diferentes formatos de mapa de cotação
# (inalterado — já é razoavelmente robusto para csv/xlsx/docx/mhtml)
# ==============================================================================
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
                            except Exception:
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
                except Exception:
                    continue

        if dfs:
            df_principal = max(dfs, key=lambda x: len(x) * len(x.columns))
            df_principal.columns = [str(c).strip() for c in df_principal.columns]
            return df_principal

    except Exception as e:
        st.error(f"Erro ao processar o arquivo MHTML: {e}")
    return pd.DataFrame()


def extrair_tabela_excel_inteligente(arquivo_excel):
    """Acha a linha de cabeçalho da tabela de itens em QUALQUER aba do
    arquivo, em vez de confiar no nome da aba — exports do TOTVS costumam
    ter uma aba "Parametros" (metadados, sem itens) antes da aba real de
    dados ("Analise da Cotação", "Cotação por Produto" etc., nome varia).
    Pontua cada linha candidata por quantas categorias de termo de
    cabeçalho ela cobre (código/produto, descrição, fornecedor/valor) e
    fica com a de maior pontuação — a linha real de cabeçalho da tabela
    de itens cobre as três; uma pergunta de metadado tipo "Descrição do
    produto ?" cobre no máximo duas.
    """
    def normalizar(txt):
        return "".join([c for c in unicodedata.normalize('NFKD', str(txt).lower()) if not unicodedata.combining(c)])

    try:
        xls = pd.ExcelFile(arquivo_excel)

        melhor_sheet, melhor_header_idx, melhor_pontuacao = xls.sheet_names[0], 0, -1
        for s in xls.sheet_names:
            df_raw = pd.read_excel(xls, sheet_name=s, header=None, dtype=str)
            for idx, row in df_raw.iterrows():
                row_norm = normalizar(" ".join(str(x) for x in row.values if pd.notna(x)))
                pontuacao = sum([
                    any(t in row_norm for t in ('codigo', 'produto', 'item')),
                    'descricao' in row_norm,
                    any(t in row_norm for t in ('fornecedor', 'vlr', 'valor', 'preco')),
                ])
                if pontuacao > melhor_pontuacao:
                    melhor_pontuacao, melhor_sheet, melhor_header_idx = pontuacao, s, idx
                if pontuacao == 3:
                    break

        df = pd.read_excel(xls, sheet_name=melhor_sheet, header=melhor_header_idx, dtype=str)
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


# ==============================================================================
# PROCESSAMENTO DO MAPA DE COTAÇÃO (VIA UPLOAD)
#
# Mudança-chave em relação à versão anterior: em vez de varrer célula a
# célula do histórico procurando o código (O(itens_mapa × linhas_histórico),
# lento e sujeito a falso-positivo), fazemos um merge por código normalizado
# contra a base_precos já consolidada — a mesma lógica de build_comparativo.py
# da skill. A "melhor cotação" por item é sempre o MENOR valor entre as
# linhas daquele código (cobre tanto o mapa já resumido, uma linha por item,
# quanto o mapa bruto multi-fornecedor com código só na primeira linha do
# grupo — usamos ffill no código antes de agrupar).
# ==============================================================================
cotacao = pd.DataFrame()
if uploaded_cot is not None:
    bar = st.progress(0)
    status_processamento = st.empty()
    status_processamento.text("Processando dados...")
    for i in range(100):
        time.sleep(0.005)
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
    status_processamento.empty()

# Número da cotação (ex.: "021132") pra nomear os arquivos exportados
# (PDF/Excel) igual ao que consta no arquivo enviado, em vez de um nome
# genérico — vem da coluna "COTACAO :" do export do TOTVS; se não achar
# (outro formato de arquivo, ou coluna ausente), cai pro nome do próprio
# arquivo enviado (que na prática já é o número da cotação, ex. "021132.xlsx").
numero_cotacao = extrair_numero_cotacao(cotacao)
if not numero_cotacao and uploaded_cot is not None:
    numero_cotacao = os.path.splitext(uploaded_cot.name)[0].strip() or None
if numero_cotacao:
    numero_cotacao = "".join(c if (c.isalnum() or c in "-_") else "_" for c in numero_cotacao)

# O comparativo cotação x histórico em si (achar colunas, achar a melhor
# cotação por item, calcular variação vs. último/médio, montar a
# "Observação") não depende de st.* — vive em processar_mapa_cotacao()
# (logica_cotacao.py) para poder ser testado sem Streamlit.
df_final, aviso_valores_estranhos = processar_mapa_cotacao(cotacao, base_precos)

if not df_final.empty:
    st.subheader("📋 Mapa de Cotação Consolidado & Comparativo Histórico")

    if aviso_valores_estranhos:
        st.warning(
            "⚠️ Uma fração grande dos itens tem variação muito alta e sistemática em relação ao "
            "histórico (dezenas de vezes o preço anterior). Isso costuma indicar que os valores do "
            "mapa e do histórico não estão na mesma unidade (ex.: valor de um lote vs. preço "
            "unitário). Confira antes de usar esse comparativo para decisão de compra."
        )

    n_sem_historico = (df_final['Observação'] == 'Sem histórico de compra').sum()
    if n_sem_historico:
        st.info(f"ℹ️ {n_sem_historico} item(ns) sem correspondência no histórico de compras — "
                f"nenhum preço de referência foi inventado para eles.")

    # ==========================================================================
    # Grade interativa (ordenação e filtro por coluna, ao estilo Excel)
    #
    # A grade recebe os valores brutos (numéricos/data), não as strings já
    # formatadas — assim ordenar e filtrar funciona pelo valor real (ex.:
    # maior/menor preço, faixa de variação %), e a formatação em R$/%% fica
    # só na exibição, via valueFormatter.
    # ==========================================================================
    colunas_moeda = ['Valor Cotado (R$)', 'Último Preço Pago (R$)', 'Preço Médio (R$)',
                      'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)']
    colunas_pct = ['Var. vs Último (%)', 'Var. vs Médio (%)']

    df_grid = df_final.copy()
    for col in colunas_moeda + colunas_pct:
        df_grid[col] = pd.to_numeric(df_grid[col].replace("", np.nan), errors='coerce')
    df_grid['Data Última Compra'] = pd.to_datetime(
        df_grid['Data Última Compra'].replace("", np.nan), errors='coerce'
    ).dt.strftime('%Y-%m-%d')

    formatter_moeda = JsCode("""
        function(params) {
            if (params.value === null || params.value === undefined) { return ''; }
            return 'R$ ' + Number(params.value).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
        }
    """)
    formatter_qtd = JsCode("""
        function(params) {
            if (params.value === null || params.value === undefined) { return ''; }
            var v = Number(params.value);
            var casas = Number.isInteger(v) ? 0 : 2;
            return v.toLocaleString('pt-BR', {minimumFractionDigits: casas, maximumFractionDigits: casas});
        }
    """)
    formatter_data = JsCode("""
        function(params) {
            if (!params.value) { return ''; }
            var p = params.value.split('-');
            return p[2] + '/' + p[1] + '/' + p[0];
        }
    """)
    formatter_pct = JsCode("""
        function(params) {
            if (params.value === null || params.value === undefined) { return ''; }
            var v = Number(params.value);
            var txt = (v > 0 ? '+' : '') + v.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2}) + '%';
            if (v > 0) { return '↑ ' + txt; }
            if (v < 0) { return '↓ ' + txt; }
            return txt;
        }
    """)
    cellstyle_pct = JsCode("""
        function(params) {
            if (params.value === null || params.value === undefined) { return {}; }
            var v = Number(params.value);
            if (v > 0) { return {color: '#c00000', fontWeight: 900}; }
            if (v < 0) { return {color: '#2ca02c', fontWeight: 900}; }
            return {color: '#555555', fontWeight: 'bold'};
        }
    """)

    t_grid = TEMAS.get(st.session_state['tema'], TEMAS['Claro'])
    tema_aggrid = 'dark' if st.session_state['tema'] == 'Escuro' else 'light'

    gb = GridOptionsBuilder.from_dataframe(df_grid)
    gb.configure_default_column(sortable=True, filter=True, resizable=True)
    # GridOptionsBuilder.from_dataframe já define autoSizeStrategy
    # "fitGridWidth" (estica as colunas pra preencher a tela, espremendo o
    # texto). Sobrescreve para "fitCellContents": cada coluna (cabeçalho e
    # célula, numa linha só) fica larga o suficiente pro maior conteúdo,
    # sem cortar nem quebrar linha — sobra vira rolagem horizontal.
    gb.configure_grid_options(autoSizeStrategy={"type": "fitCellContents"})
    gb.configure_column('Item', width=70, cellStyle={'textAlign': 'center'})
    gb.configure_column('Código', width=100, cellStyle={'textAlign': 'center'})
    gb.configure_column('Descrição', width=260, cellStyle={'textAlign': 'left'})
    gb.configure_column('Unidade', width=90, cellStyle={'textAlign': 'center'})
    gb.configure_column('Qtd', type=['numericColumn'], valueFormatter=formatter_qtd,
                         cellStyle={'textAlign': 'right'}, width=90)
    gb.configure_column('Fornecedor Cotado', width=200, cellStyle={'textAlign': 'left'})
    for col in colunas_moeda:
        gb.configure_column(col, type=['numericColumn'], valueFormatter=formatter_moeda,
                             cellStyle={'textAlign': 'right'}, width=150)
    gb.configure_column('Data Última Compra', valueFormatter=formatter_data,
                         cellStyle={'textAlign': 'center'}, width=130)
    gb.configure_column('Fornecedor Última Compra', width=200, cellStyle={'textAlign': 'left'})
    for col in colunas_pct:
        gb.configure_column(col, type=['numericColumn'], valueFormatter=formatter_pct,
                             cellStyle=cellstyle_pct, width=140)
    gb.configure_column('Observação', width=180, cellStyle={'textAlign': 'center'})
    grid_options = gb.build()

    AgGrid(
        df_grid,
        gridOptions=grid_options,
        theme=tema_aggrid,
        height=600,
        allow_unsafe_jscode=True,
        custom_css={
            ".ag-header-cell-label": {"font-weight": "bold", "justify-content": "center"},
            ".ag-cell": {"font-size": "12px"},
        },
        key="grid_mapa_cotacao",
    )

    sufixo_arquivo = f"_{numero_cotacao}" if numero_cotacao else ""

    pdf_bytes = gerar_pdf(df_final, numero_cotacao=numero_cotacao)
    placeholder_pdf.download_button(
        label="📥 PDF",
        data=pdf_bytes,
        file_name=f"mapa_de_cotacao{sufixo_arquivo}.pdf",
        mime="application/pdf",
        key="btn_pdf_top",
        use_container_width=True,
    )
    xlsx_bytes = gerar_excel(df_final, numero_cotacao=numero_cotacao)
    placeholder_xlsx.download_button(
        label="📊 Excel",
        data=xlsx_bytes,
        file_name=f"Comparativo_Cotacao{sufixo_arquivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_xlsx_top",
        use_container_width=True,
    )
elif uploaded_cot is not None:
    st.warning("⚠️ Nenhum item válido encontrado no arquivo carregado.")

st.markdown("---")

# ==============================================================================
# 🔍 BARRA DE CONSULTA COMPACTA E CONGELADA NO RODAPÉ DA TELA
#
# Agora consulta diretamente historico_bruto (DataFrame com colunas nomeadas
# já filtrado/normalizado em construir_base_precos), em vez de varrer célula
# a célula do CSV cru.
# ==============================================================================
st.markdown('<div class="footer-pesquisa">', unsafe_allow_html=True)
st.markdown("**🔍 Consulta Rápida de Histórico por Código do Item**")
codigo_pesquisa = st.text_input("Digite ou cole o código do item:", "", key="input_pesquisa_fixa")
st.markdown('</div>', unsafe_allow_html=True)

if codigo_pesquisa:
    cod_norm_pesquisa = normalizar_codigo(codigo_pesquisa)

    if historico_bruto.empty or cod_norm_pesquisa is None:
        st.warning(f"⚠️ Nenhuma compra anterior encontrada no histórico para o código **{codigo_pesquisa}**.")
    else:
        registros = historico_bruto[historico_bruto['Cod_Norm'] == cod_norm_pesquisa].sort_values('Data Emissao')

        if registros.empty:
            st.warning(f"⚠️ Nenhuma compra anterior encontrada no histórico para o código **{cod_norm_pesquisa}**.")
        else:
            st.success(f"Foram encontradas **{len(registros)}** ocorrência(s) de compra para o código **{cod_norm_pesquisa}**:")

            df_historico_item = pd.DataFrame({
                'Pedido': registros['Pedido'] if 'Pedido' in registros.columns else '',
                'Data Emissão PC': registros['Data Emissao'].dt.strftime('%d/%m/%Y'),
                'Fornecedor da Compra': registros['Nome Fornece'],
                'Status Aprov': registros['Status Aprov'],
                'Quantidade': registros['Quantidade'],
                'Prc Unitario': registros['Prc Unitario'].apply(formatar_brl),
            }).set_index('Pedido')
            st.table(df_historico_item)

            df_chart = registros[registros['Prc Unitario'] > 0].dropna(subset=['Data Emissao']).copy()
            if not df_chart.empty:
                st.markdown("#### 📈 Evolução do Preço Histórico")
                df_chart = df_chart.sort_values('Data Emissao')
                df_chart["Data Formatada"] = df_chart["Data Emissao"].dt.strftime('%d/%m/%Y')
                df_chart["Preço Formatado BR"] = df_chart["Prc Unitario"].apply(
                    lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

                fig = px.line(
                    df_chart,
                    x="Data Formatada",
                    y="Prc Unitario",
                    markers=True,
                    line_shape="spline",
                    text="Preço Formatado BR"
                )

                fig.update_traces(
                    fill='tozeroy',
                    line=dict(color='#00d2c4', width=3),
                    marker=dict(size=8, color='#00d2c4'),
                    textposition="top center"
                )

                fig.update_layout(
                    xaxis_title="Data Emissão",
                    yaxis_title="",
                    plot_bgcolor='rgba(255,255,255,0.02)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#333333'),
                    xaxis=dict(showgrid=True, gridcolor='rgba(200,200,200,0.3)'),
                    yaxis=dict(
                        showgrid=True,
                        gridcolor='rgba(200,200,200,0.3)',
                        showticklabels=False,
                    ),
                    margin=dict(l=20, r=20, t=50, b=20)
                )

                st.plotly_chart(fig, use_container_width=True)

st.markdown(
    f"<div style='text-align: center; margin-top: 30px;'>"
    f"<span class='status-badge'>ℹ️ {status_historico}</span></div>",
    unsafe_allow_html=True
)
