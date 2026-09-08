import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import docx
from fpdf import FPDF
import unicodedata
import email
from bs4 import BeautifulSoup
import datetime
import time
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

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
    ),
    'Escuro': dict(
        main_bg='#0e1117', text='#f5f5f5', body_text='#e6e6e6',
        badge_bg='#1c2e4a', badge_text='#8ab4f8', badge_border='#2c4770',
        expander_border='#333a45', expander_bg='#161a20', expander_header_bg='#1c212a',
        footer_bg='#161a20', footer_border='#4f7cff', footer_text='#f5f5f5',
        th_bg='#1f2a3f', th_text='#e8eef7', th_border='#3a4a66',
        td_border='#2a2f38', td_text='#e6e6e6', row_bg='#161a20',
    ),
}


def gerar_css(tema: str) -> str:
    t = TEMAS.get(tema, TEMAS['Claro'])
    return f"""
    <style>
    .main {{ background-color: {t['main_bg']}; }}
    [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
        background-color: {t['main_bg']} !important;
    }}
    body, p, span, label, .stMarkdown {{ color: {t['body_text']}; }}
    h1 {{ color: {t['text']}; font-family: 'Helvetica Neue', sans-serif; margin-bottom: 5px; }}
    h2, h3, h4 {{ color: {t['text']}; }}

    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 150px !important;
        max-width: 100% !important;
    }}
    header {{ visibility: hidden !important; }}
    #MainMenu {{ visibility: hidden !important; }}
    footer {{ visibility: hidden !important; }}
    div[data-baseweb="modal"], div.stDialog, div[role="dialog"] {{
        display: none !important;
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
        border: 1px solid {t['expander_border']} !important;
        background-color: {t['expander_bg']} !important;
        border-radius: 6px !important;
        box-shadow: none !important;
        margin-bottom: 20px !important;
    }}
    .streamlit-expanderHeader {{
        padding-top: 8px !important;
        padding-bottom: 8px !important;
        min-height: 40px !important;
        font-size: 14px !important;
        background-color: {t['expander_header_bg']} !important;
        border-radius: 6px !important;
        color: {t['text']} !important;
    }}
    .streamlit-expanderContent {{
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


st.markdown(gerar_css(st.session_state['tema']), unsafe_allow_html=True)

# ==============================================================================
# 1. BASE HISTÓRICA — mesma lógica da skill "analise-mapa-cotacao"
#
# Antes, o painel lia historico_compras.csv sem cabeçalho e "adivinhava"
# posições de coluna (h_row.get(2), get(4), get(10)...). Isso é frágil e não
# bate com o processo validado na análise manual. Agora o histórico é lido
# pelo export padrão de Pedidos de Compra (título na linha 1, cabeçalho na
# linha 2), com colunas nomeadas, e a base é consolidada por código de
# Produto exatamente como build_base_precos.py da skill.
# ==============================================================================

# Único arquivo-fonte: o historico_compras.csv versionado no repositório do
# GitHub, sempre no mesmo diretório do app.py. Não há upload/edição pelo
# painel — para atualizar a base, atualize o arquivo no repositório.
HISTORICO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "historico_compras.csv")

HIST_REQUIRED_COLS = ['Produto', 'Descricao.1', 'Unidade', 'Prc Unitario',
                       'Data Emissao', 'Nome Fornece', 'Quantidade', 'Status Aprov']


def limpar_valor(valor):
    if pd.isna(valor) or valor is None:
        return 0.0
    val_str = str(valor).replace('R$', '').strip()
    if not val_str or val_str.lower() in ['nan', 'total item', 'total', '##########', 'a vista', '25 dias', 'item', 'código', 'produto', 'descrição']:
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
    except Exception:
        return 0.0


def _ler_historico_bruto(caminho_historico: str) -> pd.DataFrame:
    """Lê o histórico de Pedidos de Compra, aceitando tanto .csv quanto
    .xlsx — mesmo layout em ambos os casos (título na linha 1, cabeçalho de
    colunas na linha 2, export padrão TOTVS/Protheus).

    O export brasileiro do TOTVS/Protheus normalmente sai em .csv com:
      - ';' como separador de campo (porque ',' é o separador decimal);
      - BOM no início do arquivo (necessita encoding 'utf-8-sig');
      - números como texto, com separador de milhar '.' e decimal ','
        (ex.: " 9.500,0000 ", com espaços — tratado depois por limpar_valor).
    Tentamos ';' primeiro (caso mais comum) e caímos para ',' só se as
    colunas esperadas não aparecerem — cobre também um CSV exportado no
    padrão internacional (vírgula), como o simulado durante os testes.
    """
    for sep in (';', ','):
        try:
            df = pd.read_csv(caminho_historico, header=1, sep=sep, encoding='utf-8-sig')
        except Exception:
            continue
        if 'Produto' in df.columns and 'Prc Unitario' in df.columns:
            return df
    # Nenhuma tentativa achou as colunas esperadas — devolve a última leitura
    # (com ',') para que a checagem de HIST_REQUIRED_COLS gere uma mensagem
    # de erro clara para o usuário, em vez de travar aqui.
    return pd.read_csv(caminho_historico, header=1, sep=',', encoding='utf-8-sig')


def normalizar_codigo(valor):
    """Extrai só os dígitos de um código de produto e remove zeros à esquerda,
    para casar códigos vindos com ou sem padding (ex.: '0000001268', '1268' e
    1268 devem ser tratados como o mesmo item)."""
    if pd.isna(valor):
        return None
    s = ''.join(filter(str.isdigit, str(valor)))
    if not s:
        return None
    return str(int(s))


@st.cache_data(show_spinner="Consolidando base histórica de preços...")
def construir_base_precos(caminho_historico: str, mtime: float, status_filtro: str = None):
    """
    Constrói a base histórica consolidada de preços.

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
    if not os.path.exists(caminho_historico):
        return pd.DataFrame(), pd.DataFrame(), (
            f"Base de dados indisponível — arquivo '{os.path.basename(caminho_historico)}' "
            "não encontrado no repositório."
        )

    try:
        df = _ler_historico_bruto(caminho_historico)
    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), f"Erro ao ler histórico: {e}"

    faltando = [c for c in HIST_REQUIRED_COLS if c not in df.columns]
    if faltando:
        return pd.DataFrame(), pd.DataFrame(), (
            f"Histórico inválido — colunas faltando: {faltando}. "
            "Confira se o arquivo tem título na linha 1 e cabeçalho na linha 2."
        )

    if status_filtro:
        df_f = df[df['Status Aprov'] == status_filtro].copy()
    else:
        df_f = df.copy()

    # Coerção defensiva: se o CSV/XLSX vier com preço em formato BR
    # ("1.234,56") ou como texto, limpar_valor normaliza; se já vier
    # numérico (caso comum quando exportado do Excel), pd.to_numeric resolve
    # sem precisar reprocessar linha a linha.
    if not pd.api.types.is_numeric_dtype(df_f['Prc Unitario']):
        df_f['Prc Unitario'] = df_f['Prc Unitario'].apply(limpar_valor)
    if not pd.api.types.is_numeric_dtype(df_f['Quantidade']):
        df_f['Quantidade'] = df_f['Quantidade'].apply(limpar_valor)

    df_f = df_f[df_f['Prc Unitario'].notna() & (df_f['Prc Unitario'] > 0) & df_f['Produto'].notna()].copy()
    df_f['Data Emissao'] = pd.to_datetime(df_f['Data Emissao'], errors='coerce', dayfirst=True)
    df_f['Cod_Norm'] = df_f['Produto'].apply(normalizar_codigo)
    df_f = df_f.dropna(subset=['Cod_Norm']).sort_values('Data Emissao')

    ultimo = df_f.groupby('Cod_Norm').tail(1)[
        ['Cod_Norm', 'Produto', 'Descricao.1', 'Unidade', 'Prc Unitario', 'Data Emissao', 'Nome Fornece']
    ].rename(columns={
        'Descricao.1': 'Descricao_Item',
        'Prc Unitario': 'Ultimo_Preco',
        'Data Emissao': 'Data_Ultima_Compra',
        'Nome Fornece': 'Fornecedor_Ultima_Compra',
    })

    agg = df_f.groupby('Cod_Norm').agg(
        Preco_Medio=('Prc Unitario', 'mean'),
        Preco_Minimo=('Prc Unitario', 'min'),
        Preco_Maximo=('Prc Unitario', 'max'),
        Qtd_Compras=('Prc Unitario', 'count'),
        Qtd_Total_Comprada=('Quantidade', 'sum'),
        Fornecedores_Distintos=('Nome Fornece', pd.Series.nunique),
    ).reset_index()

    base_precos = ultimo.merge(agg, on='Cod_Norm', how='left')

    data_mod = datetime.datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %H:%M')
    status_msg = f"Base atualizada em: {data_mod}"
    return base_precos, df_f, status_msg


def _mtime_or_zero(path):
    return os.path.getmtime(path) if os.path.exists(path) else 0.0


base_precos, historico_bruto, status_historico = construir_base_precos(
    HISTORICO_PATH, _mtime_or_zero(HISTORICO_PATH)
)

# ==============================================================================
# Cabeçalho + Configurações (upload do mapa de cotação e exportação)
#
# A base histórica não é mais editável pelo painel — ela vem exclusivamente
# do historico_compras.csv versionado no repositório do GitHub. Para
# atualizar os preços históricos, atualize esse arquivo no repositório.
# ==============================================================================
with st.expander("⚙️ Abrir / Fechar Configurações (Upload e Exportação)", expanded=False):
    col_exp0, col_exp1, col_exp2 = st.columns([1, 2, 1])

    with col_exp0:
        st.markdown("### 🎨 Tema")
        opcoes_tema = list(TEMAS.keys())
        tema_escolhido = st.radio(
            "Aparência do painel",
            options=opcoes_tema,
            index=opcoes_tema.index(st.session_state['tema']),
            key="seletor_tema",
            horizontal=True,
        )
        if tema_escolhido != st.session_state['tema']:
            st.session_state['tema'] = tema_escolhido
            st.rerun()

    with col_exp1:
        st.markdown("### 📁 Upload do Mapa de Cotação")
        uploaded_cot = st.file_uploader(
            "Carregar Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml)",
            type=["csv", "xlsx", "docx", "mhtml", "html"]
        )
    with col_exp2:
        st.markdown("### 📥 Exportar")
        placeholder_pdf = st.empty()
        placeholder_xlsx = st.empty()

st.title("📊 Gestão Estratégica de Compras | Mapa de Cotação")
st.markdown("---")


# ==============================================================================
# Funções de Conversão e Formatação (inalteradas)
# ==============================================================================


def formatar_brl(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except Exception:
        return ""
    if val_float <= 0:
        return ""
    return f"R$ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def formatar_qtd(valor):
    try:
        val_float = float(valor)
    except Exception:
        val_float = 0.0

    if val_float.is_integer():
        return f"{int(val_float):,}".replace(',', '.')
    else:
        return f"{val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def formatar_pct(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except Exception:
        val_float = 0.0
    return f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')


def formatar_pct_com_seta(valor):
    if valor == "" or pd.isna(valor) or valor is None:
        return ""
    try:
        val_float = float(valor)
    except Exception:
        val_float = 0.0

    val_fmt = f"{val_float:+,.2f}%".replace(',', 'X').replace('.', ',').replace('X', '.')

    if val_float > 0:
        return f"<span style='color: #c00000; font-size: 14px; font-weight: 900; white-space: nowrap;'>↑ {val_fmt}</span>"
    elif val_float < 0:
        return f"<span style='color: #2ca02c; font-size: 14px; font-weight: 900; white-space: nowrap;'>↓ {val_fmt}</span>"
    else:
        return f"<span style='color: #555555; font-size: 12px; font-weight: bold; white-space: nowrap;'>{val_fmt}</span>"


def limpar_texto_pdf(texto):
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return texto_sem_acento.encode('latin-1', 'replace').decode('latin-1')


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
    try:
        xls = pd.ExcelFile(arquivo_excel)
        sheet_name = xls.sheet_names[0]
        # Prioriza abas cujo nome sugira o mapa de cotação por produto,
        # espelhando a heurística usada na skill (Passo 2).
        for s in xls.sheet_names:
            s_low = s.lower()
            if 'cota' in s_low and 'produto' in s_low:
                sheet_name = s
                break

        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)

        header_row_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)]).lower()
            row_str_norm = "".join([c for c in unicodedata.normalize('NFKD', row_str) if not unicodedata.combining(c)])
            if 'codigo' in row_str_norm or 'descricao' in row_str_norm or 'vlr' in row_str_norm or 'preco' in row_str_norm or 'item' in row_str_norm:
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


# ==============================================================================
# Exportação — PDF
# ==============================================================================
def gerar_pdf(df):
    class PDFProfissional(FPDF):
        def __init__(self):
            super().__init__(orientation='L', unit='mm', format='A4')
            self.set_margins(left=5.0, top=19.1, right=5.0)
            self.set_auto_page_break(auto=True, margin=19.1)

        def header(self):
            self.set_fill_color(47, 85, 151)
            self.rect(5.0, 8, 287.0, 20, 'F')

            self.set_font("helvetica", "B", 14)
            self.set_text_color(255, 255, 255)
            self.set_xy(5.0, 10)
            self.cell(287.0, 6, limpar_texto_pdf("Mapa de Cotacao & Comparativo Historico"), 0, 1, "C")

            self.set_font("helvetica", "", 9)
            self.set_xy(5.0, 16)
            self.cell(287.0, 5, limpar_texto_pdf("Gestao Estrategica de Compras | Parente Andrade"), 0, 1, "C")
            self.ln(10)

        def footer(self):
            self.set_y(-12)
            self.set_font("helvetica", "I", 8)
            self.set_text_color(128, 128, 128)
            data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
            self.cell(0, 8, limpar_texto_pdf(f"Gerado em {data_hora} | Pagina {self.page_no()}"), 0, 0, "C")

    pdf = PDFProfissional()
    pdf.add_page()

    col_widths = [8, 16, 44, 8, 17, 34, 17, 34, 17, 20, 20, 20, 22]
    headers = [
        "Item", "Codigo", "Descricao", "Qtd",
        "Vl Cotado", "Forn. Cotado", "Ult. Preco",
        "Forn. Ult.", "Preco Med.", "Var vs Med(%)", "Preco Min.", "Preco Max.", "Observacao"
    ]

    pdf.set_fill_color(47, 85, 151)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 6.5)

    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 7, limpar_texto_pdf(h), border=1, fill=True, align="C")
    pdf.ln()

    pdf.set_font("helvetica", "", 6.0)

    fill = False
    for _, row in df.iterrows():
        if fill:
            pdf.set_fill_color(242, 245, 249)
        else:
            pdf.set_fill_color(255, 255, 255)

        pdf.set_text_color(0, 0, 0)
        var_val = row['Var. vs Médio (%)']

        ult_preco_val = row['Último Preço Pago (R$)']
        ult_preco_str = formatar_brl(ult_preco_val) if ult_preco_val != "" else ""
        forn_ant_str = str(row['Fornecedor Última Compra']) if ult_preco_val != "" else ""

        preco_med_val = row['Preço Médio (R$)']
        preco_med_str = formatar_brl(preco_med_val) if preco_med_val != "" else ""

        min_val = row['Preço Mín. Histórico (R$)']
        min_str = formatar_brl(min_val) if min_val != "" else ""
        max_val = row['Preço Máx. Histórico (R$)']
        max_str = formatar_brl(max_val) if max_val != "" else ""

        var_str = formatar_pct(var_val) if var_val != "" else ""

        pdf.cell(col_widths[0], 6, limpar_texto_pdf(str(row['Item'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[1], 6, limpar_texto_pdf(str(row['Código'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[2], 6, limpar_texto_pdf(str(row['Descrição'])[:30]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[3], 6, limpar_texto_pdf(str(row['Qtd'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[4], 6, limpar_texto_pdf(formatar_brl(row['Valor Cotado (R$)'])), border=1, fill=fill, align="R")
        pdf.cell(col_widths[5], 6, limpar_texto_pdf(str(row['Fornecedor Cotado'])[:20]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[6], 6, limpar_texto_pdf(ult_preco_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[7], 6, limpar_texto_pdf(forn_ant_str[:20]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[8], 6, limpar_texto_pdf(preco_med_str), border=1, fill=fill, align="R")

        if var_val != "":
            if var_val < 0:
                pdf.set_text_color(44, 160, 44)
            elif var_val > 0:
                pdf.set_text_color(192, 0, 0)
            else:
                pdf.set_text_color(0, 0, 0)

        pdf.cell(col_widths[9], 6, limpar_texto_pdf(var_str), border=1, fill=fill, align="R")

        pdf.set_text_color(0, 0, 0)
        pdf.cell(col_widths[10], 6, limpar_texto_pdf(min_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[11], 6, limpar_texto_pdf(max_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[12], 6, limpar_texto_pdf(str(row['Observação'])[:22]), border=1, fill=fill, align="L")

        pdf.ln()
        fill = not fill

    pdf_output = pdf.output(dest='S')
    if isinstance(pdf_output, str):
        return pdf_output.encode('latin1')
    return bytes(pdf_output)


# ==============================================================================
# Exportação — Excel formatado (mesmo padrão visual de build_comparativo.py
# da skill: cores por faixa de variação, moeda, %, congelamento de painel)
# ==============================================================================
FONT_XLSX = 'Arial'
HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
HEADER_FONT = Font(name=FONT_XLSX, bold=True, color='FFFFFF', size=10)
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREEN = PatternFill(start_color='C6E0B4', end_color='C6E0B4', fill_type='solid')
YELLOW = PatternFill(start_color='FFE699', end_color='FFE699', fill_type='solid')
RED = PatternFill(start_color='F8CBAD', end_color='F8CBAD', fill_type='solid')
GRAY = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')


def _var_fill(v):
    if v == "" or pd.isna(v):
        return GRAY
    if v <= -5:
        return GREEN
    if v >= 5:
        return RED
    return YELLOW


def gerar_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo_Cotacao"

    headers = ['Item', 'Cód. Produto', 'Descrição do Item', 'Unid.', 'Qtd',
               'Fornecedor Cotado', 'Valor Cotado (R$)',
               'Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
               'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)',
               'Var. vs Último (%)', 'Var. vs Médio (%)', 'Observação']

    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[1].height = 36

    for _, r in df.iterrows():
        data_str = ""
        if r['Data Última Compra'] not in ("", None) and pd.notna(r['Data Última Compra']):
            try:
                data_str = pd.to_datetime(r['Data Última Compra']).date()
            except Exception:
                data_str = r['Data Última Compra']
        ws.append([
            r['Item'], r['Código'], r['Descrição'], r['Unidade'], r['Qtd'],
            r['Fornecedor Cotado'], float(r['Valor Cotado (R$)']) if r['Valor Cotado (R$)'] != "" else None,
            float(r['Último Preço Pago (R$)']) if r['Último Preço Pago (R$)'] != "" else None,
            data_str,
            r['Fornecedor Última Compra'],
            float(r['Preço Médio (R$)']) if r['Preço Médio (R$)'] != "" else None,
            float(r['Preço Mín. Histórico (R$)']) if r['Preço Mín. Histórico (R$)'] != "" else None,
            float(r['Preço Máx. Histórico (R$)']) if r['Preço Máx. Histórico (R$)'] != "" else None,
            round(float(r['Var. vs Último (%)']), 2) if r['Var. vs Último (%)'] != "" else None,
            round(float(r['Var. vs Médio (%)']), 2) if r['Var. vs Médio (%)'] != "" else None,
            r['Observação'],
        ])

    last_row = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = Font(name=FONT_XLSX, size=10)
            cell.border = BORDER
        row[8].number_format = 'DD/MM/YYYY'
        for idx in (6, 7, 10, 11, 12):
            row[idx].number_format = '#,##0.00'
        for idx in (13, 14):
            row[idx].number_format = '+0.0"%";-0.0"%";0.0"%"'
        row[14].fill = _var_fill(row[14].value)

    widths = [7, 11, 42, 7, 7, 32, 15, 16, 16, 32, 15, 15, 15, 14, 14, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"

    leg_row = last_row + 2
    ws.cell(row=leg_row, column=1, value="Legenda (Var. vs Preço Médio):").font = Font(name=FONT_XLSX, bold=True, size=9)
    for i, (txt, fill) in enumerate([
        ("5%+ abaixo da média", GREEN), ("Dentro de ±5% da média", YELLOW),
        ("5%+ acima da média", RED), ("Sem histórico de compra", GRAY)]):
        c = ws.cell(row=leg_row + 1 + i, column=1, value=txt)
        c.font = Font(name=FONT_XLSX, size=9)
        c.fill = fill

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


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
    st.text("Processando dados...")
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

df_final = pd.DataFrame()
aviso_valores_estranhos = False

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
    c_unid = achar_coluna(cotacao, ['unidade', 'unid', 'und'])
    c_qtd = achar_coluna(cotacao, ['qtd', 'quantidade'])

    c_vlr = achar_coluna(cotacao, [
        'valor unitario', 'vlr. unitario', 'valor unit', 'vlr. unit', 'vlr unit', 'unitario',
        'preço unitario', 'preco unitario', 'preço unit', 'preco unit', 'vlr', 'preço', 'preco',
        'unit', 'vl unit', 'vl. unit', 'vl.unit', 'valor'
    ])

    c_forn = achar_coluna(cotacao, ['fornecedor', 'empresa', 'nome'])
    c_status = achar_coluna(cotacao, ['status'])

    # Forward-fill do código: cobre o formato bruto de mapa de cotação em que
    # o código só aparece na primeira linha de cada grupo de fornecedores
    # cotando o mesmo item (ver Passo 2 da skill).
    if c_cod:
        cotacao[c_cod] = cotacao[c_cod].ffill()

    if c_status and not cotacao.empty:
        df_vencedores = cotacao[cotacao[c_status].astype(str).str.contains(
            'vencedor|melhor preço|melhor preco', case=False, na=False)]
        if not df_vencedores.empty:
            cotacao = df_vencedores

    resultados_brutos = []
    item_contador = 1
    for idx, row in cotacao.iterrows():
        raw_cod = row[c_cod] if c_cod and pd.notna(row[c_cod]) else None
        if raw_cod is None:
            continue
        cod_norm = normalizar_codigo(raw_cod)
        if cod_norm is None:
            continue
        # pula linhas de cabeçalho repetido / rodapé sem sentido
        desc_check = str(row[c_desc]) if c_desc and pd.notna(row[c_desc]) else ""
        if desc_check.strip().lower() in ['descrição', 'descricao', 'nan', '']:
            continue

        num_item = str(row[c_item]) if c_item and pd.notna(row[c_item]) else f"{item_contador:04d}"
        desc = str(row[c_desc]) if c_desc and pd.notna(row[c_desc]) else 'Descrição não informada'
        unidade = str(row[c_unid]) if c_unid and pd.notna(row[c_unid]) else ''
        qtd = limpar_valor(row[c_qtd]) if c_qtd and pd.notna(row[c_qtd]) else 1.0

        valor_cotado = limpar_valor(row[c_vlr]) if c_vlr and pd.notna(row[c_vlr]) else 0.0
        if valor_cotado <= 0:
            for col_nome in row.index:
                val_tentativa = limpar_valor(row[col_nome])
                if val_tentativa > 0 and val_tentativa != qtd:
                    valor_cotado = val_tentativa
                    break
        if valor_cotado <= 0:
            continue

        forn_cotado = str(row[c_forn]) if c_forn and pd.notna(row[c_forn]) else 'Fornecedor não informado'

        item_contador += 1
        resultados_brutos.append({
            'Item': num_item,
            'Código': str(raw_cod).strip(),
            'Cod_Norm': cod_norm,
            'Descrição': desc,
            'Unidade': unidade,
            'Qtd': qtd,
            'Valor Cotado (R$)': valor_cotado,
            'Fornecedor Cotado': forn_cotado,
        })

    if resultados_brutos:
        df_bruto = pd.DataFrame(resultados_brutos)

        # Melhor cotação por item = MENOR valor entre as linhas do mesmo
        # código (Passo 2 da skill) — cobre tanto o mapa já resumido (uma
        # linha por item) quanto o mapa multi-fornecedor bruto.
        df_bruto = df_bruto.sort_values('Valor Cotado (R$)')
        df_melhor = df_bruto.drop_duplicates(subset=['Cod_Norm'], keep='first').copy()
        df_melhor = df_melhor.sort_values('Item')

        # Aviso de possível incompatibilidade de escala (lote vs. unitário),
        # igual à validação recomendada no Passo 2 da skill.
        if not base_precos.empty:
            checagem = df_melhor.merge(
                base_precos[['Cod_Norm', 'Ultimo_Preco']], on='Cod_Norm', how='left'
            )
            checagem = checagem[checagem['Ultimo_Preco'].notna() & (checagem['Ultimo_Preco'] > 0)]
            if len(checagem) > 0:
                razao = (checagem['Valor Cotado (R$)'] / checagem['Ultimo_Preco'])
                fora_da_curva = ((razao > 10) | (razao < 0.1)).mean()
                if fora_da_curva > 0.3:
                    aviso_valores_estranhos = True

        if base_precos.empty:
            df_merge = df_melhor.copy()
            for col in ['Ultimo_Preco', 'Data_Ultima_Compra', 'Fornecedor_Ultima_Compra',
                        'Preco_Medio', 'Preco_Minimo', 'Preco_Maximo']:
                df_merge[col] = np.nan
        else:
            df_merge = df_melhor.merge(
                base_precos[['Cod_Norm', 'Ultimo_Preco', 'Data_Ultima_Compra', 'Fornecedor_Ultima_Compra',
                             'Preco_Medio', 'Preco_Minimo', 'Preco_Maximo']],
                on='Cod_Norm', how='left'
            )

        def calc_var(row, campo):
            base = row[campo]
            if pd.isna(base) or base == 0:
                return ""
            return round((row['Valor Cotado (R$)'] - base) / base * 100, 2)

        def observacao(row):
            if pd.isna(row['Ultimo_Preco']):
                return 'Sem histórico de compra'
            var_medio = row['Var. vs Médio (%)']
            if var_medio == "":
                return 'Sem histórico de compra'
            if var_medio <= -5:
                return 'Abaixo da média histórica'
            if var_medio >= 5:
                return 'Acima da média histórica'
            return 'Próximo da média histórica'

        df_merge['Var. vs Último (%)'] = df_merge.apply(lambda r: calc_var(r, 'Ultimo_Preco'), axis=1)
        df_merge['Var. vs Médio (%)'] = df_merge.apply(lambda r: calc_var(r, 'Preco_Medio'), axis=1)
        df_merge['Observação'] = df_merge.apply(observacao, axis=1)

        df_merge = df_merge.rename(columns={
            'Ultimo_Preco': 'Último Preço Pago (R$)',
            'Data_Ultima_Compra': 'Data Última Compra',
            'Fornecedor_Ultima_Compra': 'Fornecedor Última Compra',
            'Preco_Medio': 'Preço Médio (R$)',
            'Preco_Minimo': 'Preço Mín. Histórico (R$)',
            'Preco_Maximo': 'Preço Máx. Histórico (R$)',
        })

        for col in ['Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
                    'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)']:
            df_merge[col] = df_merge[col].apply(lambda v: "" if pd.isna(v) else v)

        colunas_exatas = [
            'Item', 'Código', 'Descrição', 'Unidade', 'Qtd',
            'Fornecedor Cotado', 'Valor Cotado (R$)',
            'Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
            'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)',
            'Var. vs Último (%)', 'Var. vs Médio (%)', 'Observação'
        ]
        df_final = df_merge[colunas_exatas]

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
    tema_aggrid = 'alpine-dark' if st.session_state['tema'] == 'Escuro' else 'alpine'

    # Garante que nenhum texto fique cortado: cabeçalho e célula quebram
    # linha e crescem em altura (em vez de truncar com "..."), e a largura
    # inicial de cada coluna é recalculada para caber no conteúdo.
    autosize_ao_carregar = JsCode("""
        function(params) {
            var ids = [];
            var cols = (params.api.getColumns ? params.api.getColumns() : params.columnApi.getAllColumns());
            cols.forEach(function(c) { ids.push(c.getId()); });
            if (params.api.autoSizeColumns) { params.api.autoSizeColumns(ids, false); }
            else if (params.columnApi && params.columnApi.autoSizeColumns) { params.columnApi.autoSizeColumns(ids, false); }
        }
    """)

    gb = GridOptionsBuilder.from_dataframe(df_grid)
    gb.configure_default_column(
        sortable=True, filter=True, resizable=True,
        wrapText=True, autoHeight=True,
        wrapHeaderText=True, autoHeaderHeight=True,
    )
    gb.configure_grid_options(onFirstDataRendered=autosize_ao_carregar)
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

    pdf_bytes = gerar_pdf(df_final)
    placeholder_pdf.download_button(
        label="📥 PDF",
        data=pdf_bytes,
        file_name="mapa_de_cotacao_suprimentos.pdf",
        mime="application/pdf",
        key="btn_pdf_top"
    )
    xlsx_bytes = gerar_excel(df_final)
    placeholder_xlsx.download_button(
        label="📊 Excel",
        data=xlsx_bytes,
        file_name="Comparativo_Cotacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_xlsx_top"
    )
elif uploaded_cot is not None:
    st.warning("⚠️ Nenhum item válido encontrado no arquivo carregado.")
else:
    st.info("👆 Clique na caixa **⚙️ Abrir / Fechar Configurações** acima e faça o upload do seu Mapa de Cotação (.csv, .xlsx, .docx ou .mhtml) caso queira analisar um mapa completo.")

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
            st.warning(f"⚠️ Nenhuma compra anterior encontrada no histórico para o código **{codigo_pesquisa}**.")
        else:
            st.success(f"Foram encontradas **{len(registros)}** ocorrência(s) de compra para o código **{codigo_pesquisa}**:")

            df_historico_item = pd.DataFrame({
                'Data Emissão PC': registros['Data Emissao'].dt.strftime('%d/%m/%Y'),
                'Fornecedor da Compra': registros['Nome Fornece'],
                'Status Aprov': registros['Status Aprov'],
                'Quantidade': registros['Quantidade'],
                'Prc Unitario': registros['Prc Unitario'].apply(formatar_brl),
            })
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
