"""
Lógica pura (sem nenhuma dependência de Streamlit) do app "Meu Mapa de
Cotação": formatação de valores, normalização de código de produto,
consolidação da base histórica de preços e o comparativo cotação x
histórico (menor preço por item, variação vs. último/médio, observação).

Extraído de app.py para poder ser testado sem precisar de um contexto real
de execução do Streamlit (st.session_state, st.cache_data, st.secrets etc).
Nenhuma linha de lógica de negócio foi alterada nessa extração — só
movida, e onde um trecho vivia solto no meio do script (o comparativo do
mapa de cotação) ele virou uma função (`processar_mapa_cotacao`) com os
mesmos nomes de variável internos.
"""
import unicodedata
import datetime
import io

import pandas as pd
import numpy as np
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ==============================================================================
# Normalização de valores vindos de planilha/upload (moeda, código de produto,
# texto sem acento)
# ==============================================================================
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


def _normalizar_texto(valor: str) -> str:
    s = "".join(c for c in unicodedata.normalize('NFKD', str(valor)) if not unicodedata.combining(c))
    return s.strip().lower()


def _achar_coluna_normalizada(colunas, termos):
    for col in colunas:
        col_norm = _normalizar_texto(col)
        for termo in termos:
            if _normalizar_texto(termo) in col_norm:
                return col
    return None


def extrair_numero_cotacao(cotacao: pd.DataFrame):
    """Acha o número da cotação (ex.: "021132") no mapa de cotação bruto,
    ainda com as colunas originais do arquivo — o export do TOTVS traz uma
    coluna "COTACAO :" com o mesmo número repetido em toda linha. Usado
    pra nomear os arquivos exportados (PDF/Excel) com o número que consta
    no arquivo original, em vez de um nome genérico. Devolve None se não
    achar a coluna ou se ela estiver vazia (quem chama decide o fallback —
    ex.: usar o nome do arquivo enviado)."""
    if cotacao is None or cotacao.empty:
        return None
    col = _achar_coluna_normalizada(cotacao.columns, ['cotacao'])
    if not col:
        return None
    for valor in cotacao[col]:
        if pd.notna(valor) and str(valor).strip():
            return str(valor).strip()
    return None


def normalizar_codigo(valor):
    """Extrai só os dígitos de um código de produto e completa com zeros à
    esquerda até 10 dígitos — o formato padrão do código de Produto no
    TOTVS/Portal Gestão de Compras (ex.: '1268' e '0000001268' viram ambos
    '0000001268', e são tratados como o mesmo item)."""
    if pd.isna(valor):
        return None
    s = ''.join(filter(str.isdigit, str(valor)))
    if not s:
        return None
    return str(int(s)).zfill(10)


# ==============================================================================
# Consolidação da base histórica de preços (parte pura de construir_base_precos
# em app.py — recebe o histórico já lido da planilha, não acessa Google Sheets)
# ==============================================================================
def consolidar_base_precos(df: pd.DataFrame, status_filtro: str = None):
    """
    Regra de negócio (decisão explícita do usuário — "HISTÓRICO É HISTÓRICO"):
    por padrão TODAS as linhas do histórico entram na consolidação,
    independente do Status Aprov (Aprovado, Bloqueado, Rejeitado, Pendente).
    Passe status_filtro='Aprovado' apenas se quiser considerar só compras
    efetivadas em um caso específico.

    Espera que `df` já tenha as colunas internas: Produto, Descricao.1,
    Unidade, Prc Unitario, Data Emissao, Nome Fornece, Quantidade,
    Status Aprov (ver HIST_REQUIRED_COLS/MAPA_COLUNAS_PEDIDOS em app.py).

    Retorna (base_precos, df_f):
        base_precos: DataFrame consolidado por item (1 linha por Cod_Norm)
        df_f: DataFrame do histórico filtrado, linha a linha (usado na
              consulta rápida por código)
    """
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
    # A coluna DATA PEDIDO tinha ~13 mil linhas legadas gravadas em formato
    # americano (M/D/AAAA, ex.: "1/7/2026" = 7 de janeiro) misturadas com
    # linhas corretas em DD/MM/AAAA - por isso o dayfirst=False daqui (o
    # inverso do resto do sistema). Essas linhas legadas foram corrigidas
    # direto na planilha "Pedidos" em 09/09/2026 (confirmado batendo
    # DATA LIBERAÇÃO - DATA PEDIDO >= 0), entao a base agora e 100%
    # DD/MM/AAAA de verdade - dayfirst=True (igual ao Portal Gestão de
    # Compras e ao Painel do Comprador) e o certo dai em diante.
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

    return base_precos, df_f


# ==============================================================================
# Comparativo do mapa de cotação x histórico (parte pura do bloco de
# PROCESSAMENTO DO MAPA DE COTAÇÃO em app.py — recebe a tabela já extraída
# do arquivo de upload, não lida com st.file_uploader/st.progress)
# ==============================================================================
def processar_mapa_cotacao(cotacao: pd.DataFrame, base_precos: pd.DataFrame):
    """
    Em vez de varrer célula a célula do histórico procurando o código
    (O(itens_mapa × linhas_histórico), lento e sujeito a falso-positivo),
    fazemos um merge por código normalizado contra a base_precos já
    consolidada — a mesma lógica de build_comparativo.py da skill. A
    "melhor cotação" por item é sempre o MENOR valor entre as linhas
    daquele código (cobre tanto o mapa já resumido, uma linha por item,
    quanto o mapa bruto multi-fornecedor com código só na primeira linha do
    grupo — usamos ffill no código antes de agrupar).

    Retorna (df_final, aviso_valores_estranhos):
        df_final: DataFrame vazio se `cotacao` estiver vazia ou nenhum item
                   válido for encontrado; senão, uma linha por item cotado
                   com as colunas de comparação (Valor Cotado, Último Preço
                   Pago, Preço Médio/Mín/Máx, variações e Observação).
        aviso_valores_estranhos: True se uma fração grande dos itens tiver
                   variação muito alta e sistemática em relação ao
                   histórico (indício de escala diferente, ex. lote vs.
                   unitário).
    """
    aviso_valores_estranhos = False
    df_final = pd.DataFrame()

    if cotacao.empty:
        return df_final, aviso_valores_estranhos

    cotacao = cotacao.copy()
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

    # Quando a planilha traz TOTAL (valor total da linha, já líquido) e
    # Desconto (valor do desconto concedido) separados do Vl.Unitário
    # bruto, o preço comparável de verdade é o unitário com desconto já
    # aplicado: (TOTAL - Desconto) / Qtde — mesma conta que o comprador já
    # fazia manualmente numa coluna auxiliar no Excel (decisão explícita do
    # usuário: a comparação sempre tem que usar esse valor, nunca o
    # Vl.Unitário bruto, quando essas colunas existirem).
    c_total = achar_coluna(cotacao, ['total'])
    c_desconto = achar_coluna(cotacao, ['desconto'])

    # Prioriza "Razão Social" (nome real do fornecedor) sobre "Fornecedor"
    # quando os dois existem — em export TOTVS, "Fornecedor" costuma ser só
    # o código interno (ex.: "000079"), e "Razão Social" tem o nome de fato
    # (ex.: "LJ GUERRA E CIA LTDA").
    c_forn = achar_coluna(cotacao, ['razão social', 'razao social', 'nome fantasia']) \
        or achar_coluna(cotacao, ['fornecedor', 'empresa', 'nome'])
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

        valor_cotado = 0.0
        if c_total and c_desconto and qtd > 0:
            total_linha = limpar_valor(row[c_total]) if pd.notna(row[c_total]) else 0.0
            desconto_linha = limpar_valor(row[c_desconto]) if pd.notna(row[c_desconto]) else 0.0
            if total_linha > 0:
                valor_cotado = (total_linha - desconto_linha) / qtd

        if valor_cotado <= 0:
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
            'Código': cod_norm,
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

        # Representatividade (%): quanto o valor total da linha (Valor
        # Cotado × Qtd) pesa dentro do valor total da cotação inteira —
        # ajuda o comprador a enxergar quais itens realmente merecem
        # atenção/negociação (poucos itens costumam concentrar a maior
        # parte do valor, ex.: análise ABC/Pareto).
        valor_total_linha = df_merge['Valor Cotado (R$)'] * df_merge['Qtd']
        soma_valor_total = valor_total_linha.sum()
        if soma_valor_total > 0:
            df_merge['Representatividade (%)'] = (valor_total_linha / soma_valor_total * 100).round(2)
        else:
            df_merge['Representatividade (%)'] = 0.0

        colunas_exatas = [
            'Item', 'Código', 'Descrição', 'Unidade', 'Qtd',
            'Fornecedor Cotado', 'Valor Cotado (R$)',
            'Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
            'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)',
            'Var. vs Último (%)', 'Var. vs Médio (%)', 'Observação', 'Representatividade (%)'
        ]
        df_final = df_merge[colunas_exatas]

    return df_final, aviso_valores_estranhos


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
# Exportação — PDF
# ==============================================================================
def gerar_pdf(df, numero_cotacao=None):
    class PDFProfissional(FPDF):
        def __init__(self):
            super().__init__(orientation='L', unit='mm', format='A4')
            self.set_margins(left=5.0, top=19.1, right=5.0)
            self.set_auto_page_break(auto=True, margin=19.1)

        def header(self):
            self.set_fill_color(47, 85, 151)
            self.rect(5.0, 8, 287.0, 20, 'F')

            titulo = "Mapa de Cotacao & Comparativo Historico"
            if numero_cotacao:
                titulo += f" - Cotacao no {numero_cotacao}"
            self.set_font("helvetica", "B", 14)
            self.set_text_color(255, 255, 255)
            self.set_xy(5.0, 10)
            self.cell(287.0, 6, limpar_texto_pdf(titulo), 0, 1, "C")

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

    col_widths = [8, 16, 38, 8, 17, 30, 17, 30, 17, 20, 20, 20, 22, 14]
    headers = [
        "Item", "Codigo", "Descricao", "Qtd",
        "Vl Cotado", "Forn. Cotado", "Ult. Preco",
        "Forn. Ult.", "Preco Med.", "Var vs Med(%)", "Preco Min.", "Preco Max.", "Observacao", "Repr.%"
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
        repr_val = row['Representatividade (%)']
        repr_str = f"{float(repr_val):.1f}%" if repr_val != "" and pd.notna(repr_val) else ""

        pdf.cell(col_widths[0], 6, limpar_texto_pdf(str(row['Item'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[1], 6, limpar_texto_pdf(str(row['Código'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[2], 6, limpar_texto_pdf(str(row['Descrição'])[:26]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[3], 6, limpar_texto_pdf(str(row['Qtd'])), border=1, fill=fill, align="C")
        pdf.cell(col_widths[4], 6, limpar_texto_pdf(formatar_brl(row['Valor Cotado (R$)'])), border=1, fill=fill, align="R")
        pdf.cell(col_widths[5], 6, limpar_texto_pdf(str(row['Fornecedor Cotado'])[:18]), border=1, fill=fill, align="L")
        pdf.cell(col_widths[6], 6, limpar_texto_pdf(ult_preco_str), border=1, fill=fill, align="R")
        pdf.cell(col_widths[7], 6, limpar_texto_pdf(forn_ant_str[:18]), border=1, fill=fill, align="L")
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
        pdf.cell(col_widths[13], 6, limpar_texto_pdf(repr_str), border=1, fill=fill, align="R")

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


def gerar_excel(df: pd.DataFrame, numero_cotacao=None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparativo_Cotacao"

    headers = ['Item', 'Cód. Produto', 'Descrição do Item', 'Unid.', 'Qtd',
               'Fornecedor Cotado', 'Valor Cotado (R$)',
               'Último Preço Pago (R$)', 'Data Última Compra', 'Fornecedor Última Compra',
               'Preço Médio (R$)', 'Preço Mín. Histórico (R$)', 'Preço Máx. Histórico (R$)',
               'Var. vs Último (%)', 'Var. vs Médio (%)', 'Observação', 'Representatividade (%)']

    # Linha de título com o número da cotação é opcional (só quando
    # informado) para não deslocar a posição do cabeçalho no caso comum —
    # header_row segue a posição real em vez de um "1" fixo.
    if numero_cotacao:
        ws.append([f"Mapa de Cotação nº {numero_cotacao}"])
        titulo_row = ws.max_row
        ws.merge_cells(start_row=titulo_row, start_column=1, end_row=titulo_row, end_column=len(headers))
        titulo_cell = ws.cell(row=titulo_row, column=1)
        titulo_cell.font = Font(name=FONT_XLSX, bold=True, size=12)
        titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[titulo_row].height = 22

    ws.append(headers)
    header_row = ws.max_row
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[header_row].height = 36

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
            round(float(r['Representatividade (%)']), 2) if r['Representatividade (%)'] != "" else None,
        ])

    last_row = ws.max_row
    for row in ws.iter_rows(min_row=header_row + 1, max_row=last_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.font = Font(name=FONT_XLSX, size=10)
            cell.border = BORDER
        row[8].number_format = 'DD/MM/YYYY'
        for idx in (6, 7, 10, 11, 12):
            row[idx].number_format = '#,##0.00'
        for idx in (13, 14):
            row[idx].number_format = '+0.0"%";-0.0"%";0.0"%"'
        row[14].fill = _var_fill(row[14].value)
        row[16].number_format = '0.0"%"'

    widths = [7, 11, 42, 7, 7, 32, 15, 16, 16, 32, 15, 15, 15, 14, 14, 24, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{last_row}"

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
