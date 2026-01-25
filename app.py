import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from dateutil import parser as dtparser
from lxml import etree
import zipfile
import re
import os

st.set_page_config(page_title="Auditor Fiscal NFC-e", page_icon="🛡️", layout="wide")

# -----------------------------
# Utils
# -----------------------------
def norm_txt(s):
    if s is None:
        return ""
    s = str(s).lower().strip()
    repl = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüçñ", "aaaaaeeeeiiiiooooouuuucn")
    s = s.translate(repl)
    for ch in ["\n", "\t", " ", ".", ",", ";", ":", "-", "_", "/", "\\", "(", ")", "[", "]", "{", "}", "%", "º"]:
        s = s.replace(ch, " ")
    return " ".join(s.split())

def to_float(x):
    """Converte valores numéricos preservando casas decimais.

    - XML (NF-e/NFC-e) normalmente vem com decimal em ponto: 43.00
    - Excel/BR pode vir como: 1.234,56 ou 43,00
    - Não remove '.' cegamente (isso quebrava o XML e virava 4300).
    """
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)

    s = str(x).strip()
    if not s:
        return np.nan

    # remove símbolos comuns
    s = s.replace("R$", "").replace("\u00a0", " ").strip()
    s = s.replace(" ", "")

    if "," in s and "." in s:
        # assume formato BR "1.234,56"
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        # "123,45"
        s = s.replace(",", ".")
    else:
        # "43.00" (XML) ou "4300"
        pass

    try:
        return float(s)
    except:
        return np.nan


def to_date(x):
    if pd.isna(x):
        return pd.NaT
    if isinstance(x, (datetime, pd.Timestamp)):
        return pd.to_datetime(x).date()
    s = str(x).strip()
    if not s:
        return pd.NaT
    try:
        return dtparser.parse(s, dayfirst=True).date()
    except:
        return pd.NaT

def round2(x):
    if pd.isna(x):
        return np.nan
    return float(np.round(x, 2))

def detect_centavos(df, cols):
    """
    Heurística para ADM/FLEX:
    Se a mediana é "grande" e quase tudo é inteiro, assume que está em centavos.
    """
    for c in cols:
        if c in df.columns:
            s = df[c].dropna()
            if not s.empty:
                try:
                    arr = s.astype(float)
                    med = float(np.nanmedian(arr))
                    if med > 1000 and ((arr % 1) == 0).mean() > 0.9:
                        df[c] = df[c] / 100
                except:
                    pass
    return df

def excel_download(df_dict):
    """Gera Excel 'premium' (dashboard + tabela estilizada + filtros + congela painéis)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        workbook = writer.book

        # =========================
        # Formatos
        # =========================
        fmt_title = workbook.add_format({
            "bold": True, "font_size": 16, "font_color": "#1F4E79"
        })
        fmt_sub = workbook.add_format({
            "bold": True, "font_size": 11, "font_color": "#1F4E79"
        })
        fmt_kpi = workbook.add_format({
            "bold": True, "font_size": 14, "align": "center", "valign": "vcenter",
            "bg_color": "#E8F1FB", "border": 1
        })
        fmt_kpi_lbl = workbook.add_format({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": "#F3F6FB", "border": 1
        })

        fmt_header = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": "#1F4E79",
            "align": "center", "valign": "vcenter", "border": 1
        })
        fmt_text = workbook.add_format({"valign": "vcenter"})
        fmt_wrap = workbook.add_format({"valign": "vcenter", "text_wrap": True})
        fmt_num = workbook.add_format({"num_format": "#,##0.00", "valign": "vcenter"})
        fmt_int = workbook.add_format({"num_format": "0", "valign": "vcenter"})
        fmt_date = workbook.add_format({"num_format": "yyyy-mm-dd", "valign": "vcenter"})

        fmt_ok = workbook.add_format({"font_color": "#006100", "bg_color": "#C6EFCE"})
        fmt_warn = workbook.add_format({"font_color": "#9C5700", "bg_color": "#FFEB9C"})
        fmt_bad = workbook.add_format({"font_color": "#9C0006", "bg_color": "#FFC7CE"})
        fmt_div = workbook.add_format({"font_color": "#7F3F00", "bg_color": "#FCE4D6"})

        zebra = workbook.add_format({"bg_color": "#F7F7F7"})

        # =========================
        # Dashboard (Resumo)
        # =========================
        resumo_name = "RESUMO"
        ws = workbook.add_worksheet(resumo_name)
        writer.sheets[resumo_name] = ws

        ws.set_default_row(18)
        ws.set_column(0, 0, 3)
        ws.set_column(1, 1, 40)
        ws.set_column(2, 6, 18)

        ws.write(0, 1, "Auditoria NFC-e — Resumo", fmt_title)
        ws.write(2, 1, "KPIs", fmt_sub)

        def safe_len(df):
            try:
                return int(len(df))
            except:
                return 0

        sefaz_n = safe_len(df_dict.get("SEFAZ", pd.DataFrame()))
        adm_n = safe_len(df_dict.get("ADM", pd.DataFrame()))
        flex_n = safe_len(df_dict.get("FLEX", pd.DataFrame()))
        alerts_df = df_dict.get("ALERTAS", pd.DataFrame())
        alerts_n = safe_len(alerts_df)

        # KPI cards
        kpis = [("Notas SEFAZ", sefaz_n), ("Notas ADM", adm_n), ("Notas FLEX", flex_n), ("Alertas", alerts_n)]
        row = 4
        col = 1
        for i, (lbl, val) in enumerate(kpis):
            ws.write(row, col + i, lbl, fmt_kpi_lbl)
            ws.write(row + 1, col + i, val, fmt_kpi)

        # Quebras por motivo / status
        ws.write(7, 1, "Distribuição de Alertas", fmt_sub)
        ws.set_column(1, 1, 45)
        ws.set_column(2, 2, 14)

        start_row = 9
        if alerts_df is not None and not alerts_df.empty:
            # Motivos
            motivos = alerts_df["motivo"].fillna("SEM_MOTIVO").value_counts().head(20)
            ws.write(start_row, 1, "Motivo (Top 20)", fmt_header)
            ws.write(start_row, 2, "Qtde", fmt_header)
            for r, (k, v) in enumerate(motivos.items(), start=1):
                ws.write(start_row + r, 1, str(k), fmt_wrap)
                ws.write(start_row + r, 2, int(v), fmt_int)
            # Status
            st_row = start_row
            st_col = 4
            ws.set_column(st_col, st_col, 22)
            ws.set_column(st_col + 1, st_col + 1, 12)
            ws.write(st_row, st_col, "Status ADM", fmt_header)
            ws.write(st_row, st_col + 1, "Qtde", fmt_header)
            for r, (k, v) in enumerate(alerts_df["status_adm"].fillna("SEM").value_counts().items(), start=1):
                ws.write(st_row + r, st_col, str(k), fmt_text)
                ws.write(st_row + r, st_col + 1, int(v), fmt_int)

            ws.write(st_row + 6, st_col, "Status FLEX", fmt_header)
            ws.write(st_row + 6, st_col + 1, "Qtde", fmt_header)
            for r, (k, v) in enumerate(alerts_df["status_flex"].fillna("SEM").value_counts().items(), start=1):
                ws.write(st_row + 6 + r, st_col, str(k), fmt_text)
                ws.write(st_row + 6 + r, st_col + 1, int(v), fmt_int)
        else:
            ws.write(start_row, 1, "Sem alertas gerados.", fmt_text)

        # =========================
        # Abas de dados (Tabela bonita)
        # =========================
        def apply_table(sheet, df):
            ws = writer.sheets[sheet]

            # Congelar: cabeçalho + 3 colunas (data/serie/numero) quando existir
            freeze_col = 0
            for c in ["data", "serie", "numero"]:
                if c in [norm_txt(x) for x in df.columns]:
                    freeze_col = max(freeze_col, [norm_txt(x) for x in df.columns].index(c) + 1)
            ws.freeze_panes(1, freeze_col)

            # Filtro
            if len(df.columns) > 0:
                ws.autofilter(0, 0, max(0, len(df)), max(0, len(df.columns) - 1))

            # Cabeçalho
            ws.set_row(0, 20)
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)

            # Tabela do Excel (estilo)
            nrows = len(df) + 1
            ncols = len(df.columns)
            if ncols > 0:
                ws.add_table(0, 0, max(0, nrows - 1), max(0, ncols - 1), {
                    "style": "Table Style Medium 9",
                    "columns": [{"header": c} for c in df.columns],
                    "autofilter": True
                })

            # Ajuste de colunas + formatos
            for col_idx, col_name in enumerate(df.columns):
                s = df[col_name]
                cname = norm_txt(col_name)

                sample = s.astype(str).head(200).tolist()
                max_len = max([len(str(col_name))] + [len(x) for x in sample if x is not None])
                width = min(max(10, max_len + 2), 55)

                col_fmt = fmt_text
                if "motivo" == cname:
                    col_fmt = fmt_wrap
                    width = max(width, 35)
                elif "data" in cname:
                    col_fmt = fmt_date
                    width = max(width, 12)
                elif any(k in cname for k in ["serie", "numero", "cnf", "id"]) and s.dropna().apply(lambda x: str(x).isdigit()).mean() > 0.8:
                    col_fmt = fmt_int
                elif any(k in cname for k in ["valor", "base", "icms"]):
                    col_fmt = fmt_num

                ws.set_column(col_idx, col_idx, width, col_fmt)

            # Condicionais em status
            for st_col in [c for c in df.columns if norm_txt(c).startswith("status_")]:
                cidx = df.columns.get_loc(st_col)
                ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"OK","format":fmt_ok})
                ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"MULTIPLO","format":fmt_warn})
                ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"DIVERGE","format":fmt_warn})
                ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"NAO_ENCONTRADO","format":fmt_bad})
                ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"SEM_ARQUIVO","format":fmt_bad})

            # Motivo destacado
            for c in df.columns:
                if norm_txt(c) == "motivo":
                    cidx = df.columns.get_loc(c)
                    ws.conditional_format(1, cidx, max(1, len(df)), cidx, {"type":"text","criteria":"containing","value":"Divergência","format":fmt_div})

            # Zebra suave por linhas
            ws.conditional_format(1, 0, max(1, len(df)), max(0, len(df.columns) - 1), {
                "type": "formula",
                "criteria": "=MOD(ROW(),2)=0",
                "format": zebra
            })

            # Agrupamento de colunas (visual): SEFAZ | ADM | FLEX (quando existir)
            cols_norm = [norm_txt(c) for c in df.columns]
            def group(prefixes, level=1, hidden=False):
                for p in prefixes:
                    if p in cols_norm:
                        i = cols_norm.index(p)
                        ws.set_column(i, i, None, None, {"level": level, "hidden": hidden})
            # Colunas auxiliares de data_*
            group(["data_adm","data_flex"], level=1, hidden=True)

        # Escreve abas
        for name, df in df_dict.items():
            sheet = name[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)
            apply_table(sheet, df)

        # Cores das abas
        for sheet, color in [("SEFAZ", "#D9E1F2"), ("ADM", "#E2EFDA"), ("FLEX", "#FFF2CC"), ("ALERTAS", "#FCE4D6")]:
            if sheet in writer.sheets:
                writer.sheets[sheet].set_tab_color(color)

    return output.getvalue()


def find_col(df, ideas):
    cols_norm = {c: norm_txt(c) for c in df.columns}
    for want in ideas:
        w = norm_txt(want)
        for c, cn in cols_norm.items():
            if w and w in cn:
                return c
    return None

def safe_get_col(df, ideas, required=True, default=np.nan):
    col = find_col(df, ideas)
    if col is None:
        if required:
            raise KeyError(
                f"Não encontrei coluna parecida com {ideas}. "
                f"Colunas disponíveis: {list(df.columns)}"
            )
        else:
            return pd.Series([default] * len(df), index=df.index)
    return df[col]

def normalize_key(x):
    """
    Normaliza série/número para bater entre Excel e XML:
    - remove .0 (quando vem como float)
    - mantém só dígitos
    - remove zeros à esquerda (000123 -> 123)
    """
    if pd.isna(x):
        return ""
    s = str(x).strip()

    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]

    try:
        if isinstance(x, (float, np.floating)) and float(x).is_integer():
            s = str(int(x))
    except:
        pass

    dig = re.sub(r"\D+", "", s)
    if dig == "":
        return s

    dig2 = dig.lstrip("0")
    return dig2 if dig2 != "" else "0"

def looks_like_bad_header(cols):
    cols = list(cols)
    if len(cols) == 0:
        return True
    if sum(str(c).lower().startswith("unnamed") for c in cols) >= max(1, int(len(cols) * 0.6)):
        return True
    if len(cols) >= 3 and all(str(c).strip() == "" for c in cols):
        return True
    if len(cols) == 1 and len(str(cols[0])) > 25:
        return True
    if sum(isinstance(c, (int, float, datetime, pd.Timestamp, np.number)) for c in cols) >= max(1, int(len(cols) * 0.6)):
        return True
    return False

def smart_read_excel(uploaded_file):
    df1 = pd.read_excel(uploaded_file, sheet_name=0)
    if not looks_like_bad_header(df1.columns):
        return df1

    df0 = pd.read_excel(uploaded_file, sheet_name=0, header=None)

    expected = [
        "data", "data venda", "data movto", "emissao", "dt emissao", "data emissao",
        "serie", "série", "numero", "número", "vlr", "vlr total", "vlr. total",
        "valor", "total", "base", "base icms", "icms", "cfop", "aliquota", "alíquota"
    ]
    expected = [norm_txt(x) for x in expected]

    best_i = None
    best_score = -1
    max_rows = min(60, len(df0))

    for i in range(max_rows):
        row = df0.iloc[i].tolist()
        row_norm = [norm_txt(x) for x in row]
        score = 0
        for cell in row_norm:
            for e in expected:
                if e and e in cell:
                    score += 1
        has_data = any("data" in c for c in row_norm)
        has_key = any(("serie" in c) or ("numero" in c) or ("nnf" in c) or ("n nf" in c) for c in row_norm)
        if has_data and has_key:
            score += 3
        if score > best_score:
            best_score = score
            best_i = i

    if best_i is None or best_score < 2:
        df0.columns = [f"col_{i}" for i in range(df0.shape[1])]
        return df0

    header = df0.iloc[best_i].tolist()
    header = [str(h).strip() if not pd.isna(h) else "" for h in header]
    header = [h if h else f"col_{idx}" for idx, h in enumerate(header)]

    df = df0.iloc[best_i + 1:].copy()
    df.columns = header
    df = df.reset_index(drop=True)
    df = df.dropna(how="all")
    return df

# -----------------------------
# XML SEFAZ
# -----------------------------
def parse_xml(xml_bytes: bytes, filename: str | None = None):
    root = etree.fromstring(xml_bytes)

    def xtext(node, xpath):
        el = node.xpath(xpath)
        if not el:
            return None
        return el[0].text if hasattr(el[0], "text") else str(el[0])

    infNFe = root.xpath("//*[local-name()='infNFe']")
    if not infNFe:
        return None
    infNFe = infNFe[0]

    ide_nodes = infNFe.xpath(".//*[local-name()='ide']")
    if not ide_nodes:
        return None
    ide = ide_nodes[0]

    serie = xtext(ide, ".//*[local-name()='serie']")
    numero = xtext(ide, ".//*[local-name()='nNF']")
    dhEmi = xtext(ide, ".//*[local-name()='dhEmi']") or xtext(ide, ".//*[local-name()='dEmi']")
    data = to_date(dhEmi)

    total = infNFe.xpath(".//*[local-name()='ICMSTot']")
    vNF = vBC = vICMS = None
    if total:
        total = total[0]
        vNF = to_float(xtext(total, ".//*[local-name()='vNF']"))
        vBC = to_float(xtext(total, ".//*[local-name()='vBC']"))
        vICMS = to_float(xtext(total, ".//*[local-name()='vICMS']"))

    # SEFAZ (XML) já vem em reais -> NÃO divide por 100 aqui

    # Regra: arquivos XML cujo nome começa com '11' são NFC-e canceladas (padrão informado)
    cancelada = False
    if filename:
        base_name = os.path.basename(str(filename))
        cancelada = base_name.startswith('11')

    return {
        "data": data,
        "serie": normalize_key(serie),
        "numero": normalize_key(numero),
        "valor": round2(vNF),
        "base": round2(vBC),
        "icms": round2(vICMS),
        "fonte": "SEFAZ",
        "cancelada": cancelada,
    }

def load_sefaz_from_upload(uploaded):
    rows = []

    def add_xml_bytes(b, fname=None):
        r = parse_xml(b, filename=fname)
        if r:
            rows.append(r)

    for f in uploaded:
        name = (f.name or "").lower()
        content = f.getvalue()

        if name.endswith(".zip"):
            try:
                with zipfile.ZipFile(BytesIO(content), "r") as z:
                    for n in z.namelist():
                        if n.lower().endswith(".xml"):
                            add_xml_bytes(z.read(n), fname=n)
            except:
                pass
        elif name.endswith(".xml"):
            add_xml_bytes(content, fname=f.name)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["serie"] = df["serie"].astype(str).map(normalize_key)
    df["numero"] = df["numero"].astype(str).map(normalize_key)
    return df

# -----------------------------
# Standardizers
# -----------------------------
def standardize_adm(df):
    out = pd.DataFrame()

    out["data"] = safe_get_col(df, ["data venda", "data movto", "data"], required=True).apply(to_date)

    serie_col = find_col(df, ["serie", "série"])
    if serie_col:
        out["serie"] = df[serie_col].map(normalize_key)
    else:
        out["serie"] = "1"

    out["numero"] = safe_get_col(df, ["numero", "número", "nnf", "n nf"], required=True).map(normalize_key)

    out["valor"] = safe_get_col(df, ["vlr total", "vlr. total", "valor", "total", "vlt total"], required=True).apply(to_float)

    # opcionais (ADM normalmente não tem)
    out["base"] = safe_get_col(df, ["base", "vbc", "base icms"], required=False, default=np.nan).apply(to_float)
    out["icms"] = safe_get_col(df, ["icms", "vicms"], required=False, default=np.nan).apply(to_float)
    # Fiscal ADM vem em centavos -> converte para reais
    for c in ["valor", "base", "icms"]:
        out[c] = out[c] / 100
    for c in ["valor", "base", "icms"]:
        out[c] = out[c].apply(round2)

    out["fonte"] = "ADM"
    return out

def standardize_flex(df):
    out = pd.DataFrame()

    out["data"] = safe_get_col(df, ["data", "emissao", "dt emissao", "data emissao"], required=True).apply(to_date)

    serie_col = find_col(df, ["serie", "série"])
    if serie_col:
        out["serie"] = df[serie_col].map(normalize_key)
    else:
        out["serie"] = "1"

    out["numero"] = safe_get_col(df, ["numero", "número", "nfce", "nota", "nnf", "n nf"], required=True).map(normalize_key)

    out["valor"] = safe_get_col(df, ["valor", "total", "liquido", "líquido"], required=True).apply(to_float)
    out["base"]  = safe_get_col(df, ["base icms", "base", "vbc"], required=False, default=np.nan).apply(to_float)
    out["icms"]  = safe_get_col(df, ["icms", "vicms"], required=False, default=np.nan).apply(to_float)

    # FLEX pode vir em centavos
    out = detect_centavos(out, ["valor", "base", "icms"])
    for c in ["valor", "base", "icms"]:
        out[c] = out[c].apply(round2)

    grp = out.groupby(["data", "serie", "numero"], as_index=False)[["valor", "base", "icms"]].sum()
    grp["fonte"] = "FLEX"
    return grp

# -----------------------------
# Auditoria
# -----------------------------
def audit(sefaz, adm, flex):
    '''
    Auditoria focada em VALORES (ignora DATA como chave principal).

    Match principal: Série + Número
    - Se existir em ADM/FLEX com a mesma Série+Número, considera a nota "existente"
    - Se houver múltiplas linhas com a mesma Série+Número (datas diferentes), escolhe a melhor
      comparando os valores com SEFAZ (menor diferença), e marca status MULTIPLO.

    Alertas:
    - NAO_ENCONTRADO (não existe por Série+Número)
    - Divergência de VALOR / BASE / ICMS (quando existe)
    '''
    if sefaz.empty:
        return pd.DataFrame(columns=[
            "data","serie","numero",
            "status_adm","status_flex","motivo",
            "valor_sefaz","base_sefaz","icms_sefaz",
            "data_adm","serie_adm","valor_adm","base_adm","icms_adm",
            "data_flex","serie_flex","valor_flex","base_flex","icms_flex",
        ])

    k_key = ["serie", "numero"]

    # agrupa por Série+Número
    def build_group(df):
        if df.empty:
            return {}
        return { k: sub for k, sub in df.groupby(k_key) }

    adm_g  = build_group(adm)
    flex_g = build_group(flex)

    alerts = []

    def cmp(a, b):
        if pd.isna(a) and pd.isna(b):
            return True
        return np.isclose(a, b, atol=0.01, equal_nan=True)

    def pick_best(sub, r_sefaz):
        """Escolhe a linha mais provável quando há duplicidade (datas diferentes)."""
        if sub is None or len(sub) == 0:
            return None, "NAO_ENCONTRADO"
        if len(sub) == 1:
            return sub.iloc[0], "OK"

        # Score por proximidade dos valores (prioriza VALOR)
        def score(row):
            s = 0.0
            # valor sempre pesa mais
            if pd.notna(r_sefaz.get("valor", np.nan)) and pd.notna(row.get("valor", np.nan)):
                s += abs(float(r_sefaz["valor"]) - float(row["valor"])) * 100
            # base/icms se existirem
            if pd.notna(r_sefaz.get("base", np.nan)) and pd.notna(row.get("base", np.nan)):
                s += abs(float(r_sefaz["base"]) - float(row["base"]))
            if pd.notna(r_sefaz.get("icms", np.nan)) and pd.notna(row.get("icms", np.nan)):
                s += abs(float(r_sefaz["icms"]) - float(row["icms"]))
            # bônus se data bate (não é chave, mas ajuda a escolher)
            if str(row.get("data", "")) == str(r_sefaz.get("data", "")):
                s -= 0.5
            return s

        best_idx = sub.apply(score, axis=1).astype(float).idxmin()
        return sub.loc[best_idx], "MULTIPLO"

    def add_alert(r_sefaz, status_adm, status_flex, motivo, r_adm=None, r_flex=None):
        alerts.append({
            "data": r_sefaz.get("data", np.nan),
            "serie": r_sefaz.get("serie", np.nan),
            "numero": r_sefaz.get("numero", np.nan),

            "status_adm": status_adm,
            "status_flex": status_flex,
            "motivo": motivo,

            "valor_sefaz": r_sefaz.get("valor", np.nan),
            "base_sefaz": r_sefaz.get("base", np.nan),
            "icms_sefaz": r_sefaz.get("icms", np.nan),

            "data_adm":  (r_adm.get("data", np.nan)  if r_adm is not None else np.nan),
            "serie_adm": (r_adm.get("serie", np.nan) if r_adm is not None else np.nan),
            "valor_adm": (r_adm.get("valor", np.nan) if r_adm is not None else np.nan),
            "base_adm":  (r_adm.get("base", np.nan)  if r_adm is not None else np.nan),
            "icms_adm":  (r_adm.get("icms", np.nan)  if r_adm is not None else np.nan),

            "data_flex":  (r_flex.get("data", np.nan)  if r_flex is not None else np.nan),
            "serie_flex": (r_flex.get("serie", np.nan) if r_flex is not None else np.nan),
            "valor_flex": (r_flex.get("valor", np.nan) if r_flex is not None else np.nan),
            "base_flex":  (r_flex.get("base", np.nan)  if r_flex is not None else np.nan),
            "icms_flex":  (r_flex.get("icms", np.nan)  if r_flex is not None else np.nan),
        })

    for _, r in sefaz.iterrows():
        key = (r["serie"], r["numero"])

        adm_sub  = adm_g.get(key)
        flex_sub = flex_g.get(key)

        r_adm, st_adm = pick_best(adm_sub, r)
        r_flex, st_flex = pick_best(flex_sub, r)

        # Existência
        status_adm = "SEM_ARQUIVO" if adm.empty else st_adm
        status_flex = "SEM_ARQUIVO" if flex.empty else st_flex

        if status_adm in ["SEM_ARQUIVO", "NAO_ENCONTRADO"] or status_flex in ["SEM_ARQUIVO", "NAO_ENCONTRADO"]:
            motivos = []
            if status_adm == "SEM_ARQUIVO":
                motivos.append("ADM não carregou")
            elif status_adm == "NAO_ENCONTRADO":
                motivos.append("Nota NÃO encontrada no ADM (Série+Número)")
            elif status_adm == "MULTIPLO":
                motivos.append("ADM: múltiplas linhas (datas diferentes)")

            if status_flex == "SEM_ARQUIVO":
                motivos.append("FLEX não carregou")
            elif status_flex == "NAO_ENCONTRADO":
                motivos.append("Nota NÃO encontrada no FLEX (Série+Número)")
            elif status_flex == "MULTIPLO":
                motivos.append("FLEX: múltiplas linhas (datas diferentes)")

            add_alert(r, status_adm, status_flex, " | ".join(motivos), r_adm=r_adm, r_flex=r_flex)
            continue

        # Agora: existe. Se MULTIPLO, avisa mas continua comparando valores
        if status_adm == "MULTIPLO" or status_flex == "MULTIPLO":
            motivos = []
            if status_adm == "MULTIPLO":
                motivos.append("ADM: múltiplas linhas (escolhida a mais próxima por valores)")
            if status_flex == "MULTIPLO":
                motivos.append("FLEX: múltiplas linhas (escolhida a mais próxima por valores)")
            add_alert(r, status_adm, status_flex, " | ".join(motivos), r_adm=r_adm, r_flex=r_flex)

        # Divergência VALOR
        if r_adm is not None and not cmp(r.get("valor", np.nan), r_adm.get("valor", np.nan)):
            add_alert(r, status_adm, status_flex, "Divergência de VALOR (ADM)", r_adm=r_adm, r_flex=r_flex)

        if r_flex is not None and not cmp(r.get("valor", np.nan), r_flex.get("valor", np.nan)):
            add_alert(r, status_adm, status_flex, "Divergência de VALOR (FLEX)", r_adm=r_adm, r_flex=r_flex)

        # Divergência BASE/ICMS (só se existir dos dois lados)
        if r_adm is not None and pd.notna(r_adm.get("base", np.nan)) and pd.notna(r.get("base", np.nan)):
            if not cmp(r["base"], r_adm["base"]):
                add_alert(r, status_adm, status_flex, "Divergência de BASE (ADM)", r_adm=r_adm, r_flex=r_flex)

        if r_adm is not None and pd.notna(r_adm.get("icms", np.nan)) and pd.notna(r.get("icms", np.nan)):
            if not cmp(r["icms"], r_adm["icms"]):
                add_alert(r, status_adm, status_flex, "Divergência de ICMS (ADM)", r_adm=r_adm, r_flex=r_flex)

        if r_flex is not None and pd.notna(r_flex.get("base", np.nan)) and pd.notna(r.get("base", np.nan)):
            if not cmp(r["base"], r_flex["base"]):
                add_alert(r, status_adm, status_flex, "Divergência de BASE (FLEX)", r_adm=r_adm, r_flex=r_flex)

        if r_flex is not None and pd.notna(r_flex.get("icms", np.nan)) and pd.notna(r.get("icms", np.nan)):
            if not cmp(r["icms"], r_flex["icms"]):
                add_alert(r, status_adm, status_flex, "Divergência de ICMS (FLEX)", r_adm=r_adm, r_flex=r_flex)

    df_alerts = pd.DataFrame(alerts)
    if not df_alerts.empty:
        df_alerts = df_alerts.drop_duplicates().sort_values(["serie","numero","motivo"])
    return df_alerts

# -----------------------------
# UI



def normalize_table(df: pd.DataFrame, fonte: str, dividir_por_100: bool=False) -> pd.DataFrame:
    """Normaliza qualquer tabela (CSV/Excel) para o formato padrão:
    data, serie, numero, valor, base, icms, fonte
    """
    if df is None or len(df)==0:
        return pd.DataFrame(columns=["data","serie","numero","valor","base","icms","fonte"])
    cols = {norm_txt(c): c for c in df.columns}

    def pick(primary, fallbacks):
        for k in [primary]+fallbacks:
            if k in cols:
                return cols[k]
        return None

    c_data = pick("data", ["dt", "data venda", "data movto", "data_movto", "dtemi", "dhEmi".lower()])
    c_serie = pick("serie", ["série", "ser", "serie_nf", "serie nfe", "serie nfce"])
    c_num = pick("numero", ["número", "num", "n°", "nnf", "nf", "numero nf", "numero nota"])
    c_val = pick("valor", ["vlr total", "vltotal", "valor total", "vl tot", "vl total", "valor nota", "vnf", "v_nf", "total"])
    c_base = pick("base", ["base icms", "vlbcicms", "bc icms", "vbc", "v_bc", "baseicms"])
    c_icms = pick("icms", ["valor icms", "vlicms", "icms total", "vl icms", "vicms", "v_icms"])

    out = pd.DataFrame()
    out["data"] = df[c_data].astype(str).str[:10] if c_data is not None else ""
    out["serie"] = df[c_serie] if c_serie is not None else np.nan
    out["numero"] = df[c_num] if c_num is not None else np.nan

    def to_num(series):
        s = series.astype(str).str.strip()
        # remove moeda e espaços
        s = s.str.replace("R$", "", regex=False).str.replace(" ", "", regex=False)
        # pt-BR -> float
        s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce")

    out["valor"] = to_num(df[c_val]) if c_val is not None else np.nan
    out["base"] = to_num(df[c_base]) if c_base is not None else np.nan
    out["icms"] = to_num(df[c_icms]) if c_icms is not None else np.nan

    if dividir_por_100:
        out["valor"] = out["valor"] / 100.0
        out["base"] = out["base"] / 100.0
        out["icms"] = out["icms"] / 100.0

    out["fonte"] = fonte

    # limpeza / tipos
    out = out.dropna(subset=["numero"])
    out["serie"] = pd.to_numeric(out["serie"], errors="coerce").fillna(0).astype(int)
    out["numero"] = pd.to_numeric(out["numero"], errors="coerce").astype(int)
    for c in ["valor","base","icms"]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).astype(float)
    return out

def read_csv_any(file_like, fonte: str = "SEFAZ", dividir_por_100: bool = False):
    try:
        file_like.seek(0)
    except Exception:
        pass
    try:
        df = pd.read_csv(file_like, sep=None, engine="python")
    except Exception:
        file_like.seek(0)
        df = pd.read_csv(file_like)
    return normalize_table(df, fonte=fonte, dividir_por_100=dividir_por_100)

def read_sefaz_xml_file(file_like):
    # UploadedFile pode manter ponteiro no fim após reruns; use getvalue/seek
    try:
        content = file_like.getvalue()
    except Exception:
        file_like.seek(0)
        content = file_like.read()
    xml = content.decode("utf-8", errors="ignore") if isinstance(content, (bytes, bytearray)) else str(content)
    if "parse_sefaz_xml_string" in globals():
        return parse_sefaz_xml_string(xml, filename=os.path.basename(path))
    return pd.DataFrame(columns=["data","serie","numero","valor","base","icms","fonte","cancelada"])

def read_sefaz_zip(file_like):
    import zipfile as _zf
    from io import BytesIO as _BytesIO
    try:
        data = file_like.getvalue()
    except Exception:
        file_like.seek(0)
        data = file_like.read()
    buf = _BytesIO(data)
    rows=[]
    with _zf.ZipFile(buf) as z:
        for n in z.namelist():
            if n.lower().endswith(".xml"):
                xml = z.read(n).decode("utf-8", errors="ignore")
                if "parse_sefaz_xml_string" in globals():
                    df = parse_sefaz_xml_string(xml, filename=n)
                    if not df.empty:
                        rows.append(df)
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["data","serie","numero","valor","base","icms","fonte"])

def read_excel_any(file_like, fonte, dividir_por_100=False):
    try:
        file_like.seek(0)
    except Exception:
        pass
    df = pd.read_excel(file_like)
    return normalize_table(df, fonte=fonte, dividir_por_100=dividir_por_100)

def parse_sefaz_xml_string(xml: str, filename: str | None = None) -> pd.DataFrame:
    """Parse de um XML da SEFAZ (texto) e retorna DataFrame normalizado."""
    try:
        row = parse_xml(xml.encode("utf-8"), filename=filename)
        return pd.DataFrame([row])
    except Exception:
        return pd.DataFrame(columns=["data","serie","numero","valor","base","icms","fonte","cancelada"])
def money_br(v):
    try:
        if pd.isna(v):
            return "—"
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"

def render_header():
    st.markdown("""<div class="l-hero">
      <div class="l-hero-left">
        <div class="l-icon">🛡️</div>
        <div>
          <p class="l-title">Auditor Fiscal NFC-e</p>
          <p class="l-sub">Conferência inteligente SEFAZ × Fiscal ADM × Fiscal Flex</p>
        </div>
      </div>
      <div class="l-badge">✅ Sistema Operacional</div>
    </div>""", unsafe_allow_html=True)

def upload_block(col, title, desc, key, border_class, types):
    colors = {"blue":"#2563eb","green":"#059669","amber":"#b45309"}
    with col:
        st.markdown(f"""<div class="upload-card {border_class}">
          <div style="font-size:32px; line-height: 1;">📄</div>
          <div class="upload-title" style="color: {colors.get(border_class,'#111827')};">{title}</div>
          <p class="upload-desc">{desc}</p>
        """, unsafe_allow_html=True)

        up = st.file_uploader(" ", type=types, key=key, label_visibility="collapsed")
        if up is not None:
            st.markdown(f'<span class="small-pill ok">✅ {up.name}</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span class="small-pill">⬆️ Clique ou arraste o arquivo</span>', unsafe_allow_html=True)
        st.markdown("""</div>""", unsafe_allow_html=True)
        return up


def calc_metrics(sefaz_df: pd.DataFrame, adm_df: pd.DataFrame, flex_df: pd.DataFrame, alerts_df: pd.DataFrame) -> dict:
    """
    Painel (cards) baseado no dataframe final (alerts_df), porque é ele que alimenta a tabela.
    Assim, mesmo quando ADM/FLEX não forem carregados, os totais do SEFAZ aparecem.
    """
    al = alerts_df.copy() if isinstance(alerts_df, pd.DataFrame) else pd.DataFrame()
    sef = sefaz_df.copy() if isinstance(sefaz_df, pd.DataFrame) else pd.DataFrame()

    def _sum_numeric(df: pd.DataFrame, col: str) -> float:
        if df.empty or col not in df.columns:
            return 0.0
        s = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return float(s.sum())

    def _count_notes(df: pd.DataFrame) -> int:
        if df.empty:
            return 0
        if {"serie","numero"}.issubset(df.columns):
            return int(df[["serie","numero"]].dropna().drop_duplicates().shape[0])
        return int(len(df))

    # Totais do SEFAZ (preferir colunas da tabela final)
    total_valor_sefaz = _sum_numeric(al, "valor_sefaz") or _sum_numeric(sef, "valor")
    total_icms_apurado = _sum_numeric(al, "icms_sefaz") or _sum_numeric(sef, "icms")
    total_notas = _count_notes(al[al.get("valor_sefaz").notna()] if (not al.empty and "valor_sefaz" in al.columns) else al) or _count_notes(sef)

    # Contagens de status na tabela final
    conferidas = divergentes = ausentes_adm = ausentes_flex = 0
    if not al.empty:
        if "status_adm" in al.columns:
            ausentes_adm = int(al["status_adm"].astype(str).isin(["SEM_ARQUIVO","NAO_ENCONTRADO"]).sum())
        if "status_flex" in al.columns:
            ausentes_flex = int(al["status_flex"].astype(str).isin(["SEM_ARQUIVO","NAO_ENCONTRADO"]).sum())

        ok_adm = (al["status_adm"].astype(str) == "OK") if "status_adm" in al.columns else pd.Series([False]*len(al))
        ok_flex = (al["status_flex"].astype(str) == "OK") if "status_flex" in al.columns else pd.Series([False]*len(al))
        no_div = ~al["motivo"].astype(str).str.contains("Diverg", case=False, na=False) if "motivo" in al.columns else pd.Series([True]*len(al))

        conferidas = int((ok_adm & ok_flex & no_div).sum())
        divergentes = int(al["motivo"].astype(str).str.contains("Diverg", case=False, na=False).sum()) if "motivo" in al.columns else 0

    return {
        # nomes "canônicos" (interno)
        "total_valor_sefaz": total_valor_sefaz,
        "total_icms_apurado": total_icms_apurado,

        # nomes usados no painel (UI)
        "valor_total": total_valor_sefaz,
        "icms_total": total_icms_apurado,

        "total_notas": total_notas,
        "conferidas": conferidas,
        "divergentes": divergentes,
        "ausentes_adm": ausentes_adm,
        "ausentes_flex": ausentes_flex,
    }


def style_table(df):
    if df is None or df.empty:
        return df
    d = df.copy()

    def row_style(r):
        motivo = str(r.get("motivo",""))
        if "Divergência" in motivo:
            return ["background-color: #fff1f2"] * len(r)
        if r.get("status_adm") == "NAO_ENCONTRADO" or r.get("status_flex") == "NAO_ENCONTRADO":
            return ["background-color: #fffbeb"] * len(r)
        return [""] * len(r)

    sty = d.style.apply(row_style, axis=1)

    for col in [c for c in d.columns if str(c).startswith("status_")]:
        sty = sty.applymap(
            lambda v: (
                "color:#065f46;background:#d1fae5;border-radius:999px;padding:2px 8px;display:inline-block;font-weight:800;"
                if str(v)=="OK" else
                "color:#7c2d12;background:#ffedd5;border-radius:999px;padding:2px 8px;display:inline-block;font-weight:800;"
                if ("DIVERGE" in str(v) or "MULTIPLO" in str(v)) else
                "color:#7f1d1d;background:#fee2e2;border-radius:999px;padding:2px 8px;display:inline-block;font-weight:900;"
                if ("NAO_ENCONTRADO" in str(v) or "SEM_ARQUIVO" in str(v)) else ""
            ),
            subset=[col]
        )
    return sty

render_header()
st.write("")
st.markdown('<div class="l-card"><h3 style="margin:0;font-weight:900;">Importar Dados</h3><p style="margin:6px 0 0 0;color:rgba(17,24,39,.6);font-size:13px;">Carregue os arquivos das três fontes para iniciar a conferência</p></div>', unsafe_allow_html=True)
st.write("")

c1, c2, c3 = st.columns(3, gap="large")
up_sefaz = upload_block(c1, "Arquivo SEFAZ", "XML (zip), CSV ou Excel da Secretaria da Fazenda", "up_sefaz", "blue", ["xml","zip","csv","xlsx","xls"])
up_adm   = upload_block(c2, "Fiscal ADM", "Planilha Excel do sistema Fiscal ADM", "up_adm", "green", ["xlsx","xls"])
up_flex  = upload_block(c3, "Fiscal Flex", "Planilha Excel do sistema Fiscal Flex", "up_flex", "amber", ["xlsx","xls"])

sefaz_df = pd.DataFrame()
adm_df = pd.DataFrame()
flex_df = pd.DataFrame()
alerts = pd.DataFrame()

if up_sefaz is not None:
    name = up_sefaz.name.lower()
    if name.endswith(".csv"):
        sefaz_df = read_csv_any(up_sefaz, fonte="SEFAZ", dividir_por_100=False)
    elif name.endswith(".zip"):
        sefaz_df = read_sefaz_zip(up_sefaz)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        sefaz_df = read_excel_any(up_sefaz, fonte="SEFAZ", dividir_por_100=False)
    else:
        sefaz_df = read_sefaz_xml_file(up_sefaz)


if up_adm is not None:
    adm_df = read_excel_any(up_adm, fonte="ADM", dividir_por_100=True)

if up_flex is not None:
    flex_df = read_excel_any(up_flex, fonte="FLEX", dividir_por_100=False)

if up_sefaz is not None and (sefaz_df is None or sefaz_df.empty):
    st.warning("Arquivo SEFAZ carregado, mas não consegui identificar as colunas (preciso de Nº/Série/Valor). Se for CSV, confirme se tem colunas como numero/serie/valor/base/icms.")
# Auditoria (ignora data, foca valores)
if not sefaz_df.empty and "cancelada" in sefaz_df.columns:
    sefaz_df = sefaz_df[~sefaz_df["cancelada"].fillna(False)]

if not sefaz_df.empty:
    alerts = audit(sefaz_df, adm_df, flex_df)

st.write("")

m = calc_metrics(sefaz_df, adm_df, flex_df, alerts)
# Fallback: se por algum motivo o DataFrame SEFAZ não ficou disponível neste ciclo,
# calculamos os cards a partir da própria tabela de alertas (que já contém valor/base/icms da SEFAZ).
if (m.get("total_notas", 0) == 0) and (not alerts.empty) and ("valor_sefaz" in alerts.columns):
    sef = alerts[alerts["valor_sefaz"].notna()].copy()
    # chaves únicas por (serie, numero) para não contar duplicado
    if "serie" in sef.columns and "numero" in sef.columns:
        sef["_k"] = sef["serie"].astype(str) + "-" + sef["numero"].astype(str)
        total_notas = sef["_k"].nunique()
    else:
        total_notas = len(sef)
    valor_total = pd.to_numeric(sef["valor_sefaz"], errors="coerce").fillna(0).sum()
    icms_total = pd.to_numeric(sef.get("icms_sefaz", 0), errors="coerce").fillna(0).sum()

    # conferidas = OK nos 3
    conferidas = 0
    if "status_adm" in alerts.columns and "status_flex" in alerts.columns:
        conferidas = int(((alerts["status_adm"] == "OK") & (alerts["status_flex"] == "OK") & (~alerts["motivo"].astype(str).str.contains("Diverg", case=False, na=False))).sum())

    divergentes = 0
    if "motivo" in alerts.columns:
        divergentes = int(alerts["motivo"].astype(str).str.contains("Diverg", case=False, na=False).sum())

    aus_adm = 0
    if "status_adm" in alerts.columns:
        aus_adm = int(alerts["status_adm"].isin(["NAO_ENCONTRADO", "SEM_ARQUIVO"]).sum())

    aus_flex = 0
    if "status_flex" in alerts.columns:
        aus_flex = int(alerts["status_flex"].isin(["NAO_ENCONTRADO", "SEM_ARQUIVO"]).sum())

    m.update({
        "valor_total": float(valor_total),
        "icms_total": float(icms_total),
        "total_notas": int(total_notas),
        "conferidas": int(conferidas),
        "divergentes": int(divergentes),
        "ausentes_adm": int(aus_adm),
        "ausentes_flex": int(aus_flex),
    })


sum1, sum2 = st.columns(2, gap="large")
with sum1:
    st.markdown(f'<div class="grad bluepurp"><p class="grad-label">Valor Total SEFAZ</p><p class="grad-value">{money_br(m.get("valor_total"))}</p></div>', unsafe_allow_html=True)
with sum2:
    st.markdown(f'<div class="grad green"><p class="grad-label">Total ICMS Apurado</p><p class="grad-value">{money_br(m.get("icms_total"))}</p></div>', unsafe_allow_html=True)

k1,k2,k3,k4,k5 = st.columns(5, gap="large")

def kpi(col, title, num, sub, variant, emoji):
    with col:
        st.markdown(f'''
        <div class="kpi {variant}">
          <div>
            <h4>{title}</h4>
            <div class="num">{num}</div>
            <div class="sub">{sub}</div>
          </div>
          <div class="dot">{emoji}</div>
        </div>
        ''', unsafe_allow_html=True)

kpi(k1,"TOTAL DE NOTAS", m.get("total_notas", 0), "Notas analisadas", "neutral","🧾")
kpi(k2,"CONFERIDAS", m.get("conferidas", 0), "Compatíveis", "success","✅")
kpi(k3,"DIVERGENTES", m.get("divergentes", 0), "Valores diferentes", "danger","⛔")
kpi(k4,"AUSENTES ADM", m.get("ausentes_adm", 0), "Faltando no ADM", "warn","⚠️")
kpi(k5,"AUSENTES FLEX", m.get("ausentes_flex", 0), "Faltando no Flex", "warn","⚠️")

st.write("")
st.markdown('<div class="l-card"><h3 style="margin:0;font-weight:900;">Resultado da Conferência</h3><p style="margin:6px 0 0 0;color:rgba(17,24,39,.6);font-size:13px;">Clique e filtre para ver os detalhes completos</p></div>', unsafe_allow_html=True)
st.write("")

tabs = ["Todos", "Conferidos", "Divergentes", "Ausentes ADM", "Ausentes Flex"]
counts = {"Todos": m.get("total_notas", 0), "Conferidos": m.get("conferidas", 0), "Divergentes": m.get("divergentes", 0), "Ausentes ADM": m.get("ausentes_adm", 0), "Ausentes Flex": m.get("ausentes_flex", 0)}
choice = st.radio(" ", [f"{t}  ({counts[t]})" for t in tabs], horizontal=True, label_visibility="collapsed")
choice_key = choice.split("  (")[0]

search = st.text_input("Buscar (número, série ou motivo)", placeholder="Ex.: 15186, 201, Divergência de VALOR", label_visibility="collapsed")

if sefaz_df.empty:
    st.info("Carregue o arquivo SEFAZ para ver os resultados.")
else:
    base = alerts.copy()

    if choice_key == "Divergentes":
        base = base[base["motivo"].fillna("").str.contains("Divergência", case=False)] if not base.empty else base
    elif choice_key == "Ausentes ADM":
        base = base[base.get("status_adm") == "NAO_ENCONTRADO"] if not base.empty else base
    elif choice_key == "Ausentes Flex":
        base = base[base.get("status_flex") == "NAO_ENCONTRADO"] if not base.empty else base
    elif choice_key == "Conferidos":
        if base.empty:
            conf = sefaz_df.copy()
        else:
            problem = base[["serie","numero"]].dropna().drop_duplicates()
            conf = sefaz_df.merge(problem.assign(_p=1), on=["serie","numero"], how="left")
            conf = conf[conf["_p"].isna()].drop(columns=["_p"])
        conf = conf.copy()
        conf["status_adm"] = "OK"
        conf["status_flex"] = "OK"
        conf["motivo"] = "Conferido"
        base = conf

    if search and not base.empty:
        s = search.lower().strip()
        mask = pd.Series(False, index=base.index)
        for col in [c for c in base.columns if c in ["serie","numero","motivo"]]:
            mask = mask | base[col].astype(str).str.lower().str.contains(s, na=False)
        base = base[mask]

    if not base.empty:
        sort_cols = [c for c in ["serie","numero","motivo"] if c in base.columns]
        if sort_cols:
            base = base.sort_values(sort_cols)

    if base.empty:
        st.success("Sem registros nesse filtro ✅")
    else:
        st.dataframe(style_table(base), use_container_width=True, height=520)

st.write("")
d1, d2 = st.columns([1,2])
with d1:
    if not sefaz_df.empty:
        xls = excel_download({"SEFAZ": sefaz_df, "ADM": adm_df, "FLEX": flex_df, "ALERTAS": alerts})
        st.download_button("Baixar relatório", data=xls, file_name="auditoria_nfce.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with d2:
    st.caption("Dica: use os filtros (pílulas) + busca para isolar divergências e ausências rapidamente.")
