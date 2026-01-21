# -*- coding: utf-8 -*-
"""
Extrator XML -> Planilha (IBS/CBS)
- Faz upload de 1+ XML (NFe/NFCe) e opcionalmente uma planilha modelo (.xlsx)
- Extrai: Data, Número da Nota, Item/Serviço, cClassTrib, Base (vBC), vIBS, vCBS, arquivo, Fonte do valor
- Grava na aba "LANCAMENTOS" preservando fórmulas/validações existentes (Excel recalcula ao abrir)

Como rodar:
  python -m pip install -r requirements.txt
  python -m streamlit run app.py
"""
import io
import zipfile
from datetime import datetime, date
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# -----------------------------
# Page config + CSS (Figma-like)
# -----------------------------
st.set_page_config(page_title="Extrator XML - IBS/CBS", layout="wide")

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root{
  --bg:#f6f8fc;
  --card:#ffffff;
  --ink:#0f172a;
  --muted:#64748b;
  --line:#e6eaf2;
  --shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
  --radius:16px;
}

html, body, [class*="css"]  { font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
.stApp { background: var(--bg); }

/* page width */
.block-container{ max-width: 1180px; padding-top: 28px; padding-bottom: 42px; }

/* sidebar */
section[data-testid="stSidebar"]{
  background: #0b1220;
  border-right: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] *{ color: rgba(255,255,255,0.92); }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3{ color: #fff; }
section[data-testid="stSidebar"] .stButton button,
section[data-testid="stSidebar"] .stDownloadButton button{
  background: rgba(255,255,255,0.08);
  border: 1px solid rgba(255,255,255,0.14);
}
section[data-testid="stSidebar"] .stFileUploader{
  background: rgba(255,255,255,0.06);
  border: 1px solid rgba(255,255,255,0.12);
  border-radius: 14px;
  padding: 10px 12px;
}

/* headings */
.h-title{
  display:flex; align-items:center; gap:12px;
  margin: 2px 0 2px 0;
}
.h-badge{
  width: 34px; height: 34px; border-radius: 12px;
  background: #eef2ff;
  border: 1px solid var(--line);
  display:flex; align-items:center; justify-content:center;
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
}
.h-badge svg{ width: 18px; height: 18px; opacity: .85; }
.h-sub{ color: var(--muted); margin: 0 0 18px 46px; font-size: 0.98rem; }

/* generic card */
.card{
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 16px 18px;
}
.card + .card{ margin-top: 12px; }

.kpi-grid{ display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 14px; }
@media(max-width: 1200px){ .kpi-grid{ grid-template-columns: repeat(2, minmax(0, 1fr)); } }
@media(max-width: 650px){ .kpi-grid{ grid-template-columns: 1fr; } }

.kpi{
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 16px;
  box-shadow: var(--shadow);
  padding: 16px 18px;
  position: relative;
  overflow:hidden;
}
.kpi:before{
  content:"";
  position:absolute; left:0; top:0; bottom:0; width: 4px;
  background: #cbd5e1;
}
.kpi.kpi-ibs:before{ background:#2563eb; }
.kpi.kpi-cbs:before{ background:#16a34a; }
.kpi.kpi-cred:before{ background:#f59e0b; }
.kpi.kpi-total:before{ background:#a855f7; }

.kpi .label{ color: var(--muted); font-size: 0.88rem; margin-bottom: 4px; display:flex; align-items:center; justify-content:space-between; gap: 10px; }
.kpi .value{ color: var(--ink); font-size: 1.6rem; font-weight: 700; letter-spacing: -0.02em; }
.kpi .sub{ color: var(--muted); font-size: 0.86rem; margin-top: 2px; }

.panel-title{ display:flex; align-items:flex-start; gap: 10px; margin-bottom: 8px; }
.panel-title h3{ margin:0; font-size: 1.05rem; }
.panel-title .hint{ color: var(--muted); font-size: 0.86rem; margin-top: 2px; }
.icon{
  width: 28px; height: 28px; border-radius: 10px;
  border: 1px solid var(--line);
  display:flex; align-items:center; justify-content:center;
  background: #f8fafc;
}
.icon svg{ width: 16px; height: 16px; opacity:.9; }

.bar-row{ margin-top: 12px; }
.bar-label{ display:flex; justify-content:space-between; align-items:center; font-size: 0.92rem; color: var(--muted); margin-bottom: 6px; }
.bar-track{ height: 10px; background:#eef2f7; border-radius: 999px; overflow:hidden; border: 1px solid #e7ebf3;}
.bar-fill{ height:100%; border-radius: 999px; }
.bar-fill.ibs{ background:#2563eb; }
.bar-fill.cbs{ background:#16a34a; }
.bar-fill.cred{ background:#f59e0b; }
.bar-foot{ display:flex; justify-content:space-between; align-items:center; margin-top: 10px; padding-top: 10px; border-top:1px solid var(--line); }
.bar-foot strong{ font-size: 0.95rem; }
.badge-money{ font-weight: 700; }

.hr{ height:1px; background: var(--line); margin: 18px 0; }

/* dataframe look */
.stDataFrame { background: white; border-radius: 14px; border: 1px solid var(--line); overflow:hidden; }

/* inputs row spacing */
div[data-testid="stHorizontalBlock"] > div{ padding-right: 8px; }

/* === Force light canvas === */
.stApp{ background: #f6f8fc !important; }

/* === FIX: titles visible on light background === */
h1, h2, h3, h4, h5, h6,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] h4,
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4{
  color: #0f172a !important;
}
.h-sub, .small-muted, .subtitle, .stCaption, .stMarkdown p, p{
  color: #64748b !important;
}

</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

# -----------------------------
# XML helpers
# -----------------------------
def _local(tag: str) -> str:
    # "{ns}Tag" -> "Tag"
    return tag.split("}", 1)[-1] if "}" in tag else tag

def _find_text(elem: ET.Element, path: str) -> str | None:
    x = elem.find(path)
    if x is None or x.text is None:
        return None
    return x.text.strip()

def _parse_date(root: ET.Element) -> date | None:
    """
    Tenta pegar data de emissão:
      - NFe/infNFe/ide/dhEmi (ISO datetime) ou dEmi (YYYY-MM-DD)
    """
    for p in [
        ".//{*}infNFe/{*}ide/{*}dhEmi",
        ".//{*}infNFe/{*}ide/{*}dEmi",
        ".//{*}ide/{*}dhEmi",
        ".//{*}ide/{*}dEmi",
    ]:
        t = _find_text(root, p)
        if not t:
            continue
        try:
            # dhEmi pode ser "2026-01-08T10:22:33-03:00"
            if "T" in t:
                # remove timezone para parse mais simples
                base = t.split("T")[0]
                return datetime.fromisoformat(base).date() if len(base) > 10 else datetime.fromisoformat(t[:19]).date()
            return datetime.fromisoformat(t).date()
        except Exception:
            try:
                return datetime.strptime(t[:10], "%Y-%m-%d").date()
            except Exception:
                pass
    return None

def _parse_nnf(root: ET.Element) -> str | None:
    # Número da NF: ide/nNF
    for p in [".//{*}infNFe/{*}ide/{*}nNF", ".//{*}ide/{*}nNF"]:
        t = _find_text(root, p)
        if t:
            return t
    return None

def _parse_items_from_xml(xml_bytes: bytes, filename: str) -> list[dict]:
    """
    Extrai itens (det) e IBS/CBS:
      - Item/Serviço: det/prod/xProd
      - cClassTrib: imposto/IBSCBS/cClassTrib
      - Base (vBC): imposto/IBSCBS/vBC
      - vIBS / vCBS: imposto/IBSCBS/vIBS, vCBS (se existirem)
    """
    try:
        root = ET.fromstring(xml_bytes)
    except Exception:
        return []

    emissao = _parse_date(root)
    nnf = _parse_nnf(root)

    rows: list[dict] = []
    dets = root.findall(".//{*}infNFe/{*}det") or root.findall(".//{*}det")
    for det in dets:
        xprod = _find_text(det, ".//{*}prod/{*}xProd") or ""
        ibscbs = det.find(".//{*}imposto/{*}IBSCBS")
        if ibscbs is None:
            # alguns XML podem não ter IBSCBS -> ignora item
            continue

        cclass = _find_text(ibscbs, ".//{*}cClassTrib") or ""
        vbc = _find_text(ibscbs, ".//{*}vBC")
        vibs = _find_text(ibscbs, ".//{*}vIBS")
        vcbs = _find_text(ibscbs, ".//{*}vCBS")

        def _to_float(x: str | None):
            try:
                return float(x) if x not in (None, "") else None
            except Exception:
                return None

        vbc_f = _to_float(vbc)
        vibs_f = _to_float(vibs)
        vcbs_f = _to_float(vcbs)

        # Fonte do valor (base)
        fonte = "IBSCBS/vBC" if vbc_f is not None else ""

        rows.append(
            {
                "Data": emissao,
                "Numero": nnf,
                "Item/Serviço": xprod,
                "cClassTrib": cclass,
                "Valor da operação": vbc_f,
                "vIBS": vibs_f,
                "vCBS": vcbs_f,
                "arquivo": filename,
                "Fonte do valor": fonte,
            }
        )

    return rows

# -----------------------------
# Excel write helper
# -----------------------------
def _append_to_workbook(template_bytes: bytes, df: pd.DataFrame) -> bytes:
    """
    Abre o template e grava df na aba LANCAMENTOS, acrescentando linhas.
    Mantém fórmulas e formatações existentes.
    """
    bio = io.BytesIO(template_bytes)
    wb = load_workbook(bio)
    if "LANCAMENTOS" not in wb.sheetnames:
        # fallback: tenta primeira aba
        ws = wb.active
    else:
        ws = wb["LANCAMENTOS"]

    # Descobre cabeçalho na primeira linha: mapeia nomes para colunas
    header_row = 1
    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col).value
        if isinstance(v, str) and v.strip():
            headers[v.strip()] = col

    # Campos que vamos preencher (se existirem)
    fields = ["Data", "Numero", "Item/Serviço", "cClassTrib", "Valor da operação", "vIBS", "vCBS", "arquivo", "Fonte do valor"]

    # próxima linha vazia (considera que a planilha pode ter fórmulas/linhas em branco)
    next_row = ws.max_row + 1
    # tenta achar a última linha com algo na coluna "Data" (se existir)
    if "Data" in headers:
        c = headers["Data"]
        r = ws.max_row
        while r >= 2 and ws.cell(row=r, column=c).value in (None, ""):
            r -= 1
        next_row = max(r + 1, 2)

    # escreve as linhas
    for _, row in df.iterrows():
        for f in fields:
            if f not in headers:
                continue
            col = headers[f]
            val = row.get(f, None)

            cell = ws.cell(row=next_row, column=col)
            # datas: escreve como date
            if f == "Data" and pd.notna(val) and isinstance(val, date):
                cell.value = val
                cell.number_format = "yyyy-mm-dd"
            else:
                # pandas NaN -> None
                if pd.isna(val):
                    val = None
                cell.value = val
        next_row += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# -----------------------------
# UI
# -----------------------------
st.markdown(
    """
<div class="h-title">
  <div class="h-badge" aria-hidden="true">
    <svg viewBox="0 0 24 24" fill="none">
      <path d="M7 3h7l3 3v15a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" stroke="#334155" stroke-width="1.6"/>
      <path d="M14 3v4a1 1 0 0 0 1 1h4" stroke="#334155" stroke-width="1.6"/>
      <path d="M8 12h8M8 16h8" stroke="#334155" stroke-width="1.6" stroke-linecap="round"/>
    </svg>
  </div>
  <div>
    <h1 style="margin:0; font-size: 2.05rem; letter-spacing:-0.02em;">Extrator XML - IBS/CBS</h1>
  </div>
</div>
<div class="h-sub">Visualização de dados fiscais da reforma tributária</div>
<div class="hr"></div>
""",
    unsafe_allow_html=True,
)

# Sidebar: uploads
with st.sidebar:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Fonte de dados")
    st.caption("Planilha modelo (.xlsx) — opcional")
    planilha_file = st.file_uploader("Arraste e solte aqui", type=["xlsx"], label_visibility="collapsed")
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    st.caption("XML(s) — envie 1 ou mais (pode ser .xml ou .zip com xml dentro)")
    xml_files = st.file_uploader("XML(s)", type=["xml", "zip"], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

# Carrega planilha modelo: se não vier, usa placeholder (exige upload para gerar)
if planilha_file is None:
    st.info("Envie uma **planilha modelo** (.xlsx) na lateral para eu inserir os dados na aba **LANCAMENTOS** mantendo suas fórmulas.")
    template_bytes = None
else:
    template_bytes = planilha_file.read()

# Parse XMLs
rows_all: list[dict] = []
errors: list[str] = []
if xml_files:
    for f in xml_files:
        try:
            b = f.read()
            if f.name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(b)) as z:
                    xml_names = [n for n in z.namelist() if n.lower().endswith(".xml")]
                    if not xml_names:
                        errors.append(f"{f.name}: zip sem .xml")
                        continue
                    for xn in xml_names:
                        xb = z.read(xn)
                        rows = _parse_items_from_xml(xb, f"{f.name}:{xn}")
                        if not rows:
                            errors.append(f"{f.name}:{xn}: não encontrei itens com IBSCBS")
                        rows_all.extend(rows)
            else:
                rows = _parse_items_from_xml(b, f.name)
                if not rows:
                    errors.append(f"{f.name}: não encontrei itens com IBSCBS")
                rows_all.extend(rows)
        except Exception as e:
            errors.append(f"{f.name}: erro ao ler ({e})")

df = pd.DataFrame(rows_all)

# Normaliza Data
if not df.empty:
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date

# ---------- KPIs ----------
def money(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "R$ 0,00"
    try:
        return "R$ {:,.2f}".format(float(x)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def pct(x):
    try:
        return f"{float(x):.1f}%"
    except Exception:
        return ""


# --- Totais vindos da PLANILHA ---
ibs_total = 0.0
cbs_total = 0.0

if "Valor IBS" in df.columns:
    ibs_total = (
        df["Valor IBS"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
        .sum()
    )

if "Valor CBS" in df.columns:
    cbs_total = (
        df["Valor CBS"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
        .sum()
    )

base_total = float(df["Valor da operação"].fillna(0).sum()) if not df.empty and "Valor da operação" in df.columns else 0.0
total_tributos = ibs_total + cbs_total



ibs_aliq = (ibs_total / base_total * 100) if base_total else 0.0
cbs_aliq = (cbs_total / base_total * 100) if base_total else 0.0
creditos_total = 0.0  # se no futuro você tiver créditos no XML, dá para preencher aqui
st.markdown(
    f"""
<div class="kpi-grid">
  <div class="kpi kpi-ibs">
    <div class="label">
      <span>IBS Total</span>
      <span style="color: var(--muted); font-weight:600;">Alíquota: {pct(ibs_aliq)}</span>
    </div>
    <div class="value">{money(ibs_total)}</div>
    <div class="sub">Somatório de vIBS</div>
  </div>

  <div class="kpi kpi-cbs">
    <div class="label">
      <span>CBS Total</span>
      <span style="color: var(--muted); font-weight:600;">Alíquota: {pct(cbs_aliq)}</span>
    </div>
    <div class="value">{money(cbs_total)}</div>
    <div class="sub">Somatório de vCBS</div>
  </div>

  <div class="kpi kpi-cred">
    <div class="label">
      <span>Créditos</span>
      <span style="color: var(--muted); font-weight:600;">IBS + CBS</span>
    </div>
    <div class="value">{money(creditos_total)}</div>
    <div class="sub">Quando existir no XML, aparece aqui</div>
  </div>

  <div class="kpi kpi-total">
    <div class="label">
      <span>Total Tributos</span>
      <span style="color: var(--muted); font-weight:600;">Consolidado</span>
    </div>
    <div class="value">{money(total_tributos)}</div>
    <div class="sub">IBS + CBS (sem créditos)</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# Painéis (estilo Figma) — Débitos vs Créditos
c1, c2 = st.columns(2, gap="large")
ibs_deb = float(ibs_total or 0.0)
cbs_deb = float(cbs_total or 0.0)
ibs_cred = 0.0
cbs_cred = 0.0

def _bar_width(val, vmax):
    if vmax <= 0:
        return "0%"
    return f"{max(0.0, min(1.0, val / vmax)) * 100:.1f}%"

max_ibs = max(ibs_deb, ibs_cred, 1e-9)
max_cbs = max(cbs_deb, cbs_cred, 1e-9)

with c1:
    st.markdown(
        f"""
<div class="card">
  <div class="panel-title">
    <div class="icon" aria-hidden="true">
      <svg viewBox="0 0 24 24" fill="none">
        <path d="M4 18V6" stroke="#334155" stroke-width="1.7" stroke-linecap="round"/>
        <path d="M4 18h16" stroke="#334155" stroke-width="1.7" stroke-linecap="round"/>
        <path d="M8 14l3-3 3 2 4-5" stroke="#2563eb" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </div>
    <div>
      <h3>IBS - Débitos vs Créditos</h3>
      <div class="hint">Imposto sobre Bens e Serviços (Estados/Municípios)</div>
    </div>
  </div>

  <div class="bar-row">
    <div class="bar-label"><span>Débitos</span><span class="badge-money">{money(ibs_deb)}</span></div>
    <div class="bar-track"><div class="bar-fill ibs" style="width:{_bar_width(ibs_deb, max_ibs)}"></div></div>
  </div>

  <div class="bar-row">
    <div class="bar-label"><span>Créditos</span><span class="badge-money">-{money(ibs_cred)}</span></div>
    <div class="bar-track"><div class="bar-fill cred" style="width:{_bar_width(ibs_cred, max_ibs)}"></div></div>
  </div>

  <div class="bar-foot">
    <strong>Saldo a Recolher</strong>
    <span class="badge-money" style="color:#2563eb;">{money(ibs_deb - ibs_cred)}</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

with c2:
    st.markdown(
        f"""
<div class="card">
  <div class="panel-title">
    <div class="icon" aria-hidden="true">
      <svg viewBox="0 0 24 24" fill="none">
        <path d="M4 18V6" stroke="#334155" stroke-width="1.7" stroke-linecap="round"/>
        <path d="M4 18h16" stroke="#334155" stroke-width="1.7" stroke-linecap="round"/>
        <path d="M8 14l3-3 3 2 4-5" stroke="#16a34a" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    </div>
    <div>
      <h3>CBS - Débitos vs Créditos</h3>
      <div class="hint">Contribuição sobre Bens e Serviços (União)</div>
    </div>
  </div>

  <div class="bar-row">
    <div class="bar-label"><span>Débitos</span><span class="badge-money">{money(cbs_deb)}</span></div>
    <div class="bar-track"><div class="bar-fill cbs" style="width:{_bar_width(cbs_deb, max_cbs)}"></div></div>
  </div>

  <div class="bar-row">
    <div class="bar-label"><span>Créditos</span><span class="badge-money">-{money(cbs_cred)}</span></div>
    <div class="bar-track"><div class="bar-fill cred" style="width:{_bar_width(cbs_cred, max_cbs)}"></div></div>
  </div>

  <div class="bar-foot">
    <strong>Saldo a Recolher</strong>
    <span class="badge-money" style="color:#16a34a;">{money(cbs_deb - cbs_cred)}</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

# Alerts
if errors:
    st.warning("Alguns arquivos tiveram problemas:")
    for e in errors[:10]:
        st.write("•", e)
    if len(errors) > 10:
        st.caption(f"... e mais {len(errors)-10} itens")

# ---------- Filters + table ----------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown("## Itens do Documento")
st.caption("Detalhamento dos itens extraídos do XML (inclui base vBC e valores de IBS/CBS quando presentes).")

if df.empty:
    st.info("Envie XML(s) para visualizar os itens aqui.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

c1, c2, c3 = st.columns([1, 2, 1], gap="large")

with c1:
    min_d = df["Data"].min()
    max_d = df["Data"].max()
    # SEMPRE define "periodo" (evita NameError)
    periodo = st.date_input("Período", value=(min_d, max_d), min_value=min_d, max_value=max_d)

with c2:
    q = st.text_input("Buscar item", placeholder="Ex.: produto, serviço, descrição...")

with c3:
    classes = sorted([c for c in df["cClassTrib"].dropna().unique().tolist() if str(c).strip() != ""])
    pick = st.selectbox("cClassTrib", options=["(Todos)"] + classes, index=0)

df_view = df.copy()

# filtro de período (robusto)
if isinstance(periodo, (list, tuple)) and len(periodo) == 2:
    d1, d2 = periodo
    df_view["Data"] = pd.to_datetime(df_view["Data"], errors="coerce").dt.date
    df_view = df_view[(df_view["Data"] >= d1) & (df_view["Data"] <= d2)]

# busca
if q:
    qq = q.strip().lower()
    df_view = df_view[df_view["Item/Serviço"].fillna("").str.lower().str.contains(qq, na=False)]

# cClassTrib
if pick and pick != "(Todos)":
    df_view = df_view[df_view["cClassTrib"].astype(str) == str(pick)]

show_cols = ["Data", "Numero", "Item/Serviço", "cClassTrib", "Valor da operação", "vIBS", "vCBS", "arquivo", "Fonte do valor"]
show_cols = [c for c in show_cols if c in df_view.columns]

st.dataframe(df_view[show_cols], use_container_width=True, hide_index=True, height=420)

st.download_button(
    "Baixar CSV filtrado",
    data=df_view[show_cols].to_csv(index=False).encode("utf-8"),
    file_name="itens_filtrados.csv",
    mime="text/csv",
)
st.markdown("</div>", unsafe_allow_html=True)

# ---------- Generate planilha ----------
st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
st.markdown("## Gerar planilha preenchida")

if template_bytes is None:
    st.info("Para gerar a planilha, envie a **planilha modelo** (.xlsx) na lateral.")
else:
    if st.button("Gerar planilha", type="primary"):
        out_bytes = _append_to_workbook(template_bytes, df_view)
        st.success("Planilha gerada! Abra no Excel para ver as fórmulas calculando.")
        st.download_button(
            "Baixar planilha_preenchida.xlsx",
            data=out_bytes,
            file_name="planilha_preenchida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
