def normalize_table(df: pd.DataFrame, fonte: str, dividir_por_100: bool=False) -> pd.DataFrame:
    """Normaliza qualquer tabela (CSV/Excel) para o formato padrão:
    data, serie, numero, valor, base, icms, fonte

    Robusto para planilhas com cabeçalhos variados: tenta match exato (normalizado)
    e, se não encontrar, usa heurísticas por substring (ex.: 'serie', 'num', 'valor', 'icms').
    """
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=["data","serie","numero","valor","base","icms","fonte"])

    # mapa: nome_normalizado -> nome_original
    cols = {norm_txt(c): c for c in df.columns}

    def pick(primary, fallbacks):
        for k in [primary] + list(fallbacks):
            if k in cols:
                return cols[k]
        return None

    def pick_contains(tokens, prefer_tokens=None):
        """Retorna a primeira coluna cujo nome normalizado contém TODOS tokens.
        prefer_tokens (opcional) prioriza colunas que contenham algum desses tokens."""
        norm_names = list(cols.keys())
        cand = []
        for nn in norm_names:
            ok = True
            for t in tokens:
                if t not in nn:
                    ok = False
                    break
            if ok:
                cand.append(nn)
        if not cand:
            return None
        if prefer_tokens:
            # prioriza quem contém token preferencial
            for pt in prefer_tokens:
                for nn in cand:
                    if pt in nn:
                        return cols[nn]
        return cols[cand[0]]

    # tentativas (match exato)
    c_data = pick("data", ["dt", "data venda", "data movto", "data_movto", "dtemi", "dhemi"])
    c_serie = pick("serie", ["série", "ser", "serie_nf", "serie nfe", "serie nfce", "serie cupom", "serie documento"])
    c_num   = pick("numero", ["número", "num", "n nf", "nnf", "nf", "numero nf", "numero nota", "numero documento", "numero cupom", "no", "nro"])
    c_val   = pick("valor", ["vlr total", "vltotal", "valor total", "vl tot", "vl total", "valor nota", "vnf", "v_nf", "total", "vlcont", "vltotnf"])
    c_base  = pick("base", ["base icms", "vlbcicms", "bc icms", "vbc", "v_bc", "baseicms", "base_calculo", "base calculo"])
    c_icms  = pick("icms", ["valor icms", "vlicms", "icms total", "vl icms", "vicms", "v_icms", "valor_do_icms"])

    # heurísticas (se faltar)
    if c_serie is None:
        c_serie = pick_contains(["serie"])
    if c_num is None:
        c_num = pick_contains(["num"], prefer_tokens=["numero","nnf","nro"]) or pick_contains(["nf"], prefer_tokens=["numero","num"])
    if c_val is None:
        c_val = pick_contains(["valor"], prefer_tokens=["total","vl"]) or pick_contains(["total"])
    if c_icms is None:
        c_icms = pick_contains(["icms"], prefer_tokens=["valor","vl"])
    if c_base is None:
        # base costuma vir como 'base icms' / 'bc icms' / 'vbc'
        c_base = pick_contains(["base"], prefer_tokens=["icms","bc","calculo"]) or pick_contains(["bc"], prefer_tokens=["icms"])

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
    out["base"]  = to_num(df[c_base]) if c_base is not None else np.nan
    out["icms"]  = to_num(df[c_icms]) if c_icms is not None else np.nan

    if dividir_por_100:
        out["valor"] = out["valor"] / 100.0
        out["base"]  = out["base"] / 100.0
        out["icms"]  = out["icms"] / 100.0

    out["fonte"] = fonte

    # limpeza / tipos
    out = out.dropna(subset=["numero"])
    # serie pode vir vazia em algumas planilhas: tenta numérico, senão 0
    out["serie"]  = pd.to_numeric(out["serie"], errors="coerce").fillna(0).astype(int)
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

            st.markdown('<span class="small-pill">⬆️ Clique ou arraste o arquivo</span>', unsafe_allow_html=True)
        st.markdown("""</div>""", unsafe_allow_html=True)
        return up


def calc_metrics(sefaz_df: pd.DataFrame, adm_df: pd.DataFrame, flex_df: pd.DataFrame, alerts_df: pd.DataFrame) -> dict:
    """Calcula KPIs do painel.

    IMPORTANTE:
    - A tabela (alerts_df) só contém *problemas* (ausências/divergências/múltiplos).
      Então "conferidas" não pode ser contada dentro dela.
    - Todos os contadores (conferidas/divergentes/ausentes) devem ser por NOTA (série+número),
      e não por linha, porque uma mesma nota pode gerar mais de um alerta.
    """
    sef = sefaz_df.copy() if isinstance(sefaz_df, pd.DataFrame) else pd.DataFrame()
    al = alerts_df.copy() if isinstance(alerts_df, pd.DataFrame) else pd.DataFrame()

    def _sum_numeric(df: pd.DataFrame, col: str) -> float:
        if df is None or df.empty or col not in df.columns:
            return 0.0
        s = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return float(s.sum())

    def _key_df(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return pd.DataFrame(columns=['serie','numero'])
        if not {'serie','numero'}.issubset(df.columns):
            return pd.DataFrame(columns=['serie','numero'])
        k = df[['serie','numero']].dropna().copy()
        # normaliza tipos pra evitar mismatch 201 vs 201.0
        k['serie'] = k['serie'].astype(str).str.replace('.0','',regex=False).str.strip()
        k['numero'] = k['numero'].astype(str).str.replace('.0','',regex=False).str.strip()
        return k

    # Total de notas: SEMPRE vem do SEFAZ (universo)
    keys_sefaz = _key_df(sef)
    total_notas = int(keys_sefaz.drop_duplicates().shape[0]) if not keys_sefaz.empty else int(len(sef))

    # Totais monetários: preferir SEFAZ direto (é o "universo").
    # Se por algum motivo não existir, tenta pegar das colunas *_sefaz do df final.
    valor_total = _sum_numeric(sef, 'valor') or _sum_numeric(al, 'valor_sefaz')
    icms_total = _sum_numeric(sef, 'icms') or _sum_numeric(al, 'icms_sefaz')

    # Se não há alertas, então tudo conferido (desde que ADM/FLEX carregados).
    keys_alertas = _key_df(al)
    total_problemas = int(keys_alertas.drop_duplicates().shape[0]) if not keys_alertas.empty else 0

    # Divergentes (por nota)
    divergentes = 0
    if not al.empty and 'motivo' in al.columns and {'serie','numero'}.issubset(al.columns):
        mask_div = al['motivo'].astype(str).str.contains('Diverg', case=False, na=False)
        divergentes = int(_key_df(al[mask_div]).drop_duplicates().shape[0])

    # Ausentes (por nota)
    ausentes_adm = 0
    if not al.empty and 'status_adm' in al.columns and {'serie','numero'}.issubset(al.columns):
        mask_aus = al['status_adm'].astype(str).isin(['NAO_ENCONTRADO','SEM_ARQUIVO'])
        ausentes_adm = int(_key_df(al[mask_aus]).drop_duplicates().shape[0])

    ausentes_flex = 0
    if not al.empty and 'status_flex' in al.columns and {'serie','numero'}.issubset(al.columns):
        mask_aus = al['status_flex'].astype(str).isin(['NAO_ENCONTRADO','SEM_ARQUIVO'])
        ausentes_flex = int(_key_df(al[mask_aus]).drop_duplicates().shape[0])

    # Conferidas: total - notas que apareceram em QUALQUER alerta
    conferidas = max(total_notas - total_problemas, 0)

    return {
        # chaves usadas no painel atual
        'valor_total': float(valor_total),
        'icms_total': float(icms_total),
        'total_notas': int(total_notas),
        'conferidas': int(conferidas),
        'divergentes': int(divergentes),
        'ausentes_adm': int(ausentes_adm),
        'ausentes_flex': int(ausentes_flex),

        # aliases (mantém compatibilidade)
        'total_valor_sefaz': float(valor_total),
        'total_icms_apurado': float(icms_total),
    }



def style_table(df):
    if df is None or df.empty:
    # garante chaves (ADM/FLEX podem vir com cabeçalho diferente)
    df = _ensure_serie_numero(df)

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
full_df = pd.DataFrame()

if up_sefaz is not None:
    name = up_sefaz.name.lower()
    if name.endswith(".csv"):
        sefaz_df = read_csv_any(up_sefaz, fonte="SEFAZ", dividir_por_100=False)
    elif name.endswith(".zip"):
        sefaz_df = read_sefaz_zip(up_sefaz)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        sefaz_df = read_excel_any(up_sefaz, fonte="SEFAZ", dividir_por_100=False)

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
    full_df = build_full_table(sefaz_df, adm_df, flex_df, alerts)

st.write("")

m = calc_metrics(sefaz_df, adm_df, flex_df, alerts)
# Se a tabela completa existir, recalcula contadores (mais fiel ao filtro UI)
if isinstance(full_df, pd.DataFrame) and not full_df.empty:
    try:
        m['total_notas'] = int(full_df[['serie','numero']].drop_duplicates().shape[0])
        m['divergentes'] = int(full_df[full_df['motivo'].astype(str).str.contains('diverg', case=False, na=False)][['serie','numero']].drop_duplicates().shape[0])
        m['ausentes_adm'] = int(full_df[full_df['status_adm'].astype(str).isin(['NAO_ENCONTRADO','SEM_ARQUIVO'])][['serie','numero']].drop_duplicates().shape[0])
        m['ausentes_flex'] = int(full_df[full_df['status_flex'].astype(str).isin(['NAO_ENCONTRADO','SEM_ARQUIVO'])][['serie','numero']].drop_duplicates().shape[0])
        m['conferidas'] = int(full_df[full_df['motivo'].astype(str).str.lower().eq('conferido')][['serie','numero']].drop_duplicates().shape[0])
    except Exception:
        pass
# Fallback: se por algum motivo o DataFrame SEFAZ não ficou disponível neste ciclo,
# calculamos os cards a partir da própria tabela de alertas (que já contém valor/base/icms da SEFAZ).
if (m.get("total_notas", 0) == 0) and (not alerts.empty) and ("valor_sefaz" in alerts.columns):
    sef = alerts[alerts["valor_sefaz"].notna()].copy()
    # chaves únicas por (serie, numero) para não contar duplicado
    if "serie" in sef.columns and "numero" in sef.columns:
        sef["_k"] = sef["serie"].astype(str) + "-" + sef["numero"].astype(str)
        total_notas = sef["_k"].nunique()

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

    base = (full_df.copy() if isinstance(full_df, pd.DataFrame) and not full_df.empty else alerts.copy())

    if choice_key == "Divergentes":
        base = base[base["motivo"].fillna("").str.contains("diverg", case=False)] if not base.empty else base
    elif choice_key == "Ausentes ADM":
        base = base[base.get("status_adm").astype(str).isin(["NAO_ENCONTRADO","SEM_ARQUIVO"])] if not base.empty else base
    elif choice_key == "Ausentes Flex":
        base = base[base.get("status_flex").astype(str).isin(["NAO_ENCONTRADO","SEM_ARQUIVO"])] if not base.empty else base
    elif choice_key == "Conferidos":
        # conferidos agora vêm da tabela completa (motivo = Conferido)
        if base.empty:
            base = base

            base = base[base.get("motivo").astype(str).str.lower().eq("conferido")]

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