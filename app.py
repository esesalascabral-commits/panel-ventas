import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
st.set_page_config(
    page_title="Resumen de Ventas",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS
st.markdown("""
<style>
    html, body, [class*="css"], .stApp, .main {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        background-color: #f4f6fa;
    }
    p, span, div, label, h1, h2, h3, h4, h5, h6,
    .stMarkdown, .stMarkdown * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }
    p, span, div, label { font-weight: 600 !important; }
    h1, h2, h3, h4 { font-weight: 800 !important; }

    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div {
        background-color: #ffffff !important;
    }
    section[data-testid="stSidebar"] * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
    }

    div[data-testid="stMetric"] {
        background: white;
        padding: 18px 20px;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        border-left: 4px solid #1f4e79;
    }
    div[data-testid="stMetric"] * {
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        opacity: 1 !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.7rem !important;
        font-weight: 800 !important;
    }
    div[data-testid="stMetricLabel"],
    div[data-testid="stMetricLabel"] * {
        font-size: 0.85rem !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    div[data-testid="stMetricDelta"] * {
        color: inherit !important;
        -webkit-text-fill-color: inherit !important;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        height: 44px; padding: 0 22px; background-color: #eef2f7;
        border-radius: 10px; font-weight: 700 !important;
    }
    .stTabs [data-baseweb="tab"] * {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    .stTabs [aria-selected="true"] { background-color: #1f4e79 !important; }
    .stTabs [aria-selected="true"] * {
        color: white !important;
        -webkit-text-fill-color: white !important;
    }

    .header-box {
        background: linear-gradient(135deg, #1f4e79 0%, #2c7fb8 60%, #4da3d9 100%);
        padding: 28px 36px; border-radius: 16px;
        margin-bottom: 26px;
        box-shadow: 0 4px 12px rgba(31,78,121,0.15);
    }
    .header-box h1 {
        color: white !important; margin: 0;
        font-size: 2.1rem; font-weight: 800;
    }
    .header-box p {
        margin: 8px 0 0; font-size: 0.98rem;
        color: white !important; font-weight: 600;
    }
    .pill {
        display: inline-block; padding: 4px 12px; border-radius: 20px;
        background: rgba(255,255,255,0.25);
        color: white !important;
        font-size: 0.8rem; margin-right: 8px; font-weight: 700;
    }

    .highlight-card {
        background: linear-gradient(135deg, #fff 0%, #f8fafc 100%);
        padding: 20px 24px; border-radius: 14px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        margin-bottom: 16px;
    }
    .highlight-card h3 {
        margin: 0 0 8px; font-size: 0.85rem;
        color: #0f172a !important;
        text-transform: uppercase; letter-spacing: 0.8px; font-weight: 800;
    }
    .highlight-card .big-number {
        font-size: 2rem; font-weight: 800;
        color: #1f4e79 !important;
        line-height: 1.1;
    }
    .highlight-card .sub {
        font-size: 0.9rem; margin-top: 6px;
        color: #0f172a !important; font-weight: 700;
    }

    .alert-box {
        background: #fef2f2;
        border-left: 4px solid #dc2626;
        padding: 14px 18px; border-radius: 10px;
        margin-bottom: 16px;
    }
    .alert-box p {
        margin: 0; color: #991b1b !important; font-weight: 700;
    }

    .success-box {
        background: #f0fdf4;
        border-left: 4px solid #16a34a;
        padding: 14px 18px; border-radius: 10px;
        margin-bottom: 16px;
    }
    .success-box p {
        margin: 0; color: #166534 !important; font-weight: 700;
    }

    .stDataFrame [role="columnheader"], .stDataFrame [role="columnheader"] * {
        color: #0f172a !important;
        font-weight: 800 !important;
    }
    .stDataFrame [role="gridcell"], .stDataFrame [role="gridcell"] * {
        color: #0f172a !important;
        font-weight: 600 !important;
    }

    h5 {
        color: #0f172a !important;
        font-weight: 800 !important;
        font-size: 1.05rem !important;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# PLOTLY LAYOUT BASE
# ============================================================
PLOTLY_LAYOUT = dict(
    font=dict(family="system-ui, sans-serif", size=13, color="#0f172a"),
    plot_bgcolor="white", paper_bgcolor="white",
    xaxis=dict(tickfont=dict(color="#0f172a", size=12),
               title_font=dict(color="#0f172a", size=13),
               gridcolor="#f1f5f9", zeroline=False),
    yaxis=dict(tickfont=dict(color="#0f172a", size=12),
               title_font=dict(color="#0f172a", size=13),
               gridcolor="#f1f5f9", zeroline=False),
)


# ============================================================
# MAPEO DE TERMINALES
# ============================================================
TERMINALES = {
    69510721: "CAPELLI CORDOBA FLEX",
    69510718: "CAPELLI CORDOBA MINI",
    69510722: "CAPELLI SANTIAGO",
    69510723: "CAPELLI ALVARADO",
    69510726: "IDRAET PRO BEAUTY SALTA",
    69510727: "CAPELLI JUJUY - IDRAET TUCUMAN",
    69506644: "SUCURSAL 69506644",
}


# ============================================================
# UTILIDADES
# ============================================================
def formato_pesos(v):
    if pd.isna(v) or v == 0:
        return "$ 0"
    return f"$ {v:,.0f}".replace(",", ".")


def formato_pesos_full(v):
    if pd.isna(v):
        return "$ 0,00"
    return f"$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nombre_terminal(numero):
    try:
        numero = int(numero)
    except (ValueError, TypeError):
        return f"Terminal {numero}"
    return TERMINALES.get(numero, f"Terminal {numero}")


def detectar_encoding(ruta_bytes):
    """Prueba distintos encodings y devuelve el texto."""
    for enc in ["latin-1", "utf-8", "cp1252"]:
        try:
            return ruta_bytes.decode(enc)
        except (UnicodeDecodeError, AttributeError):
            continue
    return ruta_bytes.decode("latin-1", errors="ignore")


def color_diferencia(v):
    """Semáforo según magnitud de diferencia."""
    if pd.isna(v):
        return "⚪"
    a = abs(v)
    if a < 50:
        return "🟢"
    elif a < 500:
        return "🟡"
    else:
        return "🔴"


# ============================================================
# PARSER DE .LIS
# ============================================================
def parsear_lis(contenido, nombre_archivo):
    lineas = contenido.splitlines()
    d = {
        "archivo": nombre_archivo,
        "sucursal": None, "fecha": None, "hora": None,
        "vta_efectivo": 0, "vta_cheques": 0, "vta_mp_tr": 0, "vta_posnet": 0,
        "vta_del_dia": 0, "total_ingresos": 0, "total_egresos": 0,
        "diferencia_caja": 0, "cierre_z_monto": 0, "cierre_z_cant": 0,
        "cupones": [],
    }
    for linea in lineas:
        s = linea.strip()

        m = re.match(r"(.+?)\s+DETALLE INGRESOS\s+(\d{2}/\d{2}/\d{4})\s+Hora:\s*(\d+)\s*(\d+)", linea)
        if m and d["sucursal"] is None:
            d["sucursal"] = m.group(1).strip()
            d["fecha"] = m.group(2)
            d["hora"] = f"{m.group(3)}:{m.group(4)}"

        if s.startswith("VTA DEL DIA"):
            nums = re.findall(r"[\d,]+\.\d{2}", linea)
            if len(nums) >= 5:
                d["vta_efectivo"] = float(nums[0].replace(",", ""))
                d["vta_cheques"]  = float(nums[1].replace(",", ""))
                d["vta_mp_tr"]    = float(nums[2].replace(",", ""))
                d["vta_posnet"]   = float(nums[3].replace(",", ""))
                d["vta_del_dia"]  = float(nums[4].replace(",", ""))

        if s.startswith("TOTAL INGRESOS"):
            nums = re.findall(r"[\d,]+\.\d{2}", linea)
            if len(nums) >= 5:
                d["total_ingresos"] = float(nums[4].replace(",", ""))

        if s.startswith("TOTAL EGRESOS"):
            nums = re.findall(r"[\d,]+\.\d{2}", linea)
            if len(nums) >= 5:
                d["total_egresos"] = float(nums[4].replace(",", ""))

        if s.startswith("DIFERENCIA") and "CAJA" not in s:
            nums = re.findall(r"-?[\d,]+\.\d{2}", linea)
            if nums:
                d["diferencia_caja"] = float(nums[-1].replace(",", ""))

        m = re.search(r"CIERRE Z=\s*([\d,]+\.\d{2})\s*\(\s*(\d+)\)", linea)
        if m:
            d["cierre_z_monto"] = float(m.group(1).replace(",", ""))
            d["cierre_z_cant"]  = int(m.group(2))

        m = re.match(r"\s*(\d{2}/\d{2}/\d{2})\s+([A-Z]\d+)\s+([\d,]+\.\d{2})\s+(.+)$", linea)
        if m:
            d["cupones"].append({
                "fecha": m.group(1), "cupon": m.group(2),
                "importe": float(m.group(3).replace(",", "")),
                "tarjeta": m.group(4).strip(),
                "sucursal": d["sucursal"],
            })
    return d


# ============================================================
# CARGA DE ARCHIVOS
# ============================================================
st.sidebar.header("📁 Carga de archivos")
archivos = st.sidebar.file_uploader(
    "Subí los archivos .LIS",
    type=["lis", "txt"],
    accept_multiple_files=True,
)

if st.sidebar.button("🗑️ Limpiar archivos"):
    st.rerun()

cierres = []
cupones_todos = []
errores = []

if archivos:
    with st.spinner("Procesando archivos..."):
        for f in archivos:
            try:
                contenido = detectar_encoding(f.read())
                d = parsear_lis(contenido, f.name)
                if d["sucursal"] is None:
                    errores.append(f"{f.name}: no se detectó encabezado")
                    continue
                cupones_todos.extend(d.pop("cupones"))
                cierres.append(d)
            except Exception as e:
                errores.append(f"{f.name}: {str(e)}")

df = pd.DataFrame(cierres) if cierres else pd.DataFrame()
df_cupones = pd.DataFrame(cupones_todos) if cupones_todos else pd.DataFrame()

# Detectar duplicados
if not df.empty:
    dup = df[df.duplicated(subset=["sucursal", "fecha"], keep=False)]
    if not dup.empty:
        st.warning(f"⚠️ Detectadas {len(dup)} cargas duplicadas (misma sucursal + fecha).")

# ============================================================
# HEADER
# ============================================================
if not df.empty:
    fechas = sorted(df["fecha"].dropna().unique().tolist())
    fecha_txt = " · ".join(fechas) if fechas else "Sin datos"
    n_suc = df["sucursal"].nunique()
    st.markdown(f"""
    <div class="header-box">
        <h1>💰 Resumen General de Ventas</h1>
        <p>Análisis consolidado de todas las sucursales</p>
        <div style="margin-top:14px">
            <span class="pill">📅 {fecha_txt}</span>
            <span class="pill">🏪 {n_suc} sucursales</span>
            <span class="pill">🕒 {datetime.now():%H:%M}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="header-box">
        <h1>💰 Resumen General de Ventas</h1>
        <p>Análisis consolidado de todas las sucursales</p>
    </div>
    """, unsafe_allow_html=True)

if errores:
    for e in errores:
        st.error(f"❌ Error al procesar: {e}")

if not archivos:
    st.info("👈 Subí los archivos .LIS desde el panel izquierdo para comenzar.")
    st.stop()

if df.empty:
    st.error("❌ No se pudo procesar ningún archivo correctamente.")
    st.stop()

# ============================================================
# FILTROS
# ============================================================
st.sidebar.markdown("---")
st.sidebar.header("🔍 Filtros")

sucursales = sorted(df["sucursal"].dropna().unique().tolist())
suc_sel = st.sidebar.multiselect("Sucursales", sucursales, default=sucursales)

fechas_disp = sorted(df["fecha"].dropna().unique().tolist())
fecha_sel = st.sidebar.multiselect("Fechas", fechas_disp, default=fechas_disp)

if suc_sel:
    df = df[df["sucursal"].isin(suc_sel)]
    if not df_cupones.empty:
        df_cupones = df_cupones[df_cupones["sucursal"].isin(suc_sel)]

if fecha_sel:
    df = df[df["fecha"].isin(fecha_sel)]

if df.empty:
    st.warning("⚠️ No hay datos con los filtros seleccionados.")
    st.stop()

# ============================================================
# ALERTAS
# ============================================================
alertas = df[df["diferencia_caja"].abs() > 500]
if not alertas.empty:
    st.markdown(f"""
    <div class="alert-box">
        <p>⚠️ {len(alertas)} sucursal(es) con diferencia de caja superior a $500. Revisar abajo en la pestaña "Por sucursal".</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# KPIs
# ============================================================
total_venta = df["vta_del_dia"].sum()
total_cupones = df["cierre_z_cant"].sum()
ticket_prom = total_venta / total_cupones if total_cupones > 0 else 0
total_efectivo = df["vta_efectivo"].sum()
total_posnet = df["vta_posnet"].sum()
total_mp = df["vta_mp_tr"].sum()
total_egresos = df["total_egresos"].sum()
diferencia = df["diferencia_caja"].sum()

suc_top = df.loc[df["vta_del_dia"].idxmax()]
suc_top_nombre = suc_top["sucursal"]
suc_top_monto = suc_top["vta_del_dia"]
pct_top = (suc_top_monto / total_venta * 100) if total_venta > 0 else 0

col1, col2, col3, col4 = st.columns(4)
col1.metric("💰 Venta total", formato_pesos(total_venta))
col2.metric("🧾 Cupones emitidos", f"{int(total_cupones):,}".replace(",", "."))
col3.metric("🎯 Ticket promedio", formato_pesos(ticket_prom))
col4.metric(
    f"{color_diferencia(diferencia)} Diferencia caja",
    formato_pesos(diferencia),
    delta="OK" if abs(diferencia) < 100 else "Revisar",
    delta_color="normal" if abs(diferencia) < 100 else "inverse",
)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

col_a, col_b, col_c = st.columns([1.4, 1, 1])

with col_a:
    st.markdown(f"""
    <div class="highlight-card">
        <h3>🏆 Sucursal destacada</h3>
        <div class="big-number">{suc_top_nombre}</div>
        <div class="sub">{formato_pesos(suc_top_monto)} · {pct_top:.1f}% del total</div>
    </div>
    """, unsafe_allow_html=True)

with col_b:
    pct_ef = (total_efectivo / total_venta * 100) if total_venta > 0 else 0
    st.markdown(f"""
    <div class="highlight-card">
        <h3>💵 Efectivo</h3>
        <div class="big-number">{formato_pesos(total_efectivo)}</div>
        <div class="sub">{pct_ef:.1f}% del total</div>
    </div>
    """, unsafe_allow_html=True)

with col_c:
    pct_pos = (total_posnet / total_venta * 100) if total_venta > 0 else 0
    st.markdown(f"""
    <div class="highlight-card">
        <h3>💳 Posnet</h3>
        <div class="big-number">{formato_pesos(total_posnet)}</div>
        <div class="sub">{pct_pos:.1f}% del total</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

col5, col6, col7, col8 = st.columns(4)
col5.metric("📱 M.P./Transf.", formato_pesos(total_mp))
col6.metric("📤 Egresos", formato_pesos(total_egresos))
col7.metric("📊 Cierres leídos", f"{len(df)}")
col8.metric("📅 Fecha", df['fecha'].iloc[0])

st.divider()

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Resumen general",
    "🏪 Por sucursal",
    "💳 Medios de pago",
    "🎫 Cupones y tarjetas",
    "📥 Datos y exportar",
])

# ------------------------------------------------------------
# TAB 1
# ------------------------------------------------------------
with tab1:
    col_a, col_b = st.columns([1.3, 1])

    with col_a:
        st.markdown("##### 🏆 Ranking de sucursales")
        ranking = df.sort_values("vta_del_dia", ascending=True)
        fig = px.bar(
            ranking, x="vta_del_dia", y="sucursal", orientation="h",
            text="vta_del_dia",
            color="vta_del_dia",
            color_continuous_scale=["#dbeafe", "#2c7fb8", "#1f4e79"],
        )
        fig.update_traces(
            texttemplate="$%{text:,.0f}", textposition="outside",
            textfont=dict(size=13, color="#0f172a"),
            marker_line_width=0,
            hovertemplate="<b>%{y}</b><br>$ %{x:,.0f}<extra></extra>",
        )
        fig.update_layout(
            height=380, showlegend=False, coloraxis_showscale=False,
            margin=dict(l=0, r=80, t=10, b=0),
            xaxis_title="", yaxis_title="", **PLOTLY_LAYOUT,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.markdown("##### 🥧 Participación por sucursal")
        colores = ["#1f4e79", "#2c7fb8", "#4da3d9", "#7fc7e8", "#a5d8f0", "#cce9f7"]
        fig = px.pie(
            df, names="sucursal", values="vta_del_dia", hole=0.5,
            color_discrete_sequence=colores,
        )
        fig.update_traces(
            textinfo="percent",
            textfont=dict(size=14, color="white"),
            marker=dict(line=dict(color="white", width=2)),
            hovertemplate="<b>%{label}</b><br>$ %{value:,.0f}<br>%{percent}<extra></extra>",
        )
        fig.update_layout(
            height=380, showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15,
                        xanchor="center", x=0.5,
                        font=dict(size=12, color="#0f172a")),
            margin=dict(l=0, r=0, t=10, b=0), **PLOTLY_LAYOUT,
        )
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# TAB 2
# ------------------------------------------------------------
with tab2:
    st.markdown("##### Detalle por sucursal")
    tabla = df.copy()
    tabla["ticket_promedio"] = tabla.apply(
        lambda r: r["vta_del_dia"] / r["cierre_z_cant"] if r["cierre_z_cant"] > 0 else 0,
        axis=1,
    )
    tabla["alerta"] = tabla["diferencia_caja"].apply(color_diferencia)
    tabla = tabla.sort_values("vta_del_dia", ascending=False)

    st.dataframe(
        tabla[[
            "alerta", "sucursal", "fecha", "hora", "vta_del_dia", "vta_efectivo",
            "vta_mp_tr", "vta_posnet", "cierre_z_cant", "ticket_promedio",
            "total_egresos", "diferencia_caja",
        ]].style.format({
            "vta_del_dia": formato_pesos, "vta_efectivo": formato_pesos,
            "vta_mp_tr": formato_pesos, "vta_posnet": formato_pesos,
            "ticket_promedio": formato_pesos, "total_egresos": formato_pesos,
            "diferencia_caja": formato_pesos,
        }),
        use_container_width=True, hide_index=True,
    )

# ------------------------------------------------------------
# TAB 3
# ------------------------------------------------------------
with tab3:
    medios = pd.DataFrame({
        "Medio": ["Efectivo", "Cheques", "M.P./Transf.", "Posnet"],
        "Monto": [
            df["vta_efectivo"].sum(), df["vta_cheques"].sum(),
            df["vta_mp_tr"].sum(), df["vta_posnet"].sum(),
        ],
    })
    medios = medios[medios["Monto"] > 0]

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("##### Distribución de medios de pago")
        fig = px.pie(
            medios, names="Medio", values="Monto", hole=0.45,
            color_discrete_sequence=["#1f4e79", "#2c7fb8", "#4da3d9", "#7fc7e8"],
        )
        fig.update_traces(
            textinfo="percent+label",
            textfont=dict(size=13, color="white"),
            marker=dict(line=dict(color="white", width=2)),
            hovertemplate="<b>%{label}</b><br>$ %{value:,.0f}<br>%{percent}<extra></extra>",
        )
        fig.update_layout(height=400, showlegend=False, **PLOTLY_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.markdown("##### Monto por medio")
        fig = px.bar(
            medios, x="Medio", y="Monto", text="Monto",
            color="Medio",
            color_discrete_sequence=["#1f4e79", "#2c7fb8", "#4da3d9", "#7fc7e8"],
        )
        fig.update_traces(
            texttemplate="$%{text:,.0f}", textposition="outside",
            textfont=dict(size=13, color="#0f172a"),
            marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>$ %{y:,.0f}<extra></extra>",
        )
        fig.update_layout(height=400, showlegend=False, **PLOTLY_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# TAB 4
# ------------------------------------------------------------
with tab4:
    if not df_cupones.empty:
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("##### Top tarjetas por monto")
            top_tarj = (
                df_cupones.groupby("tarjeta")["importe"]
                .sum().reset_index()
                .sort_values("importe", ascending=True)
            )
            fig = px.bar(
                top_tarj, x="importe", y="tarjeta", orientation="h",
                color="importe", text="importe",
                color_continuous_scale=["#dbeafe", "#2c7fb8", "#1f4e79"],
            )
            fig.update_traces(
                texttemplate="$%{text:,.0f}", textposition="outside",
                textfont=dict(size=12, color="#0f172a"),
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>$ %{x:,.0f}<extra></extra>",
            )
            fig.update_layout(
                height=450, showlegend=False, coloraxis_showscale=False,
                xaxis_title="", yaxis_title="",
                margin=dict(l=0, r=80, t=10, b=0), **PLOTLY_LAYOUT,
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_b:
            st.markdown("##### Cantidad de cupones por tarjeta")
            cnt_tarj = df_cupones.groupby("tarjeta").size().reset_index(name="cantidad")
            fig = px.pie(
                cnt_tarj, names="tarjeta", values="cantidad", hole=0.45,
                color_discrete_sequence=["#1f4e79", "#2c7fb8", "#4da3d9", "#7fc7e8", "#a5d8f0"],
            )
            fig.update_traces(
                textinfo="percent",
                textfont=dict(size=13, color="white"),
                marker=dict(line=dict(color="white", width=2)),
                hovertemplate="<b>%{label}</b><br>%{value} cupones<br>%{percent}<extra></extra>",
            )
            fig.update_layout(
                height=450, showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.15,
                            xanchor="center", x=0.5,
                            font=dict(size=12, color="#0f172a")),
                **PLOTLY_LAYOUT,
            )
            st.plotly_chart(fig, use_container_width=True)

        st.divider()
        st.markdown("##### Detalle de cupones")
        st.dataframe(
            df_cupones.style.format({"importe": formato_pesos}),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("No hay cupones en los archivos subidos.")

# ------------------------------------------------------------
# TAB 5
# ------------------------------------------------------------
with tab5:
    st.markdown("##### 📋 Datos crudos (cierres)")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("##### 📤 Exportar a Excel")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cierres", index=False)
        if not df_cupones.empty:
            df_cupones.to_excel(writer, sheet_name="Cupones", index=False)

    st.download_button(
        "⬇️ Descargar Excel",
        data=buffer.getvalue(),
        file_name=f"resumen_ventas_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.caption("💡 Tip: los filtros del panel izquierdo actualizan todo automáticamente.")