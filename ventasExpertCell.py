# app.py
import os
import base64
import json
import zlib
import unicodedata
from datetime import datetime, date

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(
    page_title="Ventas ExpertCell",
    page_icon="📊",
    layout="wide",
)

# -------------------------------
# THEME (READ ONLY — we do NOT force anything)
# -------------------------------
try:
    theme_base = st.get_option("theme.base") or "light"  # "light" | "dark"
except Exception:
    theme_base = "light"

IS_DARK = str(theme_base).lower() == "dark"
PLOTLY_TEMPLATE = "plotly_dark" if IS_DARK else "plotly_white"

# -------------------------------
# NEUTRAL, THEME-FRIENDLY CSS (no forced colors)
# -------------------------------
st.markdown(
    """
<style>
header[data-testid="stHeader"]{ background: rgba(0,0,0,0) !important; }
header[data-testid="stHeader"] [data-testid="stToolbar"]{ background: rgba(0,0,0,0) !important; }
header[data-testid="stHeader"] button,
header[data-testid="stHeader"] svg{
  color: var(--text-color) !important;
  fill: var(--text-color) !important;
}

.stApp{
  background-color: var(--background-color) !important;
  background-image:
    radial-gradient(circle at 1px 1px, rgba(127,127,127,0.14) 1px, transparent 0) !important;
  background-size: 18px 18px !important;
  color: var(--text-color) !important;
}
.block-container{ padding-top: 1.2rem; }

section[data-testid="stSidebar"]{
  background: var(--secondary-background-color) !important;
  border-right: 1px solid rgba(127,127,127,0.25) !important;
}
section[data-testid="stSidebar"] *{ color: var(--text-color) !important; }

section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea{
  background: var(--background-color) !important;
  border: 1px solid rgba(127,127,127,0.28) !important;
  color: var(--text-color) !important;
  border-radius: 10px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div{
  background: var(--background-color) !important;
  border: 1px solid rgba(127,127,127,0.28) !important;
  border-radius: 10px !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"]{
  background: rgba(127,127,127,0.25) !important;
  color: var(--text-color) !important;
  border-radius: 999px !important;
  font-weight: 800 !important;
}

.metric-card{
  background: var(--secondary-background-color) !important;
  border-radius: 14px;
  padding: 14px 16px;
  border: 1px solid rgba(127,127,127,0.22);
  box-shadow: 0 1px 0 rgba(0,0,0,0.10);
}
.metric-label{ font-size:0.92rem; opacity:0.78; font-weight:800; }
.metric-value{ font-size:2.25rem; font-weight:900; margin-top:4px; line-height:1; }
.metric-sub{ font-size:0.9rem; opacity:0.70; margin-top:6px; }

.kpi-mini{
  background: var(--secondary-background-color) !important;
  border-radius: 14px;
  padding: 12px 14px;
  border: 1px solid rgba(127,127,127,0.22);
}
.kpi-mini .t{ font-size:0.9rem; font-weight:900; opacity:0.78; }
.kpi-mini .v{ font-size:1.6rem; font-weight:900; margin-top:4px; }

div[data-testid="stPlotlyChart"] > div{ border-radius: 16px; }
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------
# HELPERS (format)
# -------------------------------
def fmt_int(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    return f"{int(round(x)):,}"


def fmt_money_short(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "$-"
    x = float(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1_000_000:
        return f"{sign}${x/1_000_000:.2f}M"
    if x >= 1_000:
        return f"{sign}${x/1_000:.2f}K"
    return f"{sign}${x:,.2f}"


def fmt_pct(x: float | None) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "-"
    return f"{x*100:.2f}%"


def normalize_name(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().upper()
    s = " ".join(s.split())
    s = unicodedata.normalize("NFKD", s)
    s = "".join([c for c in s if not unicodedata.combining(c)])
    return s


def metric_card(label: str, value: str, sub: str | None = None):
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_mini(label: str, value: str):
    st.markdown(
        f"""
        <div class="kpi-mini">
          <div class="t">{label}</div>
          <div class="v">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def month_key_to_name_es(ym: int) -> str:
    y = ym // 100
    m = ym % 100
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    if 1 <= m <= 12:
        return f"{meses[m-1]} {y}"
    return str(ym)


def apply_plotly_theme(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


# ✅ Totals row helper (for every table shown)
def add_totals_row(
    df: pd.DataFrame,
    label_col: str,
    totals: dict,
    label: str = "TOTAL",
) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    row = {c: np.nan for c in out.columns}
    if label_col in row:
        row[label_col] = label
    for k, v in totals.items():
        if k in row:
            row[k] = v
    return pd.concat([out, pd.DataFrame([row])], ignore_index=True)


# ✅ NEW: bold totals rows in st.dataframe using Styler
def style_totals_bold(df: pd.DataFrame, label_col: str):
    def _bold_row(row):
        v = row.get(label_col, "")
        if "TOTAL" in str(v).upper():
            return ["font-weight: 800"] * len(row)
        return [""] * len(row)

    return df.style.apply(_bold_row, axis=1)


# -------------------------------
# DB (SQL Server via pyodbc)
# -------------------------------
def get_db_cfg():
    if "db" in st.secrets:
        return {
            "server": st.secrets["db"]["server"],
            "database": st.secrets["db"]["database"],
            "username": st.secrets["db"]["username"],
            "password": st.secrets["db"]["password"],
            "driver": st.secrets["db"].get("driver", "ODBC Driver 17 for SQL Server"),
        }
    return {
        "server": os.getenv("DB_SERVER", ""),
        "database": os.getenv("DB_DATABASE", ""),
        "username": os.getenv("DB_USERNAME", ""),
        "password": os.getenv("DB_PASSWORD", ""),
        "driver": os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server"),
    }


@st.cache_data(ttl=600, show_spinner=False)
def read_sql(query: str) -> pd.DataFrame:
    import pyodbc
    cfg = get_db_cfg()
    conn_str = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        "TrustServerCertificate=yes;"
    )
    with pyodbc.connect(conn_str) as conn:
        return pd.read_sql(query, conn)


# -------------------------------
# POWER QUERY → PANDAS (Consulta2)
# -------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def load_empleados() -> pd.DataFrame:
    q = r"""
    SELECT
      Region,
      Plaza,
      Tienda AS Centro,
      [Nombre Completo] AS Nombre,
      Puesto,
      RFC,
      [Jefe Inmediato],
      Estatus,
      [Fecha Ingreso],
      [Fecha Baja],
      [Canal de Venta],
      Operacion
    FROM reporte_empleado('EMPRESA_MAESTRA',1,'','') AS e
    WHERE
      [Canal de Venta] IN ('ATT', 'IZZI')
      AND [Operacion] IN ('CONTACT CENTER')
      AND [Tipo Tienda] IN ('VIRTUAL')
      AND (
        Estatus = 'ACTIVO'
        OR (
          Estatus = 'BAJA'
          AND [Fecha Baja] >= DATEADD(MONTH, -1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
          AND [Fecha Baja] <  DATEADD(MONTH,  1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
        )
      )
    """
    df = read_sql(q)
    df["Nombre"] = df["Nombre"].astype(str).str.strip()
    df["Jefe Inmediato"] = df["Jefe Inmediato"].astype(str).str.strip()
    df["Estatus"] = df["Estatus"].astype(str).str.strip()
    df["Centro"] = df["Centro"].astype(str).str.strip()
    df["Puesto"] = df["Puesto"].astype(str).str.strip()
    return df


# -------------------------------
# POWER QUERY → PANDAS (Consulta1)
# -------------------------------
def build_ventas_query(start_yyyymmdd: str, end_yyyymmdd: str) -> str:
    return f"""
    SELECT
      FOLIO,
      [PTO. DE VENTA] AS CENTRO,
      [OPERACION PDV],
      [ESTATUS],
      [EJECUTIVO],
      [FECHA DE CAPTURA],
      [PLAN],
      [RENTA SIN IMPUESTOS],
      [PRECIO],
      [SUBREGION]
    FROM reporte_ventas_no_conciliadas('EMPRESA_MAESTRA', 4, '{start_yyyymmdd}', '{end_yyyymmdd}', 1, '19000101', '20990101')
    WHERE
      [OPERACION PDV] = 'CONTACT CENTER'
      AND [PTO. DE VENTA] LIKE 'EXP ATT C CENTER%'
    """


@st.cache_data(ttl=600, show_spinner=False)
def load_ventas(start_yyyymmdd: str, end_yyyymmdd: str) -> pd.DataFrame:
    q = build_ventas_query(start_yyyymmdd, end_yyyymmdd)
    df = read_sql(q)

    df["EJECUTIVO"] = df["EJECUTIVO"].astype(str).str.strip()
    df["CENTRO"] = df["CENTRO"].astype(str).str.strip()

    def fix_centro(c: str) -> str:
        c_up = str(c).upper()
        if "JUAREZ" in c_up:
            return "EXP ATT C CENTER JUAREZ"
        if "CENTER 2" in c_up:
            return "EXP ATT C CENTER 2"
        return c

    df["CENTRO"] = df["CENTRO"].apply(fix_centro)

    df["EJECUTIVO"] = df["EJECUTIVO"].replace(
        {
            "CESAR JAHACIEL ALONSO GARCIAA": "CESAR JAHACIEL ALONSO GARCIA",
            "VICTOR BETANZO FUENTES": "VICTOR BETANZOS FUENTES",
        }
    )

    df["FECHA DE CAPTURA"] = pd.to_datetime(df["FECHA DE CAPTURA"], errors="coerce")
    df["Fecha"] = df["FECHA DE CAPTURA"].dt.date
    df["Hora"] = df["FECHA DE CAPTURA"].dt.time
    df["Año"] = df["FECHA DE CAPTURA"].dt.year
    df["Mes"] = df["FECHA DE CAPTURA"].dt.month
    df["NombreMes"] = df["FECHA DE CAPTURA"].dt.strftime("%B").str.lower()
    df["AñoMes"] = df["Año"] * 100 + df["Mes"]

    iso = df["FECHA DE CAPTURA"].dt.isocalendar()
    df["ISOYear"] = iso.year.astype(int)
    df["ISOWeek"] = iso.week.astype(int)
    df["SemanaAño"] = df["ISOWeek"]
    df["WeekKey"] = df["ISOYear"] * 100 + df["ISOWeek"]
    df["SemanaISO"] = df["ISOYear"].astype(str) + "-W" + df["ISOWeek"].astype(str).str.zfill(2)

    df["DiaSemana"] = df["FECHA DE CAPTURA"].dt.day_name().astype(str)
    df["DiaNum"] = df["FECHA DE CAPTURA"].dt.day.astype(int)

    for col in ["PRECIO", "RENTA SIN IMPUESTOS"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["CentroKey"] = np.where(df["CENTRO"].str.upper().str.contains("JUAREZ", na=False), "JV", "CC2")
    return df


def add_empleado_join(ventas: pd.DataFrame, empleados: pd.DataFrame) -> pd.DataFrame:
    emp = empleados[["Nombre", "Jefe Inmediato"]].copy()
    emp["Nombre"] = emp["Nombre"].astype(str).str.strip()

    out = ventas.merge(emp, left_on="EJECUTIVO", right_on="Nombre", how="left")
    out.rename(columns={"Jefe Inmediato": "Supervisor"}, inplace=True)
    out["Supervisor"] = out["Supervisor"].fillna("").replace({"": "BAJA"})
    out["Supervisor_norm"] = out["Supervisor"].apply(normalize_name)
    out["EJECUTIVO_norm"] = out["EJECUTIVO"].apply(normalize_name)
    return out


# -------------------------------
# METAS (MANUAL TABLE INSIDE CODE)  ✅ REPLACES POWERQUERY BASE64
# -------------------------------
# ✅ EDIT METAS HERE (by hand)
METAS_MANUAL_ROWS = [
    # --- Metas Centro ---
    {"IDCenter": "CC1", "Nivel": "Centro", "Nombre": "EDUARDO AGUILA SANCHEZ", "Centro": "CC2", "Meta": 640},
    {"IDCenter": "JV1", "Nivel": "Centro", "Nombre": "MARIA LUISA MEZA GOEL",  "Centro": "JV",  "Meta": 252},

    # --- Metas Supervisor (JV) ---
    {"IDCenter": "JV2", "Nivel": "Supervisor", "Nombre": "JORGE MIGUEL UREÑA ZARATE",        "Centro": "JV",  "Meta": 104},
    {"IDCenter": "JV3", "Nivel": "Supervisor", "Nombre": "MARIA FERNANDA MARTINEZ BISTRAIN", "Centro": "JV",  "Meta": 148},

    # --- Metas Supervisor (CC2) ---
    {"IDCenter": "CC2", "Nivel": "Supervisor", "Nombre": "ALFREDO CABRERA PADRON",          "Centro": "CC2", "Meta": 214},
    {"IDCenter": "CC4", "Nivel": "Supervisor", "Nombre": "REYNA LIZZETTE MARTINEZ GARCIA",  "Centro": "CC2", "Meta": 226},

    # ❌ JULIO is intentionally NOT included
]


def load_metas_df() -> pd.DataFrame:
    # No cache on purpose: edits to METAS_MANUAL_ROWS should reflect immediately.
    df = pd.DataFrame(METAS_MANUAL_ROWS, columns=["IDCenter", "Nivel", "Nombre", "Centro", "Meta"])
    if df.empty:
        return pd.DataFrame(columns=["IDCenter", "Nivel", "Nombre", "Centro", "Meta", "Nombre_norm"])

    df["Nombre"] = df["Nombre"].astype(str).str.strip()
    df["Centro"] = df["Centro"].astype(str).str.strip().str.upper()
    df["Nivel"] = df["Nivel"].astype(str).str.strip()
    df["Meta"] = pd.to_numeric(df["Meta"], errors="coerce")
    df["Nombre_norm"] = df["Nombre"].apply(normalize_name)
    return df


# -------------------------------
# MEASURES
# -------------------------------
def total_folios(df: pd.DataFrame) -> int:
    return int(len(df))


def total_precio(df: pd.DataFrame) -> float:
    return float(df["PRECIO"].sum(skipna=True))


def total_renta(df: pd.DataFrame) -> float:
    return float(df["RENTA SIN IMPUESTOS"].sum(skipna=True))


def arpu(df: pd.DataFrame) -> float:
    fol = total_folios(df)
    return (total_precio(df) / fol) if fol else np.nan


def arpu_siva(df: pd.DataFrame) -> float:
    fol = total_folios(df)
    return (total_renta(df) / fol) if fol else np.nan


def distinct_days_with_sales(df: pd.DataFrame, exclude_sunday: bool = True) -> int:
    if df.empty:
        return 0
    tmp = df.copy()
    tmp["Fecha_dt"] = pd.to_datetime(tmp["Fecha"])
    if exclude_sunday:
        tmp = tmp[tmp["Fecha_dt"].dt.weekday <= 5]
    return int(tmp["Fecha"].nunique())


def promedio_diario(df: pd.DataFrame) -> float:
    d = distinct_days_with_sales(df, exclude_sunday=True)
    return (total_folios(df) / d) if d else np.nan


def max_folios_dia(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    g = df.groupby("Fecha", as_index=False).size()
    return int(g["size"].max())


def daily_series(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Fecha", "Ventas"])
    out = df.groupby("Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
    out["Fecha"] = pd.to_datetime(out["Fecha"])
    return out.sort_values("Fecha")


def weekly_series(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["SemanaISO", "WeekKey", "Ventas", "PromDiarioSemana", "DiasSemana"])

    ventas_sem = (
        df.groupby(["WeekKey", "SemanaISO"], as_index=False)
        .size()
        .rename(columns={"size": "Ventas"})
    )

    tmp = df.copy()
    tmp["Fecha_dt"] = pd.to_datetime(tmp["Fecha"])
    tmp = tmp[tmp["Fecha_dt"].dt.weekday <= 5]  # lunes-sábado

    dias_sem = (
        tmp.groupby("WeekKey")["Fecha"]
        .nunique()
        .reset_index()
        .rename(columns={"Fecha": "DiasSemana"})
    )

    out = ventas_sem.merge(dias_sem, on="WeekKey", how="left")
    out["PromDiarioSemana"] = out["Ventas"] / out["DiasSemana"].replace({0: np.nan})
    return out.sort_values("WeekKey")


def filter_month(df: pd.DataFrame, ym: int) -> pd.DataFrame:
    return df[df["AñoMes"] == ym].copy()


def cut_month_mode(df: pd.DataFrame, mode: int, day_cut: int) -> pd.DataFrame:
    if mode == 1:
        return df[df["DiaNum"] <= day_cut].copy()
    return df.copy()


# -------------------------------
# GAUGE
# -------------------------------
def gauge_fig(value: float, meta: float, title: str):
    value = 0 if value is None or np.isnan(value) else float(value)
    meta = 0 if meta is None or np.isnan(meta) else float(meta)
    axis_max = max(meta, value, 1)

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            number={"font": {"size": 38}},
            title={"text": title, "font": {"size": 12}},
            gauge={
                "axis": {"range": [0, axis_max], "tickwidth": 0},
                "bar": {"color": "rgba(127,127,127,0.65)", "thickness": 0.35},
                "bgcolor": "rgba(0,0,0,0)",
                "borderwidth": 0,
                "steps": [{"range": [0, axis_max], "color": "rgba(127,127,127,0.18)"}],
            },
        )
    )
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        margin=dict(l=8, r=8, t=40, b=10),
        height=250,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


# -------------------------------
# DETALLE HELPERS (CC2 / JV + pies)
# -------------------------------
def build_detalle_matrix(df_: pd.DataFrame) -> pd.DataFrame:
    if df_.empty:
        return pd.DataFrame(columns=["Supervisor", "Ejecutivo", "TotalVentas", "MontoVendido", "ARPU"])

    by_sup = (
        df_.groupby("Supervisor", as_index=False)
        .agg(
            TotalVentas=("FOLIO", "count"),
            MontoVendido=("PRECIO", "sum"),
            ARPU=("PRECIO", lambda s: s.sum() / len(s) if len(s) else np.nan),
        )
        .sort_values("TotalVentas", ascending=False)
    )

    by_ej = (
        df_.groupby(["Supervisor", "EJECUTIVO"], as_index=False)
        .agg(
            TotalVentas=("FOLIO", "count"),
            MontoVendido=("PRECIO", "sum"),
            ARPU=("PRECIO", lambda s: s.sum() / len(s) if len(s) else np.nan),
        )
        .sort_values(["Supervisor", "TotalVentas"], ascending=[True, False])
    )

    rows = []
    for _, sr in by_sup.iterrows():
        sup = sr["Supervisor"]
        rows.append(
            {
                "Supervisor": str(sup),
                "Ejecutivo": "",
                "TotalVentas": int(sr["TotalVentas"]),
                "MontoVendido": float(sr["MontoVendido"]),
                "ARPU": float(sr["ARPU"]) if pd.notna(sr["ARPU"]) else np.nan,
            }
        )
        sub = by_ej[by_ej["Supervisor"] == sup].sort_values("TotalVentas", ascending=False)
        for _, er in sub.iterrows():
            rows.append(
                {
                    "Supervisor": "",
                    "Ejecutivo": "   " + str(er["EJECUTIVO"]),
                    "TotalVentas": int(er["TotalVentas"]),
                    "MontoVendido": float(er["MontoVendido"]),
                    "ARPU": float(er["ARPU"]) if pd.notna(er["ARPU"]) else np.nan,
                }
            )

    out = pd.DataFrame(rows, columns=["Supervisor", "Ejecutivo", "TotalVentas", "MontoVendido", "ARPU"])
    return out


def donut_compare_fig(labels: list[str], values: list[float], title: str, value_formatter):
    vals = [0.0 if (v is None or (isinstance(v, float) and np.isnan(v))) else float(v) for v in values]
    total = float(np.sum(vals))
    if total <= 0:
        return None

    texts = [value_formatter(v) for v in vals]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=vals,
                hole=0.55,
                text=texts,
                textinfo="label+text+percent",
                insidetextorientation="radial",
                sort=False,
            )
        ]
    )
    fig.update_layout(title=title, height=320)
    apply_plotly_theme(fig)
    return fig


# -------------------------------
# LOAD DATA
# -------------------------------
st.sidebar.header("⚙️ Parámetros")

# Refresh button
if "last_refresh" not in st.session_state:
    st.session_state["last_refresh"] = None

btn_cols = st.sidebar.columns([1, 1])
with btn_cols[0]:
    if st.button("🔄 Actualizar datos", use_container_width=True):
        st.cache_data.clear()
        st.session_state["last_refresh"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        st.rerun()
with btn_cols[1]:
    st.caption(f"🕒 {st.session_state['last_refresh']}" if st.session_state["last_refresh"] else "")

# ✅ Auto-detect today's date and use it as default "Fin"
default_start = "20250801"
default_start_dt = datetime.strptime(default_start, "%Y%m%d").date()
default_end_dt = date.today()

d1, d2 = st.sidebar.columns(2)
start_dt = d1.date_input("Inicio", value=default_start_dt, format="YYYY-MM-DD")
end_dt = d2.date_input("Fin", value=default_end_dt, format="YYYY-MM-DD")

if start_dt > end_dt:
    st.sidebar.error("⚠️ Inicio no puede ser mayor que Fin.")
    st.stop()

start_yyyymmdd = start_dt.strftime("%Y%m%d")
end_yyyymmdd = end_dt.strftime("%Y%m%d")

with st.spinner("Cargando datos desde SQL Server…"):
    empleados = load_empleados()
    ventas_raw = load_ventas(start_yyyymmdd, end_yyyymmdd)
    ventas = add_empleado_join(ventas_raw, empleados)
    metas = load_metas_df()

if ventas.empty:
    st.error("No hay datos en el rango seleccionado.")
    st.stop()

meses_disponibles = sorted(ventas["AñoMes"].dropna().unique().tolist())
mes_labels = {ym: month_key_to_name_es(int(ym)) for ym in meses_disponibles}

mes_sel = st.sidebar.selectbox(
    "Mes",
    options=meses_disponibles,
    format_func=lambda ym: mes_labels.get(ym, str(ym)),
    index=len(meses_disponibles) - 1,
)

center_keys = ["CC2", "JV"]
center_sel = st.sidebar.multiselect("Centro (CC2 / JV)", options=center_keys, default=center_keys)

supervisores = sorted([s for s in ventas["Supervisor"].dropna().unique().tolist()])
sup_sel = st.sidebar.multiselect("Supervisor", options=supervisores, default=[])

ejecutivos = sorted([e for e in ventas["EJECUTIVO"].dropna().unique().tolist()])
ej_sel = st.sidebar.multiselect("Ejecutivo", options=ejecutivos, default=[])

subregs = sorted([s for s in ventas["SUBREGION"].dropna().unique().tolist()])
sub_sel = st.sidebar.multiselect("Subregión", options=subregs, default=[])

# Filters (month + sidebar filters)
df_base = ventas.copy()
df_base = df_base[df_base["AñoMes"] == mes_sel]
if center_sel:
    df_base = df_base[df_base["CentroKey"].isin(center_sel)]
if sup_sel:
    df_base = df_base[df_base["Supervisor"].isin(sup_sel)]
if ej_sel:
    df_base = df_base[df_base["EJECUTIVO"].isin(ej_sel)]
if sub_sel:
    df_base = df_base[df_base["SUBREGION"].isin(sub_sel)]

# -------------------------------
# TABS
# -------------------------------
tabs = st.tabs(
    [
        "🌐 Global Mes",
        "📅 Semanas",
        "🆚 JV vs CC2",
        "📉 Mes vs Mes",
        "📋 Detalle",
        "🗺️ Región",
        "🏆 Tops",
        "🎯 Metas",
        "📈 Tendencia Ejecutivo",
    ]
)

# ======================================================
# TAB 1: Global del Mes
# ======================================================
with tabs[0]:
    st.markdown(f"## Ventas Globales del Mes — **{mes_labels[mes_sel]}**")

    colA, colB = st.columns([0.78, 0.22], gap="large")

    with colA:
        k1, k2, k3, k4 = st.columns(4, gap="medium")
        with k1:
            metric_card("Promedio de Ventas", fmt_int(promedio_diario(df_base)))
        with k2:
            metric_card("Total de Ventas", fmt_int(total_folios(df_base)))
        with k3:
            metric_card("Max Ventas en un Día", fmt_int(max_folios_dia(df_base)))
        with k4:
            metric_card("Monto Vendido", fmt_money_short(total_precio(df_base)))

        s = daily_series(df_base)
        avg_line = promedio_diario(df_base)

        fig = go.Figure()
        fig.add_trace(go.Bar(x=s["Fecha"], y=s["Ventas"], name="Total de Ventas"))
        fig.add_trace(go.Scatter(x=s["Fecha"], y=[avg_line] * len(s), mode="lines", name="Promedio Diario de Ventas"))
        fig.update_layout(
            title="Vista General de Ventas",
            xaxis_title="Fecha",
            yaxis_title="Total de Ventas",
            height=380,
        )
        apply_plotly_theme(fig)
        st.plotly_chart(fig, width="stretch")

        top_days = s.sort_values("Ventas", ascending=False).head(8).copy()
        top_days["Fecha"] = top_days["Fecha"].dt.strftime("%A, %d %B %Y")

        # ✅ ADD TOTAL ROW (month total) + ✅ BOLD TOTAL
        top_days_show = top_days.rename(columns={"Ventas": "Total de Ventas"})
        top_days_show = add_totals_row(
            top_days_show,
            label_col="Fecha",
            totals={"Total de Ventas": total_folios(df_base)},
            label="TOTAL (Mes)",
        )

        st.dataframe(
            style_totals_bold(top_days_show, label_col="Fecha").format({"Total de Ventas": "{:,.0f}"}),
            hide_index=True,
            width="stretch",
        )

        st.caption(f"🕒 Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    with colB:
        st.markdown("### ARPU")
        kpi_mini("ARPU", fmt_money_short(arpu(df_base)))
        kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_base)))

# ======================================================
# TAB 2: Semanas
# ======================================================
with tabs[1]:
    st.markdown("## Vista por Semanas")

    df_sem = ventas.copy()
    if center_sel:
        df_sem = df_sem[df_sem["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_sem = df_sem[df_sem["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_sem = df_sem[df_sem["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_sem = df_sem[df_sem["SUBREGION"].isin(sub_sel)]

    meses_sem = sorted(df_sem["AñoMes"].dropna().unique().tolist())
    sem_labels = ["Todos los meses"] + [mes_labels.get(m, str(m)) for m in meses_sem]
    sem_choice = st.selectbox("Mes (Semanas)", options=sem_labels, index=0)

    if sem_choice != "Todos los meses":
        inv = {mes_labels.get(m, str(m)): m for m in meses_sem}
        mes_sem_sel = inv.get(sem_choice)
        if mes_sem_sel is not None:
            month_rows = df_sem[df_sem["AñoMes"] == mes_sem_sel].copy()
            week_keys = month_rows["WeekKey"].dropna().unique().tolist()
            df_sem = df_sem[df_sem["WeekKey"].isin(week_keys)].copy()

    w = weekly_series(df_sem)
    prom_global_sem = np.nanmean(w["PromDiarioSemana"].values) if not w.empty else np.nan

    col1, col2, col3 = st.columns([0.33, 0.34, 0.33])
    with col1:
        metric_card("Promedio de Ventas", fmt_int(prom_global_sem))
    with col2:
        metric_card("Total de Ventas", fmt_int(total_folios(df_sem)))
    with col3:
        if not df_sem.empty:
            last_week_key = int(df_sem["WeekKey"].max())
            tmpw = df_sem[df_sem["WeekKey"] == last_week_key].copy()
            dias = distinct_days_with_sales(tmpw, exclude_sunday=True)
            metric_card("Semana actual", f"{dias} días")

    figw = go.Figure()
    figw.add_trace(go.Bar(x=w["SemanaISO"], y=w["Ventas"], name="Total de Ventas"))
    figw.add_trace(go.Scatter(x=w["SemanaISO"], y=w["PromDiarioSemana"], mode="lines+markers", name="Promedio Diario por Semana"))
    figw.update_layout(
        title="Vista General de Ventas",
        xaxis_title="SEMANA",
        yaxis_title="Total de Ventas",
        height=460,
    )
    apply_plotly_theme(figw)
    st.plotly_chart(figw, width="stretch")

    st.markdown("### ARPU (rango mostrado)")
    kpi_mini("ARPU", fmt_money_short(arpu(df_sem)))
    kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_sem)))

# ======================================================
# TAB 3: JV vs CC2
# ======================================================
with tabs[2]:
    st.markdown(f"## Ventas del Mes — Comparativo (JV vs CC2) — **{mes_labels[mes_sel]}**")

    df_m = filter_month(ventas, mes_sel)
    if sub_sel:
        df_m = df_m[df_m["SUBREGION"].isin(sub_sel)]

    df_jv = df_m[df_m["CentroKey"] == "JV"].copy()
    df_cc2 = df_m[df_m["CentroKey"] == "CC2"].copy()

    cL, cR = st.columns(2, gap="large")

    def render_group(col, title, df_):
        with col:
            st.markdown(f"### {title}")
            a, b, c, d = st.columns(4, gap="medium")
            with a:
                metric_card("Monto Vendido", fmt_money_short(total_precio(df_)))
            with b:
                metric_card("Promedio de Ventas", fmt_int(promedio_diario(df_)))
            with c:
                metric_card("Total de Ventas", fmt_int(total_folios(df_)))
            with d:
                metric_card("Max Ventas en un Día", fmt_int(max_folios_dia(df_)))

            s = daily_series(df_)
            avg = promedio_diario(df_)
            fig = go.Figure()
            fig.add_trace(go.Bar(x=s["Fecha"], y=s["Ventas"], name="Total de Ventas"))
            fig.add_trace(go.Scatter(x=s["Fecha"], y=[avg] * len(s), mode="lines", name="Promedio Diario de Venta"))
            fig.update_layout(title="Vista General de Ventas", height=340)
            apply_plotly_theme(fig)
            st.plotly_chart(fig, width="stretch")

            top = s.sort_values("Ventas", ascending=False).head(6).copy()
            top["Fecha"] = top["Fecha"].dt.strftime("%A, %d %B %Y")

            # ✅ ADD TOTAL ROW (month total for that center) + ✅ BOLD TOTAL
            top_show = top.rename(columns={"Ventas": "Total de Ventas"})
            top_show = add_totals_row(
                top_show,
                label_col="Fecha",
                totals={"Total de Ventas": total_folios(df_)},
                label="TOTAL (Mes)",
            )

            st.dataframe(
                style_totals_bold(top_show, label_col="Fecha").format({"Total de Ventas": "{:,.0f}"}),
                hide_index=True,
                width="stretch",
            )

            kpi_mini("ARPU", fmt_money_short(arpu(df_)))
            kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_)))

    render_group(cL, "JV (Juárez)", df_jv)
    render_group(cR, "CC2 (Center 2)", df_cc2)

# ======================================================
# TAB 4: Mes vs Mes
# ======================================================
with tabs[3]:
    st.markdown("## Mes vs Mes")

    months_all = sorted(ventas["AñoMes"].dropna().unique().tolist())
    if len(months_all) < 2:
        st.warning("Se requieren al menos 2 meses para comparar.")
    else:
        colS1, colS2, colS3 = st.columns([0.35, 0.35, 0.30])
        with colS1:
            mes_actual = st.selectbox(
                "Mes Actual",
                options=months_all,
                format_func=lambda ym: mes_labels.get(ym, str(ym)),
                index=len(months_all) - 1,
            )
        with colS2:
            mes_comp = st.selectbox(
                "Mes Comparado",
                options=months_all,
                format_func=lambda ym: mes_labels.get(ym, str(ym)),
                index=max(0, len(months_all) - 2),
            )
        with colS3:
            modo = st.selectbox(
                "ModoComparación",
                options=[0, 1],
                format_func=lambda x: "0 = Mes completo" if x == 0 else "1 = Cortar al día de hoy",
            )

        dfA = filter_month(ventas, mes_actual)
        dfB = filter_month(ventas, mes_comp)

        if center_sel:
            dfA = dfA[dfA["CentroKey"].isin(center_sel)]
            dfB = dfB[dfB["CentroKey"].isin(center_sel)]
        if sub_sel:
            dfA = dfA[dfA["SUBREGION"].isin(sub_sel)]
            dfB = dfB[dfB["SUBREGION"].isin(sub_sel)]

        today_day = date.today().day
        dfA_m = cut_month_mode(dfA, modo, today_day)
        dfB_m = cut_month_mode(dfB, modo, today_day)

        ventasA = total_folios(dfA_m)
        ventasB = total_folios(dfB_m)

        dif = ventasA - ventasB
        pct = (dif / ventasB) if ventasB else np.nan
        arrow = "↑" if dif > 0 else ("↓" if dif < 0 else "→")

        st.markdown(
            f"""
            <div class="metric-card" style="text-align:center;">
              <div class="metric-value" style="font-size:1.6rem;">
                {dif:+,} ventas ({fmt_pct(pct)}) {arrow}
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns(2, gap="large")
        with c1:
            metric_card("Promedio Diario", fmt_int(promedio_diario(dfA_m)))
            metric_card("Ventas", fmt_int(ventasA))
            sA = daily_series(dfA_m)
            avgA = promedio_diario(dfA_m)
            figA = go.Figure()
            figA.add_trace(go.Bar(x=sA["Fecha"], y=sA["Ventas"], name="Total Folios (Modo)"))
            figA.add_trace(go.Scatter(x=sA["Fecha"], y=[avgA] * len(sA), mode="lines", name="Promedio Diario Mes (Modo)"))
            figA.update_layout(title="Vista General de Ventas", height=340)
            apply_plotly_theme(figA)
            st.plotly_chart(figA, width="stretch")

        with c2:
            metric_card("Promedio Diario", fmt_int(promedio_diario(dfB_m)))
            metric_card("Ventas", fmt_int(ventasB))
            sB = daily_series(dfB_m)
            avgB = promedio_diario(dfB_m)
            figB = go.Figure()
            figB.add_trace(go.Bar(x=sB["Fecha"], y=sB["Ventas"], name="Total Folios (Modo)"))
            figB.add_trace(go.Scatter(x=sB["Fecha"], y=[avgB] * len(sB), mode="lines", name="Promedio Diario Mes (Modo)"))
            figB.update_layout(title="Vista General de Ventas", height=340)
            apply_plotly_theme(figB)
            st.plotly_chart(figB, width="stretch")

        emoji = "📈🔥" if dif > 0 else ("📉⚠️" if dif < 0 else "➖")
        msg = f"{emoji} {mes_labels[mes_actual]} {'mejoró' if dif>0 else ('bajó' if dif<0 else 'mantiene')} {fmt_pct(pct)} vs {mes_labels[mes_comp]} ({dif:+,} ventas)."
        st.markdown(
            f"""
            <div class="metric-card" style="text-align:center;">
              <div class="metric-value" style="font-size:1.25rem;">{msg}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ======================================================
# TAB 5: Detalle (CC2/JV separated + pies)
# ======================================================
with tabs[4]:
    st.markdown(f"## Detalle General de Ventas — **{mes_labels[mes_sel]}**")

    df_d = df_base.copy()
    df_cc2_d = df_d[df_d["CentroKey"] == "CC2"].copy()
    df_jv_d = df_d[df_d["CentroKey"] == "JV"].copy()

    mat_cc2 = build_detalle_matrix(df_cc2_d)
    mat_jv = build_detalle_matrix(df_jv_d)

    cL, cR = st.columns(2, gap="large")
    with cL:
        st.markdown("### CC2 (Center 2)")
        if mat_cc2.empty:
            st.info("Sin datos para CC2 con los filtros actuales.")
        else:
            # ✅ ADD TOTAL ROW (accurate totals from df_cc2_d) + ✅ BOLD TOTAL
            total_row_cc2 = {
                "TotalVentas": int(total_folios(df_cc2_d)),
                "MontoVendido": float(total_precio(df_cc2_d)),
                "ARPU": float(arpu(df_cc2_d)) if pd.notna(arpu(df_cc2_d)) else np.nan,
            }
            mat_cc2_show = add_totals_row(
                mat_cc2,
                label_col="Supervisor",
                totals=total_row_cc2,
                label="TOTAL",
            )

            st.dataframe(
                style_totals_bold(mat_cc2_show, label_col="Supervisor").format(
                    {
                        "TotalVentas": "{:,.0f}",
                        "MontoVendido": "${:,.2f}",
                        "ARPU": "${:,.2f}",
                    }
                ),
                hide_index=True,
                width="stretch",
                height=520,
            )

    with cR:
        st.markdown("### JV (Juárez)")
        if mat_jv.empty:
            st.info("Sin datos para JV con los filtros actuales.")
        else:
            # ✅ ADD TOTAL ROW (accurate totals from df_jv_d) + ✅ BOLD TOTAL
            total_row_jv = {
                "TotalVentas": int(total_folios(df_jv_d)),
                "MontoVendido": float(total_precio(df_jv_d)),
                "ARPU": float(arpu(df_jv_d)) if pd.notna(arpu(df_jv_d)) else np.nan,
            }
            mat_jv_show = add_totals_row(
                mat_jv,
                label_col="Supervisor",
                totals=total_row_jv,
                label="TOTAL",
            )

            st.dataframe(
                style_totals_bold(mat_jv_show, label_col="Supervisor").format(
                    {
                        "TotalVentas": "{:,.0f}",
                        "MontoVendido": "${:,.2f}",
                        "ARPU": "${:,.2f}",
                    }
                ),
                hide_index=True,
                width="stretch",
                height=520,
            )

    st.markdown("---")

    pieL, pieR = st.columns(2, gap="large")
    monto_cc2 = total_precio(df_cc2_d)
    monto_jv = total_precio(df_jv_d)
    fig_monto = donut_compare_fig(["CC2", "JV"], [monto_cc2, monto_jv], "% Monto Vendido", fmt_money_short)

    arpu_cc2_val = arpu(df_cc2_d)
    arpu_jv_val = arpu(df_jv_d)
    fig_arpu = donut_compare_fig(["CC2", "JV"], [arpu_cc2_val, arpu_jv_val], "% ARPU", fmt_money_short)

    with pieL:
        if fig_monto is None:
            st.info("Sin monto suficiente para graficar % Monto Vendido.")
        else:
            st.plotly_chart(fig_monto, width="stretch")

    with pieR:
        if fig_arpu is None:
            st.info("Sin ARPU suficiente para graficar % ARPU.")
        else:
            st.plotly_chart(fig_arpu, width="stretch")

# ======================================================
# TAB 6: Región
# ======================================================
with tabs[5]:
    st.markdown(f"## Ventas por Región — **{mes_labels[mes_sel]}**")

    df_r = df_base.copy()
    if df_r.empty:
        st.info("Sin datos con los filtros actuales.")
    else:
        reg = df_r.groupby("SUBREGION", as_index=False).size().rename(columns={"size": "Ventas"})
        reg = reg.sort_values("Ventas", ascending=False)
        fig = px.pie(reg, names="SUBREGION", values="Ventas", title="VENTAS POR REGIÓN", hole=0.25, template=PLOTLY_TEMPLATE)
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, width="stretch")

# ======================================================
# TAB 7: TOPS
# ======================================================
with tabs[6]:
    st.markdown(f"## TOP Ventas — **{mes_labels[mes_sel]}**")

    df_m_all = filter_month(ventas, mes_sel)

    st.markdown("### TOP Ventas x Ejecutivo (por Centro)")
    centro_top = st.selectbox("Centro para TOP Ejecutivo", options=["JV", "CC2"], index=0)
    df_top = df_m_all[df_m_all["CentroKey"] == centro_top].copy()

    top_ej = (
        df_top.groupby("EJECUTIVO", as_index=False)
        .size()
        .rename(columns={"size": "VENTAS"})
        .sort_values("VENTAS", ascending=False)
        .head(8)
    )

    fig1 = go.Figure()
    fig1.add_trace(
        go.Bar(
            x=top_ej["VENTAS"],
            y=top_ej["EJECUTIVO"],
            orientation="h",
            text=top_ej["VENTAS"],
            textposition="outside",
        )
    )
    fig1.update_layout(
        title="TOP EJECUTIVOS",
        xaxis_title="VENTAS",
        yaxis_title="EJECUTIVO",
        height=420,
        yaxis=dict(categoryorder="total ascending"),
    )
    apply_plotly_theme(fig1)
    st.plotly_chart(fig1, width="stretch")

    st.markdown("---")

    st.markdown("### TOP Ventas Globales — (Color por Centro)")
    df_g = df_m_all.copy()

    g = (
        df_g.groupby(["EJECUTIVO", "CentroKey"], as_index=False)
        .size()
        .rename(columns={"size": "VENTAS"})
    )
    total_exec = g.groupby("EJECUTIVO", as_index=False)["VENTAS"].sum().rename(columns={"VENTAS": "VENTAS_TOTAL"})
    dominant = g.sort_values("VENTAS", ascending=False).drop_duplicates("EJECUTIVO")
    dom = dominant.merge(total_exec, on="EJECUTIVO", how="left")
    top_global = dom.sort_values("VENTAS_TOTAL", ascending=False).head(10)

    fig2 = px.bar(
        top_global,
        x="VENTAS_TOTAL",
        y="EJECUTIVO",
        orientation="h",
        text="VENTAS_TOTAL",
        color="CentroKey",
        title="TOP EJECUTIVOS GLOBAL",
        template=PLOTLY_TEMPLATE,
    )
    fig2.update_layout(
        height=460,
        yaxis=dict(categoryorder="total ascending"),
        xaxis_title="VENTAS",
        legend_title_text="CENTRO",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    fig2.update_traces(textposition="inside", insidetextanchor="end")
    st.plotly_chart(fig2, width="stretch")

    st.markdown("---")
    st.markdown("### TOP Ventas por Equipo (cada equipo = Supervisor) — por Centro")

    centro_equipo = st.selectbox(
        "Centro para TOP por Equipo",
        options=["CC2", "JV"],
        index=0,
        key="top_equipo_centro",
    )

    df_eq = df_m_all[df_m_all["CentroKey"] == centro_equipo].copy()

    if df_eq.empty:
        st.info(f"Sin datos para {centro_equipo} en el mes seleccionado.")
    else:
        df_eq_sup = df_eq.copy()
        df_eq_sup_nb = df_eq_sup[df_eq_sup["Supervisor"].astype(str).str.upper() != "BAJA"].copy()
        if not df_eq_sup_nb.empty:
            df_eq_sup = df_eq_sup_nb

        top_sup = (
            df_eq_sup.groupby("Supervisor", as_index=False)
            .size()
            .rename(columns={"size": "VENTAS"})
            .sort_values("VENTAS", ascending=False)
            .head(6)
        )

        fig_sup = go.Figure()
        fig_sup.add_trace(
            go.Bar(
                x=top_sup["VENTAS"],
                y=top_sup["Supervisor"],
                orientation="h",
                text=top_sup["VENTAS"],
                textposition="outside",
            )
        )
        fig_sup.update_layout(
            title="TOP SUPERVISORES",
            xaxis_title="VENTAS",
            yaxis_title="SUPERVISOR",
            height=330,
            yaxis=dict(categoryorder="total ascending"),
        )
        apply_plotly_theme(fig_sup)
        st.plotly_chart(fig_sup, width="stretch")

        equipos_n = 4 if centro_equipo == "CC2" else 2

        equipos_top = (
            df_eq_sup.groupby("Supervisor", as_index=False)
            .size()
            .rename(columns={"size": "VENTAS"})
            .sort_values("VENTAS", ascending=False)
        )

        equipos = equipos_top["Supervisor"].head(min(equipos_n, len(equipos_top))).tolist()

        if not equipos:
            st.info("No hay equipos para mostrar con los filtros actuales.")
        else:
            grid = st.columns(2, gap="large")
            for i, sup in enumerate(equipos):
                df_team = df_eq[df_eq["Supervisor"] == sup].copy()

                top_exec_team = (
                    df_team.groupby("EJECUTIVO", as_index=False)
                    .size()
                    .rename(columns={"size": "VENTAS"})
                    .sort_values("VENTAS", ascending=False)
                    .head(6)
                )

                fig_team = go.Figure()
                fig_team.add_trace(
                    go.Bar(
                        x=top_exec_team["EJECUTIVO"],
                        y=top_exec_team["VENTAS"],
                        text=top_exec_team["VENTAS"],
                        textposition="outside",
                    )
                )
                fig_team.update_layout(
                    title=f"EQUIPO: {sup}",
                    xaxis_title="",
                    yaxis_title="VENTAS",
                    height=320,
                )
                apply_plotly_theme(fig_team)
                fig_team.update_layout(margin=dict(l=20, r=20, t=60, b=100))
                fig_team.update_xaxes(tickangle=-35)

                with grid[i % 2]:
                    st.plotly_chart(fig_team, width="stretch", key=f"top_team_{centro_equipo}_{normalize_name(sup)}_{i}")

    st.caption(f"Día: {datetime.now().strftime('%d/%m/%Y')}")

# ======================================================
# TAB 8: METAS
# ======================================================
with tabs[7]:
    st.markdown(f"## Metas — **{mes_labels[mes_sel]}**")

    df_mes = filter_month(ventas, mes_sel)

    if center_sel:
        df_mes = df_mes[df_mes["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_mes = df_mes[df_mes["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_mes = df_mes[df_mes["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_mes = df_mes[df_mes["SUBREGION"].isin(sub_sel)]

    metas_centro = metas[metas["Nivel"].str.lower() == "centro"].copy()
    metas_sup = metas[metas["Nivel"].str.lower() == "supervisor"].copy()

    metas_centro_f = metas_centro.copy()
    if center_sel:
        metas_centro_f = metas_centro_f[metas_centro_f["Centro"].isin([c.upper() for c in center_sel])]

    # ✅ Supervisores ACTIVOS REALES (no jefes) = empleados activos cuyo Puesto contiene SUPERV
    emp_sup = empleados.copy()
    emp_sup["Estatus"] = emp_sup["Estatus"].astype(str).str.strip().str.upper()
    emp_sup["Puesto"] = emp_sup["Puesto"].astype(str).str.strip().str.upper()
    emp_sup["Nombre"] = emp_sup["Nombre"].astype(str).str.strip()
    emp_sup["Centro"] = emp_sup["Centro"].astype(str).str.strip()

    emp_sup = emp_sup[
        (emp_sup["Estatus"] == "ACTIVO")
        & (emp_sup["Puesto"].str.contains("SUPERV", na=False))
        & (emp_sup["Centro"].str.upper().str.contains("EXP ATT C CENTER", na=False))
    ].copy()

    emp_sup["CentroKey"] = np.where(
        emp_sup["Centro"].str.upper().str.contains("JUAREZ", na=False),
        "JV",
        "CC2",
    )
    emp_sup["Supervisor"] = emp_sup["Nombre"]
    emp_sup["Supervisor_norm"] = emp_sup["Supervisor"].apply(normalize_name)

    # ❌ Hide any "JULIO ..."
    emp_sup["FirstName"] = emp_sup["Supervisor_norm"].astype(str).str.split().str[0]
    emp_sup = emp_sup[emp_sup["FirstName"] != "JULIO"].copy()
    emp_sup.drop(columns=["FirstName"], inplace=True, errors="ignore")

    # Respect sidebar supervisor filter (if exists)
    sup_norm_filter = {normalize_name(s) for s in sup_sel} if sup_sel else None
    if sup_norm_filter:
        emp_sup = emp_sup[emp_sup["Supervisor_norm"].isin(sup_norm_filter)].copy()

    # Dedup
    emp_sup = emp_sup.drop_duplicates("Supervisor_norm", keep="first")

    active_supervisores_norm = set(emp_sup["Supervisor_norm"].dropna().tolist())

    metas_sup_show = metas_sup[metas_sup["Nombre_norm"].isin(active_supervisores_norm)].copy()
    if center_sel:
        metas_sup_show["Centro"] = metas_sup_show["Centro"].astype(str).str.strip().str.upper()
        metas_sup_show = metas_sup_show[metas_sup_show["Centro"].isin([c.upper() for c in center_sel])].copy()

    st.markdown("### Metas Globales")
    achieved_global = int(len(df_mes))

    if not metas_sup_show.empty:
        meta_global = int(pd.to_numeric(metas_sup_show["Meta"], errors="coerce").fillna(0).sum())
    else:
        meta_global = int(pd.to_numeric(metas_centro_f["Meta"], errors="coerce").fillna(0).sum()) if not metas_centro_f.empty else 0

    faltan_global = meta_global - achieved_global

    colg1, colg2 = st.columns([0.72, 0.28], gap="large")
    with colg1:
        fig = gauge_fig(achieved_global, meta_global, "VISOR GLOBAL")
        st.plotly_chart(fig, width="stretch", key="gauge_global")
    with colg2:
        metric_card("Meta Global", fmt_int(meta_global))
        metric_card("Alcanzado", fmt_int(achieved_global))
        metric_card("FALTAN", fmt_int(faltan_global))

    st.markdown("---")

    st.markdown("### Metas x Centro")
    if metas_centro_f.empty:
        st.info("No hay metas de Centro (con los filtros actuales).")
    else:
        ach_center = df_mes.groupby("CentroKey").size().to_dict()
        cols = st.columns(2, gap="large")

        for i, centro in enumerate(["CC2", "JV"]):
            row = metas_centro_f[metas_centro_f["Centro"] == centro].head(1)
            if row.empty:
                continue
            meta_val = int(pd.to_numeric(row["Meta"].iloc[0], errors="coerce") or 0)
            coord = str(row["Nombre"].iloc[0])
            achieved = int(ach_center.get(centro, 0))
            faltan = meta_val - achieved

            with cols[i]:
                fig = gauge_fig(achieved, meta_val, "VISOR DE METAS X COORDINADOR")
                st.plotly_chart(fig, width="stretch", key=f"gauge_centro_{centro}")
                st.markdown(f"**COORDINADOR:** {coord}")
                st.markdown(f"**FALTAN:** {fmt_int(faltan)}")

    st.markdown("---")

    # ------------------------------------------------------
    # Metas x Supervisor — SOLO ACTIVOS REALES (mismo concepto que TOPS por equipo)
    # - CC2 vs JV side-by-side
    # - Supervisores vienen de empleados (Puesto SUPERV, ACTIVO)
    # - Achieved viene de ventas del mes (df_mes)
    # ------------------------------------------------------
    st.markdown("### Metas x Supervisor — (Solo Activos, comparativo CC2 vs JV)")

    centers_to_show = center_sel if center_sel else ["CC2", "JV"]

    show_cc2 = "CC2" in centers_to_show
    show_jv = "JV" in centers_to_show

    if show_cc2 and show_jv:
        col_cc2, col_jv = st.columns(2, gap="large")
        center_slots = {"CC2": col_cc2, "JV": col_jv}
        inner_grid_cols = 1
    else:
        center_slots = {"CC2": st.container(), "JV": st.container()}
        inner_grid_cols = 2

    metas_sup_local = metas_sup.copy()
    metas_sup_local["Centro"] = metas_sup_local["Centro"].astype(str).str.strip().str.upper()

    for centro in ["CC2", "JV"]:
        if centro not in centers_to_show:
            continue

        with center_slots[centro]:
            st.markdown(f"#### {'CC2 (Center 2)' if centro=='CC2' else 'JV (Juárez)'}")

            cand = emp_sup[emp_sup["CentroKey"] == centro][["Supervisor", "Supervisor_norm"]].copy()
            if cand.empty:
                st.info("No hay supervisores ACTIVOS (puesto supervisor) para este centro con los filtros actuales.")
                continue

            df_c = df_mes[df_mes["CentroKey"] == centro].copy()
            ach_map = df_c.groupby("Supervisor_norm").size().to_dict() if not df_c.empty else {}

            cand["Achieved"] = cand["Supervisor_norm"].map(lambda n: int(ach_map.get(n, 0)))
            cand = cand.sort_values("Achieved", ascending=False).reset_index(drop=True)

            grid_cols = st.columns(inner_grid_cols, gap="large")

            for idx, rr in cand.iterrows():
                sup_name = str(rr["Supervisor"]).strip()
                sup_norm = str(rr["Supervisor_norm"]).strip()
                achieved = int(rr["Achieved"])

                meta_val = np.nan
                mrow = metas_sup_local[
                    (metas_sup_local["Nombre_norm"] == sup_norm) & (metas_sup_local["Centro"] == centro)
                ]
                if not mrow.empty:
                    mv = pd.to_numeric(mrow["Meta"].iloc[0], errors="coerce")
                    meta_val = float(mv) if pd.notna(mv) else np.nan

                faltan = (meta_val - achieved) if pd.notna(meta_val) else np.nan

                with grid_cols[idx % inner_grid_cols]:
                    fig = gauge_fig(
                        achieved,
                        meta_val if pd.notna(meta_val) else 0,
                        f"SUPERVISOR: {sup_name}",
                    )
                    st.plotly_chart(fig, width="stretch", key=f"meta_sup_{centro}_{sup_norm}_{idx}")

                    st.markdown(f"**Meta:** {fmt_int(meta_val) if pd.notna(meta_val) else '-'}")
                    st.markdown(f"**Alcanzado:** {fmt_int(achieved)}")
                    st.markdown(f"**FALTAN:** {fmt_int(faltan) if pd.notna(faltan) else '-'}")

    st.caption(f"Día: {datetime.now().strftime('%d/%m/%Y')}")

# ======================================================
# TAB 9: Tendencia x Ejecutivo
# ======================================================
with tabs[8]:
    st.markdown("## Tendencia x Ejecutivo")

    ej = st.selectbox("Ejecutivo", options=sorted(ventas["EJECUTIVO"].dropna().unique().tolist()))
    st.markdown(f"✅ Has seleccionado: **{ej}**")

    df_e = ventas[ventas["EJECUTIVO"] == ej].copy()
    if center_sel:
        df_e = df_e[df_e["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_e = df_e[df_e["Supervisor"].isin(sup_sel)]
    if sub_sel:
        df_e = df_e[df_e["SUBREGION"].isin(sub_sel)]

    m = (
        df_e.groupby("AñoMes", as_index=False)
        .size()
        .rename(columns={"size": "Ventas"})
        .sort_values("AñoMes")
    )

    prom = float(m[m["AñoMes"] < int(m["AñoMes"].max())]["Ventas"].mean()) if len(m) > 1 else np.nan

    c1, c2 = st.columns([0.65, 0.35], gap="large")

    with c1:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=m["AñoMes"], y=m["Ventas"], mode="lines+markers", name="Ventas"))
        if not np.isnan(prom):
            fig.add_trace(go.Scatter(x=m["AñoMes"], y=[prom] * len(m), mode="lines", name="Promedio (sin mes actual)"))
        fig.update_layout(
            title="Tendencia X Ejecutivo",
            xaxis_title="Mes",
            yaxis_title="Ventas",
            height=420,
        )
        apply_plotly_theme(fig)
        st.plotly_chart(fig, width="stretch")

    with c2:
        cur = int(mes_sel)
        cur_v = int(m[m["AñoMes"] == cur]["Ventas"].iloc[0]) if (m["AñoMes"] == cur).any() else 0
        metric_card("Ventas Mes", fmt_int(cur_v))
        if not np.isnan(prom):
            dif = cur_v - prom
            metric_card("Diferencia", f"{dif:+.0f} vs promedio")

        figb = px.bar(
            pd.DataFrame({"Mes": [mes_labels[mes_sel]], "Ventas": [cur_v]}),
            x="Mes",
            y="Ventas",
            title="Ventas X Ejecutivo",
            template=PLOTLY_TEMPLATE,
        )
        figb.update_layout(
            height=300,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(figb, width="stretch")

