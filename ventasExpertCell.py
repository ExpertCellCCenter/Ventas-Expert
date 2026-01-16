# app.py
import os
import base64
import json
import zlib
import unicodedata
from datetime import datetime, date

from io import BytesIO  # ✅ ADDED (Excel download)

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
# ✅ Pylance safe default (Streamlit can stop early with st.stop(), but Pylance doesn't know that)
center_sel: list[str] = ["CC2", "JV"]

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
# ✅ EXCEL EXPORT HELPERS (ADDED)
# -------------------------------
def _safe_sheet_name(name: str) -> str:
    name = str(name or "Sheet").strip()
    bad = [":", "\\", "/", "?", "*", "[", "]"]
    for b in bad:
        name = name.replace(b, " ")
    name = " ".join(name.split())
    return (name[:31] or "Sheet")


def build_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets.items():
            if df is None:
                continue
            df_to_write = df.copy()
            sheet = _safe_sheet_name(sname)
            df_to_write.to_excel(writer, index=False, sheet_name=sheet)

            try:
                from openpyxl.utils import get_column_letter
                ws = writer.sheets[sheet]
                max_rows = min(len(df_to_write), 500)
                for i, col in enumerate(df_to_write.columns, start=1):
                    col_vals = [str(col)]
                    if max_rows > 0:
                        col_vals += [str(v) for v in df_to_write[col].iloc[:max_rows].tolist()]
                    width = min(max(len(x) for x in col_vals) + 2, 55)
                    ws.column_dimensions[get_column_letter(i)].width = max(10, width)
            except Exception:
                pass

    return output.getvalue()


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
# METAS (MANUAL TABLE)
# -------------------------------
# ✅ EDIT METAS HERE (by hand)
METAS_MANUAL_ROWS = [
    # --- Metas Centro ---
    {"IDCenter": "CC1", "Nivel": "Centro", "Nombre": "EDUARDO AGUILA SANCHEZ", "Centro": "CC2", "Meta": 596},
    {"IDCenter": "JV1", "Nivel": "Centro", "Nombre": "MARIA LUISA MEZA GOEL",  "Centro": "JV",  "Meta": 258},

    # --- Metas Supervisor (JV) ---
    {"IDCenter": "JV2", "Nivel": "Supervisor", "Nombre": "JORGE MIGUEL UREÑA ZARATE",        "Centro": "JV",  "Meta": 120},
    {"IDCenter": "JV3", "Nivel": "Supervisor", "Nombre": "MARIA FERNANDA MARTINEZ BISTRAIN", "Centro": "JV",  "Meta": 138},

    # --- Metas Supervisor (CC2) ---
    {"IDCenter": "CC2", "Nivel": "Supervisor", "Nombre": "ALFREDO CABRERA PADRON",          "Centro": "CC2", "Meta": 156},
    {"IDCenter": "CC4", "Nivel": "Supervisor", "Nombre": "REYNA LIZZETTE MARTINEZ GARCIA",  "Centro": "CC2", "Meta": 132},
    {"IDCenter": "CC3", "Nivel": "Supervisor", "Nombre": "CARLOS ALBERTO AGUILAR CANO",  "Centro": "CC2", "Meta": 168},
    {"IDCenter": "CC5", "Nivel": "Supervisor", "Nombre": "ALAN UZIEL SALAZAR AGUILAR",     "Centro": "CC2", "Meta": 140},

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
        return pd.DataFrame(
            {
                "Fecha": pd.to_datetime([]),
                "Ventas": pd.Series([], dtype="int64"),
            }
        )
    out = df.groupby("Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
    out["Fecha"] = pd.to_datetime(out["Fecha"], errors="coerce")
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

# ✅ ONLY FOR FILTER OPTIONS (do NOT change data logic):
#    - remove BAJA supervisors (people no longer with you)
#    - remove EDUARDO AGUILA SANCHEZ from Supervisor filters (he is coordinator of supervisors)
EXCLUDED_SUP_NORMS = {normalize_name("BAJA"), normalize_name("EDUARDO AGUILA SANCHEZ")}
ventas_filtros = ventas[~ventas["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS)].copy()

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

# ✅ Prevent Streamlit from keeping excluded supervisors in session state selections
if "Supervisor" in st.session_state:
    st.session_state["Supervisor"] = [
        s for s in st.session_state["Supervisor"]
        if normalize_name(s) not in EXCLUDED_SUP_NORMS
    ]

# ✅ Supervisor options (exclude BAJA + Eduardo only in the filter list)
supervisores = sorted([s for s in ventas_filtros["Supervisor"].dropna().unique().tolist()])
sup_sel = st.sidebar.multiselect("Supervisor", options=supervisores, default=[])

# ✅ Ejecutivo options (DO NOT exclude BAJA here — exclusion is ONLY in Tendencia Ejecutivo tab)
ejecutivos = sorted([e for e in ventas["EJECUTIVO"].dropna().unique().tolist()])

if "Ejecutivo" in st.session_state:
    st.session_state["Ejecutivo"] = [e for e in st.session_state["Ejecutivo"] if e in ejecutivos]
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

    # --- Selection ARPU (day selected in the bar chart) ---
    sel_arpu_label = None
    sel_arpu_val = None
    sel_arpu_siva_val = None

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
        # ✅ Enable click/selection on bars to retrieve selected day(s)
        event = st.plotly_chart(
            fig,
            width="stretch",
            key=f"t0_global_mes_{mes_sel}",
            on_select="rerun",
            selection_mode="points",
        )

        # --- Extract selected day(s) robustly (prefer point_indices to avoid date parsing quirks) ---
        sel_dates = []

        try:
            idxs = list(getattr(event.selection, "point_indices", []))
        except Exception:
            idxs = []

        # Fallback if the object behaves more like a dict
        if not idxs:
            try:
                idxs = list(event["selection"]["point_indices"])
            except Exception:
                idxs = []

        # Map indices back to the daily series (s) to get the exact date(s)
        for ix in idxs:
            try:
                ix = int(ix)
                if 0 <= ix < len(s):
                    d = pd.to_datetime(s.iloc[ix]["Fecha"], errors="coerce")
                    if pd.notna(d):
                        sel_dates.append(d.date())
            except Exception:
                pass

        # Fallback: if no indices, try reading x directly from points
        if not sel_dates:
            try:
                pts = getattr(event.selection, "points", [])
            except Exception:
                pts = []
            if not pts:
                try:
                    pts = event.get("selection", {}).get("points", [])
                except Exception:
                    pts = []

            for p in pts:
                try:
                    xval = p.get("x", None)
                    d = pd.to_datetime(xval, errors="coerce")
                    if pd.notna(d):
                        sel_dates.append(d.date())
                except Exception:
                    pass

        if sel_dates:
            sel_dates = sorted(set(sel_dates))
            df_sel_day = df_base[df_base["Fecha"].isin(sel_dates)].copy()

            if len(sel_dates) == 1:
                sel_arpu_label = sel_dates[0].strftime("%d/%m/%Y")
            else:
                sel_arpu_label = f"{sel_dates[0].strftime('%d/%m/%Y')} → {sel_dates[-1].strftime('%d/%m/%Y')}"

            sel_arpu_val = arpu(df_sel_day)
            sel_arpu_siva_val = arpu_siva(df_sel_day)

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

        # ✅ DOWNLOAD EXCEL (Global Mes) — ADDED
        sheets_gm = {
            "Top dias (tabla)": top_days_show.copy(),
            "Serie diaria (grafica)": s.rename(columns={"Ventas": "Total de Ventas"}).copy(),
        }
        if sel_dates:
            sheets_gm["Seleccion (dia)"] = df_sel_day.copy()

        st.download_button(
            "⬇️ Descargar Excel (Global Mes)",
            data=build_excel_bytes(sheets_gm),
            file_name=f"GlobalMes_{mes_sel}_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_global_mes_{mes_sel}",
            use_container_width=True,
        )

        st.caption(f"🕒 Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    with colB:
        st.markdown("### ARPU")
        kpi_mini("ARPU", fmt_money_short(arpu(df_base)))
        kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_base)))

        # ✅ Show ARPU for the selected day in the bar chart (if any)
        if sel_arpu_label:
            st.markdown("#### ARPU del día seleccionado")
            kpi_mini(f"ARPU ({sel_arpu_label})", fmt_money_short(sel_arpu_val))
            kpi_mini(f"ARPU S/IVA ({sel_arpu_label})", fmt_money_short(sel_arpu_siva_val))
        else:
            st.caption("👆 Selecciona un día en la barra para ver el ARPU de ese día.")


# ======================================================
# TAB 2: Semanas
# ======================================================
with tabs[1]:
    st.markdown("## Vista por Semanas")

    # ✅ Export holders (ADDED)
    by_day_export = None
    cmp_export = None
    interval_summary_export = None
    hour_breakdown_export = None

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
    sem_choice = st.selectbox("Mes (Semanas)", options=sem_labels, index=0, key="sem_mes_choice")

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
    st.plotly_chart(figw, width="stretch", key="t1_semanas_general")

    st.markdown("### ARPU")
    kpi_mini("ARPU", fmt_money_short(arpu(df_sem)))
    kpi_mini("ARPU S/IVA", fmt_money_short(arpu_siva(df_sem)))

    # ==========================================================
    # ✅ NEW (TAB SEMANAS): Vista por meses y semanas (como en Mes vs Mes)
    # ==========================================================
    st.markdown("---")
    st.markdown("## ✅ Vista por meses y semanas")

    df_sem_mvw_ctx = ventas.copy()
    if center_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["EJECUTIVO"].isin(ej_sel)]
    if sub_sel:
        df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["SUBREGION"].isin(sub_sel)]

    df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["FECHA DE CAPTURA"].notna()].copy()
    df_sem_mvw_ctx["M_DT"] = pd.to_datetime(df_sem_mvw_ctx["FECHA DE CAPTURA"], errors="coerce")
    df_sem_mvw_ctx = df_sem_mvw_ctx[df_sem_mvw_ctx["M_DT"].notna()].copy()

    df_sem_mvw_ctx["M_Fecha"] = df_sem_mvw_ctx["M_DT"].dt.date
    df_sem_mvw_ctx["M_Hora"] = df_sem_mvw_ctx["M_DT"].dt.hour

    if df_sem_mvw_ctx.empty:
        st.info("No hay datos disponibles para la vista por meses/semanas con los filtros actuales.")
    else:
        df_sem_mvw_ctx["M_MonthKey"] = df_sem_mvw_ctx["M_DT"].dt.strftime("%Y-%m")
        df_sem_mvw_ctx["M_MonthName"] = df_sem_mvw_ctx["M_DT"].dt.strftime("%B")
        df_sem_mvw_ctx["M_MonthLabel"] = df_sem_mvw_ctx["M_MonthKey"] + " (" + df_sem_mvw_ctx["M_MonthName"] + ")"

        month_start = df_sem_mvw_ctx["M_DT"].dt.to_period("M").dt.to_timestamp()
        first_wd = month_start.dt.weekday
        df_sem_mvw_ctx["M_WeekOfMonth"] = ((df_sem_mvw_ctx["M_DT"].dt.day + first_wd - 1) // 7) + 1

        st.markdown("### Vista por meses y semanas (Ventas)")

        month_map = (
            df_sem_mvw_ctx[["M_MonthKey", "M_MonthLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values("M_MonthKey")
        )
        m_options = month_map["M_MonthLabel"].tolist()

        # Default: mes seleccionado en sidebar + (si existe) el mes anterior por orden
        def _ym_to_monthkey_from_ym(ym_int: int) -> str:
            y = ym_int // 100
            m = ym_int % 100
            return f"{y}-{m:02d}"

        cur_key = _ym_to_monthkey_from_ym(int(mes_sel)) if mes_sel else None
        keys_sorted = month_map["M_MonthKey"].tolist()

        defaults = []
        if cur_key and cur_key in keys_sorted:
            idx = keys_sorted.index(cur_key)
            defaults_keys = [keys_sorted[idx]]
            if idx - 1 >= 0:
                defaults_keys.insert(0, keys_sorted[idx - 1])
            defaults = month_map[month_map["M_MonthKey"].isin(defaults_keys)]["M_MonthLabel"].tolist()

        if not defaults:
            defaults = m_options[-2:] if len(m_options) >= 2 else m_options

        m_sel = st.multiselect(
            "Selecciona uno o más meses (Ventas)",
            options=m_options,
            default=defaults,
            key="sem_mvw_months_multi",
        )

        df_mvw = df_sem_mvw_ctx.copy()
        if m_sel:
            df_mvw = df_mvw[df_mvw["M_MonthLabel"].isin(m_sel)].copy()
        else:
            df_mvw = df_mvw.iloc[0:0].copy()

        if df_mvw.empty:
            st.info("No hay datos para los meses seleccionados.")
        else:
            df_mvw["M_WeekLabel"] = df_mvw["M_MonthLabel"] + " - Semana " + df_mvw["M_WeekOfMonth"].astype(int).astype(str)

            w_map = (
                df_mvw[["M_MonthKey", "M_MonthLabel", "M_WeekOfMonth", "M_WeekLabel"]]
                .dropna()
                .drop_duplicates()
                .sort_values(["M_MonthKey", "M_WeekOfMonth"])
            )
            w_options = w_map["M_WeekLabel"].tolist()

            w_sel = st.multiselect(
                "Selecciona Semana(s) del mes (Ventas)",
                options=w_options,
                default=w_options,
                key="sem_mvw_weeks_multi",
            )

            if w_sel:
                df_mvw = df_mvw[df_mvw["M_WeekLabel"].isin(w_sel)].copy()
            else:
                df_mvw = df_mvw.iloc[0:0].copy()

            if df_mvw.empty:
                st.info("No hay datos para las semanas seleccionadas.")
            else:
                by_day = df_mvw.groupby("M_Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
                by_day_export = by_day.copy()  # ✅ ADDED (export)
                fig_mvw = px.bar(
                    by_day,
                    x="M_Fecha",
                    y="Ventas",
                    title="Total por día (Ventas) — filtro por Mes(es) y Semana(s)",
                    labels={"Ventas": "Ventas", "M_Fecha": "Fecha"},
                    template=PLOTLY_TEMPLATE,
                )
                fig_mvw.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_mvw, width="stretch", key="t1_sem_mvw_bar")

                st.markdown("### Comparativo día vs día (entre meses seleccionados) — Ventas")

                df_mvw["M_DiaDelMes"] = df_mvw["M_DT"].dt.day
                cmp = (
                    df_mvw.groupby(["M_MonthLabel", "M_DiaDelMes"], as_index=False)
                    .size()
                    .rename(columns={"size": "Ventas"})
                )
                cmp_export = cmp.copy()  # ✅ ADDED (export)

                fig_cmp = px.line(
                    cmp,
                    x="M_DiaDelMes",
                    y="Ventas",
                    color="M_MonthLabel",
                    markers=True,
                    title="Comparativo por día del mes (Ventas)",
                    labels={"M_DiaDelMes": "Día del mes", "Ventas": "Ventas", "M_MonthLabel": "Mes"},
                    template=PLOTLY_TEMPLATE,
                )
                fig_cmp.update_xaxes(dtick=1)
                fig_cmp.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_cmp, width="stretch", key="t1_sem_mvw_day_vs_day")

                st.markdown("#### Comparar dos intervalos de tiempo ")

                def _trim_time_to_minute_sem(tobj):
                    try:
                        return tobj.replace(second=0, microsecond=0)
                    except Exception:
                        return tobj

                def _t_sem(h: int, m: int):
                    return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M").time()

                avail_dts = df_mvw["M_DT"].dropna()
                avail_dates = sorted(df_mvw["M_Fecha"].dropna().unique().tolist())

                if avail_dts.empty or not avail_dates:
                    st.info("No hay fechas/horas disponibles para comparar con los filtros actuales (Semanas).")
                else:
                    min_dt = avail_dts.min().to_pydatetime()
                    max_dt = avail_dts.max().to_pydatetime()

                    d_def_2 = avail_dates[-1]
                    d_def_1 = avail_dates[-2] if len(avail_dates) >= 2 else avail_dates[-1]

                    s1_def = max(datetime.combine(d_def_1, _t_sem(0, 0)), min_dt)
                    e1_def = min(datetime.combine(d_def_1, _t_sem(23, 59)), max_dt)
                    s2_def = max(datetime.combine(d_def_2, _t_sem(0, 0)), min_dt)
                    e2_def = min(datetime.combine(d_def_2, _t_sem(23, 59)), max_dt)

                    if s1_def > e1_def:
                        s1_def, e1_def = min_dt, max_dt
                    if s2_def > e2_def:
                        s2_def, e2_def = min_dt, max_dt

                    cA, cB = st.columns(2, gap="large")

                    with cA:
                        st.markdown("**Fecha 1 (intervalo a comparar)**")
                        a1, a2 = st.columns(2)
                        with a1:
                            s1_date = st.date_input(
                                "Inicio (Fecha 1) — Semanas",
                                value=s1_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i1_start_date",
                            )
                        with a2:
                            s1_time = st.time_input(
                                "Inicio (Hora 1) — Semanas",
                                value=_trim_time_to_minute_sem(s1_def.time()),
                                key="sem_mvw_i1_start_time",
                            )
                        b1, b2 = st.columns(2)
                        with b1:
                            e1_date = st.date_input(
                                "Fin (Fecha 1) — Semanas",
                                value=e1_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i1_end_date",
                            )
                        with b2:
                            e1_time = st.time_input(
                                "Fin (Hora 1) — Semanas",
                                value=_trim_time_to_minute_sem(e1_def.time()),
                                key="sem_mvw_i1_end_time",
                            )

                    with cB:
                        st.markdown("**Fecha 2 (intervalo a comparar)**")
                        a1, a2 = st.columns(2)
                        with a1:
                            s2_date = st.date_input(
                                "Inicio (Fecha 2) — Semanas",
                                value=s2_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i2_start_date",
                            )
                        with a2:
                            s2_time = st.time_input(
                                "Inicio (Hora 2) — Semanas",
                                value=_trim_time_to_minute_sem(s2_def.time()),
                                key="sem_mvw_i2_start_time",
                            )
                        b1, b2 = st.columns(2)
                        with b1:
                            e2_date = st.date_input(
                                "Fin (Fecha 2) — Semanas",
                                value=e2_def.date(),
                                min_value=min_dt.date(),
                                max_value=max_dt.date(),
                                key="sem_mvw_i2_end_date",
                            )
                        with b2:
                            e2_time = st.time_input(
                                "Fin (Hora 2) — Semanas",
                                value=_trim_time_to_minute_sem(e2_def.time()),
                                key="sem_mvw_i2_end_time",
                            )

                    s1 = datetime.combine(s1_date, s1_time)
                    e1 = datetime.combine(e1_date, e1_time)
                    s2 = datetime.combine(s2_date, s2_time)
                    e2 = datetime.combine(e2_date, e2_time)

                    if s1 < min_dt: s1 = min_dt
                    if e1 > max_dt: e1 = max_dt
                    if s2 < min_dt: s2 = min_dt
                    if e2 > max_dt: e2 = max_dt

                    if s1 > e1:
                        st.warning("En Fecha 1 (Semanas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                        s1, e1 = e1, s1
                    if s2 > e2:
                        st.warning("En Fecha 2 (Semanas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                        s2, e2 = e2, s2

                    df_d1 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s1)) & (df_mvw["M_DT"] <= pd.Timestamp(e1))].copy()
                    df_d2 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s2)) & (df_mvw["M_DT"] <= pd.Timestamp(e2))].copy()

                    v1 = int(df_d1.shape[0])
                    v2 = int(df_d2.shape[0])

                    m1 = float(df_d1["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d1.columns else 0.0
                    m2 = float(df_d2["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d2.columns else 0.0

                    k1, k2, k3, k4 = st.columns(4, gap="medium")
                    with k1:
                        metric_card("Ventas (Fecha 1)", fmt_int(v1), sub=f"{s1:%d/%m/%Y %H:%M} → {e1:%d/%m/%Y %H:%M}")
                    with k2:
                        metric_card("Ventas (Fecha 2)", fmt_int(v2), sub=f"{s2:%d/%m/%Y %H:%M} → {e2:%d/%m/%Y %H:%M}")
                    with k3:
                        metric_card("Diferencia (1-2)", f"{(v1 - v2):+,}")
                    with k4:
                        metric_card("Diferencia Monto (1-2)", f"{(m1 - m2):+,.2f}")

                    comp_df = pd.DataFrame(
                        {
                            "Comparación": ["Fecha 1", "Fecha 2"],
                            "Ventas": [v1, v2],
                            "MontoVendido": [m1, m2],
                            "Inicio": [s1, s2],
                            "Fin": [e1, e2],
                        }
                    )
                    interval_summary_export = comp_df.copy()  # ✅ ADDED (export)

                    fig_dates = px.bar(
                        comp_df,
                        x="Comparación",
                        y="Ventas",
                        title="Comparativo Ventas — Fecha 1 vs Fecha 2",
                        labels={"Ventas": "Ventas"},
                        hover_data={"Inicio": True, "Fin": True, "MontoVendido": True, "Comparación": False},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_dates.update_layout(height=360, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_dates, width="stretch", key="t1_sem_mvw_interval_bar")

                    df_d1["H"] = pd.to_datetime(df_d1["M_DT"], errors="coerce").dt.hour
                    df_d2["H"] = pd.to_datetime(df_d2["M_DT"], errors="coerce").dt.hour

                    h1 = df_d1.groupby("H").size()
                    h2 = df_d2.groupby("H").size()

                    hours = list(range(0, 24))
                    hour_df = pd.DataFrame(
                        {
                            "Hora": hours,
                            "Fecha 1": [int(h1.get(h, 0)) for h in hours],
                            "Fecha 2": [int(h2.get(h, 0)) for h in hours],
                        }
                    )
                    hour_breakdown_export = hour_df.copy()  # ✅ ADDED (export)

                    hour_long = hour_df.melt(id_vars="Hora", var_name="Fecha", value_name="Ventas")

                    fig_hour2 = px.bar(
                        hour_long,
                        x="Hora",
                        y="Ventas",
                        color="Fecha",
                        barmode="group",
                        title="Comparativo por hora — Fecha 1 vs Fecha 2",
                        labels={"Ventas": "Ventas", "Hora": "Hora"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_hour2.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_hour2, width="stretch", key="t1_sem_mvw_interval_hour")

    # ✅ DOWNLOAD EXCEL (Semanas) — ADDED
    filtros_sem = pd.DataFrame(
        {
            "Filtro": ["Centro", "Supervisor", "Ejecutivo", "Subregión", "Mes (Semanas)"],
            "Valor": [
                ", ".join(center_sel) if center_sel else "Todos",
                ", ".join(sup_sel) if sup_sel else "Todos",
                ", ".join(ej_sel) if ej_sel else "Todos",
                ", ".join(sub_sel) if sub_sel else "Todos",
                str(sem_choice),
            ],
        }
    )

    sheets_sem = {
        "Filtros": filtros_sem,
        "Semanas (serie)": w.copy(),
    }
    if by_day_export is not None:
        sheets_sem["MesSemana - Total por dia"] = by_day_export
    if cmp_export is not None:
        sheets_sem["Comparativo dia vs dia"] = cmp_export
    if interval_summary_export is not None:
        sheets_sem["Intervalo resumen"] = interval_summary_export
    if hour_breakdown_export is not None:
        sheets_sem["Intervalo por hora"] = hour_breakdown_export

    st.download_button(
        "⬇️ Descargar Excel (Semanas)",
        data=build_excel_bytes(sheets_sem),
        file_name=f"Semanas_{mes_sel}_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_semanas_{mes_sel}",
        use_container_width=True,
    )

# ======================================================
# TAB 3: JV vs CC2
# ======================================================
with tabs[2]:
    st.markdown(f"## Ventas del Mes — Comparativo (JV vs CC2) — **{mes_labels[mes_sel]}**")

    df_m = filter_month(ventas, mes_sel)
    if sub_sel:
        df_m = df_m[df_m["SUBREGION"].isin(sub_sel)]
    if sup_sel:
        df_m = df_m[df_m["Supervisor"].isin(sup_sel)]
    if ej_sel:
        df_m = df_m[df_m["EJECUTIVO"].isin(ej_sel)]
    if center_sel:
        df_m = df_m[df_m["CentroKey"].isin(center_sel)]

    df_jv = df_m[df_m["CentroKey"] == "JV"].copy()
    df_cc2 = df_m[df_m["CentroKey"] == "CC2"].copy()

    cL, cR = st.columns(2, gap="large")

    def render_group(col, title, df_, key_prefix: str):
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
            st.plotly_chart(fig, width="stretch", key=f"t2_{key_prefix}_general_{mes_sel}")

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

    render_group(cL, "JV (Juárez)", df_jv, key_prefix="jv")
    render_group(cR, "CC2 (Center 2)", df_cc2, key_prefix="cc2")

# ======================================================
# TAB 4: Mes vs Mes  ✅ + Vista por meses/semanas + día vs día + intervalo
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
                key="mvm_mes_actual",
            )
        with colS2:
            mes_comp = st.selectbox(
                "Mes Comparado",
                options=months_all,
                format_func=lambda ym: mes_labels.get(ym, str(ym)),
                index=max(0, len(months_all) - 2),
                key="mvm_mes_comp",
            )
        with colS3:
            modo = st.selectbox(
                "ModoComparación",
                options=[0, 1],
                format_func=lambda x: "0 = Mes completo" if x == 0 else "1 = Cortar al día de hoy",
                key="mvm_modo",
            )

        dfA = filter_month(ventas, mes_actual)
        dfB = filter_month(ventas, mes_comp)

        if center_sel:
            dfA = dfA[dfA["CentroKey"].isin(center_sel)]
            dfB = dfB[dfB["CentroKey"].isin(center_sel)]
        if sup_sel:
            dfA = dfA[dfA["Supervisor"].isin(sup_sel)]
            dfB = dfB[dfB["Supervisor"].isin(sup_sel)]
        if ej_sel:
            dfA = dfA[dfA["EJECUTIVO"].isin(ej_sel)]
            dfB = dfB[dfB["EJECUTIVO"].isin(ej_sel)]
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
            st.plotly_chart(figA, width="stretch", key=f"t3_mesA_{mes_actual}_{modo}")

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
            st.plotly_chart(figB, width="stretch", key=f"t3_mesB_{mes_comp}_{modo}")

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

        # ==========================================================
        # ✅ NEW: Vista por meses y semanas + comparativo día vs día + intervalo (como tu ejemplo)
        # ==========================================================
        st.markdown("---")
        st.markdown("## ✅ Vista por meses y semanas ")

        df_mvm_ctx = ventas.copy()
        if center_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["CentroKey"].isin(center_sel)]
        if sup_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["Supervisor"].isin(sup_sel)]
        if ej_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["EJECUTIVO"].isin(ej_sel)]
        if sub_sel:
            df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["SUBREGION"].isin(sub_sel)]

        df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["FECHA DE CAPTURA"].notna()].copy()
        df_mvm_ctx["M_DT"] = pd.to_datetime(df_mvm_ctx["FECHA DE CAPTURA"], errors="coerce")
        df_mvm_ctx = df_mvm_ctx[df_mvm_ctx["M_DT"].notna()].copy()

        df_mvm_ctx["M_Fecha"] = df_mvm_ctx["M_DT"].dt.date
        df_mvm_ctx["M_Hora"] = df_mvm_ctx["M_DT"].dt.hour

        if df_mvm_ctx.empty:
            st.info("No hay datos disponibles para la vista por meses/semanas con los filtros actuales.")
        else:
            df_mvm_ctx["M_MonthKey"] = df_mvm_ctx["M_DT"].dt.strftime("%Y-%m")
            df_mvm_ctx["M_MonthName"] = df_mvm_ctx["M_DT"].dt.strftime("%B")
            df_mvm_ctx["M_MonthLabel"] = df_mvm_ctx["M_MonthKey"] + " (" + df_mvm_ctx["M_MonthName"] + ")"

            month_start = df_mvm_ctx["M_DT"].dt.to_period("M").dt.to_timestamp()
            first_wd = month_start.dt.weekday
            df_mvm_ctx["M_WeekOfMonth"] = ((df_mvm_ctx["M_DT"].dt.day + first_wd - 1) // 7) + 1

            st.markdown("### Vista por meses y semanas")

            m_options = sorted(df_mvm_ctx["M_MonthLabel"].dropna().unique().tolist())

            # Default: los 2 meses del comparativo (si existen), si no: todos
            def _ym_to_monthkey(ym: int) -> str:
                y = ym // 100
                m = ym % 100
                return f"{y}-{m:02d}"

            want_keys = {_ym_to_monthkey(int(mes_actual)), _ym_to_monthkey(int(mes_comp))}
            defaults = [x for x in m_options if str(x).startswith(tuple(sorted(want_keys)))]
            if not defaults:
                defaults = m_options

            m_sel = st.multiselect(
                "Selecciona uno o más meses (Ventas)",
                options=m_options,
                default=defaults,
                key="mvm_months_multi",
            )

            df_mvw = df_mvm_ctx.copy()
            if m_sel:
                df_mvw = df_mvw[df_mvw["M_MonthLabel"].isin(m_sel)].copy()
            else:
                df_mvw = df_mvw.iloc[0:0].copy()

            if df_mvw.empty:
                st.info("No hay datos para los meses seleccionados.")
            else:
                df_mvw["M_WeekLabel"] = df_mvw["M_MonthLabel"] + " - Semana " + df_mvw["M_WeekOfMonth"].astype(int).astype(str)
                w_options = sorted(df_mvw["M_WeekLabel"].dropna().unique().tolist())
                w_sel_default = w_options

                w_sel = st.multiselect(
                    "Selecciona Semana(s) del mes (Ventas)",
                    options=w_options,
                    default=w_sel_default,
                    key="mvm_weeks_multi",
                )

                if w_sel:
                    df_mvw = df_mvw[df_mvw["M_WeekLabel"].isin(w_sel)].copy()
                else:
                    df_mvw = df_mvw.iloc[0:0].copy()

                if df_mvw.empty:
                    st.info("No hay datos para las semanas seleccionadas.")
                else:
                    by_day = df_mvw.groupby("M_Fecha", as_index=False).size().rename(columns={"size": "Ventas"})
                    fig_mvw = px.bar(
                        by_day,
                        x="M_Fecha",
                        y="Ventas",
                        title="Total por día (Ventas) — filtro por Mes(es) y Semana(s)",
                        labels={"Ventas": "Ventas", "M_Fecha": "Fecha"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_mvw.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_mvw, width="stretch", key="t3_mvw_bar")

                    st.markdown("### Comparativo día vs día (mes contra mes) — Ventas")

                    df_mvw["M_DiaDelMes"] = df_mvw["M_DT"].dt.day
                    cmp = (
                        df_mvw.groupby(["M_MonthLabel", "M_DiaDelMes"], as_index=False)
                        .size()
                        .rename(columns={"size": "Ventas"})
                    )

                    fig_cmp = px.line(
                        cmp,
                        x="M_DiaDelMes",
                        y="Ventas",
                        color="M_MonthLabel",
                        markers=True,
                        title="Comparativo por día del mes (Ventas)",
                        labels={"M_DiaDelMes": "Día del mes", "Ventas": "Ventas", "M_MonthLabel": "Mes"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig_cmp.update_xaxes(dtick=1)
                    fig_cmp.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_cmp, width="stretch", key="t3_mvw_day_vs_day")

                    st.markdown("#### Comparar dos intervalos de tiempo")

                    def _trim_time_to_minute(tobj):
                        try:
                            return tobj.replace(second=0, microsecond=0)
                        except Exception:
                            return tobj

                    def _t(h: int, m: int):
                        return datetime.strptime(f"{h:02d}:{m:02d}", "%H:%M").time()

                    avail_dts = df_mvw["M_DT"].dropna()
                    avail_dates = sorted(df_mvw["M_Fecha"].dropna().unique().tolist())

                    if avail_dts.empty or not avail_dates:
                        st.info("No hay fechas/horas disponibles para comparar con los filtros actuales (Ventas).")
                    else:
                        min_dt = avail_dts.min().to_pydatetime()
                        max_dt = avail_dts.max().to_pydatetime()

                        d_def_2 = avail_dates[-1]
                        d_def_1 = avail_dates[-2] if len(avail_dates) >= 2 else avail_dates[-1]

                        s1_def = max(datetime.combine(d_def_1, _t(0, 0)), min_dt)
                        e1_def = min(datetime.combine(d_def_1, _t(23, 59)), max_dt)
                        s2_def = max(datetime.combine(d_def_2, _t(0, 0)), min_dt)
                        e2_def = min(datetime.combine(d_def_2, _t(23, 59)), max_dt)

                        if s1_def > e1_def:
                            s1_def, e1_def = min_dt, max_dt
                        if s2_def > e2_def:
                            s2_def, e2_def = min_dt, max_dt

                        cA, cB = st.columns(2, gap="large")

                        with cA:
                            st.markdown("**Fecha 1 (intervalo a comparar)**")
                            a1, a2 = st.columns(2)
                            with a1:
                                s1_date = st.date_input(
                                    "Inicio (Fecha 1) — Ventas",
                                    value=s1_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i1_start_date",
                                )
                            with a2:
                                s1_time = st.time_input(
                                    "Inicio (Hora 1) — Ventas",
                                    value=_trim_time_to_minute(s1_def.time()),
                                    key="mvm_i1_start_time",
                                )
                            b1, b2 = st.columns(2)
                            with b1:
                                e1_date = st.date_input(
                                    "Fin (Fecha 1) — Ventas",
                                    value=e1_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i1_end_date",
                                )
                            with b2:
                                e1_time = st.time_input(
                                    "Fin (Hora 1) — Ventas",
                                    value=_trim_time_to_minute(e1_def.time()),
                                    key="mvm_i1_end_time",
                                )

                        with cB:
                            st.markdown("**Fecha 2 (intervalo a comparar)**")
                            a1, a2 = st.columns(2)
                            with a1:
                                s2_date = st.date_input(
                                    "Inicio (Fecha 2) — Ventas",
                                    value=s2_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i2_start_date",
                                )
                            with a2:
                                s2_time = st.time_input(
                                    "Inicio (Hora 2) — Ventas",
                                    value=_trim_time_to_minute(s2_def.time()),
                                    key="mvm_i2_start_time",
                                )
                            b1, b2 = st.columns(2)
                            with b1:
                                e2_date = st.date_input(
                                    "Fin (Fecha 2) — Ventas",
                                    value=e2_def.date(),
                                    min_value=min_dt.date(),
                                    max_value=max_dt.date(),
                                    key="mvm_i2_end_date",
                                )
                            with b2:
                                e2_time = st.time_input(
                                    "Fin (Hora 2) — Ventas",
                                    value=_trim_time_to_minute(e2_def.time()),
                                    key="mvm_i2_end_time",
                                )

                        s1 = datetime.combine(s1_date, s1_time)
                        e1 = datetime.combine(e1_date, e1_time)
                        s2 = datetime.combine(s2_date, s2_time)
                        e2 = datetime.combine(e2_date, e2_time)

                        if s1 < min_dt: s1 = min_dt
                        if e1 > max_dt: e1 = max_dt
                        if s2 < min_dt: s2 = min_dt
                        if e2 > max_dt: e2 = max_dt

                        if s1 > e1:
                            st.warning("En Fecha 1 (Ventas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                            s1, e1 = e1, s1
                        if s2 > e2:
                            st.warning("En Fecha 2 (Ventas), el inicio es mayor que el fin. Se ajustó automáticamente.")
                            s2, e2 = e2, s2

                        df_d1 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s1)) & (df_mvw["M_DT"] <= pd.Timestamp(e1))].copy()
                        df_d2 = df_mvw[(df_mvw["M_DT"] >= pd.Timestamp(s2)) & (df_mvw["M_DT"] <= pd.Timestamp(e2))].copy()

                        v1 = int(df_d1.shape[0])
                        v2 = int(df_d2.shape[0])

                        m1 = float(df_d1["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d1.columns else 0.0
                        m2 = float(df_d2["PRECIO"].sum(skipna=True)) if "PRECIO" in df_d2.columns else 0.0

                        k1, k2, k3, k4 = st.columns(4, gap="medium")
                        with k1:
                            metric_card("Ventas (Fecha 1)", fmt_int(v1), sub=f"{s1:%d/%m/%Y %H:%M} → {e1:%d/%m/%Y %H:%M}")
                        with k2:
                            metric_card("Ventas (Fecha 2)", fmt_int(v2), sub=f"{s2:%d/%m/%Y %H:%M} → {e2:%d/%m/%Y %H:%M}")
                        with k3:
                            metric_card("Diferencia (1-2)", f"{(v1 - v2):+,}")
                        with k4:
                            metric_card("Diferencia Monto (1-2)", f"{(m1 - m2):+,.2f}")

                        comp_df = pd.DataFrame(
                            {
                                "Comparación": ["Fecha 1", "Fecha 2"],
                                "Ventas": [v1, v2],
                                "MontoVendido": [m1, m2],
                                "Inicio": [s1, s2],
                                "Fin": [e1, e2],
                            }
                        )
                        fig_dates = px.bar(
                            comp_df,
                            x="Comparación",
                            y="Ventas",
                            title="Comparativo Ventas — Fecha 1 vs Fecha 2",
                            labels={"Ventas": "Ventas"},
                            hover_data={"Inicio": True, "Fin": True, "MontoVendido": True, "Comparación": False},
                            template=PLOTLY_TEMPLATE,
                        )
                        fig_dates.update_layout(height=360, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_dates, width="stretch", key="t3_mvw_interval_bar")

                        df_d1["H"] = pd.to_datetime(df_d1["M_DT"], errors="coerce").dt.hour
                        df_d2["H"] = pd.to_datetime(df_d2["M_DT"], errors="coerce").dt.hour

                        h1 = df_d1.groupby("H").size()
                        h2 = df_d2.groupby("H").size()

                        hours = list(range(0, 24))
                        hour_df = pd.DataFrame(
                            {
                                "Hora": hours,
                                "Fecha 1": [int(h1.get(h, 0)) for h in hours],
                                "Fecha 2": [int(h2.get(h, 0)) for h in hours],
                            }
                        )
                        hour_long = hour_df.melt(id_vars="Hora", var_name="Fecha", value_name="Ventas")

                        fig_hour = px.bar(
                            hour_long,
                            x="Hora",
                            y="Ventas",
                            color="Fecha",
                            barmode="group",
                            title="Comparativo por hora — Fecha 1 vs Fecha 2",
                            labels={"Ventas": "Ventas", "Hora": "Hora"},
                            template=PLOTLY_TEMPLATE,
                        )
                        fig_hour.update_layout(height=380, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_hour, width="stretch", key="t3_mvw_interval_hour")

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
            st.plotly_chart(fig_monto, width="stretch", key="t4_pie_monto")

    with pieR:
        if fig_arpu is None:
            st.info("Sin ARPU suficiente para graficar % ARPU.")
        else:
            st.plotly_chart(fig_arpu, width="stretch", key="t4_pie_arpu")

    # ✅ DOWNLOAD EXCEL (Detalle) — ADDED
    sheets_det = {"Contexto (df_base mes)": df_d.copy()}
    try:
        if not mat_cc2.empty:
            sheets_det["Detalle CC2 (tabla)"] = mat_cc2_show.copy()
    except Exception:
        pass
    try:
        if not mat_jv.empty:
            sheets_det["Detalle JV (tabla)"] = mat_jv_show.copy()
    except Exception:
        pass

    st.download_button(
        "⬇️ Descargar Excel (Detalle)",
        data=build_excel_bytes(sheets_det),
        file_name=f"Detalle_{mes_sel}_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_detalle_{mes_sel}",
        use_container_width=True,
    )

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
        st.plotly_chart(fig, width="stretch", key="t5_region_pie")

# ======================================================
# TAB 7: TOPS
# ======================================================
with tabs[6]:
    st.markdown(f"## TOP Ventas — **{mes_labels[mes_sel]}**")

    df_m_all = filter_month(ventas, mes_sel)

    st.markdown("### TOP Ventas x Ejecutivo (por Centro)")
    centro_top = st.selectbox("Centro para TOP Ejecutivo", options=["JV", "CC2"], index=0, key="top_exec_center")
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
    st.plotly_chart(fig1, width="stretch", key=f"t6_top_ej_{centro_top}_{mes_sel}")

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
    st.plotly_chart(fig2, width="stretch", key=f"t6_top_global_{mes_sel}")

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
        st.plotly_chart(fig_sup, width="stretch", key=f"t6_top_sup_{centro_equipo}_{mes_sel}")

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

    # Metas x Supervisor — SOLO ACTIVOS REALES
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
#   ✅ UPDATED: Sanity check now follows "Filtro por meses y semanas"
#              - If multiple months selected, Meta = sum(meta per month)
#              - Ventas = ventas inside the selected months+weeks interval
#              - Meta per month is computed using the 1st day of EACH month
#              - Antigüedad is calendar days from ingreso to 1st day of month (NOT workable days)
#
#   ✅ FIXED WARNINGS (Pylance):
#      - "ventas" is not defined
#      - "filter_month" is not defined
#      (Only for type-checking; does NOT change runtime behavior)
#
#   ✅ IMPLEMENTED:
#      - Interval team plot now colors bars: GREEN=HEALTHY, RED=RISKY
# ======================================================
with tabs[8]:
    st.markdown("## Tendencia x Ejecutivo")

    # -------------------------------
    # ✅ Pylance warnings fix (no runtime impact)
    # -------------------------------
    from typing import TYPE_CHECKING
    if TYPE_CHECKING:
        ventas: pd.DataFrame
        filter_month: str
        tabs: list

    # 1) Base context (same global filters you already apply in this tab)
    df_ctx = ventas.copy()
    if center_sel:
        df_ctx = df_ctx[df_ctx["CentroKey"].isin(center_sel)]
    if sup_sel:
        df_ctx = df_ctx[df_ctx["Supervisor"].isin(sup_sel)]
    if sub_sel:
        df_ctx = df_ctx[df_ctx["SUBREGION"].isin(sub_sel)]

    # ======================================================
    # ✅ Filtro por meses y semanas (controla TODO este TAB)
    # ======================================================
    st.markdown("### Filtro por meses y semanas (solo esta pestaña)")

    df_mvw_ctx = df_ctx.copy()
    df_mvw_ctx = df_mvw_ctx[df_mvw_ctx["FECHA DE CAPTURA"].notna()].copy()
    df_mvw_ctx["T_DT"] = pd.to_datetime(df_mvw_ctx["FECHA DE CAPTURA"], errors="coerce")
    df_mvw_ctx = df_mvw_ctx[df_mvw_ctx["T_DT"].notna()].copy()

    if df_mvw_ctx.empty:
        st.info("No hay datos con fecha válida para aplicar el filtro de meses/semanas en esta pestaña.")
        df_ctx = df_ctx.iloc[0:0].copy()
    else:
        df_mvw_ctx["T_MonthKey"] = df_mvw_ctx["T_DT"].dt.strftime("%Y-%m")
        df_mvw_ctx["T_MonthName"] = df_mvw_ctx["T_DT"].dt.strftime("%B")
        df_mvw_ctx["T_MonthLabel"] = df_mvw_ctx["T_MonthKey"] + " (" + df_mvw_ctx["T_MonthName"] + ")"

        # Semana del mes
        month_start = df_mvw_ctx["T_DT"].dt.to_period("M").dt.to_timestamp()
        first_wd = month_start.dt.weekday
        df_mvw_ctx["T_WeekOfMonth"] = ((df_mvw_ctx["T_DT"].dt.day + first_wd - 1) // 7) + 1
        df_mvw_ctx["T_WeekLabel"] = (
            df_mvw_ctx["T_MonthLabel"] + " - Semana " + df_mvw_ctx["T_WeekOfMonth"].astype(int).astype(str)
        )

        month_map = (
            df_mvw_ctx[["T_MonthKey", "T_MonthLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values("T_MonthKey")
        )
        m_options = month_map["T_MonthLabel"].tolist()

        # Default: mes_sel + mes anterior (si existe)
        def _ym_to_monthkey_from_ym(ym_int: int) -> str:
            y = ym_int // 100
            m = ym_int % 100
            return f"{y}-{m:02d}"

        defaults = []
        try:
            cur_key = _ym_to_monthkey_from_ym(int(mes_sel)) if mes_sel else None
        except Exception:
            cur_key = None

        keys_sorted = month_map["T_MonthKey"].tolist()
        if cur_key and cur_key in keys_sorted:
            idx = keys_sorted.index(cur_key)
            defaults_keys = [keys_sorted[idx]]
            if idx - 1 >= 0:
                defaults_keys.insert(0, keys_sorted[idx - 1])
            defaults = month_map[month_map["T_MonthKey"].isin(defaults_keys)]["T_MonthLabel"].tolist()

        if not defaults:
            defaults = m_options[-2:] if len(m_options) >= 2 else m_options

        m_sel = st.multiselect(
            "Selecciona uno o más meses (Tendencia Ejecutivo)",
            options=m_options,
            default=defaults,
            key="tend_mvw_months_multi",
        )

        df_f = df_mvw_ctx.copy()
        if m_sel:
            df_f = df_f[df_f["T_MonthLabel"].isin(m_sel)].copy()

        # Semanas disponibles según meses elegidos
        w_map = (
            df_f[["T_MonthKey", "T_MonthLabel", "T_WeekOfMonth", "T_WeekLabel"]]
            .dropna()
            .drop_duplicates()
            .sort_values(["T_MonthKey", "T_WeekOfMonth"])
        )
        w_options = w_map["T_WeekLabel"].tolist()

        w_sel = st.multiselect(
            "Selecciona Semana(s) del mes (Tendencia Ejecutivo)",
            options=w_options,
            default=w_options,
            key="tend_mvw_weeks_multi",
        )

        if w_sel:
            df_f = df_f[df_f["T_WeekLabel"].isin(w_sel)].copy()

        df_ctx = df_f.copy()

    st.markdown("---")

    # ✅ Exclusion ONLY inside this TAB (options + visuals)
    df_ctx_opts = df_ctx[~df_ctx["Supervisor_norm"].isin(EXCLUDED_SUP_NORMS)].copy()

    # ======================================================
    # ✅ Teams plot: Ventas por Supervisor + etiqueta #Ejecutivos
    # ======================================================
    st.markdown("### Equipos (Supervisores) — Ejecutivos por equipo y Ventas")

    df_team = df_ctx_opts.copy()

    if df_team.empty:
        st.info("No hay datos para visualizar equipos con los filtros actuales.")
    else:
        team_kpi = (
            df_team.groupby("Supervisor", as_index=False)
            .agg(
                Ventas=("FOLIO", "count"),
                Ejecutivos=("EJECUTIVO", "nunique"),
            )
            .sort_values("Ventas", ascending=False)
            .reset_index(drop=True)
        )

        team_kpi["Etiqueta"] = team_kpi.apply(
            lambda r: f"{int(r['Ventas']):,} ventas  |  {int(r['Ejecutivos'])} ejecutivos",
            axis=1,
        )

        dyn_h = max(380, 160 + 52 * len(team_kpi))

        fig_team = px.bar(
            team_kpi.sort_values("Ventas", ascending=True),
            x="Ventas",
            y="Supervisor",
            orientation="h",
            text="Etiqueta",
            title="Ventas por Equipo (Supervisor) y número de Ejecutivos",
            labels={"Supervisor": "Supervisor (Equipo)", "Ventas": "Ventas"},
            template=PLOTLY_TEMPLATE,
        )
        fig_team.update_traces(textposition="outside", cliponaxis=False)
        fig_team.update_layout(
            height=dyn_h,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=320, r=40, t=70, b=30),
            showlegend=False,
            xaxis=dict(zeroline=False),
            yaxis=dict(automargin=True),
        )

        st.plotly_chart(fig_team, width="stretch", key="t8_team_supervisores_overview")

        team_show = team_kpi[["Supervisor", "Ventas", "Ejecutivos"]].copy()
        team_show = add_totals_row(
            team_show,
            label_col="Supervisor",
            totals={"Ventas": int(team_show["Ventas"].sum()), "Ejecutivos": int(team_show["Ejecutivos"].sum())},
            label="TOTAL",
        )
        st.dataframe(
            style_totals_bold(team_show, label_col="Supervisor").format({"Ventas": "{:,.0f}", "Ejecutivos": "{:,.0f}"}),
            hide_index=True,
            width="stretch",
        )

    st.markdown("---")

    # ======================================================
    # 2) Supervisor filter (TAB)
    # ======================================================
    sup_opts_tab = sorted([s for s in df_ctx_opts["Supervisor"].dropna().unique().tolist()])
    sup_options = ["Todos"] + sup_opts_tab

    if "tend_ej_sup_tab" in st.session_state and st.session_state["tend_ej_sup_tab"] not in sup_options:
        st.session_state["tend_ej_sup_tab"] = "Todos"

    sup_tab_choice = st.selectbox(
        "Supervisor (Tendencia Ejecutivo)",
        options=sup_options,
        index=0,
        key="tend_ej_sup_tab",
    )

    df_ctx2 = df_ctx_opts.copy()
    if sup_tab_choice != "Todos":
        df_ctx2 = df_ctx2[df_ctx2["Supervisor"] == sup_tab_choice].copy()

    # ======================================================
    # 3) Ejecutivo filter (TAB)
    # ======================================================
    ej_opts = sorted([e for e in df_ctx2["EJECUTIVO"].dropna().unique().tolist()])
    if not ej_opts:
        st.info("No hay ejecutivos disponibles para el supervisor seleccionado con los filtros actuales.")
    else:
        if "tend_ej_sel" in st.session_state and st.session_state["tend_ej_sel"] not in ej_opts:
            st.session_state["tend_ej_sel"] = ej_opts[0]

        ej = st.selectbox("Ejecutivo", options=ej_opts, key="tend_ej_sel")
        st.markdown(f"✅ Has seleccionado: **{ej}**")

        # 4) Data for charts
        df_e = df_ctx2[df_ctx2["EJECUTIVO"] == ej].copy()

        if df_e.empty:
            st.info("Sin datos para el ejecutivo seleccionado con los filtros actuales.")
        else:
            m = (
                df_e.groupby("AñoMes", as_index=False)
                .size()
                .rename(columns={"size": "Ventas"})
                .sort_values("AñoMes")
                .reset_index(drop=True)
            )

            m["MesDT"] = pd.to_datetime(m["AñoMes"].astype(int).astype(str) + "01", format="%Y%m%d", errors="coerce")

            prom = float(m.iloc[:-1]["Ventas"].mean()) if len(m) > 1 else np.nan

            cur_ym = int(m["AñoMes"].iloc[-1])
            cur_v = int(m["Ventas"].iloc[-1])
            cur_label = mes_labels.get(cur_ym, month_key_to_name_es(cur_ym))

            c1, c2 = st.columns([0.65, 0.35], gap="large")

            with c1:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=m["MesDT"], y=m["Ventas"], mode="lines+markers", name="Ventas"))
                if not np.isnan(prom):
                    fig.add_trace(go.Scatter(x=m["MesDT"], y=[prom] * len(m), mode="lines", name="Promedio (sin mes actual)"))
                fig.update_layout(
                    title="Tendencia X Ejecutivo",
                    xaxis_title="Mes",
                    yaxis_title="Ventas",
                    height=420,
                )
                fig.update_xaxes(tickformat="%b %Y")
                apply_plotly_theme(fig)
                st.plotly_chart(fig, width="stretch", key=f"t8_tend_line_{normalize_name(ej)}")

            with c2:
                metric_card("Ventas Mes", fmt_int(cur_v), sub=cur_label)
                if not np.isnan(prom):
                    dif = cur_v - prom
                    metric_card("Diferencia", f"{dif:+.0f} vs promedio")

                figb = px.bar(
                    pd.DataFrame({"Mes": [cur_label], "Ventas": [cur_v]}),
                    x="Mes",
                    y="Ventas",
                    title="Ventas X Ejecutivo",
                    template=PLOTLY_TEMPLATE,
                )
                figb.update_layout(
                    height=300,
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=20, r=20, t=60, b=40),
                )
                st.plotly_chart(figb, width="stretch", key=f"t8_tend_bar_{normalize_name(ej)}")

            # ======================================================
            # ✅ SANITY CHECK (INTERVAL) + DESGLOSE POR EJECUTIVO (EQUIPO)
            # ======================================================
            st.markdown("---")
            st.markdown("### ✅ Sanity check — Meta mensual por antigüedad (Healthy vs Risky)")

            DEFAULT_META = 12  # default

            # ---------- Robust column picking ----------
            def _norm_txt(x: str) -> str:
                try:
                    x = str(x)
                except Exception:
                    x = ""
                x = x.strip().lower()
                x = unicodedata.normalize("NFKD", x)
                x = "".join([c for c in x if not unicodedata.combining(c)])
                x = " ".join(x.split())
                return x

            def _pick_col(df: pd.DataFrame, candidates):
                cols = list(df.columns)
                norm_cols = {_norm_txt(c): c for c in cols}

                for cand in candidates:
                    k = _norm_txt(cand)
                    if k in norm_cols:
                        return norm_cols[k]

                for cand in candidates:
                    kc = _norm_txt(cand)
                    for nk, real in norm_cols.items():
                        if kc and kc in nk:
                            return real
                return None

            name_col = _pick_col(
                empleados,
                ["Nombre Completo", "NOMBRE COMPLETO", "Nombre", "NOMBRE", "Empleado", "EMPLEADO", "Ejecutivo", "EJECUTIVO"],
            )
            ing_col = _pick_col(
                empleados,
                ["Fecha Ingreso", "FechaIngreso", "FECHA_INGRESO", "Fecha Alta", "FechaAlta", "FECHA_ALTA"],
            )
            dias_col = _pick_col(
                empleados,
                ["Dias Activos", "Días Activos", "Dias Activo", "Días Activo", "DiasActivos", "DIAS_ACTIVOS", "Dias", "DIAS"],
            )

            # ---------- Build employee records + better name matching ----------
            emp_by_norm = {}
            emp_records = []

            _STOP = {"DE", "DEL", "LA", "LAS", "LOS", "Y"}

            def _tokens(nombre: str):
                n = normalize_name(str(nombre).strip())
                toks = [t for t in n.split() if t]
                toks = [t for t in toks if t not in _STOP]
                return set(toks)

            if name_col:
                tmp = empleados[[name_col] + ([ing_col] if ing_col else []) + ([dias_col] if dias_col else [])].copy()
                tmp[name_col] = tmp[name_col].astype(str).str.strip()
                tmp["Nombre_norm"] = tmp[name_col].apply(normalize_name)

                if ing_col:
                    tmp["IngresoDT"] = pd.to_datetime(tmp[ing_col], errors="coerce")
                else:
                    tmp["IngresoDT"] = pd.NaT

                if dias_col:
                    tmp["DiasActivos"] = pd.to_numeric(tmp[dias_col], errors="coerce")
                else:
                    tmp["DiasActivos"] = np.nan

                tmp["_has_dias"] = tmp["DiasActivos"].notna().astype(int)
                tmp = tmp.sort_values(["_has_dias", "IngresoDT"], ascending=[False, True])
                tmp = tmp.drop_duplicates("Nombre_norm", keep="first").drop(columns=["_has_dias"])

                for _, r in tmp.iterrows():
                    nn = r["Nombre_norm"]
                    rec = {
                        "norm": nn,
                        "tokens": _tokens(nn),
                        "ingreso": (pd.Timestamp(r["IngresoDT"]).normalize() if pd.notna(r["IngresoDT"]) else None),
                        "dias": (float(r["DiasActivos"]) if pd.notna(r["DiasActivos"]) else None),
                    }
                    emp_by_norm[nn] = rec
                    emp_records.append(rec)

            def resolve_emp_record(nombre_ej: str):
                n = normalize_name(str(nombre_ej).strip())
                if n in emp_by_norm:
                    return emp_by_norm[n]

                t = _tokens(nombre_ej)
                if not t or not emp_records:
                    return None

                best = None
                best_score = 0.0
                for rec in emp_records:
                    inter = len(t & rec["tokens"])
                    if inter == 0:
                        continue
                    score = inter / max(len(t), len(rec["tokens"]))
                    if score > best_score:
                        best_score = score
                        best = rec

                if best and (best_score >= 0.70 or len(t & best["tokens"]) >= 3):
                    return best
                return None

            # ---------- Calendar days from ingreso -> ref_end ----------
            def calendar_days_since(ingreso_dt: pd.Timestamp, ref_end: pd.Timestamp) -> int:
                if ingreso_dt is None or pd.isna(ingreso_dt):
                    return 0
                start = pd.Timestamp(ingreso_dt).normalize()
                end = pd.Timestamp(ref_end).normalize()
                d = int((end - start).days)
                return max(d, 0)

            def antiguedad_meses_y_dias(nombre_ej: str, ref_end: pd.Timestamp):
                rec = resolve_emp_record(nombre_ej)

                if rec and rec.get("ingreso") is not None:
                    cd = calendar_days_since(rec["ingreso"], ref_end)
                    return int(cd // 30), int(cd % 30), rec.get("ingreso")

                if rec and rec.get("dias") is not None and not pd.isna(rec["dias"]):
                    d_int = int(float(rec["dias"]))
                    return int(d_int // 30), int(d_int % 30), rec.get("ingreso")

                return None, None, None

            # ✅ meta evaluated VS the 1st day of each month
            # ✅ rule: if ingreso > 1st day of month => meta = 6
            def meta_por_antiguedad(nombre_ej: str, ref_end: pd.Timestamp) -> int:
                rec = resolve_emp_record(nombre_ej)

                if rec and rec.get("ingreso") is not None:
                    ing = pd.Timestamp(rec["ingreso"]).normalize()
                    ref0 = pd.Timestamp(ref_end).normalize()

                    if ing > ref0:
                        return 6

                    cd = calendar_days_since(ing, ref0)
                    return 12 if cd > 41 else 6

                if rec and rec.get("dias") is not None and not pd.isna(rec["dias"]):
                    return 12 if float(rec["dias"]) > 41 else 6

                return int(DEFAULT_META)

            # ======================================================
            # ✅ Interval scope (controlled by months+weeks filter)
            # ======================================================
            df_scope = df_ctx2.copy()  # already filtered by month(s)+week(s) + supervisor choice
            months_in_scope = sorted(df_scope["T_MonthKey"].dropna().unique().tolist())

            if df_scope.empty or not months_in_scope:
                st.info("No hay datos suficientes en el intervalo seleccionado (meses/semanas) para el sanity check.")
            else:
                df_e_scope = df_scope[df_scope["EJECUTIVO"] == ej].copy()
                ventas_intervalo = int(df_e_scope.shape[0])

                meta_rows = []
                meta_total = 0

                for mk in months_in_scope:
                    try:
                        yy, mm = mk.split("-")
                        ms = pd.Timestamp(year=int(yy), month=int(mm), day=1).normalize()
                    except Exception:
                        continue

                    ventas_m = int(df_e_scope[df_e_scope["T_MonthKey"] == mk].shape[0])
                    meta_m = int(meta_por_antiguedad(ej, ms))  # ✅ per-month meta at 1st day
                    meta_rows.append({"Mes": mk, "Ventas": ventas_m, "Meta": meta_m, "Delta": ventas_m - meta_m})
                    meta_total += meta_m

                meta_intervalo = int(meta_total)
                estado = "HEALTHY" if ventas_intervalo >= meta_intervalo else "RISKY"
                badge = "🟢 HEALTHY" if estado == "HEALTHY" else "🟠 RISKY"
                delta = int(ventas_intervalo - meta_intervalo)

                # Antigüedad shown using the LAST month in interval (1st day of that month)
                last_mk = months_in_scope[-1]
                try:
                    yy, mm = last_mk.split("-")
                    ref_tenure = pd.Timestamp(year=int(yy), month=int(mm), day=1).normalize()
                except Exception:
                    ref_tenure = pd.Timestamp.today().normalize()

                am, ad, ing_dt = antiguedad_meses_y_dias(ej, ref_tenure)
                antig_txt = f"{am} meses {ad} días" if (am is not None and ad is not None) else "No disponible (default meta=12)"

                csc1, csc2, csc3, csc4 = st.columns(4, gap="medium")
                with csc1:
                    st.markdown("**Intervalo evaluado**")
                    st.write(f"{months_in_scope[0]} → {months_in_scope[-1]}  |  Meses: {len(months_in_scope)}")
                with csc2:
                    metric_card(
                        "Antigüedad (al 1ro del último mes)",
                        antig_txt,
                        sub=(f"Ingreso: {ing_dt:%d/%m/%Y}" if ing_dt is not None else None),
                    )
                with csc3:
                    metric_card("Ventas (intervalo)", fmt_int(ventas_intervalo))
                with csc4:
                    metric_card("Meta / Estado", f"{meta_intervalo} — {badge}", sub=f"Delta: {delta:+d}")

                # Optional breakdown per month
                if len(meta_rows) > 1:
                    st.markdown("#### Desglose por mes (según filtro)")
                    df_break = pd.DataFrame(meta_rows)
                    st.dataframe(
                        df_break.style.format({"Ventas": "{:,.0f}", "Meta": "{:,.0f}", "Delta": "{:+,.0f}"}),
                        hide_index=True,
                        width="stretch",
                    )

                # Chart sanity (Ventas vs Meta) for the interval
                bar_color = "#2ecc71" if estado == "HEALTHY" else "#e74c3c"
                meta_color = bar_color

                fig_sc = go.Figure()
                fig_sc.add_trace(
                    go.Bar(
                        x=["Ventas (intervalo)"],
                        y=[ventas_intervalo],
                        name="Ventas",
                        text=[ventas_intervalo],
                        textposition="outside",
                        marker=dict(color=bar_color),
                    )
                )
                fig_sc.add_trace(
                    go.Scatter(
                        x=["Ventas (intervalo)"],
                        y=[meta_intervalo],
                        mode="markers+text",
                        name="Meta",
                        text=[f"Meta: {meta_intervalo}"],
                        textposition="top center",
                        marker=dict(size=12, color=meta_color),
                    )
                )

                fig_sc.update_layout(
                    title=f"Sanity check — {ej} — Intervalo (meses/semanas seleccionados)",
                    yaxis_title="Cantidad",
                    height=360,
                )

                apply_plotly_theme(fig_sc)
                fig_sc.update_traces(marker_color=bar_color, selector=dict(type="bar"))
                fig_sc.update_traces(marker=dict(color=meta_color), selector=dict(type="scatter"))
                st.plotly_chart(fig_sc, width="stretch", key=f"t8_sanity_interval_{normalize_name(ej)}")

                # ---------- Desglose por ejecutivo del equipo (interval) ----------
                st.markdown("#### ✅ Desglose por Ejecutivo (equipo) — Ventas vs Meta mensual (según intervalo)")

                g = (
                    df_scope.groupby("EJECUTIVO", as_index=False)
                    .size()
                    .rename(columns={"size": "Ventas"})
                )

                # MetaIntervalo per ejecutivo = sum(meta per month in interval
                def _meta_intervalo_for_exec(exec_name: str) -> int:
                    total = 0
                    for mk2 in months_in_scope:
                        try:
                            yy2, mm2 = mk2.split("-")
                            ms2 = pd.Timestamp(year=int(yy2), month=int(mm2), day=1).normalize()
                        except Exception:
                            continue
                        total += int(meta_por_antiguedad(exec_name, ms2))
                    return int(total)

                g["MetaIntervalo"] = g["EJECUTIVO"].apply(_meta_intervalo_for_exec)

                def _antig_str(nm: str) -> str:
                    mm, dd, _ = antiguedad_meses_y_dias(nm, ref_tenure)
                    return f"{mm}m {dd}d" if (mm is not None and dd is not None) else "N/D"

                g["Antigüedad"] = g["EJECUTIVO"].apply(_antig_str)
                g["Delta"] = g["Ventas"] - g["MetaIntervalo"]
                g["Estado"] = np.where(g["Ventas"] >= np.ceil(g["MetaIntervalo"] / 2.0), "HEALTHY", "RISKY")


                def _rank_estado(x: str) -> int:
                    return 0 if x == "RISKY" else 1

                g["_rank"] = g["Estado"].apply(_rank_estado)
                g = g.sort_values(["_rank", "Delta", "Ventas"], ascending=[True, True, False]).drop(columns=["_rank"])

                # ✅ COLOR MAP: GREEN healthy, RED risky
                STATUS_COLORS = {"HEALTHY": "#2ecc71", "RISKY": "#e74c3c"}

                fig_team_exec = px.bar(
                    g.sort_values("Ventas", ascending=True),
                    x="Ventas",
                    y="EJECUTIVO",
                    orientation="h",
                    color="Estado",  # ✅ color by state
                    color_discrete_map=STATUS_COLORS,  # ✅ force green/red
                    title="Ventas por Ejecutivo (intervalo seleccionado)",
                    hover_data={"MetaIntervalo": True, "Antigüedad": True, "Delta": True, "Estado": True},
                    template=PLOTLY_TEMPLATE,
                )
                fig_team_exec.update_layout(
                    height=min(900, 140 + 30 * len(g)),
                    margin=dict(l=20, r=20, t=70, b=20),
                    legend_title_text="Estado",
                )
                apply_plotly_theme(fig_team_exec)
                st.plotly_chart(fig_team_exec, width="stretch", key=f"t8_team_exec_interval_{normalize_name(sup_tab_choice)}")

                show = g.rename(columns={"MetaIntervalo": "Meta intervalo"}).copy()
                st.dataframe(
                    show.style.format(
                        {
                            "Ventas": "{:,.0f}",
                            "Meta intervalo": "{:,.0f}",
                            "Delta": "{:+,.0f}",
                        }
                    ),
                    hide_index=True,
                    width="stretch",
                )
