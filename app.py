"""
Calendario fenológico del maíz por grados-día (GDD) — Streamlit.

Entradas
  * Datos meteorológicos diarios (o subdiarios) con Tmax y Tmin:
    estaciones incluidas en data/estaciones (planillas SENAMHI .XLS) o un archivo subido
    (CSV tipo ejemplo_busan.csv, Excel, CSV horario/10 min).
  * Variedad de maíz desde data/gdd_variedades.xlsx (hoja «requerimientos»): fases, GDD, Tb y método.
    Fases, GDD, Tb y método siguen siendo editables.

Ejecutar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import datetime as dt
import io
import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Calendario GDD del maíz", layout="wide")

MAX_DIAS = 400          # horizonte de simulación desde la siembra
DATA = Path(__file__).parent / "data"
DIR_ESTACIONES = DATA / "estaciones"
EJEMPLO = DATA / "ejemplo_busan.csv"
MAESTRA = DIR_ESTACIONES / "Maestra.xlsx"
VARIEDADES_XLSX = DATA / "gdd_variedades.xlsx"
ESTACION_DEFECTO = "urubamba"

COLOR_CULTIVO = "#f0a500"

GENERICA = "Genérica (valores de literatura)"
ETAPAS_GENERICAS = pd.DataFrame({
    "Etapa": ["Emergencia de la panoja (VT)", "Aparición de estigmas (R1)",
              "Grano lechoso (R3)", "Momento óptimo de cosecha"],
    "GDD": [490.0, 600.0, 700.0, 880.0],
    "Descripción": [
        "Aparecen las flores masculinas",
        "Polinización; etapa de mayor necesidad de agua",
        "Los granos se llenan de jugo blanco",
        "En maíz dulce, 20–25 días después de R1",
    ],
})
NOMBRES_FASE = {
    "emergencia": "Emergencia (VE)",
    "aparicion de hojas": "Aparición de hojas",
    "panoja": "Panoja (VT)",
    "espiga": "Espiga / estigmas (R1)",
    "maduracion lechosa": "Maduración lechosa (R3)",
    "maduracion pastosa": "Maduración pastosa (R4)",
    "maduracion cornea": "Maduración córnea (R6)",
    "cosecha": "Cosecha",
}
FASES_OCULTAS = {"emergencia", "aparicion de hojas"}
METODOS = {
    "media": "Media — max(min(Tmed, Tu) − Tb, 0)",
    "acotada": "Acotada — Tmax y Tmin limitadas a [Tb, Tu]",
}


def sin_tildes(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in s if not unicodedata.combining(c)).lower()


# --------------------------------------------------------------------------
# Lectura de datos meteorológicos
# --------------------------------------------------------------------------
def leer_html_senamhi(texto: str) -> pd.DataFrame:
    """Planilla SENAMHI exportada como tabla HTML con extensión .XLS.
    Columnas: Codigo, Planilla, Año, Mes, V Dia, PT101, TM102 (Tmax), TM103 (Tmin)."""
    filas = re.split(r"<TR[^>]*>", texto, flags=re.I)[1:]
    i = next(k for k, f in enumerate(filas) if re.search(r"<TH", f, re.I))
    cab = [h.strip() for h in re.findall(r"<TH[^>]*>([^<]*)", filas[i], re.I)]
    datos = [[c.strip() for c in re.findall(r"<TD[^>]*>([^<]*)", f, re.I)] for f in filas[i + 1:]]
    return pd.DataFrame([d for d in datos if len(d) == len(cab)], columns=cab)


def unir_fecha(df: pd.DataFrame) -> pd.DataFrame:
    """Si la fecha viene en columnas separadas (año, mes, día) crea una columna «fecha»."""
    norm = {sin_tildes(c).strip(): c for c in df.columns}
    def col(*claves):
        for k, c in norm.items():
            if any(k == x or k.endswith(" " + x) for x in claves):
                return c
        return None
    a, m, d = col("ano", "year", "anio"), col("mes", "month"), col("dia", "day")
    if a and m and d and not any(x in norm for x in ("fecha", "date")):
        f = pd.to_datetime(dict(year=pd.to_numeric(df[a], errors="coerce"),
                                month=pd.to_numeric(df[m], errors="coerce"),
                                day=pd.to_numeric(df[d], errors="coerce")), errors="coerce")
        df = df.drop(columns=[a, m, d])
        df.insert(0, "fecha", f.dt.strftime("%Y-%m-%d"))
    return df


@st.cache_data(show_spinner=False)
def leer_tabla(contenido: bytes, nombre: str) -> pd.DataFrame:
    """Lee planilla SENAMHI (HTML), CSV/TXT o Excel, detectando separador y coma decimal."""
    cabeza = contenido[:2000].lstrip().lower()
    if cabeza.startswith(b"<") and (b"<table" in cabeza or b"<title" in cabeza or b"<html" in cabeza):
        try:
            texto = contenido.decode("utf-8")
        except UnicodeDecodeError:
            texto = contenido.decode("latin-1")
        return unir_fecha(leer_html_senamhi(texto))
    if nombre.lower().endswith((".xlsx", ".xls")):
        return unir_fecha(pd.read_excel(io.BytesIO(contenido)))
    try:
        contenido.decode("utf-8")
        cod = "utf-8-sig"
    except UnicodeDecodeError:
        cod = "latin-1"  # CSV guardado con Excel en Windows
    df = pd.read_csv(io.BytesIO(contenido), sep=None, engine="python", encoding=cod)
    obj = [c for c in df.columns[1:] if not pd.api.types.is_numeric_dtype(df[c])]
    if obj and df[obj[0]].astype(str).str.match(r"^-?\d+,\d+$").mean() > 0.5:
        df = pd.read_csv(io.BytesIO(contenido), sep=None, engine="python", decimal=",", encoding=cod)
    return unir_fecha(df)


def detectar_columnas(df: pd.DataFrame) -> dict:
    """Propone columnas de fecha, Tmax y Tmin según el nombre."""
    cols = list(df.columns)
    n = {c: sin_tildes(c) for c in cols}
    excl = ("rocio", "suelo", "intens", "bateria")
    ok = {c: not any(x in n[c] for x in excl) and not re.search(r"temperatura -\d", n[c]) for c in cols}

    def primero(cond):
        return next((c for c in cols if ok[c] and cond(n[c])), None)

    fecha = primero(lambda s: any(x in s for x in ("fecha", "date", "time"))) or cols[0]
    tmax = (primero(lambda s: "tm102" in s or re.search(r"\bt_?max\b|^tx$", s))
            or primero(lambda s: "temp" in s and "max" in s))
    tmin = (primero(lambda s: "tm103" in s or re.search(r"\bt_?min\b|^tn$", s))
            or primero(lambda s: "temp" in s and "min" in s))
    temp = (primero(lambda s: "temp" in s and ("media" in s or "instant" in s))
            or primero(lambda s: "temp" in s and "max" not in s and "min" not in s))
    return {"fecha": fecha, "tmax": tmax or temp, "tmin": tmin or temp}


def a_fecha(col: pd.Series) -> pd.Series:
    """Fechas ISO (AAAA-MM-DD) o día/mes/año (DD/MM/AAAA), nunca mes/día/año."""
    if pd.api.types.is_datetime64_any_dtype(col):
        return col
    texto = col.astype(str).str.strip()
    iso = texto.str.match(r"^\d{4}[-/.]").mean() > 0.5
    return pd.to_datetime(texto, errors="coerce", dayfirst=not iso)


@st.cache_data(show_spinner=False)
def a_diario(df: pd.DataFrame, c_fecha: str, c_tmax: str, c_tmin: str):
    """Serie diaria continua (tmax, tmin) con control de calidad. Agrega datos subdiarios."""
    d = pd.DataFrame({
        "fecha": a_fecha(df[c_fecha]),
        "tmax": pd.to_numeric(df[c_tmax], errors="coerce"),
        "tmin": pd.to_numeric(df[c_tmin], errors="coerce"),
    })
    d = d.dropna(subset=["fecha"])
    # Códigos de dato faltante (-999, -99.9…)
    for c in ("tmax", "tmin"):
        d.loc[(d[c] < -60) | (d[c] > 60), c] = np.nan
    d["fecha"] = d["fecha"].dt.normalize()
    agg = {"tmax": "max", "tmin": "min"}
    n = d.groupby("fecha").size()
    d = d.groupby("fecha").agg(agg)
    # Datos subdiarios: un día con pocos registros da Tmax/Tmin sesgadas → faltante
    if n.median() > 1:
        d.loc[n < 0.75 * n.median()] = np.nan
    invertidos = d["tmin"] > d["tmax"]
    d.loc[invertidos, ["tmax", "tmin"]] = np.nan
    idx = pd.date_range(d.index.min(), d.index.max(), freq="D")
    d = d.reindex(idx)
    d.index.name = "fecha"
    return d, int(invertidos.sum())


def nombre_estacion(p: Path) -> str:
    return p.stem.replace("_", " ").title()


@st.cache_data(show_spinner=False)
def codigo_estacion(ruta: str) -> int | None:
    """Código SENAMHI de la primera fila de datos de la planilla."""
    with open(ruta, "rb") as f:
        texto = f.read(20000).decode("latin-1")
    m = re.search(r"<TD[^>]*>\s*(\d+)\s*</TD>", texto, flags=re.I)
    return int(m.group(1)) if m else None


@st.cache_data(show_spinner=False)
def leer_maestra(ruta: str, modificado: float = 0) -> dict[int, str]:
    """Código (CODIGO o NCODIGO) → 'Distrito, Provincia, Departamento'."""
    m = pd.read_excel(ruta)
    ubic = {}
    for _, r in m.iterrows():
        texto = ", ".join(str(r[c]).strip().title() for c in ("DIST", "PROV", "DPTO"))
        for c in ("NCODIGO", "CODIGO"):
            if pd.notna(r[c]):
                ubic[int(r[c])] = texto
    return ubic


# --------------------------------------------------------------------------
# Variedades
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def leer_variedades(contenido: bytes) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(contenido), sheet_name="requerimientos")


def fases_de(req: pd.DataFrame) -> list[str]:
    """Fases presentes en la hoja: columnas gdd_<fase> (sin sufijos _min/_max)."""
    return [m.group(1) for c in req.columns
            if (m := re.fullmatch(r"gdd_(.+)", c)) and not re.search(r"_(min|max)$", c)]


def etapas_variedad(fila: pd.Series, fases: list[str]) -> pd.DataFrame:
    filas = []
    for f in fases:
        g = fila.get(f"gdd_{f}")
        if pd.isna(g) or sin_tildes(f) in FASES_OCULTAS:
            continue
        partes = []
        if pd.notna(fila.get(f"dias_{f}")):
            partes.append(f"~{fila[f'dias_{f}']:.0f} días desde la siembra")
        if pd.notna(fila.get(f"n_{f}")):
            partes.append(f"n={fila[f'n_{f}']:.0f} campañas")
        if pd.notna(fila.get(f"cv_{f}_pct")):
            partes.append(f"CV {fila[f'cv_{f}_pct']:.1f} %")
        filas.append({
            "Etapa": NOMBRES_FASE.get(sin_tildes(f), f.capitalize()),
            "GDD": round(float(g), 1),
            "Descripción": " · ".join(partes),
        })
    return pd.DataFrame(filas, columns=ETAPAS_GENERICAS.columns)


# --------------------------------------------------------------------------
# Cálculo de GDD
# --------------------------------------------------------------------------
def gdd_diario(tmax: pd.Series, tmin: pd.Series, tb: float, tu: float, metodo: str) -> pd.Series:
    if metodo == "acotada":   # McMaster y Wilhelm (1997), método 2
        return (tmax.clip(tb, tu) + tmin.clip(tb, tu)) / 2 - tb
    tm = (tmax + tmin) / 2    # método 1 (promedio simple)
    return (tm.clip(upper=tu) - tb).clip(lower=0)


def doy_sin_bisiesto(idx: pd.DatetimeIndex) -> np.ndarray:
    """Día del año 0..364; el 29 de febrero se asimila al 28."""
    md = idx.month * 100 + idx.day
    md = np.where(md == 229, 228, md)
    ref = pd.to_datetime([f"2001-{m // 100:02d}-{m % 100:02d}" for m in md])
    return (ref.dayofyear - 1).to_numpy()


def climatologia(serie: pd.Series) -> np.ndarray:
    """Media por día del año (365 valores); huecos rellenados por interpolación circular."""
    df = pd.DataFrame({"v": serie.to_numpy(), "doy": doy_sin_bisiesto(serie.index)})
    m = df.dropna().groupby("doy")["v"].mean().reindex(range(365))
    if m.isna().all():
        return np.zeros(365)
    tri = pd.concat([m, m, m], ignore_index=True).interpolate(limit_direction="both")
    return tri.iloc[365:730].to_numpy()


@st.cache_data(show_spinner=False)
def preparar(d: pd.DataFrame, tb: float, tu: float, metodo: str):
    g = gdd_diario(d["tmax"], d["tmin"], tb, tu, metodo)
    tmed = (d["tmax"] + d["tmin"]) / 2
    return g, climatologia(g), climatologia(tmed)


def inicio_siembra(y: int, siembra: dt.date) -> pd.Timestamp:
    try:
        return pd.Timestamp(year=y, month=siembra.month, day=siembra.day)
    except ValueError:          # 29-feb en año no bisiesto
        return pd.Timestamp(year=y, month=2, day=28)


def temporada(d, g, clim_g, clim_t, siembra: dt.date, hoy: dt.date) -> pd.DataFrame:
    """Serie desde el día siguiente a la siembra: observado hasta la fecha de referencia,
    climatología después."""
    fechas = pd.date_range(pd.Timestamp(siembra) + pd.Timedelta(days=1), periods=MAX_DIAS, freq="D")
    doy = doy_sin_bisiesto(fechas)
    pasado = fechas <= pd.Timestamp(hoy)
    g_obs = g.reindex(fechas).to_numpy()
    t_obs = ((d["tmax"] + d["tmin"]) / 2).reindex(fechas).to_numpy()
    usa_obs = pasado & ~np.isnan(g_obs)
    gv = np.where(usa_obs, g_obs, clim_g[doy])
    return pd.DataFrame({
        "das": np.arange(1, MAX_DIAS + 1),
        "gdd": gv,
        "gdd_acum": np.cumsum(gv),
        "obs": usa_obs,
        "faltante": pasado & ~usa_obs,
        "tmed": np.where(usa_obs, t_obs, clim_t[doy]),
    }, index=fechas)


def alcanza(sea: pd.DataFrame, objetivo) -> pd.Timestamp | None:
    if objetivo is None or pd.isna(objetivo):
        return None
    hit = sea.index[sea["gdd_acum"] >= objetivo]
    return hit[0] if len(hit) else None


def fmt(f) -> str:
    return f.strftime("%d/%m/%Y") if f is not None and not pd.isna(f) else "—"


# --------------------------------------------------------------------------
# Barra lateral · 1 · Datos meteorológicos
# --------------------------------------------------------------------------
st.sidebar.header("1 · Datos meteorológicos")
incluidas = sorted([p for p in DIR_ESTACIONES.glob("*")
                    if p.suffix.lower() in (".xls", ".xlsx", ".csv", ".txt") and p != MAESTRA
                    and not p.name.startswith((".", "~"))],
                   key=lambda p: sin_tildes(p.stem)) if DIR_ESTACIONES.exists() else []
if EJEMPLO.exists():
    incluidas.append(EJEMPLO)
origen = st.sidebar.radio("Origen", ["Estación incluida", "Subir archivo"], horizontal=True,
                          disabled=not incluidas, index=0 if incluidas else 1)

if origen == "Subir archivo" or not incluidas:
    archivo = st.sidebar.file_uploader(
        "Archivo con Tmax y Tmin",
        type=["csv", "txt", "xlsx", "xls"],
        help="Formatos: CSV como ejemplo_busan.csv (date,tmax,tmin); planilla SENAMHI .XLS "
             "(TM102, TM103); CSV horario o de 10 min (se agrega a diario).",
    )
    st.sidebar.download_button(
        "Descargar plantilla CSV",
        "date,tmax,tmin\n2025-01-01,28.4,17.2\n2025-01-02,29.1,16.8\n",
        file_name="plantilla_gdd.csv", mime="text/csv",
    )
    if archivo is None:
        st.caption("Suba un archivo con datos meteorológicos o elija una estación incluida.")
        st.stop()
    crudo = leer_tabla(archivo.getvalue(), archivo.name)
    fuente = archivo.name
    ubicacion = None
else:
    nombres = [nombre_estacion(p) for p in incluidas]
    i_def = next((i for i, p in enumerate(incluidas) if sin_tildes(p.stem) == ESTACION_DEFECTO), 0)
    maestra = leer_maestra(str(MAESTRA), MAESTRA.stat().st_mtime) if MAESTRA.exists() else {}
    ubicaciones = [maestra.get(codigo_estacion(str(p))) if p.suffix.lower() == ".xls" else None
                   for p in incluidas]
    i_sel = st.sidebar.selectbox(
        f"Estación ({len(incluidas)} disponibles)", range(len(incluidas)), index=i_def,
        format_func=lambda i: f"{nombres[i]} · {ubicaciones[i]}" if ubicaciones[i] else nombres[i],
        help="Distrito, provincia y departamento según Maestra.xlsx. Escriba para buscar.",
    )
    ruta = incluidas[i_sel]
    crudo = leer_tabla(ruta.read_bytes(), ruta.name)
    fuente = nombres[i_sel]
    ubicacion = ubicaciones[i_sel]

cols = list(crudo.columns)
auto = detectar_columnas(crudo)
if origen == "Subir archivo":
    with st.sidebar.expander("Columnas del archivo", expanded=True):
        def _sel(etq, clave, opciones):
            v = auto[clave]
            return st.selectbox(etq, opciones, index=opciones.index(v) if v in opciones else 0,
                                key=f"col_{clave}_{fuente}")
        c_fecha = _sel("Fecha", "fecha", cols)
        c_tmax = _sel("Temperatura máxima (°C)", "tmax", cols)
        c_tmin = _sel("Temperatura mínima (°C)", "tmin", cols)
        st.caption("Con datos horarios de una sola temperatura, elija esa columna como máxima y "
                   "como mínima: se toma el máximo y el mínimo de cada día. Los valores -999 / -99.9 "
                   "se tratan como faltantes.")
else:
    # Estaciones incluidas: columnas detectadas automáticamente
    c_fecha, c_tmax, c_tmin = auto["fecha"], auto["tmax"], auto["tmin"]
    if c_tmax is None or c_tmin is None:
        st.error(f"No se encontraron columnas de Tmax y Tmin en el archivo de la estación {fuente}.")
        st.stop()

try:
    datos, n_invertidos = a_diario(crudo, c_fecha, c_tmax, c_tmin)
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo interpretar el archivo con las columnas elegidas: {e}")
    st.stop()

validos = datos.dropna(subset=["tmax", "tmin"])
if validos.empty:
    st.error("El archivo no contiene días con Tmax y Tmin válidas. Revise la selección de columnas.")
    st.stop()

F0, FN = datos.index.min().date(), validos.index.max().date()
anios = sorted(validos.index.year.unique().tolist())

# --------------------------------------------------------------------------
# Barra lateral · 2 · Variedad y parámetros
# --------------------------------------------------------------------------
st.sidebar.header("2 · Variedad y parámetros")
req = pd.DataFrame()
try:
    if VARIEDADES_XLSX.exists():
        req = leer_variedades(VARIEDADES_XLSX.read_bytes())
except Exception as e:  # noqa: BLE001
    st.sidebar.error(f"No se pudo leer la hoja «requerimientos»: {e}")

fases = fases_de(req) if not req.empty else []
con_datos = req[req[[f"gdd_{f}" for f in fases]].notna().any(axis=1)] if fases else req
sin_datos = sorted(set(req.get("variedad", pd.Series(dtype=str))) - set(con_datos.get("variedad", [])))
opciones_var = con_datos["variedad"].tolist() + [GENERICA] if not con_datos.empty else [GENERICA]
variedad = st.sidebar.selectbox(
    "Variedad", opciones_var, format_func=lambda v: v if v == GENERICA else v.title(),
    help="Fases, GDD, Tb y método calibrados por variedad (hoja «requerimientos»).",
)
if sin_datos:
    st.sidebar.caption("Sin requerimientos calculados: " + ", ".join(v.title() for v in sin_datos))

fila_var = None if variedad == GENERICA else con_datos[con_datos["variedad"] == variedad].iloc[0]
tb_var = 10.0 if fila_var is None or pd.isna(fila_var.get("tb")) else float(fila_var["tb"])
metodo_var = "media" if fila_var is None or fila_var.get("metodo") not in METODOS else fila_var["metodo"]

# Campaña grande de la sierra: 1 de octubre más reciente con datos
siembra_def = next((d for d in (dt.date(FN.year, 10, 1), dt.date(FN.year - 1, 10, 1)) if F0 <= d <= FN), F0)
siembra = st.sidebar.date_input("Fecha de siembra", value=siembra_def, min_value=F0,
                                format="DD/MM/YYYY", key=f"siembra_{fuente}")
hoy = st.sidebar.date_input("Fecha de referencia («hoy»)", value=FN, min_value=F0,
                            format="DD/MM/YYYY", key=f"hoy_{fuente}")
c1, c2 = st.sidebar.columns(2)
tb = c1.number_input("Tb (°C)", value=tb_var, step=0.5, key=f"tb_{variedad}")
tu = c2.number_input("Tu (°C)", value=30.0, step=0.5, key=f"tu_{variedad}")
metodo = st.sidebar.selectbox("Método de cálculo", list(METODOS), format_func=METODOS.get,
                              index=list(METODOS).index(metodo_var), key=f"metodo_{variedad}")
if tu <= tb:
    st.sidebar.error("Tu debe ser mayor que Tb.")
    st.stop()

# --------------------------------------------------------------------------
# Barra lateral · 3 · Fases y GDD editables
# --------------------------------------------------------------------------
st.sidebar.header("3 · Fases y GDD")
clave_base, clave_ed = f"base_{variedad}", f"editor_{variedad}"
if clave_base not in st.session_state:
    st.session_state[clave_base] = (ETAPAS_GENERICAS.copy() if fila_var is None
                                    else etapas_variedad(fila_var, fases))
editadas = st.sidebar.data_editor(
    st.session_state[clave_base],
    num_rows="dynamic",
    hide_index=True,
    width="stretch",
    key=clave_ed,
    column_config={
        "Etapa": st.column_config.TextColumn(required=True),
        "GDD": st.column_config.NumberColumn("GDD", min_value=0, step=10, format="%.0f", required=True),
        "Descripción": st.column_config.TextColumn(width="medium"),
    },
)
st.sidebar.caption("Edite fases y GDD directamente; agregue o elimine filas con ＋ / 🗑. "
                   "Los cambios se guardan por variedad durante la sesión.")
if st.sidebar.button("Restaurar", width="stretch", help="Volver a los valores de la variedad"):
    st.session_state.pop(clave_base, None)
    st.session_state.pop(clave_ed, None)
    for k in (f"tb_{variedad}", f"tu_{variedad}", f"metodo_{variedad}"):
        st.session_state.pop(k, None)
    st.rerun()

etapas = (editadas.dropna(subset=["Etapa", "GDD"])
          .assign(Descripción=lambda x: x["Descripción"].fillna(""))
          .sort_values("GDD").reset_index(drop=True))
if etapas.empty:
    st.caption("Defina al menos una fase con su GDD en la barra lateral.")
    st.stop()
etapas["G"] = etapas["GDD"].astype(float).round().astype(int)

# --------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------
g, clim_g, clim_t = preparar(datos, tb, tu, metodo)
sea = temporada(datos, g, clim_g, clim_t, siembra, hoy)
pasado = sea.loc[: pd.Timestamp(hoy)]
acum_hoy = float(pasado["gdd"].sum())
k_hoy = (pd.Timestamp(hoy) - pd.Timestamp(siembra)).days

etapas["Fecha"] = [alcanza(sea, v) for v in etapas["G"]]
etapas["DAS"] = [(f - pd.Timestamp(siembra)).days if f is not None else None for f in etapas["Fecha"]]

actual = etapas[etapas["G"] <= acum_hoy].tail(1)
siguiente = etapas[etapas["G"] > acum_hoy].head(1)

# --------------------------------------------------------------------------
# Encabezado y avisos
# --------------------------------------------------------------------------
st.title("Calendario fenológico del maíz")
st.caption(f"Grados-día de crecimiento (GDD) · Estación: **{fuente}**"
           f"{f' ({ubicacion})' if ubicacion else ''} · {fmt(F0)} – {fmt(FN)} "
           f"({len(datos)} días, {datos['tmax'].notna().mean():.0%} con temperatura)")

if fila_var is not None:
    conf = str(fila_var.get("confiabilidad", "")).strip() or "—"
    obs_var = f" · {fila_var['observaciones']}" if pd.notna(fila_var.get("observaciones")) else ""
    st.caption(f"Variedad **{variedad.title()}** · Tb calibrada {tb_var:g} °C "
               f"(intervalo {fila_var.get('tb_intervalo', '—')} °C, método {metodo_var}) · "
               f"Campañas de calibración {fila_var.get('campanias', 0):.0f} · "
               f"Confiabilidad {conf}{obs_var}")

avisos = []
if fila_var is None:
    avisos.append(f"Variedad genérica: los GDD por fase y la Tb ({tb:g} °C) son **valores "
                  "provisionales** de literatura.")
else:
    if tb != tb_var or metodo != metodo_var:
        avisos.append(f"Tb o método modificados respecto a la calibración (Tb {tb_var:g} °C, "
                      f"{metodo_var}): los GDD de las fases dejan de ser comparables.")
    ultima = etapas["Fecha"].dropna().max() if etapas["Fecha"].notna().any() else None
    if ultima is not None:
        tmed_ciclo = sea.loc[:ultima, "tmed"].mean()
        lo, hi = fila_var.get("tmed_min"), fila_var.get("tmed_max")
        if pd.notna(lo) and pd.notna(hi) and not (lo - 1 <= tmed_ciclo <= hi + 1):
            avisos.append(f"La temperatura media del ciclo en esta estación ({tmed_ciclo:.1f} °C) está "
                          f"**fuera del rango de calibración** de la variedad ({lo:.1f}–{hi:.1f} °C): "
                          "las fechas pueden ser poco fiables.")
if hoy > FN:
    avisos.append(f"La fecha de referencia es posterior al último dato ({fmt(FN)}): "
                  "ese tramo se estima con **climatología**.")
n_falt = int(sea["faltante"].sum())
if n_invertidos:
    avisos.append(f"{n_invertidos} día(s) con Tmin > Tmax se descartaron como dato erróneo.")
if len(validos) < 365:
    avisos.append(f"La climatología se calcula con solo {len(validos)} día(s) con temperatura: "
                  "las fechas proyectadas no son confiables.")
elif len(anios) < 3:
    avisos.append(f"Solo hay {len(anios)} año(s) de datos: la climatología usada para proyectar "
                  "es poco robusta.")
for a in avisos:
    st.caption(a)
if n_falt:
    st.caption(f"{n_falt} día(s) sin temperatura entre la siembra y la fecha de referencia "
               "se sustituyeron por climatología.")

# --------------------------------------------------------------------------
# Estado actual
# --------------------------------------------------------------------------
st.subheader("Estado actual del cultivo")
st.caption(f"Siembra {fmt(siembra)} · Referencia {fmt(hoy)} · acumulación desde el día siguiente a la siembra")
m1, m2, m3 = st.columns(3)
if not actual.empty:
    m1.metric("Fase actual", actual["Etapa"].iloc[0])
    m1.caption(actual["Descripción"].iloc[0])
else:
    m1.metric("Fase actual", "Inicial")
    m1.caption("Aún no se alcanza la primera fase")
m2.metric("GDD acumulados", f"{acum_hoy:.0f} °C·d")
m2.caption(f"{max(0, k_hoy)} días desde la siembra")
if not siguiente.empty:
    s = siguiente.iloc[0]
    falta = (s["Fecha"] - pd.Timestamp(hoy)).days if s["Fecha"] is not None else None
    m3.metric("Siguiente fase", s["Etapa"])
    m3.caption(f"{fmt(s['Fecha'])} · en ~{falta} días ({s['G'] - acum_hoy:.0f} °C·d restantes)"
               if falta is not None else "No se alcanza en el horizonte simulado")
else:
    m3.metric("Siguiente fase", "Ciclo completado")
    m3.caption("Prepare el siguiente ciclo")

# Línea de tiempo
max_g = etapas["G"].max() * 1.06
fig_t = go.Figure()
fig_t.add_shape(type="rect", x0=0, x1=max_g, y0=0, y1=1, line=dict(color="#1e3a5f"),
                fillcolor="rgba(30,58,95,.25)", layer="below")
fig_t.add_shape(type="rect", x0=0, x1=min(acum_hoy, max_g), y0=0, y1=1, line_width=0,
                fillcolor=COLOR_CULTIVO, opacity=0.4, layer="below")
for i, r in etapas.iterrows():
    hecho = acum_hoy >= r["G"]
    fig_t.add_shape(type="line", x0=r["G"], x1=r["G"], y0=-0.25, y1=1.25,
                    line=dict(color=COLOR_CULTIVO if hecho else "#5f7b9e", width=2))
    arriba = i % 2 == 0
    fig_t.add_annotation(x=r["G"], y=1.35 if arriba else -0.35, yanchor="bottom" if arriba else "top",
                         text=f"<b>{r['Etapa']}</b><br>{fmt(r['Fecha'])}", showarrow=False,
                         font=dict(size=11))
fig_t.add_trace(go.Scatter(x=[acum_hoy], y=[0.5], mode="markers+text", text=[f"{acum_hoy:.0f}"],
                           textposition="middle right", marker=dict(symbol="triangle-down", size=14,
                           color="white", line=dict(color="black", width=1)),
                           hovertemplate="GDD acumulados hoy: %{x:.0f}<extra></extra>"))
fig_t.update_layout(height=200, margin=dict(l=10, r=10, t=10, b=10), showlegend=False,
                    xaxis=dict(range=[-max_g * 0.02, max_g], title="GDD acumulados (°C·d)",
                               showgrid=False),
                    yaxis=dict(visible=False, range=[-1.4, 2.4]))
st.plotly_chart(fig_t, width="stretch")

# --------------------------------------------------------------------------
# Tabla de fases
# --------------------------------------------------------------------------
st.subheader("Fecha prevista de cada fase")
filas = []
for _, r in etapas.iterrows():
    hecho = acum_hoy >= r["G"]
    es_actual = not actual.empty and r["Etapa"] == actual["Etapa"].iloc[0]
    filas.append({
        "Fase": r["Etapa"],
        "GDD": r["G"],
        "Fecha prevista": fmt(r["Fecha"]),
        "DAS": r["DAS"],
        "Estado": "🟡 En curso" if es_actual else ("✅ Superada" if hecho else "🗓️ Planificada"),
    })
st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch")
st.caption("DAS: días después de la siembra.")

# --------------------------------------------------------------------------
# Curva de GDD acumulados por año
# --------------------------------------------------------------------------
st.subheader("Curva de GDD acumulados · comparación por año")
das_fin = etapas["DAS"].dropna().max() if etapas["DAS"].notna().any() else 220
dias_curva = int(min(MAX_DIAS, max(120, das_fin * 1.15)))
fig_c = go.Figure()
paleta = ["#3f6288", "#4d7aa6", "#5b92c4", "#69aae2", "#77c2ff", "#8fd0ff", "#a7dcff"]
for i, y in enumerate(anios):
    s0 = inicio_siembra(y, siembra) + pd.Timedelta(days=1)
    tramo = g.reindex(pd.date_range(s0, periods=dias_curva, freq="D"))
    if tramo.notna().sum() < 10:
        continue
    acum = tramo.fillna(0).cumsum().where(tramo.notna().cummax() & (tramo.index <= validos.index.max()))
    fig_c.add_trace(go.Scatter(x=np.arange(1, dias_curva + 1), y=acum, mode="lines", name=str(y),
                               line=dict(color=paleta[i % len(paleta)], width=1.2), opacity=0.75,
                               hovertemplate=f"{y} · día %{{x}}: %{{y:.0f}} °C·d<extra></extra>"))
cur = sea.iloc[:dias_curva]
obs_mask = cur.index <= pd.Timestamp(hoy)
fig_c.add_trace(go.Scatter(x=cur["das"][obs_mask], y=cur["gdd_acum"][obs_mask], mode="lines",
                           name="Temporada actual (observado)",
                           line=dict(color=COLOR_CULTIVO, width=3)))
k0 = max(0, int(obs_mask.sum()) - 1)
fig_c.add_trace(go.Scatter(x=cur["das"].iloc[k0:], y=cur["gdd_acum"].iloc[k0:], mode="lines",
                           name="Proyección (climatología)",
                           line=dict(color=COLOR_CULTIVO, width=2, dash="dot")))
for _, r in etapas.iterrows():
    fig_c.add_hline(y=r["G"], line=dict(color="#7fa8d0", width=1, dash="dash"),
                    annotation_text=f"{r['Etapa']} {r['G']}", annotation_position="right",
                    annotation_font_size=10)
fig_c.update_layout(height=430, margin=dict(l=10, r=190, t=10, b=10),
                    xaxis_title="Días desde la siembra", yaxis_title="GDD acumulados (°C·d)",
                    yaxis_range=[0, etapas["G"].max() * 1.1], hovermode="closest",
                    legend=dict(orientation="h", y=-0.18))
st.plotly_chart(fig_c, width="stretch")
st.caption(f"Siembra el {siembra.strftime('%d/%m')} de cada año · {anios[0]}–{anios[-1]}")

st.markdown("---")
formula = (f"GDD = max(min((Tmax+Tmin)/2, {tu:g}) − {tb:g}, 0)" if metodo == "media" else
           f"GDD = (min(max(Tmax, {tb:g}), {tu:g}) + min(max(Tmin, {tb:g}), {tu:g}))/2 − {tb:g}")
st.caption(
    f"{formula} · método {metodo}. Acumulación desde el día siguiente a la siembra; observaciones "
    "hasta la fecha de referencia y promedio climatológico diario de la estación después. "
    "Este calendario se basa solo en la temperatura; considere también humedad del suelo, "
    "estado real del cultivo e información agrícola local."
)
