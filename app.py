import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation


# =========================
# Config
# =========================
st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")
st.title("App Incidencias / Ausentismo / Asistencia")


# =========================
# Helpers
# =========================
CABIFY_PURPLE = "1F123F"
CABIFY_LIGHT = "F5F1FC"

DROPDOWN_OPTS = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]

def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = s.replace(".", "").replace(" ", "")
    # deja el guión si existe
    return s

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def excel_to_df(file, sheet_index=0):
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def safe_col(df: pd.DataFrame, name: str):
    return name if name in df.columns else None

def as_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def maybe_filter_area(df: pd.DataFrame, only_area: str, col="Área"):
    if only_area and col in df.columns:
        return df[df[col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()
    return df

def get_date_range_from_data(dfs_dates):
    all_dates = pd.concat([s.dropna() for s in dfs_dates], ignore_index=True)
    if all_dates.empty:
        return None, None
    return all_dates.min().date(), all_dates.max().date()


# =========================
# Sidebar inputs
# =========================
with st.sidebar:
    st.header("Cargar archivos (Excel)")

    f_turnos_cod = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_turnos_colab = st.file_uploader("2) Turnos Colaboradores (Activos + Turnos)", type=["xlsx"])
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Hoja1=Inasistencias, Hoja2=Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Filtros / Reglas")

    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar.")

    min_inc_hours = st.number_input(
        "Tiempo mínimo de incidencia (horas)",
        min_value=0.0, value=0.0, step=0.25,
        help="Se considera incidencia solo si (Retraso + Salida Anticipada) >= este umbral."
    )

if not all([f_turnos_cod, f_turnos_colab, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()


# =========================
# Load data
# =========================
df_cod = excel_to_df(f_turnos_cod, 0)
df_turnos = excel_to_df(f_turnos_colab, 0)

# Detalle: Hoja1=Inasistencias, Hoja2=Asistencias (tal como tu Excel real)
df_inasist = excel_to_df(f_detalle, 0)
df_asist = excel_to_df(f_detalle, 1)

# Normalizaciones básicas
for df in [df_turnos, df_inasist, df_asist]:
    if "RUT" in df.columns:
        df["RUT_norm"] = df["RUT"].apply(normalize_rut)

# Parse fechas base
# Inasistencias: Día
if "Día" in df_inasist.columns:
    df_inasist["Fecha_base"] = df_inasist["Día"].apply(try_parse_date_any)
else:
    df_inasist["Fecha_base"] = pd.NaT

# Asistencias: Fecha Entrada
if "Fecha Entrada" in df_asist.columns:
    df_asist["Fecha_base"] = df_asist["Fecha Entrada"].apply(try_parse_date_any)
elif "Día" in df_asist.columns:
    df_asist["Fecha_base"] = df_asist["Día"].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

# Filtrar por área (si aplica)
df_turnos = maybe_filter_area(df_turnos, only_area, "Área")
df_inasist = maybe_filter_area(df_inasist, only_area, "Área")
df_asist = maybe_filter_area(df_asist, only_area, "Área")


# =========================
# Turnos (formato ancho -> largo) para planificados
# =========================
fixed_cols = [c for c in ["Nombre del Colaborador", "RUT", "Área", "Supervisor"] if c in df_turnos.columns]
date_cols = [c for c in df_turnos.columns if c not in fixed_cols]

df_turnos_long = df_turnos.melt(
    id_vars=fixed_cols,
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_Activo"
)

df_turnos_long["Fecha_dt"] = df_turnos_long["Fecha"].apply(try_parse_date_any)
df_turnos_long["RUT_norm"] = df_turnos_long["RUT"].apply(normalize_rut) if "RUT" in df_turnos_long.columns else ""

df_turnos_long["Turno_Activo"] = df_turnos_long["Turno_Activo"].astype(str).str.strip()
df_turnos_long.loc[df_turnos_long["Turno_Activo"].isin(["", "nan", "NaT", "None"]), "Turno_Activo"] = ""

# Turnos activos para cumplimiento: NO vacíos, NO L, NO "-", NO LIBRE
def is_planned_turno(x: str) -> bool:
    if x is None:
        return False
    s = str(x).strip().upper()
    if s in ["", "L", "-", "LIBRE"]:
        return False
    return True

df_turnos_long["Es_Turno_Planificado"] = df_turnos_long["Turno_Activo"].apply(is_planned_turno)


# =========================
# Rango de fechas (selector)
# =========================
min_d, max_d = get_date_range_from_data([df_inasist["Fecha_base"], df_asist["Fecha_base"], df_turnos_long["Fecha_dt"]])

with st.sidebar:
    st.subheader("Rango de fechas")
    if min_d and max_d:
        date_start, date_end = st.date_input(
            "Selecciona desde / hasta",
            value=(min_d, max_d),
            min_value=min_d,
            max_value=max_d,
        )
    else:
        st.warning("No pude inferir fechas desde los archivos. Revisar columnas de fecha.")
        st.stop()

# Normalizar por si streamlit devuelve 1 fecha
if isinstance(date_start, tuple) or isinstance(date_start, list):
    date_start, date_end = date_start

dt_start = pd.to_datetime(date_start)
dt_end = pd.to_datetime(date_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

df_asist = df_asist[(df_asist["Fecha_base"] >= dt_start) & (df_asist["Fecha_base"] <= dt_end)].copy()
df_inasist = df_inasist[(df_inasist["Fecha_base"] >= dt_start) & (df_inasist["Fecha_base"] <= dt_end)].copy()
df_turnos_long = df_turnos_long[(df_turnos_long["Fecha_dt"] >= dt_start) & (df_turnos_long["Fecha_dt"] <= dt_end)].copy()


# =========================
# Incidencias desde ASISTENCIAS
# Regla: SOLO si existe Retraso o Salida Anticipada
# y (Retraso + Salida) >= min_inc_hours
# =========================
col_retraso = safe_col(df_asist, "Retraso (horas)")
col_salida = safe_col(df_asist, "Salida Anticipada (horas)")

if not col_retraso or not col_salida:
    st.error("En Asistencias falta 'Retraso (horas)' o 'Salida Anticipada (horas)'.")
    st.stop()

df_asist["Retraso_h"] = as_num(df_asist[col_retraso])
df_asist["SalidaAnt_h"] = as_num(df_asist[col_salida])
df_asist["Inc_total_h"] = df_asist["Retraso_h"] + df_asist["SalidaAnt_h"]

mask_asist = (df_asist["Inc_total_h"] > 0) & (df_asist["Inc_total_h"] >= float(min_inc_hours))
df_asist_inc = df_asist[mask_asist].copy()

df_asist_inc["Tipo_Incidencia"] = "Marcaje/Turno"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + df_asist_inc["Retraso_h"].round(2).astype(str) +
    " | SalidaAnt_h=" + df_asist_inc["SalidaAnt_h"].round(2).astype(str)
)

# Campos de salida (los que tú pediste)
# Fecha, Nombre, Primer Apellido, Segundo Apellido, RUT, Turno, Especialidad, Supervisor, Tipo_Incidencia, Detalle, Clasificación Manual
def pick(df, col, default=""):
    return df[col] if col in df.columns else default

out_asist = pd.DataFrame({
    "Fecha": df_asist_inc["Fecha_base"],
    "Nombre": pick(df_asist_inc, "Nombre", ""),
    "Primer Apellido": pick(df_asist_inc, "Primer Apellido", ""),
    "Segundo Apellido": pick(df_asist_inc, "Segundo Apellido", ""),
    "RUT": df_asist_inc["RUT"],
    "Turno": pick(df_asist_inc, "Turno", ""),
    "Especialidad": pick(df_asist_inc, "Especialidad", ""),
    "Supervisor": pick(df_asist_inc, "Supervisor", ""),
    "Tipo_Incidencia": df_asist_inc["Tipo_Incidencia"],
    "Detalle": df_asist_inc["Detalle"],
    "Clasificación Manual": "Seleccionar",
})


# =========================
# Inasistencias (Hoja1) - solo Motivo "-" o vacío (inasistencia preliminar injustificada)
# =========================
motivo_col = safe_col(df_inasist, "Motivo")
if motivo_col:
    motivo = df_inasist[motivo_col].astype(str).str.strip()
    mask_inas = motivo.isna() | (motivo == "") | (motivo == "-")
else:
    mask_inas = pd.Series([True]*len(df_inasist))

df_inas_inc = df_inasist[mask_inas].copy()
df_inas_inc["Tipo_Incidencia"] = "Inasistencia"

out_inas = pd.DataFrame({
    "Fecha": df_inas_inc["Fecha_base"],
    "Nombre": pick(df_inas_inc, "Nombre", ""),
    "Primer Apellido": pick(df_inas_inc, "Primer Apellido", ""),
    "Segundo Apellido": pick(df_inas_inc, "Segundo Apellido", ""),
    "RUT": df_inas_inc["RUT"],
    "Turno": pick(df_inas_inc, "Turno", ""),
    "Especialidad": pick(df_inas_inc, "Especialidad", ""),
    "Supervisor": pick(df_inas_inc, "Supervisor", ""),
    "Tipo_Incidencia": df_inas_inc["Tipo_Incidencia"],
    "Detalle": (df_inas_inc[motivo_col].astype(str) if motivo_col else ""),
    "Clasificación Manual": "Seleccionar",
})


# =========================
# Consolidado Detalle
# =========================
df_detalle = pd.concat([out_asist, out_inas], ignore_index=True)

# Normalizar RUT para cruces
df_detalle["RUT_norm"] = df_detalle["RUT"].apply(normalize_rut)

# Orden / formato
df_detalle["Fecha"] = pd.to_datetime(df_detalle["Fecha"], errors="coerce")
df_detalle = df_detalle.sort_values(["Fecha", "RUT_norm"], na_position="last").reset_index(drop=True)

# Vista corta de fecha en pantalla
df_detalle_view = df_detalle.copy()
df_detalle_view["Fecha"] = df_detalle_view["Fecha"].dt.date


# =========================
# Editor (y resumen dinámico en app)
# =========================
st.subheader("Detalle de Incidencias por Comprobar")

edited = st.data_editor(
    df_detalle_view.drop(columns=["RUT_norm"], errors="ignore"),
    use_container_width=True,
    column_config={
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=DROPDOWN_OPTS
        )
    }
)

# Resumen dinámico (se recalcula con lo editado)
st.subheader("Resumen dinámico (en base a lo editado)")

edited_calc = edited.copy()
edited_calc["Fecha"] = pd.to_datetime(edited_calc["Fecha"], errors="coerce")
resumen = (
    edited_calc.groupby(["Tipo_Incidencia", "Clasificación Manual"], dropna=False)
    .size().reset_index(name="Cantidad")
    .sort_values("Cantidad", ascending=False)
)
st.dataframe(resumen, use_container_width=True)


# =========================
# Cumplimiento (TODOS los trabajadores del reporte de turnos)
# planned_turnos = conteo de fechas con turno planificado (no L / no vacío) en el rango
# unjustified = incidencias clasificadas como "Injustificada"
# cumplimiento = 1 - unjustified / planned_turnos  (si planned_turnos==0 => 1.0)
# =========================
workers_cols = [c for c in ["Nombre del Colaborador", "RUT", "Supervisor", "Área"] if c in df_turnos.columns]
df_workers = df_turnos[workers_cols].drop_duplicates().copy()
df_workers["RUT_norm"] = df_workers["RUT"].apply(normalize_rut)

planned = (
    df_turnos_long[df_turnos_long["Es_Turno_Planificado"]]
    .groupby("RUT_norm", as_index=False)
    .agg(Turnos_Planificados=("Es_Turno_Planificado", "size"))
)

# injustificadas desde lo EDITADO (en app)
edited_for_comp = edited.copy()
edited_for_comp["RUT_norm"] = edited_for_comp["RUT"].apply(normalize_rut)

unjust = (
    edited_for_comp[edited_for_comp["Clasificación Manual"] == "Injustificada"]
    .groupby("RUT_norm", as_index=False)
    .size()
    .rename(columns={"size": "Inc_Injustificadas"})
)

df_cump = df_workers.merge(planned, on="RUT_norm", how="left").merge(unjust, on="RUT_norm", how="left")
df_cump["Turnos_Planificados"] = df_cump["Turnos_Planificados"].fillna(0).astype(int)
df_cump["Inc_Injustificadas"] = df_cump["Inc_Injustificadas"].fillna(0).astype(int)

df_cump["Cumplimiento_pct"] = np.where(
    df_cump["Turnos_Planificados"] > 0,
    (1 - (df_cump["Inc_Injustificadas"] / df_cump["Turnos_Planificados"])) * 100,
    100.0
)
df_cump["Cumplimiento_pct"] = df_cump["Cumplimiento_pct"].clip(lower=0).round(2)

st.subheader("Cumplimiento (vista en app)")
st.dataframe(
    df_cump[["Nombre del Colaborador", "RUT", "Supervisor", "Turnos_Planificados", "Inc_Injustificadas", "Cumplimiento_pct"]],
    use_container_width=True
)


# =========================
# Export Excel (con validación + estilo + fórmulas para que sea dinámico en Excel)
# =========================
def build_excel_bytes(detalle_df: pd.DataFrame, workers_df: pd.DataFrame, start_d: date, end_d: date) -> BytesIO:
    wb = Workbook()

    # Estilos
    header_fill = PatternFill("solid", fgColor=CABIFY_PURPLE)
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(vertical="center", wrap_text=True)

    # ---------------- Sheet: Detalle
    ws_det = wb.active
    ws_det.title = "Detalle"

    # Asegurar tipo fecha
    dfw = detalle_df.copy()
    dfw["Fecha"] = pd.to_datetime(dfw["Fecha"], errors="coerce")

    # Orden exacto requerido
    cols = ["Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT", "Turno",
            "Especialidad", "Supervisor", "Tipo_Incidencia", "Detalle", "Clasificación Manual"]
    dfw = dfw[cols].copy()

    # Escribir dataframe
    for r_idx, row in enumerate(dataframe_to_rows(dfw, index=False, header=True), start=1):
        ws_det.append(row)

        # Header
        if r_idx == 1:
            for c in range(1, len(cols) + 1):
                cell = ws_det.cell(row=1, column=c)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border
        else:
            for c in range(1, len(cols) + 1):
                cell = ws_det.cell(row=r_idx, column=c)
                cell.border = border
                cell.alignment = center

    # Formato fecha corta en Excel (col 1)
    for r in range(2, ws_det.max_row + 1):
        ws_det.cell(row=r, column=1).number_format = "dd-mm-yyyy"

    # Autosize básico
    for col_idx in range(1, len(cols) + 1):
        ws_det.column_dimensions[chr(64 + col_idx)].width = 18
    ws_det.column_dimensions["J"].width = 40  # Detalle
    ws_det.column_dimensions["K"].width = 26  # Clasificación

    # DataValidation dropdown en "Clasificación Manual" (col K)
    dv = DataValidation(type="list", formula1=f"\"{','.join(DROPDOWN_OPTS)}\"", allow_blank=False)
    ws_det.add_data_validation(dv)
    dv.add(f"K2:K1048576")

    # Set default "Seleccionar" si viene vacío
    for r in range(2, ws_det.max_row + 1):
        c = ws_det.cell(row=r, column=11)
        if c.value is None or str(c.value).strip() == "":
            c.value = "Seleccionar"

    # ---------------- Sheet: Resumen (DINÁMICO EN EXCEL)
    ws_res = wb.create_sheet("Resumen")
    ws_res.append(["Tipo_Incidencia", "Clasificación Manual", "Cantidad"])
    for c in range(1, 4):
        cell = ws_res.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    tipos = ["Marcaje/Turno", "Inasistencia"]
    for t in tipos:
        for cl in DROPDOWN_OPTS:
            ws_res.append([t, cl, None])

    # Fórmula COUNTIFS (usa columnas: I=Tipo_Incidencia, K=Clasificación)
    for r in range(2, ws_res.max_row + 1):
        t = ws_res.cell(row=r, column=1).value
        cl = ws_res.cell(row=r, column=2).value
        ws_res.cell(row=r, column=3).value = f'=COUNTIFS(Detalle!$I:$I,"{t}",Detalle!$K:$K,"{cl}")'
        for c in range(1, 4):
            ws_res.cell(row=r, column=c).border = border
            ws_res.cell(row=r, column=c).alignment = center

    ws_res.column_dimensions["A"].width = 18
    ws_res.column_dimensions["B"].width = 28
    ws_res.column_dimensions["C"].width = 12

    # ---------------- Sheet: Cumplimiento (DINÁMICO EN EXCEL en base a clasificación)
    ws_c = wb.create_sheet("Cumplimiento")
    ws_c.append(["Nombre del Colaborador", "RUT", "Supervisor", "Turnos_Planificados", "Inc_Injustificadas", "Cumplimiento_%"])
    for c in range(1, 7):
        cell = ws_c.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Calcula planificados desde workers_df (ya viene calculado en app) -> lo pasamos estático
    # La parte dinámica será el conteo de "Injustificada" desde Detalle
    for _, row in workers_df.iterrows():
        ws_c.append([
            row.get("Nombre del Colaborador", ""),
            row.get("RUT", ""),
            row.get("Supervisor", ""),
            int(row.get("Turnos_Planificados", 0)),
            None,
            None
        ])

    # Celdas de control (fechas) en una zona discreta (H1:H2)
    ws_c["H1"] = "FechaInicio"
    ws_c["H2"] = "FechaFin"
    ws_c["I1"] = start_d
    ws_c["I2"] = end_d
    ws_c["I1"].number_format = "dd-mm-yyyy"
    ws_c["I2"].number_format = "dd-mm-yyyy"

    # Fórmulas:
    # Inc_Injustificadas = COUNTIFS(Detalle!E:E, RUT, Detalle!K:K, "Injustificada", Detalle!A:A, >=I1, Detalle!A:A, <=I2)
    # Cumplimiento = IF(D=0,1,1 - E/D)
    for r in range(2, ws_c.max_row + 1):
        rut_cell = ws_c.cell(row=r, column=2).coordinate  # B
        ws_c.cell(row=r, column=5).value = (
            f'=COUNTIFS(Detalle!$E:$E,{rut_cell},Detalle!$K:$K,"Injustificada",'
            f'Detalle!$A:$A,">="&$I$1,Detalle!$A:$A,"<="&$I$2)'
        )
        planned_cell = ws_c.cell(row=r, column=4).coordinate  # D
        unjust_cell = ws_c.cell(row=r, column=5).coordinate   # E
        ws_c.cell(row=r, column=6).value = f'=IF({planned_cell}=0,1,MAX(0,1-({unjust_cell}/{planned_cell})))'
        ws_c.cell(row=r, column=6).number_format = "0.00%"

        for c in range(1, 7):
            ws_c.cell(row=r, column=c).border = border
            ws_c.cell(row=r, column=c).alignment = center

    ws_c.column_dimensions["A"].width = 34
    ws_c.column_dimensions["B"].width = 16
    ws_c.column_dimensions["C"].width = 28
    ws_c.column_dimensions["D"].width = 18
    ws_c.column_dimensions["E"].width = 18
    ws_c.column_dimensions["F"].width = 16

    # Guardar a bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# Preparar workers_df para export (con turnos planificados)
planned_export = (
    df_turnos_long[df_turnos_long["Es_Turno_Planificado"]]
    .groupby("RUT_norm", as_index=False)
    .size()
    .rename(columns={"size": "Turnos_Planificados"})
)
workers_export = df_workers.merge(planned_export, on="RUT_norm", how="left")
workers_export["Turnos_Planificados"] = workers_export["Turnos_Planificados"].fillna(0).astype(int)

# Excel bytes (detalle = lo editado)
detalle_export = edited.copy()
detalle_export["Fecha"] = pd.to_datetime(detalle_export["Fecha"], errors="coerce")

st.subheader("Descarga")
excel_bytes = build_excel_bytes(
    detalle_df=detalle_export,
    workers_df=workers_export,
    start_d=date_start,
    end_d=date_end
)

st.download_button(
    "Descargar Excel (Cabify + Dropdown + Resumen/Cumplimiento dinámicos)",
    data=excel_bytes,
    file_name="reporte_incidencias_cabify.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
