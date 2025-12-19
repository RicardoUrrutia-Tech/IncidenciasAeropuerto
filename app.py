import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")
st.title("App Incidencias / Ausentismo / Asistencia")

# =========================
# Helpers
# =========================
CABIFY = {
    "m1": "1F123F",
    "m2": "362065",
    "m3": "4A2B8D",
    "m4": "5B34AC",
    "m5": "7145D6",
    "m6": "8A6EE4",
    "m7": "A697ED",
    "m8": "C4BDF5",
    "m9": "DFDAF8",
    "m10": "F5F1FC",
    "m11": "FAF8FE",
    "pink": "E83C96",
    "red": "E74A41",
    "orange": "EA8C2E",
    "yellow": "EFBD03",
    "blue": "4583D4",
    "green": "0C936B",  # ojo: tu input decía #OC936B (O letra), lo correcto es 0 (cero)
}

CLASIF_OPTS = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]

def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper().replace(".", "").replace(" ", "")
    return s

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def excel_to_df(file, sheet_index=0):
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def find_col(df: pd.DataFrame, candidates):
    """
    Devuelve el nombre real de columna que matchea cualquiera de candidates (case-insensitive, trim).
    """
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in norm_map:
            return norm_map[key]
    # match parcial
    for cand in candidates:
        key = str(cand).strip().lower()
        for k, real in norm_map.items():
            if k == key:
                return real
    return None

def get_num(df, colname_candidates):
    col = find_col(df, colname_candidates if isinstance(colname_candidates, list) else [colname_candidates])
    if not col:
        return pd.Series([0.0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

def safe_text_series(df, candidates, default=""):
    col = find_col(df, candidates)
    if not col:
        return pd.Series([default] * len(df))
    return df[col].astype(str).fillna(default)

def style_ws_cabify(ws):
    header_fill = PatternFill("solid", fgColor=CABIFY["m4"])
    header_font = Font(color="FFFFFF", bold=True)
    body_fill = PatternFill("solid", fgColor=CABIFY["m11"])
    alt_fill = PatternFill("solid", fgColor=CABIFY["m10"])

    thin = Side(style="thin", color=CABIFY["m7"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.fill = alt_fill if (row_idx % 2 == 0) else body_fill

    # Ajuste ancho simple
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                v = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(v))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 55)

    ws.freeze_panes = "A2"

def write_df_to_sheet(wb, sheet_name, df):
    ws = wb.create_sheet(title=sheet_name[:31])
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    return ws

def apply_dropdown(ws, df_cols, col_name, options):
    if col_name not in df_cols:
        return
    idx = list(df_cols).index(col_name) + 1
    col_letter = ws.cell(row=1, column=idx).column_letter
    dv = DataValidation(type="list", formula1=f"\"{','.join(options)}\"", allow_blank=False)
    dv.add(f"{col_letter}2:{col_letter}1048576")
    ws.add_data_validation(dv)

def set_date_format(ws, df_cols, col_name="Fecha"):
    if col_name not in df_cols:
        return
    idx = list(df_cols).index(col_name) + 1
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=idx)
        c.number_format = "dd-mm-yyyy"  # fecha corta

def to_excel_bytes(dfs: dict):
    """
    dfs: {"Hoja": df, ...}
    """
    output = BytesIO()
    wb = Workbook()
    # borrar sheet por defecto
    wb.remove(wb.active)

    for name, df in dfs.items():
        ws = write_df_to_sheet(wb, name, df)

        # dropdown en Clasificación Manual solo en la hoja principal
        if name == "Incidencias":
            apply_dropdown(ws, df.columns, "Clasificación Manual", CLASIF_OPTS)
            set_date_format(ws, df.columns, "Fecha")

        # formateo fechas en otras hojas si aplica
        if "Fecha" in df.columns:
            set_date_format(ws, df.columns, "Fecha")

        style_ws_cabify(ws)

    wb.save(output)
    output.seek(0)
    return output

# =========================
# UI Inputs
# =========================
with st.sidebar:
    st.header("Cargar archivos (Excel)")
    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_reporte_turnos = st.file_uploader("2) Reporte Turnos (Activos + Turnos)", type=["xlsx"])
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Hoja1=Inasistencias, Hoja2=Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Filtro opcional")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar.")

if not all([f_turnos, f_reporte_turnos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()

# =========================
# Load
# =========================
df_turnos = excel_to_df(f_turnos, 0)  # (por ahora no lo usamos, queda para siguientes reglas)
df_activos = excel_to_df(f_reporte_turnos, 0)

# Detalle Turnos Colaboradores:
# Hoja 1 (index 0): Inasistencias
# Hoja 2 (index 1): Asistencias
df_inasist = excel_to_df(f_detalle, 0)
df_asist = excel_to_df(f_detalle, 1)

# =========================
# Normalización columnas base (RUT, Área, etc.)
# =========================
rut_col_inas = find_col(df_inasist, ["RUT", "Rut", "rut"])
rut_col_as = find_col(df_asist, ["RUT", "Rut", "rut"])

if not rut_col_inas or not rut_col_as:
    st.error("No pude detectar la columna RUT en una de las hojas del 'Detalle Turnos Colaboradores'. Revisa que exista una columna llamada RUT/Rut.")
    st.stop()

df_inasist["RUT_norm"] = df_inasist[rut_col_inas].apply(normalize_rut)
df_asist["RUT_norm"] = df_asist[rut_col_as].apply(normalize_rut)

# Fecha base
dia_col_inas = find_col(df_inasist, ["Día", "Dia", "DIA", "día"])
df_inasist["Fecha_base"] = df_inasist[dia_col_inas].apply(try_parse_date_any) if dia_col_inas else pd.NaT

fecha_ent_col_as = find_col(df_asist, ["Fecha Entrada", "Fecha_Entrada", "Fecha entrada"])
dia_col_as = find_col(df_asist, ["Día", "Dia", "DIA", "día"])
if fecha_ent_col_as:
    df_asist["Fecha_base"] = df_asist[fecha_ent_col_as].apply(try_parse_date_any)
elif dia_col_as:
    df_asist["Fecha_base"] = df_asist[dia_col_as].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

# (Opcional) Filtrar Área
def maybe_filter_area(df):
    if not only_area:
        return df
    area_col = find_col(df, ["Área", "Area", "AREA"])
    if not area_col:
        return df
    return df[df[area_col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_inasist = maybe_filter_area(df_inasist)
df_asist = maybe_filter_area(df_asist)

# =========================
# Turnos planificados (Activos + Turnos) -> largo para contar turnos
# =========================
fixed_cols = [c for c in ["Nombre del Colaborador", "RUT", "Área", "Supervisor"] if c in df_activos.columns]
date_cols = [c for c in df_activos.columns if c not in fixed_cols]

if "RUT" not in df_activos.columns:
    # intenta detectar rut alternativo
    rut_col_act = find_col(df_activos, ["RUT", "Rut", "rut"])
    if rut_col_act and rut_col_act != "RUT":
        df_activos = df_activos.rename(columns={rut_col_act: "RUT"})

df_act_long = df_activos.melt(
    id_vars=[c for c in fixed_cols if c in df_activos.columns],
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_planificado"
)
df_act_long["Fecha_dt"] = df_act_long["Fecha"].apply(try_parse_date_any)
df_act_long["RUT_norm"] = df_act_long["RUT"].apply(normalize_rut) if "RUT" in df_act_long.columns else ""

df_act_long["Turno_planificado"] = df_act_long["Turno_planificado"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_planificado"].isin(["", "nan", "NaT", "None", "-"]), "Turno_planificado"] = ""

# =========================
# Construcción Incidencias
# =========================
inc_rows = []

# 1) Asistencias: SOLO si hay Retraso o Salida Anticipada
retr = get_num(df_asist, ["Retraso (horas)", "Retraso horas", "Retraso"])
sal = get_num(df_asist, ["Salida Anticipada (horas)", "Salida Anticipada", "Salida anticipada (horas)"])

mask_asist = (retr > 0) | (sal > 0)
df_asist_inc = df_asist[mask_asist].copy()

df_asist_inc["Fecha"] = df_asist_inc["Fecha_base"].dt.date
df_asist_inc["Nombre"] = safe_text_series(df_asist_inc, ["Nombre"], "")
df_asist_inc["Primer Apellido"] = safe_text_series(df_asist_inc, ["Primer Apellido", "Primer apellido"], "")
df_asist_inc["Segundo Apellido"] = safe_text_series(df_asist_inc, ["Segundo Apellido", "Segundo apellido"], "")
df_asist_inc["RUT"] = df_asist_inc[rut_col_as].astype(str)
df_asist_inc["Turno"] = safe_text_series(df_asist_inc, ["Turno"], "")
df_asist_inc["Especialidad"] = safe_text_series(df_asist_inc, ["Especialidad"], "")
df_asist_inc["Supervisor"] = safe_text_series(df_asist_inc, ["Supervisor"], "")

df_asist_inc["Tipo_Incidencia"] = "Marcaje/Turno"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + retr[mask_asist].astype(str).values
    + " | SalidaAnt_h=" + sal[mask_asist].astype(str).values
)

df_asist_inc["Clasificación Manual"] = "Seleccionar"

inc_rows.append(df_asist_inc[[
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Turno", "Especialidad", "Supervisor",
    "Tipo_Incidencia", "Detalle", "Clasificación Manual"
]])

# 2) Inasistencias: se listan para clasificar manualmente
df_inasist_inc = df_inasist.copy()

df_inasist_inc["Fecha"] = df_inasist_inc["Fecha_base"].dt.date
df_inasist_inc["Nombre"] = safe_text_series(df_inasist_inc, ["Nombre"], "")
df_inasist_inc["Primer Apellido"] = safe_text_series(df_inasist_inc, ["Primer Apellido", "Primer apellido"], "")
df_inasist_inc["Segundo Apellido"] = safe_text_series(df_inasist_inc, ["Segundo Apellido", "Segundo apellido"], "")
df_inasist_inc["RUT"] = df_inasist_inc[rut_col_inas].astype(str)
df_inasist_inc["Turno"] = safe_text_series(df_inasist_inc, ["Turno"], "")
df_inasist_inc["Especialidad"] = safe_text_series(df_inasist_inc, ["Especialidad"], "")
df_inasist_inc["Supervisor"] = safe_text_series(df_inasist_inc, ["Supervisor"], "")

mot = safe_text_series(df_inasist_inc, ["Motivo"], "")
df_inasist_inc["Tipo_Incidencia"] = "Inasistencia"
df_inasist_inc["Detalle"] = "Motivo=" + mot

df_inasist_inc["Clasificación Manual"] = "Seleccionar"

inc_rows.append(df_inasist_inc[[
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Turno", "Especialidad", "Supervisor",
    "Tipo_Incidencia", "Detalle", "Clasificación Manual"
]])

df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Limpieza / orden
df_incidencias["Fecha"] = pd.to_datetime(df_incidencias["Fecha"], errors="coerce")
df_incidencias = df_incidencias.sort_values(["Fecha", "RUT"], na_position="last").reset_index(drop=True)

# =========================
# UI: Editor + resumen dinámico
# =========================
st.subheader("Reporte Total de Incidencias por Comprobar")

edited = st.data_editor(
    df_incidencias,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=CLASIF_OPTS,
            required=True
        ),
    },
    key="editor_incidencias"
)

# Resumen dinámico por tipo (Injustificadas)
st.subheader("Resumen dinámico (solo Injustificada)")
df_inj = edited[edited["Clasificación Manual"] == "Injustificada"].copy()

if df_inj.empty:
    st.info("Aún no hay registros clasificados como 'Injustificada'.")
    resumen = pd.DataFrame(columns=["Tipo_Incidencia", "Cantidad"])
else:
    resumen = (
        df_inj.groupby(["Tipo_Incidencia"], dropna=False)
              .size()
              .reset_index(name="Cantidad")
              .sort_values("Cantidad", ascending=False)
    )
    st.dataframe(resumen, use_container_width=True)

# =========================
# Hoja: Cumplimiento por colaborador
# Base = turnos planificados (no vacíos) en reporte turnos
# Baja solo por injustificadas (incidencias + inasistencias)
# =========================
st.subheader("Cumplimiento por colaborador (dinámico)")

# Turnos planificados por RUT
df_act_long_valid = df_act_long[df_act_long["Turno_planificado"] != ""].copy()
turnos_plan = (
    df_act_long_valid.groupby("RUT_norm")
    .size()
    .reset_index(name="Turnos_planificados")
)

# Injustificadas por RUT
edited_tmp = edited.copy()
edited_tmp["RUT_norm"] = edited_tmp["RUT"].apply(normalize_rut)
inj_cnt = (
    edited_tmp[edited_tmp["Clasificación Manual"] == "Injustificada"]
    .groupby("RUT_norm")
    .size()
    .reset_index(name="Injustificadas")
)

cumpl = turnos_plan.merge(inj_cnt, on="RUT_norm", how="left")
cumpl["Injustificadas"] = cumpl["Injustificadas"].fillna(0).astype(int)
cumpl["Cumplimiento_%"] = (1 - (cumpl["Injustificadas"] / cumpl["Turnos_planificados"])).clip(lower=0, upper=1) * 100

# Traer nombre desde incidencias (o desde activos si quieres, aquí desde incidencias)
name_map = (
    edited_tmp.dropna(subset=["RUT_norm"])
    .drop_duplicates("RUT_norm")[["RUT_norm", "Nombre", "Primer Apellido", "Segundo Apellido"]]
)
cumpl = cumpl.merge(name_map, on="RUT_norm", how="left")

cumpl = cumpl[[
    "Nombre", "Primer Apellido", "Segundo Apellido", "RUT_norm",
    "Turnos_planificados", "Injustificadas", "Cumplimiento_%"
]].rename(columns={"RUT_norm": "RUT_norm_sin_puntos"}).sort_values("Cumplimiento_%", ascending=True)

st.dataframe(cumpl, use_container_width=True)

# =========================
# Export Excel (Cabify + dropdown real)
# =========================
st.subheader("Descarga")

# Preparar para export: fecha corta en excel (date) y texto limpio
edited_export = edited.copy()
edited_export["Fecha"] = pd.to_datetime(edited_export["Fecha"], errors="coerce")

excel_bytes = to_excel_bytes({
    "Incidencias": edited_export,
    "Resumen_Injustificadas": resumen,
    "Cumplimiento": cumpl
})

st.download_button(
    "Descargar Excel consolidado (Cabify + dropdown)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

